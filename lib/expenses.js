import { createHash, randomUUID } from 'node:crypto';
import { decrypt, encrypt } from './crypto.js';
import { HttpError } from './http.js';
import { store } from './store.js';
import { favoritesHash, normalizeFavorite, normalizeValues, rowHash } from './xlsx-sheet.js';

// Ausgaben und Favoriten liegen in der Datenbank, jeder Datensatz einzeln und mit AES-256-GCM verschluesselt. Der
// Schluessel stammt aus dem Passwort des Benutzers und wird nie gespeichert (lib/vault.js). In der Datenbank steht nur
// ein Chiffrat. Die Excel-Datei wird erst
// durch "In Excel uebertragen" beschrieben (siehe workbook-service.syncEntries); Aenderungen in Excel werden ignoriert.
//
// Collections: expenses:<vaultId> (key = ID), tombstones:<vaultId> (bereits uebertragene, dann geloeschte IDs),
// favorites (key = vaultId), meta (key = vaultId).

// In der Excel-Datei werden Ausgaben als negative Betraege gefuehrt (Zahlenformat zeigt Minus rot).
// Die App nimmt positive Euro-Betraege entgegen und speichert sie mit diesem Vorzeichen.
export const EXPENSE_SIGN = -1;
export const MAX_FAVORITES = 30;
const MAX_BATCH = 20;
const ID_PATTERN = /^[A-Za-z0-9-]{8,64}$/;


// `vault` = { id, key } (siehe lib/vault.js): id trennt die Daten, key verschluesselt sie.
const seal = (vault, obj) => ({ v: 1, d: encrypt(JSON.stringify(obj), vault.key) });
const open = (vault, value) => JSON.parse(decrypt(value.d, vault.key));
const expensesOf = (vault) => `expenses:${vault.id}`;
const tombstonesOf = (vault) => `tombstones:${vault.id}`;

// Eingabe der App: positiver Euro-Betrag -> gespeicherter Wert mit Ausgaben-Vorzeichen.
// Nur mit `refund: true` (Rueckerstattung) wird der Betrag positiv gespeichert.
export function fromInput(input) {
  const v = normalizeValues(input);
  if (!(Math.abs(v.betrag) > 0)) throw new HttpError(422, 'bad_value', 'Betrag muss groesser als 0 sein');
  return { ...v, betrag: (input.refund === true ? -EXPENSE_SIGN : EXPENSE_SIGN) * Math.abs(v.betrag) };
}

const publicEntry = (e) => ({ id: e.id, ausgabe: e.ausgabe, ort: e.ort, betrag: e.betrag, hash: rowHash(e) });

// Alle Ausgaben in Erfassungsreihenfolge (aelteste zuerst), entschluesselt, samt Uebertragungsstand.
export async function listEntries(vault) {
  const rows = await store.list(expensesOf(vault));
  return rows
    .map(({ value }) => open(vault, value))
    .sort((a, b) => a.created - b.created || a.id.localeCompare(b.id))
    .map((e) => ({ ...e, hash: rowHash(e) }));
}

const listHash = (entries) => createHash('sha256').update(entries.map((e) => `${e.id}:${e.hash}`).join('|')).digest('hex').slice(0, 16);

export async function listTombstones(vault) {
  return (await store.list(tombstonesOf(vault))).map((t) => t.key);
}

export async function readFavorites(vault) {
  const raw = await store.get('favorites', vault.id);
  const items = raw ? open(vault, raw) : [];
  return { exists: true, items, hash: favoritesHash(items) };
}

// Stand der App: Liste, Summe, Favoriten und was noch nicht in Excel steht. Aufbau wie frueher beim Lesen der Datei,
// damit die Oberflaeche nach jeder Aenderung ohne weiteres Lesen auskommt.
export async function readState(vault) {
  const [entries, tombstones, favorites, meta] = await Promise.all([
    listEntries(vault),
    listTombstones(vault),
    readFavorites(vault),
    store.get('meta', vault.id),
  ]);
  const total = entries.reduce((s, e) => s + (typeof e.betrag === 'number' ? e.betrag : 0), 0);
  return {
    rows: entries.map(publicEntry),
    summary: { count: entries.length, total: Math.round(total * 100) / 100 },
    missingIds: 0,
    listHash: listHash(entries),
    favorites,
    pending: entries.filter((e) => e.syncedHash !== e.hash).length + tombstones.length,
    lastTransfer: meta?.lastTransfer ?? null,
  };
}

const checkId = (id) => {
  if (id !== undefined && !ID_PATTERN.test(String(id))) throw new HttpError(400, 'bad_id', 'Ungueltige ID');
};

// Die ID kommt vom Client (einmal pro Formular erzeugt): Wiederholungen nach Netzfehlern erzeugen keine Dubletten.
export async function addEntries(vault, records) {
  if (!Array.isArray(records) || records.length < 1 || records.length > MAX_BATCH) {
    throw new HttpError(400, 'bad_request', `1 bis ${MAX_BATCH} Eintraege erforderlich`);
  }
  const prepared = records.map((r) => {
    checkId(r.id);
    return { id: String(r.id ?? randomUUID()), values: fromInput(r) };
  });
  if (new Set(prepared.map((p) => p.id)).size !== prepared.length) throw new HttpError(400, 'bad_id', 'IDs muessen eindeutig sein');
  const known = new Set((await store.list(expensesOf(vault))).map((r) => r.key));
  const fresh = prepared.filter((p) => !known.has(p.id));
  const now = Date.now();
  await Promise.all(fresh.map((p, i) => store.set(expensesOf(vault), p.id, seal(vault, { id: p.id, ...p.values, created: now + i }))));
  return { ids: prepared.map((p) => p.id), added: fresh.length };
}

export async function updateEntry(vault, { id, expectedHash, values: input }) {
  if (!id || !expectedHash) throw new HttpError(400, 'bad_request', 'id und expectedHash erforderlich');
  const values = fromInput(input);
  const raw = await store.get(expensesOf(vault), String(id));
  if (!raw) throw new HttpError(404, 'row_not_found', 'Eintrag existiert nicht mehr', { id });
  const current = open(vault, raw);
  if (rowHash(current) !== expectedHash) {
    throw new HttpError(409, 'conflict', 'Der Eintrag wurde in der Zwischenzeit geaendert. Es wurde nichts gespeichert.', { current: publicEntry(current) });
  }
  const next = { ...current, ...values };
  await store.set(expensesOf(vault), current.id, seal(vault, next));
  return { id: current.id, hash: rowHash(next) };
}

// Idempotent. War der Eintrag schon in Excel, wird er beim naechsten Uebertragen auch dort entfernt.
export async function deleteEntry(vault, { id, expectedHash }) {
  if (!id || !expectedHash) throw new HttpError(400, 'bad_request', 'id und expectedHash erforderlich');
  const raw = await store.get(expensesOf(vault), String(id));
  if (!raw) return { id, deleted: false };
  const current = open(vault, raw);
  if (rowHash(current) !== expectedHash) {
    throw new HttpError(409, 'conflict', 'Der Eintrag wurde in der Zwischenzeit geaendert. Es wurde nichts geloescht.', { current: publicEntry(current) });
  }
  if (current.syncedHash) await store.set(tombstonesOf(vault), current.id, { at: Date.now() });
  await store.delete(expensesOf(vault), current.id);
  return { id: current.id, deleted: true };
}

// Speichert die komplette Favoritenliste, nur mit dem Stand, den der Benutzer gesehen hat (zweites Geraet).
export async function saveFavorites(vault, { favorites, expectedHash }) {
  if (!Array.isArray(favorites) || favorites.length > MAX_FAVORITES) {
    throw new HttpError(400, 'bad_request', `Hoechstens ${MAX_FAVORITES} Favoriten erlaubt`);
  }
  if (typeof expectedHash !== 'string' || !expectedHash) throw new HttpError(400, 'bad_request', 'expectedHash erforderlich');
  const items = favorites.map(normalizeFavorite);
  const keys = items.map((f) => `${f.ausgabe}\u0000${f.ort}`.toLocaleLowerCase('de-DE'));
  if (new Set(keys).size !== keys.length) throw new HttpError(422, 'bad_value', 'Favoriten muessen eindeutig sein (Ausgabe und Ort)');
  const current = await readFavorites(vault);
  if (current.hash !== expectedHash) {
    throw new HttpError(409, 'conflict', 'Die Favoriten wurden in der Zwischenzeit auf einem anderen Geraet geaendert. Es wurde nichts gespeichert.', { favorites: current });
  }
  await store.set('favorites', vault.id, seal(vault, items));
  return { favorites: { exists: true, items, hash: favoritesHash(items) } };
}

// Nach erfolgreichem Schreiben in die Excel-Datei: Eintraege als uebertragen markieren (nur wenn sie sich seither
// nicht geaendert haben), Loesch-Vermerke entfernen.
export async function markTransferred(vault, { transferred, deletedIds }) {
  const byId = new Map(transferred.map((t) => [t.id, t.hash]));
  const rows = await store.list(expensesOf(vault));
  await Promise.all(rows.map(async ({ key, value }) => {
    const e = open(vault, value);
    if (byId.get(key) === rowHash(e) && e.syncedHash !== byId.get(key)) await store.set(expensesOf(vault), key, seal(vault, { ...e, syncedHash: byId.get(key) }));
  }));
  await Promise.all(deletedIds.map((id) => store.delete(tombstonesOf(vault), id)));
  await store.set('meta', vault.id, { ...((await store.get('meta', vault.id)) ?? {}), lastTransfer: new Date().toISOString() });
}

// "Excel leeren" -> App-Seite: alle Ausgaben und Loesch-Vermerke weg. Favoriten bleiben.
export async function clearEntries(vault) {
  await Promise.all([store.clear(expensesOf(vault)), store.clear(tombstonesOf(vault))]);
}

// Einmalige Uebernahme des bisherigen Stands: Wer die App schon mit der Excel-Datei als Speicher genutzt hat, behaelt
// Ausgaben (mit ID) und Favoriten. Danach zaehlt nur noch die Datenbank. `loadExcel` liefert { rows, favorites } oder wirft.
// Scheitert das Lesen, wird beim naechsten Mal erneut versucht (vorhandene IDs werden nie ueberschrieben).
export async function importOnce(vault, loadExcel) {
  const meta = (await store.get('meta', vault.id)) ?? {};
  if (meta.imported) return false;
  let snapshot;
  try {
    snapshot = await loadExcel();
  } catch {
    return false;
  }
  const known = new Set((await store.list(expensesOf(vault))).map((r) => r.key));
  const rows = snapshot.rows.filter((r) => r.id && ID_PATTERN.test(r.id) && typeof r.betrag === 'number' && r.ausgabe && !known.has(r.id));
  await Promise.all(rows.map((r, i) => {
    const e = { id: r.id, ausgabe: r.ausgabe, ort: r.ort, betrag: r.betrag, created: i };
    return store.set(expensesOf(vault), r.id, seal(vault, { ...e, syncedHash: rowHash(e) }));
  }));
  if (!(await store.get('favorites', vault.id)) && snapshot.favorites?.items?.length) {
    await store.set('favorites', vault.id, seal(vault, snapshot.favorites.items.slice(0, MAX_FAVORITES)));
  }
  await store.set('meta', vault.id, { ...((await store.get('meta', vault.id)) ?? {}), imported: true });
  return true;
}
