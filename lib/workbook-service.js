import { createHash, randomUUID } from 'node:crypto';
import { HttpError } from './http.js';
import {
  appendEntry,
  assignIds,
  clearEntries,
  deleteEntry,
  loadWorkbook,
  normalizeValues,
  parseEntries,
  rowHash,
  serialize,
  updateEntry,
} from './xlsx-sheet.js';

// `io` kapselt den Dateizugriff: { meta(), download(), upload(buf, etag) }.
// upload wirft PreconditionFailed, wenn die Datei seit dem Download veraendert wurde (HTTP 412 bei If-Match).
export class PreconditionFailed extends Error {}

// In der Excel-Datei werden Ausgaben als negative Betraege gefuehrt (Zahlenformat zeigt Minus rot).
// Die App nimmt positive Euro-Betraege entgegen und schreibt sie mit diesem Vorzeichen.
export const EXPENSE_SIGN = -1;

const MAX_ATTEMPTS = 4;
const ID_PATTERN = /^[A-Za-z0-9-]{8,64}$/;

// Konsistenter Stand: Metadaten vor und nach dem Download muessen denselben eTag haben.
async function snapshot(io) {
  for (let i = 0; i < 3; i++) {
    const before = await io.meta();
    const buf = await io.download();
    const after = await io.meta();
    if (before.eTag === after.eTag) return { buf, meta: after };
  }
  throw new HttpError(409, 'busy', 'Die Datei wird gerade veraendert. Bitte erneut versuchen.');
}

const fileInfo = (m) => ({ name: m.name, etag: m.eTag, ctag: m.cTag, lastModified: m.lastModifiedDateTime, lastModifiedBy: m.lastModifiedBy?.user?.displayName });

const listHash = (entries) => createHash('sha256').update(entries.map((e) => `${e.id}:${e.hash}`).join('|')).digest('hex').slice(0, 16);

function summarize(entries) {
  const total = entries.reduce((s, e) => s + (typeof e.betrag === 'number' ? e.betrag : 0), 0);
  return { count: entries.length, total: Math.round(total * 100) / 100 };
}

function describe(wb) {
  const { entries, sheet } = parseEntries(wb);
  return { sheet, rows: entries, summary: summarize(entries), missingIds: entries.filter((e) => e.id === '').length, listHash: listHash(entries) };
}

// Eingabe der App: positiver Euro-Betrag -> gespeicherter Wert mit Ausgaben-Vorzeichen.
// Nur mit `refund: true` (Rueckerstattung) wird der Betrag positiv gespeichert.
function fromInput(input) {
  const v = normalizeValues(input);
  if (!(Math.abs(v.betrag) > 0)) throw new HttpError(422, 'bad_value', 'Betrag muss groesser als 0 sein');
  return { ...v, betrag: (input.refund === true ? -EXPENSE_SIGN : EXPENSE_SIGN) * Math.abs(v.betrag) };
}

// Fuehrt `mutate(wb, describe)` auf frischem Stand aus und laedt mit If-Match hoch. Bei 412 wird auf
// frischen Daten wiederholt, damit fremde Aenderungen erhalten bleiben und nichts unbemerkt ueberschrieben wird.
async function mutateWithRetry(io, mutate) {
  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    const { buf, meta } = await snapshot(io);
    const wb = await loadWorkbook(buf);
    const outcome = await mutate(wb);
    if (outcome.skipUpload) return { ...outcome.result, retries: attempt, file: fileInfo(meta) };
    try {
      const saved = await io.upload(await serialize(wb), meta.eTag);
      return { ...outcome.result, retries: attempt, file: fileInfo(saved) };
    } catch (e) {
      if (!(e instanceof PreconditionFailed)) throw e;
    }
  }
  throw new HttpError(409, 'busy', 'Speichern nach mehreren Versuchen nicht moeglich: Datei wird laufend veraendert.');
}

export async function readSheet(io) {
  const { buf, meta } = await snapshot(io);
  return { file: fileInfo(meta), ...describe(await loadWorkbook(buf)) };
}

const MAX_BATCH = 20;

// Mehrere Eintraege (z. B. aus einem Diktat) in einem Durchgang: ein Download, ein Upload, alles oder nichts.
// Jede ID kommt vom Client, bereits vorhandene werden uebersprungen (Wiederholung ist sicher).
export async function appendRecords(io, { records }) {
  if (!Array.isArray(records) || records.length < 1 || records.length > MAX_BATCH) {
    throw new HttpError(400, 'bad_request', `1 bis ${MAX_BATCH} Eintraege erforderlich`);
  }
  const prepared = records.map((r) => {
    if (r.id !== undefined && !ID_PATTERN.test(String(r.id))) throw new HttpError(400, 'bad_id', 'Ungueltige ID');
    return { id: r.id ?? randomUUID(), values: fromInput(r) };
  });
  if (new Set(prepared.map((p) => p.id)).size !== prepared.length) throw new HttpError(400, 'bad_id', 'IDs muessen eindeutig sein');
  return mutateWithRetry(io, (wb) => {
    const known = new Set(parseEntries(wb).entries.map((e) => e.id));
    const fresh = prepared.filter((p) => !known.has(p.id));
    if (!fresh.length) return { skipUpload: true, result: { ids: prepared.map((p) => p.id), added: 0 } };
    for (const p of fresh) appendEntry(wb, { id: p.id, ...p.values });
    return { result: { ids: prepared.map((p) => p.id), added: fresh.length } };
  });
}

// Die ID kommt vom Client (einmal pro Formular erzeugt): Wiederholungen nach Netzfehlern erzeugen so keine Dubletten.
export async function appendRecord(io, input) {
  const values = fromInput(input);
  if (input.id !== undefined && !ID_PATTERN.test(String(input.id))) throw new HttpError(400, 'bad_id', 'Ungueltige ID');
  const id = input.id ?? randomUUID();
  return mutateWithRetry(io, (wb) => {
    const existing = parseEntries(wb).entries.find((e) => e.id === id);
    if (existing) return { skipUpload: true, result: { id, row: existing.row, duplicate: true } };
    return { result: { id, row: appendEntry(wb, { id, ...values }), duplicate: false } };
  });
}

export async function updateRecord(io, { id, expectedHash, values: input }) {
  if (!id || !expectedHash) throw new HttpError(400, 'bad_request', 'id und expectedHash erforderlich');
  const values = fromInput(input);
  return mutateWithRetry(io, (wb) => {
    const current = parseEntries(wb).entries.find((e) => e.id === String(id));
    if (!current) throw new HttpError(404, 'row_not_found', 'Eintrag existiert nicht mehr (in Excel geloescht?)', { id });
    if (current.hash !== expectedHash) {
      // Zeile wurde seit dem Einlesen in Excel geaendert -> nichts ueberschreiben.
      throw new HttpError(409, 'conflict', 'Der Eintrag wurde in der Zwischenzeit geaendert. Es wurde nichts gespeichert.', { current });
    }
    updateEntry(wb, current.row, values);
    return { result: { id, hash: rowHash({ ...current, ...values }) } };
  });
}

// Loescht einen Eintrag nur, wenn er noch dem Stand entspricht, den der Benutzer bestaetigt hat.
// Idempotent: Ist die ID nicht mehr vorhanden, wird nichts hochgeladen (Wiederholung nach Netzfehler ist sicher).
export async function deleteRecord(io, { id, expectedHash }) {
  if (!id || !expectedHash) throw new HttpError(400, 'bad_request', 'id und expectedHash erforderlich');
  return mutateWithRetry(io, (wb) => {
    const current = parseEntries(wb).entries.find((e) => e.id === String(id));
    if (!current) return { skipUpload: true, result: { id, deleted: false } };
    if (current.hash !== expectedHash) {
      throw new HttpError(409, 'conflict', 'Der Eintrag wurde in der Zwischenzeit geaendert. Es wurde nichts geloescht.', { current });
    }
    deleteEntry(wb, current.row);
    return { result: { id, deleted: true } };
  });
}

// Ergaenzt IDs nur bei Eintraegen ohne ID (vorhandene bleiben unveraendert). Idempotent.
export async function assignMissingIds(io) {
  return mutateWithRetry(io, (wb) => {
    const n = assignIds(wb, () => randomUUID());
    return n === 0 ? { skipUpload: true, result: { assigned: 0 } } : { result: { assigned: n } };
  });
}

// "Excel leeren": nur mit dem Stand, den der Benutzer bestaetigt hat (Anzahl/Summe aus listHash abgeleitet).
export async function clearList(io, { expectedListHash }) {
  if (!expectedListHash) throw new HttpError(400, 'bad_request', 'expectedListHash erforderlich');
  return mutateWithRetry(io, (wb) => {
    const now = describe(wb);
    if (now.listHash !== expectedListHash) {
      throw new HttpError(409, 'conflict', 'Die Liste hat sich seit der Bestaetigung geaendert. Es wurde nichts geloescht.', { summary: now.summary });
    }
    if (now.rows.some((e) => e.id === '')) {
      throw new HttpError(409, 'missing_ids', 'Es gibt Eintraege ohne ID. Bitte zuerst die IDs ergaenzen.', { missingIds: now.missingIds });
    }
    const cleared = clearEntries(wb);
    return { result: { cleared, summaryBefore: now.summary } };
  });
}
