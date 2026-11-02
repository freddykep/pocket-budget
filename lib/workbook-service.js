import { HttpError } from './http.js';
import {
  appendEntry,
  clearEntries,
  deleteEntry,
  loadWorkbook,
  parseEntries,
  readFavorites,
  serialize,
  updateEntry,
} from './xlsx-sheet.js';

// Schreibender und lesender Zugriff auf die Excel-Datei in OneDrive. Die Ausgaben selbst liegen in der Datenbank
// (lib/expenses.js); die Datei wird nur beim Uebertragen, Leeren und Bereitstellen angefasst.
// `io` kapselt den Dateizugriff: { meta(), download(), upload(buf, etag) }.
// upload wirft PreconditionFailed, wenn die Datei seit dem Download veraendert wurde (HTTP 412 bei If-Match).
export class PreconditionFailed extends Error {}

const MAX_ATTEMPTS = 4;

// Metadaten zuerst, dann Inhalt: Der Inhalt ist damit mindestens so neu wie der eTag. Hat sich die Datei
// dazwischen geaendert, scheitert der Upload mit If-Match (412) und wird auf frischem Stand wiederholt.
async function snapshot(io) {
  const meta = await io.meta();
  const buf = await io.download();
  return { buf, meta };
}

const fileInfo = (m) => ({ name: m.name, etag: m.eTag, ctag: m.cTag, lastModified: m.lastModifiedDateTime, lastModifiedBy: m.lastModifiedBy?.user?.displayName });

// Fuehrt `mutate(wb)` auf frischem Stand aus und laedt mit If-Match hoch. Bei 412 wird auf frischen Daten
// wiederholt, damit nichts unbemerkt ueberschrieben wird. Mit `keepBuffer` enthaelt das Ergebnis die fertige Datei.
async function mutateWithRetry(io, mutate, { keepBuffer = false } = {}) {
  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    const { buf, meta } = await snapshot(io);
    const wb = await loadWorkbook(buf);
    const outcome = await mutate(wb);
    if (outcome.skipUpload) return { ...outcome.result, retries: attempt, file: fileInfo(meta), ...(keepBuffer ? { buffer: buf } : {}) };
    try {
      const out = await serialize(wb);
      const saved = await io.upload(out, meta.eTag);
      return { ...outcome.result, retries: attempt, file: fileInfo(saved), ...(keepBuffer ? { buffer: out } : {}) };
    } catch (e) {
      if (!(e instanceof PreconditionFailed)) throw e;
    }
  }
  throw new HttpError(409, 'busy', 'Speichern nach mehreren Versuchen nicht moeglich: Datei wird laufend veraendert.');
}

// Liest Liste und Favoriten aus der Datei. Nur fuer die einmalige Uebernahme des bisherigen Stands.
export async function readSheet(io) {
  const { buf } = await snapshot(io);
  const wb = await loadWorkbook(buf);
  return { rows: parseEntries(wb).entries, favorites: readFavorites(wb) };
}

// "In Excel uebertragen": Die Datei wird an die Datenbank angeglichen, soweit die App sie geschrieben hat.
// Neue Eintraege werden angehaengt, geaenderte (gleiche ID) ueberschrieben, in der App geloeschte (`deletedIds`)
// entfernt. Alles andere in der Datei, auch manuelle Aenderungen an unveraenderten Eintraegen, bleibt unberuehrt.
// Wiederholung ist sicher (IDs werden nie doppelt angelegt).
export async function syncEntries(io, { entries, deletedIds }, opts) {
  return mutateWithRetry(io, (wb) => {
    let removed = 0;
    for (const id of deletedIds) {
      const row = parseEntries(wb).entries.find((e) => e.id === id); // nach jedem Loeschen ruecken die Zeilen nach
      if (row) {
        deleteEntry(wb, row.row);
        removed++;
      }
    }
    const inExcel = new Map(parseEntries(wb).entries.map((e) => [e.id, e]));
    let added = 0;
    let updated = 0;
    for (const e of entries) {
      if (e.syncedHash && e.syncedHash === e.hash) continue; // schon uebertragen und seither unveraendert: Excel nicht anfassen
      const row = inExcel.get(e.id);
      if (!row) {
        appendEntry(wb, e);
        added++;
      } else if (row.hash !== e.hash) {
        updateEntry(wb, row.row, e);
        updated++;
      }
    }
    const result = { added, updated, removed, transferred: entries.map((e) => ({ id: e.id, hash: e.hash })), deletedIds };
    return added + updated + removed === 0 ? { skipUpload: true, result } : { result };
  }, opts);
}

// "Excel leeren" -> Datei-Seite: nur die Werte der Erfassungsliste, Zeilen, Stile und Merges bleiben.
export async function clearExcel(io) {
  return mutateWithRetry(io, (wb) => {
    const cleared = clearEntries(wb);
    return cleared === 0 ? { skipUpload: true, result: { cleared } } : { result: { cleared } };
  });
}
