import { HttpError } from './http.js';
import { clearEntries, importOnce, listEntries, listTombstones, markTransferred, readState } from './expenses.js';
import { clearExcel, readSheet, syncEntries } from './workbook-service.js';

// Zusammenspiel von Datenbank (Ausgaben) und Excel-Datei. Die Datenbank ist fuehrend.

// Beim ersten Lesen: bisherigen Inhalt der Datei einmalig uebernehmen (siehe expenses.importOnce).
export const importFromExcelOnce = (vault, io) => importOnce(vault, () => readSheet(io));

// "In Excel uebertragen". Die Eintraege bleiben in der Datenbank. Mit `keepBuffer` kommt die fertige Datei mit zurueck.
export async function transferToExcel(vault, io, opts) {
  const [entries, deletedIds] = await Promise.all([listEntries(vault), listTombstones(vault)]);
  const rows = entries.map((e) => ({ id: e.id, ausgabe: e.ausgabe, ort: e.ort, betrag: e.betrag, hash: e.hash, syncedHash: e.syncedHash }));
  const result = await syncEntries(io, { entries: rows, deletedIds }, opts);
  await markTransferred(vault, result);
  const { transferred, ...out } = result;
  return { ...out, state: await readState(vault) };
}

// Excel bereitstellen: erst uebertragen, dann genau die Datei liefern, die jetzt in OneDrive liegt.
export async function exportExcel(vault, io) {
  const { buffer, file } = await transferToExcel(vault, io, { keepBuffer: true });
  return { name: file.name || 'Pocket-Budget.xlsx', buf: buffer };
}

// "Liste leeren": loescht die Ausgaben in der Datenbank UND in der Excel-Datei. Nur mit dem Stand, den der Benutzer
// bestaetigt hat. Erst die Datei (kann an einer Sperre scheitern), dann die Datenbank: Bei einem Fehler bleibt die App
// unveraendert, und ein zweiter Versuch ist sicher.
export async function clearAll(vault, io, { expectedListHash }) {
  if (!expectedListHash) throw new HttpError(400, 'bad_request', 'expectedListHash erforderlich');
  const before = await readState(vault);
  if (before.listHash !== expectedListHash) {
    throw new HttpError(409, 'conflict', 'Die Liste hat sich seit der Bestaetigung geaendert. Es wurde nichts geloescht.', { summary: before.summary });
  }
  const { cleared } = await clearExcel(io);
  await clearEntries(vault);
  return { cleared, summaryBefore: before.summary, state: await readState(vault) };
}
