import { fileIo } from './graph.js';
import { HttpError } from './http.js';
import { requireUser } from './session.js';
import { store } from './store.js';

// Liefert den Dateizugriff fuer die vom Benutzer gewaehlte Datei (aus dem Server-Store).
export async function requireFileIo(req) {
  const userId = await requireUser(req);
  const file = await store.get('files', userId);
  if (!file) throw new HttpError(409, 'no_file', 'Es ist noch keine Datei ausgewaehlt');
  return fileIo(userId, file);
}
