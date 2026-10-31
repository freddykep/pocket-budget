import { getOwnItemMeta } from '../../lib/graph.js';
import { HttpError, assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { requireUser } from '../../lib/session.js';
import { store } from '../../lib/store.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const userId = await requireUser(req);
  const { itemId } = await readJson(req);
  if (!/^[A-Za-z0-9!_-]{1,200}$/.test(String(itemId || ''))) throw new HttpError(400, 'bad_id', 'Ungueltige Datei-ID');

  // Die Pruefung laeuft mit dem Token des Benutzers: fremde oder nicht erlaubte Items liefern 403/404.
  const meta = await getOwnItemMeta(userId, itemId);
  if (!meta.file || !/\.xlsx$/i.test(meta.name)) throw new HttpError(422, 'not_xlsx', 'Bitte eine .xlsx-Datei waehlen');
  await store.set('files', userId, { driveId: meta.parentReference.driveId, itemId: meta.id, name: meta.name });
  send(res, 200, { name: meta.name, driveId: meta.parentReference.driveId, itemId: meta.id });
});
