import { listChildren } from '../lib/graph.js';
import { HttpError, method, send, wrap } from '../lib/http.js';
import { requireUser } from '../lib/session.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  const userId = await requireUser(req);
  const folder = new URL(req.url, 'http://x').searchParams.get('folder') || '';
  if (folder && !/^[A-Za-z0-9!_-]{1,200}$/.test(folder)) throw new HttpError(400, 'bad_id', 'Ungueltige Ordner-ID');
  send(res, 200, { items: await listChildren(userId, folder) });
});
