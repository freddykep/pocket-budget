import { getMe } from '../lib/graph.js';
import { method, send, wrap } from '../lib/http.js';
import { requireUser } from '../lib/session.js';
import { store } from '../lib/store.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  const userId = await requireUser(req);
  const me = await getMe(userId);
  const file = await store.get('files', userId);
  send(res, 200, {
    displayName: me.displayName,
    account: me.mail || me.userPrincipalName,
    selectedFile: file ? { name: file.name } : null,
  });
});
