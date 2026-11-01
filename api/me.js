import { getMe } from '../lib/graph.js';
import { method, send, wrap } from '../lib/http.js';
import { requireUser } from '../lib/session.js';
import { store } from '../lib/store.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  const userId = await requireUser(req);
  // Konto (Microsoft) und gewaehlte Datei (Store) sind unabhaengig: gleichzeitig abfragen.
  const [me, file] = await Promise.all([getMe(userId), store.get('files', userId)]);
  send(res, 200, {
    displayName: me.displayName,
    account: me.mail || me.userPrincipalName,
    selectedFile: file ? { name: file.name } : null,
  });
});
