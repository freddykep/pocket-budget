import { getMe } from '../lib/graph.js';
import { assertSameOrigin, method, readJson, send, wrap } from '../lib/http.js';
import { requireUser } from '../lib/session.js';
import { store } from '../lib/store.js';
import { createVault, unlockVault, vaultSalt } from '../lib/vault.js';

// GET: Konto und gewaehlte Datei. POST { action: 'create' | 'unlock' } mit Schluessel im Header: Tresor anlegen bzw.
// pruefen (gemeinsame Function wegen des 12-Functions-Limits von Vercel Hobby).
export default wrap(async (req, res) => {
  if (req.method === 'POST') {
    assertSameOrigin(req);
    const userId = await requireUser(req);
    const body = await readJson(req);
    if (body?.action === 'create') await createVault(req);
    else await unlockVault(req, userId);
    return send(res, 200, { ok: true });
  }
  method(req, 'GET');
  const userId = await requireUser(req);
  // Konto (Microsoft) und gewaehlte Datei (Store) sind unabhaengig: gleichzeitig abfragen.
  const [me, file] = await Promise.all([getMe(userId), store.get('files', userId)]);
  send(res, 200, {
    displayName: me.displayName,
    account: me.mail || me.userPrincipalName,
    selectedFile: file ? { name: file.name } : null,
    vaultSalt: vaultSalt(),
  });
});
