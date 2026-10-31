import { disconnect } from '../../lib/auth.js';
import { assertSameOrigin, method, send, wrap } from '../../lib/http.js';
import { destroySession, requireUser } from '../../lib/session.js';

// Beendet Sitzung UND serverseitige Verbindung (Refresh Token wird geloescht).
export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const userId = await requireUser(req);
  await disconnect(userId);
  await destroySession(req, res);
  send(res, 200, { ok: true });
});
