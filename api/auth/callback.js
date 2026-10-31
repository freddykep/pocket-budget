import { completeLogin } from '../../lib/auth.js';
import { HttpError, method, redirect, wrap } from '../../lib/http.js';
import { createSession } from '../../lib/session.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  const q = new URL(req.url, 'http://x').searchParams;
  if (q.get('error')) {
    // Z. B. Admin-Genehmigung erforderlich: genau das wollen wir im PoC sehen.
    const msg = `${q.get('error')}: ${(q.get('error_description') || '').split('\r\n')[0]}`.slice(0, 300);
    return redirect(res, `/?error=${encodeURIComponent(msg)}`);
  }
  if (!q.get('code')) throw new HttpError(400, 'no_code', 'Kein Code erhalten');
  const userId = await completeLogin(q.get('code'), q.get('state'));
  await createSession(res, userId);
  redirect(res, '/');
});
