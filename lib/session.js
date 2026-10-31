import { randomBytes } from 'node:crypto';
import { HttpError, parseCookies, setCookie } from './http.js';
import { store } from './store.js';

const COOKIE = 'pb_sid';
const TTL_SECONDS = 7 * 24 * 3600;

export async function createSession(res, userId) {
  const sid = randomBytes(32).toString('base64url');
  await store.set('sessions', sid, { userId, expires: Date.now() + TTL_SECONDS * 1000 });
  setCookie(res, COOKIE, sid, { maxAge: TTL_SECONDS });
}

export async function destroySession(req, res) {
  const sid = parseCookies(req)[COOKIE];
  if (sid) await store.delete('sessions', sid);
  setCookie(res, COOKIE, '', { maxAge: 0 });
}

// Der Benutzer wird ausschliesslich aus der Server-Session bestimmt, nie aus Request-Parametern.
export async function requireUser(req) {
  const sid = parseCookies(req)[COOKIE];
  const s = sid && (await store.get('sessions', sid));
  if (!s || s.expires < Date.now()) {
    if (s) await store.delete('sessions', sid);
    throw new HttpError(401, 'not_signed_in', 'Nicht angemeldet');
  }
  return s.userId;
}
