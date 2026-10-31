import { createHash, randomBytes } from 'node:crypto';
import { SCOPES, assertConfigured, config } from './config.js';
import { decrypt, encrypt } from './crypto.js';
import { HttpError } from './http.js';
import { store } from './store.js';

const b64u = (buf) => Buffer.from(buf).toString('base64url');
const tokenUrl = () => `${config.loginBase}/${config.authority}/oauth2/v2.0/token`;

// Access Tokens nur im Arbeitsspeicher; Refresh Tokens verschluesselt im Store.
const accessCache = new Map();
const inflight = new Map();

export async function buildLoginUrl() {
  assertConfigured();
  const state = b64u(randomBytes(24));
  const nonce = b64u(randomBytes(24));
  const verifier = b64u(randomBytes(48));
  const challenge = b64u(createHash('sha256').update(verifier).digest());
  await store.set('pending', state, { verifier, nonce, expires: Date.now() + 10 * 60 * 1000 });

  const url = new URL(`${config.loginBase}/${config.authority}/oauth2/v2.0/authorize`);
  url.search = new URLSearchParams({
    client_id: config.clientId,
    response_type: 'code',
    redirect_uri: config.redirectUri,
    response_mode: 'query',
    scope: SCOPES,
    state,
    nonce,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    prompt: 'select_account',
  }).toString();
  return url.toString();
}

async function tokenRequest(params) {
  const res = await fetch(tokenUrl(), {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      scope: SCOPES,
      ...params,
    }),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = new HttpError(res.status === 400 ? 401 : 502, body.error || 'token_error', body.error_description?.split('\r\n')[0] || 'Token-Anfrage fehlgeschlagen');
    err.oauthError = body.error;
    throw err;
  }
  return body;
}

function idTokenClaims(idToken) {
  // Der id_token kommt direkt vom Token-Endpunkt ueber TLS; wir pruefen aud und nonce.
  return JSON.parse(Buffer.from(idToken.split('.')[1], 'base64url').toString('utf8'));
}

export async function completeLogin(code, state) {
  const pending = state && (await store.get('pending', state));
  if (state) await store.delete('pending', state);
  if (!pending || pending.expires < Date.now()) throw new HttpError(400, 'bad_state', 'Ungueltiger oder abgelaufener Login-Vorgang');

  const tok = await tokenRequest({
    grant_type: 'authorization_code',
    code,
    redirect_uri: config.redirectUri,
    code_verifier: pending.verifier,
  });
  const claims = idTokenClaims(tok.id_token);
  if (claims.aud !== config.clientId || claims.nonce !== pending.nonce) {
    throw new HttpError(400, 'bad_id_token', 'id_token ungueltig');
  }
  if (!tok.refresh_token) throw new HttpError(502, 'no_refresh_token', 'Kein Refresh Token erhalten (offline_access?)');

  // Optionale Zugriffsliste: nur freigegebene Konten duerfen sich verbinden (nichts wird gespeichert).
  const account = String(claims.preferred_username || claims.email || '').toLowerCase();
  if (config.allowedAccounts.length && !config.allowedAccounts.includes(account)) {
    throw new HttpError(403, 'not_allowed', `Das Konto "${account}" ist fuer diese App nicht freigegeben.`);
  }

  const userId = `${claims.tid}:${claims.oid || claims.sub}`;
  await store.set('connections', userId, { refresh: encrypt(tok.refresh_token), updated: new Date().toISOString() });
  accessCache.set(userId, { token: tok.access_token, exp: Date.now() + tok.expires_in * 1000 });
  return userId;
}

export async function getAccessToken(userId) {
  const hit = accessCache.get(userId);
  if (hit && hit.exp - 60_000 > Date.now()) return hit.token;
  if (inflight.has(userId)) return inflight.get(userId);

  const p = (async () => {
    const conn = await store.get('connections', userId);
    if (!conn) throw new HttpError(401, 'reauth', 'Verbindung abgelaufen, bitte erneut anmelden');
    try {
      const tok = await tokenRequest({ grant_type: 'refresh_token', refresh_token: decrypt(conn.refresh) });
      // Refresh Tokens rotieren; das neue sofort speichern.
      if (tok.refresh_token) {
        await store.set('connections', userId, { refresh: encrypt(tok.refresh_token), updated: new Date().toISOString() });
      }
      accessCache.set(userId, { token: tok.access_token, exp: Date.now() + tok.expires_in * 1000 });
      return tok.access_token;
    } catch (e) {
      if (e.oauthError === 'invalid_grant') {
        await store.delete('connections', userId);
        accessCache.delete(userId);
        throw new HttpError(401, 'reauth', 'Verbindung wurde widerrufen oder ist abgelaufen, bitte erneut anmelden');
      }
      throw e;
    }
  })();
  inflight.set(userId, p);
  try {
    return await p;
  } finally {
    inflight.delete(userId);
  }
}

export async function disconnect(userId) {
  accessCache.delete(userId);
  await store.delete('connections', userId);
  await store.delete('files', userId);
}
