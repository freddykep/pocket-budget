import { config } from './config.js';

export class HttpError extends Error {
  constructor(status, code, message, extra = {}) {
    super(message);
    this.status = status;
    this.code = code;
    this.extra = extra;
  }
}

export function send(res, status, body) {
  res.statusCode = status;
  res.setHeader('content-type', 'application/json; charset=utf-8');
  res.setHeader('cache-control', 'no-store');
  res.end(JSON.stringify(body));
}

export function redirect(res, location) {
  res.statusCode = 302;
  res.setHeader('location', location);
  res.setHeader('cache-control', 'no-store');
  res.end();
}

export function parseCookies(req) {
  const out = {};
  for (const part of (req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}

export function setCookie(res, name, value, { maxAge } = {}) {
  const attrs = [`${name}=${encodeURIComponent(value)}`, 'Path=/', 'HttpOnly', 'SameSite=Lax'];
  if (maxAge !== undefined) attrs.push(`Max-Age=${maxAge}`);
  if (config.secureCookies) attrs.push('Secure');
  const prev = res.getHeader('set-cookie');
  res.setHeader('set-cookie', [...(prev ? [].concat(prev) : []), attrs.join('; ')]);
}

export async function readJson(req) {
  if (req.body !== undefined) return typeof req.body === 'string' ? JSON.parse(req.body || '{}') : req.body ?? {};
  const chunks = [];
  let size = 0;
  for await (const c of req) {
    size += c.length;
    if (size > 1_000_000) throw new HttpError(413, 'too_large', 'Request zu gross');
    chunks.push(c);
  }
  const text = Buffer.concat(chunks).toString('utf8');
  try {
    return text ? JSON.parse(text) : {};
  } catch {
    throw new HttpError(400, 'bad_json', 'Ungueltiges JSON');
  }
}

// CSRF-Schutz fuer zustandsaendernde Aufrufe: nur JSON, nur gleicher Origin.
export function assertSameOrigin(req) {
  if (!String(req.headers['content-type'] || '').startsWith('application/json')) {
    throw new HttpError(415, 'bad_content_type', 'Nur application/json erlaubt');
  }
  const origin = req.headers.origin;
  if (origin && new URL(origin).host !== req.headers.host) {
    throw new HttpError(403, 'bad_origin', 'Origin nicht erlaubt');
  }
}

export function method(req, allowed) {
  if (req.method !== allowed) throw new HttpError(405, 'method_not_allowed', `${allowed} erforderlich`);
}

export function wrap(handler) {
  return async (req, res) => {
    try {
      await handler(req, res);
    } catch (e) {
      if (e instanceof HttpError) {
        return send(res, e.status, { error: { code: e.code, message: e.message, ...e.extra } });
      }
      console.error(e);
      send(res, 500, { error: { code: 'internal', message: 'Interner Fehler' } });
    }
  };
}

// Binaerantwort als Download. Dateiname ASCII-sicher plus UTF-8-Variante (RFC 5987), damit Umlaute erhalten bleiben.
export function sendFile(res, buf, name, type) {
  const ascii = name.replace(/[^\x20-\x7e]|["\%]/g, '_');
  res.statusCode = 200;
  res.setHeader('content-type', type);
  res.setHeader('content-length', buf.length);
  res.setHeader('content-disposition', `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`);
  res.setHeader('cache-control', 'no-store');
  res.end(buf);
}
