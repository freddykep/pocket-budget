// Lokaler Entwicklungsserver. Auf Vercel uebernehmen die Functions in /api und der Ordner /public.
import { existsSync, statSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { extname, join, normalize } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = import.meta.dirname;
const publicDir = join(root, 'public');
const port = Number(process.env.PORT) || 3000;

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json',
  '.webmanifest': 'application/manifest+json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
};

// Entspricht den Headern in vercel.json.
const SECURITY_HEADERS = {
  'content-security-policy': "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'",
  'x-content-type-options': 'nosniff',
  'referrer-policy': 'same-origin',
};

const server = createServer(async (req, res) => {
  const { pathname } = new URL(req.url, 'http://localhost');
  for (const [k, v] of Object.entries(SECURITY_HEADERS)) res.setHeader(k, v);

  if (pathname.startsWith('/api/')) {
    // Nur Kleinbuchstaben, Schraegstriche und Bindestriche: kein Path Traversal.
    const file = join(root, `${pathname}.js`);
    if (!/^\/api\/[a-z/-]+$/.test(pathname) || !existsSync(file)) {
      res.statusCode = 404;
      return res.end('{"error":{"code":"not_found"}}');
    }
    const mod = await import(pathToFileURL(file));
    return mod.default(req, res);
  }

  const rel = pathname === '/' ? '/index.html' : pathname;
  const file = normalize(join(publicDir, rel));
  if (!file.startsWith(publicDir) || !/^[\w./-]+$/.test(rel) || !existsSync(file) || !statSync(file).isFile()) {
    res.statusCode = 404;
    return res.end('Not found');
  }
  res.setHeader('content-type', TYPES[extname(file)] || 'application/octet-stream');
  res.setHeader('cache-control', 'no-cache');
  res.end(await readFile(file));
});

server.listen(port, () => console.log(`PoC laeuft auf http://localhost:${port}`));
