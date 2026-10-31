// Lokaler Entwicklungsserver. Auf Vercel uebernehmen die Functions in /api und der Ordner /public.
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = import.meta.dirname;
const port = Number(process.env.PORT) || 3000;

const server = createServer(async (req, res) => {
  const { pathname } = new URL(req.url, 'http://localhost');
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
  if (pathname === '/' || pathname === '/index.html') {
    res.setHeader('content-type', 'text/html; charset=utf-8');
    res.setHeader('cache-control', 'no-store');
    return res.end(await readFile(join(root, 'public', 'index.html')));
  }
  res.statusCode = 404;
  res.end('Not found');
});

server.listen(port, () => console.log(`PoC laeuft auf http://localhost:${port}`));
