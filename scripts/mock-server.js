// Nur fuer lokale UI-Tests: serviert public/ und ersetzt Microsoft/OneDrive durch eine In-Memory-Kopie
// von testdata/test.xlsx. Der Excel-Service (lib/workbook-service.js) ist derselbe wie in Produktion.
// Start: node scripts/mock-server.js   (Port 3001)
import { readFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { extname, join } from 'node:path';
import { HttpError, readJson, send, sendFile, wrap } from '../lib/http.js';
import { PreconditionFailed, exportFile, appendRecord, appendRecords, assignMissingIds, clearList, deleteRecord, readSheet, saveFavorites, updateRecord } from '../lib/workbook-service.js';

const publicDir = join(import.meta.dirname, '..', 'public');
const fixture = process.env.FIXTURE || join(import.meta.dirname, '..', 'testdata', 'test.xlsx');
const TYPES = { '.html': 'text/html', '.css': 'text/css', '.js': 'text/javascript', '.webmanifest': 'application/manifest+json', '.svg': 'image/svg+xml', '.png': 'image/png' };

const drive = { buf: readFileSync(fixture), v: 1, signedIn: true, hasFile: true };
const io = {
  meta: async () => ({ name: 'test.xlsx', eTag: `"v${drive.v}"`, cTag: `c${drive.v}`, lastModifiedDateTime: new Date().toISOString() }),
  download: async () => Buffer.from(drive.buf),
  upload: async (b, etag) => {
    if (etag !== `"v${drive.v}"`) throw new PreconditionFailed();
    drive.buf = b; drive.v++;
    return io.meta();
  },
};

const routes = {
  '/api/me': async () => ({ displayName: 'Test Person', account: 'test@example.org', selectedFile: drive.hasFile ? { name: 'test.xlsx' } : null }),
  '/api/files': async () => ({ items: [{ id: 'f1', name: 'PocketBudget-PoC', isFolder: true, isXlsx: false }, { id: 'x1', name: 'test.xlsx', isFolder: false, isXlsx: true }] }),
  '/api/file/select': async () => { drive.hasFile = true; return { name: 'test.xlsx' }; },
  '/api/excel/read': () => readSheet(io),
  '/api/excel/append': async (req) => { const b = await readJson(req); return Array.isArray(b.records) ? appendRecords(io, b) : appendRecord(io, b); },
  '/api/excel/update': async (req) => { const b = await readJson(req); return Array.isArray(b.favorites) ? saveFavorites(io, b) : updateRecord(io, b); },
  '/api/excel/delete': async (req) => deleteRecord(io, await readJson(req)),
  '/api/excel/clear': async (req) => clearList(io, await readJson(req)),
  '/api/excel/assign-ids': () => assignMissingIds(io),
  '/api/auth/logout': async () => { drive.signedIn = false; return { ok: true }; },
  // Testhilfe: simuliert eine Aenderung direkt in Excel (bumpt eTag) bzw. gesperrte Datei.
  '/mock/touch': async () => { drive.v++; return { ok: true }; },
};

createServer(wrap(async (req, res) => {
  const { pathname } = new URL(req.url, 'http://x');
  if (pathname === '/api/excel/read' && drive.signedIn && new URL(req.url, 'http://x').searchParams.get('download') === '1') {
    const f = await exportFile(io);
    return sendFile(res, f.buf, f.name, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }
  const route = routes[pathname];
  if (route) {
    if (!drive.signedIn && pathname.startsWith('/api/') && pathname !== '/api/auth/logout') throw new HttpError(401, 'not_signed_in', 'Nicht angemeldet');
    return send(res, 200, await route(req));
  }
  if (pathname === '/api/auth/login') { drive.signedIn = true; res.statusCode = 302; res.setHeader('location', '/'); return res.end(); }
  const rel = pathname === '/' ? '/index.html' : pathname;
  try {
    const body = readFileSync(join(publicDir, rel));
    res.setHeader('content-type', TYPES[extname(rel)] || 'application/octet-stream');
    res.setHeader('cache-control', 'no-store');
    res.end(body);
  } catch {
    res.statusCode = 404;
    res.end('Not found');
  }
})).listen(3001, () => console.log('Mock-Server auf http://localhost:3001'));
