// Nur fuer lokale UI-Tests: serviert public/ und ersetzt Microsoft/OneDrive durch eine In-Memory-Kopie
// von testdata/test.xlsx. Datenbank (verschluesselte Datei in einem Temp-Ordner) und Excel-Service sind dieselben wie in Produktion.
// Start: node scripts/mock-server.js   (Port 3001)
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { createServer } from 'node:http';
import { extname, join } from 'node:path';
import { HttpError, readJson, send, sendFile, wrap } from '../lib/http.js';
import { addEntries, deleteEntry, readState, saveFavorites, updateEntry } from '../lib/expenses.js';
import { clearAll, exportExcel, importFromExcelOnce, transferToExcel } from '../lib/transfer.js';
import { PreconditionFailed } from '../lib/workbook-service.js';
import { createVault, requireVault, unlockVault, vaultSalt } from '../lib/vault.js';

process.env.DATA_DIR ??= mkdtempSync(join(tmpdir(), 'pocket-budget-mock-'));
process.env.TOKEN_ENC_KEY ??= Buffer.alloc(32, 7).toString('base64');
const USER = 'mock:user'; // Microsoft-Konto (Sitzung); der Tresor kommt aus dem Passwort im Browser
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
  '/api/me': async (req) => {
    if (req.method === 'POST') { const b = await readJson(req); await (b.action === 'create' ? createVault(req) : unlockVault(req, USER)); return { ok: true }; }
    return { displayName: 'Test Person', account: 'test@example.org', selectedFile: drive.hasFile ? { name: 'test.xlsx' } : null, vaultSalt: vaultSalt() };
  },
  '/api/files': async () => ({ items: [{ id: 'f1', name: 'PocketBudget-PoC', isFolder: true, isXlsx: false }, { id: 'x1', name: 'test.xlsx', isFolder: false, isXlsx: true }] }),
  '/api/file/select': async () => { drive.hasFile = true; return { name: 'test.xlsx' }; },
  '/api/excel/read': async (req) => { const v = await requireVault(req); await importFromExcelOnce(v, io); return readState(v); },
  '/api/excel/append': async (req) => { const v = await requireVault(req); const b = await readJson(req); const many = Array.isArray(b.records); const r = await addEntries(v, many ? b.records : [b]); const state = await readState(v); return many ? { ...r, state } : { id: r.ids[0], duplicate: r.added === 0, state }; },
  '/api/excel/update': async (req) => { const v = await requireVault(req); const b = await readJson(req); const r = Array.isArray(b.favorites) ? await saveFavorites(v, b) : await updateEntry(v, b); return { ...r, state: await readState(v) }; },
  '/api/excel/delete': async (req) => { const v = await requireVault(req); return { ...(await deleteEntry(v, await readJson(req))), state: await readState(v) }; },
  '/api/excel/clear': async (req) => clearAll(await requireVault(req), io, await readJson(req)),
  '/api/excel/sync': async (req) => transferToExcel(await requireVault(req), io),
  '/api/auth/logout': async () => { drive.signedIn = false; return { ok: true }; },
  // Testhilfe: simuliert eine Aenderung direkt in Excel (bumpt eTag) bzw. gesperrte Datei.
  '/mock/touch': async () => { drive.v++; return { ok: true }; },
};

createServer(wrap(async (req, res) => {
  const { pathname } = new URL(req.url, 'http://x');
  if (pathname === '/api/excel/read' && drive.signedIn && new URL(req.url, 'http://x').searchParams.get('download') === '1') {
    const f = await exportExcel(await requireVault(req), io);
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
