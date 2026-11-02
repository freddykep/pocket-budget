import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { randomBytes } from 'node:crypto';
import { join } from 'node:path';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { decrypt, encrypt } from '../lib/crypto.js';
import { addEntries, deleteEntry, listEntries, readFavorites, readState, saveFavorites, updateEntry } from '../lib/expenses.js';
import { store } from '../lib/store.js';
import { clearAll, exportExcel, importFromExcelOnce, transferToExcel } from '../lib/transfer.js';
import { VAULT_HEADER, createVault, requireVault, unlockVault, vaultFromKey, vaultSalt } from '../lib/vault.js';
import { PreconditionFailed, readSheet } from '../lib/workbook-service.js';

process.env.TOKEN_ENC_KEY = Buffer.alloc(32, 7).toString('base64');
process.env.DATA_DIR = mkdtempSync(join(tmpdir(), 'pocket-budget-test-'));

// Fixture: die echte Testdatei (verbundene Zellen B:D, ID in A, Ort E, Betrag F). Bleibt ausserhalb des Repos.
const FIXTURE = new URL('../testdata/test.xlsx', import.meta.url);
const skip = existsSync(FIXTURE) ? false : 'testdata/test.xlsx fehlt';
const original = () => (skip ? null : readFileSync(FIXTURE));

// Simuliert OneDrive: eTag-Zaehler, If-Match-Pruefung und Hook fuer parallele Aenderungen "in Excel".
function fakeDrive(buf = original()) {
  const state = { buf, v: 1, beforeUpload: null, uploads: 0 };
  const io = {
    meta: async () => ({ name: 'test.xlsx', eTag: `"v${state.v}"`, cTag: `c${state.v}`, lastModifiedDateTime: 'x' }),
    download: async () => Buffer.from(state.buf),
    upload: async (b, etag) => {
      if (state.beforeUpload) { const h = state.beforeUpload; state.beforeUpload = null; await h(); }
      if (etag !== `"v${state.v}"`) throw new PreconditionFailed();
      state.buf = b; state.v++; state.uploads++;
      return io.meta();
    },
  };
  // Unabhaengiger "Excel-Benutzer": ExcelJS schreibt die Datei komplett anders heraus als unsere Engine.
  state.externalEdit = async (fn) => {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(state.buf);
    fn(wb.worksheets[0]); state.buf = Buffer.from(await wb.xlsx.writeBuffer()); state.v++;
  };
  return { io, state };
}

const entries = async (zipBuf) => {
  const zip = await JSZip.loadAsync(zipBuf);
  const out = {};
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) out[name] = await zip.files[name].async('string');
  return out;
};
const rowXml = (sheet, n) => new RegExp(`<row r="${n}"[\\s\\S]*?</row>`).exec(sheet)?.[0];
const merges = (sheet) => [...sheet.matchAll(/<mergeCell ref="([^"]+)"/g)].map((m) => m[1]);
const SHEET = 'xl/worksheets/sheet1.xml';

let counter = 0;
const newKey = () => randomBytes(32).toString('base64url');
const newUser = () => vaultFromKey(newKey()); // Tresor { id, key }
const reqWith = (key) => ({ headers: { [VAULT_HEADER]: key } });
const add = (u, r) => addEntries(u, [r]);
const sync = (u, io, opts) => transferToExcel(u, io, opts);
const excelRows = async (io) => (await readSheet(io)).rows;

test('liest die Struktur der Datei (einmalige Uebernahme): ID in A, Ausgabe in B:D (verbunden), Ort in E, Betrag in F', { skip }, async () => {
  const { io } = fakeDrive();
  const d = await readSheet(io);
  assert.equal(d.rows.length, 12);
  assert.equal(d.rows[0].ausgabe, 'Wietzorke');
  assert.equal(d.rows[0].ort, '', 'Ort " " gilt als leer');
  assert.equal(d.rows[0].betrag, -26);
  assert.equal(d.rows[11].id, '57583a13');
  assert.equal(d.rows.filter((r) => r.id === '').length, 11);
});

test('Erfassen, Aendern, Loeschen und Favoriten beruehren die Excel-Datei nicht', { skip }, async () => {
  const { state } = fakeDrive();
  const u = newUser();
  await add(u, { id: 'erf-0001-aaaa', ausgabe: 'Brot', ort: 'Baecker', betrag: '2,50' });
  const [e] = await listEntries(u);
  await updateEntry(u, { id: e.id, expectedHash: e.hash, values: { ausgabe: 'Brot', ort: '', betrag: 3 } });
  await saveFavorites(u, { favorites: [{ ausgabe: 'Brot', ort: '', betrag: 3 }], expectedHash: (await readFavorites(u)).hash });
  const st = await readState(u);
  assert.equal(st.summary.count, 1);
  assert.equal(st.summary.total, -3);
  assert.equal(st.pending, 1);
  assert.equal(state.uploads, 0);
  assert.equal(state.v, 1);
});

test('Uebertragen: neue Zeile erhaelt Stil, Merge B:D, ID; nichts anderes aendert sich; Eintraege bleiben in der DB', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  const before = await entries(state.buf);
  await add(u, { id: 'neu-0001-aaaa', ausgabe: 'Tom & Jerry <3 "x"', ort: 'REWE', betrag: '84,50' });
  const r = await sync(u, io);
  assert.equal(r.added, 1);
  const after = await entries(state.buf);

  for (const name of Object.keys(before)) if (name !== SHEET) assert.equal(after[name], before[name], `${name} unveraendert`);
  for (let i = 1; i <= 13; i++) assert.equal(rowXml(after[SHEET], i), rowXml(before[SHEET], i), `Zeile ${i} byte-identisch`);
  const row14 = rowXml(after[SHEET], 14);
  assert.match(row14, /<c r="B14" s="5"/);
  assert.match(row14, /<c r="E14" s="3"/);
  assert.match(row14, /<c r="F14" s="4"><v>-84.5<\/v><\/c>/, 'Betrag als negative Zahl mit Zellformat');
  assert.deepEqual(merges(after[SHEET]), [...merges(before[SHEET]), 'B14:D14']);

  const e = (await excelRows(io)).at(-1);
  assert.deepEqual([e.id, e.ausgabe, e.ort, e.betrag], ['neu-0001-aaaa', 'Tom & Jerry <3 "x"', 'REWE', -84.5]);
  assert.equal((await listEntries(u)).length, 1, 'Eintrag bleibt in der DB');
  assert.equal(r.state.pending, 0);
  assert.ok(r.state.lastTransfer);

  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(state.buf);
  assert.equal(wb.worksheets[0].getCell('B14').value, 'Tom & Jerry <3 "x"');
  assert.equal(wb.worksheets[0].getCell('C14').isMerged, true);
});

test('Uebertragen ist wiederholbar: zweiter Lauf laedt nichts hoch, keine Dubletten', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await addEntries(u, [{ id: 'wdh-0001-aaaa', ausgabe: 'A', betrag: 1 }, { id: 'wdh-0002-aaaa', ausgabe: 'B', betrag: 2 }]);
  await sync(u, io);
  assert.equal(state.uploads, 1);
  const again = await sync(u, io);
  assert.equal(again.added + again.updated + again.removed, 0);
  assert.equal(state.uploads, 1);
  assert.equal((await excelRows(io)).length, 14);
});

test('Vorzeichen: Standard negativ, nur refund:true speichert positiv', async () => {
  const u = newUser();
  await add(u, { id: 'neg-0001-aaaa', ausgabe: 'A', betrag: '10' });
  await add(u, { id: 'neg-0002-aaaa', ausgabe: 'B', betrag: '-10' });
  await add(u, { id: 'pos-0001-aaaa', ausgabe: 'R', betrag: '10', refund: true });
  await add(u, { id: 'neg-0003-aaaa', ausgabe: 'C', betrag: '10', refund: 'true' });
  const by = Object.fromEntries((await listEntries(u)).map((e) => [e.id, e.betrag]));
  assert.deepEqual(by, { 'neg-0001-aaaa': -10, 'neg-0002-aaaa': -10, 'pos-0001-aaaa': 10, 'neg-0003-aaaa': -10 });
  const row = (await listEntries(u)).find((e) => e.id === 'pos-0001-aaaa');
  await updateEntry(u, { id: row.id, expectedHash: row.hash, values: { ausgabe: 'R', ort: '', betrag: 10 } });
  assert.equal((await listEntries(u)).find((e) => e.id === 'pos-0001-aaaa').betrag, -10);
});

test('Mehrere Eintraege: Reihenfolge bleibt, Wiederholung erzeugt keine Dubletten, ungueltige Eingaben speichern nichts', async () => {
  const u = newUser();
  const recs = [{ id: 'mul-0001-aaaa', ausgabe: 'Eins', betrag: 1 }, { id: 'mul-0002-aaaa', ausgabe: 'Zwei', betrag: 2, refund: true }];
  assert.equal((await addEntries(u, recs)).added, 2);
  assert.equal((await addEntries(u, recs)).added, 0);
  assert.deepEqual((await listEntries(u)).map((e) => e.ausgabe), ['Eins', 'Zwei']);
  for (const bad of [{ ausgabe: '', betrag: 1 }, { ausgabe: 'x', betrag: 'abc' }, { ausgabe: 'x', betrag: 0 }]) {
    await assert.rejects(add(u, bad), (e) => e.status === 422);
  }
  await assert.rejects(add(u, { id: 'a b', ausgabe: 'x', betrag: 1 }), (e) => e.code === 'bad_id');
  await assert.rejects(addEntries(u, [{ ausgabe: 'x', betrag: 1 }, { ausgabe: 'y', betrag: 0 }]), (e) => e.status === 422);
  assert.equal((await listEntries(u)).length, 2);
});

test('Aendern und Loeschen: Konflikt bei veraltetem Stand, Loeschen ist idempotent, fremde Nutzer sehen nichts', async () => {
  const u = newUser();
  await add(u, { id: 'chg-0001-aaaa', ausgabe: 'A', betrag: 1 });
  const [e] = await listEntries(u);
  await updateEntry(u, { id: e.id, expectedHash: e.hash, values: { ausgabe: 'B', betrag: 1 } });
  await assert.rejects(updateEntry(u, { id: e.id, expectedHash: e.hash, values: { ausgabe: 'C', betrag: 1 } }), (x) => x.code === 'conflict');
  await assert.rejects(deleteEntry(u, { id: e.id, expectedHash: e.hash }), (x) => x.code === 'conflict');
  await assert.rejects(updateEntry(u, { id: 'gibt-es-nicht', expectedHash: 'x', values: { ausgabe: 'C', betrag: 1 } }), (x) => x.code === 'row_not_found');
  assert.equal((await listEntries(newUser())).length, 0, 'anderer Benutzer sieht die Eintraege nicht');
  const now = (await listEntries(u))[0];
  assert.equal((await deleteEntry(u, { id: now.id, expectedHash: now.hash })).deleted, true);
  assert.equal((await deleteEntry(u, { id: now.id, expectedHash: now.hash })).deleted, false);
});

test('Nach dem Uebertragen: Aenderung und Loeschen wirken beim naechsten Uebertragen auch in Excel, fremde Zeilen bleiben', { skip }, async () => {
  const { io } = fakeDrive();
  const u = newUser();
  await addEntries(u, [{ id: 'syn-0001-aaaa', ausgabe: 'Eins', betrag: 1 }, { id: 'syn-0002-aaaa', ausgabe: 'Zwei', betrag: 2 }, { id: 'syn-0003-aaaa', ausgabe: 'Drei', betrag: 3 }]);
  await sync(u, io);
  const [e1, e2] = await listEntries(u);
  await updateEntry(u, { id: e1.id, expectedHash: e1.hash, values: { ausgabe: 'Eins neu', ort: 'Ort', betrag: 5 } });
  await deleteEntry(u, { id: e2.id, expectedHash: e2.hash });
  assert.equal((await readState(u)).pending, 2, 'eine Aenderung, ein Loeschen');
  const r = await sync(u, io);
  assert.deepEqual([r.added, r.updated, r.removed], [0, 1, 1]);
  const rows = await excelRows(io);
  assert.equal(rows.length, 14, '12 Zeilen aus der Datei + 2 Eintraege');
  assert.deepEqual(rows.slice(12).map((x) => [x.id, x.ausgabe, x.ort, x.betrag]), [['syn-0001-aaaa', 'Eins neu', 'Ort', -5], ['syn-0003-aaaa', 'Drei', '', -3]]);
  assert.equal(rows[0].ausgabe, 'Wietzorke', 'vorhandene Zeilen bleiben');
  assert.equal((await readState(u)).pending, 0);
});

test('Nie uebertragene Eintraege loeschen: kein Loesch-Vermerk, Excel bleibt unberuehrt', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { id: 'nie-0001-aaaa', ausgabe: 'A', betrag: 1 });
  const [e] = await listEntries(u);
  await deleteEntry(u, { id: e.id, expectedHash: e.hash });
  assert.equal((await readState(u)).pending, 0);
  await sync(u, io);
  assert.equal(state.uploads, 0);
});

test('Aenderungen in Excel werden ignoriert: die App zeigt weiter die DB', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { id: 'ign-0001-aaaa', ausgabe: 'Original', betrag: 1 });
  await sync(u, io);
  await state.externalEdit((ws) => { ws.getCell('B14').value = 'In Excel geaendert'; });
  const st = await readState(u);
  assert.equal(st.rows[0].ausgabe, 'Original');
  assert.equal(st.pending, 0);
  await add(u, { id: 'ign-0002-aaaa', ausgabe: 'Zweiter', betrag: 2 });
  await sync(u, io);
  assert.equal((await excelRows(io)).at(-2).ausgabe, 'In Excel geaendert', 'unveraenderte Eintraege werden nicht erneut geschrieben');
});

test('Uebertragen: parallele Excel-Aenderung zwischen Lesen und Speichern geht nicht verloren', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { ausgabe: 'Kaffee', ort: 'Bonn', betrag: 4 });
  state.beforeUpload = () => state.externalEdit((ws) => { ws.getCell('E2').value = 'Koeln'; });
  const r = await sync(u, io);
  assert.equal(r.retries, 1);
  const rows = await excelRows(io);
  assert.equal(rows[0].ort, 'Koeln', 'fremde Aenderung blieb erhalten');
  assert.equal(rows.length, 13);
});

test('Verschluesselt in der DB: weder Text, Ort, Betrag noch Favoriten sind im Klartext gespeichert', async () => {
  const u = newUser();
  await add(u, { id: 'enc-0001-aaaa', ausgabe: 'Geheimausgabe', ort: 'Geheimort', betrag: '123,45' });
  await saveFavorites(u, { favorites: [{ ausgabe: 'Geheimfavorit', ort: '', betrag: '' }], expectedHash: (await readFavorites(u)).hash });
  const raw = JSON.stringify([...(await store.list(`expenses:${u.id}`)), await store.get('favorites', u.id)]);
  for (const secret of ['Geheimausgabe', 'Geheimort', '123', 'Geheimfavorit']) assert.ok(!raw.includes(secret), `${secret} nicht im Klartext`);
  assert.equal((await listEntries(u))[0].ausgabe, 'Geheimausgabe', 'Lesen entschluesselt');
  const [row] = await store.list(`expenses:${u.id}`);
  row.value.d = `${row.value.d.slice(0, -3)}AAA`;
  await store.set(`expenses:${u.id}`, row.key, row.value);
  await assert.rejects(listEntries(u), 'manipulierte Daten werden erkannt');
});

test('Leeren: loescht DB und Excel (nur Werte), Favoriten bleiben, naechster Eintrag landet in Zeile 2', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { id: 'clr-0001-aaaa', ausgabe: 'Alt', betrag: 1 });
  await saveFavorites(u, { favorites: [{ ausgabe: 'Fav', ort: '', betrag: '' }], expectedHash: (await readFavorites(u)).hash });
  await sync(u, io);
  const before = await entries(state.buf);
  const st = await readState(u);
  await assert.rejects(clearAll(u, io, { expectedListHash: 'falsch' }), (e) => e.code === 'conflict');
  assert.equal((await listEntries(u)).length, 1, 'bei Konflikt nichts geloescht');

  const r = await clearAll(u, io, { expectedListHash: st.listHash });
  assert.equal(r.cleared, 13);
  assert.equal(r.summaryBefore.count, 1);
  assert.equal(r.state.summary.count, 0);
  assert.equal(r.state.favorites.items.length, 1);
  const after = await entries(state.buf);
  assert.deepEqual(merges(after[SHEET]), merges(before[SHEET]), 'Merges unveraendert');
  assert.equal(rowXml(after[SHEET], 1), rowXml(before[SHEET], 1), 'Kopfzeile unveraendert');
  assert.match(rowXml(after[SHEET], 5), /<c r="B5" s="5"\/>/, 'Stil bleibt, Wert weg');
  assert.equal((await excelRows(io)).length, 0);

  await add(u, { ausgabe: 'Neu', ort: 'dm', betrag: 18.95 });
  await sync(u, io);
  assert.equal((await excelRows(io))[0].ausgabe, 'Neu', 'landet in Zeile 2');
  assert.equal(merges((await entries(state.buf))[SHEET]).length, merges(before[SHEET]).length, 'kein zusaetzlicher Merge');
});

test('Leeren: bei fremden Inhalten in der Datei bricht es ab und laesst die DB unveraendert', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { ausgabe: 'Bleibt', betrag: 1 });
  await state.externalEdit((ws) => { ws.getCell('H5').value = 'Notiz'; });
  const st = await readState(u);
  const uploads = state.uploads;
  await assert.rejects(clearAll(u, io, { expectedListHash: st.listHash }), (e) => e.code === 'not_a_list');
  assert.equal(state.uploads, uploads);
  assert.equal((await listEntries(u)).length, 1);
});

test('Excel bereitstellen: uebertraegt zuerst und liefert genau die Datei aus OneDrive', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  await add(u, { ausgabe: 'Frisch', ort: '', betrag: 9 });
  const out = await exportExcel(u, io);
  assert.equal(out.name, 'test.xlsx');
  assert.ok(out.buf.equals(state.buf), 'identisch mit der Datei in OneDrive');
  assert.equal((await excelRows(io)).at(-1).ausgabe, 'Frisch');
  const again = await exportExcel(u, io);
  assert.ok(again.buf.equals(state.buf));
  assert.equal(state.uploads, 1, 'zweites Bereitstellen ohne Aenderung laedt nichts hoch');
});

test('Favoriten: speichern, Reihenfolge, Grenzen, Dubletten und Konflikt mit veraltetem Stand', async () => {
  const u = newUser();
  const fav = (ausgabe, ort = '', betrag = '') => ({ ausgabe, ort, betrag });
  const h0 = (await readFavorites(u)).hash;
  const r = await saveFavorites(u, { favorites: [fav('Brot', 'Baecker', 2.5), fav('Bus')], expectedHash: h0 });
  assert.deepEqual(r.favorites.items.map((f) => f.ausgabe), ['Brot', 'Bus']);
  await assert.rejects(saveFavorites(u, { favorites: [fav('X')], expectedHash: h0 }), (e) => e.code === 'conflict' && e.extra.favorites.items.length === 2);
  const h1 = r.favorites.hash;
  await assert.rejects(saveFavorites(u, { favorites: [fav('A'), fav('a')], expectedHash: h1 }), (e) => e.code === 'bad_value');
  await assert.rejects(saveFavorites(u, { favorites: [fav('')], expectedHash: h1 }), (e) => e.status === 422);
  await assert.rejects(saveFavorites(u, { favorites: Array.from({ length: 31 }, (_, i) => fav(`F${i}`)), expectedHash: h1 }), (e) => e.status === 400);
  assert.equal((await readFavorites(u)).items.length, 2);
  assert.equal((await readFavorites(newUser())).items.length, 0);
});

test('Einmalige Uebernahme: Eintraege mit ID aus der Datei, danach zaehlt nur die DB', { skip }, async () => {
  const { io, state } = fakeDrive();
  const u = newUser();
  assert.equal(await importFromExcelOnce(u, io), true);
  const st = await readState(u);
  assert.deepEqual(st.rows.map((r) => r.id), ['57583a13'], 'nur Zeilen mit ID');
  assert.equal(st.pending, 0, 'schon in Excel vorhanden');
  assert.equal(await importFromExcelOnce(u, io), false);
  await state.externalEdit((ws) => { ws.getCell('B14').value = 'x'; });
  assert.equal((await readState(u)).rows.length, 1);
});

test('Einmalige Uebernahme: scheitert das Lesen, wird es spaeter erneut versucht', { skip }, async () => {
  const { io } = fakeDrive();
  const u = newUser();
  const broken = { ...io, download: async () => { throw new Error('offline'); } };
  assert.equal(await importFromExcelOnce(u, broken), false);
  assert.equal(await importFromExcelOnce(u, io), true);
});

test('fehlende Pflichtspalten werden mit den gefundenen Ueberschriften gemeldet', async () => {
  const wb = new ExcelJS.Workbook(); wb.addWorksheet('S').addRow(['Foo', 'Bar']);
  const { io } = fakeDrive(Buffer.from(await wb.xlsx.writeBuffer()));
  await assert.rejects(readSheet(io), (e) => e.code === 'bad_layout' && /"Foo", "Bar"/.test(e.message));
});

test('Token-Verschluesselung: Roundtrip und Manipulationserkennung', () => {
  const c = encrypt('refresh-token-geheim');
  assert.ok(!c.includes('geheim'));
  assert.equal(decrypt(c), 'refresh-token-geheim');
  const parts = c.split('.'); parts[3] = `${parts[3].slice(0, -2)}AA`;
  assert.throws(() => decrypt(parts.join('.')));
});

// Variante wie in Daniels Datei: Ausgabe verbunden in A:C, Ort D, Betrag E, ID ganz hinten in F.
async function backIdFixture() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Liste');
  ws.addRow(['Ausgabe', null, null, 'Ort', 'Betrag', 'ID']);
  for (const [t, o, b] of [['Brot', '', -2.5], ['Bus', 'Hamburg', -3]]) {
    const r = ws.addRow([t, null, null, o, b]);
    ws.mergeCells(`A${r.number}:C${r.number}`);
    r.getCell(5).numFmt = '#,##0.00 "€"';
    r.getCell(1).border = { left: { style: 'thin' }, top: { style: 'thin' }, bottom: { style: 'thin' } };
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

test('Layout mit ID hinten und Ausgabe in A:C: uebertragen und leeren', async () => {
  const { io, state } = fakeDrive(await backIdFixture());
  const u = newUser();
  await add(u, { id: 'bak-0001-aaaa', ausgabe: 'Kaffee', ort: '', betrag: 4.2 });
  await sync(u, io);
  const sheet = (await entries(state.buf))[SHEET];
  assert.ok(merges(sheet).includes('A4:C4'), 'neuer verbundener Bereich A4:C4');
  assert.match(rowXml(sheet, 4), /<c r="E4"[^>]*><v>-4.2<\/v>/);
  assert.match(rowXml(sheet, 4), /<c r="F4"[^>]*t="inlineStr"/, 'ID steht hinten in F');

  await clearAll(u, io, { expectedListHash: (await readState(u)).listHash });
  assert.deepEqual(merges((await entries(state.buf))[SHEET]), merges(sheet));
  assert.equal((await readSheet(io)).rows.length, 0);
});

test('Aufrufe: Erfassen ohne Excel-Zugriff, Uebertragen = Metadaten, Inhalt, Upload; nichts zu tun = kein Upload', { skip }, async () => {
  const { io } = fakeDrive();
  const calls = [];
  const traced = {
    meta: async () => { calls.push('meta'); return io.meta(); },
    download: async () => { calls.push('download'); return io.download(); },
    upload: async (b, e) => { calls.push('upload'); return io.upload(b, e); },
  };
  const u = newUser();
  await add(u, { ausgabe: 'X', betrag: 1 });
  await readState(u);
  assert.deepEqual(calls, []);
  await sync(u, traced);
  assert.deepEqual(calls, ['meta', 'download', 'upload']);
  calls.length = 0;
  await sync(u, traced);
  assert.deepEqual(calls, ['meta', 'download']);
});

test('Tresor: anlegen, entsperren, doppeltes Anlegen und unbekanntes Passwort', async () => {
  const key = newKey();
  await assert.rejects(requireVault(reqWith(key)), (e) => e.code === 'vault_locked', 'ohne Tresor kein Zugriff');
  await assert.rejects(unlockVault(reqWith(key), 'ms:a'), (e) => e.code === 'vault_unknown', 'Entsperren legt nichts an');
  await assert.rejects(requireVault(reqWith(key)), (e) => e.code === 'vault_locked');
  const v = await createVault(reqWith(key));
  assert.equal((await requireVault(reqWith(key))).id, v.id);
  assert.equal((await unlockVault(reqWith(key), 'ms:a')).id, v.id);
  await assert.rejects(createVault(reqWith(key)), (e) => e.code === 'vault_exists');
  for (const bad of [undefined, '', 'zu-kurz', Buffer.alloc(16).toString('base64url')]) {
    await assert.rejects(requireVault(reqWith(bad)), (e) => e.code === 'vault_locked');
  }
  assert.match(vaultSalt(), /^[0-9a-f]{64}$/);
});

test('Tresor: jedes Passwort hat eigene Daten; mit dem Schluessel eines anderen ist nichts zu lesen', async () => {
  const keyA = newKey();
  const keyB = newKey();
  const a = await createVault(reqWith(keyA));
  const b = await createVault(reqWith(keyB));
  assert.notEqual(a.id, b.id);
  await add(a, { id: 'tre-0001-aaaa', ausgabe: 'Nur fuer A', betrag: 5 });
  assert.equal((await listEntries(b)).length, 0, 'B sieht nichts von A');
  assert.equal((await listEntries(a)).length, 1);
  // Selbst mit A's Datensatz unter B's Schluessel: nicht entschluesselbar.
  const [row] = await store.list(`expenses:${a.id}`);
  await store.set(`expenses:${b.id}`, row.key, row.value);
  await assert.rejects(listEntries(b));
  assert.equal(JSON.stringify(row.value).includes('Nur fuer A'), false);
});

test('Tresor: nach 5 Fehlversuchen wird gesperrt, richtiges Passwort setzt den Zaehler zurueck', async () => {
  const ok = newKey();
  await createVault(reqWith(ok));
  const user = 'ms:throttle';
  for (let i = 0; i < 5; i++) await assert.rejects(unlockVault(reqWith(newKey()), user), (e) => e.code === 'vault_unknown');
  await assert.rejects(unlockVault(reqWith(ok), user), (e) => e.code === 'vault_throttled', 'auch das richtige Passwort ist kurz gesperrt');
  await store.delete('attempts', user);
  await assert.rejects(unlockVault(reqWith(newKey()), user), (e) => e.code === 'vault_unknown');
  assert.ok(await unlockVault(reqWith(ok), user));
  assert.equal(await store.get('attempts', user), null);
});
