import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { test } from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { decrypt, encrypt } from '../lib/crypto.js';
import {
  PreconditionFailed,
  appendRecord,
  assignMissingIds,
  clearList,
  readSheet,
  updateRecord,
} from '../lib/workbook-service.js';

process.env.TOKEN_ENC_KEY = Buffer.alloc(32, 7).toString('base64');

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

test('liest die echte Struktur: ID in A, Ausgabe in B:D (verbunden), Ort in E, Betrag in F', { skip }, async () => {
  const { io } = fakeDrive();
  const d = await readSheet(io);
  assert.equal(d.rows.length, 12);
  assert.equal(d.rows[0].ausgabe, 'Wietzorke');
  assert.equal(d.rows[0].ort, '', 'Ort " " gilt als leer');
  assert.equal(d.rows[0].betrag, -26);
  assert.equal(d.rows[11].id, '57583a13');
  assert.equal(d.rows[11].ausgabe, 'Neue Ausgabe');
  assert.equal(d.missingIds, 11);
  assert.equal(d.summary.count, 12);
  assert.equal(d.summary.total, -26 - 5.5 - 10.99 - 7.8 - 26 - 6.23 - 16.2 - 50.8 - 46 - 497.49 - 2.49 - 24.9);
});

test('append: neue Zeile erhaelt Stil, Merge B:D, ID; nichts anderes aendert sich', { skip }, async () => {
  const { io, state } = fakeDrive();
  const before = await entries(state.buf);
  const r = await appendRecord(io, { id: 'neu-0001-aaaa', ausgabe: 'Tom & Jerry <3 "x"', ort: 'REWE', betrag: '84,50' });
  assert.equal(r.row, 14);
  const after = await entries(state.buf);

  for (const name of Object.keys(before)) if (name !== SHEET) assert.equal(after[name], before[name], `${name} unveraendert`);
  for (let n = 1; n <= 13; n++) assert.equal(rowXml(after[SHEET], n), rowXml(before[SHEET], n), `Zeile ${n} byte-identisch`);

  const row14 = rowXml(after[SHEET], 14);
  assert.match(row14, /<c r="B14" s="5"/);
  assert.match(row14, /<c r="C14" s="6"/);
  assert.match(row14, /<c r="D14" s="7"/);
  assert.match(row14, /<c r="E14" s="3"/);
  assert.match(row14, /<c r="F14" s="4"><v>-84.5<\/v><\/c>/, 'Betrag als negative Zahl mit Zellformat');
  assert.deepEqual(merges(after[SHEET]), [...merges(before[SHEET]), 'B14:D14']);
  assert.match(after[SHEET], /<mergeCells count="13">/);
  assert.match(after[SHEET], /<dimension ref="A1:F14"\/>/);

  const d = await readSheet(io);
  const e = d.rows.at(-1);
  assert.deepEqual([e.id, e.ausgabe, e.ort, e.betrag], ['neu-0001-aaaa', 'Tom & Jerry <3 "x"', 'REWE', -84.5]);
});

test('Ausgabedatei ist eine gueltige Arbeitsmappe (ExcelJS oeffnet sie, Merges vorhanden)', { skip }, async () => {
  const { io, state } = fakeDrive();
  await appendRecord(io, { ausgabe: 'Test', ort: '', betrag: 1 });
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(state.buf);
  const ws = wb.worksheets[0];
  assert.equal(ws.getCell('B14').value, 'Test');
  assert.equal(ws.getCell('C14').isMerged, true);
  assert.equal(ws.getCell('B2').master.address, 'B2');
});

test('append ist idempotent: gleiche ID zweimal -> ein Eintrag', { skip }, async () => {
  const { io, state } = fakeDrive();
  await appendRecord(io, { id: 'dup-0001-aaaa', ausgabe: 'Kaffee', ort: '', betrag: 3 });
  const uploads = state.uploads;
  const again = await appendRecord(io, { id: 'dup-0001-aaaa', ausgabe: 'Kaffee', ort: '', betrag: 3 });
  assert.equal(again.duplicate, true);
  assert.equal(state.uploads, uploads, 'kein zweiter Upload');
  assert.equal((await readSheet(io)).rows.length, 13);
});

test('append: parallele Excel-Aenderung zwischen Lesen und Speichern geht nicht verloren', { skip }, async () => {
  const { io, state } = fakeDrive();
  state.beforeUpload = () => state.externalEdit((ws) => { ws.getCell('E2').value = 'Koeln'; });
  const r = await appendRecord(io, { ausgabe: 'Kaffee', ort: 'Bonn', betrag: 4 });
  assert.equal(r.retries, 1);
  const d = await readSheet(io);
  assert.equal(d.rows[0].ort, 'Koeln', 'fremde Aenderung blieb erhalten');
  assert.equal(d.rows.length, 13, 'neue Zeile ist vorhanden');
});

test('IDs ergaenzen: nur fehlende, vorhandene bleiben, zweiter Lauf aendert nichts', { skip }, async () => {
  const { io, state } = fakeDrive();
  const r = await assignMissingIds(io);
  assert.equal(r.assigned, 11);
  const d = await readSheet(io);
  assert.equal(d.missingIds, 0);
  assert.equal(d.rows[11].id, '57583a13', 'vorhandene ID unveraendert');
  assert.equal(new Set(d.rows.map((x) => x.id)).size, 12, 'alle eindeutig');
  assert.match(rowXml((await entries(state.buf))[SHEET], 2), /<c r="A2"[^>]*t="inlineStr"/);
  const uploads = state.uploads;
  assert.equal((await assignMissingIds(io)).assigned, 0);
  assert.equal(state.uploads, uploads);
});

test('update: aendert nur die Zellen der Zeile; Merge, Stil und ID bleiben', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assignMissingIds(io);
  const seen = (await readSheet(io)).rows[0];
  const before = await entries(state.buf);
  await updateRecord(io, { id: seen.id, expectedHash: seen.hash, values: { ausgabe: 'Wochenmarkt', ort: 'Bonn', betrag: 12.5 } });
  const after = await entries(state.buf);
  for (let n = 3; n <= 13; n++) assert.equal(rowXml(after[SHEET], n), rowXml(before[SHEET], n), `Zeile ${n} unveraendert`);
  assert.deepEqual(merges(after[SHEET]), merges(before[SHEET]));
  const row2 = rowXml(after[SHEET], 2);
  assert.match(row2, /<c r="F2" s="2"><v>-12.5<\/v>/);
  const e = (await readSheet(io)).rows[0];
  assert.deepEqual([e.id, e.ausgabe, e.ort, e.betrag], [seen.id, 'Wochenmarkt', 'Bonn', -12.5]);
});

test('update: Konflikt, wenn dieselbe Zeile in Excel geaendert wurde -> nichts ueberschrieben', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assignMissingIds(io);
  const seen = (await readSheet(io)).rows[0];
  await state.externalEdit((ws) => { ws.getCell('F2').value = -99; });
  const before = state.uploads;
  await assert.rejects(
    updateRecord(io, { id: seen.id, expectedHash: seen.hash, values: { ausgabe: 'x', ort: '', betrag: 1 } }),
    (e) => e.status === 409 && e.code === 'conflict' && e.extra.current.betrag === -99,
  );
  assert.equal(state.uploads, before);
});

test('update: Race zwischen Snapshot und Upload auf derselben Zeile wird als Konflikt erkannt', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assignMissingIds(io);
  const seen = (await readSheet(io)).rows[0];
  state.beforeUpload = () => state.externalEdit((ws) => { ws.getCell('F2').value = -77; });
  await assert.rejects(
    updateRecord(io, { id: seen.id, expectedHash: seen.hash, values: { ausgabe: 'x', ort: '', betrag: 1 } }),
    (e) => e.code === 'conflict',
  );
  assert.equal((await readSheet(io)).rows[0].betrag, -77);
});

test('update: Eintrag in Excel geloescht -> 404', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assignMissingIds(io);
  const seen = (await readSheet(io)).rows[0];
  await state.externalEdit((ws) => { ws.spliceRows(2, 1); });
  await assert.rejects(updateRecord(io, { id: seen.id, expectedHash: seen.hash, values: { ausgabe: 'x', ort: '', betrag: 1 } }), (e) => e.status === 404);
});

test('Excel leeren: nur Werte weg, Zeilen/Stile/Merges bleiben; naechster Eintrag landet in Zeile 2', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assignMissingIds(io);
  const d = await readSheet(io);
  const before = await entries(state.buf);
  const r = await clearList(io, { expectedListHash: d.listHash });
  assert.equal(r.cleared, 12);
  assert.equal(r.summaryBefore.count, 12);

  const after = await entries(state.buf);
  for (const name of Object.keys(before)) if (name !== SHEET) assert.equal(after[name], before[name]);
  assert.deepEqual(merges(after[SHEET]), merges(before[SHEET]), 'Merges unveraendert');
  assert.equal(rowXml(after[SHEET], 1), rowXml(before[SHEET], 1), 'Kopfzeile unveraendert');
  assert.match(rowXml(after[SHEET], 5), /<c r="B5" s="5"\/>/, 'Stil bleibt, Wert weg');
  assert.match(rowXml(after[SHEET], 5), /<c r="F5" s="4"\/>/);
  const empty = await readSheet(io);
  assert.equal(empty.rows.length, 0);
  assert.equal(empty.summary.total, 0);

  const a = await appendRecord(io, { ausgabe: 'Neu', ort: 'dm', betrag: 18.95 });
  assert.equal(a.row, 2);
  assert.equal(merges((await entries(state.buf))[SHEET]).length, 12, 'kein zusaetzlicher Merge');
});

test('Excel leeren: blockiert bei geaenderter Liste, fehlenden IDs und fremden Inhalten', { skip }, async () => {
  const { io, state } = fakeDrive();
  await assert.rejects(clearList(io, { expectedListHash: 'x' }), (e) => e.code === 'conflict');
  await assignMissingIds(io);
  const d = await readSheet(io);
  await assert.rejects(clearList(io, { expectedListHash: 'falsch' }), (e) => e.code === 'conflict');

  await state.externalEdit((ws) => { ws.getCell('H5').value = 'Notiz'; });
  const d2 = await readSheet(io);
  const before = state.uploads;
  await assert.rejects(clearList(io, { expectedListHash: d2.listHash }), (e) => e.code === 'not_a_list');
  assert.equal(state.uploads, before);
  assert.equal((await readSheet(io)).rows.length, d.rows.length);
});

test('Excel leeren: blockiert, wenn Eintraege ohne ID existieren', { skip }, async () => {
  const { io } = fakeDrive();
  const d = await readSheet(io);
  await assert.rejects(clearList(io, { expectedListHash: d.listHash }), (e) => e.code === 'missing_ids');
});

test('fehlende Pflichtspalten werden mit den gefundenen Ueberschriften gemeldet', async () => {
  const wb = new ExcelJS.Workbook(); wb.addWorksheet('S').addRow(['Foo', 'Bar']);
  const { io } = fakeDrive(Buffer.from(await wb.xlsx.writeBuffer()));
  await assert.rejects(readSheet(io), (e) => e.code === 'bad_layout' && /"Foo", "Bar"/.test(e.message));
});

test('Eingaben werden validiert', { skip }, async () => {
  const { io, state } = fakeDrive();
  for (const bad of [{ ausgabe: '', betrag: 1 }, { ausgabe: 'x', betrag: 'abc' }, { ausgabe: 'x', betrag: 0 }]) {
    await assert.rejects(appendRecord(io, bad), (e) => e.status === 422);
  }
  await assert.rejects(appendRecord(io, { id: 'a b', ausgabe: 'x', betrag: 1 }), (e) => e.code === 'bad_id');
  assert.equal(state.uploads, 0);
});

test('Token-Verschluesselung: Roundtrip und Manipulationserkennung', () => {
  const c = encrypt('refresh-token-geheim');
  assert.ok(!c.includes('geheim'));
  assert.equal(decrypt(c), 'refresh-token-geheim');
  const parts = c.split('.'); parts[3] = parts[3].slice(0, -2) + 'AA';
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

test('Layout mit ID hinten und Ausgabe in A:C: lesen, anhaengen, IDs ergaenzen, leeren', async () => {
  const { io, state } = fakeDrive(await backIdFixture());
  const d = await readSheet(io);
  assert.deepEqual(d.rows.map((r) => [r.ausgabe, r.ort, r.betrag]), [['Brot', '', -2.5], ['Bus', 'Hamburg', -3]]);

  await assignMissingIds(io);
  const a = await appendRecord(io, { ausgabe: 'Kaffee', ort: '', betrag: 4.2 });
  assert.equal(a.row, 4);
  const sheet = (await entries(state.buf))[SHEET];
  assert.ok(merges(sheet).includes('A4:C4'), 'neuer verbundener Bereich A4:C4');
  assert.match(rowXml(sheet, 4), /<c r="E4"[^>]*><v>-4.2<\/v>/);
  assert.match(rowXml(sheet, 4), /<c r="F4"[^>]*t="inlineStr"/, 'ID steht hinten in F');

  const d2 = await readSheet(io);
  await clearList(io, { expectedListHash: d2.listHash });
  const sheet2 = (await entries(state.buf))[SHEET];
  assert.deepEqual(merges(sheet2), merges(sheet));
  assert.equal((await readSheet(io)).rows.length, 0);
});
