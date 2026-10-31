import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { HttpError } from './http.js';

// Bearbeitet die .xlsx als Archiv: nur betroffene Zeilen/Zellen im XML des ersten Blatts werden neu
// geschrieben. Alle anderen Archivteile (Stile, Spaltenbreiten, Designs, weitere Blaetter) und alle
// unveraenderten Zeilen bleiben unveraendert. Verbundene Zellen werden nie aufgeloest.
//
// Erwartetes Layout (Zeile 1 = Ueberschriften, Spalten nur anhand der Ueberschriften gefunden):
//   ID | Ausgabe (verbundener Bereich bis zur naechsten Spalte) | Ort (optional) | Betrag

const ALIASES = { id: 'id', ausgabe: 'ausgabe', ausgabentext: 'ausgabe', ort: 'ort', betrag: 'betrag' };
const TRACKED = ['id', 'ausgabe', 'ort', 'betrag'];

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&');
// eslint-disable-next-line no-control-regex
const stripInvalid = (s) => s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

export const colNum = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
export function colName(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function parseAttrs(raw) {
  const o = {};
  for (const m of raw.matchAll(/([\w:.-]+)="([^"]*)"/g)) o[m[1]] = m[2];
  return o;
}
const renderAttrs = (o) => Object.entries(o).map(([k, v]) => `${k}="${v}"`).join(' ');

function textOf(xml) {
  return [...xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => unesc(m[1])).join('');
}

function parseSst(xml) {
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g)].map((m) => (m[1] ? textOf(m[1]) : ''));
}

function parseCells(inner, rowNo) {
  const cells = [];
  for (const m of inner.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const attrs = parseAttrs(m[1]);
    const ref = /^([A-Z]+)(\d+)$/.exec(attrs.r || '');
    if (!ref || Number(ref[2]) !== rowNo) throw new HttpError(422, 'bad_xlsx', 'Unerwartete Zellstruktur im Tabellenblatt');
    cells.push({ col: colNum(ref[1]), attrs, inner: m[2] ?? '', raw: m[0], dirty: false });
  }
  return cells;
}

function parseSheet(xml) {
  const m = /<sheetData\s*>([\s\S]*?)<\/sheetData>/.exec(xml) || /<sheetData\s*\/>/.exec(xml);
  if (!m) throw new HttpError(422, 'bad_xlsx', 'Tabellenblatt hat keine Daten (sheetData fehlt)');
  const rows = [];
  for (const r of (m[1] ?? '').matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const attrs = parseAttrs(r[1]);
    const no = Number(attrs.r);
    rows.push({ r: no, attrsRaw: r[1], cells: parseCells(r[2] ?? '', no), raw: r[0], dirty: false });
  }
  const mergeBlock = /<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/.exec(xml);
  const merges = mergeBlock ? [...mergeBlock[1].matchAll(/<mergeCell\s+ref="([^"]+)"\s*\/>/g)].map((x) => x[1]) : [];
  return {
    prefix: xml.slice(0, m.index),
    suffix: xml.slice(m.index + m[0].length),
    rows,
    merges,
    addedMerges: [],
  };
}

export async function loadWorkbook(buf) {
  let zip;
  try {
    zip = await JSZip.loadAsync(buf);
  } catch {
    throw new HttpError(422, 'bad_xlsx', 'Datei ist keine lesbare .xlsx-Arbeitsmappe');
  }
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wbXml || !relsXml) throw new HttpError(422, 'bad_xlsx', 'Datei ist keine lesbare .xlsx-Arbeitsmappe');
  const sheet = /<sheet\b[^>]*?name="([^"]*)"[^>]*?r:id="([^"]+)"[^>]*\/>/.exec(wbXml);
  if (!sheet) throw new HttpError(422, 'bad_layout', 'Arbeitsmappe enthaelt kein Blatt');
  const rel = [...relsXml.matchAll(/<Relationship\b([^>]*)\/>/g)].map((x) => parseAttrs(x[1])).find((a) => a.Id === sheet[2]);
  if (!rel) throw new HttpError(422, 'bad_xlsx', 'Blatt-Verknuepfung fehlt');
  const path = `xl/${rel.Target.replace(/^\/?(xl\/)?/, '')}`;
  const xml = await zip.file(path)?.async('string');
  if (!xml) throw new HttpError(422, 'bad_xlsx', 'Tabellenblatt fehlt in der Datei');
  const sstXml = await zip.file('xl/sharedStrings.xml')?.async('string');
  return { zip, path, sheetName: unesc(sheet[1]), sst: sstXml ? parseSst(sstXml) : [], doc: parseSheet(xml) };
}

function cellValue(cell, sst) {
  const t = cell.attrs.t;
  const v = /<v>([\s\S]*?)<\/v>/.exec(cell.inner)?.[1];
  if (t === 'inlineStr') return textOf(cell.inner);
  if (v === undefined) return '';
  if (t === 's') return sst[Number(v)] ?? '';
  if (t === 'str') return unesc(v);
  if (t === 'b') return v === '1';
  if (t === 'e') return '';
  return v === '' ? '' : Number(v);
}

const hasFormula = (cell) => /<f\b/.test(cell.inner);
const isEmpty = (v) => v === '' || v == null || (typeof v === 'string' && v.trim() === '');
const getRow = (wb, r) => wb.doc.rows.find((x) => x.r === r);
const getCell = (row, col) => row?.cells.find((c) => c.col === col);

export function getLayout(wb) {
  const header = getRow(wb, 1);
  const cols = {};
  const found = [];
  for (const c of header?.cells ?? []) {
    const raw = String(cellValue(c, wb.sst));
    if (raw !== '') found.push(raw);
    const key = ALIASES[raw.trim().toLowerCase()];
    if (key && !cols[key]) cols[key] = c.col;
  }
  const missing = ['id', 'ausgabe', 'betrag'].filter((k) => !cols[k]);
  if (missing.length) {
    throw new HttpError(
      422,
      'bad_layout',
      `Spalten fehlen in Zeile 1 von Blatt "${wb.sheetName}": ${missing.join(', ')}. Gefundene Ueberschriften: ${found.map((f) => `"${f}"`).join(', ') || '(keine)'}`,
    );
  }
  // Der Ausgabe-Bereich reicht bis zur naechsten bekannten Spalte rechts davon (z. B. B:D bei Ort in E).
  const next = Math.min(...['id', 'ort', 'betrag'].map((k) => cols[k]).filter((n) => n > cols.ausgabe));
  return { cols, spanEnd: Number.isFinite(next) ? next - 1 : cols.ausgabe };
}

export function normalizeValues(input) {
  const ausgabe = stripInvalid(String(input.ausgabe ?? input.ausgabentext ?? '').trim());
  const ort = stripInvalid(String(input.ort ?? '').trim());
  if (!ausgabe) throw new HttpError(422, 'bad_value', 'Ausgabe darf nicht leer sein');
  if (ausgabe.length > 200 || ort.length > 100) throw new HttpError(422, 'bad_value', 'Text zu lang');
  const betrag = typeof input.betrag === 'number' ? input.betrag : Number(String(input.betrag ?? '').replace(',', '.'));
  if (!Number.isFinite(betrag)) throw new HttpError(422, 'bad_value', 'Betrag ist keine Zahl');
  return { ausgabe, ort, betrag: Math.round(betrag * 100) / 100 };
}

export function rowHash(r) {
  const b = typeof r.betrag === 'number' ? String(Math.round(r.betrag * 100) / 100) : '';
  return createHash('sha256').update(JSON.stringify([String(r.id ?? '').trim(), String(r.ausgabe ?? '').trim(), String(r.ort ?? '').trim(), b])).digest('hex').slice(0, 16);
}

const EMPTY_CELL = { attrs: {}, inner: '' };

function entryOf(wb, layout, row) {
  const val = (k) => (layout.cols[k] ? cellValue(getCell(row, layout.cols[k]) ?? EMPTY_CELL, wb.sst) : '');
  const betrag = val('betrag');
  return {
    row: row.r,
    id: String(val('id')).trim(),
    ausgabe: String(val('ausgabe')).trim(),
    ort: String(val('ort')).trim(),
    betrag: typeof betrag === 'number' ? betrag : '',
  };
}

const hasContent = (e) => e.id !== '' || e.ausgabe !== '' || e.ort !== '' || e.betrag !== '';

export function parseEntries(wb) {
  const layout = getLayout(wb);
  const entries = [];
  for (const row of wb.doc.rows) {
    if (row.r < 2) continue;
    const e = entryOf(wb, layout, row);
    if (!hasContent(e)) continue;
    e.hash = rowHash(e);
    entries.push(e);
  }
  return { sheet: wb.sheetName, entries, hasId: true };
}

// ---- Schreiben -------------------------------------------------------------------------------

function setCell(wb, row, col, value, styleFrom) {
  let cell = getCell(row, col);
  if (cell && hasFormula(cell)) {
    throw new HttpError(422, 'formula_cell', `Zelle ${colName(col)}${row.r} enthaelt eine Formel und wird nicht ueberschrieben`);
  }
  if (!cell) {
    const attrs = { r: `${colName(col)}${row.r}` };
    if (styleFrom?.attrs.s !== undefined) attrs.s = styleFrom.attrs.s;
    cell = { col, attrs, inner: '', raw: '', dirty: true };
    row.cells.push(cell);
    row.cells.sort((a, b) => a.col - b.col);
  }
  delete cell.attrs.t;
  if (value === '' || value == null) {
    cell.inner = '';
  } else if (typeof value === 'number') {
    cell.inner = `<v>${value}</v>`;
  } else {
    cell.attrs.t = 'inlineStr';
    cell.inner = `<is><t xml:space="preserve">${esc(stripInvalid(String(value)))}</t></is>`;
  }
  cell.dirty = true;
  row.dirty = true;
}

function ensureMerge(wb, layout, r) {
  if (layout.spanEnd <= layout.cols.ausgabe) return;
  const start = `${colName(layout.cols.ausgabe)}${r}:`;
  const all = [...wb.doc.merges, ...wb.doc.addedMerges];
  if (all.some((m) => m.startsWith(start))) return;
  wb.doc.addedMerges.push(`${start}${colName(layout.spanEnd)}${r}`);
}

function lastContentRow(wb, layout) {
  let last = 1;
  for (const row of wb.doc.rows) if (row.r >= 2 && hasContent(entryOf(wb, layout, row))) last = Math.max(last, row.r);
  return last;
}

export function appendEntry(wb, { id, ausgabe, ort, betrag }) {
  const layout = getLayout(wb);
  const last = lastContentRow(wb, layout);
  const target = last + 1;
  let row = getRow(wb, target);
  const template = last >= 2 ? getRow(wb, last) : null;
  if (!row) {
    // Neue Zeile mit Zeilenformat und Zellstilen der letzten Datenzeile.
    row = { r: target, attrsRaw: (template?.attrsRaw ?? ` r="${target}"`).replace(/\br="\d+"/, `r="${target}"`).replace(/\s*spans="[^"]*"/, ''), cells: [], raw: '', dirty: true };
    for (const tc of template?.cells ?? []) {
      row.cells.push({ col: tc.col, attrs: { r: `${colName(tc.col)}${target}`, ...(tc.attrs.s !== undefined ? { s: tc.attrs.s } : {}) }, inner: '', raw: '', dirty: true });
    }
    const idx = wb.doc.rows.findIndex((x) => x.r > target);
    wb.doc.rows.splice(idx === -1 ? wb.doc.rows.length : idx, 0, row);
  }
  const styleFrom = (col) => getCell(template, col);
  setCell(wb, row, layout.cols.id, id, styleFrom(layout.cols.id));
  setCell(wb, row, layout.cols.ausgabe, ausgabe, styleFrom(layout.cols.ausgabe));
  if (layout.cols.ort) setCell(wb, row, layout.cols.ort, ort, styleFrom(layout.cols.ort));
  setCell(wb, row, layout.cols.betrag, betrag, styleFrom(layout.cols.betrag));
  ensureMerge(wb, layout, target);
  return target;
}

export function updateEntry(wb, rowNo, { ausgabe, ort, betrag }) {
  const layout = getLayout(wb);
  const row = getRow(wb, rowNo);
  setCell(wb, row, layout.cols.ausgabe, ausgabe);
  if (layout.cols.ort) setCell(wb, row, layout.cols.ort, ort);
  setCell(wb, row, layout.cols.betrag, betrag);
}

// Vergibt IDs nur an Eintraege ohne ID. Vorhandene IDs bleiben unveraendert.
export function assignIds(wb, makeId) {
  const layout = getLayout(wb);
  let n = 0;
  for (const row of wb.doc.rows) {
    if (row.r < 2) continue;
    const e = entryOf(wb, layout, row);
    if (hasContent(e) && e.id === '') {
      setCell(wb, row, layout.cols.id, makeId());
      n++;
    }
  }
  return n;
}

// Prueft, ob die Datei sicher als Erfassungsliste behandelt werden kann (Voraussetzung fuer "Excel leeren").
export function assertErfassungsliste(wb) {
  const layout = getLayout(wb);
  const known = new Set([layout.cols.id, layout.cols.ort, layout.cols.betrag].filter(Boolean));
  for (let c = layout.cols.ausgabe; c <= layout.spanEnd; c++) known.add(c);
  for (const row of wb.doc.rows) {
    if (row.r < 2) continue;
    for (const cell of row.cells) {
      if (known.has(cell.col)) {
        if (hasFormula(cell)) throw new HttpError(422, 'not_a_list', `Zelle ${colName(cell.col)}${row.r} enthaelt eine Formel. Die Datei wird nicht als reine Erfassungsliste behandelt.`);
      } else if (hasFormula(cell) || !isEmpty(cellValue(cell, wb.sst))) {
        throw new HttpError(422, 'not_a_list', `Zelle ${colName(cell.col)}${row.r} enthaelt weitere Inhalte neben der Erfassungsliste.`);
      }
    }
  }
  if (wb.zip.file(/^xl\/worksheets\/sheet\d+\.xml$/).length > 1) {
    throw new HttpError(422, 'not_a_list', 'Die Arbeitsmappe enthaelt mehrere Blaetter. Die Datei wird nicht als reine Erfassungsliste behandelt.');
  }
  return layout;
}

// Leert nur die Werte der Datenzeilen (ID, Ausgabe, Ort, Betrag). Zeilen, Stile und verbundene Zellen bleiben.
export function clearEntries(wb) {
  const layout = assertErfassungsliste(wb);
  let count = 0;
  for (const row of wb.doc.rows) {
    if (row.r < 2) continue;
    if (!hasContent(entryOf(wb, layout, row))) continue;
    count++;
    for (const key of TRACKED) {
      if (layout.cols[key] && getCell(row, layout.cols[key])) setCell(wb, row, layout.cols[key], '');
    }
  }
  return count;
}

// ---- Serialisierung --------------------------------------------------------------------------

function renderCell(c) {
  const a = renderAttrs(c.attrs);
  return c.inner ? `<c ${a}>${c.inner}</c>` : `<c ${a}/>`;
}

function renderRow(row) {
  if (!row.dirty) return row.raw;
  const inner = row.cells.map((c) => (c.dirty ? renderCell(c) : c.raw)).join('');
  return inner ? `<row${row.attrsRaw}>${inner}</row>` : `<row${row.attrsRaw}/>`;
}

const AFTER_MERGE = ['phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'tableParts', 'extLst'];

export async function serialize(wb) {
  const { doc } = wb;
  let prefix = doc.prefix;
  let suffix = doc.suffix;

  // Ausdehnung des Datenbereichs nachziehen, falls Zeilen hinzugekommen sind.
  const dim = /<dimension\s+ref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"\s*\/>/.exec(prefix);
  if (dim) {
    const maxRow = Math.max(...doc.rows.map((r) => r.r), Number(dim[4] ?? dim[2]));
    const maxCol = Math.max(...doc.rows.flatMap((r) => r.cells.map((c) => c.col)), colNum(dim[3] ?? dim[1]));
    const ref = `${dim[1]}${dim[2]}:${colName(maxCol)}${maxRow}`;
    if (ref !== `${dim[1]}${dim[2]}:${dim[3] ?? dim[1]}${dim[4] ?? dim[2]}`) prefix = prefix.replace(dim[0], `<dimension ref="${ref}"/>`);
  }

  if (doc.addedMerges.length) {
    const added = doc.addedMerges.map((m) => `<mergeCell ref="${m}"/>`).join('');
    const total = doc.merges.length + doc.addedMerges.length;
    const block = /<mergeCells\b[^>]*>([\s\S]*?)<\/mergeCells>/.exec(suffix);
    if (block) {
      suffix = suffix.replace(block[0], `<mergeCells count="${total}">${block[1]}${added}</mergeCells>`);
    } else {
      const tag = new RegExp(`<(${AFTER_MERGE.join('|')})\\b`).exec(suffix);
      const at = tag ? tag.index : suffix.lastIndexOf('</worksheet>');
      suffix = `${suffix.slice(0, at)}<mergeCells count="${total}">${added}</mergeCells>${suffix.slice(at)}`;
    }
  }

  const xml = `${prefix}<sheetData>${doc.rows.map(renderRow).join('')}</sheetData>${suffix}`;
  wb.zip.file(wb.path, xml);
  return Buffer.from(await wb.zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}
