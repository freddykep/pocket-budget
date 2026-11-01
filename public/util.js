// Reine Hilfsfunktionen (ohne DOM), damit sie auch in Node getestet werden koennen.

// Betrag aus Benutzereingabe. Ausgaben sind immer negativ gespeichert, deshalb wird ein eingegebenes
// Vorzeichen ignoriert. Gibt NaN zurueck, wenn die Eingabe nicht eindeutig lesbar ist.
export function parseAmount(input) {
  let s = String(input ?? '').trim().replace(/[\s€]/g, '');
  if (!s) return NaN;
  s = s.replace(/^[-+−]/, '');
  if (s.includes(',')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, ''); // 1.234 oder 1.234.567 = Tausenderpunkte (deutsche Schreibweise)
  }
  if (!/^\d+(\.\d{1,2})?$/.test(s)) return NaN;
  return Math.round(Number(s) * 100) / 100;
}

const euro = new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' });
export const formatEuro = (n) => euro.format(n);

// Gespeicherte Ausgaben sind negativ; angezeigt werden sie ohne Vorzeichen. Positive Werte (z. B. Erstattung) mit "+".
export function displayAmount(betrag) {
  if (typeof betrag !== 'number') return '';
  return betrag > 0 ? `+${formatEuro(betrag)}` : formatEuro(Math.abs(betrag));
}

// Summe aller Ausgaben als positive Zahl.
export const totalExpenses = (summaryTotal) => Math.round(-summaryTotal * 100) / 100 || 0;

const MESSAGES = {
  offline: 'Keine Internetverbindung. Bitte Verbindung prüfen und erneut versuchen.',
  timeout: 'Die Anfrage hat zu lange gedauert. Bitte erneut versuchen.',
  not_signed_in: 'Du bist nicht angemeldet.',
  reauth: 'Die Microsoft-Anmeldung ist abgelaufen. Bitte neu verbinden.',
  no_file: 'Es ist noch keine Excel-Datei ausgewählt.',
  forbidden: 'Microsoft verweigert den Zugriff auf diese Datei.',
  not_found: 'Die Excel-Datei wurde nicht gefunden. Wurde sie verschoben oder gelöscht?',
  locked: 'Die Excel-Datei ist gerade gesperrt, zum Beispiel weil sie in Excel geöffnet ist. Bitte in Excel schließen und erneut versuchen.',
  busy: 'Die Excel-Datei wird gerade verändert. Bitte gleich noch einmal versuchen.',
  conflict: 'Dieser Eintrag wurde in der Zwischenzeit in Excel geändert. Es wurde nichts überschrieben.',
  row_not_found: 'Dieser Eintrag existiert nicht mehr (in Excel gelöscht?).',
  missing_ids: 'Einige Einträge haben noch keine technische ID. Bitte zuerst die IDs ergänzen.',
  bad_value: null,
  bad_layout: null,
  not_a_list: null,
  not_allowed: null,
  graph_error: 'Microsoft meldet einen Fehler. Bitte später erneut versuchen.',
  internal: 'Unerwarteter Fehler. Bitte später erneut versuchen.',
};

export class ApiError extends Error {
  constructor(code, message, data) {
    super(message);
    this.code = code;
    this.data = data;
  }
}

// Verstaendliche Meldung; Codes ohne feste Meldung nutzen den (nicht sensiblen) Text des Servers.
export function errorMessage(err) {
  const fixed = MESSAGES[err?.code];
  return fixed || err?.message || MESSAGES.internal;
}

// ---- Spracheingabe: Diktat -> Ausgabe / Ort / Betrag ------------------------------------------
// Das Ergebnis ist nur ein Vorschlag; der Benutzer prueft und korrigiert ihn immer vor dem Speichern.

const PARTICIPLE = /^(ge[a-zäöüß]{3,}(?:t|en)|eingekauft|besorgt|bezahlt|ausgegeben)$/i;
const PARTICIPLE_TEXT = { getankt: 'Tanken', eingekauft: 'Einkauf', gekauft: 'Einkauf', gegessen: 'Essen', bezahlt: '', ausgegeben: '' };
const AMOUNT_PATTERNS = [
  [/(\d+)\s*komma\s*(\d{1,2})\s*(?:euro|€)/i, (m) => parseAmount(`${m[1]},${m[2]}`)],
  [/(\d+)\s*(?:euro|€)\s*(\d{1,2})(?!\d)(?:\s*cent)?/i, (m) => parseAmount(`${m[1]},${m[2].padEnd(2, '0')}`)],
  [/(\d[\d.,]*)\s*(?:euro|€)/i, (m) => parseAmount(m[1])],
  [/(?:^|\s)(\d+(?:[.,]\d{1,2})?)(?=\s|$)/, (m) => parseAmount(m[1])],
];
const upper = (s) => (s ? s[0].toLocaleUpperCase('de-DE') + s.slice(1) : s);
const participle = (word) => PARTICIPLE_TEXT[word.toLowerCase()] ?? word;

export function parseSpeech(input) {
  const transcript = String(input ?? '').trim();
  let text = ` ${transcript} `;
  let betrag = null;

  for (const [re, toAmount] of AMOUNT_PATTERNS) {
    const m = re.exec(text);
    if (!m) continue;
    const value = toAmount(m);
    if (Number.isNaN(value) || value <= 0) continue;
    betrag = value;
    text = text.replace(m[0], ' ');
    break;
  }

  text = text.replace(/\s+/g, ' ').trim();
  for (let i = 0; i < 3; i++) text = text.replace(/^(ich habe|ich hab|habe|hab|eine neue ausgabe|neue ausgabe|ausgabe|notiere|für|fuer|noch)\s+/i, '');
  text = text.replace(/\s+(ausgegeben|bezahlt)$/i, '').replace(/^(ausgegeben|bezahlt)$/i, '');

  let desc = '';
  let place = '';
  const prep = /^(.*?)\b(?:bei|beim|im|in der|in dem|in|vom|von|auf dem|am)\s+(.+)$/i.exec(text);
  if (prep) {
    const tokens = prep[2].split(' ');
    const idx = tokens.findIndex((t, i) => i > 0 && PARTICIPLE.test(t));
    place = (idx > 0 ? tokens.slice(0, idx) : tokens).join(' ');
    const extra = idx > 0 ? tokens.slice(idx).map(participle).filter(Boolean).join(' ') : '';
    desc = [prep[1].trim(), extra].filter(Boolean).join(' ');
  } else {
    desc = text.split(' ').map((t) => (PARTICIPLE.test(t) ? participle(t) : t)).filter(Boolean).join(' ');
  }
  if (!desc && place) {
    desc = place; // "42 Euro im Restaurant": das Restaurant ist die Ausgabe
    place = '';
  }
  return { ausgabe: upper(desc), ort: upper(place), betrag, transcript };
}

// ---- Mehrere Eintraege in einem Diktat --------------------------------------------------------
// "Rewe 15 Euro, Edeka 10 Euro" -> zwei Vorschlaege. Getrennt wird an Komma (nicht in Zahlen), Semikolon
// und Woertern wie "und"/"dann"; ohne Trennzeichen an den Betraegen: endet das Diktat jeweils mit dem Betrag
// ("Rewe 15 Euro Edeka 10 Euro") oder beginnt es damit ("15 Euro Rewe 10 Euro Edeka").

const MARKED_AMOUNT = /(\d+)\s*komma\s*(\d{1,2})\s*(?:euro|€)|(\d+)\s*(?:euro|€)\s*(\d{1,2})(?!\d)(?!\s*(?:euro|€))(?:\s*cent)?|(\d[\d.,]*)\s*(?:euro|€)/gi;
const BARE_AMOUNT = /(?<=^|\s)\d+(?:[.,]\d{1,2})?(?=\s|$|,)/g;
const SEPARATOR = /\s*(?:;|(?<!\d),|,(?!\d))\s*|\s+(und dann|und|dann|außerdem|ausserdem|sowie)\s+/gi;
const LEADING_FILLER = /^(?:(?:ich habe|ich hab|habe|hab|eine neue ausgabe|neue ausgabe|ausgabe|notiere|für|fuer|noch|erstens|zuerst)\s+)*/i;
const MAX_VOICE_ENTRIES = 10;

// Betragsstellen eines Textes; Betraege mit "Euro" haben Vorrang vor blossen Zahlen ("3 Brötchen 15 Euro").
function amountSpans(text) {
  const spans = (re) => [...text.matchAll(re)].map((m) => ({ start: m.index, end: m.index + m[0].length }));
  const marked = spans(MARKED_AMOUNT);
  return marked.length ? marked : spans(BARE_AMOUNT);
}

function splitPieces(text) {
  const pieces = [];
  const seps = [''];
  let last = 0;
  for (const m of text.matchAll(SEPARATOR)) {
    pieces.push(text.slice(last, m.index));
    seps.push(m[1] ? ` ${m[1]} ` : ' ');
    last = m.index + m[0].length;
  }
  pieces.push(text.slice(last));
  return { pieces, seps };
}

// Stuecke ohne Betrag gehoeren zum folgenden (am Ende: zum vorherigen) Stueck: "Brot und Butter 5 Euro".
function mergeAmountless(pieces, seps) {
  const out = [];
  let carry = '';
  pieces.forEach((p, i) => {
    const text = carry ? `${carry}${seps[i]}${p}` : p;
    if (amountSpans(text).length) { out.push(text); carry = ''; } else carry = text;
  });
  if (carry) {
    if (out.length) out[out.length - 1] += ` ${carry}`;
    else out.push(carry);
  }
  return out;
}

function splitAtAmounts(piece) {
  const spans = amountSpans(piece);
  if (spans.length < 2) return [piece];
  const startsWithAmount = !piece.slice(0, spans[0].start).replace(LEADING_FILLER, '').trim();
  const cuts = startsWithAmount ? spans.map((s) => s.start) : spans.map((s) => s.end);
  const parts = [];
  let from = 0;
  for (const cut of startsWithAmount ? cuts.slice(1) : cuts.slice(0, -1)) {
    parts.push(piece.slice(from, cut));
    from = cut;
  }
  parts.push(piece.slice(from));
  return parts;
}

export function parseSpeechMulti(input) {
  const transcript = String(input ?? '').trim();
  const { pieces, seps } = splitPieces(transcript);
  const segments = mergeAmountless(pieces, seps).flatMap(splitAtAmounts).slice(0, MAX_VOICE_ENTRIES);
  const entries = segments.map((s) => parseSpeech(s)).filter((e) => e.ausgabe || e.betrag !== null);
  return { entries: entries.length ? entries : [parseSpeech(transcript)], transcript };
}

// ---- Vorschlaege und Dublettenpruefung (nur aus den bereits geladenen Zeilen) ------------------

const norm = (s) => String(s ?? '').trim().toLocaleLowerCase('de-DE');

// Gleicher Text, Ort und Betrag unter den letzten Eintraegen. Die Excel-Datei hat kein Datum, deshalb zaehlt
// "zuletzt erfasst" statt "am selben Tag". `betrag` ist der Wert mit Vorzeichen, wie er gespeichert wuerde.
export function findDuplicate(rows, { ausgabe, ort, betrag }, recent = 20) {
  return rows.slice(-recent).find((r) => norm(r.ausgabe) === norm(ausgabe) && norm(r.ort) === norm(ort)
    && typeof r.betrag === 'number' && Math.abs(r.betrag - betrag) < 0.005) ?? null;
}

// Haeufig/zuletzt genutzte Texte und Orte (neueste zuerst) sowie der letzte Ort/Betrag je Ausgabentext.
export function buildSuggestions(rows, limit = 200) {
  const ausgabe = new Map();
  const ort = new Map();
  const last = new Map();
  for (const r of [...rows].reverse()) {
    const a = String(r.ausgabe ?? '').trim();
    const o = String(r.ort ?? '').trim();
    if (a && !ausgabe.has(norm(a))) ausgabe.set(norm(a), a);
    if (o && !ort.has(norm(o))) ort.set(norm(o), o);
    if (a && !last.has(norm(a)) && typeof r.betrag === 'number') last.set(norm(a), { ort: o, betrag: Math.abs(r.betrag) });
  }
  return { ausgabe: [...ausgabe.values()].slice(0, limit), ort: [...ort.values()].slice(0, limit), last };
}

// ---- Favoriten: manuell hinterlegte Vorlagen (Ausgabe, Ort, optional Betrag) -------------------

export const MAX_FAVORITES = 30;

// Prueft eine Favoriten-Eingabe. Liefert { value } oder { errors }. Betrag ist optional (leer = beim Erfassen eintippen).
export function validateFavorite({ ausgabe = '', ort = '', betrag = '' } = {}) {
  const errors = {};
  const a = String(ausgabe).trim();
  const o = String(ort).trim();
  const b = String(betrag).trim();
  if (!a) errors.ausgabe = 'Bitte eine Beschreibung eingeben.';
  else if (a.length > 200) errors.ausgabe = 'Maximal 200 Zeichen.';
  if (o.length > 100) errors.ort = 'Maximal 100 Zeichen.';
  let amount = '';
  if (b) {
    const n = parseAmount(b);
    if (Number.isNaN(n) || n <= 0) errors.betrag = 'Bitte einen Betrag wie 84,50 eingeben oder das Feld leer lassen.';
    else amount = n;
  }
  return Object.keys(errors).length ? { errors } : { value: { ausgabe: a, ort: o, betrag: amount } };
}

// Liest die Favoriten aus der Excel-Datei defensiv ein (von Hand bearbeitete Zeilen koennen unvollstaendig sein).
export function sanitizeFavorites(raw) {
  if (!Array.isArray(raw)) return [];
  const out = [];
  for (const f of raw) {
    if (!f || typeof f !== 'object') continue;
    const v = validateFavorite({ ausgabe: f.ausgabe, ort: f.ort, betrag: typeof f.betrag === 'number' ? String(f.betrag) : '' });
    if (v.value && !out.some((o) => sameFavorite(o, v.value))) out.push(v.value);
  }
  return out.slice(0, MAX_FAVORITES);
}

// Gleicher Favorit = gleiche Ausgabe und gleicher Ort (ohne Beachtung von Gross-/Kleinschreibung).
export function sameFavorite(a, b) {
  const k = (f) => `${f.ausgabe}\u0000${f.ort}`.toLocaleLowerCase('de-DE');
  return k(a) === k(b);
}
