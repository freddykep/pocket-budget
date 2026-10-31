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
