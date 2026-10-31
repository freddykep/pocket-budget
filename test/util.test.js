import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ApiError, buildSuggestions, displayAmount, errorMessage, findDuplicate, parseAmount, parseSpeech, parseSpeechMulti, totalExpenses } from '../public/util.js';

test('parseAmount: deutsche und englische Schreibweisen', () => {
  const ok = { '84,50': 84.5, '84.50': 84.5, '84': 84, '1.234,56': 1234.56, '1.234': 1234, '1.234.567': 1234567, '0,5': 0.5, ' 42 € ': 42, '65,00': 65 };
  for (const [input, expected] of Object.entries(ok)) assert.equal(parseAmount(input), expected, input);
});

test('parseAmount: Vorzeichen wird ignoriert (Ausgaben sind immer negativ gespeichert)', () => {
  assert.equal(parseAmount('-26'), 26);
  assert.equal(parseAmount('−5,5'), 5.5);
});

test('parseAmount: ungueltige Eingaben ergeben NaN', () => {
  for (const bad of ['', 'abc', '12,345', '1,2,3', '12..5', '€', null, undefined]) assert.ok(Number.isNaN(parseAmount(bad)), String(bad));
});

test('displayAmount und Summe', () => {
  assert.match(displayAmount(-84.5), /^84,50\s€$/);
  assert.match(displayAmount(12), /^\+12,00\s€$/);
  assert.equal(displayAmount(''), '');
  assert.equal(totalExpenses(-123.45), 123.45);
  assert.equal(totalExpenses(0), 0);
});

test('errorMessage: feste Texte, sonst Servertext', () => {
  assert.match(errorMessage(new ApiError('locked', 'x')), /gesperrt/);
  assert.match(errorMessage(new ApiError('offline', 'x')), /Internetverbindung/);
  assert.equal(errorMessage(new ApiError('bad_layout', 'Spalten fehlen: id')), 'Spalten fehlen: id');
  assert.match(errorMessage(new Error('boom')), /boom/);
  assert.match(errorMessage(undefined), /Unerwarteter Fehler/);
});


test('parseSpeech: Beispiele aus der Anforderung', () => {
  const pick = (t) => { const r = parseSpeech(t); return [r.ausgabe, r.ort, r.betrag]; };
  assert.deepEqual(pick('42 Euro im Restaurant'), ['Restaurant', '', 42]);
  assert.deepEqual(pick('Wocheneinkauf bei Rewe 84 Euro 50'), ['Wocheneinkauf', 'Rewe', 84.5]);
  assert.deepEqual(pick('65 Euro bei Aral getankt'), ['Tanken', 'Aral', 65]);
});

test('parseSpeech: weitere Formulierungen', () => {
  const pick = (t) => { const r = parseSpeech(t); return [r.ausgabe, r.ort, r.betrag]; };
  assert.deepEqual(pick('Kaffee bei Starbucks 3,50 Euro'), ['Kaffee', 'Starbucks', 3.5]);
  assert.deepEqual(pick('12 Komma 90 Euro Pizza'), ['Pizza', '', 12.9]);
  assert.deepEqual(pick('3 Euro 20 Cent Brötchen'), ['Brötchen', '', 3.2]);
  assert.deepEqual(pick('Ich habe 20 Euro für Blumen ausgegeben'), ['Blumen', '', 20]);
  assert.deepEqual(pick('Bahnticket 29,90'), ['Bahnticket', '', 29.9]);
});

test('parseSpeech: ohne Betrag bleibt der Betrag leer', () => {
  const r = parseSpeech('Bahnticket');
  assert.equal(r.betrag, null);
  assert.equal(r.ausgabe, 'Bahnticket');
  assert.equal(parseSpeech('').betrag, null);
});

test('parseSpeechMulti: mehrere Eintraege in einem Diktat', () => {
  const pick = (t) => parseSpeechMulti(t).entries.map((e) => [e.ausgabe, e.ort, e.betrag]);
  const two = [['Rewe', '', 15], ['Edeka', '', 10]];
  assert.deepEqual(pick('Rewe 15 Euro, Edeka 10 Euro'), two);
  assert.deepEqual(pick('Rewe 15 Euro Edeka 10 Euro'), two, 'ohne Trennzeichen');
  assert.deepEqual(pick('15 Euro Rewe 10 Euro Edeka'), two, 'Betrag zuerst');
  assert.deepEqual(pick('Rewe 15 Euro und Edeka 10 Euro'), two);
  assert.deepEqual(pick('Rewe 15, Edeka 10'), two, 'ohne Euro');
  assert.deepEqual(pick('Kaffee 3,50 Euro und Brötchen 2,20 Euro und Tanken 40 Euro'), [['Kaffee', '', 3.5], ['Brötchen', '', 2.2], ['Tanken', '', 40]]);
});

test('parseSpeechMulti: Einzeleintraege bleiben ein Eintrag', () => {
  const pick = (t) => parseSpeechMulti(t).entries.map((e) => [e.ausgabe, e.ort, e.betrag]);
  assert.deepEqual(pick('Wocheneinkauf bei Rewe 84 Euro 50'), [['Wocheneinkauf', 'Rewe', 84.5]]);
  assert.deepEqual(pick('Brot und Butter 5 Euro'), [['Brot und Butter', '', 5]]);
  assert.deepEqual(pick('12 Komma 90 Euro Pizza'), [['Pizza', '', 12.9]]);
  assert.deepEqual(pick('Rewe 3 Brötchen 15 Euro'), [['Rewe 3 Brötchen', '', 15]]);
  assert.deepEqual(pick('Bahnticket'), [['Bahnticket', '', null]]);
});

test('findDuplicate: gleicher Text, Ort und Betrag unter den letzten Eintraegen', () => {
  const rows = [{ ausgabe: 'Rewe', ort: '', betrag: -15 }, { ausgabe: 'Edeka', ort: 'Bonn', betrag: -10 }];
  assert.ok(findDuplicate(rows, { ausgabe: ' rewe ', ort: '', betrag: -15 }));
  assert.equal(findDuplicate(rows, { ausgabe: 'Rewe', ort: '', betrag: 15 }), null, 'Rueckerstattung ist keine Dublette');
  assert.equal(findDuplicate(rows, { ausgabe: 'Edeka', ort: '', betrag: -10 }), null, 'anderer Ort');
  assert.equal(findDuplicate(rows, { ausgabe: 'Rewe', ort: '', betrag: -15 }, 1), null, 'nur die letzten Eintraege zaehlen');
});

test('buildSuggestions: eindeutig, neueste zuerst, letzter Ort/Betrag je Text', () => {
  const rows = [
    { ausgabe: 'Kaffee', ort: 'Alt', betrag: -3 },
    { ausgabe: 'Rewe', ort: 'Bonn', betrag: -20 },
    { ausgabe: 'kaffee', ort: 'Neu', betrag: -4.5 },
  ];
  const s = buildSuggestions(rows);
  assert.deepEqual(s.ausgabe, ['kaffee', 'Rewe']);
  assert.deepEqual(s.ort, ['Neu', 'Bonn', 'Alt']);
  assert.deepEqual(s.last.get('kaffee'), { ort: 'Neu', betrag: 4.5 });
});
