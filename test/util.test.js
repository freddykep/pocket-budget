import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ApiError, displayAmount, errorMessage, parseAmount, totalExpenses } from '../public/util.js';

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
