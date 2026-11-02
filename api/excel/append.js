import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { addEntries, readState } from '../../lib/expenses.js';
import { requireUser } from '../../lib/session.js';
import { requireVault } from '../../lib/vault.js';

// Ein Eintrag ({ ausgabe, ... }) oder mehrere ({ records: [...] }). Eine gemeinsame Function, weil der
// Vercel-Hobby-Plan hoechstens 12 Functions pro Deployment erlaubt.
export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  // Beides noetig: Microsoft-Sitzung (App-Zugang) und Tresor-Schluessel (Passwort). Erst die Sitzung: ohne sie
  // darf nicht erkennbar sein, ob zu einem Schluessel ein Tresor existiert.
  await requireUser(req);
  const vault = await requireVault(req);
  const body = await readJson(req);
  const many = Array.isArray(body?.records);
  const r = await addEntries(vault, many ? body.records : [body]);
  const state = await readState(vault);
  send(res, 200, many ? { ...r, state } : { id: r.ids[0], duplicate: r.added === 0, state });
});
