import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { readState, saveFavorites, updateEntry } from '../../lib/expenses.js';
import { requireUser } from '../../lib/session.js';
import { requireVault } from '../../lib/vault.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  // Beides noetig: Microsoft-Sitzung (App-Zugang) und Tresor-Schluessel (Passwort). Erst die Sitzung: ohne sie
  // darf nicht erkennbar sein, ob zu einem Schluessel ein Tresor existiert.
  await requireUser(req);
  const vault = await requireVault(req);
  const body = await readJson(req);
  // Ein Eintrag ({ id, ... }) oder die Favoritenliste ({ favorites }): gemeinsame Function wegen des 12-Functions-Limits.
  const r = Array.isArray(body?.favorites) ? await saveFavorites(vault, body) : await updateEntry(vault, body);
  send(res, 200, { ...r, state: await readState(vault) });
});
