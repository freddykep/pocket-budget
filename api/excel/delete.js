import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { deleteEntry, readState } from '../../lib/expenses.js';
import { requireUser } from '../../lib/session.js';
import { requireVault } from '../../lib/vault.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  // Beides noetig: Microsoft-Sitzung (App-Zugang) und Tresor-Schluessel (Passwort). Erst die Sitzung: ohne sie
  // darf nicht erkennbar sein, ob zu einem Schluessel ein Tresor existiert.
  await requireUser(req);
  const vault = await requireVault(req);
  const r = await deleteEntry(vault, await readJson(req));
  send(res, 200, { ...r, state: await readState(vault) });
});
