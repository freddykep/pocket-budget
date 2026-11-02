import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { requireVault } from '../../lib/vault.js';
import { clearAll } from '../../lib/transfer.js';

// Leert Datenbank und Excel-Datei.
export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const io = await requireFileIo(req); // prueft zuerst die Microsoft-Sitzung
  const vault = await requireVault(req);
  send(res, 200, await clearAll(vault, io, await readJson(req)));
});
