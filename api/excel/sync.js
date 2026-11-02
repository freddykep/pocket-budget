import { assertSameOrigin, method, send, wrap } from '../../lib/http.js';
import { transferToExcel } from '../../lib/transfer.js';
import { requireFileIo } from '../../lib/selected.js';
import { requireVault } from '../../lib/vault.js';

// "In Excel uebertragen": schreibt neue, geaenderte und geloeschte Eintraege aus der Datenbank in die Excel-Datei.
export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const io = await requireFileIo(req); // prueft zuerst die Microsoft-Sitzung
  const vault = await requireVault(req);
  send(res, 200, await transferToExcel(vault, io));
});
