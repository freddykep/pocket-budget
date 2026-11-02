import { method, send, sendFile, wrap } from '../../lib/http.js';
import { readState } from '../../lib/expenses.js';
import { requireFileIo } from '../../lib/selected.js';
import { requireVault } from '../../lib/vault.js';
import { exportExcel, importFromExcelOnce } from '../../lib/transfer.js';

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Standard: Stand aus der Datenbank als JSON. Mit ?download=1: Excel bereitstellen, d. h. erst in die Datei uebertragen,
// dann diese Datei liefern (nichts wird geloggt). Beides in einer Function, weil der Vercel-Hobby-Plan hoechstens
// 12 Functions erlaubt.
export default wrap(async (req, res) => {
  method(req, 'GET');
  const io = await requireFileIo(req); // prueft zuerst die Microsoft-Sitzung
  const vault = await requireVault(req);
  if (new URL(req.url, 'http://x').searchParams.get('download') === '1') {
    const { name, buf } = await exportExcel(vault, io);
    return sendFile(res, buf, name, XLSX);
  }
  await importFromExcelOnce(vault, io);
  send(res, 200, await readState(vault));
});
