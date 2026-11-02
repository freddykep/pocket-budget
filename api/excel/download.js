import { method, sendFile, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { exportFile } from '../../lib/workbook-service.js';

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Unveraenderte Kopie der verknuepften Excel-Datei zum Teilen. Nichts wird gespeichert oder geloggt.
export default wrap(async (req, res) => {
  method(req, 'GET');
  const { name, buf } = await exportFile(await requireFileIo(req));
  sendFile(res, buf, name, XLSX);
});
