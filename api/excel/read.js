import { method, send, sendFile, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { exportFile, readSheet } from '../../lib/workbook-service.js';

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// Standard: Inhalt als JSON. Mit ?download=1: unveraenderte Kopie der Datei zum Teilen (nichts wird gespeichert
// oder geloggt). Beides in einer Function, weil der Vercel-Hobby-Plan hoechstens 12 Functions erlaubt.
export default wrap(async (req, res) => {
  method(req, 'GET');
  const io = await requireFileIo(req);
  if (new URL(req.url, 'http://x').searchParams.get('download') === '1') {
    const { name, buf } = await exportFile(io);
    return sendFile(res, buf, name, XLSX);
  }
  send(res, 200, await readSheet(io));
});
