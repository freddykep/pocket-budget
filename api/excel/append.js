import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { appendRecord, appendRecords } from '../../lib/workbook-service.js';

// Ein Eintrag ({ ausgabe, ... }) oder mehrere ({ records: [...] }). Eine gemeinsame Function, weil der
// Vercel-Hobby-Plan hoechstens 12 Functions pro Deployment erlaubt.
export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const io = await requireFileIo(req);
  const body = await readJson(req);
  send(res, 200, Array.isArray(body?.records) ? await appendRecords(io, body) : await appendRecord(io, body));
});
