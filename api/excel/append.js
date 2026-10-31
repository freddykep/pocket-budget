import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { appendRecord } from '../../lib/workbook-service.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const io = await requireFileIo(req);
  send(res, 200, await appendRecord(io, await readJson(req)));
});
