import { assertSameOrigin, method, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { assignMissingIds } from '../../lib/workbook-service.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  send(res, 200, await assignMissingIds(await requireFileIo(req)));
});
