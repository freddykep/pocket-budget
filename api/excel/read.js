import { method, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { readSheet } from '../../lib/workbook-service.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  send(res, 200, await readSheet(await requireFileIo(req)));
});
