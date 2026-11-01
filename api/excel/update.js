import { assertSameOrigin, method, readJson, send, wrap } from '../../lib/http.js';
import { requireFileIo } from '../../lib/selected.js';
import { saveFavorites, updateRecord } from '../../lib/workbook-service.js';

export default wrap(async (req, res) => {
  method(req, 'POST');
  assertSameOrigin(req);
  const io = await requireFileIo(req);
  const body = await readJson(req);
  // Ein Eintrag ({ id, ... }) oder die Favoritenliste ({ favorites }): gemeinsame Function wegen des 12-Functions-Limits.
  send(res, 200, Array.isArray(body?.favorites) ? await saveFavorites(io, body) : await updateRecord(io, body));
});
