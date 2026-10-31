import { buildLoginUrl } from '../../lib/auth.js';
import { method, redirect, wrap } from '../../lib/http.js';

export default wrap(async (req, res) => {
  method(req, 'GET');
  redirect(res, await buildLoginUrl());
});
