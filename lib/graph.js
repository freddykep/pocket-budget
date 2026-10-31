import { config } from './config.js';
import { getAccessToken } from './auth.js';
import { HttpError } from './http.js';
import { PreconditionFailed } from './workbook-service.js';

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const META_FIELDS = 'id,name,eTag,cTag,size,lastModifiedDateTime,lastModifiedBy,file,folder,parentReference';

async function graph(userId, path, { method = 'GET', headers = {}, body } = {}) {
  const token = await getAccessToken(userId);
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(config.graphBase + path, {
      method,
      body,
      headers: { authorization: `Bearer ${token}`, ...headers },
    });
    if ((res.status === 429 || res.status === 503) && attempt < 3) {
      const wait = Number(res.headers.get('retry-after')) || 2 ** attempt;
      await new Promise((r) => setTimeout(r, Math.min(wait, 10) * 1000));
      continue;
    }
    return res;
  }
}

async function jsonOrThrow(res, what) {
  if (res.ok) return res.json();
  const body = await res.json().catch(() => ({}));
  const msg = body.error?.message || res.statusText;
  if (res.status === 401) throw new HttpError(401, 'reauth', 'Microsoft hat den Zugriff abgelehnt, bitte erneut anmelden');
  if (res.status === 403) throw new HttpError(403, 'forbidden', `${what}: Zugriff verweigert`);
  if (res.status === 404) throw new HttpError(404, 'not_found', `${what}: nicht gefunden`);
  throw new HttpError(502, 'graph_error', `${what}: ${msg}`, { graphCode: body.error?.code });
}

export async function getMe(userId) {
  return jsonOrThrow(await graph(userId, '/me?$select=displayName,mail,userPrincipalName'), 'Konto');
}

export async function listChildren(userId, folderId) {
  const base = folderId ? `/me/drive/items/${encodeURIComponent(folderId)}` : '/me/drive/root';
  const data = await jsonOrThrow(
    await graph(userId, `${base}/children?$select=id,name,folder,file,size,parentReference&$top=200`),
    'Ordner',
  );
  return data.value
    .map((i) => ({ id: i.id, name: i.name, isFolder: !!i.folder, isXlsx: !!i.file && /\.xlsx$/i.test(i.name) }))
    .filter((i) => i.isFolder || i.isXlsx);
}

// Metadaten der Datei im eigenen Drive des angemeldeten Benutzers (Graph erzwingt die Berechtigung).
export async function getOwnItemMeta(userId, itemId) {
  return jsonOrThrow(await graph(userId, `/me/drive/items/${encodeURIComponent(itemId)}?$select=${META_FIELDS}`), 'Datei');
}

// Zugriff auf die gespeicherte Datei; driveId/itemId stammen aus dem Server-Store, nie aus dem Request.
export function fileIo(userId, { driveId, itemId }) {
  const base = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
  return {
    async meta() {
      return jsonOrThrow(await graph(userId, `${base}?$select=${META_FIELDS}`), 'Datei');
    },
    async download() {
      const res = await graph(userId, `${base}/content`);
      if (!res.ok) await jsonOrThrow(res, 'Download');
      return Buffer.from(await res.arrayBuffer());
    },
    async upload(buf, etag) {
      const res = await graph(userId, `${base}/content`, {
        method: 'PUT',
        headers: { 'content-type': XLSX, 'if-match': etag },
        body: buf,
      });
      if (res.status === 412) throw new PreconditionFailed();
      if (res.status === 423) throw new HttpError(423, 'locked', 'Datei ist gesperrt (z. B. in Excel geoeffnet). Bitte spaeter erneut versuchen.');
      return jsonOrThrow(res, 'Upload');
    },
  };
}
