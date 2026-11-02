import { createHash, createHmac, timingSafeEqual } from 'node:crypto';
import { config } from './config.js';
import { decrypt, encrypt } from './crypto.js';
import { HttpError } from './http.js';
import { store } from './store.js';

// Tresor des Benutzers: Der Browser leitet aus dem Passwort einen 32-Byte-Schluessel ab (PBKDF2, siehe public/vault.js)
// und schickt ihn bei jeder Anfrage im Header mit. Der Server
//  - leitet daraus die Tresor-ID ab (HMAC) und trennt damit die Daten verschiedener Passwoerter,
//  - verschluesselt/entschluesselt damit die Ausgaben,
//  - speichert den Schluessel NIE (weder Datenbank noch Log).
// Ein anderes Passwort ergibt einen anderen Tresor; wer das Passwort kennt, kommt in den Tresor. Ein vergessenes
// Passwort laesst sich nicht zuruecksetzen.

export const VAULT_HEADER = 'x-vault-key';
const CHECK_TEXT = 'pocket-budget-vault';
const MAX_FAILS = 5;
const LOCK_MS = 15 * 60 * 1000;

// Oeffentlicher Salt dieser Installation (verhindert fertige Passwort-Tabellen fuer alle Installationen zugleich).
export const vaultSalt = () => createHash('sha256').update(`pocket-budget vault salt|${config.encKey}`).digest('hex');

export function vaultFromKey(b64) {
  const key = Buffer.from(String(b64 || ''), 'base64url');
  if (key.length !== 32) throw new HttpError(401, 'vault_locked', 'Bitte mit deinem Passwort entsperren');
  const id = createHmac('sha256', key).update('pocket-budget vault id').digest('hex');
  return { id, key };
}

export const vaultFromRequest = (req) => vaultFromKey(req.headers[VAULT_HEADER]);

// Nur fuer bestehende Tresore: ein unbekannter Schluessel (falsches Passwort) legt nichts an und liest nichts.
export async function requireVault(req) {
  const vault = vaultFromRequest(req);
  const meta = await store.get('meta', vault.id);
  if (!meta?.check) throw new HttpError(401, 'vault_locked', 'Bitte mit deinem Passwort entsperren');
  return vault;
}

export async function createVault(req) {
  const vault = vaultFromRequest(req);
  if (await store.get('meta', vault.id)) {
    throw new HttpError(409, 'vault_exists', 'Zu diesem Passwort gibt es schon einen Tresor. Bitte entsperren statt neu anlegen.');
  }
  await store.set('meta', vault.id, { check: encrypt(CHECK_TEXT, vault.key), created: new Date().toISOString() });
  return vault;
}

// Entsperren mit Bremse: Nach zu vielen Fehlversuchen (pro Microsoft-Konto) wird kurz gesperrt, damit niemand Passwoerter durchprobiert.
export async function unlockVault(req, userId) {
  const now = Date.now();
  const attempts = (await store.get('attempts', userId)) ?? { count: 0, until: 0 };
  if (attempts.until > now) {
    throw new HttpError(429, 'vault_throttled', `Zu viele Fehlversuche. Bitte in ${Math.ceil((attempts.until - now) / 60000)} Minuten erneut versuchen.`);
  }
  const vault = vaultFromRequest(req);
  const meta = await store.get('meta', vault.id);
  let ok = false;
  try {
    const plain = meta?.check ? Buffer.from(decrypt(meta.check, vault.key)) : null;
    ok = !!plain && plain.length === CHECK_TEXT.length && timingSafeEqual(plain, Buffer.from(CHECK_TEXT));
  } catch {
    ok = false;
  }
  if (!ok) {
    const count = attempts.count + 1;
    await store.set('attempts', userId, count >= MAX_FAILS ? { count: 0, until: now + LOCK_MS } : { count, until: 0 });
    throw new HttpError(401, 'vault_unknown', 'Dieses Passwort ist nicht bekannt. Bitte Eingabe prüfen oder einen neuen Tresor anlegen.');
  }
  if (attempts.count) await store.delete('attempts', userId);
  return vault;
}
