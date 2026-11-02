import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { config } from './config.js';

// AES-256-GCM. Format: v1.<iv>.<tag>.<ciphertext> (base64url)
function key() {
  const k = Buffer.from(config.encKey || '', 'base64');
  if (k.length !== 32) throw new Error('TOKEN_ENC_KEY muss 32 Bytes (base64) lang sein. Erzeugen: npm run gen-key');
  return k;
}

// Ohne `dataKey` gilt der Server-Schluessel (Tokens), sonst der uebergebene (Tresor des Benutzers, siehe lib/vault.js).
export function encrypt(plain, dataKey = key()) {
  const iv = randomBytes(12);
  const cipher = createCipheriv('aes-256-gcm', dataKey, iv);
  const ct = Buffer.concat([cipher.update(plain, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return ['v1', iv, tag, ct].map((p) => (typeof p === 'string' ? p : p.toString('base64url'))).join('.');
}

export function decrypt(payload, dataKey = key()) {
  const [v, iv, tag, ct] = String(payload).split('.');
  if (v !== 'v1' || !iv || !tag || !ct) throw new Error('Ungueltiges Verschluesselungsformat');
  const decipher = createDecipheriv('aes-256-gcm', dataKey, Buffer.from(iv, 'base64url'));
  decipher.setAuthTag(Buffer.from(tag, 'base64url'));
  return Buffer.concat([decipher.update(Buffer.from(ct, 'base64url')), decipher.final()]).toString('utf8');
}
