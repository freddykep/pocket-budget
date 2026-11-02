// Tresor-Schluessel im Browser. Aus dem Passwort wird mit PBKDF2 ein 32-Byte-Schluessel abgeleitet; nur dieser
// (nie das Passwort) geht zum Server. Er liegt in sessionStorage: nach dem Schliessen der App ist wieder das Passwort
// noetig (der Passwort-Manager des Handys fuellt es per Fingerabdruck/Gesicht aus).

const KEY = 'pb_vault_key';
export const MIN_PASSWORD = 10;
const ITERATIONS = 600_000;

let memoryKey = null; // Rueckfall, falls sessionStorage nicht verfuegbar ist

const toBase64Url = (bytes) => btoa(String.fromCharCode(...bytes)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

export async function deriveKey(password, salt) {
  const enc = new TextEncoder();
  const material = await crypto.subtle.importKey('raw', enc.encode(String(password).normalize('NFKC')), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits(
    { name: 'PBKDF2', hash: 'SHA-256', salt: enc.encode(`pocket-budget|${salt}`), iterations: ITERATIONS },
    material,
    256,
  );
  return toBase64Url(new Uint8Array(bits));
}

export function setVaultKey(key) {
  memoryKey = key;
  try { sessionStorage.setItem(KEY, key); } catch { /* bleibt im Arbeitsspeicher */ }
}

export function clearVaultKey() {
  memoryKey = null;
  try { sessionStorage.removeItem(KEY); } catch { /* nichts zu tun */ }
}

export function currentVaultKey() {
  try { return sessionStorage.getItem(KEY) ?? memoryKey; } catch { return memoryKey; }
}
