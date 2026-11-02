import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { neon } from '@neondatabase/serverless';
import { config } from './config.js';

// Key-Value-Speicher. Mit DATABASE_URL (Neon Postgres, z. B. auf Vercel) wird eine Tabelle genutzt,
// sonst eine lokale JSON-Datei (nur Entwicklung). Gespeichert werden Sitzungen, Login-Vorgaenge, das
// VERSCHLUESSELTE Refresh Token, die Datei-ID und die Ausgaben/Favoriten (nur verschluesselt, siehe lib/expenses.js).
// Collections: pending (Login-Vorgaenge), sessions, connections, files, expenses:<userId>, tombstones:<userId>,
// favorites, meta.

// ---- Lokale Datei ----------------------------------------------------------------------------
let chain = Promise.resolve();
const file = () => join(config.dataDir, 'store.json');

async function load() {
  try {
    return JSON.parse(await readFile(file(), 'utf8'));
  } catch (e) {
    if (e.code === 'ENOENT') return {};
    throw e;
  }
}

async function save(db) {
  await mkdir(config.dataDir, { recursive: true });
  const tmp = `${file()}.tmp`;
  await writeFile(tmp, JSON.stringify(db, null, 2), 'utf8');
  // Unter Windows scheitert das Umbenennen kurz, wenn Virenscanner/Indexer die Datei gerade oeffnen.
  for (let attempt = 0; ; attempt++) {
    try {
      return await rename(tmp, file());
    } catch (e) {
      if (!['EPERM', 'EBUSY', 'EACCES'].includes(e.code) || attempt >= 5) throw e;
      await new Promise((r) => setTimeout(r, 20 * (attempt + 1)));
    }
  }
}

// Schreibzugriffe nacheinander, damit parallele Requests keine Daten ueberschreiben.
function exclusive(fn) {
  const run = chain.then(fn);
  chain = run.catch(() => {});
  return run;
}

const fileStore = {
  get: (collection, key) => exclusive(async () => (await load())[collection]?.[key] ?? null),
  set: (collection, key, value) =>
    exclusive(async () => {
      const db = await load();
      (db[collection] ??= {})[key] = value;
      await save(db);
    }),
  delete: (collection, key) =>
    exclusive(async () => {
      const db = await load();
      if (db[collection]?.[key] !== undefined) {
        delete db[collection][key];
        await save(db);
      }
    }),
  list: (collection) => exclusive(async () => Object.entries((await load())[collection] ?? {}).map(([key, value]) => ({ key, value }))),
  clear: (collection) =>
    exclusive(async () => {
      const db = await load();
      if (db[collection]) {
        delete db[collection];
        await save(db);
      }
    }),
};

// ---- Neon Postgres ---------------------------------------------------------------------------
let sqlClient;
let ready;

async function sql() {
  sqlClient ??= neon(config.databaseUrl);
  ready ??= sqlClient`CREATE TABLE IF NOT EXISTS kv (
    collection text NOT NULL,
    key text NOT NULL,
    value jsonb NOT NULL,
    updated_at timestamptz NOT NULL DEFAULT now(),
    PRIMARY KEY (collection, key)
  )`;
  try {
    await ready;
  } catch (e) {
    ready = undefined;
    throw e;
  }
  return sqlClient;
}

const pgStore = {
  async get(collection, key) {
    const q = await sql();
    const rows = await q`SELECT value FROM kv WHERE collection = ${collection} AND key = ${key}`;
    return rows[0]?.value ?? null;
  },
  async set(collection, key, value) {
    const q = await sql();
    if (collection === 'pending') {
      // Abgelaufene Login-Vorgaenge nebenbei aufraeumen.
      await q`DELETE FROM kv WHERE collection IN ('pending', 'sessions') AND (value->>'expires')::bigint < ${Date.now()}`;
    }
    await q`INSERT INTO kv (collection, key, value) VALUES (${collection}, ${key}, ${JSON.stringify(value)}::jsonb)
            ON CONFLICT (collection, key) DO UPDATE SET value = EXCLUDED.value, updated_at = now()`;
  },
  async delete(collection, key) {
    const q = await sql();
    await q`DELETE FROM kv WHERE collection = ${collection} AND key = ${key}`;
  },
  async list(collection) {
    const q = await sql();
    return q`SELECT key, value FROM kv WHERE collection = ${collection}`;
  },
  async clear(collection) {
    const q = await sql();
    await q`DELETE FROM kv WHERE collection = ${collection}`;
  },
};

const backend = () => (config.databaseUrl ? pgStore : fileStore);

export const store = {
  get: (c, k) => backend().get(c, k),
  set: (c, k, v) => backend().set(c, k, v),
  delete: (c, k) => backend().delete(c, k),
  // Alle Eintraege einer Collection als [{ key, value }] (Reihenfolge nicht garantiert).
  list: (c) => backend().list(c),
  // Entfernt eine ganze Collection.
  clear: (c) => backend().clear(c),
};
