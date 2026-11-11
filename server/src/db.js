// Penyimpanan data Aktara dengan antarmuka async yang sama untuk dua mesin:
//  - PostgreSQL (produksi, mis. Neon)  : bila DATABASE_URL diisi
//  - SQLite bawaan Node (dev & tes)   : selain itu
// Query ditulis dengan placeholder $1, $2, ... (gaya Postgres); untuk SQLite diterjemahkan otomatis.
"use strict";

const SKEMA_PG = `
CREATE TABLE IF NOT EXISTS kantor (
  id SERIAL PRIMARY KEY,
  nama TEXT NOT NULL,
  kode_lisensi TEXT NOT NULL UNIQUE,
  paket TEXT NOT NULL DEFAULT 'dasar',
  maks_perangkat INTEGER NOT NULL DEFAULT 2,
  berlaku_sampai TEXT NOT NULL,
  dibuat TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS perangkat (
  id SERIAL PRIMARY KEY,
  kantor_id INTEGER NOT NULL REFERENCES kantor(id),
  device_id TEXT NOT NULL,
  nama TEXT,
  aktif INTEGER NOT NULL DEFAULT 1,
  terakhir_online TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE (kantor_id, device_id)
);
CREATE TABLE IF NOT EXISTS nomor_akta (
  kantor_id INTEGER NOT NULL REFERENCES kantor(id),
  tahun INTEGER NOT NULL,
  bulan INTEGER NOT NULL,
  terakhir INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (kantor_id, tahun, bulan)
);
CREATE TABLE IF NOT EXISTS repertorium (
  id SERIAL PRIMARY KEY,
  kantor_id INTEGER NOT NULL REFERENCES kantor(id),
  tahun INTEGER NOT NULL,
  bulan INTEGER NOT NULL,
  nomor INTEGER NOT NULL,
  tanggal TEXT NOT NULL,
  sifat_akta TEXT NOT NULL,
  penghadap TEXT NOT NULL,
  dibuat_oleh TEXT,
  dibuat TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
  UNIQUE (kantor_id, tahun, bulan, nomor)
);
CREATE TABLE IF NOT EXISTS audit (
  id SERIAL PRIMARY KEY,
  kantor_id INTEGER,
  device_id TEXT,
  aksi TEXT NOT NULL,
  detail TEXT,
  waktu TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);`;

const SKEMA_SQLITE = SKEMA_PG
  .replace(/SERIAL PRIMARY KEY/g, "INTEGER PRIMARY KEY AUTOINCREMENT")
  .replace(/TIMESTAMPTZ/g, "TEXT");

async function bukaPostgres(url) {
  const { Pool } = require("pg");
  // Neon & layanan cloud memakai SSL; ?sslmode=disable hanya untuk Postgres lokal/CI
  const ssl = /sslmode=disable/.test(url) ? false : { rejectUnauthorized: true };
  const pool = new Pool({ connectionString: url.replace(/[?&]sslmode=disable/, ""), ssl, max: 5 });
  await pool.query(SKEMA_PG);
  return {
    mesin: "postgres",
    async all(sql, p) { return (await pool.query(sql, p || [])).rows; },
    async get(sql, p) { return (await pool.query(sql, p || [])).rows[0]; },
    async run(sql, p) { const r = await pool.query(sql, p || []); return { changes: r.rowCount, rows: r.rows }; },
    tutup() { return pool.end(); }
  };
}

function bukaSqlite(path) {
  const { DatabaseSync } = require("node:sqlite");
  const db = new DatabaseSync(path || ":memory:");
  db.exec("PRAGMA foreign_keys = ON;");
  db.exec(SKEMA_SQLITE);
  const ubah = (sql) => sql.replace(/\$\d+/g, "?");
  return {
    mesin: "sqlite",
    async all(sql, p) { return db.prepare(ubah(sql)).all(...(p || [])); },
    async get(sql, p) { return db.prepare(ubah(sql)).get(...(p || [])); },
    async run(sql, p) {
      if (/\bRETURNING\b/i.test(sql)) { const rows = db.prepare(ubah(sql)).all(...(p || [])); return { changes: rows.length, rows }; }
      const r = db.prepare(ubah(sql)).run(...(p || [])); return { changes: Number(r.changes), rows: [] };
    },
    tutup() { db.close(); }
  };
}

// buka({ url }) -> Postgres ; buka({ path }) -> SQLite
async function buka(opsi) {
  opsi = opsi || {};
  return opsi.url ? bukaPostgres(opsi.url) : bukaSqlite(opsi.path);
}

module.exports = { buka };
