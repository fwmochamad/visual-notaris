// Penyimpanan data Aktara. Pengembangan & tes: SQLite bawaan Node (node:sqlite).
// Produksi: skema yang sama dipindah ke PostgreSQL (lihat server/README.md).
"use strict";
const { DatabaseSync } = require("node:sqlite");

const SKEMA = `
CREATE TABLE IF NOT EXISTS kantor (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nama TEXT NOT NULL,
  kode_lisensi TEXT NOT NULL UNIQUE,
  paket TEXT NOT NULL DEFAULT 'dasar',
  maks_perangkat INTEGER NOT NULL DEFAULT 2,
  berlaku_sampai TEXT NOT NULL,
  dibuat TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS perangkat (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kantor_id INTEGER NOT NULL REFERENCES kantor(id),
  device_id TEXT NOT NULL,
  nama TEXT,
  aktif INTEGER NOT NULL DEFAULT 1,
  terakhir_online TEXT NOT NULL DEFAULT (datetime('now')),
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
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kantor_id INTEGER NOT NULL REFERENCES kantor(id),
  tahun INTEGER NOT NULL,
  bulan INTEGER NOT NULL,
  nomor INTEGER NOT NULL,
  tanggal TEXT NOT NULL,
  sifat_akta TEXT NOT NULL,
  penghadap TEXT NOT NULL,
  dibuat_oleh TEXT,
  dibuat TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE (kantor_id, tahun, bulan, nomor)
);
CREATE TABLE IF NOT EXISTS audit (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  kantor_id INTEGER,
  device_id TEXT,
  aksi TEXT NOT NULL,
  detail TEXT,
  waktu TEXT NOT NULL DEFAULT (datetime('now'))
);`;

function buka(path) {
  const db = new DatabaseSync(path || ":memory:");
  db.exec("PRAGMA foreign_keys = ON;");
  db.exec(SKEMA);
  return db;
}

// Menjalankan fungsi dalam satu transaksi (dipakai agar nomor akta tidak pernah ganda).
function transaksi(db, fn) {
  db.exec("BEGIN IMMEDIATE");
  try { const r = fn(); db.exec("COMMIT"); return r; }
  catch (e) { db.exec("ROLLBACK"); throw e; }
}

module.exports = { buka, transaksi };
