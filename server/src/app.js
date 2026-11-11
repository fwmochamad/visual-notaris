// API Aktara: lisensi & aktivasi, nomor akta terpusat, repertorium. Tanpa dependensi eksternal.
"use strict";
const http = require("node:http");
const crypto = require("node:crypto");
const { transaksi } = require("./db");

const BATAS_OFFLINE_HARI = 30;

class HttpError extends Error { constructor(status, pesan) { super(pesan); this.status = status; } }

function buatApp({ db, rahasia, kunciAdmin, sekarang = () => new Date() }) {
  if (!rahasia || !kunciAdmin) throw new Error("rahasia dan kunciAdmin wajib diisi");

  // ---------- Token perangkat (HMAC, tanpa library JWT) ----------
  const b64 = (s) => Buffer.from(s).toString("base64url");
  const tanda = (data) => crypto.createHmac("sha256", rahasia).update(data).digest("base64url");
  function buatToken(isi) { const d = b64(JSON.stringify(isi)); return d + "." + tanda(d); }
  function bacaToken(token) {
    const [d, s] = String(token || "").split(".");
    if (!d || !s) throw new HttpError(401, "Token tidak ada");
    const a = Buffer.from(tanda(d)), b = Buffer.from(s);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) throw new HttpError(401, "Token tidak valid");
    return JSON.parse(Buffer.from(d, "base64url").toString());
  }

  function kodeLisensiBaru() {
    const h = crypto.randomBytes(8).toString("hex").toUpperCase();
    return "AKT-" + h.slice(0, 4) + "-" + h.slice(4, 8) + "-" + h.slice(8, 12) + "-" + h.slice(12, 16);
  }
  const hariIni = () => sekarang().toISOString().slice(0, 10);
  const audit = (kantorId, deviceId, aksi, detail) =>
    db.prepare("INSERT INTO audit (kantor_id, device_id, aksi, detail) VALUES (?,?,?,?)")
      .run(kantorId, deviceId, aksi, detail ? JSON.stringify(detail) : null);

  // ---------- Otentikasi ----------
  function perangkatDariToken(req) {
    const t = bacaToken((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const p = db.prepare(`SELECT p.*, k.berlaku_sampai, k.nama AS nama_kantor FROM perangkat p
      JOIN kantor k ON k.id = p.kantor_id WHERE p.kantor_id = ? AND p.device_id = ?`).get(t.k, t.d);
    if (!p || !p.aktif) throw new HttpError(401, "Perangkat tidak terdaftar atau sudah dinonaktifkan");
    if (p.berlaku_sampai < hariIni()) throw new HttpError(402, "Lisensi kantor sudah habis masa berlakunya");
    db.prepare("UPDATE perangkat SET terakhir_online = datetime('now') WHERE id = ?").run(p.id);
    return p;
  }
  function wajibAdmin(req) {
    const a = Buffer.from(String(req.headers["x-admin-key"] || "")), b = Buffer.from(kunciAdmin);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) throw new HttpError(403, "Bukan admin");
  }

  // ---------- Validasi ----------
  function wajib(obj, kolom) {
    for (const k of kolom) if (obj[k] === undefined || obj[k] === null || obj[k] === "") throw new HttpError(400, "Kolom '" + k + "' wajib diisi");
  }
  function periode(o) {
    const tahun = parseInt(o.tahun, 10), bulan = parseInt(o.bulan, 10);
    if (!(tahun >= 2000 && tahun <= 2100) || !(bulan >= 1 && bulan <= 12)) throw new HttpError(400, "Tahun/bulan tidak valid");
    return { tahun, bulan };
  }

  // ---------- Rute ----------
  const rute = {
    "GET /api/health": () => ({ status: "ok", versi: "0.1.0" }),

    // Admin Aktara membuat kantor pelanggan + kode lisensi
    "POST /api/admin/kantor": (req, body) => {
      wajibAdmin(req); wajib(body, ["nama", "berlaku_sampai"]);
      const kode = kodeLisensiBaru();
      const r = db.prepare("INSERT INTO kantor (nama, kode_lisensi, paket, maks_perangkat, berlaku_sampai) VALUES (?,?,?,?,?)")
        .run(body.nama, kode, body.paket || "dasar", body.maks_perangkat || 2, body.berlaku_sampai);
      return { id: Number(r.lastInsertRowid), kode_lisensi: kode };
    },

    // Aktivasi instan dari add-in: kode lisensi + id perangkat -> token
    "POST /api/aktivasi": (req, body) => {
      wajib(body, ["kode_lisensi", "device_id"]);
      const k = db.prepare("SELECT * FROM kantor WHERE kode_lisensi = ?").get(String(body.kode_lisensi).trim().toUpperCase());
      if (!k) throw new HttpError(404, "Kode lisensi tidak ditemukan");
      if (k.berlaku_sampai < hariIni()) throw new HttpError(402, "Lisensi sudah habis masa berlakunya");
      return transaksi(db, () => {
        const ada = db.prepare("SELECT * FROM perangkat WHERE kantor_id = ? AND device_id = ?").get(k.id, body.device_id);
        if (!ada || !ada.aktif) {
          const jml = db.prepare("SELECT COUNT(*) n FROM perangkat WHERE kantor_id = ? AND aktif = 1").get(k.id).n;
          if (jml >= k.maks_perangkat) throw new HttpError(409, "Jumlah perangkat sudah mencapai batas paket (" + k.maks_perangkat + "). Nonaktifkan perangkat lama di portal.");
          if (ada) db.prepare("UPDATE perangkat SET aktif = 1, nama = ? WHERE id = ?").run(body.nama_perangkat || ada.nama, ada.id);
          else db.prepare("INSERT INTO perangkat (kantor_id, device_id, nama) VALUES (?,?,?)").run(k.id, body.device_id, body.nama_perangkat || null);
        }
        audit(k.id, body.device_id, "aktivasi");
        return {
          token: buatToken({ k: k.id, d: body.device_id }),
          kantor: k.nama, paket: k.paket, berlaku_sampai: k.berlaku_sampai, batas_offline_hari: BATAS_OFFLINE_HARI
        };
      });
    },

    "GET /api/lisensi": (req) => {
      const p = perangkatDariToken(req);
      return { kantor: p.nama_kantor, berlaku_sampai: p.berlaku_sampai, batas_offline_hari: BATAS_OFFLINE_HARI, dicek: sekarang().toISOString() };
    },

    // Nomor akta berikutnya, atomik per kantor per bulan -> tidak pernah ganda antar komputer
    "POST /api/nomor-akta": (req, body) => {
      const p = perangkatDariToken(req);
      const { tahun, bulan } = periode(body);
      const nomor = transaksi(db, () => {
        db.prepare("INSERT INTO nomor_akta (kantor_id, tahun, bulan, terakhir) VALUES (?,?,?,0) ON CONFLICT DO NOTHING").run(p.kantor_id, tahun, bulan);
        db.prepare("UPDATE nomor_akta SET terakhir = terakhir + 1 WHERE kantor_id = ? AND tahun = ? AND bulan = ?").run(p.kantor_id, tahun, bulan);
        return db.prepare("SELECT terakhir FROM nomor_akta WHERE kantor_id = ? AND tahun = ? AND bulan = ?").get(p.kantor_id, tahun, bulan).terakhir;
      });
      audit(p.kantor_id, p.device_id, "nomor-akta", { tahun, bulan, nomor });
      return { tahun, bulan, nomor };
    },

    // Catat akta final ke repertorium
    "POST /api/repertorium": (req, body) => {
      const p = perangkatDariToken(req);
      wajib(body, ["nomor", "tanggal", "sifat_akta", "penghadap"]);
      const { tahun, bulan } = periode(body);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(body.tanggal)) throw new HttpError(400, "Format tanggal harus YYYY-MM-DD");
      try {
        const r = db.prepare(`INSERT INTO repertorium (kantor_id, tahun, bulan, nomor, tanggal, sifat_akta, penghadap, dibuat_oleh)
          VALUES (?,?,?,?,?,?,?,?)`).run(p.kantor_id, tahun, bulan, parseInt(body.nomor, 10), body.tanggal, body.sifat_akta, body.penghadap, p.nama || p.device_id);
        audit(p.kantor_id, p.device_id, "repertorium", { tahun, bulan, nomor: body.nomor });
        return { id: Number(r.lastInsertRowid) };
      } catch (e) {
        if (/UNIQUE/.test(e.message)) throw new HttpError(409, "Nomor akta " + body.nomor + " bulan ini sudah tercatat di repertorium");
        throw e;
      }
    },

    "GET /api/repertorium": (req, body, query) => {
      const p = perangkatDariToken(req);
      const { tahun, bulan } = periode(query);
      const baris = db.prepare(`SELECT nomor, tanggal, sifat_akta, penghadap FROM repertorium
        WHERE kantor_id = ? AND tahun = ? AND bulan = ? ORDER BY nomor`).all(p.kantor_id, tahun, bulan);
      if (query.format === "csv") {
        const esc = (v) => '"' + String(v).replace(/"/g, '""') + '"';
        const csv = ["No. Urut;Nomor Akta;Tanggal;Sifat Akta;Nama Penghadap"]
          .concat(baris.map((b, i) => [i + 1, b.nomor, b.tanggal, esc(b.sifat_akta), esc(b.penghadap)].join(";"))).join("\r\n");
        return { __raw: "﻿" + csv, type: "text/csv; charset=utf-8", nama: "repertorium-" + tahun + "-" + String(bulan).padStart(2, "0") + ".csv" };
      }
      return { tahun, bulan, jumlah: baris.length, data: baris };
    }
  };

  // ---------- Server HTTP ----------
  return http.createServer((req, res) => {
    // CORS: add-in berjalan dari domain hosting add-in
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Admin-Key");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    if (req.method === "OPTIONS") { res.writeHead(204); return res.end(); }

    const url = new URL(req.url, "http://x");
    const fn = rute[req.method + " " + url.pathname];
    let data = "";
    req.on("data", (c) => { data += c; if (data.length > 1e6) req.destroy(); });
    req.on("end", () => {
      try {
        if (!fn) throw new HttpError(404, "Tidak ditemukan");
        let body = {};
        if (data) { try { body = JSON.parse(data); } catch (e) { throw new HttpError(400, "JSON tidak valid"); } }
        const hasil = fn(req, body, Object.fromEntries(url.searchParams));
        if (hasil && hasil.__raw !== undefined) {
          res.writeHead(200, { "Content-Type": hasil.type, "Content-Disposition": 'attachment; filename="' + hasil.nama + '"' });
          return res.end(hasil.__raw);
        }
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify(hasil));
      } catch (e) {
        const status = e.status || 500;
        if (status === 500) console.error(e);
        res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: status === 500 ? "Kesalahan server" : e.message }));
      }
    });
  });
}

module.exports = { buatApp };
