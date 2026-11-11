// API Aktara: lisensi & aktivasi, nomor akta terpusat, repertorium. Tanpa dependensi eksternal.
"use strict";
const http = require("node:http");
const crypto = require("node:crypto");

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
    db.run("INSERT INTO audit (kantor_id, device_id, aksi, detail) VALUES ($1,$2,$3,$4)",
      [kantorId, deviceId, aksi, detail ? JSON.stringify(detail) : null]);

  // ---------- Otentikasi ----------
  async function perangkatDariToken(req) {
    const t = bacaToken((req.headers.authorization || "").replace(/^Bearer\s+/i, ""));
    const p = await db.get(`SELECT p.*, k.berlaku_sampai, k.nama AS nama_kantor FROM perangkat p
      JOIN kantor k ON k.id = p.kantor_id WHERE p.kantor_id = $1 AND p.device_id = $2`, [t.k, t.d]);
    if (!p || !p.aktif) throw new HttpError(401, "Perangkat tidak terdaftar atau sudah dinonaktifkan");
    if (p.berlaku_sampai < hariIni()) throw new HttpError(402, "Lisensi kantor sudah habis masa berlakunya");
    await db.run("UPDATE perangkat SET terakhir_online = CURRENT_TIMESTAMP WHERE id = $1", [p.id]);
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
    "GET /api/health": async () => ({ status: "ok", versi: "0.2.0", db: db.mesin }),

    // Admin Aktara membuat kantor pelanggan + kode lisensi
    "POST /api/admin/kantor": async (req, body) => {
      wajibAdmin(req); wajib(body, ["nama", "berlaku_sampai"]);
      const kode = kodeLisensiBaru();
      const r = await db.run("INSERT INTO kantor (nama, kode_lisensi, paket, maks_perangkat, berlaku_sampai) VALUES ($1,$2,$3,$4,$5) RETURNING id",
        [body.nama, kode, body.paket || "dasar", body.maks_perangkat || 2, body.berlaku_sampai]);
      return { id: Number(r.rows[0].id), kode_lisensi: kode };
    },

    // Aktivasi instan dari add-in: kode lisensi + id perangkat -> token
    "POST /api/aktivasi": async (req, body) => {
      wajib(body, ["kode_lisensi", "device_id"]);
      const k = await db.get("SELECT * FROM kantor WHERE kode_lisensi = $1", [String(body.kode_lisensi).trim().toUpperCase()]);
      if (!k) throw new HttpError(404, "Kode lisensi tidak ditemukan");
      if (k.berlaku_sampai < hariIni()) throw new HttpError(402, "Lisensi sudah habis masa berlakunya");
      const ada = await db.get("SELECT * FROM perangkat WHERE kantor_id = $1 AND device_id = $2", [k.id, body.device_id]);
      if (!ada || !ada.aktif) {
        const jml = Number((await db.get("SELECT COUNT(*) AS n FROM perangkat WHERE kantor_id = $1 AND aktif = 1", [k.id])).n);
        if (jml >= k.maks_perangkat) throw new HttpError(409, "Jumlah perangkat sudah mencapai batas paket (" + k.maks_perangkat + "). Nonaktifkan perangkat lama di portal.");
        if (ada) await db.run("UPDATE perangkat SET aktif = 1, nama = $1 WHERE id = $2", [body.nama_perangkat || ada.nama, ada.id]);
        else await db.run("INSERT INTO perangkat (kantor_id, device_id, nama) VALUES ($1,$2,$3)", [k.id, body.device_id, body.nama_perangkat || null]);
      }
      await audit(k.id, body.device_id, "aktivasi");
      return {
        token: buatToken({ k: k.id, d: body.device_id }),
        kantor: k.nama, paket: k.paket, berlaku_sampai: k.berlaku_sampai, batas_offline_hari: BATAS_OFFLINE_HARI
      };
    },

    "GET /api/lisensi": async (req) => {
      const p = await perangkatDariToken(req);
      return { kantor: p.nama_kantor, berlaku_sampai: p.berlaku_sampai, batas_offline_hari: BATAS_OFFLINE_HARI, dicek: sekarang().toISOString() };
    },

    // Nomor akta berikutnya: satu pernyataan upsert yang atomik -> tidak pernah ganda antar komputer
    "POST /api/nomor-akta": async (req, body) => {
      const p = await perangkatDariToken(req);
      const { tahun, bulan } = periode(body);
      const r = await db.run(`INSERT INTO nomor_akta (kantor_id, tahun, bulan, terakhir) VALUES ($1,$2,$3,1)
        ON CONFLICT (kantor_id, tahun, bulan) DO UPDATE SET terakhir = nomor_akta.terakhir + 1 RETURNING terakhir`,
        [p.kantor_id, tahun, bulan]);
      const nomor = Number(r.rows[0].terakhir);
      await audit(p.kantor_id, p.device_id, "nomor-akta", { tahun, bulan, nomor });
      return { tahun, bulan, nomor };
    },

    // Catat akta final ke repertorium
    "POST /api/repertorium": async (req, body) => {
      const p = await perangkatDariToken(req);
      wajib(body, ["nomor", "tanggal", "sifat_akta", "penghadap"]);
      const { tahun, bulan } = periode(body);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(body.tanggal)) throw new HttpError(400, "Format tanggal harus YYYY-MM-DD");
      try {
        const r = await db.run(`INSERT INTO repertorium (kantor_id, tahun, bulan, nomor, tanggal, sifat_akta, penghadap, dibuat_oleh)
          VALUES ($1,$2,$3,$4,$5,$6,$7,$8) RETURNING id`,
          [p.kantor_id, tahun, bulan, parseInt(body.nomor, 10), body.tanggal, body.sifat_akta, body.penghadap, p.nama || p.device_id]);
        await audit(p.kantor_id, p.device_id, "repertorium", { tahun, bulan, nomor: body.nomor });
        return { id: Number(r.rows[0].id) };
      } catch (e) {
        if (/UNIQUE|duplicate key/i.test(e.message)) throw new HttpError(409, "Nomor akta " + body.nomor + " bulan ini sudah tercatat di repertorium");
        throw e;
      }
    },

    "GET /api/repertorium": async (req, body, query) => {
      const p = await perangkatDariToken(req);
      const { tahun, bulan } = periode(query);
      const baris = await db.all(`SELECT nomor, tanggal, sifat_akta, penghadap FROM repertorium
        WHERE kantor_id = $1 AND tahun = $2 AND bulan = $3 ORDER BY nomor`, [p.kantor_id, tahun, bulan]);
      if (query.format === "csv") {
        const esc = (v) => '"' + String(v).replace(/"/g, '""') + '"';
        const csv = ["No. Urut;Nomor Akta;Tanggal;Sifat Akta;Nama Penghadap"]
          .concat(baris.map((b, i) => [i + 1, b.nomor, b.tanggal, esc(b.sifat_akta), esc(b.penghadap)].join(";"))).join("\r\n");
        return { __raw: "\ufeff" + csv, type: "text/csv; charset=utf-8", nama: "repertorium-" + tahun + "-" + String(bulan).padStart(2, "0") + ".csv" };
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
    req.on("end", async () => {
      try {
        if (!fn) throw new HttpError(404, "Tidak ditemukan");
        let body = {};
        if (data) { try { body = JSON.parse(data); } catch (e) { throw new HttpError(400, "JSON tidak valid"); } }
        const hasil = await fn(req, body, Object.fromEntries(url.searchParams));
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
