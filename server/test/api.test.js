"use strict";
const test = require("node:test");
const assert = require("node:assert");
const { buka } = require("../src/db");
const { buatApp } = require("../src/app");

const ADMIN = "kunci-admin-tes";
let server, base;

test.before(async () => {
  // TEST_DATABASE_URL diisi -> tes dijalankan terhadap Postgres sungguhan (dipakai di CI)
  const db = await buka(process.env.TEST_DATABASE_URL ? { url: process.env.TEST_DATABASE_URL } : { path: ":memory:" });
  server = buatApp({ db, rahasia: "rahasia-tes", kunciAdmin: ADMIN });
  await new Promise((r) => server.listen(0, r));
  base = "http://127.0.0.1:" + server.address().port;
});
test.after(async () => { server.close(); });

async function panggil(method, path, body, headers) {
  const r = await fetch(base + path, {
    method, headers: Object.assign({ "Content-Type": "application/json" }, headers || {}),
    body: body ? JSON.stringify(body) : undefined
  });
  const type = r.headers.get("content-type") || "";
  return { status: r.status, data: type.includes("json") ? await r.json() : await r.text() };
}
const auth = (t) => ({ Authorization: "Bearer " + t });

async function kantorBaru(opsi) {
  const r = await panggil("POST", "/api/admin/kantor",
    Object.assign({ nama: "Kantor Notaris Uji, S.H., M.Kn.", berlaku_sampai: "2099-12-31", maks_perangkat: 2 }, opsi || {}), { "X-Admin-Key": ADMIN });
  assert.strictEqual(r.status, 200);
  return r.data.kode_lisensi;
}
async function aktivasi(kode, device) {
  return panggil("POST", "/api/aktivasi", { kode_lisensi: kode, device_id: device, nama_perangkat: device });
}

test("health", async () => {
  assert.deepStrictEqual((await panggil("GET", "/api/health")).data.status, "ok");
});

test("admin wajib kunci yang benar", async () => {
  const r = await panggil("POST", "/api/admin/kantor", { nama: "X", berlaku_sampai: "2099-01-01" }, { "X-Admin-Key": "salah" });
  assert.strictEqual(r.status, 403);
});

test("aktivasi instan dan batas jumlah perangkat", async () => {
  const kode = await kantorBaru();
  assert.match(kode, /^AKT-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}$/);
  assert.strictEqual((await aktivasi(kode, "pc-1")).status, 200);
  assert.strictEqual((await aktivasi(kode.toLowerCase(), "pc-2")).status, 200, "kode tidak peka huruf besar/kecil");
  assert.strictEqual((await aktivasi(kode, "pc-1")).status, 200, "aktivasi ulang perangkat sama boleh");
  const r = await aktivasi(kode, "pc-3");
  assert.strictEqual(r.status, 409);
  assert.match(r.data.error, /batas paket/);
});

test("kode lisensi salah atau kedaluwarsa ditolak", async () => {
  assert.strictEqual((await aktivasi("AKT-0000-0000-0000-0000", "pc")).status, 404);
  const kode = await kantorBaru({ berlaku_sampai: "2020-01-01" });
  assert.strictEqual((await aktivasi(kode, "pc")).status, 402);
});

test("token palsu ditolak", async () => {
  assert.strictEqual((await panggil("GET", "/api/lisensi", null, auth("abc.def"))).status, 401);
  assert.strictEqual((await panggil("GET", "/api/lisensi")).status, 401);
});

test("cek lisensi memberi batas offline 30 hari", async () => {
  const kode = await kantorBaru();
  const { data } = await aktivasi(kode, "pc-1");
  const r = await panggil("GET", "/api/lisensi", null, auth(data.token));
  assert.strictEqual(r.status, 200);
  assert.strictEqual(r.data.batas_offline_hari, 30);
});

test("nomor akta berurutan, tidak ganda walau diminta bersamaan dari 2 komputer", async () => {
  const kode = await kantorBaru();
  const t1 = (await aktivasi(kode, "pc-1")).data.token, t2 = (await aktivasi(kode, "pc-2")).data.token;
  const minta = (t) => panggil("POST", "/api/nomor-akta", { tahun: 2026, bulan: 11 }, auth(t));
  const hasil = await Promise.all([minta(t1), minta(t2), minta(t1), minta(t2), minta(t1), minta(t2)]);
  const nomor = hasil.map((h) => h.data.nomor).sort((a, b) => a - b);
  assert.deepStrictEqual(nomor, [1, 2, 3, 4, 5, 6]);
  // Bulan berikutnya mulai dari 1 lagi
  assert.strictEqual((await panggil("POST", "/api/nomor-akta", { tahun: 2026, bulan: 12 }, auth(t1))).data.nomor, 1);
});

test("nomor akta antar kantor terpisah", async () => {
  const tA = (await aktivasi(await kantorBaru(), "a")).data.token;
  const tB = (await aktivasi(await kantorBaru(), "b")).data.token;
  await panggil("POST", "/api/nomor-akta", { tahun: 2026, bulan: 11 }, auth(tA));
  assert.strictEqual((await panggil("POST", "/api/nomor-akta", { tahun: 2026, bulan: 11 }, auth(tB))).data.nomor, 1);
});

test("periode tidak valid ditolak", async () => {
  const t = (await aktivasi(await kantorBaru(), "pc")).data.token;
  assert.strictEqual((await panggil("POST", "/api/nomor-akta", { tahun: 2026, bulan: 13 }, auth(t))).status, 400);
});

test("repertorium: catat, tolak nomor ganda, ekspor CSV", async () => {
  const t = (await aktivasi(await kantorBaru(), "pc")).data.token;
  const akta = { tahun: 2026, bulan: 11, nomor: 1, tanggal: "2026-11-11", sifat_akta: "Kuasa", penghadap: 'Budi "Santoso"' };
  assert.strictEqual((await panggil("POST", "/api/repertorium", akta, auth(t))).status, 200);
  assert.strictEqual((await panggil("POST", "/api/repertorium", Object.assign({}, akta, { nomor: 2, sifat_akta: "PPJB", penghadap: "Ani; Budi" }), auth(t))).status, 200);
  assert.strictEqual((await panggil("POST", "/api/repertorium", akta, auth(t))).status, 409);
  assert.strictEqual((await panggil("POST", "/api/repertorium", Object.assign({}, akta, { nomor: 3, tanggal: "11/11/2026" }), auth(t))).status, 400);

  const j = await panggil("GET", "/api/repertorium?tahun=2026&bulan=11", null, auth(t));
  assert.strictEqual(j.data.jumlah, 2);
  const csv = await panggil("GET", "/api/repertorium?tahun=2026&bulan=11&format=csv", null, auth(t));
  const baris = csv.data.replace(/^﻿/, "").split("\r\n");
  assert.strictEqual(baris.length, 3);
  assert.strictEqual(baris[1], '1;1;2026-11-11;"Kuasa";"Budi ""Santoso"""');
  assert.strictEqual(baris[2], '2;2;2026-11-11;"PPJB";"Ani; Budi"');
});

test("JSON rusak dan rute tak dikenal", async () => {
  const r = await fetch(base + "/api/aktivasi", { method: "POST", body: "{rusak" });
  assert.strictEqual(r.status, 400);
  assert.strictEqual((await panggil("GET", "/api/tidak-ada")).status, 404);
});
