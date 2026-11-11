/* Koneksi add-in ke Aktara API. ES5 + XMLHttpRequest (jalan di engine IE11 Word 2016).
   Prinsip: fitur inti tidak pernah menunggu server. Bila offline, nomor akta diambil dari cadangan lokal
   dan ditandai "sementara" agar dicek ulang. Lisensi berlaku offline sampai 30 hari sejak cek terakhir. */
(function (root) {
  "use strict";
  var K = { url: "ak.server", token: "ak.token", dicek: "ak.lisensiDicek", kantor: "ak.kantor", device: "ak.device", batas: "ak.batasOffline" };
  // Alamat API: config.js (diisi saat deploy) > server dev lokal > produksi
  var BAWAAN = root.AKTARA_API || ((location.hostname === "localhost") ? "https://localhost:8080" : "https://aktara-api.onrender.com");

  function get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function set(k, v) { try { if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) {} }

  function deviceId() {
    var id = get(K.device);
    if (!id) {
      id = "dev-" + new Date().getTime().toString(36) + "-" + Math.random().toString(36).slice(2, 10);
      set(K.device, id);
    }
    return id;
  }

  function panggil(method, path, body, cb) {
    var x = new XMLHttpRequest();
    x.open(method, (get(K.url) || BAWAAN) + path, true);
    x.setRequestHeader("Content-Type", "application/json");
    var t = get(K.token); if (t) x.setRequestHeader("Authorization", "Bearer " + t);
    x.timeout = 8000;
    x.onload = function () {
      var data = {}; try { data = JSON.parse(x.responseText); } catch (e) {}
      if (x.status >= 200 && x.status < 300) cb(null, data);
      else cb({ status: x.status, pesan: data.error || ("HTTP " + x.status) });
    };
    x.onerror = x.ontimeout = function () { cb({ status: 0, pesan: "Tidak dapat terhubung ke server" }); };
    x.send(body ? JSON.stringify(body) : null);
  }

  function aktivasi(kode, namaPerangkat, cb) {
    panggil("POST", "/api/aktivasi", { kode_lisensi: kode, device_id: deviceId(), nama_perangkat: namaPerangkat }, function (err, d) {
      if (err) return cb(err);
      set(K.token, d.token); set(K.kantor, d.kantor); set(K.batas, String(d.batas_offline_hari));
      set(K.dicek, String(new Date().getTime()));
      cb(null, d);
    });
  }

  /* Status lisensi: "aktif" | "offline" (masih dalam toleransi) | "kedaluwarsa" | "belum" */
  function statusLisensi() {
    if (!get(K.token)) return { status: "belum" };
    var dicek = parseInt(get(K.dicek) || "0", 10), batas = parseInt(get(K.batas) || "30", 10);
    var sisaHari = batas - Math.floor((new Date().getTime() - dicek) / 86400000);
    return { status: sisaHari > 0 ? "aktif" : "kedaluwarsa", kantor: get(K.kantor), sisaHariOffline: sisaHari };
  }
  function cekLisensi(cb) {
    panggil("GET", "/api/lisensi", null, function (err) {
      if (!err) { set(K.dicek, String(new Date().getTime())); return cb(null, statusLisensi()); }
      if (err.status === 401 || err.status === 402) { set(K.token, null); return cb(null, { status: err.status === 402 ? "kedaluwarsa" : "belum", pesan: err.pesan }); }
      var s = statusLisensi(); if (s.status === "aktif") s.status = "offline"; cb(null, s);
    });
  }

  /* Nomor akta: dari server bila online; bila offline, cadangan lokal bertanda sementara */
  function nomorAkta(cb) {
    var d = new Date(), tahun = d.getFullYear(), bulan = d.getMonth() + 1;
    var kunci = "ak.nomor." + tahun + "-" + bulan;
    panggil("POST", "/api/nomor-akta", { tahun: tahun, bulan: bulan }, function (err, r) {
      if (!err) { set(kunci, String(r.nomor)); return cb(null, { nomor: r.nomor, sementara: false }); }
      var n = parseInt(get(kunci) || "0", 10) + 1; set(kunci, String(n));
      cb(null, { nomor: n, sementara: true, pesan: err.pesan });
    });
  }

  function catatRepertorium(akta, cb) {
    var d = new Date(akta.tanggal + "T00:00:00");
    panggil("POST", "/api/repertorium", {
      tahun: d.getFullYear(), bulan: d.getMonth() + 1, nomor: akta.nomor, tanggal: akta.tanggal,
      sifat_akta: akta.sifat_akta, penghadap: akta.penghadap
    }, cb);
  }

  function aturServer(url) { set(K.url, url ? url.replace(/\/+$/, "") : null); }
  function urlServer() { return get(K.url) || BAWAAN; }
  function keluar() { set(K.token, null); set(K.kantor, null); }

  root.AktaraServer = { aktivasi: aktivasi, statusLisensi: statusLisensi, cekLisensi: cekLisensi, nomorAkta: nomorAkta,
    catatRepertorium: catatRepertorium, aturServer: aturServer, urlServer: urlServer, keluar: keluar, deviceId: deviceId };
})(this);
