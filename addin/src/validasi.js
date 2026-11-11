/* Validasi data para pihak. ES5. */
(function (root) {
  "use strict";
  /* Memeriksa Nomor Induk Kependudukan (NIK): 16 digit, kode provinsi, tanggal & bulan lahir.
     Mengembalikan "" bila sah, atau pesan yang menjelaskan masalahnya. */
  function cekNIK(nik) {
    var n = String(nik || "").replace(/\s/g, "");
    if (!/^\d{16}$/.test(n)) return "Nomor Induk Kependudukan (NIK) harus 16 digit angka (sekarang " + n.length + " karakter).";
    var prov = +n.slice(0, 2), tgl = +n.slice(6, 8), bln = +n.slice(8, 10);
    if (prov < 11 || prov > 96) return "Dua digit pertama NIK (kode provinsi) tidak dikenal: " + n.slice(0, 2) + ".";
    if (tgl > 40) tgl -= 40; // perempuan: tanggal lahir + 40
    if (tgl < 1 || tgl > 31) return "Digit ke-7 dan ke-8 NIK (tanggal lahir) tidak valid: " + n.slice(6, 8) + ".";
    if (bln < 1 || bln > 12) return "Digit ke-9 dan ke-10 NIK (bulan lahir) tidak valid: " + n.slice(8, 10) + ".";
    return "";
  }
  /* Data yang tersirat dari NIK: jenis kelamin & tanggal lahir (dd-mm-yy) */
  function infoNIK(nik) {
    var n = String(nik || "").replace(/\s/g, "");
    if (cekNIK(n)) return null;
    var tgl = +n.slice(6, 8), perempuan = tgl > 40;
    if (perempuan) tgl -= 40;
    return { perempuan: perempuan, tanggal: (tgl < 10 ? "0" : "") + tgl + "-" + n.slice(8, 10) + "-" + n.slice(10, 12),
      sapaan: perempuan ? "Nyonya/Nona" : "Tuan" };
  }
  var api = { cekNIK: cekNIK, infoNIK: infoNIK };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.AktaraValidasi = api;
})(this);
