/* Terbilang Bahasa Indonesia untuk akta notaris. Kompatibel ES5 (IE11 / Word 2016 Windows). */
(function (root) {
  "use strict";
  var SATUAN = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"];
  var BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var SKALA = [[1e12, "triliun"], [1e9, "miliar"], [1e6, "juta"]];

  function bilang(n) {
    n = Math.floor(n);
    if (n < 12) return SATUAN[n];
    if (n < 20) return bilang(n - 10) + " belas";
    if (n < 100) return join(bilang(Math.floor(n / 10)) + " puluh", bilang(n % 10));
    if (n < 200) return join("seratus", bilang(n - 100));
    if (n < 1000) return join(bilang(Math.floor(n / 100)) + " ratus", bilang(n % 100));
    if (n < 2000) return join("seribu", bilang(n - 1000));
    if (n < 1e6) return join(bilang(Math.floor(n / 1000)) + " ribu", bilang(n % 1000));
    for (var i = 0; i < SKALA.length; i++) {
      if (n >= SKALA[i][0]) return join(bilang(Math.floor(n / SKALA[i][0])) + " " + SKALA[i][1], bilang(n % SKALA[i][0]));
    }
    throw new Error("Angka terlalu besar");
  }
  function join(a, b) { return b ? a + " " + b : a; }

  function terbilang(n) {
    n = Number(n);
    if (isNaN(n) || n < 0) throw new Error("Angka tidak valid");
    return n === 0 ? "nol" : bilang(n);
  }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function ribuan(n) { return String(Math.floor(n)).replace(/\B(?=(\d{3})+(?!\d))/g, "."); }

  /* "11-11-2026 (sebelas November dua ribu dua puluh enam)" */
  function tanggal(d) {
    return pad(d.getDate()) + "-" + pad(d.getMonth() + 1) + "-" + d.getFullYear() +
      " (" + terbilang(d.getDate()) + " " + BULAN[d.getMonth()] + " " + terbilang(d.getFullYear()) + ")";
  }
  /* "10.30 WIB (sepuluh lewat tiga puluh menit Waktu Indonesia Barat)" */
  function jam(h, m, zona) {
    var Z = { WIB: "Waktu Indonesia Barat", WITA: "Waktu Indonesia Tengah", WIT: "Waktu Indonesia Timur" };
    zona = zona || "WIB";
    return pad(h) + "." + pad(m) + " " + zona + " (" + (h === 0 ? "nol" : terbilang(h)) +
      " lewat " + (m === 0 ? "nol" : terbilang(m)) + " menit " + Z[zona] + ")";
  }
  /* "Rp1.500.000,00 (satu juta lima ratus ribu Rupiah)" */
  function rupiah(n) {
    return "Rp" + ribuan(n) + ",00 (" + terbilang(n) + " Rupiah)";
  }
  function angka(n) { return n + " (" + terbilang(n) + ")"; }

  var api = { terbilang: terbilang, tanggal: tanggal, jam: jam, rupiah: rupiah, angka: angka, BULAN: BULAN };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.Terbilang = api;
})(this);
