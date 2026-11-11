var assert = require("assert");
var T = require("../src/terbilang.js");
var cases = [
  [0, "nol"], [1, "satu"], [10, "sepuluh"], [11, "sebelas"], [12, "dua belas"], [19, "sembilan belas"],
  [20, "dua puluh"], [99, "sembilan puluh sembilan"], [100, "seratus"], [111, "seratus sebelas"],
  [250, "dua ratus lima puluh"], [1000, "seribu"], [1001, "seribu satu"], [2026, "dua ribu dua puluh enam"],
  [11000, "sebelas ribu"], [100000, "seratus ribu"], [1500000, "satu juta lima ratus ribu"],
  [1000000000, "satu miliar"], [2500000000000, "dua triliun lima ratus miliar"]
];
cases.forEach(function (c) { assert.strictEqual(T.terbilang(c[0]), c[1], c[0]); });
assert.strictEqual(T.tanggal(new Date(2026, 10, 11)), "11-11-2026 (sebelas November dua ribu dua puluh enam)");
assert.strictEqual(T.jam(10, 30), "10.30 WIB (sepuluh lebih tiga puluh menit Waktu Indonesia Barat)");
assert.strictEqual(T.rupiah(1500000), "Rp.1.500.000,- (satu juta lima ratus ribu rupiah)");
assert.strictEqual(T.rupiah(367000000), "Rp.367.000.000,- (tiga ratus enam puluh tujuh juta rupiah)");
assert.strictEqual(T.jam(12, 30), "12.30 WIB (dua belas lebih tiga puluh menit Waktu Indonesia Barat)");
assert.strictEqual(T.luas(72), "72 M\u00b2 (tujuh puluh dua Meter persegi)");
assert.strictEqual(T.nomorAkta(22), "22.-");
console.log("Semua " + (cases.length + 7) + " tes lulus");
