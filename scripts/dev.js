// Mode pengembangan lokal: satu perintah menyalakan
//   - hosting add-in  https://localhost:3000  (folder addin/src)
//   - Aktara API      https://localhost:8080  (database dev.db, lisensi uji otomatis dibuat)
// Sertifikat: jalankan sekali `npx office-addin-dev-certs install`.
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const os = require("node:os");
const https = require("node:https");
const { buka } = require("../server/src/db");
const { buatApp } = require("../server/src/app");

const dir = path.join(os.homedir(), ".office-addin-dev-certs");
const certPath = path.join(dir, "localhost.crt"), keyPath = path.join(dir, "localhost.key");
if (!fs.existsSync(certPath)) {
  console.error("Sertifikat belum ada. Jalankan dulu:  npx office-addin-dev-certs install");
  process.exit(1);
}
const tls = { cert: fs.readFileSync(certPath), key: fs.readFileSync(keyPath) };

// 1. Hosting file add-in
const root = path.join(__dirname, "..", "addin", "src");
const MIME = { ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8", ".png": "image/png", ".css": "text/css" };
https.createServer(tls, (req, res) => {
  const p = path.normalize(decodeURIComponent(new URL(req.url, "https://x").pathname)).replace(/^([\\/]\.\.)+/, "");
  const file = path.join(root, p === path.sep || p === "/" ? "taskpane.html" : p);
  if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); return res.end("Tidak ditemukan"); }
  res.writeHead(200, { "Content-Type": MIME[path.extname(file)] || "application/octet-stream", "Cache-Control": "no-store" });
  fs.createReadStream(file).pipe(res);
}).listen(3000, () => console.log("Add-in   : https://localhost:3000/taskpane.html"));

// 2. API dengan lisensi uji
const db = buka(path.join(__dirname, "..", "server", "dev.db"));
const KODE = "AKT-DEV0-0000-0000-0001";
if (!db.prepare("SELECT 1 FROM kantor WHERE kode_lisensi = ?").get(KODE)) {
  db.prepare("INSERT INTO kantor (nama, kode_lisensi, paket, maks_perangkat, berlaku_sampai) VALUES (?,?,?,?,?)")
    .run("Kantor Notaris Uji Coba", KODE, "kantor", 5, "2099-12-31");
}
const app = buatApp({ db, rahasia: "rahasia-dev", kunciAdmin: "admin-dev" });
https.createServer(tls, app.listeners("request")[0]).listen(8080, () => {
  console.log("API      : https://localhost:8080/api/health");
  console.log("Kode lisensi uji: " + KODE);
});
