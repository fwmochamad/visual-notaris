"use strict";
const fs = require("node:fs");
const https = require("node:https");
const { buka } = require("./db");
const { buatApp } = require("./app");

const { AKTARA_SECRET, AKTARA_ADMIN_KEY, DATABASE_URL, AKTARA_DB = "aktara.db", PORT = 8080, AKTARA_TLS_CERT, AKTARA_TLS_KEY } = process.env;
if (!AKTARA_SECRET || !AKTARA_ADMIN_KEY) {
  console.error("Set AKTARA_SECRET dan AKTARA_ADMIN_KEY terlebih dahulu.");
  process.exit(1);
}

(async () => {
  // DATABASE_URL (Postgres, mis. Neon) untuk produksi; tanpa itu memakai SQLite lokal.
  const db = await buka(DATABASE_URL ? { url: DATABASE_URL } : { path: AKTARA_DB });
  const app = buatApp({ db, rahasia: AKTARA_SECRET, kunciAdmin: AKTARA_ADMIN_KEY });
  const pesan = () => console.log("Aktara API (" + db.mesin + ") di port " + PORT);
  if (AKTARA_TLS_CERT && AKTARA_TLS_KEY) {
    https.createServer({ cert: fs.readFileSync(AKTARA_TLS_CERT), key: fs.readFileSync(AKTARA_TLS_KEY) }, app.listeners("request")[0]).listen(PORT, pesan);
  } else {
    app.listen(PORT, pesan); // di Render, HTTPS ditangani platform
  }
})().catch((e) => { console.error("Gagal memulai server:", e.message); process.exit(1); });
