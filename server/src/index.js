"use strict";
const fs = require("node:fs");
const https = require("node:https");
const { buka } = require("./db");
const { buatApp } = require("./app");

const { AKTARA_SECRET, AKTARA_ADMIN_KEY, AKTARA_DB = "aktara.db", PORT = 8080, AKTARA_TLS_CERT, AKTARA_TLS_KEY } = process.env;
if (!AKTARA_SECRET || !AKTARA_ADMIN_KEY) {
  console.error("Set AKTARA_SECRET dan AKTARA_ADMIN_KEY terlebih dahulu.");
  process.exit(1);
}
const app = buatApp({ db: buka(AKTARA_DB), rahasia: AKTARA_SECRET, kunciAdmin: AKTARA_ADMIN_KEY });
if (AKTARA_TLS_CERT && AKTARA_TLS_KEY) {
  // HTTPS langsung (mis. pengujian lokal dengan add-in). Di produksi biasanya HTTPS ditangani reverse proxy.
  https.createServer({ cert: fs.readFileSync(AKTARA_TLS_CERT), key: fs.readFileSync(AKTARA_TLS_KEY) }, app.listeners("request")[0])
    .listen(PORT, () => console.log("Aktara API (HTTPS) di port " + PORT));
} else {
  app.listen(PORT, () => console.log("Aktara API di port " + PORT));
}
