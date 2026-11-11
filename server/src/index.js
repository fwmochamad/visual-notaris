"use strict";
const { buka } = require("./db");
const { buatApp } = require("./app");

const { AKTARA_SECRET, AKTARA_ADMIN_KEY, AKTARA_DB = "aktara.db", PORT = 8080 } = process.env;
if (!AKTARA_SECRET || !AKTARA_ADMIN_KEY) {
  console.error("Set AKTARA_SECRET dan AKTARA_ADMIN_KEY terlebih dahulu.");
  process.exit(1);
}
buatApp({ db: buka(AKTARA_DB), rahasia: AKTARA_SECRET, kunciAdmin: AKTARA_ADMIN_KEY })
  .listen(PORT, () => console.log("Aktara API berjalan di port " + PORT));
