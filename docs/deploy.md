# Deploy Aktara (paket gratis, tanpa kartu kredit)

| Bagian | Layanan | Alamat |
|---|---|---|
| File add-in Edisi Modern | GitHub Pages (otomatis lewat `.github/workflows/pages.yml`) | `https://fwmochamad.github.io/visual-notaris/` |
| Server API | Render, paket free, region Singapura (`render.yaml`) | `https://aktara-api.onrender.com` |
| Database | Neon Postgres, paket free, region Singapura | via `DATABASE_URL` |

## 1. Neon
1. Daftar di neon.com (login GitHub), buat project region **AWS Asia Pacific (Singapore)**.
2. Salin *connection string* (format `postgresql://...?sslmode=require`). Jangan disimpan di repo.

## 2. Render
1. Daftar di render.com (login GitHub) › **New › Blueprint** › pilih repo `visual-notaris` (file `render.yaml` terbaca otomatis).
2. Isi `DATABASE_URL` dengan connection string Neon. `AKTARA_SECRET` dan `AKTARA_ADMIN_KEY` dibuat otomatis — lihat nilainya di tab Environment.
3. Setelah live, cek `https://aktara-api.onrender.com/api/health` → `"db":"postgres"`.
   Bila nama `aktara-api` sudah dipakai orang lain, Render memberi nama lain: set variabel repo GitHub `AKTARA_API` ke alamat tersebut.

## 3. GitHub Pages
Settings › Pages › Source: **GitHub Actions**. Setiap push ke `addin/` menerbitkan ulang otomatis.
Manifest untuk dipasang di Word: `https://fwmochamad.github.io/visual-notaris/manifest.xml`.

## 4. Membuat lisensi pelanggan
```
curl -X POST https://aktara-api.onrender.com/api/admin/kantor \
  -H "X-Admin-Key: <AKTARA_ADMIN_KEY>" -H "Content-Type: application/json" \
  -d '{"nama":"Kantor Notaris Contoh","berlaku_sampai":"2027-12-31","maks_perangkat":3}'
```

## Batasan paket gratis
- Server Render tidur setelah 15 menit tanpa trafik; bangun ±1 menit. Add-in tetap jalan (nomor SEMENTARA, lisensi offline 30 hari).
- Data tersimpan di Singapura. Sebelum memakai data klien sungguhan, pindahkan ke cloud lokasi Indonesia.
