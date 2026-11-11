# Aktara

Add-in Microsoft Word untuk kantor Notaris dan PPAT di Indonesia: template akta, terbilang,
data para pihak, garis akta, minuta → salinan, dan penomoran akta. Langsung di Word,
tanpa aplikasi baru.

## Struktur
| Folder | Isi |
|---|---|
| `addin/` | **Edisi Modern** — add-in Office.js (WordApi 1.1, ES5): Word 2016, 2019, 2021, 2024, 365, Mac, Web |
| `klasik/` | **Edisi Klasik** — template global VBA (.dotm): Word 2007, 2010, 2013 (dan 2016+ Windows) |
| `server/` | API lisensi, nomor akta terpusat, repertorium (Node.js, tanpa dependensi) |
| `docs/` | Spesifikasi produk |

## Mulai
```
cd addin && npm test
cd server && npm test
```
## Uji di Word (Windows)
```
npm run certs      # sekali saja: pasang sertifikat HTTPS lokal
npm run dev        # nyalakan add-in (port 3000) + API (port 8080)
npm run sideload   # di jendela lain: daftarkan add-in ke Word dan buka Word
```
Di Word: tab Home → **Aktara** → tab Lisensi → kode `AKT-DEV0-0000-0000-0001` → Aktifkan.

Status: v0.1 — add-in tersambung ke server untuk aktivasi lisensi, nomor akta, dan repertorium; mode offline tetap jalan.
