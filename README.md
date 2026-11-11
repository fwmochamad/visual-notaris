# Aktara

Add-in Microsoft Word untuk kantor Notaris dan PPAT di Indonesia: template akta, terbilang,
data para pihak, garis akta, minuta → salinan, dan penomoran akta. Langsung di Word,
tanpa aplikasi baru.

## Struktur
| Folder | Isi |
|---|---|
| `addin/` | Add-in Office.js (WordApi 1.1, ES5) — Word 2016, 2019, 2021, 2024, 365, Mac, Web |
| `server/` | API lisensi, nomor akta terpusat, repertorium (Node.js, tanpa dependensi) |
| `docs/` | Spesifikasi produk |

## Mulai
```
cd addin && npm test
cd server && npm test
```
Cara memasang add-in ke Word: lihat `addin/README.md`.

Status: v0.1 — add-in tersambung ke server untuk aktivasi lisensi, nomor akta, dan repertorium; mode offline tetap jalan.
