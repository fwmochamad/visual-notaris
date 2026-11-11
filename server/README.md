# Aktara Server v0.1

API untuk lisensi, nomor akta terpusat, dan repertorium. Tanpa dependensi eksternal (Node.js ≥ 22.5).
Pengembangan memakai SQLite bawaan Node; produksi direncanakan PostgreSQL dengan skema yang sama.

## Menjalankan
```
AKTARA_SECRET=ganti-rahasia AKTARA_ADMIN_KEY=ganti-kunci-admin npm start   # port 8080
npm test                                                                  # 11 tes API
```

## Endpoint
| Method | Path | Fungsi |
|---|---|---|
| GET | /api/health | Cek server |
| POST | /api/admin/kantor | (header `X-Admin-Key`) Buat kantor pelanggan + kode lisensi `AKT-XXXX-…` |
| POST | /api/aktivasi | Kode lisensi + id perangkat → token; dibatasi jumlah perangkat per paket |
| GET | /api/lisensi | Cek lisensi; add-in boleh offline 30 hari sejak cek terakhir |
| POST | /api/nomor-akta | Nomor akta berikutnya per kantor per bulan, atomik (tidak ganda antar komputer) |
| POST | /api/repertorium | Catat akta final; nomor ganda ditolak |
| GET | /api/repertorium?tahun=&bulan=&format=csv | Daftar/ekspor repertorium bulanan (CSV dibuka di Excel) |

Semua aksi tercatat di tabel `audit`.

## Contoh membuat pelanggan
```
curl -X POST localhost:8080/api/admin/kantor -H "X-Admin-Key: $AKTARA_ADMIN_KEY" \
  -H "Content-Type: application/json" \
  -d '{"nama":"Kantor Notaris Contoh, S.H., M.Kn.","berlaku_sampai":"2027-12-31","maks_perangkat":3}'
```

## Belum ada (berikutnya)
Portal admin web, database klien, sinkronisasi template, HTTPS/deploy produksi, PostgreSQL.
