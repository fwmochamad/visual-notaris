# Aktara — Prototipe Add-in Word (v0.1)

Add-in Word untuk kantor Notaris/PPAT. Hanya memakai **WordApi 1.1**, sehingga berjalan di:
Word 2016, 2019, 2021, 2024, Microsoft 365 (Windows & Mac), dan Word Online.
Kode ditulis ES5 agar kompatibel dengan engine IE11 di Word 2016 Windows.

## Fitur prototipe
| Tab | Fungsi |
|---|---|
| Akta | Template akta (Kuasa, PPJB) dengan isian otomatis, penomoran akta per bulan, ubah minuta jadi SALINAN |
| Terbilang | Tanggal, jam (WIB/WITA/WIT), Rupiah, dan angka biasa dalam huruf |
| Pihak | Isi data notaris & para pihak sekali, terisi di semua tempat dalam akta |
| Garis | Garis strip pengisi akhir baris (tab-leader), menyesuaikan otomatis saat teks diedit |

## Mencoba di komputer sendiri
1. Pasang Node.js, lalu di folder ini jalankan:
   `npx office-addin-dev-certs install` lalu `npx http-server src -p 3000 -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key`
2. Tambahkan ikon `icon-16.png`, `icon-32.png`, `icon-80.png` ke folder `src` (sementara boleh gambar apa saja).
3. Sideload `manifest.xml`:
   - **Windows**: taruh manifest di folder yang dibagikan (shared folder), lalu di Word: File › Options › Trust Center › Trusted Add-in Catalogs, tambahkan alamat folder tersebut. Buka Insert › My Add-ins › Shared Folder.
   - **Word Online / 365**: Insert › Add-ins › Upload My Add-in › pilih `manifest.xml`.
   - **Mac**: salin manifest ke `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`.
4. Tombol **Aktara** muncul di tab Home.

## Tes
`node test/terbilang.test.js` — 22 tes konversi terbilang.

## Catatan batasan
- Word 2010/2013 tidak didukung oleh teknologi add-in modern; lihat dokumen spesifikasi untuk jalur pendamping (VSTO).
- Nomor akta disimpan lokal per komputer; versi produksi memakai server lisensi/sinkronisasi kantor.
