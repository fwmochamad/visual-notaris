# Aktara Edisi Klasik (Word 2007, 2010, 2013 — juga jalan di 2016+ Windows)

Add-in berbasis template global Word (`Aktara.dotm`) dengan tab ribbon **Aktara**.
Memakai server API yang sama dengan Edisi Modern (lisensi, nomor akta, repertorium).

| Grup ribbon | Tombol |
|---|---|
| Akta | Akta Kuasa, Akta PPJB, Isi Data Pihak, Jadikan Salinan (dokumen baru), Catat Repertorium |
| Terbilang | Tanggal, Jam, Rupiah, Angka, Terbilangkan Pilihan |
| Garis Akta | Pasang Garis, Hapus Garis |
| Aktara | Lisensi (aktivasi, alamat server) |

## Struktur
- `src/*.bas` — kode VBA (sumber utama, dikelola di git)
- `ribbon/customUI.xml` — definisi tab ribbon
- `build.py` — membangun `dist/Aktara.dotm` (ribbon tanpa makro)

## Merakit Aktara.dotm (sekali per rilis)
1. `python klasik/build.py`
2. Buka `klasik/dist/Aktara.dotm` di Word (File › Open).
3. Alt+F11 › File › Import File… › impor keempat file `klasik/src/*.bas`.
4. Simpan (Ctrl+S), tutup Word.
5. Uji: Alt+F8 › `AktaraJalankanTes` › Run (harus 24 lulus).

## Memasang di komputer kantor
Salin `Aktara.dotm` ke folder STARTUP Word: tekan Win+R, ketik `%APPDATA%\Microsoft\Word\STARTUP`, Enter, lalu tempel file.
Buka Word → tab **Aktara** muncul. Jika muncul peringatan makro, klik **Enable Content**
(produksi: template ditandatangani sertifikat code-signing agar tidak muncul peringatan).

## Uji lokal dengan server dev
Jalankan `npm run dev` di root repo, lalu di Word: Aktara › Lisensi › server `https://localhost:8080`,
kode `AKT-DEV0-0000-0000-0001`.
