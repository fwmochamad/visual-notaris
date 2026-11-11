Attribute VB_Name = "AktaraTemplate"
' DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual.
Option Explicit

Public Function DaftarTemplate() As Variant
    DaftarTemplate = Array("kuasa", "kuasamenjual", "jualbeli", "sewa", "pendiriancv", "pernyataan")
End Function

Public Function NamaTemplate(ByVal kode As String) As String
    Select Case kode
    Case "kuasa": NamaTemplate = "Kuasa"
    Case "kuasamenjual": NamaTemplate = "Kuasa Menjual"
    Case "jualbeli": NamaTemplate = "Perjanjian Pengikatan Jual Beli (PPJB)"
    Case "sewa": NamaTemplate = "Sewa Menyewa"
    Case "pendiriancv": NamaTemplate = "Pendirian Persekutuan Komanditer (CV)"
    Case "pernyataan": NamaTemplate = "Pernyataan"
    End Select
End Function

' Label lengkap (tanpa singkatan) untuk setiap isian
Public Function LabelIsian(ByVal tag As String) As String
    Select Case tag
    Case "HARI": LabelIsian = "Hari Penandatanganan"
    Case "TANGGAL": LabelIsian = "Tanggal Penandatanganan (dalam huruf)"
    Case "JAM": LabelIsian = "Pukul Penandatanganan (Waktu Indonesia Barat/Tengah/Timur)"
    Case "NOTARIS": LabelIsian = "Nama Lengkap Notaris beserta Gelar"
    Case "KEDUDUKAN": LabelIsian = "Tempat Kedudukan Notaris (Kabupaten/Kota)"
    Case "SAPAAN_PIHAK1": LabelIsian = "Sapaan Pihak Pertama (Tuan/Nyonya/Nona)"
    Case "NAMA_PIHAK1": LabelIsian = "Nama Lengkap Pihak Pertama"
    Case "TTL_PIHAK1": LabelIsian = "Tempat dan Tanggal Lahir Pihak Pertama"
    Case "PEKERJAAN_PIHAK1": LabelIsian = "Pekerjaan Pihak Pertama"
    Case "ALAMAT_PIHAK1": LabelIsian = "Alamat Tempat Tinggal Pihak Pertama"
    Case "NIK_PIHAK1": LabelIsian = "Nomor Induk Kependudukan (NIK) Pihak Pertama"
    Case "SAPAAN_PIHAK2": LabelIsian = "Sapaan Pihak Kedua (Tuan/Nyonya/Nona)"
    Case "NAMA_PIHAK2": LabelIsian = "Nama Lengkap Pihak Kedua"
    Case "TTL_PIHAK2": LabelIsian = "Tempat dan Tanggal Lahir Pihak Kedua"
    Case "PEKERJAAN_PIHAK2": LabelIsian = "Pekerjaan Pihak Kedua"
    Case "ALAMAT_PIHAK2": LabelIsian = "Alamat Tempat Tinggal Pihak Kedua"
    Case "NIK_PIHAK2": LabelIsian = "Nomor Induk Kependudukan (NIK) Pihak Kedua"
    Case "OBJEK_KUASA": LabelIsian = "Hal yang Dikuasakan"
    Case "OBJEK": LabelIsian = "Uraian Objek"
    Case "HARGA": LabelIsian = "Harga (angka dan huruf)"
    Case "SERTIPIKAT": LabelIsian = "Jenis dan Nomor Sertipikat Hak atas Tanah"
    Case "LUAS": LabelIsian = "Luas Tanah (Meter persegi, angka dan huruf)"
    Case "LETAK": LabelIsian = "Letak Tanah/Bangunan (alamat lengkap)"
    Case "NOMOR_OBJEK_PAJAK": LabelIsian = "Nomor Objek Pajak (NOP) Pajak Bumi dan Bangunan (PBB)"
    Case "JANGKA_WAKTU": LabelIsian = "Jangka Waktu Sewa"
    Case "TANGGAL_MULAI": LabelIsian = "Tanggal Mulai Sewa"
    Case "TANGGAL_AKHIR": LabelIsian = "Tanggal Berakhir Sewa"
    Case "PERUNTUKAN": LabelIsian = "Peruntukan Objek Sewa"
    Case "NAMA_CV": LabelIsian = "Nama Persekutuan Komanditer (Commanditaire Vennootschap/CV)"
    Case "KEDUDUKAN_CV": LabelIsian = "Tempat Kedudukan Persekutuan Komanditer (CV)"
    Case "MAKSUD_TUJUAN": LabelIsian = "Maksud dan Tujuan Usaha"
    Case "MODAL": LabelIsian = "Jumlah Modal (angka dan huruf)"
    Case "SETORAN_PIHAK1": LabelIsian = "Setoran Modal Pihak Pertama (angka dan huruf)"
    Case "SETORAN_PIHAK2": LabelIsian = "Setoran Modal Pihak Kedua (angka dan huruf)"
    Case "ISI_PERNYATAAN": LabelIsian = "Isi Pernyataan"
    Case Else: LabelIsian = Replace(tag, "_", " ")
    End Select
End Function

' Mengisi judul(), isi() dan sifat untuk kode template. False bila kode tidak dikenal.
Public Function AmbilTemplate(ByVal kode As String, ByRef judul() As String, ByRef isi() As String, ByRef sifat As String) As Boolean
    Select Case kode
    Case "kuasa"
        sifat = "Kuasa"
        ReDim judul(0 To 0)
        judul(0) = "KUASA"
        ReDim isi(0 To 11)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "{SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pemberi Kuasa."
        isi(5) = "Penghadap dikenal oleh saya, Notaris."
        isi(6) = "Penghadap menerangkan dengan ini memberi kuasa kepada {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK2};"
        isi(7) = "- selanjutnya disebut Penerima Kuasa."
        isi(8) = "^KHUSUS"
        isi(9) = "Untuk {OBJEK_KUASA}."
        isi(10) = "Untuk keperluan tersebut, Penerima Kuasa berhak menghadap pejabat dan instansi yang berwenang, memberikan keterangan, membuat, menandatangani dan menyerahkan surat-surat yang diperlukan, serta melakukan segala tindakan yang dianggap perlu dan berguna untuk tercapainya maksud pemberian kuasa ini, tidak ada satu tindakan pun yang dikecualikan."
        isi(11) = "Kuasa ini diberikan dengan hak untuk melimpahkan kuasa ini kepada orang lain (substitusi)."
    Case "kuasamenjual"
        sifat = "Kuasa Menjual"
        ReDim judul(0 To 0)
        judul(0) = "KUASA MENJUAL"
        ReDim isi(0 To 11)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "{SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pemberi Kuasa."
        isi(5) = "Penghadap dikenal oleh saya, Notaris."
        isi(6) = "Penghadap menerangkan dengan ini memberi kuasa kepada {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK2};"
        isi(7) = "- selanjutnya disebut Penerima Kuasa."
        isi(8) = "^KHUSUS"
        isi(9) = "Untuk dan atas nama Pemberi Kuasa menjual, melepaskan dan menyerahkan kepada siapa pun juga, termasuk kepada Penerima Kuasa sendiri, dengan harga dan syarat-syarat yang dianggap baik oleh Penerima Kuasa, atas sebidang tanah {SERTIPIKAT}, seluas {LUAS}, terletak di {LETAK}, dengan Nomor Objek Pajak {NOMOR_OBJEK_PAJAK}."
        isi(10) = "Untuk keperluan tersebut, Penerima Kuasa berhak menghadap Pejabat Pembuat Akta Tanah dan instansi yang berwenang, menerima uang harga penjualan dan memberikan tanda penerimaannya, menandatangani akta jual beli dan surat-surat lain yang diperlukan, serta melakukan segala tindakan yang dianggap perlu dan berguna."
        isi(11) = "Kuasa ini tidak dapat ditarik kembali dan tidak berakhir karena sebab-sebab yang tercantum dalam Pasal 1813 Kitab Undang-Undang Hukum Perdata."
    Case "jualbeli"
        sifat = "Perjanjian Pengikatan Jual Beli"
        ReDim judul(0 To 0)
        judul(0) = "PERJANJIAN PENGIKATAN JUAL BELI"
        ReDim isi(0 To 8)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "I. {SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pihak Pertama."
        isi(5) = "II. {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK2};"
        isi(6) = "- selanjutnya disebut Pihak Kedua."
        isi(7) = "Para penghadap dikenal oleh saya, Notaris."
        isi(8) = "Para penghadap menerangkan bahwa Pihak Pertama mengikatkan diri untuk menjual kepada Pihak Kedua, yang mengikatkan diri untuk membeli dari Pihak Pertama, {OBJEK}, dengan harga sebesar {HARGA}."
    Case "sewa"
        sifat = "Perjanjian Sewa Menyewa"
        ReDim judul(0 To 0)
        judul(0) = "PERJANJIAN SEWA MENYEWA"
        ReDim isi(0 To 12)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "I. {SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pihak Yang Menyewakan."
        isi(5) = "II. {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK2};"
        isi(6) = "- selanjutnya disebut Pihak Penyewa."
        isi(7) = "Para penghadap dikenal oleh saya, Notaris."
        isi(8) = "Para penghadap menerangkan bahwa Pihak Yang Menyewakan dengan ini menyewakan kepada Pihak Penyewa, yang menerangkan menyewa dari Pihak Yang Menyewakan, {OBJEK}, terletak di {LETAK}, untuk dipergunakan sebagai {PERUNTUKAN}."
        isi(9) = "^Pasal 1"
        isi(10) = "Sewa menyewa ini dilangsungkan untuk jangka waktu {JANGKA_WAKTU}, terhitung sejak tanggal {TANGGAL_MULAI} dan berakhir pada tanggal {TANGGAL_AKHIR}."
        isi(11) = "^Pasal 2"
        isi(12) = "Uang sewa untuk seluruh jangka waktu tersebut disepakati sebesar {HARGA}, yang telah dibayar lunas oleh Pihak Penyewa kepada Pihak Yang Menyewakan, dan akta ini berlaku pula sebagai tanda penerimaan yang sah (kuitansi)."
    Case "pendiriancv"
        sifat = "Pendirian Persekutuan Komanditer"
        ReDim judul(0 To 1)
        judul(0) = "PENDIRIAN"
        judul(1) = "PERSEKUTUAN KOMANDITER"
        ReDim isi(0 To 15)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "I. {SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "II. {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK2};"
        isi(5) = "Para penghadap dikenal oleh saya, Notaris."
        isi(6) = "Para penghadap menerangkan bahwa dengan ini mereka mendirikan suatu Persekutuan Komanditer (Commanditaire Vennootschap) dengan anggaran dasar sebagai berikut :"
        isi(7) = "^Pasal 1"
        isi(8) = "^NAMA DAN TEMPAT KEDUDUKAN"
        isi(9) = "Persekutuan ini bernama {NAMA_CV} dan berkedudukan di {KEDUDUKAN_CV}."
        isi(10) = "^Pasal 2"
        isi(11) = "^MAKSUD DAN TUJUAN"
        isi(12) = "Maksud dan tujuan persekutuan ialah {MAKSUD_TUJUAN}."
        isi(13) = "^Pasal 3"
        isi(14) = "^MODAL"
        isi(15) = "Modal persekutuan ditetapkan sebesar {MODAL}, yang disetor oleh Pihak Pertama sebesar {SETORAN_PIHAK1} selaku sekutu pengurus (sekutu komplementer), dan oleh Pihak Kedua sebesar {SETORAN_PIHAK2} selaku sekutu diam (sekutu komanditer)."
    Case "pernyataan"
        sifat = "Pernyataan"
        ReDim judul(0 To 0)
        judul(0) = "PERNYATAAN"
        ReDim isi(0 To 7)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "{SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk dengan Nomor Induk Kependudukan {NIK_PIHAK1};"
        isi(4) = "Penghadap dikenal oleh saya, Notaris."
        isi(5) = "Penghadap menerangkan dengan ini menyatakan sebagai berikut :"
        isi(6) = "{ISI_PERNYATAAN}"
        isi(7) = "Penghadap menyatakan bertanggung jawab sepenuhnya atas kebenaran pernyataan tersebut di atas, dan membebaskan Notaris serta para saksi dari segala tuntutan apa pun yang timbul di kemudian hari."
    Case Else
        Exit Function
    End Select
    AmbilTemplate = True
End Function
