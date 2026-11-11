Attribute VB_Name = "AktaraTemplate"
' DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual.
Option Explicit

Public Function DaftarTemplate() As Variant
    DaftarTemplate = Array("kuasa", "jualbeli", "skmht")
End Function

' Mengisi judul(), isi() dan sifat untuk kode template. False bila kode tidak dikenal.
Public Function AmbilTemplate(ByVal kode As String, ByRef judul() As String, ByRef isi() As String, ByRef sifat As String) As Boolean
    Select Case kode
    Case "kuasa"
        sifat = "Kuasa"
        ReDim judul(0 To 0)
        judul(0) = "KUASA"
        ReDim isi(0 To 9)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "{SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pemberi Kuasa."
        isi(5) = "Penghadap dikenal oleh saya, Notaris."
        isi(6) = "Penghadap menerangkan dengan ini memberi kuasa kepada {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK2};"
        isi(7) = "- selanjutnya disebut Penerima Kuasa."
        isi(8) = "^KHUSUS"
        isi(9) = "Untuk {OBJEK_KUASA}."
    Case "jualbeli"
        sifat = "Perjanjian Pengikatan Jual Beli"
        ReDim judul(0 To 0)
        judul(0) = "PERJANJIAN PENGIKATAN JUAL BELI"
        ReDim isi(0 To 8)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "I. {SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK1};"
        isi(4) = "- selanjutnya disebut Pihak Pertama."
        isi(5) = "II. {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, Warga Negara Indonesia, {PEKERJAAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK2};"
        isi(6) = "- selanjutnya disebut Pihak Kedua."
        isi(7) = "Para penghadap dikenal oleh saya, Notaris."
        isi(8) = "Para penghadap menerangkan bahwa Pihak Pertama mengikatkan diri untuk menjual kepada Pihak Kedua, yang mengikatkan diri untuk membeli dari Pihak Pertama, {OBJEK}, dengan harga sebesar {HARGA}."
    Case "skmht"
        sifat = "Surat Kuasa Membebankan Hak Tanggungan"
        ReDim judul(0 To 1)
        judul(0) = "SURAT KUASA"
        judul(1) = "MEMBEBANKAN HAK TANGGUNGAN"
        ReDim isi(0 To 23)
        isi(0) = "Pada hari ini, {HARI}, tanggal {TANGGAL}."
        isi(1) = "Pukul {JAM}."
        isi(2) = "Hadir dihadapan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya kenal dan akan disebut pada bagian akhir akta ini :"
        isi(3) = "I. {SAPAAN_PIHAK1} {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, {PEKERJAAN_PIHAK1}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK1}, {PERSETUJUAN_PASANGAN};"
        isi(4) = "- bertempat tinggal di {ALAMAT_PIHAK1}."
        isi(5) = "Pemegang hak atas tanah yang akan dijadikan Obyek Hak Tanggungan, selanjutnya disebut Pemberi Kuasa."
        isi(6) = "II. {SAPAAN_PIHAK2} {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, {JABATAN_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Kartu Tanda Penduduk Nomor {NIK_PIHAK2};"
        isi(7) = "- menurut keterangannya dalam hal ini bertindak dalam jabatannya tersebut berdasarkan {DASAR_KEWENANGAN}, oleh karena itu sah bertindak untuk dan atas nama {KREDITOR}, berkedudukan di {KEDUDUKAN_KREDITOR};"
        isi(8) = "- selanjutnya disebut Penerima Kuasa."
        isi(9) = "Para penghadap dikenal oleh saya, Notaris."
        isi(10) = "Pemberi Kuasa menerangkan dengan ini memberi kuasa kepada Penerima Kuasa :"
        isi(11) = "^KHUSUS"
        isi(12) = "Untuk membebankan Hak Tanggungan guna menjamin pelunasan utang {DEBITOR}, selaku Debitor, sejumlah {NILAI_UTANG} / sejumlah uang yang dapat ditentukan dikemudian hari berdasarkan perjanjian utang-piutang dengan {KREDITOR} selaku Kreditor, yang dibuktikan dengan {PERJANJIAN_KREDIT}, berikut penambahan, perubahan, perpanjangan serta pembaruannya, sampai sejumlah Nilai Tanggungan Peringkat {PERINGKAT} sebesar {NILAI_TANGGUNGAN}, atas Obyek Hak Tanggungan berupa hak atas tanah yang diuraikan dibawah ini :"
        isi(13) = "- {JENIS_HAK} Nomor {NOMOR_HAK}, sebagaimana diuraikan dalam Surat Ukur {SURAT_UKUR}, seluas {LUAS}, dengan Nomor Identifikasi Bidang Tanah (NIB) : {NIB}, menurut Sertipikat (Tanda Bukti Hak) dari {KANTOR_PERTANAHAN}, dan Nomor Obyek Pajak (NOP) : {NOP}, terletak di :"
        isi(14) = "- Provinsi : {PROVINSI},"
        isi(15) = "- Kabupaten/Kota : {KABUPATEN},"
        isi(16) = "- Kecamatan : {KECAMATAN},"
        isi(17) = "- Desa/Kelurahan : {DESA},"
        isi(18) = "- Blok/Jalan : {JALAN}."
        isi(19) = "Yang diperoleh oleh Pemegang Hak berdasarkan {DASAR_PEROLEHAN}."
        isi(20) = "Sertipikat dan bukti pemilikan tersebut diperlihatkan kepada saya, Notaris, untuk keperluan pembuatan Surat Kuasa Membebankan Hak Tanggungan ini."
        isi(21) = "Obyek Hak Tanggungan ini meliputi pula bangunan, tanaman dan hasil karya yang telah ada atau akan ada yang merupakan satu kesatuan dengan tanah tersebut."
        isi(22) = "Kuasa ini meliputi kuasa untuk menghadap dimana perlu, memberikan keterangan, memperlihatkan dan menyerahkan surat-surat, menandatangani Akta Pemberian Hak Tanggungan serta surat-surat lain yang diperlukan, memilih domisili, mendaftarkan Hak Tanggungan, serta memberikan dan menyetujui janji-janji sebagaimana dimaksud dalam Pasal 11 ayat (2) Undang-Undang Nomor 4 Tahun 1996 tentang Hak Tanggungan, yang disetujui oleh Pemberi Kuasa."
        isi(23) = "Kuasa ini tidak dapat ditarik kembali dan tidak berakhir oleh sebab apapun juga, kecuali karena telah dilaksanakan atau karena telah habis jangka waktunya sebagaimana diatur dalam Pasal 15 Undang-Undang Nomor 4 Tahun 1996."
    Case Else
        Exit Function
    End Select
    AmbilTemplate = True
End Function
