Attribute VB_Name = "AktaraTemplate"
' DIHASILKAN OTOMATIS dari templates/akta.json oleh scripts/gen_templates.py - jangan diedit manual.
Option Explicit

Public Function DaftarTemplate() As Variant
    DaftarTemplate = Array("kuasa", "jualbeli")
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
    Case Else
        Exit Function
    End Select
    AmbilTemplate = True
End Function
