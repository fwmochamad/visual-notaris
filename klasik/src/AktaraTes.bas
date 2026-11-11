Attribute VB_Name = "AktaraTes"
' Tes Edisi Klasik. Jalankan: Alt+F8 > AktaraJalankanTes > Run.
' Kasus uji sama dengan addin/test/terbilang.test.js agar kedua edisi identik.
Option Explicit

Private gagal As Long, lulus As Long

Private Sub Cek(ByVal hasil As String, ByVal harap As String, ByVal nama As String)
    If hasil = harap Then
        lulus = lulus + 1
    Else
        gagal = gagal + 1
        Debug.Print "GAGAL " & nama & ": dapat '" & hasil & "', harap '" & harap & "'"
    End If
End Sub

Public Sub AktaraJalankanTes()
    gagal = 0: lulus = 0
    Dim kasus As Variant, i As Long
    kasus = Array(0, "nol", 1, "satu", 10, "sepuluh", 11, "sebelas", 12, "dua belas", 19, "sembilan belas", _
        20, "dua puluh", 99, "sembilan puluh sembilan", 100, "seratus", 111, "seratus sebelas", _
        250, "dua ratus lima puluh", 1000, "seribu", 1001, "seribu satu", 2026, "dua ribu dua puluh enam", _
        11000, "sebelas ribu", 100000, "seratus ribu", 1500000, "satu juta lima ratus ribu", _
        1000000000#, "satu miliar", 2500000000000#, "dua triliun lima ratus miliar")
    For i = 0 To UBound(kasus) Step 2
        Cek AktaraTerbilang.Terbilang(kasus(i)), kasus(i + 1), CStr(kasus(i))
    Next
    Cek AktaraTerbilang.TerbilangTanggal(DateSerial(2026, 11, 11)), "11-11-2026 (sebelas November dua ribu dua puluh enam)", "tanggal"
    Cek AktaraTerbilang.TerbilangJam(10, 30), "10.30 WIB (sepuluh lebih tiga puluh menit Waktu Indonesia Barat)", "jam"
    Cek AktaraTerbilang.TerbilangRupiah(1500000), "Rp.1.500.000,- (satu juta lima ratus ribu rupiah)", "rupiah"
    Cek AktaraTerbilang.TerbilangRupiah(367000000), "Rp.367.000.000,- (tiga ratus enam puluh tujuh juta rupiah)", "rupiah 2"
    Cek AktaraTerbilang.TerbilangJam(12, 30), "12.30 WIB (dua belas lebih tiga puluh menit Waktu Indonesia Barat)", "jam 2"
    Cek AktaraTerbilang.TerbilangLuas(72), "72 M" & ChrW(178) & " (tujuh puluh dua Meter persegi)", "luas"
    Cek AktaraTerbilang.FormatNomorAkta(22), "22.-", "nomor"
    Cek AktaraServer.JsonAmbil("{""token"":""a.b"",""nomor"":7,""kantor"":""Kantor \""X\""""}", "nomor"), "7", "json angka"
    Cek AktaraServer.JsonAmbil("{""token"":""a.b"",""nomor"":7,""kantor"":""Kantor \""X\""""}", "kantor"), "Kantor ""X""", "json teks"
    MsgBox "Tes Aktara: " & lulus & " lulus, " & gagal & " gagal." & IIf(gagal > 0, vbCr & "Lihat Immediate Window (Ctrl+G) di editor VBA.", ""), _
        IIf(gagal > 0, vbExclamation, vbInformation), "Aktara"
End Sub
