Attribute VB_Name = "AktaraTerbilang"
' Aktara Edisi Klasik - terbilang Bahasa Indonesia.
' Port dari addin/src/terbilang.js; keluaran harus identik (dicek oleh AktaraTes).
Option Explicit

Private Function Satuan(ByVal n As Long) As String
    Satuan = Choose(n + 1, "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas")
End Function

Private Function Gabung(ByVal a As String, ByVal b As String) As String
    If Len(b) > 0 Then Gabung = a & " " & b Else Gabung = a
End Function

' Memakai Double agar bisa sampai triliun (Long hanya sampai ~2,1 miliar).
Private Function Bilang(ByVal n As Double) As String
    n = Int(n)
    If n < 12 Then
        Bilang = Satuan(CLng(n))
    ElseIf n < 20 Then
        Bilang = Bilang(n - 10) & " belas"
    ElseIf n < 100 Then
        Bilang = Gabung(Bilang(Int(n / 10)) & " puluh", Bilang(n - Int(n / 10) * 10))
    ElseIf n < 200 Then
        Bilang = Gabung("seratus", Bilang(n - 100))
    ElseIf n < 1000 Then
        Bilang = Gabung(Bilang(Int(n / 100)) & " ratus", Bilang(n - Int(n / 100) * 100))
    ElseIf n < 2000 Then
        Bilang = Gabung("seribu", Bilang(n - 1000))
    ElseIf n < 1000000# Then
        Bilang = Gabung(Bilang(Int(n / 1000)) & " ribu", Bilang(n - Int(n / 1000) * 1000))
    ElseIf n < 1000000000# Then
        Bilang = Gabung(Bilang(Int(n / 1000000#)) & " juta", Bilang(n - Int(n / 1000000#) * 1000000#))
    ElseIf n < 1000000000000# Then
        Bilang = Gabung(Bilang(Int(n / 1000000000#)) & " miliar", Bilang(n - Int(n / 1000000000#) * 1000000000#))
    ElseIf n < 1E+15 Then
        Bilang = Gabung(Bilang(Int(n / 1000000000000#)) & " triliun", Bilang(n - Int(n / 1000000000000#) * 1000000000000#))
    Else
        Err.Raise vbObjectError + 1, "Aktara", "Angka terlalu besar"
    End If
End Function

Public Function Terbilang(ByVal n As Double) As String
    If n < 0 Then Err.Raise vbObjectError + 2, "Aktara", "Angka tidak valid"
    If n = 0 Then Terbilang = "nol" Else Terbilang = Bilang(n)
End Function

Public Function NamaBulan(ByVal m As Integer) As String
    NamaBulan = Choose(m, "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember")
End Function

Private Function Pad2(ByVal n As Long) As String
    Pad2 = Right$("0" & CStr(n), 2)
End Function

Private Function Ribuan(ByVal n As Double) As String
    Dim s As String, hasil As String
    s = Format$(Int(n), "0")
    Do While Len(s) > 3
        hasil = "." & Right$(s, 3) & hasil
        s = Left$(s, Len(s) - 3)
    Loop
    Ribuan = s & hasil
End Function

' "11-11-2026 (sebelas November dua ribu dua puluh enam)"
Public Function TerbilangTanggal(ByVal d As Date) As String
    TerbilangTanggal = Pad2(Day(d)) & "-" & Pad2(Month(d)) & "-" & Year(d) & _
        " (" & Terbilang(Day(d)) & " " & NamaBulan(Month(d)) & " " & Terbilang(Year(d)) & ")"
End Function

' "10.30 WIB (sepuluh lebih tiga puluh menit Waktu Indonesia Barat)"
Public Function TerbilangJam(ByVal jam As Long, ByVal menit As Long, Optional ByVal zona As String = "WIB") As String
    Dim nama As String
    Select Case UCase$(zona)
        Case "WITA": nama = "Waktu Indonesia Tengah"
        Case "WIT": nama = "Waktu Indonesia Timur"
        Case Else: zona = "WIB": nama = "Waktu Indonesia Barat"
    End Select
    TerbilangJam = Pad2(jam) & "." & Pad2(menit) & " " & UCase$(zona) & " (" & Terbilang(jam) & _
        " lebih " & Terbilang(menit) & " menit " & nama & ")"
End Function

' Format praktik akta: "Rp.367.000.000,- (tiga ratus enam puluh tujuh juta rupiah)"
Public Function TerbilangRupiah(ByVal n As Double) As String
    TerbilangRupiah = "Rp." & Ribuan(n) & ",- (" & Terbilang(n) & " rupiah)"
End Function

' "72 M2 (tujuh puluh dua Meter persegi)" - M dengan superskrip 2
Public Function TerbilangLuas(ByVal n As Double) As String
    TerbilangLuas = Format$(n, "0") & " M" & ChrW(178) & " (" & Terbilang(n) & " Meter persegi)"
End Function

' Nomor akta: "22.-"
Public Function FormatNomorAkta(ByVal n As Long) As String
    FormatNomorAkta = CStr(n) & ".-"
End Function

Public Function TerbilangAngka(ByVal n As Double) As String
    TerbilangAngka = Format$(n, "0") & " (" & Terbilang(n) & ")"
End Function

' Desimal: "2,5" -> "dua koma lima" (digit di belakang koma dibaca satu per satu)
Public Function TerbilangDesimal(ByVal teks As String) As String
    Dim p() As String, h As String, i As Long, c As String
    p = Split(Replace(teks, ".", ","), ",")
    h = Terbilang(CDbl(IIf(Len(p(0)) = 0, "0", p(0))))
    If UBound(p) >= 1 Then
        If Len(p(1)) > 0 Then
            h = h & " koma"
            For i = 1 To Len(p(1))
                c = Mid$(p(1), i, 1)
                If c = "0" Then h = h & " nol" Else h = h & " " & Terbilang(CDbl(c))
            Next
        End If
    End If
    TerbilangDesimal = h
End Function

' "12,5% (dua belas koma lima persen)"
Public Function TerbilangPersen(ByVal teks As String) As String
    TerbilangPersen = Replace(teks, ".", ",") & "% (" & TerbilangDesimal(teks) & " persen)"
End Function

' Memeriksa Nomor Induk Kependudukan (NIK). Mengembalikan teks kosong bila sah, atau pesan masalahnya.
Public Function CekNIK(ByVal nik As String) As String
    Dim n As String, i As Long, prov As Long, tgl As Long, bln As Long
    n = Replace(nik, " ", "")
    If Len(n) <> 16 Then CekNIK = "Nomor Induk Kependudukan (NIK) harus 16 digit angka (sekarang " & Len(n) & " karakter).": Exit Function
    For i = 1 To 16
        If Not Mid$(n, i, 1) Like "#" Then CekNIK = "Nomor Induk Kependudukan (NIK) hanya boleh berisi angka.": Exit Function
    Next
    prov = CLng(Left$(n, 2)): tgl = CLng(Mid$(n, 7, 2)): bln = CLng(Mid$(n, 9, 2))
    If prov < 11 Or prov > 96 Then CekNIK = "Dua digit pertama NIK (kode provinsi) tidak dikenal: " & Left$(n, 2) & ".": Exit Function
    If tgl > 40 Then tgl = tgl - 40
    If tgl < 1 Or tgl > 31 Then CekNIK = "Digit ke-7 dan ke-8 NIK (tanggal lahir) tidak valid: " & Mid$(n, 7, 2) & ".": Exit Function
    If bln < 1 Or bln > 12 Then CekNIK = "Digit ke-9 dan ke-10 NIK (bulan lahir) tidak valid: " & Mid$(n, 9, 2) & ".": Exit Function
    CekNIK = ""
End Function
