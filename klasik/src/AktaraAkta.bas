Attribute VB_Name = "AktaraAkta"
' Aktara Edisi Klasik - fitur akta & tombol ribbon. Word 2007, 2010, 2013 (juga jalan di 2016+).
Option Explicit

Private Const STYLE_GARIS As String = "Akta Garis"

' ================= Ribbon callbacks =================
Public Sub AkOnLoad(ribbon As IRibbonUI)
End Sub

Public Sub AkTemplate(control As IRibbonControl): BuatTemplate control.Tag: End Sub
Public Sub AkPersen(control As IRibbonControl): SisipPersen: End Sub
Public Sub AkCek(control As IRibbonControl): PeriksaAkta: End Sub
Public Sub AkLuas(control As IRibbonControl): SisipLuas: End Sub
Public Sub AkHalaman(control As IRibbonControl): AturHalamanAkta: End Sub
Public Sub AkIsiData(control As IRibbonControl): IsiDataPihak: End Sub
Public Sub AkSalinan(control As IRibbonControl): JadikanSalinan: End Sub
Public Sub AkRepertorium(control As IRibbonControl): CatatRepertorium: End Sub
Public Sub AkGaris(control As IRibbonControl): PasangGaris: End Sub
Public Sub AkHapusGaris(control As IRibbonControl): HapusGaris: End Sub
Public Sub AkTanggal(control As IRibbonControl): SisipTanggal: End Sub
Public Sub AkJam(control As IRibbonControl): SisipJam: End Sub
Public Sub AkRupiah(control As IRibbonControl): SisipAngka True: End Sub
Public Sub AkAngka(control As IRibbonControl): SisipAngka False: End Sub
Public Sub AkTerbilangPilihan(control As IRibbonControl): TerbilangkanPilihan: End Sub
Public Sub AkLisensi(control As IRibbonControl): AktaraServer.DialogLisensi: End Sub

' ================= Terbilang =================
Private Sub Sisip(ByVal teks As String)
    Selection.TypeText teks
End Sub

Public Sub SisipTanggal()
    Dim s As String
    s = InputBox("Tanggal (dd-mm-yyyy):", "Aktara - Terbilang tanggal", Format$(Date, "dd-mm-yyyy"))
    If Len(s) = 0 Then Exit Sub
    Dim p() As String: p = Split(Replace(s, "/", "-"), "-")
    If UBound(p) <> 2 Then MsgBox "Format tanggal harus dd-mm-yyyy", vbExclamation, "Aktara": Exit Sub
    On Error GoTo Salah
    Sisip AktaraTerbilang.TerbilangTanggal(DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0))))
    Exit Sub
Salah:
    MsgBox "Tanggal tidak valid", vbExclamation, "Aktara"
End Sub

Public Sub SisipJam()
    Dim s As String, z As String
    s = InputBox("Jam (hh.mm):", "Aktara - Terbilang jam", Format$(Time, "hh.nn"))
    If Len(s) = 0 Then Exit Sub
    z = InputBox("Zona waktu (WIB / WITA / WIT):", "Aktara - Terbilang jam", "WIB")
    Dim p() As String: p = Split(Replace(s, ":", "."), ".")
    If UBound(p) <> 1 Or Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then MsgBox "Format jam harus hh.mm", vbExclamation, "Aktara": Exit Sub
    Sisip AktaraTerbilang.TerbilangJam(CLng(p(0)), CLng(p(1)), z)
End Sub

Private Function HanyaAngka(ByVal s As String) As String
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "#" Then HanyaAngka = HanyaAngka & c
    Next
End Function

Public Sub SisipAngka(ByVal rupiah As Boolean)
    Dim s As String
    s = HanyaAngka(InputBox(IIf(rupiah, "Nominal Rupiah:", "Angka:"), "Aktara - Terbilang"))
    If Len(s) = 0 Then Exit Sub
    If rupiah Then Sisip AktaraTerbilang.TerbilangRupiah(CDbl(s)) Else Sisip AktaraTerbilang.TerbilangAngka(CDbl(s))
End Sub

Public Sub SisipPersen()
    Dim s As String
    s = Trim$(InputBox("Persen (mis. 12,5):", "Aktara - Terbilang persen"))
    If Len(s) = 0 Then Exit Sub
    Sisip AktaraTerbilang.TerbilangPersen(s)
End Sub

' Periksa akta sebelum cetak: isian kosong, paragraf tanpa garis, nomor akta
Public Sub PeriksaAkta()
    Dim cc As ContentControl, p As Paragraph, kosong As String, nKosong As Long, tanpaGaris As Long, adaNomor As Boolean, t As String
    Dim sudah As Object: Set sudah = CreateObject("Scripting.Dictionary")
    For Each cc In ActiveDocument.ContentControls
        If Len(AmbilIsi(cc)) = 0 And Not sudah.Exists(cc.Tag) Then
            sudah.Add cc.Tag, 1: nKosong = nKosong + 1
            kosong = kosong & vbCr & "   - " & cc.Title
        End If
    Next
    For Each p In ActiveDocument.Paragraphs
        t = p.Range.Text
        If Left$(t, 5) = "Nomor" And InStr(t, ":") > 0 Then adaNomor = True
        If Len(t) > 1 And p.Alignment <> 1 And p.Alignment <> 2 Then
            If Right$(Left$(t, Len(t) - 1), 1) <> vbTab Then tanpaGaris = tanpaGaris + 1
        End If
    Next
    Dim pesan As String
    pesan = IIf(adaNomor, "[v] Nomor akta tercantum.", "[!] Nomor akta belum tercantum.") & vbCr
    pesan = pesan & IIf(nKosong = 0, "[v] Semua isian sudah terisi.", "[!] " & nKosong & " isian masih kosong:" & kosong) & vbCr
    pesan = pesan & IIf(tanpaGaris = 0, "[v] Semua paragraf sudah bergaris.", "[!] " & tanpaGaris & " paragraf belum diberi garis akta.")
    MsgBox pesan, IIf(adaNomor And nKosong = 0 And tanpaGaris = 0, vbInformation, vbExclamation), "Aktara - Periksa Akta"
End Sub

Public Sub SisipLuas()
    Dim s As String
    s = HanyaAngka(InputBox("Luas (M2):", "Aktara - Terbilang luas"))
    If Len(s) = 0 Then Exit Sub
    Sisip AktaraTerbilang.TerbilangLuas(CDbl(s))
End Sub

' Halaman akta sesuai praktik: A4, margin kiri 5,5 cm, kanan 1 cm, atas 1,8 cm, bawah 2,2 cm
Public Sub AturHalamanAkta()
    With ActiveDocument.PageSetup
        .PaperSize = 7 ' wdPaperA4
        .LeftMargin = CentimetersToPoints(5.5)
        .RightMargin = CentimetersToPoints(1)
        .TopMargin = CentimetersToPoints(1.8)
        .BottomMargin = CentimetersToPoints(2.2)
    End With
    ' Posisi tab garis mengikuti lebar halaman baru
    On Error Resume Next
    Dim st As Style: Set st = ActiveDocument.Styles(STYLE_GARIS)
    On Error GoTo 0
    If Not st Is Nothing Then PastikanStyle
    StatusBar = "Aktara: halaman akta diatur (A4, margin kiri 5,5 cm)."
End Sub

' Pilih angka di dokumen (mis. 1500000) lalu klik -> "1500000 (satu juta lima ratus ribu)"
Public Sub TerbilangkanPilihan()
    Dim s As String: s = HanyaAngka(Selection.Text)
    If Len(s) = 0 Then MsgBox "Pilih angka di dokumen terlebih dahulu.", vbInformation, "Aktara": Exit Sub
    If InStr(Selection.Text, ",") > 0 Then
        Selection.Text = Trim$(Selection.Text) & " (" & AktaraTerbilang.TerbilangDesimal(Trim$(Selection.Text)) & ")"
    Else
        Selection.Text = Trim$(Selection.Text) & " (" & AktaraTerbilang.Terbilang(CDbl(s)) & ")"
    End If
    Selection.Collapse 0 ' wdCollapseEnd
End Sub

' ================= Garis akta =================
Private Sub PastikanStyle()
    Dim st As Style
    On Error Resume Next
    Set st = ActiveDocument.Styles(STYLE_GARIS)
    On Error GoTo 0
    If st Is Nothing Then
        Set st = ActiveDocument.Styles.Add(STYLE_GARIS, 1) ' wdStyleTypeParagraph
        st.BaseStyle = ActiveDocument.Styles(-1).NameLocal ' wdStyleNormal
        st.Font.Name = "Courier New"
        st.Font.Size = 12
        st.ParagraphFormat.SpaceAfter = 0
        st.ParagraphFormat.LineSpacingRule = 4 ' wdLineSpaceExactly
        st.ParagraphFormat.LineSpacing = 25
        st.ParagraphFormat.Alignment = 3 ' rata kiri-kanan
    End If
    ' Tab kanan ber-leader strip tepat di batas kanan teks halaman ini
    Dim lebar As Single
    With ActiveDocument.PageSetup
        lebar = .PageWidth - .LeftMargin - .RightMargin
    End With
    st.ParagraphFormat.TabStops.ClearAll
    st.ParagraphFormat.TabStops.Add Position:=lebar, Alignment:=2, Leader:=2 ' wdAlignTabRight, wdTabLeaderDashes (2 = strip)
End Sub

Public Sub PasangGaris()
    PastikanStyle
    Dim p As Paragraph, n As Long, r As Range, rata As Long, tebal As Long
    For Each p In Selection.Paragraphs
        ' Lewati judul (rata tengah/kanan) dan paragraf kosong
        If p.Alignment = 1 Or p.Alignment = 2 Or Len(p.Range.Text) <= 1 Then GoTo Berikut
        ' Pertahankan perataan & huruf tebal
        rata = p.Alignment: tebal = p.Range.Font.Bold
        p.Style = ActiveDocument.Styles(STYLE_GARIS)
        p.Alignment = rata
        If tebal = -1 Then p.Range.Font.Bold = True
        Set r = p.Range
        r.MoveEnd 1, -1 ' tanpa tanda paragraf
        If Right$(r.Text, 1) <> vbTab Then r.InsertAfter vbTab: n = n + 1
Berikut:
    Next
    StatusBar = "Aktara: garis dipasang pada " & n & " paragraf."
End Sub

Public Sub HapusGaris()
    Dim p As Paragraph, n As Long, r As Range
    For Each p In Selection.Paragraphs
        Set r = p.Range
        r.MoveEnd 1, -1
        Do While Len(r.Text) > 0 And Right$(r.Text, 1) = vbTab
            r.Characters(r.Characters.Count).Delete
            Set r = p.Range: r.MoveEnd 1, -1
            n = n + 1
        Loop
    Next
    StatusBar = "Aktara: garis dihapus (" & n & ")."
End Sub

' ================= Template akta (dari AktaraTemplate, sumber templates/akta.json) =================
Public Sub BuatTemplate(ByVal kode As String)
    If Not AktaraServer.BolehPakai() Then Exit Sub
    Dim judul() As String, isi() As String, sifat As String
    If Not AktaraTemplate.AmbilTemplate(kode, judul, isi, sifat) Then MsgBox "Template tidak dikenal: " & kode, vbExclamation, "Aktara": Exit Sub
    Dim nomor As Long, sementara As Boolean
    nomor = AktaraServer.NomorAktaBerikutnya(sementara)

    Dim r As Range, awal As Long, i As Long, baris As String, p As Paragraph, tengah As Object
    Set tengah = CreateObject("Scripting.Dictionary")
    Set r = Selection.Range
    r.Collapse 1
    awal = r.Start
    For i = LBound(judul) To UBound(judul)
        r.InsertAfter judul(i) & vbCr: tengah.Add tengah.Count, "judul"
    Next
    r.InsertAfter "Nomor : " & AktaraTerbilang.FormatNomorAkta(nomor) & vbCr: tengah.Add tengah.Count, "nomor"
    For i = LBound(isi) To UBound(isi)
        baris = isi(i)
        If Left$(baris, 1) = "^" Then
            r.InsertAfter Mid$(baris, 2) & vbCr: tengah.Add tengah.Count, "tengah"
        Else
            r.InsertAfter baris & vbCr: tengah.Add tengah.Count, ""
        End If
    Next
    ' Format praktik akta: Courier New 12, spasi tepat 25 pt, rata kiri-kanan, judul di tengah
    i = 0
    For Each p In ActiveDocument.Range(awal, r.End).Paragraphs
        If i < tengah.Count Then
            p.Range.Font.Name = "Courier New": p.Range.Font.Size = 12
            p.LineSpacingRule = 4: p.LineSpacing = 25: p.SpaceAfter = 0
            If tengah(i) = "" Then p.Alignment = 3 Else p.Alignment = 1
            If tengah(i) = "judul" Then p.Range.Font.Bold = True
        End If
        i = i + 1
    Next
    BungkusIsian ActiveDocument.Range(awal, r.End)

    ActiveDocument.Variables("aktara.nomor").Value = CStr(nomor)
    ActiveDocument.Variables("aktara.sifat").Value = sifat
    If sementara Then
        MsgBox "Server tidak terjangkau. Nomor akta " & nomor & " adalah nomor SEMENTARA - cek ulang sebelum akta ditandatangani.", vbExclamation, "Aktara"
    Else
        StatusBar = "Aktara: akta " & sifat & " nomor " & nomor & " dibuat. Klik 'Isi Data Pihak'."
    End If
End Sub

' Ubah setiap {FIELD} menjadi content control ber-tag (Word 2007+)
Private Sub BungkusIsian(ByVal area As Range)
    ' area.End ikut bergeser otomatis saat content control ditambahkan
    Dim r As Range, tag As String, cc As ContentControl
    Set r = area.Duplicate
    With r.Find
        .ClearFormatting
        .Text = "\{[A-Z0-9_]@\}"
        .MatchWildcards = True
        .Forward = True
        .Wrap = 0 ' wdFindStop
    End With
    Do While r.Find.Execute
        If r.Start >= area.End Then Exit Do
        tag = Mid$(r.Text, 2, Len(r.Text) - 2)
        Set cc = ActiveDocument.ContentControls.Add(1, r) ' wdContentControlText
        cc.Tag = tag
        cc.Title = AktaraTemplate.LabelIsian(tag)
        r.SetRange cc.Range.End + 1, area.End
        If r.Start >= area.End Then Exit Do
        With r.Find
            .Text = "\{[A-Z0-9_]@\}"
            .MatchWildcards = True
            .Wrap = 0
        End With
    Loop
End Sub

' Tanya setiap isian unik sekali, lalu isi semua tempat dengan tag yang sama
Public Sub IsiDataPihak()
    Dim cc As ContentControl, sudah As Object, nilai As String, n As Long
    Set sudah = CreateObject("Scripting.Dictionary")
    If ActiveDocument.ContentControls.Count = 0 Then MsgBox "Belum ada isian. Buat akta dari template dulu.", vbInformation, "Aktara": Exit Sub
    For Each cc In ActiveDocument.ContentControls
        If Len(cc.Tag) > 0 And Not sudah.Exists(cc.Tag) Then
            nilai = InputBox(cc.Title & ":" & vbCr & "(kosongkan untuk melewati)", "Aktara - Data para pihak", AmbilIsi(cc))
            ' Nomor Induk Kependudukan (NIK) diperiksa; boleh dikoreksi atau tetap dipakai
            Do While Left$(cc.Tag, 4) = "NIK_" And Len(nilai) > 0
                If Len(AktaraTerbilang.CekNIK(nilai)) = 0 Then Exit Do
                If MsgBox(AktaraTerbilang.CekNIK(nilai) & vbCr & vbCr & "Perbaiki sekarang?", vbYesNo + vbExclamation, "Aktara - Periksa NIK") = vbNo Then Exit Do
                nilai = InputBox(cc.Title & ":", "Aktara - Data para pihak", nilai)
            Loop
            sudah.Add cc.Tag, nilai
        End If
    Next
    For Each cc In ActiveDocument.ContentControls
        If sudah.Exists(cc.Tag) Then
            If Len(sudah(cc.Tag)) > 0 Then cc.Range.Text = sudah(cc.Tag): n = n + 1
        End If
    Next
    StatusBar = "Aktara: " & n & " isian diperbarui."
End Sub

Private Function AmbilIsi(ByVal cc As ContentControl) As String
    If cc.ShowingPlaceholderText Then Exit Function
    If cc.Range.Text Like "{*}" Then Exit Function
    AmbilIsi = cc.Range.Text
End Function

' ================= Salinan =================
Public Sub JadikanSalinan()
    Dim kedudukan As String
    kedudukan = InputBox("Kedudukan notaris (mis. Kota Bandung):", "Aktara - Jadikan SALINAN")
    If StrPtr(kedudukan) = 0 Then Exit Sub
    ' Salinan dibuat sebagai dokumen baru, minuta tidak diubah
    Dim src As Document, dst As Document
    Set src = ActiveDocument
    Set dst = Documents.Add
    dst.Range.FormattedText = src.Range.FormattedText
    Dim r As Range
    Set r = dst.Range(0, 0)
    r.InsertBefore "SALINAN" & vbCr
    dst.Paragraphs(1).Alignment = 2 ' kanan
    dst.Paragraphs(1).Range.Font.Bold = True
    Set r = dst.Range
    r.InsertParagraphAfter
    r.InsertAfter "- Diberikan sebagai SALINAN yang sama bunyinya." & vbCr & _
        "Notaris di " & IIf(Len(kedudukan) > 0, kedudukan, "{KEDUDUKAN}") & vbCr & "(tanda tangan dan cap jabatan)"
    StatusBar = "Aktara: SALINAN dibuat di dokumen baru."
End Sub

' ================= Repertorium =================
Public Sub CatatRepertorium()
    Dim nomor As String, sifat As String, penghadap As String, cc As ContentControl
    On Error Resume Next
    nomor = ActiveDocument.Variables("aktara.nomor").Value
    sifat = ActiveDocument.Variables("aktara.sifat").Value
    On Error GoTo 0
    If Len(nomor) = 0 Then MsgBox "Dokumen ini belum dibuat dari template Aktara.", vbInformation, "Aktara": Exit Sub
    For Each cc In ActiveDocument.ContentControls
        If cc.Tag = "NAMA_PIHAK1" Or cc.Tag = "NAMA_PIHAK2" Then
            If Len(AmbilIsi(cc)) > 0 And InStr(penghadap, AmbilIsi(cc)) = 0 Then
                penghadap = penghadap & IIf(Len(penghadap) > 0, "; ", "") & AmbilIsi(cc)
            End If
        End If
    Next
    If Len(penghadap) = 0 Then penghadap = "-"
    Dim pesan As String
    If AktaraServer.CatatRepertorium(CLng(nomor), Date, sifat, penghadap, pesan) Then
        MsgBox "Akta nomor " & nomor & " tercatat di repertorium.", vbInformation, "Aktara"
    Else
        MsgBox "Gagal mencatat repertorium: " & pesan, vbExclamation, "Aktara"
    End If
End Sub
