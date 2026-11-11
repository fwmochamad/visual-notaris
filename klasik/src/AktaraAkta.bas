Attribute VB_Name = "AktaraAkta"
' Aktara Edisi Klasik - fitur akta & tombol ribbon. Word 2007, 2010, 2013 (juga jalan di 2016+).
Option Explicit

Private Const STYLE_GARIS As String = "Akta Garis"

' ================= Ribbon callbacks =================
Public Sub AkOnLoad(ribbon As IRibbonUI)
End Sub

Public Sub AkKuasa(control As IRibbonControl): BuatTemplate "kuasa": End Sub
Public Sub AkPPJB(control As IRibbonControl): BuatTemplate "jualbeli": End Sub
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

' Pilih angka di dokumen (mis. 1500000) lalu klik -> "1500000 (satu juta lima ratus ribu)"
Public Sub TerbilangkanPilihan()
    Dim s As String: s = HanyaAngka(Selection.Text)
    If Len(s) = 0 Then MsgBox "Pilih angka di dokumen terlebih dahulu.", vbInformation, "Aktara": Exit Sub
    Selection.Text = Trim$(Selection.Text) & " (" & AktaraTerbilang.Terbilang(CDbl(s)) & ")"
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
        st.ParagraphFormat.LineSpacingRule = 1 ' wdLineSpace1pt5
        st.ParagraphFormat.Alignment = 0 ' kiri
    End If
    ' Tab kanan ber-leader strip tepat di batas kanan teks halaman ini
    Dim lebar As Single
    With ActiveDocument.PageSetup
        lebar = .PageWidth - .LeftMargin - .RightMargin
    End With
    st.ParagraphFormat.TabStops.ClearAll
    st.ParagraphFormat.TabStops.Add Position:=lebar, Alignment:=2, Leader:=5 ' wdAlignTabRight, wdTabLeaderDashes
End Sub

Public Sub PasangGaris()
    PastikanStyle
    Dim p As Paragraph, n As Long, r As Range
    For Each p In Selection.Paragraphs
        p.Style = ActiveDocument.Styles(STYLE_GARIS)
        Set r = p.Range
        r.MoveEnd 1, -1 ' tanpa tanda paragraf
        If Right$(r.Text, 1) <> vbTab Then r.InsertAfter vbTab: n = n + 1
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

' ================= Template akta =================
Private Function IsiTemplate(ByVal kode As String, ByRef judul As String) As Variant
    Dim pembuka As String
    pembuka = "Berhadapan dengan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya, Notaris, kenal dan akan disebut pada bagian akhir akta ini:"
    If kode = "kuasa" Then
        judul = "KUASA"
        IsiTemplate = Array(pembuka, _
            "Tuan/Nyonya {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Nomor Induk Kependudukan {NIK_PIHAK1};", _
            "- selanjutnya disebut Pemberi Kuasa.", _
            "Penghadap menerangkan dengan ini memberi kuasa kepada Tuan/Nyonya {NAMA_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Nomor Induk Kependudukan {NIK_PIHAK2};", _
            "- selanjutnya disebut Penerima Kuasa.", _
            "KHUSUS untuk {OBJEK_KUASA}.")
    Else
        judul = "PERJANJIAN PENGIKATAN JUAL BELI"
        IsiTemplate = Array(pembuka, _
            "I. Tuan/Nyonya {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Nomor Induk Kependudukan {NIK_PIHAK1};", _
            "- selanjutnya disebut Pihak Pertama.", _
            "II. Tuan/Nyonya {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Nomor Induk Kependudukan {NIK_PIHAK2};", _
            "- selanjutnya disebut Pihak Kedua.", _
            "Para penghadap menerangkan bahwa Pihak Pertama mengikatkan diri untuk menjual kepada Pihak Kedua {OBJEK} dengan harga {HARGA}.")
    End If
End Function

Public Sub BuatTemplate(ByVal kode As String)
    If Not AktaraServer.BolehPakai() Then Exit Sub
    Dim judul As String, isi As Variant, nomor As Long, sementara As Boolean
    isi = IsiTemplate(kode, judul)
    nomor = AktaraServer.NomorAktaBerikutnya(sementara)

    Dim r As Range, awal As Long, i As Long
    Set r = Selection.Range
    r.Collapse 1 ' wdCollapseStart
    awal = r.Start
    r.InsertAfter judul & vbCr & "Nomor : " & AktaraTerbilang.TerbilangAngka(nomor) & vbCr & _
        "Pada hari ini, {HARI}, tanggal {TANGGAL}, pukul {JAM}." & vbCr
    For i = LBound(isi) To UBound(isi)
        r.InsertAfter isi(i) & vbCr
    Next
    ActiveDocument.Range(awal, awal).Paragraphs(1).Alignment = 1 ' tengah
    ActiveDocument.Range(awal, awal).Paragraphs(1).Range.Font.Bold = True
    ActiveDocument.Range(awal, awal).Paragraphs(1).Next.Alignment = 1
    BungkusIsian ActiveDocument.Range(awal, r.End)

    ActiveDocument.Variables("aktara.nomor").Value = CStr(nomor)
    ActiveDocument.Variables("aktara.sifat").Value = StrConv(judul, vbProperCase)
    If sementara Then
        MsgBox "Server tidak terjangkau. Nomor akta " & nomor & " adalah nomor SEMENTARA - cek ulang sebelum akta ditandatangani.", vbExclamation, "Aktara"
    Else
        StatusBar = "Aktara: akta " & judul & " nomor " & nomor & " dibuat. Klik 'Isi Data Pihak'."
    End If
End Sub

' Ubah setiap {FIELD} menjadi content control ber-tag (Word 2007+)
Private Sub BungkusIsian(ByVal area As Range)
    Dim r As Range, tag As String, cc As ContentControl, batas As Long
    batas = area.End
    Set r = area.Duplicate
    With r.Find
        .ClearFormatting
        .Text = "\{[A-Z0-9_]@\}"
        .MatchWildcards = True
        .Forward = True
        .Wrap = 0 ' wdFindStop
    End With
    Do While r.Find.Execute
        If r.Start >= batas Then Exit Do
        tag = Mid$(r.Text, 2, Len(r.Text) - 2)
        Set cc = ActiveDocument.ContentControls.Add(1, r) ' wdContentControlText
        cc.Tag = tag
        cc.Title = Replace(tag, "_", " ")
        r.SetRange cc.Range.End + 1, batas
        If r.Start >= batas Then Exit Do
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
