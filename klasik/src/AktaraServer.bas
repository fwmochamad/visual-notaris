Attribute VB_Name = "AktaraServer"
' Aktara Edisi Klasik - koneksi ke Aktara API (sama dengan Edisi Modern).
' Data lokal disimpan di registry: HKCU\Software\VB and VBA Program Settings\Aktara
Option Explicit

Private Const APP As String = "Aktara"
Private Const SERVER_PRODUKSI As String = "https://aktara-api-production.up.railway.app"
Private Const SERVER_DEV As String = "https://localhost:8080"

Private Function Ambil(ByVal kunci As String, Optional ByVal bawaan As String = "") As String
    Ambil = GetSetting(APP, "Lisensi", kunci, bawaan)
End Function
Private Sub Simpan(ByVal kunci As String, ByVal nilai As String)
    SaveSetting APP, "Lisensi", kunci, nilai
End Sub

Public Function UrlServer() As String
    UrlServer = Ambil("server", SERVER_PRODUKSI)
End Function

Public Function DeviceId() As String
    DeviceId = Ambil("device")
    If Len(DeviceId) = 0 Then
        Randomize
        DeviceId = "klasik-" & Environ$("COMPUTERNAME") & "-" & Hex$(CLng(Rnd * 2147483647#))
        Simpan "device", DeviceId
    End If
End Function

' ---------- JSON minimal (hanya untuk respons Aktara API yang sederhana) ----------
Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r"): s = Replace(s, vbLf, "\n"): s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function

Public Function JsonAmbil(ByVal json As String, ByVal kunci As String) As String
    Dim p As Long, q As Long, c As String
    p = InStr(json, """" & kunci & """:")
    If p = 0 Then Exit Function
    p = p + Len(kunci) + 3
    If Mid$(json, p, 1) = """" Then
        q = p + 1
        Do While q <= Len(json)
            c = Mid$(json, q, 1)
            If c = "\" Then
                q = q + 2
            ElseIf c = """" Then
                Exit Do
            Else
                q = q + 1
            End If
        Loop
        JsonAmbil = Replace(Replace(Mid$(json, p + 1, q - p - 1), "\""", """"), "\\", "\")
    Else
        q = p
        Do While q <= Len(json) And InStr(",}]", Mid$(json, q, 1)) = 0
            q = q + 1
        Loop
        JsonAmbil = Trim$(Mid$(json, p, q - p))
    End If
End Function

' ---------- HTTP ----------
' Mengembalikan status HTTP (0 = tidak terhubung). Teks respons di argumen respons.
Private Function Panggil(ByVal metode As String, ByVal path As String, ByVal body As String, ByRef respons As String) As Long
    Dim x As Object
    On Error GoTo Gagal
    Set x = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    x.setTimeouts 5000, 5000, 8000, 8000
    x.Open metode, UrlServer() & path, False
    x.setRequestHeader "Content-Type", "application/json"
    If Len(Ambil("token")) > 0 Then x.setRequestHeader "Authorization", "Bearer " & Ambil("token")
    If Len(body) > 0 Then x.send body Else x.send
    respons = x.responseText
    Panggil = x.Status
    Exit Function
Gagal:
    respons = "{""error"":""Tidak dapat terhubung ke server (" & JsonEsc(Err.Description) & ")""}"
    Panggil = 0
End Function

' ---------- Lisensi ----------
Public Function StatusLisensi(ByRef sisaHari As Long) As String
    If Len(Ambil("token")) = 0 Then StatusLisensi = "belum": Exit Function
    Dim dicek As Date, t As String
    t = Ambil("dicek", Format$(Date, "yyyy-mm-dd"))
    dicek = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 6, 2)), CInt(Mid$(t, 9, 2)))
    sisaHari = CLng(Ambil("batas", "30")) - DateDiff("d", dicek, Date)
    If sisaHari > 0 Then StatusLisensi = "aktif" Else StatusLisensi = "kedaluwarsa"
End Function

Public Function BolehPakai() As Boolean
    Dim sisa As Long, st As String, resp As String, kode As Long
    st = StatusLisensi(sisa)
    If st = "belum" Then
        MsgBox "Aktara belum diaktifkan. Klik tombol Lisensi pada tab Aktara.", vbInformation, APP
        Exit Function
    End If
    ' Perbarui cek lisensi bila online (diam-diam)
    kode = Panggil("GET", "/api/lisensi", "", resp)
    If kode = 200 Then
        Simpan "dicek", Format$(Date, "yyyy-mm-dd"): BolehPakai = True
    ElseIf kode = 401 Or kode = 402 Then
        Simpan "token", ""
        MsgBox JsonAmbil(resp, "error") & vbCr & "Aktifkan ulang lewat tombol Lisensi.", vbExclamation, APP
    Else
        BolehPakai = (st = "aktif")
        If Not BolehPakai Then MsgBox "Lisensi perlu dicek ulang. Sambungkan internet.", vbExclamation, APP
    End If
End Function

Public Sub DialogLisensi()
    Dim sisa As Long, st As String, info As String
    st = StatusLisensi(sisa)
    info = "Server: " & UrlServer() & vbCr & "Perangkat: " & DeviceId() & vbCr & vbCr
    If st <> "belum" Then
        If MsgBox(info & "Status: aktif untuk " & Ambil("kantor") & vbCr & "Berlaku offline " & sisa & " hari lagi." & vbCr & vbCr & _
                  "Ganti lisensi / server?", vbYesNo + vbQuestion, APP) = vbNo Then Exit Sub
    End If
    Dim url As String, kodeLis As String, nama As String, resp As String, status As Long
    url = InputBox("Alamat server (biarkan bawaan kecuali kantor memakai server lokal;" & vbCr & "untuk uji coba: " & SERVER_DEV & "):", APP, UrlServer())
    If Len(url) = 0 Then Exit Sub
    Simpan "server", url
    kodeLis = InputBox("Kode lisensi (AKT-XXXX-XXXX-XXXX-XXXX):", APP)
    If Len(kodeLis) = 0 Then Exit Sub
    nama = InputBox("Nama komputer ini (mis. PC Resepsionis):", APP, Environ$("COMPUTERNAME"))
    Simpan "token", ""
    status = Panggil("POST", "/api/aktivasi", "{""kode_lisensi"":""" & JsonEsc(Trim$(kodeLis)) & """,""device_id"":""" & JsonEsc(DeviceId()) & _
        """,""nama_perangkat"":""" & JsonEsc(nama) & """}", resp)
    If status = 200 Then
        Simpan "token", JsonAmbil(resp, "token")
        Simpan "kantor", JsonAmbil(resp, "kantor")
        Simpan "batas", JsonAmbil(resp, "batas_offline_hari")
        Simpan "dicek", Format$(Date, "yyyy-mm-dd")
        MsgBox "Aktivasi berhasil untuk " & JsonAmbil(resp, "kantor") & ".", vbInformation, APP
    Else
        MsgBox "Aktivasi gagal: " & JsonAmbil(resp, "error"), vbExclamation, APP
    End If
End Sub

' ---------- Nomor akta ----------
Public Function NomorAktaBerikutnya(ByRef sementara As Boolean) As Long
    Dim resp As String, kunci As String
    kunci = "nomor." & Year(Date) & "-" & Month(Date)
    If Panggil("POST", "/api/nomor-akta", "{""tahun"":" & Year(Date) & ",""bulan"":" & Month(Date) & "}", resp) = 200 Then
        NomorAktaBerikutnya = CLng(JsonAmbil(resp, "nomor"))
        sementara = False
    Else
        NomorAktaBerikutnya = CLng(Ambil(kunci, "0")) + 1
        sementara = True
    End If
    Simpan kunci, CStr(NomorAktaBerikutnya)
End Function

' ---------- Repertorium ----------
Public Function CatatRepertorium(ByVal nomor As Long, ByVal tgl As Date, ByVal sifat As String, ByVal penghadap As String, ByRef pesan As String) As Boolean
    Dim resp As String
    If Panggil("POST", "/api/repertorium", "{""tahun"":" & Year(tgl) & ",""bulan"":" & Month(tgl) & ",""nomor"":" & nomor & _
        ",""tanggal"":""" & Format$(tgl, "yyyy-mm-dd") & """,""sifat_akta"":""" & JsonEsc(sifat) & """,""penghadap"":""" & JsonEsc(penghadap) & """}", resp) = 200 Then
        CatatRepertorium = True
    Else
        pesan = JsonAmbil(resp, "error")
    End If
End Function
