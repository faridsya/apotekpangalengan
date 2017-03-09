Attribute VB_Name = "fungsi2"
Public DataUserID As String
Public Aktivitas As String
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nsize As Long, ByVal lpFileName As String) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public status_isi As String

Public gsFileName As String
Public gsDrive As String
Public gsPath As String
Public gErrFormName As String
Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" _
(ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Sub emti(currentForm As Form)
    Dim ctl As Control
    
    For Each ctl In currentForm.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is XPText Then ctl.Text = ""
    Next
End Sub
Public Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Public Function IsPrinterInstalled() As Boolean
On Error Resume Next
Dim strDummy As String
strDummy = Printer.DeviceName
  If err.Number Then
     IsPrinterInstalled = False
  Else
     IsPrinterInstalled = True
  End If
End Function
Sub HandleError(strLoc As String, strError$, lError As Long, varModule As Variant)

    Dim nCursorType As Integer

    nCursorType = Screen.MousePointer

    Screen.MousePointer = vbNormal
    MsgBox strLoc & ": " & strError & " (" & lError & ")", vbExclamation, varModule
    Screen.MousePointer = nCursorType

End Sub

Public Function judul()
judul = "Point of sales"
End Function

Public Function AwalKataKapital(strKalimat As String)
Dim i As Integer
Dim Temp As String
Dim Lokasi As Integer
Dim huruf As String * 1
  Temp$ = ""
  For i% = 1 To Len(strKalimat)
    huruf = Chr(Asc(Mid(strKalimat, i%, 1)))
    If Len(Trim(huruf)) < 1 Then
      Lokasi% = i% + 1
    End If
    If i% = Lokasi% Or i% = 1 Then
       Temp$ = Temp$ + UCase(Chr(Asc(Mid(strKalimat, _
               i%, 1))))
    Else
       Temp$ = Temp$ + LCase(Chr(Asc(Mid(strKalimat, _
                i%, 1))))
    End If
  Next i
  AwalKataKapital = Temp$
End Function

Public Sub RekamKegiatan(Aktivitas As String)
Dim sNamaFile As String
Dim NoFile As Integer
Dim JamSekarang As String
Dim TglHariIni As String
  'Supaya seragam, gunakan format tanggal yang seragam
    TglHariIni = Format(Now, "dd MMMM yyyy")
    JamSekarang = Format(Now, "hh:mm:ss")

bl = Format(Now, "M")
Th = Format(Now, "YYYY")
tg = Format(Now, "d")
  'Inisialisasi nama file log

  sNamaFile = App.Path & "\Loglogin.log"
  'Ambil penanganan sembarang file berikutnya dari
  'OS Windows, agar tidak bentrok dengan FreeFile.
   NoFile = FreeFile
      NoFile = FreeFile

  'Simpan ke log dengan mode Append (tambahkan yg
  'sudah ada). Kita menggunakan karakter koma untuk
  'memisahkan antara item (field) yg satu dgn yg lain.
   Open sNamaFile For Append As #NoFile
     'Rekam ke file log sekarang.....!
     Print #NoFile, DataUserID & "," & Aktivitas & "," & TglHariIni & "," & JamSekarang
   Close #NoFile  'Tutup file log jika sudah disimpan

   SetAttr App.Path & "\Loglogin.log", vbHidden + vbSystem

End Sub
Public Sub RekamKegiatan2(Aktivitas As String)
Dim sNamaFile As String
Dim NoFile As Integer
Dim JamSekarang As String
Dim TglHariIni As String
  'Supaya seragam, gunakan format tanggal yang seragam
    TglHariIni = Format(Now, "dd MMMM yyyy")
    JamSekarang = Format(Now, "hh:mm:ss")
    buat_folder
bl = Format(Now, "M")
Th = Format(Now, "YYYY")
tg = Format(Now, "d")
  'Inisialisasi nama file log
  
  sNamaFile = App.Path & "\Rekaman kegiatan\Tahun " & Th & "\Bulan " & bl & "\Tanggal " & tg & "\Loglogin.tx"
  'Ambil penanganan sembarang file berikutnya dari
  'OS Windows, agar tidak bentrok dengan FreeFile.
   NoFile = FreeFile
  'Simpan ke log dengan mode Append (tambahkan yg
  'sudah ada). Kita menggunakan karakter koma untuk
  'memisahkan antara item (field) yg satu dgn yg lain.
   Open sNamaFile For Append As #NoFile
     'Rekam ke file log sekarang.....!
     Print #NoFile, DataUserID & "," & Aktivitas & "," & TglHariIni & "," & JamSekarang
   Close #NoFile  'Tutup file log jika sudah disimpan
End Sub

Public Sub Ketengah(ByVal Frm As Form)
On Error Resume Next
Frm.Left = Screen.Width / 2 - Frm.Width / 2
Frm.Top = Screen.Height / 2 - Frm.Height / 2
  End Sub
Sub Translate()  'Encrypt/Decrypt Password
Dim i As Integer
Dim Lokasi As Integer
code = "1234567890" 'Ini kode/kunci utk melakukan encrypt/decrypt
  Temp$ = ""
  For i% = 1 To Len(DataString)
      Lokasi% = (i% Mod Len(code)) + 1
      'Gunakan logika XOR utk kombinasi encrypt/decrypt
      Temp$ = Temp$ + Chr$(Asc(Mid$(DataString, i%, 1)) Xor _
      Asc(Mid$(code, Lokasi%, 1)))
  Next i%
End Sub

Function validasiAngka(KeyAscii As Integer) As Integer
    Dim strValid As String
   
    strValid = "0123456789"
    If InStr(1, strValid, Chr$(KeyAscii)) = 0 And Not (KeyAscii = vbKeyBack) And Not (KeyAscii = 13) Then
        validasiAngka = 0
        MsgBox "Harus angka", vbCritical, "Peringatan"
    Else
        validasiAngka = KeyAscii
    End If
End Function
Function validasiAngka2(KeyAscii As Integer) As Integer
    Dim strValid As String
   
    strValid = "0123456789"
    If InStr(1, strValid, Chr$(KeyAscii)) = 0 And Not (KeyAscii = vbKeyBack) And Not (KeyAscii = 13) And Not (KeyAscii = 44) And Not (KeyAscii = 46) Then
        validasiAngka2 = 0
        MsgBox "Format salah", vbCritical, "Peringatan"
    Else
        validasiAngka2 = KeyAscii
    End If
End Function

  Public Function GetFileSize(file As Variant) As String
    On Error Resume Next
    Dim Bytes As Long
    Const Kb As Long = 1024
    Const Mb As Long = 1024 * Kb
    Const Gb As Long = 1024 * Mb
    Bytes = FileLen(file)
    If Bytes < Kb Then
        GetFileSize = Format(Bytes) & " bytes"
    ElseIf Bytes < Mb Then
        GetFileSize = Format(Bytes / Kb, "0.00") & " Kb"
    ElseIf Bytes < Gb Then
        GetFileSize = Format(Bytes / Mb, "0.00") & " Mb"
    Else
        GetFileSize = Format(Bytes / Gb, "0.00") & " Gb"
    End If
End Function

Function StripPath(nPath As String) As String
If Right(nPath, 1) = "\" Then
   StripPath = nPath
Else
   StripPath = nPath & "\"
End If
End Function

Public Function ReadINI(strFile As String, strKey As String, strName As String) As String
Dim intLen As Integer
Dim strText As String
'strText = Space(255)
strText = "                                                                                                    "
intLen = GetPrivateProfileString(strKey, strName, "", strText, Len(strText), strFile)
If intLen > -1 Then
    strText = Left(strText, intLen)
Else
End
End If
ReadINI = strText
End Function

Function buat_folder()
Dim X As Byte
Dim Y As SECURITY_ATTRIBUTES
Dim z As String
bl = Format(Now, "M")
Th = Format(Now, "YYYY")
tg = Format(Now, "d")
z = App.Path & "\Rekaman Kegiatan"
Y.lpSecurityDescriptor = 0
Y.bInheritHandle = True
Y.nLength = Len(Y)
X = CreateDirectory(z, Y)
z = App.Path & "\Rekaman Kegiatan\Tahun " & Th & ""
Y.lpSecurityDescriptor = 0
Y.bInheritHandle = True
Y.nLength = Len(Y)
X = CreateDirectory(z, Y)
z = App.Path & "\Rekaman Kegiatan\Tahun " & Th & "\Bulan " & bl & ""
Y.lpSecurityDescriptor = 0
Y.bInheritHandle = True
Y.nLength = Len(Y)
X = CreateDirectory(z, Y)
z = App.Path & "\Rekaman Kegiatan\Tahun " & Th & "\Bulan " & bl & "\Tanggal " & tg & ""
Y.lpSecurityDescriptor = 0
Y.bInheritHandle = True
Y.nLength = Len(Y)
X = CreateDirectory(z, Y)


If X = 1 Then
Else
Exit Function
End If

End Function

Sub demo()
Dim X
Dim Y
Dim jumlah
Dim sisa

X = GetSetting("Y", "Y", "Y")
jumlah = val(X) + 1
SaveSetting "Y", "Y", "Y", jumlah
sisa = 31 - jumlah
If sisa = 30 Then
MsgBox "Program ini hanya dapat di gunakan 30 kali"
End If
If sisa > 0 Then
MsgBox "Sisa pemakaian " & sisa & " Kali"
Else

MsgBox "Batas waktu pemakaian sudah habis" + vbCrLf + _
"silahkan beli full version dengan" + vbCrLf + _
"sms/menghubungi saya di 0857 2216 9724 ....", vbOKOnly, "Info"
Unload Mnutama
End If
End Sub



