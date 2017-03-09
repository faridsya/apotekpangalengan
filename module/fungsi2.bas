Attribute VB_Name = "fungsi2"
Public DataUserID As String
Public mu As String
Public namatoko As String
Public tgjwb As String
Public almttoko As String
Public telptoko As String
Public subnama As String
Public idproc As String
Public idpc As String
Public aktipasi As Boolean
Public matu As String
Public idb As String
Public gno As String
Public gtgl As Date
Public gnom As String
Public pss As String
Public hbrg As String
Public hpju As String
Public hpb As String
Public hcus As String
Public hsup As String
Public hpo As String
Public isipromosi As String
Public posisi As String
Public databes As String
Public rhj As Boolean
Public isibeli As Boolean
Public adalogo As Boolean
Public cttn1 As String
Public cttn2 As String
Public cttn3 As String
Public cttn4 As String
Public cttn5 As String
Public cttn6 As String
Public kodesip As Integer
Public sip As Boolean
Public Aktivitas As String
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public status_isi As String
Public pilih As String
Public jumbay As String
Public jumbar As String
Public pembayaran As Currency
Public jugir As Currency
Public junai As Currency

Public gsFileName As String
Public gsDrive As String
Public gsPath As String
Public gErrFormName As String
Private Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" _
(ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function FindWindow Lib "user32" Alias _
   "FindWindowA" (ByVal lpClassName As String, _
   ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias _
  "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
   ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
   
Public Declare Function SetTimer& Lib "user32" _
  (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal _
   lpTimerFunc&)

Private Declare Function KillTimer& Lib "user32" _
  (ByVal hwnd&, ByVal nIDEvent&)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Const EM_SETPASSWORDCHAR = &HCC
Public Const NV_INPUTBOX As Long = &H5000&
Public Sub djie_get_motherboard()
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor", , 48)
    For Each objItem In colItems

        idproc = objItem.ProcessorId
        Next
End Sub

Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, _
   ByVal idEvent&, ByVal dwTime&)

    Dim EditHwnd As Long

' CHANGE APP.TITLE TO YOUR INPUT BOX TITLE.

    EditHwnd = FindWindowEx(FindWindow("#32770", App.Title), _
       0, "Edit", "")

    Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
    KillTimer hwnd, idEvent
End Sub

Public Function TerbilangDesimal(InputCurrency As String, Optional MataUang As String = "rupiah") As String
 Dim strInput As String
 Dim strBilangan As String
 Dim strPecahan As String
   On Error GoTo Pesan
   Dim strValid As String, huruf As String * 1
   Dim I As Integer
   'Periksa setiap karakter yg diketikkan ke kotak
   'UserID
   strValid = "1234567890,"
   For I% = 1 To Len(InputCurrency)
     huruf = Chr(Asc(Mid(InputCurrency, I%, 1)))
     If InStr(strValid, huruf) = 0 Then
       Set AngkaTerbilang = Nothing
       MsgBox "Harus karakter angka!", _
              vbCritical, "Karakter Tidak Valid"
       Exit Function
     End If
   Next I%
 
 If InputCurrency = "" Then Exit Function
 If Len(Trim(InputCurrency)) > 15 Then GoTo Pesan
 
 strInput = CStr(InputCurrency) 'Konversi ke string
 'Periksa apakah ada tanda "," jika ya berarti pecahan
 If InStr(1, strInput, ",", vbBinaryCompare) Then
      
  strBilangan = Left(strInput, InStr(1, strInput, _
                ",", vbBinaryCompare) - 1)
  'strBilangan = Right(strInput, InStr(1, strInput, _
  '              ".", vbBinaryCompare) - 2)
  strPecahan = Trim(Right(strInput, Len(strInput) - Len(strBilangan) - 1))
  
  If MataUang <> "" Then
      
  If CLng(Trim(strPecahan)) > 99 Then
     strInput = Format(Round(CDbl(strInput), 2), "#0.00")
     strPecahan = Format((Right(strInput, Len(strInput) - Len(strBilangan) - 1)), "00")
    End If
    
    If Len(Trim(strPecahan)) = 1 Then
       strInput = Format(Round(CDbl(strInput), 2), _
                  "#0.00")
       strPecahan = Format((Right(strInput, _
          Len(strInput) - Len(strBilangan) - 1)), "00")
    End If
    
    If CLng(Trim(strPecahan)) = 0 Then
    TerbilangDesimal = (KonversiBilangan(strBilangan) & MataUang & " " & KonversiBilangan(strPecahan))
 Else
  TerbilangDesimal = (KonversiBilangan(strBilangan) & MataUang & " " & KonversiBilangan(strPecahan) & "sen")
    End If
  Else
    TerbilangDesimal = (KonversiBilangan(strBilangan) & "koma " & KonversiPecahan(strPecahan))
  End If
  
 Else
    TerbilangDesimal = (KonversiBilangan(strInput))
  End If
 Exit Function
Pesan:
  TerbilangDesimal = "(maksimal 15 digit)"
End Function
Private Function KonversiPecahan(strAngka As String) As String
Dim I%, strJmlHuruf$, Urai$, Kar$
 If strAngka = "" Then Exit Function
    strJmlHuruf = Trim(strAngka)
    Urai = ""
    Kar = ""
    For I = 1 To Len(strJmlHuruf)
      'Tampung setiap satu karakter ke Kar
      Kar = Mid(strAngka, I, 1)
      Urai = Urai & kata(CInt(Kar))
    Next I
    KonversiPecahan = Urai
End Function

'Fungsi ini untuk menterjemahkan setiap satu angka ke 'kata
Private Function kata(angka As Byte) As String
   Select Case angka
          Case 1: kata = "satu "
          Case 2: kata = "dua "
          Case 3: kata = "tiga "
          Case 4: kata = "empat "
          Case 5: kata = "lima "
          Case 6: kata = "enam "
          Case 7: kata = "tujuh "
          Case 8: kata = "delapan "
          Case 9: kata = "sembilan "
          Case 0: kata = "nol "
   End Select
End Function

'Ini untuk mengkonversi nilai bilangan sebelum pecahan
Private Function KonversiBilangan(strAngka As String) As String
Dim strJmlHuruf$, intPecahan As Integer, strPecahan$, Urai$, Bil1$, strTot$, Bil2$
 Dim x, y, z As Integer

 If strAngka = "" Then Exit Function
    strJmlHuruf = Trim(strAngka)
    x = 0
    y = 0
    Urai = ""
    While (x < Len(strJmlHuruf))
      x = x + 1
      strTot = Mid(strJmlHuruf, x, 1)
      y = y + val(strTot)
      z = Len(strJmlHuruf) - x + 1
      Select Case val(strTot)
      'Case 0
       '   Bil1 = "NOL "
      Case 1
          If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
              Bil1 = "satu "
          ElseIf (z = 4) Then
              If (x = 1) Then
                  Bil1 = "se"
              Else
                  Bil1 = "satu "
              End If
          ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
              x = x + 1
              strTot = Mid(strJmlHuruf, x, 1)
              z = Len(strJmlHuruf) - x + 1
              Bil2 = ""
              Select Case val(strTot)
              Case 0
                  Bil1 = "sepuluh "
              Case 1
                  Bil1 = "sebelas "
              Case 2
                  Bil1 = "dua belas "
              Case 3
                  Bil1 = "tiga belas "
              Case 4
                  Bil1 = "empat belas "
              Case 5
                  Bil1 = "lima belas "
              Case 6
                  Bil1 = "enam belas "
              Case 7
                  Bil1 = "tujuh belas "
              Case 8
                  Bil1 = "delapan belas "
              Case 9
                  Bil1 = "sembilan belas "
              End Select
          Else
              Bil1 = "se"
          End If
      
      Case 2
          Bil1 = "dua "
      Case 3
          Bil1 = "tiga "
      Case 4
          Bil1 = "empat "
      Case 5
          Bil1 = "lima "
      Case 6
          Bil1 = "enam "
      Case 7
          Bil1 = "tujuh "
      Case 8
          Bil1 = "delapan "
      Case 9
          Bil1 = "sembilan "
      Case Else
          Bil1 = ""
      End Select
       
      If (val(strTot) > 0) Then
         If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
            Bil2 = "puluh "
         ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
            Bil2 = "ratus "
         Else
            Bil2 = ""
         End If
      Else
         Bil2 = ""
      End If
      If (y > 0) Then
          Select Case z
          Case 4
              Bil2 = Bil2 + "ribu "
              y = 0
          Case 7
              Bil2 = Bil2 + "juta "
              y = 0
          Case 10
              Bil2 = Bil2 + "milyar "
              y = 0
          Case 13
              Bil2 = Bil2 + "trilyun "
              y = 0
          End Select
      End If
      Urai = Urai + Bil1 + Bil2
  Wend
  KonversiBilangan = Urai
End Function

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
Function buat_folder2()
On Error Resume Next
Dim x As Byte
Dim y As SECURITY_ATTRIBUTES
Dim z As String
bl = Format(Now, "M")
Th = Format(Now, "YYYY")
tg = Format(Now, "d")
z = "D:\backup database"
y.lpSecurityDescriptor = 0
y.bInheritHandle = True
y.nLength = Len(y)
x = CreateDirectory(z, y)



End Function
Public Function bekap()
On Error Resume Next
FileName = "D:\backup database\" + "Penjualan" + ".mdb"
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile serper & "\Penjualan.mdb", FileName
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
Dim I As Integer
Dim Temp As String
Dim Lokasi As Integer
Dim huruf As String * 1
  Temp$ = ""
  For I% = 1 To Len(strKalimat)
    huruf = Chr(Asc(Mid(strKalimat, I%, 1)))
    If Len(Trim(huruf)) < 1 Then
      Lokasi% = I% + 1
    End If
    If I% = Lokasi% Or I% = 1 Then
       Temp$ = Temp$ + UCase(Chr(Asc(Mid(strKalimat, _
               I%, 1))))
    Else
       Temp$ = Temp$ + LCase(Chr(Asc(Mid(strKalimat, _
                I%, 1))))
    End If
  Next I
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
   sNamaFile = serper & "\loglogin.txt"
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

   'SetAttr serper & "\Loglogin.log", vbHidden + vbSystem

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
  
  sNamaFile = serper & "\Rekaman kegiatan\Tahun " & Th & "\Bulan " & bl & "\Tanggal " & tg & "\Loglogin.tx"
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

Public Sub Ketengah(ByVal frm As Form)
On Error Resume Next
frm.Left = Screen.Width / 2 - frm.Width / 2
frm.Top = Screen.Height / 2 - frm.Height / 2
  End Sub
Sub Translate()  'Encrypt/Decrypt Password
Dim I As Integer
Dim Lokasi As Integer
code = "1234567890" 'Ini kode/kunci utk melakukan encrypt/decrypt
  Temp$ = ""
  For I% = 1 To Len(DataString)
      Lokasi% = (I% Mod Len(code)) + 1
      'Gunakan logika XOR utk kombinasi encrypt/decrypt
      Temp$ = Temp$ + Chr$(Asc(Mid$(DataString, I%, 1)) Xor _
      Asc(Mid$(code, Lokasi%, 1)))
  Next I%
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
Public Function nama_toko()
nama_toko = namatoko
End Function
Public Function nama_toko2()
nama_toko2 = subnama
End Function
Public Function almt()
almt = almttoko
End Function
Public Function almt2()
almt2 = telptoko
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
Function validasiAngka3(KeyAscii As Integer) As Integer
    Dim strValid As String
   
    strValid = "0123456789"
    If InStr(1, strValid, Chr$(KeyAscii)) = 0 And Not (KeyAscii = vbKeyBack) And Not (KeyAscii = 13) And Not (KeyAscii = 44) And Not (KeyAscii = 46) Then
        validasiAngka3 = 0
        MsgBox "Harus Memilih", vbCritical, "Peringatan"
    Else
        validasiAngka3 = KeyAscii
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
Dim x As Byte
Dim y As SECURITY_ATTRIBUTES
Dim z As String
bl = Format(Now, "M")
Th = Format(Now, "YYYY")
tg = Format(Now, "d")
z = serper & "\Rekaman Kegiatan"
y.lpSecurityDescriptor = 0
y.bInheritHandle = True
y.nLength = Len(y)
x = CreateDirectory(z, y)
z = serper & "\Rekaman Kegiatan\Tahun " & Th & ""
y.lpSecurityDescriptor = 0
y.bInheritHandle = True
y.nLength = Len(y)
x = CreateDirectory(z, y)
z = serper & "\Rekaman Kegiatan\Tahun " & Th & "\Bulan " & bl & ""
y.lpSecurityDescriptor = 0
y.bInheritHandle = True
y.nLength = Len(y)
x = CreateDirectory(z, y)
z = serper & "\Rekaman Kegiatan\Tahun " & Th & "\Bulan " & bl & "\Tanggal " & tg & ""
y.lpSecurityDescriptor = 0
y.bInheritHandle = True
y.nLength = Len(y)
x = CreateDirectory(z, y)


If x = 1 Then
Else
Exit Function
End If

End Function

Sub demo()
Dim x
Dim y
Dim jumlah
Dim sisa

x = GetSetting("iii", "iia", "iia")
jumlah = val(x) + 1
SaveSetting "iii", "iia", "iia", jumlah
sisa = 31 - jumlah
If sisa = 30 Then
MsgBox "Program ini hanya dapat di gunakan 30 kali"
End If
If sisa > 0 Then
MsgBox "Sisa pemakaian " & sisa & " Kali"
Else

MsgBox "Batas waktu pemakaian sudah habis" + vbCrLf + _
"silahkan beli full version dengan" + vbCrLf + _
"sms/menghubungi saya di 02291500183 ....", vbOKOnly, "Info"
Unload Mnutama
End If
End Sub
Public Function txtGotFocus()
Dim obj
Set obj = Form1.ActiveControl
    If TypeOf obj Is TextBox Then
        obj.SelStart = 0
        obj.SelLength = Len(obj.Text)
    End If
    
    
    

End Function


