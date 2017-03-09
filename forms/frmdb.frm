VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmdb 
   Caption         =   "Pengaturan database"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPText txtport 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmdb.frx":0000
      TabIndex        =   12
      Top             =   1320
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmdb.frx":0078
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmdb.frx":00EE
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmdb.frx":016E
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmdb.frx":01EE
      TabIndex        =   8
      Top             =   240
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5040
      OleObjectBlob   =   "frmdb.frx":025E
      Top             =   4080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Buat database"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin XPControls.XPText txtfolder 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   5280
      Width           =   2175
   End
   Begin XPControls.XPText txtpass 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPText txtuser 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPText txtip 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
        Private Const SYNCHRONIZE       As Long = &H100000
    Private Const INFINITE          As Long = &HFFFF


Private Sub Command1_Click()
SaveSetting "apotekbaleendah", "frmdb", "txtip.text", txtip.Text
SaveSetting "apotekbaleendah", "frmdb", "txtuser.text", txtuser.Text
SaveSetting "apotekbaleendah", "frmdb", "txtpass.text", txtpass.Text
SaveSetting "apotekbaleendah", "frmdb", "txtfolder.text", txtfolder.Text
SaveSetting "apotekbaleendah", "frmdb", "txtport.text", txtport.Text

serper = GetSetting("apotekbaleendah", "frmdb", "txtip.text", "localhost")
userdb = GetSetting("apotekbaleendah", "frmdb", "txtuser.text", "root")
passdb = GetSetting("apotekbaleendah", "frmdb", "txtpass.text", "tujuh7")
portdb = GetSetting("apotekbaleendah", "frmdb", "txtport.text", "3306")

fldr = App.Path & "\mysql"
mysqlfolder = GetSetting("apotekbaleendah", "frmdb", "txtfolder.text", fldr)

If serper = "localhost" Then
posisi = "Server"
Else
posisi = "Client"
End If

If DSNDelete("apotekbaleendah", "MySQL ODBC 5.1 Driver") Then
End If

MsgBox "Pengaturan berhasil", vbInformation, judul
End Sub

Private Sub Command2_Click()
LocTextFile = App.Path & "\penjualan.sql"
 If Not Dir$(LocTextFile) = "" Then
Dim cmd As String

FileName = Chr(34) & App.Path & "\penjualan.sql" & Chr(34)
Screen.MousePointer = vbHourglass
DoEvents
If MsgBox("Yakin buat database baru?", vbYesNo, judul) = vbNo Then Exit Sub
cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysql" & Chr(34) & " -h" & serper & " -u" & userdb & " -p" & passdb & " -ecreate database penjualan """
'cmd = "C:\Appserv\MySQL\bin\mysql -hlocalhost -uroot -ptujuh7 penjualan3 < " & Filename & ""
    Call execCommand(cmd)

cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysql" & Chr(34) & " -h" & serper & " -u" & userdb & " -p" & passdb & "  < " & FileName & """"
'cmd = "C:\Appserv\MySQL\bin\mysql -hlocalhost -uroot -ptujuh7 penjualan3 < " & Filename & ""
    Call execCommand(cmd)

    Screen.MousePointer = vbDefault
    MsgBox "Berhasil,berhasil"
Else
MsgBox "Tidak ada database yang sesuai", vbInformation
End If

End Sub

Private Sub Command3_Click()
On Error GoTo erol
If posisi = "Client" Then
MsgBox "Hanya untuk komputer server", vbInformation, judul
Exit Sub
End If
If MsgBox("Yakin buat database baru?", vbYesNo, judul) = vbNo Then Exit Sub
If Dir$(mysqlfolder & "\" & namadb, vbDirectory) = "" Then
MkDir mysqlfolder & "\" & namadb
End If

Dim fso
    Dim sfol As String, dfol As String
    sfol = App.Path & "\database & " \ " & namadb ' change to match the source folder path"
    dfol = mysqlfolder & "\data & " \ " & namadb ' change to match the destination folder path"
    Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFolder sfol, dfol
        MsgBox "Database baru berrhasil dibuat", vbInformation, judul
        Exit Sub
erol:
        If err.Description <> vbNullString Then
    MsgBox "Settingan masih salah", vbCritical, judul
End If

End Sub

Private Sub CopyFolder()
    
End Sub
Private Sub Dir1_Change()
txtfolder.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
Exit Sub
err:
MsgBox "Drive kosong"
End Sub

Private Sub Form_Load()
serper = GetSetting("apotekbaleendah", "frmdb", "txtip.text", "localhost")
userdb = GetSetting("apotekbaleendah", "frmdb", "txtuser.text", "root")
passdb = GetSetting("apotekbaleendah", "frmdb", "txtpass.text", "tujuh7")
portdb = GetSetting("apotekbaleendah", "frmdb", "txtport.text", "3306")

fldr = App.Path & "\mysql"
mysqlfolder = GetSetting("apotekbaleendah", "frmdb", "txtfolder.text", fldr)
txtfolder.Text = mysqlfolder
txtip.Text = serper
txtuser.Text = userdb
txtpass.Text = passdb
txtport.Text = portdb
    Skinpath = App.Path & "\skin\galaxy.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub execCommand(ByVal cmd As String)
        Dim Result  As Long
        Dim lPid    As Long
        Dim lHnd    As Long
        Dim lRet    As Long

        cmd = "cmd /c " & cmd
        Result = Shell(cmd, vbHide)

        lPid = Result
        If lPid <> 0 Then
            lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
            If lHnd <> 0 Then
                lRet = WaitForSingleObject(lHnd, INFINITE)
                CloseHandle (lHnd)
            End If
        End If
    End Sub
Sub konek()
    Set jual = New adodb.Connection
       jual.CursorLocation = adUseClient
jual.Open "DRIVER={MySQL ODBC 5.1 Driver};" _
                & "SERVER=" & serperdatabes & "" _
                & ";DATABASE=penjualan;" _
                & ";USER=" & userdb & "" _
                & ";PORT=3306;" _
                & ";PASSWORD=" & passdb & "" _
                & ";OPTION=3;"

End Sub

