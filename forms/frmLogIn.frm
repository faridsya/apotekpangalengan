VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Begin VB.Form frmxLogIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ClipControls    =   0   'False
   Icon            =   "frmLogIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtbaru 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtupdate 
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3960
      OleObjectBlob   =   "frmLogIn.frx":08CA
      Top             =   2880
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Atur IP"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin XPControls.XPText Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
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
      PasswordChar    =   "*"
   End
   Begin XPControls.XPText text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
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
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   6615
      TabIndex        =   9
      Top             =   960
      Width           =   6615
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   240
      ScaleHeight     =   30
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin XPControls.XPText serv 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label3 
      Caption         =   "IP Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Login !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "frmLogIn.frx":0AFE
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Username and Password to login."
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmxLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Private Const SYNCHRONIZE       As Long = &H100000
    Private Const INFINITE          As Long = &HFFFF
                Dim fso As New FileSystemObject

Dim I As Integer
Dim A, b, c
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
Private Sub Command1_Click()
On Error GoTo erol
my
installodbc
Exit Sub
erol:
If err.Number = -2147467259 Then
 MsgBox ("Server/database belum terhubung,cek settingan databasenya!")
 

End If

End Sub
Private Sub Command2_Click()
mengecil Me
End
End Sub
Sub my()
If posisi = "Server" Then
serperdatabes = "localhost"
pss = App.Path

serperreport = App.Path & "\report2"

Else
pss = "\\" & serper & "\poss"

serperdatabes = serper
serperreport = "\\" & serper & "\poss\report2"

End If

    If ConnectDb2(Text2.Text, Text1.Text) Then
        'MsgBox "Connect Kedatabase " & UCase(Text3.Text) & " Sukses!!", vbInformation, "Sukses"
      Unload Me
        Mnutama.Show
          
     With Mnutama
    .StatusBar1.Panels(1).Text = nama_toko

    .StatusBar1.Panels(2).Text = rshusus!UserName
     DataUserID = rshusus!UserName
    .StatusBar1.Panels(4).Text = GetFileSize(StripPath(serper) & "" + dbName + " ")
    .brg.Enabled = IIf(rshusus.Fields(4) = "0", False, True)
    .langgan.Enabled = IIf(rshusus.Fields(5) = "0", False, True)
    .mnmutasi.Enabled = IIf(rshusus.Fields(6) = "0", False, True)
    .ps.Enabled = IIf(rshusus.Fields(7) = "0", False, True)
    .supp.Enabled = IIf(rshusus.Fields(8) = "0", False, True)
    .tb.Enabled = IIf(rshusus.Fields(9) = "0", False, True)
    .jb.Enabled = IIf(rshusus.Fields(10) = "0", False, True)
    .rb.Enabled = IIf(rshusus.Fields(11) = "0", False, True)
    .reju.Enabled = IIf(rshusus.Fields(12) = "0", False, True)
    .keu.Enabled = IIf(rshusus.Fields(13) = "0", False, True)
    .lprn.Enabled = IIf(rshusus.Fields(14) = "0", False, True)
    .guna.Enabled = IIf(rshusus.Fields(15) = "0", False, True)
    .bup.Enabled = IIf(rshusus.Fields(16) = "0", False, True)
    .ad.Enabled = IIf(rshusus.Fields(17) = "0", False, True)
    .del.Enabled = IIf(rshusus.Fields(18) = "0", False, True)
    .mnmservis.Enabled = IIf(rshusus.Fields("cek16") = "0", False, True)
    .mnteknisi.Enabled = IIf(rshusus.Fields("cek17") = "0", False, True)
        .mnlapak.Enabled = IIf(rshusus.Fields("cek20") = "0", False, True)

    .Toolbar1.Buttons(1).Enabled = IIf(rshusus.Fields(4) = "0", False, True)
    .Toolbar1.Buttons(2).Enabled = IIf(rshusus.Fields(8) = "0", False, True)
    .Toolbar1.Buttons(3).Enabled = IIf(rshusus.Fields(10) = "0", False, True)
    .Toolbar1.Buttons(4).Enabled = IIf(rshusus.Fields(16) = "0", False, True)
    .Toolbar1.Buttons(5).Enabled = IIf(rshusus.Fields(15) = "0", False, True)

    .Toolbar1.Buttons(6).Enabled = IIf(rshusus.Fields(5) = "0", False, True)
    .Toolbar1.Buttons(7).Enabled = IIf(rshusus.Fields(6) = "0", False, True)
    .Toolbar1.Buttons(8).Enabled = IIf(rshusus.Fields(9) = "0", False, True)



    '.Toolbar1.Buttons(5).Enabled = IIf(rshusus.Fields(15) = "0", False, True)
    .StatusBar1.Panels(8).Text = rshusus!nama
        
        End With
        rhj = IIf(rshusus.Fields(19) = "0", False, True)



    Else
    If I <= 1 Then
 If MsgBox(" Percobaan ke-" & I + 1 & " salah", vbOKOnly, "Login") = vbOK Then

    If I = 1 Then
     If MsgBox(" Ini percobaan terakhirmu", vbOKOnly, "Login") = vbOK Then
     
     End If
     End If
 I = I + 1
 Text2.SetFocus
 Exit Sub
 End If
 Else
 If MsgBox("Sudah cukup layau", vbOKOnly, "Login") = vbOK Then
 
 Unload Me
 End If
 End If
    End If

End Sub
Sub bikin()
Dim DriverODBC As String
Dim NameDSN As String
DriverODBC = String(255, Chr(32))
NameDSN = "penjualan"
    'Have SQL drivers been installed?


    

    'Does the DSN name already exist?
    If (MySQLDSNWanted(NameDSN)) = True Then
        
    Else
        If Not MakeMySQLDSN(DriverODBC, NameDSN) Then
        End If
    End If
    

End Sub

Private Sub Command3_Click()
If Command3.Caption = "&Atur IP" Then
Command3.Caption = "&Confirm"
serv.Visible = True
 Label3.Visible = True
 Else
 Command3.Caption = "&Atur IP"
 serv_KeyPress (13)
 End If
End Sub

Private Sub Command4_Click()
prosesupdate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
frmdb.Show
'Command3.Visible = True
End If
End Sub

Private Sub Form_Load()
'Text1.Text = "admin"
'Text2.Text = "admin"
'Command1.TabIndex = 0
txtbaru.Text = "171214"
idpc = GetSetting("apotekbaleendah", "frmxLogIn", "text3.text", "")
aktipasi = GetSetting("apotekbaleendah", "frmid", "Text3.text", False)
text3.Text = LCase(idpc)
If idpc = "" Then
djie_get_motherboard
idpc = LCase(idproc)
'idpc = Right(idpc, Len(idpc) - 8)
'idpc = idpc + RandomString(6)
text3.Text = LCase(idpc)
SaveSetting "apotekbaleendah", "frmxLogIn", "text3.text", text3.Text

End If
serper = GetSetting("apotekbaleendah", "frmdb", "txtip.text", "localhost")
userdb = GetSetting("apotekbaleendah", "frmdb", "txtuser.text", "root")
passdb = GetSetting("apotekbaleendah", "frmdb", "txtpass.text", "tujuh7")
portdb = GetSetting("apotekbaleendah", "frmdb", "txtport.text", "tujuh7")
versiupdate = GetSetting("apotekbaleendah", "frmxLogIn", "txtupdate.text", "0")
'If versiupdate <> txtbaru.Text Then
'prosesupdate
'txtupdate.Text = txtbaru.Text
'SaveSetting "apotekbaleendah", "frmdb", "frmxLogIn", txtupdate.Text
'End If

fldr = App.Path & "\mysql"
mysqlfolder = GetSetting("apotekbaleendah", "frmdb", "txtfolder.text", fldr)

serv.Text = serper
databes = "Mysql"

If serper = "localhost" Then
posisi = "Server"
Else
posisi = "Client"
End If


End Sub
Sub prosesupdate()
LocTextFile = App.Path & "\update.sql"
 If Not Dir$(LocTextFile) = "" Then

Dim cmd As String
FileName = Chr(34) & App.Path & "\update.sql" & Chr(34)
Screen.MousePointer = vbHourglass
    DoEvents
cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysql" & Chr(34) & " -h" & serperdatabes & " -u" & userdb & " -p" & passdb & " apotekbaleendah < " & FileName & """"
'cmd = "C:\Appserv\MySQL\bin\mysql -hlocalhost -uroot -ptujuh7 penjualan3 < " & Filename & ""
    Call execCommand(cmd)

    Screen.MousePointer = vbDefault
    MsgBox "Selesai!"
Else
MsgBox "Update GAGAL!", vbInformation
End If

End Sub
Private Sub serv_GotFocus()
serv.ToolTipText = "Ketik ip atau nama komputer server lalu tekan enter"
End Sub
Private Sub serv_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

SaveSetting "apotekbaleendah", "frmxLogIn", "serv.text", serv.Text
serper = GetSetting("apotekbaleendah", "frmdb", "txtip.text", "localhost")
serv.Visible = False
Label3.Visible = False
Command3.Visible = False

End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
    A = Mid(text4, 1, 1)
    b = Right(text4, Len(text4) - 1)
    c = b & A
    text4 = c
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
End Sub

