VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form restore 
   Caption         =   "Ambil database"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5880
   Icon            =   "restore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Ambil database kosong"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5280
      OleObjectBlob   =   "restore.frx":0CCA
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Restore"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "restore"
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
Public Sub b4()
LocTextFile = Dir1.Path & "\apotekbaleendah.sql"
 If Not Dir$(LocTextFile) = "" Then

Dim cmd As String
FileName = Chr(34) & Dir1.Path & "\apotekbaleendah.sql" & Chr(34)
Screen.MousePointer = vbHourglass
    DoEvents
cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysql" & Chr(34) & " -h" & serperdatabes & " -u" & userdb & " -p" & passdb & " apotekbaleendah < " & FileName & """"
'cmd = "C:\Appserv\MySQL\bin\mysql -hlocalhost -uroot -ptujuh7 penjualan3 < " & Filename & ""
    Call execCommand(cmd)
fso.DeleteFolder (App.Path & "\gambar")
     Set fso = New FileSystemObject
     
    fso.CopyFolder Dir1.Path & "\gambar", pss & "\gambar"
     
Set fso = Nothing

    Screen.MousePointer = vbDefault
    MsgBox "Berhasil,berhasil"
Else
MsgBox "Tidak ada database yang sesuai", vbInformation
End If
End Sub

Public Sub b3()
Screen.MousePointer = vbHourglass
LocTextFile = Dir1.Path & "\penjualan.mdb"
 If Not Dir$(LocTextFile) = "" Then

FileName = serper & "\penjualan" + ".mdb"
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile Dir1.Path + "\penjualan.mdb", FileName
MsgBox "Sukses,program akan menutup.", vbInformation
Screen.MousePointer = vbDefault
End
Else
MsgBox "Tidak ada database yang sesuai", vbInformation
End If
End Sub



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo erol
If MsgBox("Yakin akan mengambil database baru?", vbYesNo) = vbYes Then
If databes = "Akses" Then

b3
Else
b4
End If

End If
erol:
If err.Description <> vbNullString Then
MsgBox "Database tidak ditemukan", vbInformation
End If

End Sub


Private Sub Drive1_Change()
On Error GoTo err
Dir1.Path = Drive1.Drive
Exit Sub
err:
MsgBox "Drive kosong"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\database.doc" _
                 , vbNullString, vbNullString, 1
End If

End Sub

Private Sub Form_Load()
transaksi.cekaktip

Dim Arq As String
    Skinpath = App.Path & "\skin\winaqua.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
Ketengah Me

Dir1.Path = Arq
End Sub


