VERSION 5.00
Begin VB.Form back 
   Caption         =   "Back up database"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4635
   Icon            =   "back.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Back up dengan gambar obat"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Triwulan Keempat"
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Triwulan Ketiga"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Triwulan Kedua"
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Triwulan pertama"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back &Up"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
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
Attribute VB_Name = "back"
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

Public Sub b2()
Dim cmd As String
FileName = Chr(34) & Dir1.Path & "\apotekbaleendah.sql" & Chr(34)
Screen.MousePointer = vbHourglass
    DoEvents
cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysqldump" & Chr(34) & " -h" & serperdatabes & " -u" & userdb & " -p" & passdb & " --routines --comments apotekbaleendah > " & FileName & """"
'cmd = "C:\Appserv\MySQL\bin\mysqldump -uroot -ptujuh7 --no-create-info --insert-ignore  penjualan > " & Filename & ""
    Call execCommand(cmd)
    If Check1.Value = Checked Then

     Set fso = New FileSystemObject
    fso.CopyFolder pss & "\gambar", Dir1.Path & "\gambar"
     
Set fso = Nothing
End If

    Screen.MousePointer = vbDefault
End Sub


Public Sub b1()
Screen.MousePointer = vbHourglass
FileName = "" + Dir1.Path + "\" + "Penjualan" + ".mdb"
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.CopyFile serper & "\Penjualan.mdb", FileName
    Screen.MousePointer = vbDefault

End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If databes = "Akses" Then
b1
Else
b2
End If
MsgBox "Back up database sukses !", vbInformation
End Sub

Private Sub Command3_Click()
Dim cmd, no1, no2 As String
Dim thn As Integer
thn = Format(Combo1.Text, "yy")
no1 = "PJ" & thn & "01001"
no2 = "PJ" & thn & "04001"

FileName = Chr(34) & Dir1.Path & "\penjualan.sql" & Chr(34)
    DoEvents

   ' cmd = Chr(34) & Chr(34) & pss & "\mysql\bin\mysqldump" & Chr(34) & " -h" & serperdatabes & " -uroot -ptujuh7 --no-create-info --insert-ignore  penjualan penjualan --where=no_penjualan between='bebas'> " & Filename & """"
'cmd = "G:\program files\toko\MySQL\bin\mysqldump -uroot -ptujuh7 --no-create-info --insert-ignore  penjualan tblbarang > " & Filename & ""
      'cmd = "G:\Program Files\Toko\mysql\bin\mysqldump -uroot -ptujuh7 --comments penjualan > " & Filename & ""
cmd = "H:\Appserv\MySQL\bin\mysqldump -uroot -ptujuh7 --no-create-info --insert-ignore  penjualan penjualan --where=no_penjualan >='" & no1 & "' and no_penjualan <='" & no1 & "'> " & FileName & ""

    Call execCommand(cmd)

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
On Error Resume Next
transaksi.cekaktip

Dim Arq As String
    Skinpath = App.Path & "\skin\winaqua.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
Ketengah Me

Dir1.Path = Arq
End Sub



