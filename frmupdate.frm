VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Update"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmupdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Update Database"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update File"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub Command1_Click()
If MsgBox("Update files?", vbYesNo) = vbNo Then Exit Sub
If fso.FolderExists(App.Path & "\report2") = False Then
MsgBox "Salah mengkopy update an", vbCritical
Exit Sub
End If
sumber = App.Path & "\files\*.rpt"
tujuan = App.Path & "\report2"
Set fso = New Scripting.FileSystemObject
fso.CopyFile sumber, tujuan, True
sumber = App.Path & "\files\*.exe"
tujuan = App.Path
Set fso = New Scripting.FileSystemObject
fso.CopyFile sumber, tujuan, True
MsgBox "Sudah", vbInformation

End Sub

Private Sub Command2_Click()
On Error GoTo erol
If MsgBox("Update database?", vbYesNo) = vbNo Then Exit Sub
konekdb

konek.Execute "alter table tservis ADD brgservis varchar(255)"
Exit Sub
erol:
If Err.Number = -2147217900 Then
MsgBox "Sudah update", vbCritical, judul
Else
MsgBox "Gagal konek database", vbCritical, judul
End If
End Sub

