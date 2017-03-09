VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmrepair 
   Caption         =   "Perbaiki table"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7290
   Icon            =   "frmrepair.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Perbaiki table"
      Height          =   615
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Lakukan bila transaksi mendadak ada yg erorr setelah sekian lama berlangsung baik,tidak akan menghilangkan data"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6975
   End
End
Attribute VB_Name = "frmrepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If MsgBox("Perbaiki tabel-tabel?", vbYesNo, judul) = vbNo Then Exit Sub
Set RS = New Recordset
RS.Open "show tables from apotekbaleendah", jual, adOpenStatic, adLockOptimistic
ProgressBar1.Visible = True

ProgressBar1.max = RS.RecordCount

RS.MoveFirst
ps = 0
ProgressBar1.Value = ps
Do While Not RS.EOF
ps = ps + 1

jual.Execute "repair table " & RS.Fields(0) & ""
ProgressBar1.Value = ps
RS.MoveNext
Loop

MsgBox "Sudah", vbInformation, judul
ProgressBar1.Visible = False
End Sub
