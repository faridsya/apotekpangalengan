VERSION 5.00
Begin VB.Form caribrg 
   Caption         =   "Cari Data"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Kode barang"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cari"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Nama barang"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "caribrg.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Kata Kunci"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label B 
      Caption         =   "Berdasarkan :"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "caribrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" Then
Text1.Text = Replace(Text1.Text, "'", "''")

If Option1.Value = True Then
sql = "select * from tblbarang where kode_brg = '" & Text1.Text & "' "
Set rsbarang = jual.Execute(sql)
Set Barang.dbgrid1.DataSource = rsbarang
Barang.teks
 Unload Me
Else
If Option2.Value = True Then
sql = "select * from tblbarang where deskripsi like '%" & Text1.Text & "%' "
 Set rsbarang = jual.Execute(sql)
Set Barang.dbgrid1.DataSource = rsbarang
Barang.teks
Unload Me
 Barang.Cmdkeluar.SetFocus


Else
MsgBox "Belum dipilih", vbCritical, "Penjualan"
End If
End If

Barang.dbgrid1.Refresh
Else
MsgBox "Isi  donk textboxnya!!", , "Peringatan"
Text1.SetFocus

End If
End Sub

Private Sub Command1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Image1_Click()
Command1_Click
End Sub

Private Sub Option1_Click()
Text1.SetFocus
End Sub

Private Sub Option2_Click()
Text1.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
