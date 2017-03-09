VERSION 5.00
Begin VB.Form satuann 
   Caption         =   "Tambah satuan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "="
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Satuan"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama barang"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Kode barang"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "satuann"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
If Text1.Text = "" Then
MsgBox "Isi nama satuan"
Text1.SetFocus
Else
If val(Text2.Text) = 0 Then
MsgBox "Konversi jangan nol"
Text2.SetFocus
Else
jual.Execute "insert into satuan values('" & Label5.Caption & "','" & Text1.Text & "','" & Text2.Text & "','')"
MsgBox "Data berhasil disimpan"
Me.Hide
Barang.Visible = True
End If
End If
erol:
If err.Description <> vbNullString Then
    MsgBox "Satuan sudah ada", vbCritical, "Penjualan"
End If

End Sub

Private Sub Form_Load()
Ketengah Me
End Sub

Private Sub Text1_Change()
Label8.Caption = Text1.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan_Click
End If
End Sub
