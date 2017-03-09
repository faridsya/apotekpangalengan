VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Info..."
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "beli"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "jual"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   120
      ScaleHeight     =   30
      ScaleWidth      =   4575
      TabIndex        =   2
      Top             =   3480
      Width           =   4575
   End
   Begin VB.PictureBox ctrlLiner2 
      BackColor       =   &H00C0C0C0&
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   1440
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Kontak : 082116969006"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "HG SOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   3372
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1470
      Left            =   -360
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Set RS = New Recordset
RS.Open "select * from keuangan where keterangan='Transaksi penjualan' order by no_transaksi", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub

RS.MoveFirst
For I = 1 To RS.RecordCount
Set rsd = New Recordset
rsd.Open "select count(no_transaksi) as jum from keuangan where no_transaksi='" & RS!no_transaksi & "'", jual, adOpenStatic, adLockOptimistic
If rsd!jum > 1 Then
jual.Execute "delete from keuangan where no_transaksi='" & RS!no_transaksi & "' and keterangan='Transaksi penjualan' order by pemasukan asc limit " & rsd!jum - 1 & ""
End If
RS.MoveNext
Next
MsgBox "berhasil"

End Sub

Private Sub Command3_Click()
Set RS = New Recordset
RS.Open "select * from keuangan where keterangan='Transaksi pembelian' order by no_transaksi", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
RS.MoveFirst
For I = 1 To RS.RecordCount
Set rsd = New Recordset
rsd.Open "select count(no_transaksi) as jum from keuangan where no_transaksi='" & RS!no_transaksi & "'", jual, adOpenStatic, adLockOptimistic
If rsd!jum > 1 Then
jual.Execute "delete from keuangan where no_transaksi='" & RS!no_transaksi & "' and keterangan='Transaksi pembelian' order by pengeluaran asc limit " & rsd!jum - 1 & ""
End If
RS.MoveNext
Next
MsgBox "berhasil"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
Command2.Visible = True
Command3.Visible = True
End If
End Sub

Private Sub Form_Load()
Command2.Visible = False
Command3.Visible = False

Label5.Caption = "HG SOFT"
End Sub

