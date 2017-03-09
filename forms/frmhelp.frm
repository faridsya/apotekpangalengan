VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmhelp 
   Caption         =   "Bantuan"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "frmhelp.frx":0000
      TabIndex        =   8
      Top             =   5040
      Width           =   4695
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4800
      OleObjectBlob   =   "frmhelp.frx":015A
      Top             =   2760
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Transaksi penjualan"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Transaksi pembelian"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Bayar piutang retail"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Isi data akun tambahan"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Isi Neraca awal"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Isi Master Pelanggan"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Isi Master Barang"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Umum"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\umum.doc" _
                 , vbNullString, vbNullString, 1

End Sub

Private Sub Command10_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pembelian.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command11_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\byrpiutang.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command2_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\barang.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command3_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pelanggan.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command4_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\neracaawal.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command5_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\akun.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Command9_Click()
ShellExecute Me.hwnd, "open", App.Path & "\panduan\penjualan.doc" _
                 , vbNullString, vbNullString, 1
End Sub

Private Sub Form_Load()
Skinpath = App.Path & "\skin\golden.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub
