VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form hps 
   Caption         =   "Kosongkan data Tabel"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "hps.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "hps.frx":324A
      TabIndex        =   11
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Kosongkan"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Kosongkan"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Kosongkan"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kosongkan"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kosongkan"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kosongkan"
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Kosongkan data data gudang"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Kosongkan data data obat"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Kosongkan data-data pelanggan"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Kosongkan data-data supplier"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   $"hps.frx":32E6
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "hps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from penjualan"
jual.Execute "Delete from pembelian"
jual.Execute "Delete from detilbeli"
jual.Execute "Delete from pemesanan"
jual.Execute "Delete from detilpesan"

jual.Execute "Delete from penjualan"
jual.Execute "Delete from detiljual"
jual.Execute "Delete from penjualans"
jual.Execute "Delete from detiljuals"

jual.Execute "Delete from piutang"
jual.Execute "Delete from hutang"
jual.Execute "Delete from giro"
jual.Execute "Delete from retur_beli"
jual.Execute "Delete from detilreturbeli"
jual.Execute "Delete from detilreturjual"

jual.Execute "Delete from retur_jual"
jual.Execute "Delete from byr_hutang"
jual.Execute "Delete from byr_piutang"
jual.Execute "Delete from sesuai"
jual.Execute "Delete from keuangan"

  jual.Execute "update tblsupplier set jumlah_hutang = 0  "
  jual.Execute "update pelanggan set jumlah_piutang = 0  "

MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command2_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from tblsupplier"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command3_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from pelanggan"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command4_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from tblbarang"
jual.Execute "Delete from satuan"
jual.Execute "Delete from paket"
jual.Execute "Delete from mutasigudang"


MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command5_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from gudang where kode_gudang!='utama'"


MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command6_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from servis"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command7_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub

jual.Execute "Delete from teknisi"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command8_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub
jual.Execute "Delete from tservis"
jual.Execute "Delete from tservis_dtl2"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Command9_Click()
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub
jual.Execute "Delete from bayar_pajak"
MsgBox "Berhasil dihapus", vbInformation

End Sub

Private Sub Form_Load()
Ketengah Me
'transaksi.cekaktip
End Sub
