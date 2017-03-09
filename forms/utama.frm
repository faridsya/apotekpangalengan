VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm Mnutama 
   BackColor       =   &H80000004&
   Caption         =   "Sistem informasi Apotek"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6015
   Icon            =   "utama.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "utama.frx":0E42
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "utama.frx":AD2C0
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   8000
      Top             =   2160
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":AD4F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1200
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":ADD00
            Key             =   "pelanggan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B04B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B0B35
            Key             =   "jual"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B35B7
            Key             =   "back"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B6039
            Key             =   "user"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B6E8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":B7697
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":BC4A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "utama.frx":BC98A
            Key             =   "out"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3525
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3263
            MinWidth        =   3263
            Text            =   "JAYA AGUNG PLASTIK"
            TextSave        =   "JAYA AGUNG PLASTIK"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "utama.frx":BCDDC
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1331
            MinWidth        =   1331
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2648
            MinWidth        =   2648
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "utama.frx":BD176
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1984
            MinWidth        =   1984
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   3240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5715
      ButtonWidth     =   2566
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Produk"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Supplier"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Penjualan"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Back Up"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&User"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pe&langgan"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Mutasi gudang"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pembelian"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Keluar"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu brg 
         Caption         =   "&Obat"
      End
      Begin VB.Menu langgan 
         Caption         =   "&Pelanggan"
      End
      Begin VB.Menu mndokter 
         Caption         =   "&Dokter"
      End
      Begin VB.Menu mnapoteker 
         Caption         =   "&Apoteker"
      End
      Begin VB.Menu mnakun 
         Caption         =   "Akun tambahan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnmpaket 
         Caption         =   "Racikan"
      End
      Begin VB.Menu bang 
         Caption         =   "Ban&k"
         Visible         =   0   'False
      End
      Begin VB.Menu mcpu 
         Caption         =   "Cetak Label harga"
      End
      Begin VB.Menu mnbar 
         Caption         =   "Cetak barcode"
      End
      Begin VB.Menu ps 
         Caption         =   "Penyesuaian sto&k"
      End
      Begin VB.Menu mnmservis 
         Caption         =   "Master Servis"
         Visible         =   0   'False
      End
      Begin VB.Menu mnsls 
         Caption         =   "Sales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnteknisi 
         Caption         =   "Teknisi"
         Visible         =   0   'False
      End
      Begin VB.Menu supp 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mngudang 
         Caption         =   "Gudang"
      End
      Begin VB.Menu dato 
         Caption         =   "Identitas Apotek"
      End
      Begin VB.Menu frmnrcawal 
         Caption         =   "Neraca awal"
         Visible         =   0   'False
      End
      Begin VB.Menu mnpromosi 
         Caption         =   "Setting promosi"
      End
      Begin VB.Menu mnjadwal 
         Caption         =   "Jadwal shift"
      End
   End
   Begin VB.Menu trans 
      Caption         =   "&Transaksi"
      Begin VB.Menu mnmutasi 
         Caption         =   "Mutasi Gudang       "
         Shortcut        =   {F1}
      End
      Begin VB.Menu tb 
         Caption         =   "Pem&belian barang      "
         Shortcut        =   {F2}
      End
      Begin VB.Menu jb 
         Caption         =   "Pen&jualan barang"
         Shortcut        =   {F3}
      End
      Begin VB.Menu rb 
         Caption         =   "Ret&ur beli                        "
         Shortcut        =   {F4}
      End
      Begin VB.Menu reju 
         Caption         =   "Re&tur jual                   "
         Shortcut        =   {F5}
      End
      Begin VB.Menu mndo 
         Caption         =   "Delivery Order   "
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnpesan 
         Caption         =   "Purchase Order"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mntrx 
         Caption         =   "Transaksi tambahan      "
         Visible         =   0   'False
      End
   End
   Begin VB.Menu keu 
      Caption         =   "&Keuangan"
      Begin VB.Menu bapi 
         Caption         =   "Pembayaran &piutang"
      End
      Begin VB.Menu bahut 
         Caption         =   "&Pembayaran hutang"
      End
      Begin VB.Menu pmsk 
         Caption         =   "Pe&masukan"
      End
      Begin VB.Menu kluar 
         Caption         =   "Pe&ngeluaran"
      End
      Begin VB.Menu mnbyrpjk 
         Caption         =   "&Pembayaran Pajak"
      End
   End
   Begin VB.Menu lprn 
      Caption         =   "&Laporan"
      Begin VB.Menu mnlapak 
         Caption         =   "Akunting"
         Visible         =   0   'False
         WindowList      =   -1  'True
         Begin VB.Menu mnlapjurnal 
            Caption         =   "Jurnal"
         End
         Begin VB.Menu mnlaprl 
            Caption         =   "Laba rugi"
         End
         Begin VB.Menu mnlapneraca 
            Caption         =   "Neraca"
         End
         Begin VB.Menu mnlapbb 
            Caption         =   "Buku Besar"
         End
      End
      Begin VB.Menu mnlapbrg 
         Caption         =   "Barang"
         Begin VB.Menu mnlembar 
            Caption         =   "Lembar Stokopname"
         End
         Begin VB.Menu db 
            Caption         =   "Data &obat"
         End
         Begin VB.Menu stk 
            Caption         =   "S&tok obat"
         End
         Begin VB.Menu lps 
            Caption         =   "Pen&yesuain stok"
         End
         Begin VB.Menu mnhistori 
            Caption         =   "Histori harga"
         End
         Begin VB.Menu mnlapgud 
            Caption         =   "Stok Gudang"
         End
         Begin VB.Menu mts 
            Caption         =   "&Mutasi Gudang"
         End
         Begin VB.Menu mnlapjualbeli 
            Caption         =   "Keluar masuk barang"
         End
         Begin VB.Menu mnoutin 
            Caption         =   "Kartu Stok"
         End
         Begin VB.Menu mnlapex 
            Caption         =   "&Expired"
         End
      End
      Begin VB.Menu mnlapkeu 
         Caption         =   "Keuangan"
         Begin VB.Menu mnlapinfo 
            Caption         =   "Info per faktur"
         End
         Begin VB.Menu j 
            Caption         =   "Pen&jualan"
         End
         Begin VB.Menu b 
            Caption         =   "Pem&belian"
         End
         Begin VB.Menu lapkom 
            Caption         =   "Komisi dokter"
         End
         Begin VB.Menu rtj 
            Caption         =   "Retur Jual"
         End
         Begin VB.Menu retbel 
            Caption         =   "&Retur Beli"
         End
         Begin VB.Menu pht 
            Caption         =   "Piutang"
         End
         Begin VB.Menu htg 
            Caption         =   "Hutang"
         End
         Begin VB.Menu pmbp 
            Caption         =   "Pembayaran Piutang"
         End
         Begin VB.Menu pmbh 
            Caption         =   "Pembayaran Hutang"
         End
         Begin VB.Menu mnlapshift 
            Caption         =   "Shift"
         End
         Begin VB.Menu tran 
            Caption         =   "Arus kas"
         End
         Begin VB.Menu mnlapbyrpjk 
            Caption         =   "Pembayaran pajak"
         End
      End
      Begin VB.Menu ds 
         Caption         =   "Data &Supplier"
      End
      Begin VB.Menu plg 
         Caption         =   "Data Pelan&ggan"
      End
      Begin VB.Menu lagi 
         Caption         =   "La&ba rugi"
      End
      Begin VB.Menu mksr 
         Caption         =   "Ka&sir"
         Visible         =   0   'False
      End
      Begin VB.Menu kmb 
         Caption         =   "Keluar ma&suk barang"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu out 
      Caption         =   "&Utilities"
      Begin VB.Menu ebot 
         Caption         =   "A&bout"
      End
      Begin VB.Menu mnhelp 
         Caption         =   "Bantuan"
         Visible         =   0   'False
      End
      Begin VB.Menu kalk 
         Caption         =   "&Kalkulator"
      End
      Begin VB.Menu guna 
         Caption         =   "&Pengguna"
      End
      Begin VB.Menu gpass 
         Caption         =   "&Ganti password"
      End
      Begin VB.Menu bup 
         Caption         =   "&Back up database"
      End
      Begin VB.Menu mnrepair 
         Caption         =   "&Repair tables"
      End
      Begin VB.Menu ad 
         Caption         =   "A&mbil database"
      End
      Begin VB.Menu sp 
         Caption         =   "&Setting Printer"
      End
      Begin VB.Menu del 
         Caption         =   "&Hapus Data"
      End
      Begin VB.Menu keluar 
         Caption         =   "&Keluar"
      End
   End
End
Attribute VB_Name = "Mnutama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
Dim teksjalan, aa, cc, bb As String

  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Private Const SYNCHRONIZE       As Long = &H100000
    Private Const INFINITE          As Long = &HFFFF

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



Private Sub ad_Click()
restore.Show 1
End Sub

Private Sub aha_Click()
huruff.Show 1
End Sub

Private Sub b_Click()
lapbeli.Show 1
End Sub

Private Sub bahut_Click()
bayar_hutang.Show
End Sub

Private Sub bank_Click()

End Sub

Private Sub bang_Click()
bank.Show 1
End Sub

Private Sub bapi_Click()
bayar_piutang.Show
End Sub

Private Sub bb_Click()
End Sub


Private Sub dato_Click()
data.Show
End Sub

Private Sub db_Click()
lapbrg.Show 1
End Sub

Private Sub del_Click()
hps.Show 1
End Sub

Private Sub dnt_Click()
kasbank.Show 1
End Sub

Private Sub ds_Click()
With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\supplier.rpt"
  .RetrieveDataFiles
.Formulas(2) = "tgjwb='" & tgjwb & "'"
  .WindowTitle = "Laporan Data Supplier"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
End With
Pesan:
If err.Description <> vbNullString Then
End If

End Sub

Private Sub ebot_Click()
frmAbout.Show 1
End Sub

Private Sub frmnrcawal_Click()
frmneraca.Show
End Sub



Private Sub gpass_Click()
pass.Show 1
End Sub

Private Sub lg_Click()
End Sub


Private Sub htg_Click()
laphutang.Show 1
End Sub

Private Sub kalk_Click()
calc.Show 1
End Sub

Private Sub keluar_Click()
buat_folder2

bekap

End
End Sub


Private Sub kluar_Click()
keluarr.Show
End Sub

Private Sub kur_Click()
End Sub

Private Sub lagi_Click()
laba.Show 1
End Sub

Private Sub langgan_Click()
pelanggan.Show 1
End Sub

Private Sub lpbb_Click()
End Sub

Private Sub lapkom_Click()
lapkomisi.Show
End Sub

Private Sub lps_Click()
lapsesuai.Show 1
End Sub

Private Sub mcpu_Click()
harga.Show
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
Dim Skinpath As String
namatoko = GetSetting("apotekbaleendah", "data", "text1.text", "Toko Farid")
almttoko = GetSetting("apotekbaleendah", "data", "text2.text", "Alamat Farid")
telptoko = GetSetting("apotekbaleendah", "data", "text3.text", "telepon toko")
subnama = GetSetting("apotekbaleendah", "data", "text4.text", "sub nama farid")
tgjwb = GetSetting("apotekbaleendah", "data", "text5.text", "Farid")

adalogo = GetSetting("apotekbaleendah", "data", "adalogo", False)
Me.Top = 0
ceksip
'demo
'konek
'konek2
teksjalan = "                                      Sistem Informasi Penjualan"
hbrg = GetSetting("apotekbaleendah", "huruff", "text1.text", "")
hpju = GetSetting("apotekbaleendah", "huruff", "text2.text", "")
hpb = GetSetting("apotekbaleendah", "huruff", "text3.text", "")
hcus = GetSetting("apotekbaleendah", "huruff", "text4.text", "")
hsup = GetSetting("apotekbaleendah", "huruff", "text5.text", "")
hpo = GetSetting("apotekbaleendah", "huruff", "text6.text", "")
cttn1 = GetSetting("apotekbaleendah", "frmfoot", "text1.text", "")
cttn2 = GetSetting("apotekbaleendah", "frmfoot", "text2.text", "")
cttn3 = GetSetting("apotekbaleendah", "frmfoot", "text3.text", "")
cttn4 = GetSetting("apotekbaleendah", "frmfoot", "text4.text", "")
cttn5 = GetSetting("apotekbaleendah", "frmfoot", "text5.text", "")
cttn6 = GetSetting("apotekbaleendah", "frmfoot", "text6.text", "")
'Skinpath = App.Path & "\skin\steelblue.skn"
    'Skin1.LoadSkin Skinpath
    'Skin1.ApplySkin Me.hwnd
End Sub
Private Sub ceksip()
Set RS = New Recordset
RS.Open "select status from transshift where status='y'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
sip = True
Else
sip = False
End If
End Sub

Private Sub MDIForm_Resize()
'Me.WindowState = 0
End Sub

Private Sub mksr_Click()
kasir.Show
End Sub

Private Sub mnakun_Click()
frmakun.Show
End Sub

Private Sub mnapoteker_Click()
frmapoteker.Show
End Sub

Private Sub mnbar_Click()
frmbarcode.Show
End Sub

Private Sub mnbyrpjk_Click()
frmbyrpjk.Show
End Sub

Private Sub mndo_Click()
frmdo.Show
End Sub

Private Sub mndokter_Click()
frmdokter.Show
End Sub

Private Sub mngudang_Click()
frmgudang.Show
End Sub

Private Sub mnhelp_Click()
frmhelp.Show
End Sub

Private Sub mnhistori_Click()
laphistori.Show
End Sub

Private Sub mnjadwal_Click()
frmjadwal.Show
End Sub

Private Sub mnlapbb_Click()
lapbb.Show
End Sub

Private Sub mnlapbyrpjk_Click()
lapbayarpajak.Show
End Sub

Private Sub mnlapex_Click()
lapexpired.Show
End Sub

Private Sub mnlapgud_Click()
frmstokbrg.Show
End Sub

Private Sub mnlapinfo_Click()
lapfaktur.Show
End Sub

Private Sub mnlapjualbeli_Click()
lapjualbeli.Show
End Sub

Private Sub mnlapjurnal_Click()
lapjurnal.Show
End Sub

Private Sub mnlapneraca_Click()
lapneraca.Show
End Sub

Private Sub mnlaprl_Click()
laprl.Show
End Sub

Private Sub mnlapservis_Click()
lapservis.Show
End Sub

Private Sub mnlapshift_Click()
lapshift.Show
End Sub

Private Sub mnlembar_Click()
With CrystalReport1
  .Reset
  .ReportFileName = serperreport & "\opname.rpt"
  .RetrieveDataFiles
.Formulas(2) = "tgjwb='" & tgjwb & "'"
  .WindowTitle = "Laporan Data Supplier"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
End With
Pesan:
If err.Description <> vbNullString Then
End If

End Sub

Private Sub mnmpaket_Click()
frmpaket.Show
End Sub

Private Sub mnmservis_Click()
frmservis.Show
End Sub

Private Sub mnmutasi_Click()
frmmutasi.Show
End Sub

Private Sub mnoutin_Click()
lapbrg3.Show
End Sub

Private Sub mnpesan_Click()
pemesanan.Show
End Sub

Private Sub mnpromosi_Click()
frmfoot.Show
End Sub

Private Sub mnreg_Click()
frmreg.Show
End Sub

Private Sub mnrepair_Click()
frmrepair.Show
End Sub

Private Sub mnsls_Click()
frmsls.Show
End Sub

Private Sub mntek_Click()
frmproses.Show
End Sub

Private Sub mnteknisi_Click()
frmteknisi.Show
End Sub

Private Sub mntrx_Click()
frmtransaksi.Show
End Sub

Private Sub mts_Click()
lapmutasi.Show
End Sub

Private Sub pesbar_Click()
pemesanan.Show 1
End Sub

Private Sub psg_Click()
sesuai2.Show 1
End Sub

Private Sub sp_Click()
setting.Show 1
End Sub

Private Sub suja_Click()
jalan.Show 1
End Sub


Sub akhir()
buat_folder2
If databes = "Akses" Then
bekap
Else
Dim cmd, FileName As String
FileName = Chr(34) & "D:\backup database\apotekbaleendah.sql" & Chr(34)
    DoEvents

    cmd = Chr(34) & Chr(34) & mysqlfolder & "\bin\mysqldump" & Chr(34) & " -h" & serperdatabes & " -u" & userdb & " -p" & passdb & " --routines --comments apotekbaleendah > " & FileName & """"

    'cmd = "H:\Appserv\MySQL\bin\mysqldump -uroot -ptujuh7 --comments penjualan > " & Filename & ""
    Call execCommand(cmd)

End If
End

End Sub

Private Sub pht_Click()
lappiutang.Show 1
End Sub


Private Sub plg_Click()
lapplg.Show 1
End Sub

Private Sub pmbh_Click()
lapbyru.Show 1
End Sub

Private Sub pmbp_Click()
lapbyrp.Show 1
End Sub

Private Sub pmsk_Click()
masuk.Show
End Sub

Private Sub ps_Click()
sesuai.Show 1
End Sub

Private Sub reju_Click()
frmreju.Show
End Sub

Private Sub retbel_Click()
lapretur.Show
End Sub

Private Sub rtj_Click()
lapreturj.Show 1
End Sub



Private Sub stk_Click()
stokbrg.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(5).Text = Format(Now, "hh:mm:ss")
StatusBar1.Panels(6).Text = Format(Now, "dd mmmm yyyy")
End Sub

Private Sub brg_Click()
Barang.Show
End Sub

Private Sub bup_Click()
back.Show 1
End Sub

Private Sub guna_Click()
pengguna.Show 1
End Sub

Private Sub j_Click()
lapjual.Show 1
End Sub

Private Sub jb_Click()
'If resheight <= 600 Then
Set RS = New Recordset
RS.Open "select status,id from transshift where status='y'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
sip = True
kodesip = RS!id
transaksi.Show

Else
sip = False
frmbuka.Show

End If
'Else
'transaksi2.Show
'End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

'akhir
End
End Sub

Private Sub rb_Click()
frmreli.Show
End Sub

Private Sub supp_Click()
suppl.Show 1
End Sub

Private Sub tb_Click()
pembelian.Show
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1: brg_Click
    Case 2: supp_Click
    Case 3: jb_Click
    Case 4: bup_Click
    Case 5: guna_Click
        Case 6: langgan_Click
    Case 7: mnmutasi_Click
    Case 8: tb_Click

        Case 9: akhir

End Select

End Sub

Private Sub tran_Click()
cash.Show 1
End Sub
