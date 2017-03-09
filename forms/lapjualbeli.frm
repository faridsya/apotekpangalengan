VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapjualbeli 
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbgol 
      Height          =   315
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Golongan"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Barang"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "lapjualbeli.frx":0000
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox cmbbarang 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "lapjualbeli.frx":0234
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "lapjualbeli.frx":02AE
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   128647171
      CurrentDate     =   37623
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddd, d MMMM yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   128647171
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "lapjualbeli.frx":0324
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "lapjualbeli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbbarang_Click()
Option1.Value = True
End Sub

Private Sub cmbgol_Change()
Option2.Value = True
End Sub

Private Sub Command1_Click()
Dim vawal, vakhir As Double
jual.Execute "drop view if exists summasuk1,vsumkeluar,vkmbrg,vsumretur,vsumretur2,vsumbeli"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `summasuk1` AS select `v1`.`Kode_brg` AS `kode_brg`,`v1`.`deskripsi` AS `deskripsi`,`v1`.`tanggal` AS `tanggal`,sum((`v1`.`masuk1` - `v1`.`keluar1`)) AS `sum(masuk1)` from `vsesuai` `v1` where (`v1`.`tanggal` > _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') group by `v1`.`Kode_brg`;"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vsumbeli` AS select `v2`.`kode_brg` AS `kode_brg`,`v2`.`deskripsi` AS `deskripsi`,`v2`.`tanggal` AS `tanggal`,sum(`v2`.`terbeli`) AS `sum(terbeli)` from `vpembelian` `v2` where (`v2`.`tanggal` > _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') group by `v2`.`kode_brg`;"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vsumkeluar` AS select `v2`.`kode_brg` AS `kode_brg`,`v2`.`deskripsi` AS `deskripsi`,`v2`.`tanggal` AS `tanggal`,sum(`v2`.`terjual`) AS `sum(terjual)` from `vpenjualan` `v2` where (`v2`.`tanggal` > _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') group by `v2`.`kode_brg`;"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vsumretur2` AS select `v1`.`kode_brg` AS `kode_brg`,`v1`.`deskripsi` AS `deskripsi`,`v1`.`tanggal` AS `tanggal`,sum(`v1`.`retur2`) AS `sumretur2` from `vretur2` `v1` where (`v1`.`tanggal` > _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') group by `v1`.`kode_brg`;"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vsumretur` AS select `v1`.`kode_brg` AS `kode_brg`,`v1`.`deskripsi` AS `deskripsi`,`v1`.`tanggal` AS `tanggal`,sum(`v1`.`retur`) AS `sumretur` from `vretur` `v1` where (`v1`.`tanggal` > _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "') group by `v1`.`kode_brg`;"


jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `apotekbaleendah`.`vkmbrg` AS select coalesce(`s1`.`sum(masuk1)`,0) AS `summasuk1`,coalesce(`s3`.`sumretur`,0) AS `sumretur`,coalesce(`s5`.`sumretur2`,0) AS `sumretur2`,`b`.`Kode_brg` AS `kode_brg`,`b`.`Deskripsi` AS `deskripsi`,coalesce((`v1`.`masuk1` - `v1`.`keluar1`),0) AS `masuk1`,coalesce(`v3`.`retur`,0) AS `retur`,coalesce(`v5`.`retur2`,0) AS `retur2`,coalesce(`v2`.`terjual`,0) AS `terjual`," _
& "coalesce(`v4`.`terbeli`,0) AS `terbeli`,`b`.`Stok` AS `stok`,coalesce(`s2`.`sum(terjual)`,0) AS `sum(terjual)`,coalesce(`s4`.`sum(terbeli)`,0) AS `sum(terbeli)`,(((((`b`.`Stok` + coalesce(`s2`.`sum(terjual)`,0)) - coalesce(`s1`.`sum(masuk1)`,0)) - coalesce(`s4`.`sum(terbeli)`,0)) - coalesce(`s3`.`sumretur`,0)) + coalesce(`s5`.`sumretur2`,0)) AS `stokakhir`," _
& "((((((((((`b`.`Stok` + coalesce(`s5`.`sumretur2`,0)) + coalesce(`s2`.`sum(terjual)`,0)) - coalesce(`s4`.`sum(terbeli)`,0)) - coalesce(`s1`.`sum(masuk1)`,0)) - coalesce(`s3`.`sumretur`,0)) + coalesce(`v2`.`terjual`,0)) - coalesce(`v4`.`terbeli`,0)) - coalesce((`v1`.`masuk1` - `v1`.`keluar1`),0)) - coalesce(`v3`.`retur`,0)) + coalesce(`v5`.`retur2`,0)) AS `stokawl` " _
& " from ((((((`apotekbaleendah`.`tblbarang` `B` left join `apotekbaleendah`.`vsesuai` `v1` on(((`b`.`Kode_brg` = `v1`.`Kode_brg`) and (`v1`.`tanggal` between _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "')))) left join `apotekbaleendah`.`vretur` `v3` on(((`b`.`Kode_brg` = `v3`.`kode_brg`) and (`v3`.`tanggal` BETWEEN _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'))) left join `apotekbaleendah`.`vretur2` `v5` on(((`b`.`Kode_brg` = `v5`.`kode_brg`) and (`v5`.`tanggal` BETWEEN _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'))) " _
& " left join `apotekbaleendah`.`vpenjualan` `v2` on(((`v2`.`kode_brg` = `b`.`Kode_brg`) and (`v2`.`tanggal` BETWEEN _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'))) " _
& " left join `apotekbaleendah`.`vpembelian` `v4` on(((`v4`.`kode_brg` = `b`.`Kode_brg`) and (`v4`.`tanggal` BETWEEN _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' and _utf8'" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'))) left join `apotekbaleendah`.`summasuk1` `s1` on((`s1`.`kode_brg` = `b`.`Kode_brg`))) left join `apotekbaleendah`.`vsumkeluar` `s2` on((`s2`.`kode_brg` = `b`.`Kode_brg`))) left join `apotekbaleendah`.`vsumretur` `s3` on((`s3`.`kode_brg` = `b`.`Kode_brg`))) left join `apotekbaleendah`.`vsumretur2` `s5` on((`s5`.`kode_brg` = `b`.`Kode_brg`))) " _
& " left join `apotekbaleendah`.`vsumbeli` `s4` on((`s4`.`kode_brg` = `b`.`Kode_brg`))) group by `b`.`Kode_brg`;" _


Set RS = New Recordset
sql = "select stokawl,stokakhir from vkmbrg where deskripsi='" & cmbbarang.Text & "'"
Set RS = jual.Execute(sql)
If RS.EOF Then Exit Sub

vawal = RS!stokawl
vakhir = RS!stokakhir




With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\jualbeli.rpt"
  .RetrieveDataFiles
.Formulas(1) = "awl='" & vawal & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{vjualbeli.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {vjualbeli.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# and {tblbarang.deskripsi}='" & cmbbarang.Text & "'"
    Else
    .SelectionFormula = "{vjualbeli.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {vjualbeli.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# and {tblbarang.golongan}='" & cmbgol.Text & "'"
End If
  
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode : '+'" & Format(DTPicker1.Value, "dd MMM yyyy") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM yyyy") & "'"
Else
.Formulas(0) = "waktu='Periode : '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
End With
Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If
End Sub
Private Sub databarang()
Set RS = New Recordset
cmbbarang.Clear
sql = "select deskripsi from tblbarang order by deskripsi"
Set RS = jual.Execute(sql)
If RS.EOF Then Exit Sub
RS.MoveFirst
Do While Not RS.EOF
cmbbarang.AddItem RS!deskripsi
RS.MoveNext
Loop
cmbbarang.ListIndex = 0
End Sub
Private Sub datagolongan()
'On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbgol.Clear
sql = "select distinct golongan from tblbarang where golongan is not null and golongan!='' order by golongan"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbgol.AddItem rsbarang!golongan
rsbarang.MoveNext
 Loop
  End If
  cmbgol.ListIndex = 0
rsbarang.Close
End Sub

Private Sub Form_Load()
Ketengah Me
databarang
datagolongan
Option1.Value = True
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")

Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub
