VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapmos 
   Caption         =   "MOST TOP SELLING"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   7800
      OleObjectBlob   =   "lapmos.frx":0000
      TabIndex        =   15
      Top             =   6240
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "lapmos.frx":0088
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "lapmos.frx":010E
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "lapmos.frx":017A
      TabIndex        =   12
      Top             =   1560
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "lapmos.frx":01E8
      Top             =   5760
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Bulan"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Semua"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   6240
      Width           =   2655
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Text            =   "Combo5"
      Top             =   3480
      Width           =   1815
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "lapmos.frx":041C
      Left            =   1320
      List            =   "lapmos.frx":0444
      TabIndex        =   5
      Text            =   "bln"
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "lapmos.frx":046F
      Left            =   2040
      List            =   "lapmos.frx":0497
      TabIndex        =   4
      Text            =   "tahun"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvbrg 
      Height          =   5655
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rank"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode barang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama barang"
         Object.Width           =   4234
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kategori"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Satuan"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Jumlah item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Jumlah untung"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Jumlah penjualan"
         Object.Width           =   2647
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Periode :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "lapmos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdcetak_Click()
simpan
With CrystalReport1
  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\mos.rpt"
  .RetrieveDataFiles
  If Combo5.Text = "untung" Then
  .SortFields(0) = "-{lapmos.untung}"
  Else
    If Combo5.Text = "barang" Then
  .SortFields(0) = "-{lapmos.item}"
  Else
    If Combo5.Text = "jual" Then
  .SortFields(0) = "-{lapmos.jual}"
  End If
  End If
  End If

  

  

  .WindowTitle = "Laporan Most Top Selling"
q = "1 / Combo1.Text / 2000"
A = MonthName(Combo4.Text)
b = Combo3.Text
.Formulas(0) = "waktu='Periode: '+ '" & A & "'+'-'+'" & b & "'"
.Formulas(2) = "tgjwb='" & tgjwb & "'"

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



Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command1.Caption = "Mohon tunggu"
dbgrid
lvbrg.Refresh
Command1.Enabled = True
Command1.Caption = "&Proses"

End Sub
Sub simpan()
jual.Execute "delete from lapmos"
For z = 1 To lvbrg.ListItems.count
sql = "insert into lapmos values('" & lvbrg.ListItems(z).SubItems(1) & "','" & Replace(lvbrg.ListItems(z).SubItems(2), "'", "''") & "','" & _
lvbrg.ListItems(z).SubItems(3) & "','" & lvbrg.ListItems(z).SubItems(4) & "','" & lvbrg.ListItems(z).SubItems(5) & "','" & Format(lvbrg.ListItems(z).SubItems(6), Number) & "','" & Format(lvbrg.ListItems(z).SubItems(7), Number) & "')"
jual.Execute (sql)



    Next z

End Sub
Private Sub Form_Load()
Ketengah Me
ktgr
Combo2.AddItem "Semua"
Combo2.AddItem ">0"
Combo2.AddItem "=0"
Combo2.Text = "Semua"
Combo5.AddItem "untung"
Combo5.AddItem "barang"
Combo5.AddItem "jual"
Combo5.Text = "untung"
Combo4.Text = Format(Now, "MM")
Combo3.Text = Format(Now, "YYYY")
Option2.Value = True
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub ktgr()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
Combo1.AddItem "Semua"
sql = "select * from tblbarang order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo1.AddItem rsbarang!kategori
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    With Combo1
    For I = 0 To .ListCount - 1
      For j = .ListCount To (I + 1) Step -1
         If .List(j) = .List(I) Then
           .RemoveItem j
         End If
      Next j
    Next I
  End With
Combo1.Text = "Semua"

  End Sub
Sub dbgrid()
On Error Resume Next
If Combo5.Text = "untung" Then
v = "utg"
Else
If Combo5.Text = "barang" Then
v = "jum"
Else
v = "ttl"
End If
End If
If Option2.Value = True Then
sql = "select baru2.* from(select baru1.*,@rank:=@rank+1 rank from (select t.kode_brg,deskripsi,kategori,t.satuan,t1.kode,coalesce(jum,0) jum,coalesce(ttl,0) ttl ,coalesce(utg,0) utg from tblbarang t LEFT join (select d.kode_brg kode,coalesce(sum(jumlah_brg),0) jum,coalesce(sum(d.total),0) ttl,coalesce(sum((`d`.`Harga_jual` - `d`.`Harga_beli`) * (`d`.`Jumlah_brg`) - (`d`.`diskon`)),0) utg from " _
& "penjualan p,detiljual d where p.no_penjualan=d.no_penjualan and month(tanggal)='" & Combo4.Text & "' and year(tanggal)='" & Combo3.Text & "' group by d.kode_brg) t1 on t.kode_brg=t1.kode) baru1 " _
& "join (select @rank:=0) baru2 order by " & v & " desc) baru2"

'sql anyar= "select baru2.* from(select baru1.*,@rank:=@rank+1 rank from (select t.kode_brg,deskripsi,kategori,t.satuan,t1.kode,coalesce(jum,0) jum,coalesce(ttl,0) ttl ,coalesce(utg,0) utg from tblbarang t LEFT join (select d.kode_brg kode,coalesce(sum(jumlah_brg-teretur),0) jum,coalesce(sum(d.total-(d.teretur*d.harga_jual*(d.total/(d.total+d.diskon)))),0) ttl,coalesce(sum((`d`.`Harga_jual` - `d`.`Harga_beli`) * (`d`.`Jumlah_brg`-d.teretur) - (`d`.`diskon`-((d.teretur/d.jumlah_brg)*d.diskon))),0) utg from " _
& "penjualan p,detiljual d where p.no_penjualan=d.no_penjualan and month(tanggal)='" & Combo4.Text & "' and year(tanggal)='" & Combo3.Text & "' group by d.kode_brg) t1 on t.kode_brg=t1.kode) baru1 " _
& "join (select @rank:=0) baru2 order by " & v & " desc) baru2"
Else
sql = "select baru1.* from (select t2.*,@rank:=@rank+1 rank from(select `t`.`Kode_brg`,`t`.`Deskripsi` ,`t`.`kategori` ,`t`.`Satuan` ,coalesce(sum(`d`.`Jumlah_brg`),0) AS `jum`,coalesce(sum(d.total),0) ttl,coalesce(sum((`d`.`Harga_jual` - `d`.`Harga_beli`) * (`d`.`Jumlah_brg`) - (`d`.`diskon`)),0) utg " _
& "from (`penjualan`.`tblbarang` `t` left join `penjualan`.`detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) " _
& "group by `t`.`Kode_brg` ) t2 join (select @rank:=0) t1 order by  " & v & " desc) baru1"

' anyar sql = "select baru1.* from (select t2.*,@rank:=@rank+1 rank from(select `t`.`Kode_brg`,`t`.`Deskripsi` ,`t`.`kategori` ,`t`.`Satuan` ,coalesce(sum(`d`.`Jumlah_brg`-d.teretur),0) AS `jum`,coalesce(sum(d.total-(d.teretur*d.harga_jual*(d.total/(d.total+d.diskon)))),0) ttl,coalesce(sum((`d`.`Harga_jual` - `d`.`Harga_beli`) * (`d`.`Jumlah_brg`-d.teretur) - (`d`.`diskon`-((d.teretur/d.jumlah_brg)*d.diskon))),0) utg " _
& "from (`penjualan`.`tblbarang` `t` left join `penjualan`.`detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) " _
& "group by `t`.`Kode_brg` ) t2 join (select @rank:=0) t1 order by  " & v & " desc) baru1"

End If
jual.Execute (sql)

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvbrg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , ![rank])
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![kategori]

                l.SubItems(4) = ![satuan]
                l.SubItems(5) = IIf(IsNull(rstrans![jum] = True), 0, rstrans![jum])

               l.SubItems(6) = IIf(IsNull(rstrans![utg] = True), 0, Format(rstrans![utg], "#,#"))

 l.SubItems(7) = IIf(IsNull(rstrans![ttl] = True), 0, Format(rstrans![ttl], "#,#"))

    .MoveNext
    Loop
End With


  



End Sub
Sub Dbgrid2()
On Error Resume Next
If Combo5.Text = "untung" Then
v = "utg"
Else
If Combo5.Text = "barang" Then
v = "jum"
Else
v = "ttl"
End If
End If
If Option2.Value = True Then
sql = "select baru2.* from(select baru1.*,@rank:=@rank+1 rank from (select t.kode_brg,deskripsi,kategori,t.satuan,t1.kode,coalesce(jum,0) jum,coalesce(ttl,0) ttl ,coalesce(utg,0) utg from tblbarang t LEFT join (select d.kode_brg kode,coalesce(sum(jumlah_brg),0) jum,coalesce(sum(d.total),0) ttl,coalesce(sum((`d`.`Harga_jual` - `d`.`Harga_beli`) * `d`.`Jumlah_brg` - `d`.`diskon`),0) utg from " _
& "penjualan p,detiljual d where p.no_penjualan=d.no_penjualan and month(tanggal)='" & Combo4.Text & "' and year(tanggal)='" & Combo3.Text & "' group by d.kode_brg) t1 on t.kode_brg=t1.kode) baru1 " _
& "join (select @rank:=0) baru2 order by " & v & " desc) baru2 where deskripsi like '%" & Text1.Text & "%'"
Else
sql = "select baru1.* from (select t2.*,@rank:=@rank+1 rank from(select `t`.`Kode_brg`,`t`.`Deskripsi` ,`t`.`kategori` ,`t`.`Satuan` ,coalesce(sum(`d`.`Jumlah_brg`),0) AS `jum`,coalesce(sum(`d`.`Total`),0) AS `ttl`,coalesce(sum((((`d`.`Harga_jual` - `d`.`Harga_beli`) * `d`.`Jumlah_brg`) - `d`.`diskon`)),0) AS `utg` " _
& "from (`penjualan`.`tblbarang` `t` left join `penjualan`.`detiljual` `d` on((`d`.`Kode_brg` = `t`.`Kode_brg`))) " _
& "group by `t`.`Kode_brg` ) t2 join (select @rank:=0) t1 order by  " & v & " desc) baru1 where deskripsi like '%" & Text1.Text & "%'"
End If
jual.Execute (sql)

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvbrg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , ![rank])
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![kategori]

                l.SubItems(4) = ![satuan]
                l.SubItems(5) = IIf(IsNull(rstrans![jum] = True), 0, rstrans![jum])

               l.SubItems(6) = IIf(IsNull(rstrans![utg] = True), 0, Format(rstrans![utg], "#,#"))

 l.SubItems(7) = IIf(IsNull(rstrans![ttl] = True), 0, Format(rstrans![ttl], "#,#"))

    .MoveNext
    Loop
End With


  



End Sub

Private Sub Text1_Change()
Dbgrid2
End Sub
