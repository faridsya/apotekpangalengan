VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} retur 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "retur.dsx":0000
End
Attribute VB_Name = "retur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
sql = "select retur_beli.no_retur,retur_beli.tanggal,pembelian.tanggal_pembelian,tblsupplier.supplier,detilreturbeli.nama_barang,detilbeli.kembali_brg,detilreturbeli.total from pembelian,detilbeli,tblsupplier,retur_beli,detilreturbeli where detilreturbeli.no_retur=retur_beli.no_retur and detilreturbeli.no_pembelian=pembelian.no_pembelian and detilreturbeli.no_pembelian=detilbeli.no_pembelian and pembelian.no_pembelian=detilbeli.no_pembelian and retur_beli.id_supplier=tblsupplier.id_supplier and retur_beli.tanggal between '" & Format(lapretur.DTPicker1.Value, "YYYY-mm-dd") & "' and '" & Format(lapretur.DTPicker2.Value, "YYYY-mm-dd") & "'"
Set ado.Recordset = jual.Execute(sql)
Set rsbarang = New Recordset
sql = "select (detilbeli.kembali_brg+detilbeli.kembali_uang+detilbeli.Teretur+detilbeli.kembali_uang2) as ttl from pembelian,detilbeli,tblsupplier,retur_beli,detilreturbeli where detilreturbeli.no_retur=retur_beli.no_retur and detilreturbeli.no_pembelian=pembelian.no_pembelian and detilreturbeli.no_pembelian=detilbeli.no_pembelian and pembelian.no_pembelian=detilbeli.no_pembelian and retur_beli.id_supplier=tblsupplier.id_supplier and retur_beli.tanggal between '" & Format(lapretur.DTPicker1.Value, "YYYY-mm-dd") & "' and '" & Format(lapretur.DTPicker2.Value, "YYYY-mm-dd") & "' "
Set rsbarang = jual.Execute(sql)
Set RS = New Recordset
sql = "select (detilbeli.kembali_uang + detilbeli.kembali_uang2) as ttl from pembelian,detilbeli,tblsupplier,retur_beli,detilreturbeli where detilreturbeli.no_retur=retur_beli.no_retur and detilreturbeli.no_pembelian=pembelian.no_pembelian and detilreturbeli.no_pembelian=detilbeli.no_pembelian and pembelian.no_pembelian=detilbeli.no_pembelian and retur_beli.id_supplier=tblsupplier.id_supplier and retur_beli.tanggal between '" & Format(lapretur.DTPicker1.Value, "YYYY-mm-dd") & "' and '" & Format(lapretur.DTPicker2.Value, "YYYY-mm-dd") & "'"
Set RS = jual.Execute(sql)
If lapretur.DTPicker1 = lapretur.DTPicker2 Then
tgl.Text = "Periode :" + Format(lapretur.DTPicker1.Value, "dd MMM YYYY")
Else
tgl.Text = "Periode :" + Format(lapretur.DTPicker1.Value, "dd MMM YYYY") + "-" + Format(lapretur.DTPicker1.Value, "dd MMM YYYY")
End If
End Sub

Private Sub Detail_Format()
On Error Resume Next
teretur.Text = rsbarang!ttl
kembu.Text = RS!ttl
End Sub
