VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} faktur 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "faktur.dsx":0000
End
Attribute VB_Name = "faktur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
sql = "select tblbarang.deskripsi,detiljual.jumlah_brg,penjualan.total_diskon,detiljual.harga_jual,detiljual.diskon,detiljual.diskon,detiljual.total,penjualan.no_penjualan,penjualan.tanggal,penjualan.jumlah,penjualan.total,penjualan.ppn,pelanggan.nama from tblbarang,penjualan,pelanggan,detiljual where tblbarang.kode_brg=detiljual.kode_brg and penjualan.id_pelanggan=pelanggan.id_pelanggan and penjualan.no_penjualan=detiljual.no_penjualan and penjualan.no_penjualan='" & transaksi.notr.Text & "'"
Set ado.Recordset = jual.Execute(sql)
Set rsbarang = New Recordset
sql = "select sum(diskon) as ttl from detiljual where no_penjualan='" & transaksi.notr.Text & "'"
Set rsbarang = jual.Execute(sql)
'ado.Source = "select sum(detiljual.total) as ttl from detiljual where no_penjualan='" & transaksi.notr.Text & "'"
End Sub

 Sub ActiveReport_Terminate()
Unload Me
transaksi.Show
End Sub

Private Sub Detail_Format()
alamat.Text = almt & "                              " & almt2
nama_brg.Text = ado.Recordset!deskripsi
no.Text = val(no.Text) + 1
dis.Text = Format(ado.Recordset!total_diskon - rsbarang!ttl, "#,#")
ppn.Text = Format(ado.Recordset!ppn, "#,#")
total.Text = Format(ado.Recordset!total + ado.Recordset!ppn, "#,#")
'dis.Text = RS!total_diskon - val(dis.Text)
End Sub

