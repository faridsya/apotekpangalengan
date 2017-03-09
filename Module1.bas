Attribute VB_Name = "Module1"
Global konek As adodb.Connection
Sub konekdb()
Set konek = New adodb.Connection
konek.CursorLocation = adUseClient
konek.Open "DRIVER={MySQL ODBC 5.1 Driver};server=localhost;database=penjualan;user=root;password=tujuh7;"

End Sub

