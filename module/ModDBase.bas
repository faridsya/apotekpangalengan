Attribute VB_Name = "ModDBase"

Global jual As adodb.Connection
Public rsbarang As New adodb.Recordset
Public rsbarang2 As New adodb.Recordset
Public rsbarang3 As New adodb.Recordset
Public rsbank As New adodb.Recordset
Public rse1 As New adodb.Recordset
Public rse2 As New adodb.Recordset
Public rse3 As New adodb.Recordset
Public rse4 As New adodb.Recordset
Public rsr As New adodb.Recordset
Public rsstn As New adodb.Recordset
Public rsd As New adodb.Recordset
Public rsbyr As New adodb.Recordset
Public pjgh As Integer
Public rssupp As New adodb.Recordset
Public rspengguna As New adodb.Recordset
Public rsreport As New adodb.Recordset
Public rsbeli As New adodb.Recordset
Public rspo As New adodb.Recordset
Public rsplg As New adodb.Recordset
Public RS2 As New adodb.Recordset
Public rsmt As New adodb.Recordset
Public serper As String
Public serperreport As String

Public passdb As String
Public userdb As String
Public mysqlfolder As String
Public namadb As String
Public portdb As String
Public versiupdate As String


Public serperdatabes As String
Public rshusus As New adodb.Recordset

Public rsbelid As New adodb.Recordset

Public RS As New adodb.Recordset
Public rstrans As New adodb.Recordset
Public rsmurid As New adodb.Recordset
Public Const dbName = "Penjualan.mdb"
Public Edit As Boolean
Public code, DataString, Temp As String





Public Function ConnectDb(user As String, pass As String) As Boolean
Dim sql As String

    ConnectDb = False
    Set jual = New adodb.Connection
       jual.CursorLocation = adUseClient
'jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & serper & "\penjualan.mdb;Jet OLEDB:Database Password=tujuh;"
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & serper & "\penjualan.mdb;Jet OLEDB:Database Password=tujuh;"

     ConnectDb = True
    sql = "select * from pengguna where username='" & user & "' and password='" & pass & "'"
Set rshusus = jual.Execute(sql)
If Not rshusus.EOF Then

 ConnectDb = True
 Else
    ConnectDb = False

 End If

  
    Exit Function
End Function
Public Function ConnectDb2(user As String, pass As String) As Boolean
Dim sql As String

    ConnectDb2 = False
    Set jual = New adodb.Connection
       jual.CursorLocation = adUseClient
jual.Open "DRIVER={MySQL ODBC 5.1 Driver};" _
                & "SERVER=" & serperdatabes & "" _
                & ";DATABASE=apotekbaleendah" _
                & ";USER=" & userdb & "" _
                & ";PORT=" & portdb & ";" _
                & ";PASSWORD=" & passdb & "" _
                & ";OPTION=3;"

     ConnectDb2 = True
     Set rshusus = New Recordset
    sql = "select * from pengguna where username='" & user & "' and password=md5('" & pass & "')"
Set rshusus = jual.Execute(sql)
If Not rshusus.EOF Then

 ConnectDb2 = True
 Else
    ConnectDb2 = False

 End If

  
    Exit Function
End Function

Sub Main()

frmxLogIn.Show
End Sub






