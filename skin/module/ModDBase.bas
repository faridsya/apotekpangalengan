Attribute VB_Name = "ModDBase"
Option Explicit
Global jual As adodb.Connection
Public rsbarang As New adodb.Recordset
Public rssupp As New adodb.Recordset
Public rspengguna As New adodb.Recordset
Public rsreport As New adodb.Recordset
Public rsbeli As New adodb.Recordset
Public rspo As New adodb.Recordset
Public rsplg As New adodb.Recordset

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
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/Penjualan.mdb;Jet OLEDB:Database Password=tujuh;"
     ConnectDb = True
    sql = "select * from pengguna where username='" & user & "' and password='" & pass & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
 ConnectDb = True
 Else
    ConnectDb = False

 End If

  
    Exit Function
End Function

Sub Main()
frmxLogIn.Show
End Sub






