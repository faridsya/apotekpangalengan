Attribute VB_Name = "mysqlDSN"
'************************************************************
'**    Created by Vipul Patel                              **
'**    http://www.mr-cad.com/                              **
'**    email: vipul_p44@yahoo.com                          **
'**    Use this as you like, no copyright, no restrictions **
'**    I assume NO Responsibility for this code.           **
'**    Copy, Use, Revise, Or even Distribute as your code  **
'**    Thank You !                                         **
'************************************************************

Private Const KEY_QUERY_VALUE = &H1
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_DWORD = 4

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Function DoesKeyExist(RegKeyPath As String, _
    RegKeyName As String, _
    ByRef RegKeyValue As String) As Boolean
    Dim DoesIt As Boolean
    Dim Result As Long
    Dim hKey As Long
    Result = RegOpenKeyEx(HKEY_LOCAL_MACHINE, RegKeyPath, 0&, KEY_QUERY_VALUE, hKey)


    If Result <> ERROR_SUCCESS Then
        DoesKeyExist = False
        Exit Function
    End If
    Result = RegQueryValueEx(hKey, RegKeyName, 0&, REG_SZ, ByVal RegKeyValue, Len(RegKeyValue))
    RegCloseKey (hKey)


    If Result <> ERROR_SUCCESS Then
        DoesKeyExist = False
        Exit Function
    End If
    DoesKeyExist = True
End Function


Public Function checkMySQLDriver(ByRef DriverODBC As String) As Boolean
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim DoesIt As Boolean
    
    
    DoesIt = False
    'edit here to change the driver information
    RegKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\MySQL ODBC 5.1 Driver"
    RegKeyName = "Driver"
    RegKeyValue = String(255, Chr(32))
    


    If DoesKeyExist(RegKeyPath, RegKeyName, RegKeyValue) Then
        DriverODBC = RegKeyValue
        DoesIt = True
    Else
        DoesIt = False
    End If
    
    checkMySQLDriver = DoesIt
End Function


Public Function MySQLDSNWanted(NameDSN As String) As Boolean
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim DoesIt As Boolean
    
    RegKeyPath = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
    RegKeyName = NameDSN
    RegKeyValue = String(255, Chr(32))
    
    If DoesKeyExist(RegKeyPath, RegKeyName, RegKeyValue) Then
        DoesIt = True
    Else
        DoesIt = False
    End If
    
    MySQLDSNWanted = DoesIt
    
End Function


Public Function MakeMySQLDSN(DriverODBC As String, _
    NameDSN As String) As Boolean
    
    Dim hKey As Long
    Dim RegKeyPath As String
    Dim RegKeyName As String
    Dim RegKeyValue As String
    Dim lKeyValue As Long
    Dim Result As Long
    Dim lSize As Long
    Dim szEmpty As String
    
    szEmpty = Chr(0)
    
    
    lSize = 4
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
    NameDSN, hKey)
    

    If Result <> ERROR_SUCCESS Then
        MakeMySQLDSN = False
        Exit Function
    End If
    
    ' For User ID- Uncomment and add back if you like,
    ' this make a blank registry entry for UID
    'Result = RegSetValueExString(hKey, "UID", 0&, REG_SZ, _
    'szEmpty, Len(szEmpty))
    
    'Camel for vipul patel
    'Edit the next line to reflect your Server Name
    RegKeyValue = serperdatabes
    Result = RegSetValueExString(hKey, "Server", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    'Working with this for now this is the Driver name
    RegKeyValue = DriverODBC
    Result = RegSetValueExString(hKey, "Driver", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    'Edit the Next Line to Revise Description
    RegKeyValue = "Ini"
    Result = RegSetValueExString(hKey, "Description", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    'Edit the next line to Revise Database Name
    RegKeyValue = "apotekbaleendah"
    Result = RegSetValueExString(hKey, "Database", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    'Working with this for now this is the user logged on
    RegKeyValue = userdb
    Result = RegSetValueExString(hKey, "User", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
                    
    'Edit the Next Line to Revise Database Password
    RegKeyValue = passdb
    Result = RegSetValueExString(hKey, "Password", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    'Edit the next line to Revise Port
    RegKeyValue = portdb
    Result = RegSetValueExString(hKey, "Port", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
                    
    'Working with this for now this is for Option
    RegKeyValue = ""
    Result = RegSetValueExString(hKey, "Option", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
                    
    'Working with this for now this is for Stmt
    RegKeyValue = ""
    Result = RegSetValueExString(hKey, "Stmt", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
                    
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, _
    RegKeyPath, _
    hKey)


    If Result <> ERROR_SUCCESS Then
        MakeSQLDSN = False
        Exit Function
    End If
    Result = RegSetValueExString(hKey, "ImplicitCommitSync", 0&, REG_SZ, _
    szEmpty, Len(szEmpty))
    RegKeyValue = "Yes"
    Result = RegSetValueExString(hKey, "UserCommitSync", 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    lKeyValue = 2048
    Result = RegSetValueExLong(hKey, "MaxBufferSize", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 5
    Result = RegSetValueExLong(hKey, "PageTimeout", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    lKeyValue = 3
    Result = RegSetValueExLong(hKey, "Threads", 0&, REG_DWORD, _
    lKeyValue, 4)
    
    Result = RegCloseKey(hKey)
    Result = RegCreateKey(HKEY_LOCAL_MACHINE, _
    "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", _
    hKey)
    
    If Result <> ERROR_SUCCESS Then
        MakeSQLDSN = False
        Exit Function
    End If
    
    RegKeyValue = "MySQL ODBC 5.1 Driver"
    Result = RegSetValueExString(hKey, NameDSN, 0&, REG_SZ, _
    RegKeyValue, Len(RegKeyValue))
    
    Result = RegCloseKey(hKey)
    MakeMySQLDSN = True
End Function
Sub installodbc()
Dim DriverODBC As String
Dim NameDSN As String

NameDSN = "apotekbaleendah"
DriverODBC = String(255, Chr(32))
If Not checkMySQLDriver(DriverODBC) Then
    End If

    'Does the DSN name already exist?
    If (MySQLDSNWanted(NameDSN)) = True Then
    Else
        If Not MakeMySQLDSN(DriverODBC, NameDSN) Then
        End If
    End If
End Sub
