VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form data 
   Caption         =   "IDENTITAS APOTEK"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "data.frx":0000
      TabIndex        =   14
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Hapus Logo"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2655
      Left            =   6840
      TabIndex        =   12
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Logo"
      TabPicture(0)   =   "data.frx":007E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8280
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Open"
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "data.frx":009A
      Top             =   3360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Full Version"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "data.frx":02CE
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "data.frx":033A
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "data.frx":03A4
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "data.frx":0420
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adostream As New adodb.Stream
Private Sub Command1_Click()
On Error Resume Next
SaveSetting "apotekbaleendah", "data", "Text1.text", Text1.Text
namatoko = GetSetting("apotekbaleendah", "data", "text1.text", "Toko Farid")
SaveSetting "apotekbaleendah", "data", "Text5.text", text5.Text
tgjwb = GetSetting("apotekbaleendah", "data", "Text5.text", "Farid")

SaveSetting "apotekbaleendah", "data", "Text2.text", Text2.Text
almttoko = GetSetting("apotekbaleendah", "data", "text2.text", "Alamat farid")
SaveSetting "apotekbaleendah", "data", "Text3.text", text3.Text
telptoko = GetSetting("apotekbaleendah", "data", "text3.text", "Telepon arid")
SaveSetting "apotekbaleendah", "data", "Text4.text", text4.Text
subnama = GetSetting("apotekbaleendah", "data", "text4.text", "Sub nama toko")
Set rsd = New Recordset
rsd.Open "select nama_toko from data_toko", jual, adOpenStatic, adLockOptimistic
If rsd.EOF Then
jual.Execute "insert into data_toko(nama_toko) values('toko')"
End If

Set RS = New Recordset
RS.Open "select * from data_toko", jual, adOpenStatic, adLockOptimistic
If SavePictureToDB(RS, cd1.FileName) = True Then
adalogo = True
       'MsgBox "simpan mahasiswabarberhasil"
       Else
       adalogo = False

    End If
    SaveSetting "apotekbaleendah", "data", "adalogo", adalogo

    RS.Update

MsgBox "Data berhasil diperbarui.", vbInformation, judul
End Sub
Public Function LoadPictureFromDB(RS As adodb.Recordset, _
foto As Image)

    On Error GoTo errLoad
    'Jika record tidak ada
    If RS Is Nothing Then
        Exit Function
    End If
   
    Set adostream = New adodb.Stream
   
    adostream.Type = adTypeBinary
    adostream.Open
   
    adostream.Write RS!logo
    
    'proses menyimpan ke bentuk file
    adostream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
    foto.Picture = LoadPicture("C:\Temp.bmp")
    'proses menghapus file temp.bmp
    Kill ("C:\Temp.bmp")
    LoadPictureFromDB = True

    Exit Function
errLoad:
    LoadPictureFromDB = False
End Function


Public Function SavePictureToDB(RS As adodb.Recordset, _
    sFileName As String)

    On Error GoTo errSimpan
    Dim oPict As StdPicture
    Set oPict = LoadPicture(sFileName)
    'jika gambar tida ditemukan
    If sFileName = "" Then
        jual.Execute "delete from data_toko"
        SavePictureToDB = False
        
        Exit Function
        Else
    End If

    Set adostream = New adodb.Stream
    adostream.Type = adTypeBinary
    adostream.Open
    adostream.LoadFromFile sFileName
      RS!logo = adostream.Read
    Image1.Picture = LoadPicture(sFileName)
    adostream.Close
    SavePictureToDB = True

Exit Function
errSimpan:
    SavePictureToDB = False
End Function


Private Sub Command2_Click()
frmid.Show
End Sub

Private Sub Command3_Click()
cd1.ShowOpen
Image1.Picture = LoadPicture(cd1.FileName)
End Sub
Sub tampillogo()
Set RS = New Recordset
RS.Open "select * from data_toko", jual, adOpenStatic, adLockOptimistic


If LoadPictureFromDB(RS, Image1) Then
End If

End Sub
Private Sub Command4_Click()
tampillogo
End Sub

Private Sub Command5_Click()
If MsgBox("Yakin hapus logo?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from data_toko"
Set Image1.Picture = Nothing
adalogo = False
SaveSetting "apotekbaleendah", "data", "adalogo", adalogo
End Sub

Private Sub Form_Load()
On Error Resume Next
tampillogo
namatoko = GetSetting("apotekbaleendah", "data", "text1.text", "Toko Farid")
almttoko = GetSetting("apotekbaleendah", "data", "text2.text", "Alamat toko")
telptoko = GetSetting("apotekbaleendah", "data", "text3.text", "Telepon farid")
tgjwb = GetSetting("apotekbaleendah", "data", "Text5.text", "Farid")

Text1.Text = namatoko
Text2.Text = almttoko
text3.Text = telptoko
text4.Text = subnama
text5.Text = tgjwb
'If aktipasi = False Then
'Command1.Enabled = False
'Command2.Enabled = True


'Else
'Command1.Enabled = True
'Command2.Enabled = False

'End If
    Skinpath = App.Path & "\skin\galaxy.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
