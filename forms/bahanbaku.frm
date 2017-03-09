VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form bahanbaku 
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9315
   Begin TabDlg.SSTab Tab1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Bahan Baku"
      TabPicture(0)   =   "bahanbaku.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbgrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Skin1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SkinLabel1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "MSComm1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "SkinLabel2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "XPCheck1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Cmdkeluar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdtambah"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Cmdedit"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Cmdhapus"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Cmdsimpan"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Cmdbatal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Cmdcari"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Supplier"
      TabPicture(1)   =   "bahanbaku.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ket2"
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(2)=   "cmdcaris"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Cmdcari 
         Caption         =   "&Cari"
         Height          =   255
         Left            =   6480
         TabIndex        =   13
         Top             =   4680
         Width           =   735
      End
      Begin VB.CommandButton Cmdbatal 
         Caption         =   "&Batal"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   4680
         Width           =   735
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Cmdkeluar 
         Caption         =   "&Keluar"
         Height          =   255
         Left            =   5520
         TabIndex        =   12
         Top             =   4680
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "S&impan gambar"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak barco&de"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cop&y barcode"
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CheckBox XPCheck1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   5040
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Kode barang otomatis"
         Height          =   255
         Left            =   6600
         TabIndex        =   16
         Top             =   5040
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "bahanbaku.frx":0038
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   6960
         Top             =   5040
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "bahanbaku.frx":00B4
         TabIndex        =   18
         Top             =   5040
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   720
         OleObjectBlob   =   "bahanbaku.frx":012A
         Top             =   6480
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   23
         Top             =   5400
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BackColor       =   13750737
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin penjualan.xFrame Frame 
         Height          =   3495
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6165
         Caption         =   ""
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeaderGradientBottom=   12611136
         Begin XPControls.XPText text2 
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPControls.XPText text5 
            Height          =   285
            Left            =   1560
            TabIndex        =   6
            Top             =   2640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XPControls.XPText text6 
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   3120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Top             =   600
            Width           =   1695
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   3720
            ScaleHeight     =   975
            ScaleWidth      =   4575
            TabIndex        =   26
            Top             =   2400
            Width           =   4575
         End
         Begin VB.ComboBox text3 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   1680
            Width           =   1695
         End
         Begin VB.ComboBox text4 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   2160
            Width           =   1695
         End
         Begin XPControls.XPText text8 
            Height          =   285
            Left            =   5400
            TabIndex        =   9
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1440
            TabIndex        =   25
            Top             =   0
            Width           =   1815
         End
         Begin VB.ComboBox text7 
            Height          =   315
            Left            =   5400
            TabIndex        =   8
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Nama bahan baku"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Mata uang"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Stok"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Kode bahan baku"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Harga beli"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Satuan 
            Caption         =   "Satuan"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Stol minimal"
            Height          =   255
            Left            =   4200
            TabIndex        =   28
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Supplier"
            Height          =   255
            Left            =   4200
            TabIndex        =   27
            Top             =   600
            Width           =   855
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket2 
         Height          =   375
         Left            =   -74880
         OleObjectBlob   =   "bahanbaku.frx":035E
         TabIndex        =   35
         Top             =   6240
         Width           =   4695
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   5775
         Left            =   -74880
         TabIndex        =   36
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   10186
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin penjualan.ThemedButton cmdcaris 
         Height          =   375
         Left            =   -72000
         TabIndex        =   37
         Top             =   6720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         Caption         =   "&Cari"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "bahanbaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim status As Byte
Dim flag As String
Private Sub DataGrid2_DblClick()
'On Error Resume Next
If Edit = False Then
text7.Text = DataGrid2.Columns(1).Text
Tab1.Tab = 0
text8.SetFocus
End If

End Sub
Private Sub Cmdcaris_Click()
Dim kata As String

If cmdcaris.Caption = "&Cari" Then
kata = InputBox("Masukkan kode supplier atau nama supplier", "Cari...")
If kata = "" Then Exit Sub
sql = "select* from tblsupplier where id_supplier='" & kata & "' or supplier like '%" & kata & "%' "
Set rsplg = New Recordset
Set rsplg = jual.Execute(sql)

If Not rsplg.EOF Then
Set DataGrid2.DataSource = rsplg
cmdcaris.Caption = "&Refresh"
Else
MsgBox "Tidak ada", vbOKOnly, judul
dbgrid3
End If
Else
dbgrid3
cmdcaris.Caption = "&Cari"
End If

End Sub
Sub dbgrid3()

Set rssupp = New Recordset
sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set DataGrid2.DataSource = rssupp
End Sub

Private Sub Check1_Click()
    SaveSetting "penjualan", "Barang", "Check1.value", Check1.Value

End Sub

Private Sub cmdbatal_Click()
awal
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub cmdcari_Click()
If Cmdcari.Caption = "&Refresh" Then
dbgrid
dbgrid1.Refresh
Cmdcari.Caption = "&Cari"
Cmdkeluar.SetFocus
Else
caribrg2.Show 1
End If

End Sub

Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
teks
frame.Enabled = True
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True

End Sub

Sub teks()
On Error Resume Next
cmdtambah.Enabled = True
Text1.Text = dbgrid1.Columns(0).Text
text2.Text = dbgrid1.Columns(1).Text
Text3.Text = dbgrid1.Columns(2).Text
text4.Text = dbgrid1.Columns(3).Text
Text5.Text = dbgrid1.Columns(4).Text
If text4.Text = "Rupiah" Then

text6.Text = dbgrid1.Columns(5).Text
Else
If text4.Text = "Dollar" Then
text6.Text = dbgrid1.Columns(6).Text
End If
End If
text7.Text = dbgrid1.Columns(7).Text
text8.Text = dbgrid1.Columns(8).Text
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from tblbarang2 where kode_brg='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
teks
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub
Private Sub mt()
On Error Resume Next
 text4.Clear
 text4.AddItem "Dollar"
 text4.AddItem "Rupiah"
  End Sub

Private Sub Cmdsimpan_Click()
'On Error GoTo erol
Set rsbarang = New Recordset
A = Text1.Text
rsbarang.Open "select * from tblbarang2 where kode_brg='" & A & "' ", jual, adOpenStatic, adLockOptimistic

If Edit = False Then
Set rsbarang3 = New Recordset
A = text2.Text
rsbarang3.Open "select * from tblbarang2 where deskripsi='" & A & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsbarang3.EOF Then
MsgBox "Nama Produk sudah terdaftar di database"
text2.Text = ""
text2.SetFocus
rs3barang.Close
Exit Sub
End If




rsbarang.AddNew
ubah
rsbarang.Close

dbgrid


If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Exit Sub
Else
awal
End If
Else

If text4.Text <> rsbarang.Fields(3) Then
jual.Execute "delete from tblbarang2 where kode_brg='" & Text1.Text & "'"
rsbarang.AddNew
ubah
dbgrid
Else
ubah
End If
End If
dbgrid

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
frame.Enabled = True
End If
End Sub

Sub ubah()
'On Error Resume Next
Dim dis As String
rsbarang.Fields(0) = Text1.Text
rsbarang.Fields(1) = text2.Text
rsbarang.Fields(2) = Text3.Text
rsbarang.Fields(3) = text4.Text
rsbarang.Fields(4) = val(Text5.Text)

If text4.Text = "Rupiah" Then

rsbarang.Fields(5) = text6.Text
Else
If text4.Text = "Dollar" Then
rsbarang.Fields(6) = text6.Text

End If
End If
rsbarang.Fields(7) = text7.Text
rsbarang.Fields(8) = val(text8.Text)

rsbarang.Update
frame.Enabled = False

End Sub
Sub cmdtambah_Click()
ket2.Caption = "Dobel klik untuk mengirimkan data"

Edit = False
tambah
kosong
Text1.SetFocus
If Check1.Value = Checked Then
kode_oto
text2.SetFocus

End If
kbrg
stn

daba = 0
End Sub
Sub kode_oto()
Dim j As Integer
Dim no As String
Set rsbarang = New Recordset
sql = "Select kode_brg from tblbarang2 where kode_brg like 'Bb%'order by kode_brg Desc"
Set rsbarang = jual.Execute(sql)
If rsbarang.EOF = True Then
Text1.Text = "Bb0001"
Else
j = val(Right(rsbarang(0), 4))
no = "Bb" + Format(Str(j + 1), "0000")
Text1.Text = no

End If
End Sub
Sub buat_gmb()
On Error Resume Next
    Dim sName, retVal, retSave
    Dim ObjGifImg As GIF
    Screen.MousePointer = 11
     'Cdlg1.InitDir = (App.Path & "\Program Barcode\Gambar Barcode\" & Trim(Text1.Text) & ".gif")
    'LblLokasi.Caption = (App.Path & "Program Barcode\Gambar Barcode\" & Trim(Text1.Text) & ".gif")
    
    DoEvents
    Picture1.Picture = Picture1.Image
    Set ObjGifImg = New GIF
    ObjGifImg.SaveGIF Picture1.Image, gsFileName & "\" & Trim(Text1.Text) & ".gif", Picture1.hDC, False, Picture1.Point(0, 0)
    Set ObjGifImg = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
cmdtambah_Click
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If status_isi = "pemesanan aktif" Then
pemesanan.Show
End If

End Sub

Private Sub MSComm1_OnComm()

Dim CheckMyScan As String

Dim CheckForCR As String

Dim MyText As String

Dim CountMe As Integer

Dim Counter As Integer

Dim Number As Integer

On Error GoTo Mscomm11:

If MSComm1.CommEvent = 2 And MSComm1.InBufferCount > 0 Then

CheckMyScan = MSComm1.Input

MyText = CheckMyScan

Text1.Text = CheckMyScan

CountMe = Len(Text1.Text)

Number = 0

Do Until Counter = CountMe

Text1.SelStart = Number

Text1.SelLength = Len(Text1.Text)

CheckForCR = Text1.SelText



If CheckForCR = vbCr Then

Text1.Text = (MyText) - vbEnter

MSComm1.PortOpen = False

MSComm1.PortOpen = True

text2.SetFocus

Exit Sub

End If

DoEvents

Number = Number + 1

Counter = Counter + 1

Loop

End If

Exit Sub

Mscomm11:

MsgBox "A error in reading this bar code", vbOKOnly, "POS"

End Sub

Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
Cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
dbgrid1.Enabled = False

End Sub
Sub kosong()
Text1.Text = ""
text2.Text = ""
Text3.Text = ""
text4.Text = ""
Text5.Text = ""
text6.Text = ""
text7.Text = ""
text8.Text = ""

End Sub

Private Sub Command1_Click()
If Text1.Text <> "" Then
frmFileDialog.Show 1
End If

End Sub

Private Sub Command2_Click()
If IsPrinterInstalled Then

If Text1.Text <> "" Then

    frmPrint.Picture1.Picture = Me.Picture1.Image
    frmPrint.PrintForm
    Printer.EndDoc
End If
Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

End Sub

Private Sub Command3_Click()
    Clipboard.Clear
    Clipboard.SetData Picture1.Image, 2


End Sub

Private Sub dbgrid1_Click()
teks


frame.Enabled = False
Cmdsimpan.Enabled = False

End Sub
Private Sub dbgrid1_GotFocus()
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set jual = New adodb.Connection
        jual.CursorLocation = adUseClient
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/Penjualan.mdb;Jet OLEDB:Database Password=tujuh;"

XPCheck1.Value = GetSetting("Penjualan", "Barang", "XPCheck1.value", Checked)
Check1.Value = GetSetting("Penjualan", "Barang", "Check1.value", Checked)

kbrg
stn
awal
supp
dbgrid3
mt
Ketengah Me
dbgrid
Dim Arq As String
    SkinPath = App.Path & "\skin\winaqua.skn"
    Skin1.LoadSkin SkinPath
    Skin1.ApplySkin Me.hWnd

Arq = ReadINI(App.Path & "\Backup.ini", "Backup_Dir", "Backup_Dir")
RekamKegiatan ("Masuk form barang")
End Sub
Private Sub kbrg()
On Error Resume Next
  Dim i As Long
  Dim j As Long

Text3.Clear
sql = "select * from tblbarang2 order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text3.AddItem rsbarang!Kategori
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    With Text3
    For i = 0 To .ListCount - 1
      For j = .ListCount To (i + 1) Step -1
         If .List(j) = .List(i) Then
           .RemoveItem j
         End If
      Next j
    Next i
  End With


  End Sub
Private Sub stn()
On Error Resume Next

  Dim i As Long
  Dim j As Long
Set rsbarang = New Recordset

Text3.Clear
sql = "select * from tblbarang2 order by kode_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text3.AddItem rsbarang!Satuan
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    With Text3
    For i = 0 To .ListCount - 1
      For j = .ListCount To (i + 1) Step -1
         If .List(j) = .List(i) Then
           .RemoveItem j
         End If
      Next j
    Next i
  End With

  End Sub

Private Sub awal()
ket2.Caption = ""

frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True
dbgrid1.Enabled = True
Edit = True
End Sub

Sub dbgrid()

sql = "select * from tblbarang2"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang


End Sub



Private Sub m_Click()

End Sub

Private Sub text1_Change()
If XPCheck1.Value = Checked Then

If Edit = False Then
Picture1.Visible = True

Call DrawBarcode(Text1, Picture1)
End If
Else
Picture1.Visible = False
End If
End Sub

Private Sub KODE_Change()
kode = A

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"     ' Set the focus to the next control.
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub text3_Click()
            SendKeys "{tab}"     ' Set the focus to the next control.

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4.SetFocus
End If
End Sub

Private Sub text4_Click()
            SendKeys "{tab}"     ' Set the focus to the next control.

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

If KeyAscii = 13 Then
Text5.SetFocus
End If

End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 13 Then
text6.SetFocus
End If

End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 13 Then
text7.SetFocus
End If
If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)
If KeyAscii = 13 Then
text8.SetFocus
End If

End Sub

Private Sub text7_Click()
            SendKeys "{tab}"     ' Set the focus to the next control.

End Sub

Private Sub supp()

  Dim i As Long
  Dim j As Long

text7.Clear
Set rsbarang = New Recordset
sql = "select * from tblsupplier order by id_supplier"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text7.AddItem rsbarang!supplier
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

If KeyAscii = 13 Then
Cmdsimpan_Click
End If
End Sub

Private Sub XPCheck1_Click()
    
    SaveSetting "penjualan", "Barang", "XPCheck1.value", XPCheck1.Value

End Sub
