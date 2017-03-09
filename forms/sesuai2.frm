VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form sesuai2 
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6930
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4560
      OleObjectBlob   =   "sesuai2.frx":0000
      Top             =   6120
   End
   Begin apotekbaleendah.ThemedButton Cmdkeluar 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   5760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Keluar"
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
   Begin apotekbaleendah.ThemedButton Cmdsimpan 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Simpan"
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
   Begin apotekbaleendah.ThemedButton Cmdbatal 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Bata&l"
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
   Begin apotekbaleendah.ThemedButton cmdtambah 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Baru"
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8175
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Penyesuain stok gudang"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   5415
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "sesuai2.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Barang"
      TabPicture(1)   =   "sesuai2.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "dbgrid1"
      Tab(1).Control(2)=   "Text5"
      Tab(1).ControlCount=   3
      Begin apotekbaleendah.xFrame frame 
         Height          =   3975
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7011
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
         Begin ACTIVESKINLibCtl.SkinLabel stn2 
            Height          =   255
            Left            =   3000
            OleObjectBlob   =   "sesuai2.frx":026C
            TabIndex        =   26
            Top             =   2400
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel stn 
            Height          =   255
            Left            =   5040
            OleObjectBlob   =   "sesuai2.frx":02CA
            TabIndex        =   25
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   4440
            OleObjectBlob   =   "sesuai2.frx":0328
            TabIndex        =   24
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "sesuai2.frx":0386
            TabIndex        =   22
            Top             =   1680
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Ke toko"
            Height          =   255
            Left            =   4680
            TabIndex        =   23
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   3000
            Width           =   2415
         End
         Begin VB.ComboBox text4 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            Top             =   1680
            Width           =   2655
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   7
            Top             =   2400
            Width           =   975
            _ExtentX        =   1720
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   18
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   115867651
            CurrentDate     =   40299
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Menambah stok"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Mengurang stok"
            Height          =   255
            Left            =   2520
            TabIndex        =   16
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Berat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Alasan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   20
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "stok :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   8
            Left            =   3840
            TabIndex        =   19
            Top             =   2400
            Width           =   615
         End
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -71160
         TabIndex        =   13
         Top             =   4560
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6800
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
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   4560
         Width           =   3255
      End
   End
End
Attribute VB_Name = "sesuai2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode, stri As String
Dim setok As Double
Private Sub Cmdbatal_Click()
awal
Label1.Caption = ""
dbgrid
kosong
cmdtambah.SetFocus
End Sub
Private Sub alsn1()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
sql = "select * from sesuai where jenis='Menambah' order by Alasan"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo1.AddItem rsbarang!alasan
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


  End Sub
Private Sub alsn2()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
sql = "select * from sesuai where jenis='Mengurang' order by Alasan"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo1.AddItem rsbarang!alasan
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


  End Sub

Private Sub kbrg()

text4.Clear
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text4.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

Private Sub Cmdcari_Click()
Dim kata As String

If Cmdcari.Caption = "&Cari" Then
Cmdcari.Caption = "&Refresh"
Frame.Enabled = False
kata = InputBox("Masukkan id pelanggan atau nama pelanggan", "Cari...")
If StrPtr(kata) = 0 Then Exit Sub

If kata = "" Then Exit Sub
cmdsimpan.Enabled = False

sql = "select* from pelanggan where id_pelanggan='" & kata & "' or nama like '%" & kata & "%' "
Set rsplg = New Recordset
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
Set dbgrid1.DataSource = rsplg
Tab1.Tab = 1
Else
MsgBox "Tidak ada", vbOKOnly, judul
Cmdcari.Caption = "&Cari"

dbgrid
awal
End If
Else
dbgrid
Cmdcari.Caption = "&Cari"
End If
End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub


Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
Frame.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True
End If
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from pelanggan where id_pelanggan='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, judul
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
If Text2.Text = "" Then Exit Sub
If Option1.Value = True Then
jls = "Menambah"
Else
If Option3.Value = True Then
jls = "Toko"

Else
jls = "Mengurang"

End If
End If
jual.Execute "insert into sesuai values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & jls & "','" & kode & "'," & val(Text2.Text) & ",'" & Combo1.Text & "')"


jual.Execute "update tblbarang set stok=" & val(stok.Caption) & " where kode_brg='" & kode & "'"
Frame.Enabled = False

If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
dbgrid

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, judul
Frame.Enabled = True
End If

End Sub

Private Sub cmdtambah_Click()

Edit = False
tambah
kosong
Label1.Caption = "Dobel klik untuk mengirimkan nama item"
Tab1.Tab = 0
kbrg
text4.SetFocus
End Sub
Sub tambah()
Edit = False
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Frame.Enabled = True
cmdbatal.Enabled = True
End Sub
Sub kosong()
Text2.Text = ""
text4.Text = ""
stok.Caption = ""
stn.Caption = ""
stn2.Caption = ""
Combo1.Text = ""
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsimpan.SetFocus
End If
End Sub

Private Sub dbgrid1_DblClick()
On Error Resume Next

text4.Text = dbgrid1.Columns(1)
Tab1.Tab = 0
text4_Click
Text2.SetFocus
End Sub



Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = 13 Then
dbgrid1_DblClick
End If

End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdsimpan.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
Text2.Text = val(Text2.Text) * 15
End If
End Sub

Private Sub Form_Load()
Option1.Value = True
Option1_Click
  
awal
Label1.Caption = ""
kbrg
Ketengah Me
dbgrid
DTPicker1.Value = Now
Tab1.Tab = 0
Dim Arq As String
    Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub awal()
Frame.Enabled = False
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
dbgrid1.Enabled = True

End Sub
Sub dbgrid()

sql = "select gudang.kode_brg,tblbarang.deskripsi,tblbarang.satuan,tblbarang.stok as stok_toko,gudang.stok_gudang from tblbarang,gudang where tblbarang.kode_brg=gudang.kode_brg"
Set rsplg = jual.Execute(sql)

Set dbgrid1.DataSource = rsplg


End Sub

Private Sub jb_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub


Private Sub Option1_Click()
alsn1
Combo1.Text = ""
Combo1.Enabled = True

End Sub

Private Sub Option2_Click()
alsn2
Combo1.Text = ""
Combo1.Enabled = True

End Sub

Private Sub Option3_Click()
Combo1.Text = "Toko"
Combo1.Enabled = False
End Sub

Private Sub Text2_Change()
If (val(Text2.Text) > val(stok)) And Option2.Value = True Then
MsgBox "Stok tidak mencukupi", , judul
Text2.Text = ""
End If
If Option1.Value = True Then
stok.Caption = setok + val(Text2.Text)
Else
stok.Caption = setok - val(Text2.Text)
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If
End Sub

Private Sub text1_Click()
sql = "select * from bahanbaku where nama_item='" & Text1.Text & "' "
Set RS = New Recordset
Set RS = jual.Execute(sql)
If Not RS.EOF Then

kode = RS!kode_item

stok.Caption = Format(RS!stok_berat, "#,##.##")

stokk.Caption = RS!stok_karung
Text2.SetFocus
End If

End Sub



Private Sub text4_Click()
On Error Resume Next
stri = Replace(text4.Text, "'", "''")

setok = 0

sql = "select * from gudang,tblbarang where tblbarang.kode_brg=gudang.kode_brg and  (gudang.kode_brg='" & stri & "' or tblbarang.deskripsi='" & stri & "')"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then

kode = rsbarang!kode_brg


setok = rsbarang!stok_gudang
stok.Caption = setok
stn.Caption = rsbarang!satuan
stn2.Caption = rsbarang!satuan

Text2.SetFocus
Else
Tab1.Tab = 1
text5.Text = text4.Text
'kolom2_KeyPress (13)
End If

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4_Click
End If
End Sub
Private Sub text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text7.SetFocus
End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text8.SetFocus
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan_Click
End If
End Sub

Private Sub Text5_Change()
Set dbgrid1.DataSource = Nothing
stri = Replace(text5.Text, "'", "''")

Set rssupp = New Recordset
sql = "select gudang.kode_brg,tblbarang.deskripsi,tblbarang.satuan,tblbarang.stok as stok_toko,gudang.stok_gudang from tblbarang,gudang where tblbarang.kode_brg=gudang.kode_brg and deskripsi like'%" & stri & "%'"

Set rspo = jual.Execute(sql)
Set dbgrid1.DataSource = rspo

End Sub
