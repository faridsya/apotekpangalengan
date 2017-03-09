VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form sesuai 
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6930
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Baru"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8175
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   4800
         OleObjectBlob   =   "sesuai.frx":0000
         Top             =   360
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Penyesuain stok obat"
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
         TabIndex        =   3
         Top             =   120
         Width           =   4695
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "sesuai.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data obat"
      TabPicture(1)   =   "sesuai.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5"
      Tab(1).Control(1)=   "dbgrid1"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Data penyesuaian"
      TabPicture(2)   =   "sesuai.frx":026C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command1"
      Tab(2).Control(1)=   "lv1"
      Tab(2).ControlCount=   2
      Begin apotekbaleendah.xFrame frame 
         Height          =   4575
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   6375
         _extentx        =   11245
         _extenty        =   8070
         caption         =   ""
         enabled         =   -1  'True
         font            =   "sesuai.frx":0288
         fontbold        =   0   'False
         fontitalic      =   0   'False
         fontsize        =   8.25
         fontstrikethru  =   0   'False
         fontunderline   =   0   'False
         headergradientbottom=   12611136
         Begin VB.ComboBox cmbakun 
            Height          =   315
            Left            =   2280
            TabIndex        =   29
            Top             =   3960
            Visible         =   0   'False
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "sesuai.frx":02B4
            TabIndex        =   27
            Top             =   2280
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel stn 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "sesuai.frx":0312
            TabIndex        =   26
            Top             =   2280
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Mengurang stok"
            Height          =   255
            Left            =   3120
            TabIndex        =   21
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Menambah stok"
            Height          =   255
            Left            =   720
            TabIndex        =   20
            Top             =   480
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   2280
            TabIndex        =   19
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   123994115
            CurrentDate     =   40299
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   2280
            TabIndex        =   18
            Top             =   3000
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
         Begin VB.ComboBox text4 
            Height          =   315
            Left            =   2280
            TabIndex        =   17
            Top             =   1680
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   600
            OleObjectBlob   =   "sesuai.frx":0370
            TabIndex        =   16
            Top             =   1680
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2280
            TabIndex        =   15
            Top             =   3480
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   600
            OleObjectBlob   =   "sesuai.frx":03EA
            TabIndex        =   14
            Top             =   2280
            Width           =   1215
         End
         Begin VB.ComboBox cmbgudang 
            Height          =   315
            Left            =   2280
            TabIndex        =   13
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Akun"
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
            Left            =   600
            TabIndex        =   28
            Top             =   3960
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   4320
            TabIndex        =   25
            Top             =   2280
            Width           =   615
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
            Left            =   600
            TabIndex        =   24
            Top             =   3480
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
            Left            =   600
            TabIndex        =   23
            Top             =   3000
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
            Left            =   600
            TabIndex        =   22
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   -71160
         TabIndex        =   5
         Top             =   4560
         Width           =   2775
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   1
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
      Begin MSComctlLib.ListView lv1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   6
         Top             =   600
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6588
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tanggal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jenis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kode barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama barang"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id "
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Alasan"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Gudang"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   4560
         Width           =   3255
      End
   End
End
Attribute VB_Name = "sesuai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim kode, stri, kodegdg As String
Dim setok As Double
Sub dbgridse()
On Error Resume Next

Set rstrans = New Recordset

sql = "select nama_gudang,tanggal,id,jenis,sesuai.kode_brg,deskripsi,jumlah,alasan from sesuai,tblbarang,gudang where sesuai.kode_brg=tblbarang.kode_brg and sesuai.kode_gudang=gudang.kode_gudang order by id desc"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
LV1.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = Format(![tanggal], "dd MMM yyyy")

                l.SubItems(2) = ![jenis]
l.SubItems(3) = ![kode_brg]
l.SubItems(4) = ![deskripsi]
l.SubItems(5) = ![jumlah]
l.SubItems(6) = ![id]
l.SubItems(7) = ![alasan]
l.SubItems(8) = ![nama_gudang]

    .MoveNext
    Loop
End With

End Sub

Private Sub cmbgudang_Click()
getkodegdg
Set rsd = New Recordset
rsd.Open " select jumlah from stokgudang where kode_brg='" & kode & "' and kode_gudang='" & kodegdg & "'", jual, adOpenStatic, adLockOptimistic
setok = rsd!jumlah
stok.Caption = setok
End Sub

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
sql = "select * from sesuai where jenis='Menambah' group by alasan order by Alasan"
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
    


  End Sub
Private Sub alsn2()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
sql = "select * from sesuai where jenis='Mengurang' group by alasan order by Alasan"
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
    


  End Sub

Private Sub kbrg()

text4.Clear
sql = "select deskripsi from tblbarang order by deskripsi"
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


Private Sub Cmdsimpan_Click()
On Error GoTo erol
If cmbgudang.Text = "" Then Exit Sub
getkodegdg
If Text2.Text = "" Then Exit Sub
If Option1.Value = True Then
jls = "Menambah"
Else
jls = "Mengurang"
End If



jual.Execute "insert into sesuai values('','" & Format(DTPicker1, "YYYY-mm-dd") & "','" & jls & "','" & kode & "'," & val(Text2.Text) & ",'" & Combo1.Text & "','" & kodegdg & "','" & cmbakun.Text & "')"
Frame.Enabled = False

If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
dbgrid
Exit Sub
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
Combo1.Text = ""
cmbgudang.Text = ""
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsimpan.SetFocus
End If
End Sub

Private Sub Command1_Click()
If LV1.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin dibatalkan?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from sesuai where id='" & LV1.SelectedItem.SubItems(6) & "'"
dbgridse
dbgrid
MsgBox "Berhasil dibatalkan", vbInformation, judul
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
Sub lisakun()
If Option1.Value = True Then
cmbakun.Clear
cmbakun.AddItem "Koreksi stok"
cmbakun.AddItem "Pendapatan stok bertambah"
cmbakun.Text = "Koreksi stok"
Else
cmbakun.Clear
cmbakun.AddItem "Koreksi stok"
cmbakun.AddItem "Beban stok berkurang"
cmbakun.Text = "Koreksi stok"

End If
End Sub
Private Sub datagdg()

cmbgudang.Clear
sql = "select nama_gudang,jumlah from gudang,stokgudang where gudang.kode_gudang=stokgudang.kode_gudang and kode_brg='" & kode & "' group by stokgudang.kode_gudang order by nama_gudang"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
cmbgudang.AddItem rsplg!nama_gudang
rsplg.MoveNext
 Loop
  rsplg.MoveFirst
  End If

rsplg.Close
  End Sub
  Sub getkodegdg()
Set rsd = New Recordset
rsd.Open "select kode_gudang from gudang where nama_gudang='" & cmbgudang.Text & "' or kode_gudang='" & cmbgudang.Text & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsd.EOF Then
kodegdg = rsd!kode_gudang
Else
kodegdg = "utama"
End If

End Sub

Private Sub Form_Load()
Option1.Value = True
Option1_Click
'dbgridse
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

sql = "select * from tblbarang order by deskripsi"
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
lisakun
End Sub

Private Sub Option2_Click()
alsn2
lisakun
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 2 Then
dbgridse

End If
End Sub

Private Sub Text2_Change()
If (val(Text2.Text) > val(setok)) And Option2.Value = True Then
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

sql = "select * from tblbarang  where deskripsi='" & stri & "' or kode_brg='" & stri & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then

kode = rsbarang!kode_brg
stn.Caption = rsbarang!satuan



cmbgudang.SetFocus
Else
Tab1.Tab = 1
text5.Text = text4.Text
'kolom2_KeyPress (13)
End If
datagdg
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

sql = "select  * from tblbarang where deskripsi like'%" & stri & "%'"
Set rspo = jual.Execute(sql)
Set dbgrid1.DataSource = rspo

End Sub

