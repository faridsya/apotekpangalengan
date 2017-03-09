VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmutasi 
   Caption         =   "Mutasi gudang"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6915
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcari 
      Caption         =   "&Cari"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   8
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   5280
         OleObjectBlob   =   "frmmutasi.frx":0000
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
         Caption         =   "Mutasi Gudang"
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
         Left            =   600
         TabIndex        =   9
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "frmmutasi.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data barang"
      TabPicture(1)   =   "frmmutasi.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SkinLabel1"
      Tab(1).Control(1)=   "txtcari"
      Tab(1).Control(2)=   "LV1"
      Tab(1).Control(3)=   "Label1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Data Mutasi"
      TabPicture(2)   =   "frmmutasi.frx":026C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lv2"
      Tab(2).ControlCount=   1
      Begin apotekbaleendah.xFrame frame 
         Height          =   3615
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6376
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
         Begin VB.CommandButton Command1 
            Caption         =   "&Cari"
            Height          =   255
            Left            =   3720
            TabIndex        =   30
            Top             =   1080
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel lblstok1 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "frmmutasi.frx":0288
            TabIndex        =   29
            Top             =   1560
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmmutasi.frx":02E6
            TabIndex        =   28
            Top             =   1560
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmmutasi.frx":034C
            TabIndex        =   27
            Top             =   2040
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel lblstok2 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "frmmutasi.frx":03B2
            TabIndex        =   26
            Top             =   2040
            Width           =   1095
         End
         Begin XPControls.XPText txtkode 
            Height          =   285
            Left            =   1800
            TabIndex        =   25
            Top             =   1080
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
         Begin XPControls.XPText txtjum 
            Height          =   285
            Left            =   1800
            TabIndex        =   24
            Top             =   2640
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSComCtl2.DTPicker tgl 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dddd, d MMMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Top             =   600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMM yyyy"
            Format          =   115539971
            CurrentDate     =   37623
         End
         Begin VB.ComboBox cmb1 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   1560
            Width           =   1695
         End
         Begin VB.ComboBox cmb2 
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ke Gudang"
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
            Height          =   375
            Index           =   8
            Left            =   360
            TabIndex        =   20
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Dari gudang"
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
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   1560
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
            Left            =   360
            TabIndex        =   18
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Kode barang"
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
            Height          =   375
            Index           =   7
            Left            =   360
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
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
            Height          =   375
            Index           =   4
            Left            =   360
            TabIndex        =   16
            Top             =   2640
            Width           =   1335
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   -74760
         OleObjectBlob   =   "frmmutasi.frx":0410
         TabIndex        =   13
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   -72960
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Stok utama"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   14
         Top             =   720
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id sales"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal"
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Dari gudang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ke gudang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   3840
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmmutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kodegdg, kodegdg2 As String

Private Sub cmb1_Click()
getkodegdg

If cmb2.Text = cmb1.Text Then
MsgBox "Tidak boleh dari gudang yang sama", vbCritical, judul
Exit Sub
End If
Set RS = New Recordset
RS.Open "select coalesce(jumlah,0) as jumlah from stokgudang where kode_gudang='" & kodegdg & "' and kode_brg='" & txtkode.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Tidak ada di gudang ini", vbCritical, judul
Exit Sub
Else
lblstok1.Caption = RS!jumlah
End If
End Sub
  Sub getkodegdg()
Set rsd = New Recordset
rsd.Open "select kode_gudang from gudang where nama_gudang='" & cmb1.Text & "' or kode_gudang='" & cmb1.Text & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsd.EOF Then
kodegdg = rsd!kode_gudang
Else
kodegdg = "utama"
End If

End Sub
  Sub getkodegdg2()
Set rsd = New Recordset
rsd.Open "select kode_gudang from gudang where nama_gudang='" & cmb2.Text & "' or kode_gudang='" & cmb2.Text & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsd.EOF Then
kodegdg2 = rsd!kode_gudang
Else
kodegdg2 = "utama"
End If

End Sub

Private Sub cmb1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub cmb2_Click()
getkodegdg2

If cmb2.Text = cmb1.Text Then
MsgBox "Tidak boleh dari gudang yang sama", vbCritical, judul
Exit Sub
End If
Set RS = New Recordset
RS.Open "select coalesce(jumlah,0) as jumlah from stokgudang where kode_gudang='" & kodegdg2 & "' and kode_brg='" & txtkode.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Tidak ada di gudang ini", vbCritical, judul
Exit Sub
Else
lblstok2.Caption = RS!jumlah
End If

End Sub

Private Sub Cmdbatal_Click()
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
dbgridtrans

kosong
cmdtambah.SetFocus
End Sub

Private Sub Cmdcari_Click()
Tab1.Tab = 1
txtcari.SetFocus

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub

Sub teks()

cmdtambah.Enabled = True
With lv2.SelectedItem
tgl.Value = .SubItems(2)
txtkode.Text = .SubItems(3)
cmb1.Text = .SubItems(5)
cmb2.Text = .SubItems(6)
txtjum.Text = .SubItems(7)

End With
End Sub
Private Sub cmdedit_Click()
Tab1.Tab = 0
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
frame.Enabled = True

Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from mutasigudang where id='" & lv2.SelectedItem.SubItems(1) & "'"
jual.Execute (sql)
dbgridtrans

kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdsimpan_Click()
 On Error GoTo erol
If val(txtjum.Text) = 0 Then Exit Sub
If cmb1.Text = "" Or cmb2.Text = "" Then Exit Sub
If cmb1.Text = cmb2.Text Then
MsgBox "Tidak boleh dari gudang yang sama", vbCritical, judul
Exit Sub
End If



If Edit = False Then
jual.Execute "insert into mutasigudang values('','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & txtkode.Text & "','" & kodegdg & "','" & kodegdg2 & "','" & val(txtjum.Text) & "')"
If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
Else
jual.Execute "delete from mutasigudang where id='" & lv2.SelectedItem.SubItems(1) & "'"
jual.Execute "insert into mutasigudang values('','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & txtkode.Text & "','" & kodegdg & "','" & kodegdg2 & "','" & val(txtjum.Text) & "')"

frame.Enabled = False
MsgBox "Data berhasil diubah", vbInformation, judul
End If
dbgridtrans
Edit = True
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
frame.Enabled = True
End If

End Sub

Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
Label1.Caption = ""

Tab1.Tab = 0
Command1.SetFocus

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
End Sub
Sub kosong()
txtjum.Text = ""
cmb1.Text = ""
cmb2.Text = ""
txtkode.Text = ""
tgl.Value = Now
End Sub




Private Sub Command1_Click()
Tab1.Tab = 1
txtcari.SetFocus
End Sub

Private Sub Form_Load()
Edit = True
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
cust1
cust2
tgl.Value = noow
Ketengah Me
dbgrid
dbgridtrans
Tab1.Tab = 0
Dim Arq As String
    Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

Arq = ReadINI(App.Path & "\Backup.ini", "Backup_Dir", "Backup_Dir")
End Sub
Private Sub awal()
frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True

End Sub
Sub dbgrid()
Dim l As ListItem
LV1.ListItems.Clear
Set RS = New Recordset
RS.Open "Select * from tblbarang order by deskripsi", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = IIf(IsNull(![deskripsi]) = True, "", ![deskripsi])
  l.SubItems(3) = IIf(IsNull(![satuan]) = True, "", ![satuan])
    l.SubItems(4) = IIf(IsNull(![stok]) = True, "", ![stok])

    .MoveNext
    Loop
End With


End Sub
Sub Dbgrid2()
Dim l As ListItem
LV1.ListItems.Clear

Set RS = New Recordset
RS.Open "Select * from tblbarang where deskripsi like '%" & txtcari.Text & "%' order by deskripsi", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = IIf(IsNull(![deskripsi]) = True, "", ![deskripsi])
  l.SubItems(3) = IIf(IsNull(![satuan]) = True, "", ![satuan])
    l.SubItems(4) = IIf(IsNull(![stok]) = True, "", ![stok])

    .MoveNext
    Loop
End With


End Sub
Sub dbgridtrans()
Dim l As ListItem
lv2.ListItems.Clear
Set RS = New Recordset
RS.Open "Select id,m.kode_brg,t.deskripsi,tanggal,m.dari,m.ke,m.jumlah from mutasigudang m,tblbarang t where m.kode_brg=t.kode_brg order by tanggal desc,deskripsi desc", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
         l.SubItems(1) = ![id]
        l.SubItems(2) = Format(![tanggal], "dd mmm yyyy")
        l.SubItems(3) = ![kode_brg]
  l.SubItems(4) = ![deskripsi]
    l.SubItems(5) = ![dari]
  l.SubItems(6) = ![ke]
  l.SubItems(7) = ![jumlah]

    .MoveNext
    Loop
End With


End Sub

Private Sub cust1()

cmb1.Clear
sql = "select nama_gudang from gudang where nama_gudang <>'' order by nama_gudang"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
cmb1.AddItem rsplg!nama_gudang
rsplg.MoveNext
 Loop
  rsplg.MoveFirst
  End If

rsplg.Close

  End Sub
Private Sub cust2()

cmb2.Clear

sql = "select nama_gudang from gudang where nama_gudang <>'' order by nama_gudang"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
cmb2.AddItem rsplg!nama_gudang
rsplg.MoveNext
 Loop
  rsplg.MoveFirst
  End If

rsplg.Close

  End Sub

Private Sub LV1_DblClick()
If LV1.ListItems.count = 0 Then Exit Sub
txtkode.Text = LV1.SelectedItem.SubItems(1)
Tab1.Tab = 0
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LV1_DblClick
End If

End Sub

Private Sub LV2_Click()
If lv2.ListItems.count = 0 Then Exit Sub
teks
Cmdedit.Enabled = True
Cmdhapus.Enabled = True

End Sub

Private Sub lv2_DblClick()
If lv2.ListItems.count = 0 Then Exit Sub
Tab1.Tab = 0
teks
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub

Private Sub lv2_KeyPress(KeyAscii As Integer)
If LV1.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
lv2_DblClick
End If
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
txtcari.SetFocus
Cmdcari.Enabled = True
End If
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If LV1.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
LV1.SetFocus
End If
End Sub

Private Sub txtjum_Change()
If Edit = False Then
If val(txtjum.Text) > val(lblstok1.Caption) Then
MsgBox "Stok kurang", vbCritical, judul
txtjum.Text = "0"
Exit Sub
End If
End If
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmb1.SetFocus
End If
End Sub
Private Sub cmb1KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmb2.SetFocus
End If
End Sub
Private Sub cmb2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
txtjum.SetFocus
End If

End Sub

Private Sub txtjum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 13 Then
Cmdsimpan.SetFocus
End If

End Sub

Private Sub txtcari_Change()
Dbgrid2
End Sub
