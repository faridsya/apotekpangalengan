VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kurs 
   Caption         =   "Kurs"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6930
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcari 
      Caption         =   "&Cari"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   6360
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton Cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   6360
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   10
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6120
         OleObjectBlob   =   "kurs.frx":0000
         Top             =   120
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kurs rupiah terhadap US dollar"
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
         Left            =   240
         TabIndex        =   11
         Top             =   120
         Width           =   6135
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "kurs.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data "
      TabPicture(1)   =   "kurs.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrid1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5953
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
      Begin penjualan.xFrame frame 
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6495
         _extentx        =   11456
         _extenty        =   8493
         colorscheme     =   0
         bordercolor     =   0
         buttoncolor     =   0
         buttonhighlightcolor=   0
         colorscheme     =   0
         caption         =   ""
         enabled         =   -1  'True
         expanded        =   0   'False
         font            =   "kurs.frx":026C
         fontbold        =   0   'False
         fontitalic      =   0   'False
         fontsize        =   8.25
         fontstrikethru  =   0   'False
         fontunderline   =   0   'False
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   3360
            OleObjectBlob   =   "kurs.frx":0298
            TabIndex        =   17
            Top             =   1560
            Width           =   1455
         End
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   1320
            TabIndex        =   1
            Top             =   1560
            Width           =   1935
            _ExtentX        =   3413
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
            Height          =   255
            Left            =   1320
            TabIndex        =   16
            Top             =   960
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
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
            Format          =   66322435
            CurrentDate     =   37623
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
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nilai 1 US$ ="
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
            Index           =   9
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   15
         Top             =   3960
         Width           =   3255
      End
   End
End
Attribute VB_Name = "kurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode As String
Private Sub cmdbatal_Click()
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub cmdcari_Click()
Dim kata As String

If Cmdcari.Caption = "&Cari" Then
Cmdcari.Caption = "&Refresh"
Frame.Enabled = False
kata = InputBox("Masukkan kode bahan baku atau nama bahan baku", "Cari...")
If StrPtr(kata) = 0 Then Exit Sub

If kata = "" Then Exit Sub
Cmdsimpan.Enabled = False

sql = "select* from tblbarang2 where kode_brg='" & kata & "' or deskripsi like '%" & kata & "%' "
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


Sub teks()
On Error Resume Next
tgl.Value = dbgrid1.Columns(0).Text
Text1.Text = dbgrid1.Columns(1).Text


End Sub
Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
Frame.Enabled = True

Cmdsimpan.Enabled = True
cmdbatal.Enabled = True
End If
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from kurs where tanggal='" & tgl.Value & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
If MsgBox("Simpan data?", vbYesNo) = vbYes Then
Set rsplg = New Recordset
rsplg.Open "select * from kurs'", jual, adOpenStatic, adLockOptimistic
If Edit = False Then

rsplg.AddNew
ubah
rsplg.Close
Else
ubah
End If
Set rsmt = New Recordset
rsmt.Open "SELECT * from kurs order by tanggal", jual, adOpenStatic, adLockOptimistic
rsmt.MoveLast
matu = rsmt!nilai

dbgrid
End If
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
Frame.Enabled = True
End If

End Sub
Sub ubah()
rsplg.Fields(0) = Format(tgl.Value, "dd MMM YYYY hh:mm")
rsplg.Fields(1) = Text1.Text

rsplg.Update
Frame.Enabled = False


End Sub

Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
Label1.Visible = False
Tab1.Tab = 0
Text1.SetFocus

End Sub

Sub tambah()
Edit = False
dbgrid1.Enabled = True

Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Frame.Enabled = True
cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
End Sub
Sub kosong()
Text1.Text = ""
End Sub
Sub idoto()
Dim j As Integer
Dim no As String
Set rsplg = New Recordset
sql = "Select id_pelanggan from pelanggan order by id_pelanggan Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
Text1.Text = "CUS0001"
Else
j = val(Right(rsplg(0), 4))
no = "CUS" + Format(Str(j + 1), "0000")
Text1.Text = no
End If
End Sub

Private Sub dbgrid1_DblClick()
On Error Resume Next
teks
Tab1.Tab = 0
End Sub



Private Sub Form_Load()
 Set jual = New adodb.Connection
        jual.CursorLocation = adUseClient
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/Penjualan.mdb;Jet OLEDB:Database Password=tujuh;"

awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
tgl.Format = dtpCustom
tgl.CustomFormat = "dd MMM yyyy HH:mm"
tgl.Value = Now
Ketengah Me
dbgrid
Tab1.Tab = 0
Dim Arq As String
    SkinPath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin SkinPath
    Skin1.ApplySkin Me.hWnd

Arq = ReadINI(App.Path & "\Backup.ini", "Backup_Dir", "Backup_Dir")
RekamKegiatan ("Masuk form Kurs")
End Sub
Private Sub awal()
Frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True
dbgrid1.Enabled = False
Label1.Caption = "Dobel klik untuh hapus atau edit"
End Sub
Sub dbgrid()

sql = "select * from kurs order by tanggal"
Set rsplg = jual.Execute(sql)

Set dbgrid1.DataSource = rsplg


End Sub


Private Sub text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub
