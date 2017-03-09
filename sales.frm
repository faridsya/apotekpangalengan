VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmsls 
   Caption         =   "Sales"
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
      TabIndex        =   5
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
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
      TabIndex        =   4
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
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
      TabIndex        =   6
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
      TabIndex        =   9
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   5160
         OleObjectBlob   =   "sales.frx":0000
         Top             =   240
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Sales"
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
         TabIndex        =   10
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "sales.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data sales"
      TabPicture(1)   =   "sales.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrid1"
      Tab(1).Control(1)=   "Label1"
      Tab(1).ControlCount=   2
      Begin penjualan.xFrame frame 
         Height          =   3615
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
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
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   2040
            TabIndex        =   16
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
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
            Locked          =   -1  'True
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   2040
            TabIndex        =   15
            Top             =   960
            Width           =   3495
            _ExtentX        =   6165
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
         Begin XPControls.XPText Text3 
            Height          =   285
            Left            =   2040
            TabIndex        =   14
            Top             =   1440
            Width           =   3495
            _ExtentX        =   6165
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
         Begin XPControls.XPText Text4 
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   1920
            Width           =   3495
            _ExtentX        =   6165
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
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID Sales"
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
            TabIndex        =   20
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama sales"
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
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Telepon"
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
            Index           =   7
            Left            =   360
            TabIndex        =   18
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            Height          =   495
            Index           =   8
            Left            =   360
            TabIndex        =   17
            Top             =   1440
            Width           =   1815
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3375
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
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
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   11
         Top             =   3840
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmsls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdbatal_Click()
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub cmdcari_Click()
frame.Enabled = False
kata = InputBox("Masukkan id sales atau nama sales", "Cari...")
If kata = "" Then Exit Sub
Cmdsimpan.Enabled = False
sql = "select* from sales where id_sales='" & kata & "' or nama_sales like '%" & kata & "%' "
Set rssupp = New Recordset
Set rssupp = jual.Execute(sql)
If Not rssupp.EOF Then
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
Set dbgrid1.DataSource = rssupp
Tab1.Tab = 1
Else
MsgBox "Tidak ada", vbOKOnly, "Penjualan"
dbgrid
awal
End If

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub

Sub teks()

cmdtambah.Enabled = True
Text1.Text = rssupp.Fields(0)
Text2.Text = rssupp.Fields(1)
Text3.Text = rssupp.Fields(2)
Text4.Text = rssupp.Fields(3)
End Sub
Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
frame.Enabled = True

Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End If
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from sales where id_sales='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
Set rssupp = New Recordset
a = Text1.Text
rssupp.Open "select * from sales where id_sales='" & a & "'", jual, adOpenStatic, adLockOptimistic
If Edit = False Then

rssupp.AddNew
ubah
rssupp.Close
If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
Else
ubah
End If
dbgrid
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
frame.Enabled = True
End If

End Sub
Sub ubah()
rssupp.Fields(0) = Text1.Text
rssupp.Fields(1) = Text2.Text
rssupp.Fields(2) = Text3.Text
rssupp.Fields(3) = Text4.Text

rssupp.Update
frame.Enabled = False


End Sub

Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
Label1.Caption = ""
idoto
Tab1.Tab = 0
Text2.SetFocus

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
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
a = "Sls" & "%"
sql = "Select id_sales from sales where id_sales like '" & a & "' order by id_sales Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
Text1.Text = "Sls001"
Else
j = val(Right(rsplg(0), 3))
No = "Sls" + Format(Str(j + 1), "000")
Text1.Text = No
End If
End Sub

Private Sub dbgrid1_DblClick()
On Error Resume Next
teks
Tab1.Tab = 0
awal
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub



Private Sub Form_Load()


awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"

Ketengah Me
dbgrid
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
dbgrid1.Enabled = True

End Sub
Sub dbgrid()
Set rssupp = New Recordset
sql = "select * from sales"
Set rssupp = jual.Execute(sql)

Set dbgrid1.DataSource = rssupp


End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
Cmdsimpan.Enabled = False
Cmdcari.Enabled = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan_Click
End If

End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
End Sub
