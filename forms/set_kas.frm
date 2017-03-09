VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form set_kas 
   Caption         =   "Setting kas"
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
      TabIndex        =   3
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
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
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   11
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
         Caption         =   "Data Supplier"
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
         TabIndex        =   12
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "set_kas.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Kas"
      TabPicture(1)   =   "set_kas.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dbgrid1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3375
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BackColor       =   13750737
         HeadLines       =   1
         RowHeight       =   15
         WrapCellPointer =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
      Begin apotekbaleendah.xFrame frame 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
         BorderColor     =   0
         ButtonColor     =   0
         ButtonHighlightColor=   0
         ColorScheme     =   0
         Caption         =   ""
         Enabled         =   -1  'True
         Expanded        =   0   'False
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
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   5
            Top             =   600
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
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   1200
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
            Caption         =   "KAS"
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
            TabIndex        =   15
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Deskripsi"
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
            Top             =   1200
            Width           =   1815
         End
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   3255
      End
   End
End
Attribute VB_Name = "set_kas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdbatal_Click()
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub Cmdcari_Click()
Dim kata As String

Frame.Enabled = False
kata = InputBox("Masukkan nomor akun atau keterangan akun", "Cari...")
If StrPtr(kata) = 0 Then Exit Sub

If kata = "" Then Exit Sub
cmdsimpan.Enabled = False
sql = "select* from tblakunkas where kodeakun='" & kata & "' or namaakun like '%" & kata & "%' "
Set rssupp = New Recordset
Set rssupp = jual.Execute(sql)
If Not rssupp.EOF Then
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
teks
dbgrid
Tab1.Tab = 0
Else
MsgBox "Tidak ada", vbOKOnly, "Penjualan"
dbgrid
awal
End If

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub

Private Sub dbgrid1_GotFocus()
Text1.Text = dbgrid1.Columns.Item(0)

Cmdhapus.Enabled = True
End Sub

Sub teks()

cmdtambah.Enabled = True
Text1.Text = dbgrid1.Columns.Item(0)

Text2.Text = dbgrid1.Columns.Item(1)

End Sub
Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
Frame.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True
End If
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from tblakunkas where kodeakun='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
Set rssupp = New Recordset
A = Text1.Text
rssupp.Open "select * from tblakunkas where kodeakun='" & A & "'", jual, adOpenStatic, adLockOptimistic
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
Frame.Enabled = True
End If

End Sub
Sub ubah()
rssupp.Fields(0) = Text1.Text
rssupp.Fields(1) = Text2.Text

rssupp.Update
Frame.Enabled = False


End Sub

Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
Label1.Caption = ""
Tab1.Tab = 0
Text1.SetFocus

End Sub
Sub tambah()
Edit = False
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Frame.Enabled = True
cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
dbgrid1.Enabled = False
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rssupp = New Recordset
sql = "Select KODEakun from tblakunkas order by kodeakun Desc"
Set rssupp = jual.Execute(sql)
If rssupp.EOF = True Then
Text1.Text = "SP0001"
Else
j = val(Right(rssupp(0), 4))
No = "SP" + Format(Str(j + 1), "0000")
Text1.Text = No
End If
End Sub

Private Sub dbgrid1_DblClick()
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

End Sub
Private Sub awal()
Frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True
dbgrid1.Enabled = True

End Sub
Sub dbgrid()

sql = "select * from tblakunkaS "
Set dbgrid1.DataSource = jual.Execute(sql)

End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
cmdsimpan.Enabled = False
Cmdcari.Enabled = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan_Click
End If
End Sub
