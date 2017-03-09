VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form suppl 
   Caption         =   "Supplier"
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
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
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
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   14
      Top             =   0
      Width           =   8175
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "supp.frx":0000
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
         TabIndex        =   15
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "supp.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data supplier"
      TabPicture(1)   =   "supp.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Skin1"
      Tab(1).Control(1)=   "dbgrid1"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -74520
         OleObjectBlob   =   "supp.frx":0902
         Top             =   3240
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3375
         Left            =   -75000
         TabIndex        =   13
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
      Begin apotekbaleendah.xFrame frame 
         Height          =   3855
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5953
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
         Begin XPControls.XPText text6 
            Height          =   285
            Left            =   1920
            TabIndex        =   23
            Top             =   3000
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
            Locked          =   -1  'True
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   6
            Top             =   1080
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
            Left            =   1920
            TabIndex        =   7
            Top             =   1560
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
            Left            =   1920
            TabIndex        =   8
            Top             =   2040
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
         Begin XPControls.XPText Text5 
            Height          =   285
            Left            =   1920
            TabIndex        =   9
            Top             =   2520
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
            Caption         =   "Jumlah hutang(Rp)"
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
            Left            =   120
            TabIndex        =   24
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID Supplier"
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
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Left            =   120
            TabIndex        =   20
            Top             =   1080
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
            Left            =   120
            TabIndex        =   19
            Top             =   2520
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
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person:"
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
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Width           =   2415
         End
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   3840
         Width           =   3255
      End
   End
End
Attribute VB_Name = "suppl"
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

If Cmdcari.Caption = "&Cari" Then
Cmdcari.Caption = "&Refresh"

Frame.Enabled = False
kata = InputBox("Masukkan id supplier atau nama suplier", "Cari...")
If kata = "" Then Exit Sub
cmdsimpan.Enabled = False
sql = "select* from tblsupplier where id_supplier='" & kata & "' or supplier like '%" & kata & "%' "
Set rssupp = New Recordset
Set rssupp = jual.Execute(sql)
If Not rssupp.EOF Then
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
Set dbgrid1.DataSource = rssupp
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

cmdtambah.Enabled = True
Text1.Text = rssupp.Fields(0)
Text2.Text = rssupp.Fields(1)
text3.Text = rssupp.Fields(2)
text4.Text = rssupp.Fields(3)
text5.Text = rssupp.Fields(4)
text6.Text = rssupp.Fields(5)

End Sub
Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
Frame.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True
Text1.Enabled = False
End If
End Sub

Private Sub cmdhapus_Click()
Set rsplg = New Recordset
rsplg.Open "select id_supplier from pembelian where id_supplier='" & Text1.Text & "' limit 1", jual, adOpenStatic, adLockOptimistic
If Not rsplg.EOF Then
MsgBox "Tidak dapat dihapus,Supplier ini sudah melakukan transaksi!", vbCritical, judul
Exit Sub
End If

If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from tblsupplier where id_supplier='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
awal
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
If Text2.Text = "" Then
MsgBox "Nama tidak boleh kosong", vbInformation, judul
Text2.SetFocus
Exit Sub
End If

Set RS = New Recordset
RS.Open "select * from tblsupplier where supplier='" & Text2.Text & "' and id_supplier <>'" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama sudah ada,bedakan walau 1 huruf", vbInformation, judul
Text2.SetFocus
Exit Sub
End If

Set rssupp = New Recordset
A = Text1.Text
rssupp.Open "select * from tblsupplier where id_supplier='" & A & "'", jual, adOpenStatic, adLockOptimistic
If Edit = False Then

jual.Execute "insert into tblsupplier values('" & Text1.Text & "','" & Text2.Text & "','" & text3.Text & "','" & text4.Text & "','" & text5.Text & "',0)"
Frame.Enabled = False

If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Exit Sub
Else
awal
End If
Else
jual.Execute "update tblsupplier set supplier='" & Text2.Text & "',kontak_person='" & text3.Text & "',alamat='" & text4.Text & "',no_telp='" & text5.Text & "' where id_supplier='" & Text1.Text & "'"
Frame.Enabled = False

End If
dbgrid

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
Frame.Enabled = True
End If

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
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Frame.Enabled = True
cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
dbgrid1.Enabled = False
Text1.Enabled = True
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""

End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rssupp = New Recordset
A = hsup & "%"
sql = "Select id_supplier from tblsupplier where id_supplier like '" & A & "' order by id_supplier Desc"
Set rssupp = jual.Execute(sql)
If rssupp.EOF = True Then
Text1.Text = hsup + "0001"
Else
j = val(Right(rssupp(0), 4))
No = hsup + Format(Str(j + 1), "0000")
Text1.Text = No
End If
End Sub

Private Sub Command1_Click()
faktur.Show
End Sub

Private Sub dbgrid1_DblClick()
On Error Resume Next

teks
Tab1.Tab = 0
awal
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub



Private Sub Form_Activate()
If hsup = "" Then
hsup = "S-"
End If

End Sub

Private Sub Form_Load()
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
If hsup = "" Then
hsup = "S-"
End If

Ketengah Me
dbgrid
Tab1.Tab = 0
Dim Arq As String
    Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

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

sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set dbgrid1.DataSource = rssupp


End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
cmdsimpan.Enabled = False
Cmdcari.Enabled = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4.SetFocus
End If
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text5.SetFocus
End If
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan_Click
End If
End Sub
