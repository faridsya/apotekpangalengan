VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form pelanggan 
   Caption         =   "Pelanggan"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6930
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   6480
      OleObjectBlob   =   "pelanggan.frx":0000
      Top             =   7440
   End
   Begin VB.CommandButton Cmdcari 
      Caption         =   "&Cari"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   4680
      TabIndex        =   11
      Top             =   7320
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   17
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6240
         OleObjectBlob   =   "pelanggan.frx":0234
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
         Caption         =   "Data Pelanggan"
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
         TabIndex        =   18
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11456
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "pelanggan.frx":0468
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data pelanggan"
      TabPicture(1)   =   "pelanggan.frx":0484
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "dbgrid1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Data sales"
      TabPicture(2)   =   "pelanggan.frx":04A0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvplg"
      Tab(2).Control(1)=   "SkinLabel21"
      Tab(2).Control(2)=   "txtcrp"
      Tab(2).ControlCount=   3
      Begin apotekbaleendah.xFrame frame 
         Height          =   5775
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   10186
         Caption         =   "xFrame1"
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
         Begin XPControls.XPText Text6 
            Height          =   285
            Left            =   2160
            TabIndex        =   34
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
         End
         Begin XPControls.XPText Text4 
            Height          =   285
            Left            =   2160
            TabIndex        =   3
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
         Begin XPControls.XPText Text3 
            Height          =   285
            Left            =   2160
            TabIndex        =   2
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
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   2160
            TabIndex        =   1
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
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   2160
            TabIndex        =   33
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
         Begin XPControls.XPText text5 
            Height          =   285
            Left            =   2160
            TabIndex        =   4
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
         Begin XPControls.XPText text7 
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Top             =   3480
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
         Begin XPControls.XPText text8 
            Height          =   285
            Left            =   2160
            TabIndex        =   7
            Top             =   3960
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
         Begin XPControls.XPText text9 
            Height          =   285
            Left            =   2160
            TabIndex        =   32
            Top             =   4440
            Width           =   2295
            _ExtentX        =   4048
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
         Begin VB.CommandButton Command1 
            Caption         =   "Cari"
            Height          =   375
            Left            =   4680
            TabIndex        =   8
            Top             =   4440
            Width           =   975
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
            Left            =   240
            TabIndex        =   31
            Top             =   1560
            Width           =   2415
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
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   30
            Top             =   2040
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
            Left            =   240
            TabIndex        =   29
            Top             =   3000
            Width           =   1335
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
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID pelanggan"
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
            TabIndex        =   27
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
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
            Left            =   240
            TabIndex        =   26
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
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
            TabIndex        =   25
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
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
            Index           =   4
            Left            =   240
            TabIndex        =   24
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID sales"
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
            Index           =   10
            Left            =   240
            TabIndex        =   23
            Top             =   4440
            Width           =   1335
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3495
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6165
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
      Begin MSComctlLib.ListView lvplg 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   20
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8493
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id sales"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama "
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Alamat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "No_telp"
            Object.Width           =   3528
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   375
         Left            =   -74760
         OleObjectBlob   =   "pelanggan.frx":04BC
         TabIndex        =   21
         Top             =   5520
         Width           =   2295
      End
      Begin XPControls.XPText txtcrp 
         Height          =   375
         Left            =   -72240
         TabIndex        =   22
         Top             =   5520
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   3960
         Width           =   3255
      End
   End
End
Attribute VB_Name = "pelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub dbgridplg()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from sales order by nama_sales"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_sales]
        l.SubItems(2) = ![nama_sales]
                                l.SubItems(3) = ![alamat_sales]

                l.SubItems(4) = ![telp_sales]
                
                
    .MoveNext
    Loop
End With


End Sub
Sub dbgridplg2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from sales where id_sales like '" & txtcrp.Text & "%' or nama_sales like '%" & txtcrp.Text & "%' order by nama_sales"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_sales]
        l.SubItems(2) = ![nama_sales]
                                l.SubItems(3) = ![alamat_sales]

                l.SubItems(4) = ![telp_sales]
    .MoveNext
    Loop
End With


End Sub

Private Sub Command1_Click()
Tab1.Tab = 2
End Sub

Private Sub lvplg_DblClick()
text9.Text = lvplg.SelectedItem.SubItems(1)
Tab1.Tab = 0
End Sub

Private Sub txtcrp_KeyPress(KeyAscii As Integer)
If lvplg.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
lvplg.SetFocus
End If
End Sub
Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub

Private Sub txtcrp_Change()
dbgridplg2

End Sub

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
frame.Enabled = False
kata = InputBox("Masukkan id pelanggan atau nama pelanggan", "Cari...")
If StrPtr(kata) = 0 Then Exit Sub

If kata = "" Then Exit Sub
Cmdsimpan.Enabled = False

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


Sub teks()
On Error Resume Next
cmdtambah.Enabled = True
Text1.Text = rsplg.Fields(0)
Text2.Text = rsplg.Fields(1)
text3.Text = rsplg.Fields(2)
text4.Text = rsplg.Fields(3)
Text5.Text = rsplg.Fields(4)
text6.Text = rsplg.Fields(5)
text7.Text = rsplg.Fields(6)
text8.Text = rsplg.Fields(7)
text9.Text = rsplg.Fields(9)

End Sub
Private Sub cmdedit_Click()
If Tab1.Tab = 0 Then
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
frame.Enabled = True
Text1.Enabled = False
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True


End If
End Sub

Private Sub cmdhapus_Click()
Set rsplg = New Recordset
rsplg.Open "select id_pelanggan from penjualan where id_pelanggan='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rsplg.EOF Then
MsgBox "Tidak dapat dihapus,sudah melakukan transaksi!"
Exit Sub
End If
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from pelanggan where id_pelanggan='" & Text1.Text & "'"
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
RS.Open "select * from pelanggan where nama='" & Text2.Text & "' and id_pelanggan<>'" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama sudah ada,bedakan walau 1 huruf", vbInformation, judul
Text2.SetFocus
Exit Sub
End If


Set rsplg = New Recordset
A = Text1.Text
rsplg.Open "select * from pelanggan where id_pelanggan='" & A & "'", jual, adOpenStatic, adLockOptimistic
If Edit = False Then
jual.Execute "insert into pelanggan values('" & Text1.Text & "','" & Text2.Text & "','" & text3.Text & "','" & text4.Text & "','" & Text5.Text & "','" & text6.Text & "','" & text7.Text & "','" & text8.Text & "',0,'" & text9.Text & "')"
frame.Enabled = False

If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
Else
jual.Execute "update pelanggan set nama='" & Text2.Text & "',kontak_person='" & text3.Text & "',alamat='" & text4.Text & "',kota='" & Text5.Text & "',telepon='" & text6.Text & "',fax='" & text7.Text & "',email='" & text8.Text & "',id_sales='" & text9.Text & "' where id_pelanggan='" & Text1.Text & "'"
frame.Enabled = False

End If
dbgrid
Exit Sub
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
Text1.Enabled = True
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
text3.Text = ""
text4.Text = ""
Text5.Text = ""
text6.Text = ""
text7.Text = ""
text8.Text = ""
text9.Text = ""

End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
A = hcus & "%"
sql = "Select id_pelanggan from pelanggan where id_pelanggan like '" & A & "' order by id_pelanggan Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
Text1.Text = hcus + "0001"
Else
j = val(Right(rsplg(0), 4))
No = hcus + Format(Str(j + 1), "0000")
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



Private Sub Form_Activate()
If hcus = "" Then
hcus = "CUS"
End If

End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pelanggan.doc" _
                 , vbNullString, vbNullString, 1
End If

End Sub

Private Sub Form_Load()

 dbgridplg
awal
Label1.Caption = "*Dobel klik untuk edit atau hapus"
If hcus = "" Then
hcus = "CUS"
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

sql = "select * from pelanggan order by id_pelanggan"
Set rsplg = jual.Execute(sql)

Set dbgrid1.DataSource = rsplg


End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
Cmdsimpan.Enabled = False
Cmdcari.Enabled = True
w = Screen.Width
m = Me.Width
Me.Width = Screen.Width - 4000
Tab1.Width = w - 4550
dbgrid1.Width = w - 5000

Ketengah Me
Else
Me.Width = 7100
Tab1.Width = 6730
dbgrid1.Width = 6495
If Tab1.Tab = 2 Then
txtcrp.SetFocus
End If

Ketengah Me

End If
End Sub



'Private Sub text1_Change()
'Dim jum1, jum2 As Currency
'Set RS = New Recordset
'RS.Open "Select sum(jumlah_piutang) as jp from piutang where id_pelanggan='" & text1.Text & "'", jual, adOpenStatic, adLockOptimistic
'jum1 = RS!jp
'RS.Close
'RS.Open "Select sum(jumlah_byr) as jp from piutang where id_pelanggan='" & text1.Text & "'", jual, adOpenStatic, adLockOptimistic
'jum2 = RS!jp
'RS.Close
'txtp.Text = jum1 - jum2
'End Sub

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
Text5.SetFocus
End If
End Sub
Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text6.SetFocus
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
Cmdsimpan.SetFocus
End If
End Sub

Private Sub text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan.SetFocus
End If
End Sub
