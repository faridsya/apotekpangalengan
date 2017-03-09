VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdo 
   Caption         =   "Delivery Order"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   9120
      OleObjectBlob   =   "frmdo.frx":0000
      Top             =   7080
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Deliveri order"
      TabPicture(0)   =   "frmdo.frx":0234
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ket"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SkinLabel18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LV1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "XPCheck1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "XPCheck2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CrystalReport1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TCmdhapus"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdkeluar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdsimpan"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdbatal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdtambah"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Data obat"
      TabPicture(1)   =   "frmdo.frx":0250
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SkinLabel11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvbrg"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtcari"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Data Pelanggan"
      TabPicture(2)   =   "frmdo.frx":026C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtcrp"
      Tab(2).Control(1)=   "lvplg"
      Tab(2).Control(2)=   "SkinLabel9"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Data DO"
      TabPicture(3)   =   "frmdo.frx":0288
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text2"
      Tab(3).Control(1)=   "cmdhps"
      Tab(3).Control(2)=   "Command5"
      Tab(3).Control(3)=   "ListView20"
      Tab(3).Control(4)=   "ListView10"
      Tab(3).Control(5)=   "SkinLabel4"
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtcrp 
         Height          =   405
         Left            =   -72000
         TabIndex        =   15
         Top             =   6360
         Width           =   2895
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   6240
         Width           =   3255
      End
      Begin VB.Frame Frame1 
         Height          =   1575
         Left            =   -69480
         TabIndex        =   14
         Top             =   4800
         Width           =   3135
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmdo.frx":02A4
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmdo.frx":0314
            TabIndex        =   45
            Top             =   720
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmdo.frx":0386
            TabIndex        =   46
            Top             =   1200
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel hj 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmdo.frx":03F8
            TabIndex        =   47
            Top             =   1200
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel hb 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmdo.frx":0456
            TabIndex        =   48
            Top             =   720
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmdo.frx":04B4
            TabIndex        =   49
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   255
         Left            =   -73080
         TabIndex        =   17
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   255
         Left            =   -71640
         TabIndex        =   18
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdkeluar 
         Caption         =   "&Keluar"
         Height          =   255
         Left            =   -70320
         TabIndex        =   19
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton TCmdhapus 
         Caption         =   "&Hapus"
         Height          =   255
         Left            =   -66840
         TabIndex        =   20
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   -67680
         TabIndex        =   36
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CommandButton cmdhps 
         Caption         =   "&Hapus"
         Height          =   300
         Left            =   -74760
         TabIndex        =   33
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cetak &DO"
         Height          =   300
         Left            =   -73200
         TabIndex        =   32
         Top             =   3360
         Width           =   1695
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   -73440
         Top             =   6600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin XPControls.XPCheck XPCheck2 
         Height          =   195
         Left            =   -69000
         TabIndex        =   27
         Top             =   4560
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   344
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
      Begin XPControls.XPCheck XPCheck1 
         Height          =   255
         Left            =   -68640
         TabIndex        =   26
         Top             =   6480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Caption         =   "Tampilkan keterangan"
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
      Begin MSDataGridLib.DataGrid Dbgrid2 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   6
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   9128
         _Version        =   393216
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
      Begin VB.ComboBox kolom2 
         Height          =   315
         Left            =   -72000
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   6120
         Width           =   2055
      End
      Begin XPControls.XPCombo kolom 
         Height          =   315
         Left            =   -74880
         TabIndex        =   8
         Top             =   6120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         Text            =   "XPCombo1"
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
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   10
         Top             =   4800
         Width           =   4815
         Begin VB.CommandButton Command2 
            Caption         =   "&Cari"
            Height          =   255
            Left            =   3960
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin XPControls.XPText text4 
            Height          =   285
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
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
            Left            =   1680
            TabIndex        =   2
            Top             =   960
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
         Begin XPControls.XPText text5 
            Height          =   285
            Left            =   5040
            TabIndex        =   3
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
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
         Begin XPControls.XPCombo satuan 
            Height          =   315
            Left            =   1680
            TabIndex        =   29
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmdo.frx":0512
            TabIndex        =   41
            Top             =   600
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmdo.frx":057C
            TabIndex        =   42
            Top             =   960
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmdo.frx":05E6
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel stn 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmdo.frx":0656
            TabIndex        =   54
            Top             =   960
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
      Begin VB.Frame Frame 
         Height          =   1575
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton Command1 
            Caption         =   "&Cari"
            Height          =   255
            Left            =   8040
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin XPControls.XPText text3 
            Height          =   285
            Left            =   5640
            TabIndex        =   25
            Top             =   360
            Width           =   2055
            _ExtentX        =   3625
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
            TabIndex        =   1
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   157941763
            CurrentDate     =   40299
         End
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   0
            Top             =   240
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmdo.frx":06B4
            TabIndex        =   38
            Top             =   360
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmdo.frx":071C
            TabIndex        =   39
            Top             =   840
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmdo.frx":078E
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   11
         Top             =   2160
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4048
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Kode barang"
            Object.Width           =   2577
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama barang"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   34
         Top             =   3840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5106
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Kode barang"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Jumlah beli"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   35
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5318
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
            Text            =   "No pemesanan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID Supplier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Supplier"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView20 
         Height          =   2895
         Left            =   -74760
         TabIndex        =   28
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5106
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
            Text            =   "No DO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID Pelaanggan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Pelanggan"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView ListView10 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   30
         Top             =   3720
         Width           =   9975
         _ExtentX        =   17595
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Kode barang"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jumlah terjual"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvbrg 
         Height          =   5535
         Left            =   360
         TabIndex        =   31
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   9763
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kategori"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Stok"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Harga Jual"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvplg 
         Height          =   5775
         Left            =   -74760
         TabIndex        =   37
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   10186
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
            Text            =   "Id pelanggan"
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
            Text            =   "No.Telp"
            Object.Width           =   3528
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   -68640
         OleObjectBlob   =   "frmdo.frx":0804
         TabIndex        =   50
         Top             =   4560
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Left            =   360
         OleObjectBlob   =   "frmdo.frx":0890
         TabIndex        =   51
         Top             =   6240
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   -74520
         OleObjectBlob   =   "frmdo.frx":0904
         TabIndex        =   52
         Top             =   6360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   -70920
         OleObjectBlob   =   "frmdo.frx":097E
         TabIndex        =   53
         Top             =   3360
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket 
         Height          =   255
         Left            =   -74640
         OleObjectBlob   =   "frmdo.frx":0A1E
         TabIndex        =   55
         Top             =   6480
         Width           =   5895
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   -69360
         TabIndex        =   12
         Top             =   6120
         Width           =   2775
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -69240
         TabIndex        =   13
         Top             =   6120
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim kode, nbrg, stri As String
Dim st, cosbli, cosju As Currency



Sub dbgridpo()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select o.nomor_do,o.tgl_do,o.id_pelanggan,nama from deliveriorder o join pelanggan p on o.id_pelanggan=p.id_pelanggan where nomor_do like '" & Text2.Text & "%' or nama like '%" & Text2.Text & "%' order by nomor_do desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView20.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView20.ListItems.Add(, , ListView20.ListItems.count + 1)
        l.SubItems(1) = ![nomor_do]
        l.SubItems(2) = Format(![tgl_do], "dd MMM yyyy")
        l.SubItems(3) = ![id_pelanggan]

                l.SubItems(4) = ![nama]

    .MoveNext
    Loop
End With

End Sub


Private Sub Cmdbatal_Click()

awal
kosong
cmdtambah.SetFocus
Label2.Caption = ""
Label1.Caption = ""
End Sub

Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
Frame.Enabled = True
Frame2.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True

End Sub



Private Sub cmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub cmdhps_Click()
If ListView10.ListItems.count = 0 Then Exit Sub

If MsgBox("Yakin akan menghapus no DO " & ListView20.SelectedItem.SubItems(1) & "", vbYesNo, judul) = vbNo Then Exit Sub

jual.Execute "delete from deliveriorder where nomor_do='" & ListView20.SelectedItem.SubItems(1) & "'"

dbgridpo
ListView10.ListItems.Clear

End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
'On Error GoTo erol
Set RS = New Recordset

If text3.Text = "" Then
MsgBox "Pelanggan belum dipilih"
text3.SetFocus
Exit Sub
 End If
If MsgBox("Simpan data?", vbYesNo, judul) = vbNo Then Exit Sub

simpandata
If XPCheck2.Value = Checked Then
If MsgBox("Cetak DO?", vbYesNo, judul) = vbYes Then
cetak

Else
awal
cmdtambah.Enabled = True
cmdtambah.Refresh
cmdtambah.SetFocus
End If
Else
cmdtambah.SetFocus
End If

Label2.Caption = ""
Label1.Caption = ""
dbgridpo
erol:
If err.Description <> vbNullString Then
MsgBox "Data belum lengkap", vbCritical, "Penjualan"
Frame.Enabled = True
End If

End Sub
Sub simpandata()
'On Error GoTo erol
Set RS = New Recordset
sql = "select id_sales from pelanggan where id_pelanggan='" & text3.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
id = RS!id_sales
RS.Close
Else
id = ""
End If
sql = "insert into deliveriorder values('" & Text1.Text & "','" & Format(DTPicker1, "YYYY-mm-dd") & "','" & id & "','" & text3.Text & "')"

jual.Execute (sql)

Set RS = New Recordset
For z = 1 To LV1.ListItems.count
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.ListItems(z).SubItems(2) & "' and satuan='" & LV1.ListItems(z).SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic

sql = "insert into deliveriorder_detil values(null,'" & Text1.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) * rsstn!konversi & "','" & LV1.ListItems(z).SubItems(5) & "','" & LV1.ListItems(z).SubItems(4) & "',0)"
jual.Execute (sql)
    Next z
      awal
 MsgBox "Data berhasil disimpan!", vbInformation, judul: Exit Sub

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "POS"
Frame.Enabled = True
End If


End Sub


Sub cetak()

With CrystalReport1
.Reset
  

  .ReportFileName = serperreport & "\do.rpt"
  .RetrieveDataFiles
  .WindowTitle = "PO"
  .SelectionFormula = "{deliveriorder.nomor_do}='" & Text1.Text & "'"
    .Formulas(0) = "almt='" & almt & "'"
    .Formulas(1) = "nama='" & nama_toko & "'"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
End With
'Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Tidak ada data"
End If

End Sub
Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
kosong2
GetNumber
text3.SetFocus


End Sub
Sub tambah()
Edit = False
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Label1.Caption = "*Dobel klik untuk mengirimkan data"

Label2.Caption = "*Dobel klik untuk mengirimkan data"
Frame.Enabled = True
cmdbatal.Enabled = True
LV1.Enabled = False
Frame2.Enabled = True

End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
Text1.Text = "DO-" + Format(Now, "YY-") + Format(Now, "MM-") + Format(Now, "DD-") + Format(Now, "Hmmss")

End Sub
Sub GetNumber()
On Error GoTo salah
Dim v As Integer
    Dim counter As String
    Dim Hitung As Integer
    Dim tgl As String
    A = "DO" & "%"

sql = "Select nomor_do from deliveriorder where nomor_do like '" & A & "' order by nomor_do"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "DO" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![nomor_do], 8) <> "DO" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "DO" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!nomor_do, 2)) + 1
               counter = "DO" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        Text1.Text = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description
End Sub


Private Sub Command1_Click()
Tab1.Tab = 2
End Sub

Private Sub Command2_Click()
Tab1.Tab = 1
End Sub

Private Sub Command5_Click()
Text1.Text = ListView20.SelectedItem.SubItems(1)
cetak
End Sub


Private Sub ListView20_Click()
'On Error Resume Next
Set rstrans = New Recordset


sql = "select d.kode_brg,deskripsi,d.jumlah_brg2,d.satuan,d.jumlah_jual from deliveriorder_detil d join tblbarang b on d.kode_brg=b.kode_brg and d.nomor_do='" & ListView20.SelectedItem.SubItems(1) & "'"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView10.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView10.ListItems.Add(, , ListView10.ListItems.count + 1)
           l.SubItems(1) = ListView10.ListItems.count & "."

        l.SubItems(2) = ![kode_brg]
        l.SubItems(3) = ![deskripsi]
        l.SubItems(5) = ![jumlah_brg2]
        l.SubItems(4) = ![satuan]
        l.SubItems(6) = ![jumlah_jual]
    .MoveNext
    Loop
End With
'Text1.Text = ListView20.SelectedItem.SubItems(1)
'Text1_KeyPress (13)

End Sub

Private Sub lvbrg_DblClick()
If lvbrg.ListItems.count = 0 Then Exit Sub
Tab1.Tab = 0
text4.Text = lvbrg.SelectedItem.SubItems(1)
text4_KeyPress (13)
End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub

Private Sub lvplg_DblClick()
If lvplg.ListItems.count = 0 Then Exit Sub
text3.Text = lvplg.SelectedItem.SubItems(1)
Tab1.Tab = 0
text4.SetFocus
End Sub

Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub

 Private Sub satuan_Change()
If satuan.Text = "" Then
stn.Caption = ""
Else
stn.Caption = satuan.Text
End If
End Sub

Private Sub satuan_Click()
text6.SetFocus
End Sub

Private Sub satuan_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub
 
Private Sub tampil_stn()
On Error Resume Next
satuan.Clear
sql = "select * from satuan where kode_brg='" & kode & "' order by satuan"
Set rsstn = jual.Execute(sql)
If Not rsbarang.EOF Then
rsstn.MoveFirst
 Do While Not rsstn.EOF
satuan.AddItem rsstn!satuan
rsstn.MoveNext
 Loop
  End If
rsstn.Close


  End Sub

Sub kosong()
Text1.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""

LV1.ListItems.Clear

End Sub
Sub kosong2()
text4.Text = ""
text5.Text = ""
text6.Text = ""
stok = ""
hb = ""
satuan.Text = ""
stn.Caption = ""
hj = ""
End Sub



Private Sub Form_Load()
Dim sql As String


  
XPCheck1.Value = GetSetting("mekarlaksana", "frmdo", "XPCheck1.value", Checked)
XPCheck2.Value = GetSetting("mekarlaksana", "frmdo", "XPCheck2.value", Checked)


Tab1.Tab = 0

Edit = True
dbgrid
dbgridpo
awal
Ketengah Me
 dbgrids
 DTPicker1 = Format(Now)
     Skinpath = App.Path & "\skin\green.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub



Sub dbgrid()
'On Error Resume Next

Set rstrans = New Recordset

stri = Replace(txtcari.Text, "'", "''")

sql = "select * from tblbarang where kode_brg like '" & stri & "%' or deskripsi like '%" & stri & "%' order by deskripsi"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvbrg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , lvbrg.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
        l.SubItems(3) = ![kategori]
        l.SubItems(4) = ![satuan]
        l.SubItems(5) = ![stok]
        l.SubItems(6) = Format(![harga_jual], "#,#")
 
    .MoveNext
    Loop
End With


End Sub
Sub dbgrids()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from pelanggan where id_pelanggan like '" & txtcrp.Text & "%' or nama like '%" & txtcrp.Text & "%' order by nama"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_pelanggan]
        l.SubItems(2) = ![nama]
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![Telepon]

    .MoveNext
    Loop
End With


End Sub

Private Sub awal()
Frame.Enabled = False
Frame2.Enabled = False
Edit = True
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
LV1.Enabled = True
End Sub

Private Sub hb_Click()
text5.Text = hb.Caption
text6.SetFocus
End Sub




Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
txtcari.SetFocus
Else
If Tab1.Tab = 2 Then
txtcrp.SetFocus
End If
End If
End Sub

Private Sub Text2_Change()
dbgridpo
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text3.SetFocus
End If
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sql = "select* from pemesanan where no_pemesanan='" & Text1.Text & "'"
If jual.Execute(sql).EOF Then
text3.SetFocus
Else
sql = "select pemesanan.tanggal_pesan,tblsupplier.supplier,detilpesan.kode_brg,tblbarang.deskripsi,detilpesan.satuan,detilpesan.jumlah_brg from detilpesan,pemesanan,tblbarang,tblsupplier where pemesanan.no_pemesanan=detilpesan.no_pemesanan and detilpesan.kode_brg=tblbarang.kode_brg and pemesanan.id_supplier=tblsupplier.id_supplier and pemesanan.no_pemesanan='" & Text1.Text & "'"
Set RS = New Recordset
Set RS = jual.Execute(sql)
DTPicker1.Value = RS![tanggal_pesan]
text3.Text = RS!Supplier
LV1.ListItems.Clear
With RS
.MoveFirst

    Do While Not .EOF
   Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
   l.SubItems(1) = LV1.ListItems.count & "."
    l.SubItems(2) = ![kode_brg]
    l.SubItems(3) = ![deskripsi]
    l.SubItems(4) = ![satuan]

l.SubItems(5) = ![jumlah_brg]
    
    
    .MoveNext
    Loop
End With

End If
End If
End Sub

Private Sub text3_Click()
text4.SetFocus
End Sub

Private Sub text3_GotFocus()
If XPCheck1.Value = Checked Then

ket.Caption = "Nama pelanggan bisa diambil dari tab data pelanggan kemudian dobel klik"
End If
End Sub

Private Sub Text3_LostFocus()
ket.Caption = ""
End Sub

Sub proses()
text4.Text = Replace(text4.Text, "'", "''")

sql = "select * from tblbarang where deskripsi='" & text4.Text & "' or kode_brg='" & text4.Text & "'"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)

stok = RS!stok
cosbli = RS.Fields("Harga_beli")
hb = Format(Str(cosbli), "#,##0") + "/" + rsbarang.Fields("satuan")
text5.Text = cosbli
cosju = RS!harga_jual
hj.Caption = Format(Str(cosju), "#,##0") + "/" + rsbarang!satuan
kode = RS!kode_brg
nbrg = RS!deskripsi
text6.SetFocus
End Sub



Private Sub text4_GotFocus()
If XPCheck1.Value = Checked Then

ket.Caption = "Nama barang bisa diambil dari tab data barang"
End If
End Sub


Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text <> "" Then
text4.Text = Replace(text4.Text, "'", "''")

sql = "select * from tblbarang where kode_brg='" & text4.Text & "'"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)

stok.Caption = rsbarang!stok
cosbli = rsbarang.Fields("Harga_beli")
hb = Format(Str(cosbli), "#,##0") + "/" + rsbarang.Fields("satuan")
text5.Text = cosbli
cosju = rsbarang!harga_jual
hj.Caption = Format(Str(cosju), "#,##0") + "/" + rsbarang!satuan

kode = rsbarang!kode_brg
nbrg = rsbarang!deskripsi
text6.SetFocus
tampil_stn
satuan.Text = rsbarang.Fields("satuan")
Else
cmdsimpan.SetFocus
End If
End If
End Sub

Private Sub Text4_LostFocus()
ket.Caption = ""
End Sub



Private Sub text6_GotFocus()
ket.Caption = "Masukkan jumlah barang yang dipesan"
End Sub

Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
If text6.Text = "" Then MsgBox "Jumlah harus diisi!", vbCritical, judul: Exit Sub
st = val(text5.Text) * val(text6.Text)
Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari
If cari Is Nothing Then
isigrid
kosong2
text4.SetFocus
Else
MsgBox "Barang sudah terdaftar!", vbCritical, judul
text4.Text = ""
text4.SetFocus

Exit Sub
End If
End If
End Sub


Sub isigrid()
    Set butir = LV1.ListItems.Add
    With butir
           .SubItems(1) = LV1.ListItems.count & "."

    .SubItems(2) = kode
    .SubItems(3) = nbrg
    .SubItems(4) = satuan.Text
    .SubItems(5) = text6.Text
    End With
  LV1.Enabled = True
   With butir
End With
End Sub



Private Sub Text6_LostFocus()
ket.Caption = ""
End Sub

Private Sub txtcari_Change()
dbgrid
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lvbrg.ListItems.count = 0 Then Exit Sub
lvbrg.SetFocus
End If
End Sub

Private Sub txtcrp_Change()
dbgrids
End Sub

Private Sub txtcrp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lvplg.ListItems.count = 0 Then Exit Sub
lvplg.SetFocus
End If
End Sub

Private Sub XPCheck1_Click()
    
    SaveSetting "mekarlaksana", "frmdo", "XPCheck1.value", XPCheck1.Value

End Sub



Private Sub XPCheck2_Click()
    SaveSetting "mekarlaksana", "frmdo", "XPCheck2.value", XPCheck2.Value

End Sub
