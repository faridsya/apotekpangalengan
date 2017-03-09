VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmreg 
   Caption         =   "Pendaftaran servis"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13573
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "frmreg.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(1)=   "cmdout"
      Tab(0).Control(2)=   "cmdbatal"
      Tab(0).Control(3)=   "cmdsimpan"
      Tab(0).Control(4)=   "cmdtambah"
      Tab(0).Control(5)=   "Skin1"
      Tab(0).Control(6)=   "CrystalReport1"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Data Servis"
      TabPicture(1)   =   "frmreg.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(2)=   "Baram"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "lvdtl2"
      Tab(1).Control(5)=   "lvdtl"
      Tab(1).Control(6)=   "lvpaket"
      Tab(1).Control(7)=   "txtcari"
      Tab(1).Control(8)=   "cmdhapus"
      Tab(1).Control(9)=   "cmdubah"
      Tab(1).Control(10)=   "cmbstts2"
      Tab(1).Control(11)=   "Command3"
      Tab(1).Control(12)=   "Command4"
      Tab(1).Control(13)=   "cmdhps"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Pelanggan Servis"
      TabPicture(2)   =   "frmreg.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lvplg"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtcari2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SkinLabel9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   375
         Left            =   840
         OleObjectBlob   =   "frmreg.frx":0054
         TabIndex        =   45
         Top             =   6840
         Width           =   1575
      End
      Begin VB.TextBox txtcari2 
         Height          =   375
         Left            =   2520
         TabIndex        =   44
         Top             =   6840
         Width           =   2895
      End
      Begin VB.CommandButton cmdhps 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   -61560
         TabIndex        =   42
         Top             =   4920
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Faktur Pengambilan"
         Height          =   375
         Left            =   -65040
         TabIndex        =   41
         Top             =   4920
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Faktur pendaftaran"
         Height          =   375
         Left            =   -67080
         TabIndex        =   40
         Top             =   4920
         Width           =   1695
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   -61200
         Top             =   7320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cmbstts2 
         Height          =   315
         Left            =   -73920
         TabIndex        =   34
         Top             =   480
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -63480
         OleObjectBlob   =   "frmreg.frx":00CE
         Top             =   7200
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "&Ubah"
         Height          =   375
         Left            =   -62880
         TabIndex        =   22
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -62040
         TabIndex        =   21
         Top             =   7200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   -74640
         TabIndex        =   0
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   -72840
         TabIndex        =   9
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   -71400
         TabIndex        =   20
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdout 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   -69720
         TabIndex        =   19
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Frame frame 
         Height          =   6015
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   13455
         Begin VB.CommandButton Command5 
            Caption         =   "Cari"
            Height          =   255
            Left            =   5640
            TabIndex        =   46
            Top             =   1680
            Width           =   615
         End
         Begin XPControls.XPText txtjum 
            Height          =   285
            Left            =   2280
            TabIndex        =   7
            Top             =   5160
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":0302
            TabIndex        =   36
            Top             =   5160
            Width           =   1455
         End
         Begin VB.ComboBox cmbstts 
            Height          =   315
            Left            =   2280
            TabIndex        =   8
            Top             =   5520
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":037A
            TabIndex        =   33
            Top             =   5520
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Input"
            Height          =   375
            Left            =   4200
            TabIndex        =   6
            Top             =   4680
            Width           =   1335
         End
         Begin XPControls.XPText txtbrg 
            Height          =   285
            Left            =   2280
            TabIndex        =   5
            Top             =   4680
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
         Begin XPControls.XPText txtkel 
            Height          =   855
            Left            =   2280
            TabIndex        =   4
            Top             =   3600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1508
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
         Begin XPControls.XPText txttelp 
            Height          =   285
            Left            =   2280
            TabIndex        =   3
            Top             =   3120
            Width           =   3135
            _ExtentX        =   5530
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
         Begin XPControls.XPText txtalmt 
            Height          =   855
            Left            =   2280
            TabIndex        =   2
            Top             =   2160
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1508
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
         Begin XPControls.XPText txtnama 
            Height          =   285
            Left            =   2280
            TabIndex        =   1
            Top             =   1680
            Width           =   3255
            _ExtentX        =   5741
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":03F2
            TabIndex        =   32
            Top             =   4680
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":0466
            TabIndex        =   31
            Top             =   3600
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":04D2
            TabIndex        =   30
            Top             =   3120
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":053E
            TabIndex        =   29
            Top             =   2160
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":05A8
            TabIndex        =   28
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox cmbskin 
            Height          =   315
            Left            =   6720
            TabIndex        =   13
            Text            =   "Combo1"
            Top             =   5160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   375
            Left            =   12120
            TabIndex        =   12
            Top             =   5040
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel matuc 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmreg.frx":0618
            TabIndex        =   14
            Top             =   3000
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel Ko 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":0676
            TabIndex        =   15
            Top             =   480
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmreg.frx":06EA
            TabIndex        =   16
            Top             =   1080
            Width           =   1335
         End
         Begin XPControls.XPText txtnmr 
            Height          =   285
            Left            =   2280
            TabIndex        =   17
            Top             =   480
            Width           =   2655
            _ExtentX        =   4683
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
         Begin MSComctlLib.ListView lv1 
            Height          =   4575
            Left            =   6720
            TabIndex        =   18
            Top             =   480
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   8070
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   794
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nama brg"
               Object.Width           =   3598
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Jumlah"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Status barang"
               Object.Width           =   2540
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
            Left            =   2280
            TabIndex        =   27
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
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
            CustomFormat    =   "dd MMM yyyy"
            Format          =   126353411
            CurrentDate     =   37623
         End
      End
      Begin XPControls.XPText txtcari 
         Height          =   375
         Left            =   -71040
         TabIndex        =   23
         Top             =   4920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
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
      Begin MSComctlLib.ListView lvpaket 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   24
         Top             =   840
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7011
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No.Servis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tgl Msk"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Alamat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Telp"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Keluhan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Teknisi"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Jenis Servis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Biaya Servis"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Biaya barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Kerusakan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "Perkiraan Selesai"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "Tgl ambil"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvdtl 
         Height          =   1935
         Left            =   -74640
         TabIndex        =   25
         Top             =   5520
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3413
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama barang"
            Object.Width           =   2823
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jumlah"
            Object.Width           =   2363
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status garansi"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvdtl2 
         Height          =   1935
         Left            =   -68880
         TabIndex        =   37
         Top             =   5520
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3413
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
            Text            =   "Kode brg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Jumlah"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Diskon"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sub total"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvplg 
         Height          =   5775
         Left            =   840
         TabIndex        =   43
         Top             =   840
         Width           =   13095
         _ExtentX        =   23098
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama "
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Alamat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "No.Telp"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Barang diganti"
         Height          =   255
         Left            =   -68760
         TabIndex        =   39
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Baram 
         Caption         =   "Barang Masuk"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cari nama pelanggan atau nomor servis"
         Height          =   375
         Left            =   -74520
         TabIndex        =   26
         Top             =   4920
         Width           =   3135
      End
   End
   Begin VB.Menu mn1 
      Caption         =   "Set status "
      Visible         =   0   'False
      Begin VB.Menu mnst1 
         Caption         =   "Set status 'setuju'"
      End
      Begin VB.Menu mnst2 
         Caption         =   "Pengambilan"
      End
      Begin VB.Menu mnst3 
         Caption         =   "Set 'batal'"
      End
   End
End
Attribute VB_Name = "frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ubah As Boolean
Dim sum As Currency
Dim tmbh As String
Sub dbgridplg()
On Error Resume Next

Set rstrans = New Recordset
'sql = "select t.*,nama_teknisi,nama_servis from tservis t left join servis s on t.kode_servis=s.kode_servis left join teknisi k on t.id_teknisi=k.id_teknisi where no_servis!='' " & tmbh & " order by no_servis desc"
sql = "select nama_cst,almt_cst,telp_cst from tservis group by nama_cst,almt_cst order by nama_cst"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![nama_cst]
        l.SubItems(2) = ![almt_cst]
        l.SubItems(3) = ![telp_cst]
        

    .MoveNext
    Loop
End With
End Sub
Sub dbgridplg2()
On Error Resume Next

Set rstrans = New Recordset
'sql = "select t.*,nama_teknisi,nama_servis from tservis t left join servis s on t.kode_servis=s.kode_servis left join teknisi k on t.id_teknisi=k.id_teknisi where no_servis!='' " & tmbh & " order by no_servis desc"
sql = "select nama_cst,almt_cst,telp_cst from tservis where nama_cst like '%" & txtcari2.Text & "%' group by nama_cst,almt_cst order by nama_cst"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![nama_cst]
        l.SubItems(2) = ![almt_cst]
        l.SubItems(3) = ![telp_cst]
        

    .MoveNext
    Loop
End With
End Sub

Sub cetakdaftar()
On Error Resume Next
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\fakturdaftar.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{tservis.no_servis}='" & txtnmr.Text & "'"
    .Formulas(0) = "nama='" & nama_toko & "'"
    .Formulas(1) = "almt='" & almt & "' + space(2) +'" & almt2 & "'"

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

End Sub
Sub cetakfaktur()
On Error Resume Next
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\fakturambil.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{tservis.no_servis}='" & txtnmr.Text & "'"
    .Formulas(0) = "nama='" & nama_toko & "'"
.Formulas(1) = "almt='" & almt & "' + space(2) +'" & almt2 & "'"
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

End Sub

Sub dbgrid()
On Error Resume Next

Set rstrans = New Recordset
sql = "select t.*,nama_teknisi,nama_servis from tservis t left join servis s on t.kode_servis=s.kode_servis left join teknisi k on t.id_teknisi=k.id_teknisi where no_servis!='' " & tmbh & " order by no_servis desc"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvpaket.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvpaket.ListItems.Add(, , lvpaket.ListItems.count + 1)
        l.SubItems(1) = ![no_servis]
        l.SubItems(2) = Format(![tgl_msk], "dd-mmm-yyyy")
        l.SubItems(3) = ![nama_cst]
        l.SubItems(4) = ![almt_cst]
        l.SubItems(5) = ![telp_cst]
        l.SubItems(6) = ![keluhan]
        l.SubItems(7) = ![status_servis]
        l.SubItems(8) = ![nama_teknisi]
        l.SubItems(9) = ![nama_servis]
        l.SubItems(10) = Format(![biaya_servis], "#,#")
        l.SubItems(11) = Format(![biaya_brg], "#,#")
        l.SubItems(12) = Format(![biaya_brg] + ![biaya_servis], "#,#")
        l.SubItems(13) = ![kerusakan]
        l.SubItems(14) = ![perkiraan] + " hari"
        l.SubItems(15) = Format(![tgl_out], "dd-mmm-yyyy")
        
       

    .MoveNext
    Loop
End With
End Sub
Sub Dbgrid2()
On Error Resume Next

Set rstrans = New Recordset
sql = "select t.*,nama_teknisi,nama_servis from tservis t left join servis s on t.kode_servis=s.kode_servis left join teknisi k on t.id_teknisi=k.id_teknisi where (no_servis like '%" & txtcari.Text & "%' or nama_cst like '%" & txtcari.Text & "%') " & tmbh & " order by no_servis desc"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvpaket.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvpaket.ListItems.Add(, , lvpaket.ListItems.count + 1)
        l.SubItems(1) = ![no_servis]
        l.SubItems(2) = Format(![tgl_msk], "dd-mmm-yyyy")
        l.SubItems(3) = ![nama_cst]
        l.SubItems(4) = ![almt_cst]
        l.SubItems(5) = ![telp_cst]
        l.SubItems(6) = ![keluhan]
        l.SubItems(7) = ![status_servis]
        l.SubItems(8) = ![nama_teknisi]
        l.SubItems(9) = ![nama_servis]
        l.SubItems(10) = Format(![biaya_servis], "#,#")
        l.SubItems(11) = Format(![biaya_brg], "#,#")
        l.SubItems(12) = Format(![biaya_brg] + ![biaya_servis], "#,#")
        l.SubItems(13) = ![kerusakan]
        l.SubItems(14) = ![perkiraan] + " hari"
        l.SubItems(15) = Format(![tgl_out], "dd-mmm-yyyy")
        

    .MoveNext
    Loop
End With
End Sub
Private Sub cmbjns_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub cmbmatu_Click()
txtharga.SetFocus
matuc.Caption = cmbmatu.Text
End Sub

Private Sub cmbskin_Click()
Skinpath = App.Path & "\skin\" & cmbskin.Text & ".skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub

Private Sub cmbstts2_Click()
If cmbstts2.Text = "Semua" Then

tmbh = ""
Else
tmbh = " and status_servis='" & cmbstts2.Text & "'"
End If
dbgrid

End Sub

Private Sub cmbstts2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub cmdbatal_Click()
awal
kosong
End Sub

Private Sub cmdhapus_Click()
With lvpaket.SelectedItem
If .SubItems(7) <> "daftar" Then
MsgBox "Sudah dalam proses!", vbCritical, judul
Exit Sub
End If
End With
If lvpaket.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin data ini dihapus?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from tservis where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
lvdtl.ListItems.Clear
Dbgrid2
MsgBox "Data berhasil dihapus", vbInformation, judul
txtcari.SetFocus

End Sub

Private Sub cmdhps_Click()
If lvpaket.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin data ini dihapus?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from tservis where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
jual.Execute "delete from tservis_dtl2 where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"

lvdtl.ListItems.Clear
Dbgrid2
MsgBox "Data berhasil dihapus", vbInformation, judul
txtcari.SetFocus

End Sub

Private Sub cmdout_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
If txtnmr.Text = "" Or txtnama.Text = "" Or txttelp.Text = "" Or lv1.ListItems.count = 0 Then Exit Sub

If MsgBox("Simpan data?", vbYesNo, judul) = vbNo Then Exit Sub
If ubah = False Then
ktr = ""
Else
ktr = " and no_servis!='" & txtnmr.Text & "'"
End If
Set RS = New Recordset
RS.Open "select no_servis from tservis where no_servis='" & txtnmr.Text & "' " & ktr & "", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nomor servis sudah ada,bedakan walau satu huruf", vbCritical, judul
Exit Sub
End If


If ubah = True Then
jual.Execute "delete from tservis where no_servis='" & txtnmr.Text & "'"
End If
jual.Execute "insert into tservis(no_servis,tgl_msk,nama_cst,almt_cst,telp_cst,keluhan,status_servis) values('" & txtnmr.Text & "','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & txtnama.Text & "','" & txtalmt.Text & "','" & txttelp.Text & "','" & txtkel.Text & "','daftar')"
For I = 1 To lv1.ListItems.count
jual.Execute "insert into tservis_dtl1 values('" & txtnmr.Text & "','" & lv1.ListItems(I).SubItems(1) & "','" & lv1.ListItems(I).SubItems(2) & "','" & lv1.ListItems(I).SubItems(3) & "')"
Next I


If ubah = False Then
If MsgBox("Data berhasil disimpan,cetak faktur?", vbYesNo, judul) = vbYes Then
cetakdaftar
Else
awal
End If
Else
MsgBox "Data berhasil diubah", vbInformation, judul
End If
dbgrid
End Sub
Private Sub cmdtambah_Click()
txtnmr.Enabled = True
ubah = False
tambah
kosong
idoto
txtnama.SetFocus
End Sub

Sub awal()
frame.Enabled = False
cmdtambah.Enabled = True
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
End Sub

Sub tambah()
frame.Enabled = True
kosong
cmdtambah.Enabled = False
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End Sub

Private Sub cmdubah_Click()
With lvpaket.SelectedItem
If .SubItems(7) <> "daftar" Then
MsgBox "Sudah dalam proses!", vbCritical, judul
Exit Sub
End If
End With
If lvpaket.ListItems.count = 0 Then Exit Sub
With lvpaket.SelectedItem
If .SubItems(7) = "diambil" Then
MsgBox "Sudah diambil"
Exit Sub
End If
End With
Tab1.Tab = 0
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
txtnmr.Text = lvpaket.SelectedItem.SubItems(1)
txtnmr_KeyPress (13)
ubah = True
frame.Enabled = True

End Sub

Private Sub Command1_Click()
If Not lv1.SelectedItem Is Nothing Then
lv1.ListItems.Remove lv1.SelectedItem.Index
End If
For I = 1 To lv1.ListItems.count
    lv1.ListItems(I).Text = I
Next I
End Sub

Private Sub Command2_Click()
If txtnmr.Text = "" Or txtnama.Text = "" Or txttelp.Text = "" Or txtbrg.Text = "" Then
MsgBox "Data belum lengkap", vbCritical, judul
Exit Sub
End If
If txtjum.Text = "" Then
txtjum.Text = "1"
End If
isigrid
End Sub
Sub isigrid()
On Error Resume Next
Set cari = lv1.FindItem(txtbrg.Text, 1, , 1)
lv1.SelectedItem = cari

If Not cari Is Nothing Then
MsgBox "Barang sudah terdaftar"
txtbrg.SetFocus
Exit Sub
End If

    Set butir = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
    With butir

    .SubItems(1) = txtbrg.Text
    .SubItems(2) = txtjum.Text
    .SubItems(3) = cmbstts.Text
    End With
    txtbrg.Text = ""
    txtjum.Text = ""
txtbrg.SetFocus
 
End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
A = "SRV" & "%"
sql = "Select no_servis from tservis where no_servis like '" & A & "' order by no_servis Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
txtnmr.Text = "SRV" + "000001"
Else
j = val(Right(rsplg(0), 6))
No = "SRV" + Format(Str(j + 1), "000000")
txtnmr.Text = No
End If
End Sub
Private Sub Command3_Click()
If MsgBox("Cetak faktur pendaftaran?", vbYesNo, judul) = vbNo Then Exit Sub
txtnmr.Text = lvpaket.SelectedItem.SubItems(1)
cetakdaftar
End Sub

Private Sub Command4_Click()
If lvpaket.SelectedItem.SubItems(7) <> "diambil" Then
MsgBox "Status terakhir harus 'diambil'"
Exit Sub
End If

If MsgBox("Cetak faktur pengambilan?", vbYesNo, judul) = vbNo Then Exit Sub
txtnmr.Text = lvpaket.SelectedItem.SubItems(1)
cetakfaktur

End Sub

Private Sub Command5_Click()
Tab1.Tab = 2
txtcari2.SetFocus
End Sub

Private Sub Form_Load()
adskin
awal
dbgridplg
tgl.Value = Now
cmbstts.Clear
cmbstts.AddItem "Garansi"
cmbstts.AddItem "Habis/tidak garansi"
cmbstts.Text = "Habis/tidak garansi"
cmbstts2.AddItem "Semua"
cmbstts2.AddItem "daftar"
cmbstts2.AddItem "tunggu konfirmasi"
cmbstts2.AddItem "setuju"
cmbstts2.AddItem "perbaikan"
cmbstts2.AddItem "selesai"
cmbstts2.AddItem "diambil"
cmbstts2.AddItem "batal"

cmbstts2.Text = "Semua"
cmdhps.Enabled = IIf(rshusus.Fields(18) = False, False, True)

tmbh = ""
cmbstts.AddItem "Semua"
dbgrid

Tab1.Tab = 0
    Skinpath = App.Path & "\skin\chizh.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub




Sub adskin()
With cmbskin
.AddItem "B-studio"
.AddItem "CHIZH"
.AddItem "Cooller"
.AddItem "DiamanteGreen"
.AddItem "Dogmas2"
.AddItem "executive2"
.AddItem "galaxy"
.AddItem "golden"
.AddItem "green"
.AddItem "innovex"
.AddItem "jade2"
.AddItem "Le-black"
.AddItem "Linuxglome"
.AddItem "LinuxUbuntu"
.AddItem "LinuxUbuntu2"
.AddItem "mac"
.AddItem "magnificblue"
.AddItem "media"
.AddItem "mediablue"
.AddItem "metallic"
.AddItem "msn2008"
.AddItem "mxs39"
.AddItem "mxs42"
.AddItem "mxs45"
.AddItem "mxs54"
.AddItem "mxs59"
.AddItem "mxs58"
.AddItem "mxs61"
.AddItem "mxs100"
.AddItem "oncyan"
.AddItem "orange"
.AddItem "orion"
.AddItem "orionnext"
.AddItem "paper"
.AddItem "plasmoid"
.AddItem "plastic"
.AddItem "retro"
.AddItem "steelblue"
.AddItem "tdpanther"
.AddItem "tiger3"
.AddItem "topsecret"
.AddItem "triton"
.AddItem "vistablue"
.AddItem "vvv"
.AddItem "web-ii"
.AddItem "winaqua"
.AddItem "winaqua2"
.AddItem "winter"
.AddItem "zhelezo"
.AddItem "xfactor3"






End With
End Sub

Private Sub lvbrg_DblClick()
If lvbrg.ListItems.count = 0 Then Exit Sub
txtkode2.Text = lvbrg.SelectedItem.SubItems(1)
txtkode2_KeyPress (13)
Tab1.Tab = 0
End Sub
Sub kosong()
txtnmr.Text = ""
txtnama.Text = ""
txttelp.Text = ""
txtbrg.Text = ""
txtkel.Text = ""
txtjum.Text = ""
lv1.ListItems.Clear
End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub

Private Sub lvpaket_Click()
On Error Resume Next
If lvpaket.ListItems.count = 0 Then Exit Sub
Set rstrans = New Recordset
sql = "select p.* from tservis_dtl1 p,tservis t where t.no_servis=p.no_servis and t.no_servis='" & lvpaket.SelectedItem.SubItems(1) & "' order by nama_brg"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvdtl.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvdtl.ListItems.Add(, , lvdtl.ListItems.count + 1)
        
        l.SubItems(1) = !nama_brg
    l.SubItems(2) = !jumlah
    l.SubItems(3) = !stts_garansi
   
    
        
    .MoveNext
    Loop
End With

Set rstrans = New Recordset
sql = "select p.*,deskripsi from tservis_dtl2 p,tblbarang t where t.kode_brg=p.kode_brg and p.no_servis='" & lvpaket.SelectedItem.SubItems(1) & "' order by deskripsi"

Set rstrans = jual.Execute(sql)
lvdtl2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvdtl2.ListItems.Add(, , lvdtl2.ListItems.count + 1)
        
        l.SubItems(1) = !kode_brg
    l.SubItems(2) = !deskripsi
    l.SubItems(3) = !jum_brg
     l.SubItems(4) = !harga
      l.SubItems(5) = !diskon
       l.SubItems(6) = Format(!subttl, "#,#")
   
    
        
    .MoveNext
    Loop
End With

End Sub

Private Sub lvpaket_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
lvpaket_Click

End If

End Sub

Private Sub lvpaket_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If lvpaket.ListItems.count = 0 Then Exit Sub
If lvpaket.SelectedItem.SubItems(1) = "" Then Exit Sub
If Button = vbRightButton Then
    Me.PopupMenu mn1
    End If

End Sub

Private Sub lvplg_DblClick()
If lvplg.ListItems.count = 0 Then Exit Sub

txtnama.Text = lvplg.SelectedItem.SubItems(1)
txtalmt.Text = lvplg.SelectedItem.SubItems(2)
txttelp.Text = lvplg.SelectedItem.SubItems(3)
Tab1.Tab = 0
End Sub

Private Sub mnst1_Click()
If lvpaket.SelectedItem.SubItems(7) <> "tunggu konfirmasi" Then
MsgBox "Status terakhir harus 'tunggu konfirmasi'"
Exit Sub
End If
If MsgBox("Proses ke 'setuju'?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "update tservis set status_servis='setuju' where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
MsgBox "Update status berhasil", vbInformation, judul
dbgrid
End Sub

Private Sub mnst2_Click()
If lvpaket.SelectedItem.SubItems(7) <> "selesai" Then
MsgBox "Status terakhir harus 'selesai'"
Exit Sub
End If

If MsgBox("Proses pengambilan oleh konsumen?", vbYesNo, judul) = vbNo Then Exit Sub
With lvpaket.SelectedItem
frmout.lblnmr.Caption = .SubItems(1)
frmout.lblbiaya.Caption = Format(.SubItems(11), "#,#")
frmout.lblbiaya2.Caption = Format(.SubItems(10), "#,#")
frmout.lblttl.Caption = Format(.SubItems(12), "#,#")
End With
frmout.Show
End Sub

Private Sub mnst3_Click()
If lvpaket.SelectedItem.SubItems(7) <> "tunggu konfirmasi" Then
MsgBox "Status terakhir harus 'tunggu konfirmasi'"
Exit Sub
End If
If MsgBox("Proses ke 'batal'?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "update tservis set status_servis='batal' where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
MsgBox "Update status berhasil", vbInformation, judul
dbgrid
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
txtcari.Text = ""
txtcari.SetFocus
Else
If Tab1.Tab = 2 Then
txtcari2.Text = ""
txtcari2.SetFocus
End If
End If
End Sub

Private Sub txtcari_Change()
Dbgrid2
End Sub

Private Sub txtdp_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub txtfst_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtfst.Text = "" Then
Cmdsimpan.SetFocus
Else
End If
End If
End Sub



Private Sub txtcari2_Change()
dbgridplg2
End Sub

Private Sub txtcari2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lvplg.ListItems.count = 0 Then Exit Sub
lvplg.SetFocus
End If
End Sub

Private Sub txtdisk_Change()
lblsub.Caption = val(Format(txtjum.Text, Number)) * val(Format(txtharga.Text, Number)) - val(txtdisk.Text)

End Sub

Private Sub txtdisk_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3_Click
End If

End Sub

Private Sub txtdiskp_Change()
End Sub

Private Sub txtdiskp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3_Click
End If

End Sub

Private Sub txtharga_Change()
txtdiskp_Change
End Sub

Private Sub txtharga_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub txtcari2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then
lvplg.SetFocus
End If
End Sub

Private Sub txtjum_Change()
'lblsub.Caption = val(Format(txtjum.Text, Number)) * val(Format(txtharga.Text, Number)) - val(Format(txtdisk.Text, Number))
End Sub

Private Sub txtjum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2_Click
End If
End Sub

Private Sub txtnmr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New Recordset
RS.Open "select * from tservis where no_servis='" & txtnmr.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
With RS
tgl.Value = !tgl_msk
txtnama.Text = !nama_cst
txtalmt.Text = !almt_cst
txttelp.Text = !telp_cst
txtkel.Text = !keluhan
End With


lv1.ListItems.Clear
Set rsd = New Recordset
sql = ""

rsd.Open "select p.* from tservis_dtl1 p,tservis t where t.no_servis=p.no_servis and t.no_servis='" & txtnmr.Text & "' order by nama_brg", jual, adOpenStatic, adLockOptimistic

With rsd
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
         l.SubItems(1) = !nama_brg
    l.SubItems(2) = !jumlah
    l.SubItems(3) = !stts_garansi
        

    .MoveNext
    Loop
End With
End If
End If
End Sub

Private Sub txtlama_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub


Private Sub txtkode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New Recordset
RS.Open "select deskripsi,harga_jual from tblbarang where kode_brg='" & txtkode2.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Barang tidak terdaftar", vbCritical, judul
Else
txtnama2.Text = RS!deskripsi
txtharga.Text = RS!harga_jual
txtjum.SetFocus
End If
End If
End Sub

