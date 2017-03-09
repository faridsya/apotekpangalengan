VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmpaket 
   Caption         =   "Data Racikan"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   14445
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   13320
      OleObjectBlob   =   "frmpaket.frx":0000
      Top             =   7320
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7095
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "frmpaket.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdtambah"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdsimpan"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdbatal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdout"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frame"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Data Racikan"
      TabPicture(1)   =   "frmpaket.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdhapus"
      Tab(1).Control(1)=   "cmdubah"
      Tab(1).Control(2)=   "SkinLabel8"
      Tab(1).Control(3)=   "txtcari"
      Tab(1).Control(4)=   "lvpaket"
      Tab(1).Control(5)=   "lvdtl"
      Tab(1).Control(6)=   "Label2"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Data obat"
      TabPicture(2)   =   "frmpaket.frx":026C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtcari2"
      Tab(2).Control(1)=   "lvbrg"
      Tab(2).Control(2)=   "Label1"
      Tab(2).ControlCount=   3
      Begin VB.Frame frame 
         Height          =   5655
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   13455
         Begin ACTIVESKINLibCtl.SkinLabel lblharga 
            Height          =   255
            Left            =   2280
            OleObjectBlob   =   "frmpaket.frx":0288
            TabIndex        =   38
            Top             =   1680
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmpaket.frx":02E6
            TabIndex        =   37
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            Caption         =   "Pilih obat"
            Height          =   3255
            Left            =   360
            TabIndex        =   25
            Top             =   1920
            Width           =   4695
            Begin VB.CommandButton Command3 
               Caption         =   "Input"
               Height          =   255
               Left            =   3720
               TabIndex        =   8
               Top             =   2280
               Width           =   855
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Cari"
               Height          =   255
               Left            =   3720
               TabIndex        =   3
               Top             =   360
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel lblsub 
               Height          =   255
               Left            =   1680
               OleObjectBlob   =   "frmpaket.frx":034E
               TabIndex        =   34
               Top             =   2760
               Width           =   2055
            End
            Begin XPControls.XPText txtdisk 
               Height          =   285
               Left            =   2280
               TabIndex        =   7
               Top             =   2280
               Width           =   1335
               _ExtentX        =   2355
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
            Begin XPControls.XPText txtdiskp 
               Height          =   285
               Left            =   1680
               TabIndex        =   6
               Top             =   2280
               Width           =   495
               _ExtentX        =   873
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
               Left            =   1680
               TabIndex        =   5
               Top             =   1800
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
            Begin XPControls.XPText txtharga 
               Height          =   285
               Left            =   1680
               TabIndex        =   4
               Top             =   1320
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
            Begin XPControls.XPText txtnama2 
               Height          =   285
               Left            =   1680
               TabIndex        =   33
               Top             =   840
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
               Locked          =   -1  'True
            End
            Begin XPControls.XPText txtkode2 
               Height          =   285
               Left            =   1680
               TabIndex        =   32
               Top             =   360
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":03AC
               TabIndex        =   31
               Top             =   2760
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":0418
               TabIndex        =   30
               Top             =   2280
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":0482
               TabIndex        =   29
               Top             =   1800
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":04EC
               TabIndex        =   26
               Top             =   360
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":055C
               TabIndex        =   27
               Top             =   840
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmpaket.frx":05CC
               TabIndex        =   28
               Top             =   1320
               Width           =   2175
            End
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   375
            Left            =   12120
            TabIndex        =   19
            Top             =   5040
            Width           =   1215
         End
         Begin VB.ComboBox cmbskin 
            Height          =   315
            Left            =   5400
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   5160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel matuc 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmpaket.frx":063E
            TabIndex        =   18
            Top             =   3000
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel Ko 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmpaket.frx":069C
            TabIndex        =   20
            Top             =   480
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   495
            Left            =   480
            OleObjectBlob   =   "frmpaket.frx":0712
            TabIndex        =   21
            Top             =   1080
            Width           =   1335
         End
         Begin XPControls.XPText txtkode 
            Height          =   285
            Left            =   2280
            TabIndex        =   1
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
         End
         Begin XPControls.XPText txtnama 
            Height          =   525
            Left            =   2280
            TabIndex        =   2
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   926
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
            MultiLine       =   -1  'True
         End
         Begin MSComctlLib.ListView lv1 
            Height          =   4575
            Left            =   5280
            TabIndex        =   22
            Top             =   480
            Width           =   8055
            _ExtentX        =   14208
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   794
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Kode brg"
               Object.Width           =   2470
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Nama brg"
               Object.Width           =   3598
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Harga jual"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Jumlah"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Diskon"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Sub Total"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.CommandButton cmdout 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   480
         TabIndex        =   0
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -64560
         TabIndex        =   13
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "&Ubah"
         Height          =   375
         Left            =   -65760
         TabIndex        =   12
         Top             =   3720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   -73680
         OleObjectBlob   =   "frmpaket.frx":0788
         TabIndex        =   11
         Top             =   4200
         Width           =   2295
      End
      Begin XPControls.XPText txtcari 
         Height          =   375
         Left            =   -72600
         TabIndex        =   23
         Top             =   3720
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
         Height          =   3135
         Left            =   -74160
         TabIndex        =   24
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5530
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
            Text            =   "Kode paket"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama paket"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Harga"
            Object.Width           =   2011
         EndProperty
      End
      Begin XPControls.XPText txtcari2 
         Height          =   375
         Left            =   -71760
         TabIndex        =   35
         Top             =   6240
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
      Begin MSComctlLib.ListView lvbrg 
         Height          =   5655
         Left            =   -74640
         TabIndex        =   36
         Top             =   480
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   9975
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
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "HArga jual 2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Harga Jual 3"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvdtl 
         Height          =   2415
         Left            =   -73200
         TabIndex        =   39
         Top             =   4560
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4260
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama brg"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Harga jual"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sub Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cari nama racikan"
         Height          =   375
         Left            =   -74040
         TabIndex        =   41
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cari nama obat"
         Height          =   255
         Left            =   -74520
         TabIndex        =   40
         Top             =   6360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmpaket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ubah As Boolean
Dim sum As Currency
Sub dbgridbrg()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblbarang order by deskripsi"
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
  l.SubItems(7) = Format(![Harga_jual2], "#,#")
 l.SubItems(8) = Format(![Harga_jual3], "#,#")
    .MoveNext
    Loop
End With


End Sub
Sub dbgridcari()
On Error Resume Next

Set rstrans = New Recordset

stri = Replace(txtcari2.Text, "'", "''")

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
  l.SubItems(7) = Format(![Harga_jual2], "#,#")
 l.SubItems(8) = Format(![Harga_jual3], "#,#")
    .MoveNext
    Loop
End With

End Sub
Sub dbgrid()
On Error Resume Next

Set rstrans = New Recordset
sql = "select * from paket  order by nama_pkt"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvpaket.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvpaket.ListItems.Add(, , lvpaket.ListItems.count + 1)
        l.SubItems(1) = ![kode_pkt]
        l.SubItems(2) = ![nama_pkt]
        l.SubItems(3) = ![harga_pkt]
       

    .MoveNext
    Loop
End With
End Sub
Sub Dbgrid2()
On Error Resume Next

Set rstrans = New Recordset
sql = "select * from paket where nama_pkt like '%" & txtcari.Text & "%' order by nama_pkt"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvpaket.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvpaket.ListItems.Add(, , lvpaket.ListItems.count + 1)
        l.SubItems(1) = ![kode_pkt]
        l.SubItems(2) = ![nama_pkt]
        l.SubItems(3) = ![harga_pkt]

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

Private Sub Cmdbatal_Click()
awal
kosong
End Sub

Private Sub cmdhapus_Click()
If lvpaket.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin data ini dihapus?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from paket where kode_pkt='" & lvpaket.SelectedItem.SubItems(1) & "'"
lvdtl.ListItems.Clear
Dbgrid2
MsgBox "Data berhasil dihapus", vbInformation, judul
txtcari.SetFocus

End Sub

Private Sub cmdout_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
If txtkode.Text = "" Or txtnama.Text = "" Or LV1.ListItems.count = 0 Then Exit Sub

If MsgBox("Simpan data?", vbYesNo, judul) = vbNo Then Exit Sub
If ubah = False Then
ktr = ""
Else
ktr = " and kode_pkt!='" & txtkode.Text & "'"
End If
Set RS = New Recordset
RS.Open "select kode_pkt from paket where kode_pkt='" & txtkode.Text & "' " & ktr & "", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Kode paket sudah ada,bedakan walau satu huruf", vbCritical, judul
Exit Sub
End If

Set RS = New Recordset
RS.Open "select nama_pkt from paket where kode_pkt='" & txtkode.Text & "' " & ktr & "", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama paket sudah ada,bedakan walau satu huruf", vbCritical, judul
Exit Sub
End If
If ubah = True Then
jual.Execute "delete from paket where kode_pkt='" & txtkode.Text & "'"
End If
jual.Execute "insert into paket values('" & txtkode.Text & "','" & txtnama.Text & "','" & sum & "')"
For I = 1 To LV1.ListItems.count
jual.Execute "insert into paket_detil values('" & txtkode.Text & "','" & LV1.ListItems(I).SubItems(1) & "','" & LV1.ListItems(I).SubItems(3) & "','" & LV1.ListItems(I).SubItems(4) & "','" & LV1.ListItems(I).SubItems(5) & "',0,'" & LV1.ListItems(I).SubItems(6) & "')"
Next I


If ubah = False Then
If MsgBox("Data berhasil disimpan,tambah data?", vbYesNo, judul) = vbYes Then
cmdtambah_Click
Else
awal
End If
Else
MsgBox "Data berhasil diubah", vbInformation, judul
End If
dbgrid
End Sub
Sub ttl()
lblharga.Caption = ""
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(I).SubItems(6)
Next I
lblharga.Caption = Format(sum, "#,#")
End Sub
Private Sub cmdtambah_Click()
txtkode.Enabled = True
ubah = False
tambah
kosong
txtkode.SetFocus
End Sub

Sub awal()
Frame.Enabled = False
cmdtambah.Enabled = True
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
End Sub

Sub tambah()
Frame.Enabled = True
kosong
cmdtambah.Enabled = False
cmdsimpan.Enabled = True
cmdbatal.Enabled = True
End Sub

Private Sub cmdubah_Click()
If lvpaket.ListItems.count = 0 Then Exit Sub
Tab1.Tab = 0
txtkode.Enabled = True
cmdsimpan.Enabled = True
cmdbatal.Enabled = True
txtkode.Text = lvpaket.SelectedItem.SubItems(1)
txtkode_KeyPress (13)
ubah = True
Frame.Enabled = True

End Sub

Private Sub Command1_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
For I = 1 To LV1.ListItems.count
    LV1.ListItems(I).Text = I
Next I
ttl
End Sub

Private Sub Command2_Click()
Tab1.Tab = 2
End Sub

Private Sub Command3_Click()
If txtkode2.Text = "" Or txtjum.Text = "" Or txtnama2.Text = "" Or txtharga.Text = "" Then Exit Sub
isigrid

End Sub

Private Sub Form_Load()
adskin
dbgrid
dbgridbrg
awal
Tab1.Tab = 0
    Skinpath = App.Path & "\skin\chizh.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Sub isigrid()
On Error Resume Next
Set cari = LV1.FindItem(txtkode2.Text, 1, , 1)
LV1.SelectedItem = cari

If Not cari Is Nothing Then
MsgBox "Barang sudah terdaftar"
txtfst.SetFocus
Exit Sub
End If

    Set butir = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
    With butir

    .SubItems(1) = txtkode2.Text
    .SubItems(2) = txtnama2.Text
    .SubItems(3) = val(Format(txtharga.Text, Number))
    .SubItems(4) = val(Format(txtjum.Text, Number))
    .SubItems(5) = val(Format(txtdisk.Text, Number))
    .SubItems(6) = val(Format(txtharga.Text, Number)) * val(Format(txtjum.Text)) - val(Format(txtdisk.Text, Number))
    
    
    

 End With
 ttl
kosong2
Command2.SetFocus
 
End Sub
Sub kosong2()
txtkode2.Text = ""
txtnama2.Text = ""
txtjum.Text = ""
txtharga.Text = ""
txtdiskp.Text = ""
txtdisk.Text = ""
lblsub.Caption = ""
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
txtkode.Text = ""
txtnama.Text = ""
txtnama2.Text = ""
txtharga.Text = ""
LV1.ListItems.Clear
txtkode2.Text = ""
txtjum.Text = ""
lblsub.Caption = ""
txtdiskp.Text = ""
txtdisk.Text = ""
End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub

Private Sub lvpaket_Click()
On Error Resume Next

Set rstrans = New Recordset
sql = "select p.*,deskripsi,satuan from paket_detil p,tblbarang t where t.kode_brg=p.kode_brg and kode_pkt='" & lvpaket.SelectedItem.SubItems(1) & "' order by deskripsi"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvdtl.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvdtl.ListItems.Add(, , lvdtl.ListItems.count + 1)
        
        l.SubItems(1) = !kode_brg
    l.SubItems(2) = !deskripsi
    l.SubItems(3) = !harga_jual
    l.SubItems(4) = !jumlah
    l.SubItems(5) = !diskon
    l.SubItems(6) = !subttl
    
        
    .MoveNext
    Loop
End With

End Sub

Private Sub lvpaket_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
lvpaket_Click

End If

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
cmdsimpan.SetFocus
Else
End If
End If
End Sub

Private Sub txtcari2_Change()
dbgridcari
End Sub

Private Sub txtcari2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lvbrg.ListItems.count = 0 Then Exit Sub
lvbrg.SetFocus
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
If val(txtdiskp.Text) > 100 Then
MsgBox "ERORR", , judul
txtdiskp.Text = ""
txtdiskp.SetFocus
Exit Sub
End If

txtdisk.Text = val(txtdiskp.Text) / 100 * val(Format(txtjum.Text, Number)) * val(Format(txtharga.Text, Number))
lblsub.Caption = val(Format(txtjum.Text, Number)) * val(Format(txtharga.Text, Number)) - val(txtdisk.Text)

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

Private Sub txtjum_Change()
'lblsub.Caption = val(Format(txtjum.Text, Number)) * val(Format(txtharga.Text, Number)) - val(Format(txtdisk.Text, Number))
txtdiskp_Change
End Sub

Private Sub txtjum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command3_Click
End If
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New Recordset
RS.Open "select * from paket where kode_pkt='" & txtkode.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
With RS
txtnama.Text = !nama_pkt

lblharga.Caption = !harga_pkt

End With


LV1.ListItems.Clear
Set rsd = New Recordset
rsd.Open "select p.*,deskripsi,satuan from paket_detil p,tblbarang t where t.kode_brg=p.kode_brg and kode_pkt='" & txtkode.Text & "' order by deskripsi", jual, adOpenStatic, adLockOptimistic

With rsd
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = !kode_brg
    l.SubItems(2) = !deskripsi
    l.SubItems(3) = !harga_jual
    l.SubItems(4) = !jumlah
    l.SubItems(5) = !diskon
    l.SubItems(6) = !subttl
        

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
