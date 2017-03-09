VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmproses 
   Caption         =   "Proses Teknisi"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15105
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   13800
      OleObjectBlob   =   "frmproses.frx":0000
      Top             =   6720
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   13361
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "frmproses.frx":0234
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdout"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdbatal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdsimpan"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Data Paket"
      TabPicture(1)   =   "frmproses.frx":0250
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Baram"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "lvdtl2"
      Tab(1).Control(5)=   "lvdtl"
      Tab(1).Control(6)=   "lvpaket"
      Tab(1).Control(7)=   "txtcari"
      Tab(1).Control(8)=   "cmbstts2"
      Tab(1).Control(9)=   "cmdubah"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Data barang"
      TabPicture(2)   =   "frmproses.frx":026C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtcari2"
      Tab(2).Control(1)=   "lvbrg"
      Tab(2).Control(2)=   "Label1"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdubah 
         Caption         =   "Update data"
         Height          =   255
         Left            =   -62760
         TabIndex        =   48
         Top             =   4800
         Width           =   2415
      End
      Begin VB.ComboBox cmbstts2 
         Height          =   315
         Left            =   -73920
         TabIndex        =   46
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Update"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2160
         TabIndex        =   30
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton cmdout 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   3960
         TabIndex        =   29
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Frame frame 
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   14655
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmproses.frx":0288
            TabIndex        =   40
            Top             =   2160
            Width           =   1575
         End
         Begin XPControls.XPText txtbiaya 
            Height          =   285
            Left            =   2280
            TabIndex        =   39
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
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
         Begin VB.ComboBox cmbjns 
            Height          =   315
            Left            =   2280
            TabIndex        =   3
            Text            =   "Combo1"
            Top             =   1800
            Width           =   2415
         End
         Begin VB.ComboBox cmbteknisi 
            Height          =   315
            Left            =   2280
            TabIndex        =   0
            Top             =   240
            Width           =   2295
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmproses.frx":02FE
            TabIndex        =   38
            Top             =   240
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   480
            OleObjectBlob   =   "frmproses.frx":036A
            TabIndex        =   37
            Top             =   1800
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmproses.frx":03E0
            TabIndex        =   36
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cmbskin 
            Height          =   315
            Left            =   5400
            TabIndex        =   24
            Text            =   "Combo1"
            Top             =   5160
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   375
            Left            =   12120
            TabIndex        =   23
            Top             =   5040
            Width           =   1215
         End
         Begin VB.Frame Frame1 
            Caption         =   "Barang yang harus diganti"
            Height          =   3255
            Left            =   360
            TabIndex        =   12
            Top             =   2520
            Width           =   4695
            Begin VB.CommandButton Command2 
               Caption         =   "Cari"
               Height          =   255
               Left            =   3720
               TabIndex        =   4
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Input"
               Height          =   255
               Left            =   3720
               TabIndex        =   8
               Top             =   2280
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel lblsub 
               Height          =   255
               Left            =   1680
               OleObjectBlob   =   "frmproses.frx":0446
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
               TabIndex        =   16
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
               OleObjectBlob   =   "frmproses.frx":04A4
               TabIndex        =   17
               Top             =   2760
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmproses.frx":0510
               TabIndex        =   18
               Top             =   2280
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmproses.frx":057A
               TabIndex        =   19
               Top             =   1800
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmproses.frx":05E4
               TabIndex        =   20
               Top             =   360
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmproses.frx":0658
               TabIndex        =   21
               Top             =   840
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   375
               Left            =   360
               OleObjectBlob   =   "frmproses.frx":06CC
               TabIndex        =   22
               Top             =   1320
               Width           =   2175
            End
         End
         Begin ACTIVESKINLibCtl.SkinLabel matuc 
            Height          =   255
            Left            =   4080
            OleObjectBlob   =   "frmproses.frx":073E
            TabIndex        =   25
            Top             =   3000
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel Ko 
            Height          =   375
            Left            =   480
            OleObjectBlob   =   "frmproses.frx":079C
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   495
            Left            =   480
            OleObjectBlob   =   "frmproses.frx":081C
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
         End
         Begin XPControls.XPText txtlama 
            Height          =   285
            Left            =   2280
            TabIndex        =   1
            Top             =   600
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
         Begin XPControls.XPText txtrusak 
            Height          =   645
            Left            =   2280
            TabIndex        =   2
            Top             =   1080
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1138
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
            TabIndex        =   28
            Top             =   480
            Width           =   9135
            _ExtentX        =   16113
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
            NumItems        =   8
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
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "hpp"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin XPControls.XPText txtcari 
         Height          =   375
         Left            =   -71400
         TabIndex        =   31
         Top             =   4800
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
      Begin XPControls.XPText txtcari2 
         Height          =   375
         Left            =   -71760
         TabIndex        =   32
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
         TabIndex        =   33
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
      Begin MSComctlLib.ListView lvpaket 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   41
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
            Text            =   "Perkiraan"
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
         Left            =   -74880
         TabIndex        =   42
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
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jumlah"
            Object.Width           =   2010
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status garansi"
            Object.Width           =   2999
         EndProperty
      End
      Begin MSComctlLib.ListView lvdtl2 
         Height          =   1935
         Left            =   -69360
         TabIndex        =   43
         Top             =   5520
         Width           =   8775
         _ExtentX        =   15478
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
      Begin VB.Label Label4 
         Caption         =   "Status"
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Barang diganti"
         Height          =   255
         Left            =   -67200
         TabIndex        =   45
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Baram 
         Caption         =   "Barang Masuk"
         Height          =   255
         Left            =   -74880
         TabIndex        =   44
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Cari nama barang"
         Height          =   255
         Left            =   -74520
         TabIndex        =   35
         Top             =   6360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Cari nomor servis/nama pelanggan"
         Height          =   375
         Left            =   -74520
         TabIndex        =   34
         Top             =   4800
         Width           =   2895
      End
   End
   Begin VB.Menu mn1 
      Caption         =   "Set status Perbaikan"
      Visible         =   0   'False
      Begin VB.Menu mnst1 
         Caption         =   "Set Status 'perbaikan'"
      End
      Begin VB.Menu mnst2 
         Caption         =   "Set status 'selesai'"
      End
   End
End
Attribute VB_Name = "frmproses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ubah As Boolean
Dim sum, hb, ttlhpp As Currency
Dim tmbh As String
Sub ttlb()
sum1 = 0
For I = 1 To lv1.ListItems.count
sum1 = sum1 + val(lv1.ListItems(I).SubItems(4)) * val(lv1.ListItems(I).SubItems(7))
Next I
ttlhpp = sum1
End Sub

Private Sub cmbjns_Click()
Set RS = New Recordset
RS.Open "select biaya_servis from servis where nama_servis='" & cmbjns.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
txtbiaya.Text = "0"
Else
txtbiaya.Text = RS!biaya_servis
Command2.SetFocus
End If
End Sub

Private Sub cmbjns_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

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

Private Sub cmbteknisi_Click()
txtlama.SetFocus
End Sub

Private Sub cmbteknisi_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

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
        l.SubItems(14) = ![perkiraan]
        l.SubItems(15) = Format(![tgl_out], "dd-mmm-yyyy")
        
        
       

    .MoveNext
    Loop
End With
End Sub
Private Sub listeknisi()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbteknisi.Clear
sql = "select nama_teknisi from teknisi order by nama_teknisi"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbteknisi.AddItem rsbarang!nama_teknisi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

  End Sub
Private Sub lisservis()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbjns.Clear
sql = "select nama_servis from servis order by nama_servis"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbjns.AddItem rsbarang!nama_servis
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

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
        l.SubItems(14) = ![perkiraan]
        l.SubItems(15) = Format(![tgl_out], "dd-mmm-yyyy")
        
        

    .MoveNext
    Loop
End With
End Sub

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

Private Sub cmbmatu_Click()
txtharga.SetFocus
matuc.Caption = cmbmatu.Text
End Sub

Private Sub cmbskin_Click()
Skinpath = App.Path & "\skin\" & cmbskin.Text & ".skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub

Private Sub cmdbatal_Click()
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
If cmbteknisi.Text = "" Or txtlama.Text = "" Then Exit Sub

If MsgBox("Update data?", vbYesNo, judul) = vbNo Then Exit Sub

Set RS = New Recordset
RS.Open "select id_teknisi from teknisi where nama_teknisi='" & cmbteknisi.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Teknisi tidak terdaftar", vbCritical
Exit Sub
Else
idteknisi = RS!id_teknisi
End If
Set RS = New Recordset

RS.Open "select kode_servis from servis where nama_servis='" & cmbjns.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Jenis Servis tidak terdaftar", vbCritical
Exit Sub
Else
idservis = RS!kode_servis
End If
jual.Execute "update tservis set perkiraan='" & txtlama.Text & "',kerusakan='" & txtrusak.Text & "',kode_servis='" & idservis & "',id_teknisi='" & idteknisi & "',biaya_servis='" & Format(txtbiaya.Text, Number) & "',biaya_brg='" & sum & "',status_servis='tunggu konfirmasi',hpp='" & ttlhpp & "' where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
For I = 1 To lv1.ListItems.count
jual.Execute "insert into tservis_dtl2 values('" & lvpaket.SelectedItem.SubItems(1) & "','" & lv1.ListItems(I).SubItems(1) & "','" & lv1.ListItems(I).SubItems(3) & "','" & lv1.ListItems(I).SubItems(4) & "','" & lv1.ListItems(I).SubItems(5) & "','" & lv1.ListItems(I).SubItems(6) & "','belum')"
Next I
kosong
lv1.ListItems.Clear
MsgBox "data berhasil diupdate", vbInformation, judul

dbgrid
Tab1.Tab = 1
End Sub
Sub ttl()
sum = 0
For I = 1 To lv1.ListItems.count
sum = sum + lv1.ListItems(I).SubItems(6)
Next I
End Sub


Sub awal()
frame.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
End Sub



Private Sub cmdubah_Click()
If lvpaket.ListItems.count = 0 Then Exit Sub
With lvpaket.SelectedItem
If .SubItems(7) <> "daftar" Then
MsgBox "Sudah diproses", vbCritical, judul
Exit Sub
End If

End With
Tab1.Tab = 0
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
frame.Enabled = True

End Sub

Private Sub Command1_Click()
If Not lv1.SelectedItem Is Nothing Then
lv1.ListItems.Remove lv1.SelectedItem.Index
End If
For I = 1 To lv1.ListItems.count
    lv1.ListItems(I).Text = I
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
listeknisi
lisservis
awal
cmbstts2.AddItem "Semua"
cmbstts2.AddItem "daftar"
cmbstts2.AddItem "tunggu konfirmasi"
cmbstts2.AddItem "setuju"
cmbstts2.AddItem "perbaikan"
cmbstts2.AddItem "selesai"
cmbstts2.AddItem "diambil"
cmbstts2.Text = "Semua"

Tab1.Tab = 1
    Skinpath = App.Path & "\skin\chizh.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Sub isigrid()
On Error Resume Next
Set cari = lv1.FindItem(txtkode2.Text, 1, , 1)
lv1.SelectedItem = cari

If Not cari Is Nothing Then
MsgBox "Barang sudah terdaftar"
txtfst.SetFocus
Exit Sub
End If

    Set butir = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
    With butir

    .SubItems(1) = txtkode2.Text
    .SubItems(2) = txtnama2.Text
    .SubItems(3) = val(Format(txtharga.Text, Number))
    .SubItems(4) = val(Format(txtjum.Text, Number))
    .SubItems(5) = val(Format(txtdisk.Text, Number))
    .SubItems(6) = val(Format(txtharga.Text, Number)) * val(Format(txtjum.Text)) - val(Format(txtdisk.Text, Number))
    .SubItems(7) = hb
    
    

 End With
 ttl
 ttlb
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
txtlama.Text = ""
txtrusak.Text = ""
cmbteknisi.Text = ""
cmbjns.Text = ""
txtbiaya.Text = ""
txtharga.Text = ""
lv1.ListItems.Clear
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

Private Sub mnst1_Click()
If lvpaket.SelectedItem.SubItems(7) <> "setuju" Then Exit Sub

If MsgBox("Proses ke perbaikan?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "update tservis set status_servis='perbaikan' where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
MsgBox "Update status berhasil", vbInformation, judul
dbgrid
End Sub

Private Sub mnst2_Click()
On Error Resume Next
If lvpaket.SelectedItem.SubItems(7) <> "perbaikan" Then Exit Sub

If MsgBox("Proses ke status selesai perbaikan?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "update tservis set status_servis='selesai' where no_servis='" & lvpaket.SelectedItem.SubItems(1) & "'"
MsgBox "Update status berhasil", vbInformation, judul
Dbgrid2
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
On Error Resume Next
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


lv1.ListItems.Clear
Set rsd = New Recordset
rsd.Open "select p.*,deskripsi,satuan from paket_detil p,tblbarang t where t.kode_brg=p.kode_brg and kode_pkt='" & txtkode.Text & "' order by deskripsi", jual, adOpenStatic, adLockOptimistic

With rsd
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
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
RS.Open "select deskripsi,harga_jual,harga_beli from tblbarang where kode_brg='" & txtkode2.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Barang tidak terdaftar", vbCritical, judul
Else
hb = RS!harga_beli
txtnama2.Text = RS!deskripsi
txtharga.Text = RS!harga_jual
txtjum.SetFocus
End If
End If
End Sub

