VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form pembelian 
   Caption         =   "Pembelian"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   9180
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   16193
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Pembelian"
      TabPicture(0)   =   "pembelian.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SkinLabel22"
      Tab(0).Control(1)=   "Cmdhapus"
      Tab(0).Control(2)=   "Cmdproses"
      Tab(0).Control(3)=   "Cmdkeluar"
      Tab(0).Control(4)=   "Cmdsimpan"
      Tab(0).Control(5)=   "cmdbatal"
      Tab(0).Control(6)=   "LV1"
      Tab(0).Control(7)=   "ket"
      Tab(0).Control(8)=   "Frame"
      Tab(0).Control(9)=   "Frame2"
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(11)=   "Frame3"
      Tab(0).Control(12)=   "cmdtambah"
      Tab(0).Control(13)=   "kete"
      Tab(0).Control(14)=   "Check1"
      Tab(0).Control(15)=   "ListView5"
      Tab(0).Control(16)=   "Check2"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Data obat"
      TabPicture(1)   =   "pembelian.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "ktr"
      Tab(1).Control(2)=   "txtcari"
      Tab(1).Control(3)=   "SkinLabel30"
      Tab(1).Control(4)=   "dbgrid1"
      Tab(1).Control(5)=   "kolom2"
      Tab(1).Control(6)=   "kolom"
      Tab(1).Control(7)=   "lvbrg"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Data PO"
      TabPicture(2)   =   "pembelian.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text10"
      Tab(2).Control(1)=   "Skin1"
      Tab(2).Control(2)=   "ListView3"
      Tab(2).Control(3)=   "ListView4"
      Tab(2).Control(4)=   "SkinLabel20"
      Tab(2).Control(5)=   "Label2"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Data supplier"
      TabPicture(3)   =   "pembelian.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvplg"
      Tab(3).Control(1)=   "Cmdcari"
      Tab(3).Control(2)=   "Dbgrid2"
      Tab(3).Control(3)=   "SkinLabel21"
      Tab(3).Control(4)=   "txtcrp"
      Tab(3).Control(5)=   "Label3"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Data Pembelian"
      TabPicture(4)   =   "pembelian.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "SkinLabel23"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "ListView2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "ListView1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdhps"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Command4"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Command1"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Text9"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Command3"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "CrystalReport1"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4320
         Top             =   3720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cetak faktur"
         Height          =   375
         Left            =   1800
         TabIndex        =   101
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "No dokumen otomatis"
         Height          =   255
         Left            =   -66960
         TabIndex        =   100
         Top             =   960
         Width           =   2175
      End
      Begin MSComctlLib.ListView lvplg 
         Height          =   6015
         Left            =   -74760
         TabIndex        =   86
         Top             =   480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   10610
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id supplier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama "
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Contact person"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Alamat"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "No.Telp"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvbrg 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   82
         Top             =   600
         Width           =   11055
         _ExtentX        =   19500
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
            Text            =   "Kode obat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama obat"
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
            Text            =   "Harga beli"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "HArga jual Grosir"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Supplier"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   81
         Top             =   2280
         Visible         =   0   'False
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4683
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
            Text            =   "Jumlah beli"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   -71400
         TabIndex        =   80
         Top             =   7080
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   3120
         TabIndex        =   72
         Top             =   7680
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   6000
         TabIndex        =   76
         Top             =   7680
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   13200
         TabIndex        =   71
         Top             =   4680
         Width           =   1335
      End
      Begin VB.CommandButton cmdhps 
         Caption         =   "&Hapus"
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   70
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No faktur otomatis"
         Height          =   255
         Left            =   -72840
         TabIndex        =   69
         Top             =   960
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel kete 
         Height          =   375
         Left            =   -70440
         OleObjectBlob   =   "pembelian.frx":008C
         TabIndex        =   66
         Top             =   8760
         Width           =   3855
      End
      Begin VB.ComboBox kolom 
         Height          =   315
         Left            =   -74760
         TabIndex        =   65
         Top             =   5640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin apotekbaleendah.ThemedButton cmdtambah 
         Height          =   375
         Left            =   -74760
         TabIndex        =   0
         Top             =   4920
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Tambah"
         font            =   "pembelian.frx":00EA
         forecolor       =   0
         mouseicon       =   "pembelian.frx":0116
      End
      Begin VB.CommandButton Cmdcari 
         Caption         =   "&Cari supplier"
         Height          =   255
         Left            =   -70920
         TabIndex        =   56
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -67920
         TabIndex        =   40
         Top             =   5880
         Width           =   3735
         Begin XPControls.XPText ttlppnt 
            Height          =   285
            Left            =   1800
            TabIndex        =   97
            Top             =   1680
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":06B0
            TabIndex        =   96
            Top             =   1680
            Width           =   1455
         End
         Begin XPControls.XPText XPText6 
            Height          =   285
            Left            =   1800
            TabIndex        =   54
            Top             =   1320
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
            Locked          =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0714
            TabIndex        =   53
            Top             =   1320
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":079A
            TabIndex        =   52
            Top             =   960
            Width           =   1335
         End
         Begin XPControls.XPText XPText2 
            Height          =   285
            Left            =   1800
            TabIndex        =   51
            Top             =   960
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
            Locked          =   -1  'True
         End
         Begin XPControls.XPText bayar 
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   2040
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0810
            TabIndex        =   47
            Top             =   2040
            Width           =   1335
         End
         Begin XPControls.XPText text8 
            Height          =   285
            Left            =   1800
            TabIndex        =   44
            Top             =   600
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
            Locked          =   -1  'True
         End
         Begin XPControls.XPText text7 
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Top             =   240
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
            Locked          =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0882
            TabIndex        =   42
            Top             =   600
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":08F6
            TabIndex        =   41
            Top             =   240
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":096C
            TabIndex        =   58
            Top             =   2400
            Visible         =   0   'False
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   255
            Left            =   1800
            TabIndex        =   59
            Top             =   2400
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   163184643
            CurrentDate     =   40299
         End
      End
      Begin VB.ComboBox kolom2 
         Height          =   315
         Left            =   -72000
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Height          =   2415
         Left            =   -70560
         TabIndex        =   27
         Top             =   5880
         Width           =   2535
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":09E0
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0A50
            TabIndex        =   29
            Top             =   720
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0AC2
            TabIndex        =   30
            Top             =   1200
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel hj 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "pembelian.frx":0B34
            TabIndex        =   31
            Top             =   1200
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel hb 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "pembelian.frx":0B92
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "pembelian.frx":0BF0
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -74760
         OleObjectBlob   =   "pembelian.frx":0C4E
         Top             =   8220
      End
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   -74760
         TabIndex        =   24
         Top             =   5280
         Width           =   4095
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0E82
            TabIndex        =   95
            Top             =   3120
            Width           =   1695
         End
         Begin XPControls.XPText txtppn 
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            Top             =   2760
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
            Locked          =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0EEC
            TabIndex        =   94
            Top             =   2760
            Width           =   1455
         End
         Begin XPControls.XPText txtbatch 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   1320
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0F50
            TabIndex        =   93
            Top             =   1320
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":0FBE
            TabIndex        =   92
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cmbgudang 
            Height          =   315
            Left            =   1800
            TabIndex        =   90
            Top             =   960
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel label4 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "pembelian.frx":1028
            TabIndex        =   68
            Top             =   1680
            Width           =   735
         End
         Begin XPControls.XPCombo satuan 
            Height          =   315
            Left            =   1800
            TabIndex        =   8
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":1086
            TabIndex        =   67
            Top             =   600
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":10F0
            TabIndex        =   35
            Top             =   2040
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":115A
            TabIndex        =   26
            Top             =   1680
            Width           =   735
         End
         Begin XPControls.XPText text6 
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Top             =   2040
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
         Begin XPControls.XPText text5 
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   1680
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
         Begin XPControls.XPCombo text4 
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   240
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":11CC
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
         Begin XPControls.XPText XPText4 
            Height          =   285
            Left            =   2280
            TabIndex        =   15
            Top             =   2400
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
         Begin XPControls.XPText XPText3 
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   2400
            Width           =   375
            _ExtentX        =   661
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":123E
            TabIndex        =   48
            Top             =   2400
            Width           =   1095
         End
         Begin XPControls.XPText XPText5 
            Height          =   285
            Left            =   1800
            TabIndex        =   49
            Top             =   3480
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
            Locked          =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "pembelian.frx":12B6
            TabIndex        =   50
            Top             =   3480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   1800
            OleObjectBlob   =   "pembelian.frx":1326
            TabIndex        =   91
            Top             =   720
            Width           =   1455
         End
         Begin XPControls.XPText txtppnp 
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   2760
            Width           =   375
            _ExtentX        =   661
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
            TabIndex        =   14
            Top             =   3120
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
            CheckBox        =   -1  'True
            CustomFormat    =   "dd MMM yyyy"
            Format          =   163446787
            CurrentDate     =   37623
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   23
         Top             =   2400
         Visible         =   0   'False
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5953
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
      Begin VB.Frame Frame 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   19
         Top             =   360
         Width           =   10215
         Begin XPControls.XPText txtdok 
            Height          =   285
            Left            =   7800
            TabIndex        =   2
            Top             =   240
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
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   255
            Left            =   4680
            TabIndex        =   89
            Top             =   1440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "pembelian.frx":139A
            TabIndex        =   45
            Top             =   1440
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "pembelian.frx":1414
            TabIndex        =   34
            Top             =   960
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   7800
            TabIndex        =   5
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   163446787
            CurrentDate     =   40299
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   3
            Top             =   960
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
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   1
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "pembelian.frx":148A
            TabIndex        =   20
            Top             =   1320
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "pembelian.frx":1502
            TabIndex        =   21
            Top             =   960
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel Nomor 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "pembelian.frx":1582
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   7800
            TabIndex        =   6
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   163446787
            CurrentDate     =   40299
         End
         Begin XPControls.XPCombo text3 
            Height          =   315
            Left            =   1920
            TabIndex        =   4
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
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
            Locked          =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "pembelian.frx":15F2
            TabIndex        =   98
            Top             =   240
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket 
         Height          =   375
         Left            =   -74880
         OleObjectBlob   =   "pembelian.frx":1664
         TabIndex        =   39
         Top             =   9960
         Width           =   10695
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   46
         Top             =   2160
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
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
         NumItems        =   13
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
            Text            =   "Harga beli"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jumlah"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Diskon"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Sub total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "kode gdg"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "batch"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ppn"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "exp"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Dbgrid2 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   55
         Top             =   600
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8916
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
      Begin apotekbaleendah.ThemedButton cmdbatal 
         Height          =   375
         Left            =   -73440
         TabIndex        =   60
         Top             =   4920
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Batal"
         font            =   "pembelian.frx":16C2
         forecolor       =   0
         mouseicon       =   "pembelian.frx":16EE
      End
      Begin apotekbaleendah.ThemedButton Cmdsimpan 
         Height          =   375
         Left            =   -68160
         TabIndex        =   61
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Simpan"
         font            =   "pembelian.frx":1C88
         forecolor       =   0
         mouseicon       =   "pembelian.frx":1CB4
      End
      Begin apotekbaleendah.ThemedButton Cmdkeluar 
         Height          =   375
         Left            =   -72120
         TabIndex        =   62
         Top             =   4920
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Keluar"
         font            =   "pembelian.frx":224E
         forecolor       =   0
         mouseicon       =   "pembelian.frx":227A
      End
      Begin apotekbaleendah.ThemedButton Cmdproses 
         Height          =   375
         Left            =   -70680
         TabIndex        =   63
         Top             =   4920
         Visible         =   0   'False
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Proses"
         font            =   "pembelian.frx":2814
         forecolor       =   0
         mouseicon       =   "pembelian.frx":2840
      End
      Begin apotekbaleendah.ThemedButton Cmdhapus 
         Height          =   375
         Left            =   -65880
         TabIndex        =   64
         Top             =   4920
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Hapus"
         font            =   "pembelian.frx":2DDA
         forecolor       =   0
         mouseicon       =   "pembelian.frx":2E06
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   120
         TabIndex        =   73
         Top             =   4080
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6165
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
         NumItems        =   9
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
            Text            =   "Kode obat"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama obat"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Harga beli"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jumlah"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Sub total"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No transaksi"
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
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Supplier"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "pembelian.frx":33A0
         TabIndex        =   75
         Top             =   7680
         Width           =   2775
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3015
         Left            =   -73920
         TabIndex        =   77
         Top             =   600
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   3135
         Left            =   -74520
         TabIndex        =   78
         Top             =   3720
         Width           =   10575
         _ExtentX        =   18653
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
            Text            =   "Kode obat"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama obat"
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
            Text            =   "Jumlah Beli"
            Object.Width           =   2540
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   375
         Left            =   -74040
         OleObjectBlob   =   "pembelian.frx":347E
         TabIndex        =   79
         Top             =   7080
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
         Height          =   375
         Left            =   -74520
         OleObjectBlob   =   "pembelian.frx":3524
         TabIndex        =   83
         Top             =   6600
         Width           =   2295
      End
      Begin XPControls.XPText txtcari 
         Height          =   375
         Left            =   -70800
         TabIndex        =   84
         Top             =   6480
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
      Begin ACTIVESKINLibCtl.SkinLabel ktr 
         Height          =   375
         Left            =   -68640
         OleObjectBlob   =   "pembelian.frx":35CE
         TabIndex        =   85
         Top             =   6480
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   375
         Left            =   -74640
         OleObjectBlob   =   "pembelian.frx":362C
         TabIndex        =   87
         Top             =   6720
         Width           =   2295
      End
      Begin XPControls.XPText txtcrp 
         Height          =   375
         Left            =   -72240
         TabIndex        =   88
         Top             =   6720
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   375
         Left            =   -67560
         OleObjectBlob   =   "pembelian.frx":36BE
         TabIndex        =   99
         Top             =   5520
         Width           =   3375
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   -66840
         TabIndex        =   57
         Top             =   8280
         Width           =   3255
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   -66480
         TabIndex        =   38
         Top             =   7920
         Width           =   2775
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   -66480
         TabIndex        =   37
         Top             =   8880
         Width           =   2655
      End
   End
End
Attribute VB_Name = "pembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode, nbrg, id, mu, mu2, ket1, kete2, stn_baku, stri, stn, kodegdg As String
Dim st, cosbli, cosju, sel As Currency
Dim po, nutupbeli, ada As Boolean
Dim dey As Integer
Sub dbgridplg()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblsupplier order by supplier"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_supplier]
        l.SubItems(2) = ![Supplier]
                                l.SubItems(3) = ![Kontak_person]

                l.SubItems(4) = ![alamat]
                
                l.SubItems(5) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub
Private Sub datagdg()

cmbgudang.Clear
sql = "select nama_gudang from gudang,stokgudang where gudang.kode_gudang=stokgudang.kode_gudang and kode_brg='" & kode & "' group by stokgudang.kode_gudang order by nama_gudang"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
cmbgudang.AddItem rsplg!nama_gudang
rsplg.MoveNext
 Loop
  rsplg.MoveFirst
  End If

rsplg.Close
cmbgudang.Text = "utama"
  End Sub
Sub dbgridplg2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblsupplier where id_supplier like '" & txtcrp.Text & "%' or supplier like '%" & txtcrp.Text & "%' order by supplier"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_supplier]
        l.SubItems(2) = ![Supplier]
                                l.SubItems(3) = ![Kontak_person]

                l.SubItems(4) = ![alamat]
                
                l.SubItems(5) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub

Private Sub Check2_Click()
SaveSetting "apotekbaleendah", "pembelian", "Check2.value", Check2.Value
End Sub

Private Sub Command2_Click()
Tab1.Tab = 3
End Sub

Private Sub Command3_Click()
'On Error Resume Next
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\fakturbeli.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{pembelian.no_pembelian}='" & ListView2.SelectedItem.SubItems(1) & "'"
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

Private Sub lvbrg_DblClick()
On Error Resume Next
If Edit = False Then
If text3.Text = "" Then
    If lvbrg.SelectedItem.SubItems(8) <> "" Then
    text3.Text = lvbrg.SelectedItem.SubItems(8)
    End If
Else
    If lvbrg.SelectedItem.SubItems(8) <> "" Then

    If lvbrg.SelectedItem.SubItems(8) <> text3.Text Then
MsgBox "Barang yang dipilih salah supplier,jika rubah supplier harus diganti di master barang dahulu", vbInformation, judul
Exit Sub
Else
End If
End If

End If
text4.Text = lvbrg.SelectedItem.SubItems(1)
Tab1.Tab = 0
text4_Click
Else
MsgBox "Tekan tombol baru dulu"

End If

End Sub

Private Sub lvbrg_GotFocus()
ktr.Caption = "Dobel klik atau enter untuk mengirimkan data"

End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub

Sub dbgrid()
On Error Resume Next

Set rstrans = New Recordset


sql = "select t.*,supplier from tblbarang t left join tblsupplier s on t.id_supplier=s.id_supplier order by deskripsi"
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
                l.SubItems(6) = Format(![harga_beli], "#,#")
                l.SubItems(7) = Format(![Harga_jual2], "#,#")
                l.SubItems(8) = ![Supplier]

    .MoveNext
    Loop
End With


End Sub

Sub dbgridcari()
On Error Resume Next
Set rstrans = New Recordset
stri = Replace(txtcari.Text, "'", "''")


sql = "select t.*,supplier from tblbarang t left join tblsupplier s on t.id_supplier=s.id_supplier where kode_brg like '" & stri & "%' or deskripsi like '%" & stri & "%' or supplier like '%" & stri & "%' order by deskripsi"
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
                l.SubItems(6) = Format(![harga_beli], "#,#")
  l.SubItems(7) = Format(![Harga_jual2], "#,#")
                l.SubItems(8) = ![Supplier]

    .MoveNext
    Loop
End With

End Sub

Sub dbgridpo()
On Error Resume Next

Set rstrans = New Recordset


sql = "select pemesanan.no_pemesanan,pemesanan.tanggal_pesan,pemesanan.id_supplier,tblsupplier.supplier from pemesanan,tblsupplier where pemesanan.id_supplier=tblsupplier.id_supplier order by no_pemesanan"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_pemesanan]
        l.SubItems(2) = Format(![tanggal_pesan], "dd MMM yyyy")
                                l.SubItems(3) = ![id_supplier]

                l.SubItems(4) = ![Supplier]

    .MoveNext
    Loop
End With

End Sub

Sub dbgridtrans()
On Error Resume Next

Set rstrans = New Recordset


sql = "select pembelian.no_pembelian,pembelian.tanggal_pembelian,pembelian.total,pembelian.total_diskon,pembelian.total_stlh_diskon,pembelian.id_supplier,tblsupplier.supplier from pembelian,tblsupplier where pembelian.id_supplier=tblsupplier.id_supplier order by tanggal_pembelian desc,no_pembelian desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_pembelian]
        l.SubItems(2) = Format(![tanggal_pembelian], "dd MMM yyyy")
                                l.SubItems(3) = ![id_supplier]

                l.SubItems(4) = ![Supplier]
                l.SubItems(5) = Format(![total], "#,#")
                l.SubItems(6) = Format(![total_diskon], "#,#")
  l.SubItems(7) = Format(![total_stlh_diskon], "#,#")

    .MoveNext
    Loop
End With

End Sub
Sub dbgridpo2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select pemesanan.no_pemesanan,pemesanan.tanggal_pesan,pemesanan.id_supplier,tblsupplier.supplier from pemesanan,tblsupplier where pemesanan.id_supplier=tblsupplier.id_supplier and (pemesanan.no_pemesanan like '" & Text10.Text & "%' or  tblsupplier.supplier like '%" & Text10.Text & "%')  order by no_pemesanan"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_pemesanan]
        l.SubItems(2) = Format(![tanggal_pesan], "dd MMM yyyy")
                                l.SubItems(3) = ![id_supplier]

                l.SubItems(4) = ![Supplier]

    .MoveNext
    Loop
End With

End Sub

Sub dbgridtrans2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select pembelian.no_pembelian,pembelian.tanggal_pembelian,pembelian.total,pembelian.total_diskon,pembelian.total_stlh_diskon,pembelian.id_supplier,tblsupplier.supplier from pembelian,tblsupplier where pembelian.id_supplier=tblsupplier.id_supplier and (pembelian.no_pembelian like'" & text9.Text & "%' or pembelian.no_pemesanan like'" & text9.Text & "%' or pembelian.id_supplier like'" & text9.Text & "%' or tblsupplier.supplier like'" & text9.Text & "%') order by no_pembelian desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_pembelian]
        l.SubItems(2) = Format(![tanggal_pembelian], "dd MMM yyyy")
                                l.SubItems(3) = ![id_supplier]

                l.SubItems(4) = ![Supplier]
                l.SubItems(5) = Format(![total], "#,#")
                l.SubItems(6) = Format(![total_diskon], "#,#")
  l.SubItems(7) = Format(![total_stlh_diskon], "#,#")

    .MoveNext
    Loop
End With

End Sub

Private Sub bayar_GotFocus()
kete.Caption = "Tekan enter bila pembayaran nol,atau masukkan berapapun sebagian yang baru dibayar,atau ketikkan jumlah yang sama bila dibayar lunas."
ListView5.Visible = False
End Sub

Private Sub bayar_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
pembayaran = val(bayar.Text)
priksa
End If
End Sub

Private Sub bayar_LostFocus()
kete.Caption = ""
End Sub

Private Sub Check1_Click()
    SaveSetting "apotekbaleendah", "pembelian", "Check1.value", Check1.Value

End Sub

Private Sub Cmdbatal_Click()
awal
kosong
cmdtambah.SetFocus
Label2.Caption = ""
Label1.Caption = ""
Label3.Caption = ""
End Sub
Private Sub priksa()
sel = val(bayar.Text) - val(XPText6.Text)
If sel > 0 Then
MsgBox "Pembayaran jangan melebihi total pembelian", vbInformation, judul
bayar.Text = ""
bayar.SetFocus
Exit Sub
End If
If (0 <= val(bayar.Text) < val(XPText6.Text)) Then
ket1 = "B"
kete2 = "BL"
If val(bayar.Text) < val(XPText6.Text) Then
MsgBox "Hutang akan bertambah " & Format(-1 * sel, "Rp#,#0.#0") & " ", vbInformation
End If
SkinLabel18.Visible = True
DTPicker3.Visible = True
DTPicker3.SetFocus
Else
If val(bayar.Text) = val(XPText6.Text) Then
ket1 = "C"
kete2 = "L"
Else
ket1 = "CB"
kete2 = "BL"
End If
End If


If sel = 0 Then
If MsgBox("Simpan data?", vbYesNo) = vbYes Then
pilih = ""
nutupbeli = False
'mulai
pilih = "KAS"
Cmdsimpan_Click
nutupbeli = True
'akhir
    'tanya.Show

End If
End If

End Sub
Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
Frame.Enabled = True
Frame2.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True

End Sub

Private Sub Cmdcari_Click()
Dim kata As String

If Cmdcari.Caption = "&Cari supplier" Then
kata = InputBox("Masukkan id supplier atau nama suplier", "Cari...")
If kata = "" Then Exit Sub
sql = "select* from tblsupplier where id_supplier='" & kata & "' or supplier like '%" & kata & "%' "
Set rssupp = New Recordset
Set rssupp = jual.Execute(sql)

If Not rssupp.EOF Then
Set Dbgrid2.DataSource = rssupp
Cmdcari.Caption = "&Refresh"
Else
MsgBox "Tidak ada", vbOKOnly, judul
dbgrids
End If
Else
dbgrids
Cmdcari.Caption = "&Cari supplier"
End If

End Sub

Private Sub cmdcari2_Click()
Dim kata As String

If cmdcari2.Caption = "&Cari PO" Then
kata = InputBox("Masukkan no pemesanan", "Cari...")
If kata = "" Then Exit Sub
sql = "select* from tblpemesanan where no_pemesanan='" & kata & "'"
Set rspo = New Recordset
Set rspo = jual.Execute(sql)

If Not rspo.EOF Then
Set dbgrid3.DataSource = rspo
cmdcari2.Caption = "&Refresh"
Else
MsgBox "Tidak ada", vbOKOnly, judul
dbgridpo
End If
Else
dbgridpo
cmdcari2.Caption = "&Cari PO"
End If

End Sub

Private Sub cmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
ttlppn
diskon
ttl2
ttl_item
For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub cmdhps_Click()
If ListView1.ListItems.count = 0 Then Exit Sub

If MsgBox("Yakin akan menghapus no faktur " & ListView2.SelectedItem.SubItems(1) & " dan merubah segala transakasi yang berhubungan dengan no faktur ini?", vbYesNo, jdul) = vbNo Then Exit Sub

hapus_faktur
dbgridtrans
ListView1.ListItems.Clear

End Sub
Sub hapus_faktur()
jual.Execute "delete from pembelian where no_pembelian='" & ListView2.SelectedItem.SubItems(1) & "'"

End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdproses_Click()
On Error Resume Next
If text4.Text <> "" Then
text6_KeyPress (13)
End If
End Sub

Private Sub Cmdproses_GotFocus()
ket.Caption = "Tekan enter/klick tombol proses untuk memproses dan memasukkan ke dalam tabel"
End Sub

Private Sub Cmdproses_LostFocus()
ket.Caption = ""
End Sub
Private Sub Cmdsimpan_Click()
'On Error GoTo erol
Dim cekharga As Currency

simpandata
dbgrid
dbgridtrans
SkinLabel18.Visible = False
DTPicker3.Visible = False

Label2.Caption = ""
Label1.Caption = ""
Label3.Caption = ""

For x = 1 To LV1.ListItems.count
Set RS = New Recordset
RS.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(x).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic

If RS!satuan = LV1.ListItems(x).SubItems(4) Then

cekharga = LV1.ListItems(x).SubItems(5)
Else
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.ListItems(x).SubItems(2) & "' and satuan='" & LV1.ListItems(x).SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic
cekharga = val(LV1.ListItems(x).SubItems(5)) / rsstn!konversi
End If

If val(RS!harga_beli) <> val(cekharga) Then
 If MsgBox("Terjadi perbedaan harga beli " & RS!deskripsi & " / " & RS!satuan & "" & vbCrLf & _
  vbCrLf & "Harga beli lama=" & Format(RS!harga_beli, "#,##0") & "" & vbCrLf & _
 "Harga beli baru=" & Format(cekharga, "#,##0") & "" & vbCrLf & _
  vbCrLf & "Memakai harga beli baru?", vbYesNo) = vbYes Then
 RS.Fields("harga_beli") = cekharga
 jual.Execute "update tblbarang set harga_beli=" & cekharga & " where kode_brg='" & LV1.ListItems(x).SubItems(2) & "'"
  If MsgBox("Harga ecer lama=" & Format(RS!harga_jual, "#,##0") & "" & vbCrLf & "Rubah harga ecer?", vbYesNo) = vbYes Then
  HJbr = InputBox("Harga ecer lama=" & Format(RS!harga_jual, "#,##0") & "" & vbCrLf & "Masukkan harga ecer baru:")
  jual.Execute "update tblbarang set harga_jual=" & HJbr & " where kode_brg='" & LV1.ListItems(x).SubItems(2) & "'"
  End If
  If MsgBox("Harga grosir lama=" & Format(RS!Harga_jual2, "#,##0") & "" & vbCrLf & "Rubah harga grosir?", vbYesNo) = vbYes Then
  HJbr = InputBox("Harga grosir lama=" & Format(RS!Harga_jual2, "#,##0") & "" & vbCrLf & "Masukkan harga grosir baru:")
  jual.Execute "update tblbarang set harga_jual2=" & HJbr & " where kode_brg='" & LV1.ListItems(x).SubItems(2) & "'"
  End If
    If MsgBox("Harga resep lama=" & Format(RS!Harga_jual3, "#,##0") & "" & vbCrLf & "Rubah harga resep?", vbYesNo) = vbYes Then
  HJbr = InputBox("Harga grosir lama=" & Format(RS!Harga_jual3, "#,##0") & "" & vbCrLf & "Masukkan harga resep baru:")
  jual.Execute "update tblbarang set harga_jual3=" & HJbr & " where kode_brg='" & LV1.ListItems(x).SubItems(2) & "'"
  End If

 End If
 End If
 Next x

Exit Sub

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
End If


End Sub



Sub simpandata()
'On Error GoTo erol
Set RS = New Recordset
sql = "select id_supplier from tblsupplier where supplier='" & text3.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
id = RS!id_supplier
RS.Close
Else
 MsgBox "Pilih  Supplier dulu"
 text3.SetFocus
 Exit Sub
id = ""
End If
sql = "insert into pembelian(no_pembelian,no_pemesanan,tanggal_pembelian,id_supplier,cash,hari,no_dokumen) values('" & Text1.Text & "','" & Text2.Text & "','" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','" & id & "','" & val(Format(bayar.Text, Number)) & "','" & dey & "','" & txtdok.Text & "')"


jual.Execute (sql)

Set RS = New Recordset
For z = 1 To LV1.ListItems.count
Set RS = New Recordset

RS.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
ada = False
jual.Execute "insert into tblbarang(kode_brg,deskripsi,satuan,stok,harga_beli,harga_jual) values('" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(3) & "','" & LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(5) & "',0)"
jual.Execute "insert into satuan values('" & LV1.ListItems(z).SubItems(2) & "','" & LV1.ListItems(z).SubItems(4) & "',1,'Utama','" & LV1.ListItems(z).SubItems(6) & "')"
Else
ada = True
End If
Set RS = New Recordset


RS.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic

If LV1.ListItems(z).SubItems(4) <> RS!satuan Then
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.ListItems(z).SubItems(2) & "' and satuan='" & LV1.ListItems(z).SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic


sql = "insert into detilbeli values('" & Text1.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) / val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(6) * val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(11) & "','" & LV1.ListItems(z).SubItems(8) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(9) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(12) & "')"
jual.Execute (sql)

Else
ns = val(st) + LV1.ListItems(z).SubItems(6)
sql = "insert into detilbeli values('" & Text1.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(11) & "','" & LV1.ListItems(z).SubItems(8) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(9) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(12) & "')"
jual.Execute (sql)

End If

Set RS = New Recordset
RS.Open "Select * from detilpesan where no_pemesanan='" & Text2.Text & "' and kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
jual.Execute "update detilpesan set jumlah_beli=(jumlah_beli + " & val(LV1.ListItems(z).SubItems(6)) & ") where no_pemesanan='" & Text2.Text & "' and kode_brg='" & LV1.ListItems(z).SubItems(2) & "' "
End If
    Next z
    MsgBox "Data Pembelian sudah tersimpan", vbInformation
    pilih = ""
      awal
 cmdtambah.SetFocus

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
Frame.Enabled = True
End If


End Sub
Sub cmdtambah_Click()
Edit = False
dey = 0
tambah
kosong
kosong2
po = False
If Check2.Value = Checked Then
GetNumber2
End If
If Check1.Value = Checked Then
GetNumber
text3.SetFocus
Else
Text1.SetFocus

End If
End Sub
Sub GetNumber()
On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl As String
    A = hpb & "%"
sql = "Select no_pembelian from pembelian where no_pembelian like '" & A & "'order by no_pembelian "
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = hpb + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_pembelian], pjgh + 6) <> hpb + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = hpb + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_pembelian, 2)) + 1
               counter = hpb + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        Text1.Text = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description
End Sub
Sub GetNumber2()
On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl As String
    A = "ND%"
sql = "Select no_dokumen from pembelian where no_dokumen like '" & A & "'order by no_dokumen"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "ND" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_dokumen], 8) <> "ND" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "ND" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_dokumen, 2)) + 1
               counter = "ND" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        txtdok.Text = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description
End Sub
Sub tambah()
Edit = False
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Label1.Caption = "*Dobel klik untuk mengirimkan data"
Label2.Caption = "*Dobel klik untuk mengirimkan data"
Frame.Enabled = True
cmdbatal.Enabled = True
Frame2.Enabled = True
Label3.Caption = "*Dobel klik untuk mengirimkan data"
End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
sql = "Select no_penjualan from penjualan where no_penjualan like 'PB-%'order by no_penjualan Desc"
Set RS = jual.Execute(sql)
If RS.EOF = True Then
Text1.Text = "PB-" + Format(Now, "YY-") + Format(Now, "MM-") + Format(Now, "dd-") + "001"

Else
j = val(Right(RS(0), 3))
br = "PB-" + Format(Now, "YY-") + Format(Now, "MM") + "-" + Format(Now, "dd-") + Format(Str(j + 1), "000")
Text1.Text = br
No = br

nmr = Format(Mid(RS(0), 8, 2))
thn = Format(Mid(RS(0), 5, 2))
txt = Format(Mid(No, 8, 2))
thnn = Format(Mid(No, 5, 2))

If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
Text1.Text = "PB-" + Format(Now, "YY-") + Format(Now, "MM-") + Format(Now, "dd-") + "001"
End If
End If
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""
text7.Text = ""
XPText3.Text = ""
XPText4.Text = ""
XPText5.Text = ""
XPText2.Text = ""
text8.Text = ""
bayar.Text = ""
txtppn.Text = ""
txtppnp.Text = ""
XPText6.Text = ""
LV1.ListItems.Clear
txtdok.Text = ""
End Sub



Private Sub Command1_Click()
dbgridtrans
End Sub




Private Sub dbgrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dbgrid3_DblClick
End If
End Sub

Private Sub DTPicker3_GotFocus()
kete.Caption = "Tentukan tanggal jatuh tempo pembayaran,bila sudah tekan enter atau klik simpan."

End Sub

Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
dey = DTPicker3.Value - DTPicker2.Value
If MsgBox("Simpan data?", vbYesNo) = vbYes Then
pilih = ""
If val(bayar.Text) > 0 Then
nutupbeli = False
'mulai
pilih = "KAS"
Cmdsimpan_Click
nutupbeli = True
'akhir

    'tanya.Show
Else
Cmdsimpan_Click
End If
End If
End If

End Sub

Private Sub DTPicker3_LostFocus()
kete.Caption = ""
End Sub

Private Sub Form_Activate()
'kbrg

If pilih = "KAS" Or pilih = "BANK" Then
Cmdsimpan_Click
End If
If hpb = "" Then
hpb = "PB"
End If
pjgh = Len(hpb)
nutupbeli = True
End Sub

Private Sub Form_Deactivate()
If nutupbeli = True Then
Unload Me
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
Cmdsimpan_Click
Else
If KeyCode = vbKeyF2 Then
cmdtambah_Click
Else
If KeyCode = vbKeyF3 Then
Tab1.Tab = 1
txtcari.SetFocus
Else
If KeyCode = vbKeyF4 Then
Tab1.Tab = 3
txtcrp.SetFocus
Else
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pembelian.doc" _
                 , vbNullString, vbNullString, 1

End If
End If
End If

End If
End If
End Sub

Private Sub info_Click()
MsgBox "Pengisian kurs terakhir 1 US$=Rp" & matu & ""
End Sub

Private Sub TCmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
ttlppn
diskon
For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub ListView2_Click()
On Error Resume Next
text4.Enabled = True
If ListView2.ListItems.count = 0 Then Exit Sub

If ListView2.ListItems.count <> 0 Then
cmdhps.Enabled = True
End If
ListView1.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select detilbeli.kode_brg,tblbarang.deskripsi,detilbeli.satuan,detilbeli.harga_beli,detilbeli.jumlah_brg,detilbeli.diskon,detilbeli.total from detilbeli,tblbarang where detilbeli.no_pembelian='" & ListView2.SelectedItem.SubItems(1) & "' and tblbarang.kode_brg=detilbeli.kode_brg", jual, adOpenStatic, adLockOptimistic
If Not rse3.EOF Then

With rse3
.MoveFirst
    Do While Not .EOF
   Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
   l.SubItems(1) = ListView1.ListItems.count & "."
        l.SubItems(2) = !kode_brg
              l.SubItems(3) = ![deskripsi]
                            l.SubItems(4) = ![satuan]
Set rse4 = New Recordset
rse4.Open "select* from satuan where kode_brg='" & l.SubItems(2) & "'and satuan='" & l.SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic


              l.SubItems(5) = ![harga_beli] * rse4!konversi
                l.SubItems(6) = ![jumlah_brg] / rse4!konversi
              l.SubItems(7) = ![diskon]
                      l.SubItems(8) = ![total]
                      

    .MoveNext
    Loop
End With
End If
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
ListView2_Click

End If

End Sub

Private Sub ListView3_Click()
On Error Resume Next
Set rstrans = New Recordset


sql = "select detilpesan.kode_brg,tblbarang.deskripsi,detilpesan.jumlah_brg2,detilpesan.satuan,detilpesan.jumlah_beli from detilpesan,pemesanan,tblbarang where pemesanan.no_pemesanan=detilpesan.no_pemesanan and detilpesan.kode_brg=tblbarang.kode_brg and pemesanan.no_pemesanan='" & ListView3.SelectedItem.SubItems(1) & "'"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView4.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView4.ListItems.Add(, , ListView4.ListItems.count + 1)
           l.SubItems(1) = ListView4.ListItems.count & "."

        l.SubItems(2) = ![kode_brg]
        l.SubItems(3) = ![deskripsi]
            l.SubItems(5) = ![jumlah_brg2]

                l.SubItems(4) = ![satuan]
l.SubItems(6) = ![jumlah_beli]
    .MoveNext
    Loop
End With

End Sub

Private Sub ListView3_DblClick()
If ListView3.ListItems.count = 0 Then Exit Sub
Text2.Text = ListView3.SelectedItem.SubItems(1)
Text2_KeyPress (13)
Tab1.Tab = 0
End Sub

Private Sub ListView5_DblClick()
text4.Text = ListView5.SelectedItem.SubItems(2)
text4_Click
satuan.Text = ListView5.SelectedItem.SubItems(4)
text5.SetFocus
End Sub

Private Sub satuan_Change()
If satuan.Text = "" Then
Label4.Caption = ""
Else
Label4.Caption = "/" & satuan.Text
End If
End Sub

Private Sub satuan_Click()
text5.SetFocus
End Sub

Private Sub satuan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text5.SetFocus
End If
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)

If Tab1.Tab = 0 Then
cmdtambah.Caption = "&Tambah"
Else
If Tab1.Tab = 1 Then
txtcari.Text = ""
txtcari.SetFocus
ktr.Caption = ""
Else
If Tab1.Tab = 3 Then
txtcrp.Text = ""
txtcrp.SetFocus
End If
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text3.SetFocus
End If
End Sub

Private Sub Text10_Change()
dbgridpo2
End Sub

Private Sub Text2_Change()
If Text2.Text = "" Then
po = False
End If

End Sub

Private Sub Text2_GotFocus()
kete.Caption = "Isi nomor PO lalu enter"
End Sub

Private Sub text4_GotFocus()
kete.Caption = "Pilih barang,atau bisa diambil di tab data barang lalu dobel klik"
End Sub

Private Sub Text4_LostFocus()
kete.Caption = ""
End Sub

Private Sub Text5_Change()
XPText3_Change
End Sub

Private Sub Text5_GotFocus()
kete.Caption = "Tekan enter untuk menyesuaikan harga dengan harga di master barang,bila beda isi harga baru kemudian tekan tab"
End Sub

Sub kosong2()
text4.Text = ""
text5.Text = ""
text6.Text = ""
stok = ""
hb = ""
hj = ""
XPText3.Text = ""
XPText4.Text = ""
XPText5.Text = ""
txtppn.Text = ""
txtppnp.Text = ""
satuan.Text = ""
End Sub
Private Sub dbgrid1_DblClick()
On Error Resume Next
If Edit = False Then
text4.Text = dbgrid1.Columns(1)
Tab1.Tab = 0
text4_Click
text5.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus
End If
End Sub
Private Sub dbgrid3_DblClick()
On Error Resume Next

If Edit = False Then

Text2.Text = dbgrid3.Columns(0).Text
text3.Text = dbgrid3.Columns(2).Text
DTPicker1 = dbgrid3.Columns(1)
Tab1.Tab = 0
Text2_KeyPress (13)
DataGrid1.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus

End If
End Sub

Private Sub Form_Load()

 Check1.Value = GetSetting("apotekbaleendah", "pembelian", "Check1.value", Checked)
  Check2.Value = GetSetting("apotekbaleendah", "pembelian", "Check2.value", Checked)

 If hpb = "" Then
 hpb = "PB"
 End If
pjgh = Len(hpb)

Tab1.Tab = 0
pilih = ""
Edit = True
supp
dbgridplg
'kbrg
cmdhps.Enabled = IIf(rshusus.Fields(18) = False, False, True)
awal
Ketengah Me
kolom.Text = "Semua"
 dbgridpo
 dbgrids
 dbgrid
 datagdg
 dbgridtrans
 DTPicker1 = Format(Now)
  DTPicker2 = Format(Now)
  DTPicker3 = Format(Now + 14)
tgl.Value = Format(Now)

     Skinpath = App.Path & "\skin\green.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
kolom.Clear
kolom.AddItem "Semua"
kolom.AddItem "Deskripsi"
kolom.Text = "Semua"

End Sub

Private Sub kbrg()

text4.Clear
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text4.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub
  Private Sub tampil_stn()
On Error Resume Next
satuan.Clear
sql = "select * from satuan where kode_brg='" & text4.Text & "' order by satuan"
Set rsstn = New Recordset
Set rsstn = jual.Execute(sql)
If Not rsstn.EOF Then
rsstn.MoveFirst
 Do While Not rsstn.EOF
satuan.AddItem rsstn!satuan
rsstn.MoveNext
 Loop
  End If
rsstn.Close


  End Sub

Private Sub dbgrid2_DblClick()
On Error Resume Next

If Edit = False Then
text3.Text = Dbgrid2.Columns(1)
Tab1.Tab = 0
text4.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus

End If
End Sub

Sub dbgrids()
Set rssupp = New Recordset

sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set Dbgrid2.DataSource = rssupp


End Sub


Private Sub awal()
Frame.Enabled = False
Frame2.Enabled = False
Edit = True
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
kete.Caption = ""

End Sub

Private Sub hb_Click()
text5.Text = hb.Caption
text6.SetFocus
End Sub

Sub isigrid()
getkodegdg
    Set butir = LV1.ListItems.Add
    With butir
    .SubItems(1) = LV1.ListItems.count & "."
    .SubItems(2) = kode
    .SubItems(3) = nbrg
    .SubItems(4) = satuan.Text
    .SubItems(5) = text5.Text
    .SubItems(6) = text6.Text
    .SubItems(7) = val(XPText4.Text)
    .SubItems(8) = XPText5.Text
    .SubItems(9) = kodegdg
    .SubItems(10) = txtbatch.Text
    .SubItems(11) = txtppn.Text
    .SubItems(12) = Format(tgl.Value, "yyyy-mm-dd")
    End With
  LV1.Enabled = True
ttl
diskon
ttl2
ttl_item
ttlppn
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not Text2.Text = "" Then
sql = "select * from pemesanan,tblsupplier where no_pemesanan='" & Text2.Text & "' and pemesanan.id_supplier=tblsupplier.id_supplier"
Set RS = New Recordset
Set RS = jual.Execute(sql)
If Not RS.EOF Then
po = True
text3.Text = RS!Supplier
DTPicker1.Value = RS!tanggal_pesan

isitabel

Else
MsgBox "Tidak ada nomor PO ini!", vbInformation, judul
End If
Else
text3.SetFocus
End If
End If
End Sub

Private Sub text3_Click()
text4.SetFocus
End Sub

 Sub text4_Click()
stri = Replace(text4.Text, "'", "''")
If text3.Text = "" Then
MsgBox "Isi supplier dulu"
text3.SetFocus
Exit Sub
End If

Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari
text5.Text = ""
If cari Is Nothing Then

sql = "select * from tblbarang where  kode_brg='" & stri & "'"
Set RS = New Recordset
Set RS = jual.Execute(sql)


proses
isibeli = True

'Barang.text3.SetFocus
Else
text4.SetFocus
text4.Text = ""
MsgBox "Nama barang sudah terdaftar"
End If
End Sub
Sub proses()
On Error Resume Next
sql = "select * from tblbarang where kode_brg='" & stri & "'"


Set RS = jual.Execute(sql)
If Not RS.EOF Then
tampil_stn

stok = RS!stok
cosbli = RS.Fields("Harga_beli")
hb = "Rp" & Format(Str(cosbli), "#,##.") + "/" + RS.Fields("satuan")
cosju = RS!harga_jual
hj.Caption = "Rp" & Format(Str(cosju), "#,##") + "/" + RS!satuan

stn_baku = RS!satuan
kode = RS!kode_brg
nbrg = RS!deskripsi
stn = RS!satuan
satuan.Text = stn

txtbatch.SetFocus

XPText5.Text = val(text6.Text) * val(text5.Text) - val(XPText4.Text) + val(txtppn.Text)
datagdg
Else
brgbaru
datagdg
End If
'barang baru

End Sub
Sub getkodegdg()
Set rsd = New Recordset
rsd.Open "select kode_gudang from gudang where nama_gudang='" & cmbgudang.Text & "' or kode_gudang='" & cmbgudang.Text & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsd.EOF Then
kodegdg = rsd!kode_gudang
Else
kodegdg = "utama"
End If

End Sub

Private Sub cmbgudang_Click()
getkodegdg
Set rsd = New Recordset
rsd.Open "select coalesce(jumlah,0) as jumlah from stokgudang where kode_gudang='" & kodegdg & "' and kode_brg='" & kode & "'", jual, adOpenStatic, adLockOptimistic
stok.Caption = IIf(rsd.EOF, "0", rsd!jumlah)

text5.SetFocus
End Sub

Sub brgbaru()
kode = text4.Text

MsgBox "Barang baru,isi beberapa data", vbInformation, judul
k = "salah"
Do While k = "salah"

nbrg = InputBox("Masukkan nama barang:")
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where deskripsi='" & nbrg & "'", jual, adOpenStatic, adLockOptimistic
If rsbarang.EOF Then
k = "benar"
Else
MsgBox "Nama barang sudah ada,coba lagi!", vbInformation, judul

End If
Loop
satuan.SetFocus

End Sub
Sub databaru()

If MsgBox("Barang baru,tambah ke database?", vbYesNo) = vbYes Then
stok = ""
hb = ""
hj = ""
Me.Hide
Barang.Show
Barang.cmdtambah_Click


Barang.Text2.Text = text4.Text
Barang.text4.Text = "0"
Barang.text4.Locked = True
Barang.text5.Text = "0"
Barang.text5.Locked = True
Barang.text7.Text = "0"
Barang.kode_oto

Barang.text3.SetFocus
Else
awal
kosong
kosong2
End If
End Sub


Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" Then
bayar.SetFocus
Else
text4_Click
End If
End If

End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 13 And text5.Text = "" Then
     If satuan.Text <> stn_baku Then
      Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rsstn.EOF Then
    text5.Text = cosbli * rsstn!konversi
    End If
      Else
            text5.Text = cosbli
End If
text6.SetFocus

Else
If KeyAscii = 13 And text5.Text <> "" Then

text6.SetFocus
End If
End If
End Sub

Private Sub Text5_LostFocus()
kete.Caption = ""
End Sub

Private Sub Text6_Change()
XPText3_Change
End Sub

Private Sub text6_GotFocus()
kete.Caption = "Isi jumlah,lalu tekan enter"
End Sub

Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)


If KeyAscii = 13 Then

If text6.Text = "" Then
text6.Text = "1"
End If

If text5.Text <> "" And satuan.Text <> "" Then


Set RS = New Recordset
stri = Replace(text4.Text, "'", "''")
isigrid

'DataGrid1.SetFocus
text4.SetFocus
kosong2
Else
MsgBox "Harga beli/satuan jangan kosong", vbCritical
text5.SetFocus
End If
End If
End Sub
Sub ttl_item()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(I).SubItems(6)
Next I
text7.Text = sum

End Sub

Sub ttl()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(I).SubItems(6) * LV1.ListItems(I).SubItems(5)
Next I

text8.Text = sum

End Sub
Sub ttl2()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(I).SubItems(8)
Next I
XPText6.Text = sum

End Sub
Sub diskon()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(7))
Next I
XPText2.Text = sum

End Sub
Sub ttlppn()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(11))
Next I
ttlppnt.Text = sum

End Sub
Sub isitabel()

On Error Resume Next
Set rstrans = New Recordset
ListView5.Visible = True

sql = "select detilpesan.kode_brg,tblbarang.deskripsi,detilpesan.jumlah_brg2,detilpesan.satuan,detilpesan.jumlah_beli from detilpesan,pemesanan,tblbarang where pemesanan.no_pemesanan=detilpesan.no_pemesanan and detilpesan.kode_brg=tblbarang.kode_brg and pemesanan.no_pemesanan='" & ListView3.SelectedItem.SubItems(1) & "'"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView5.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView5.ListItems.Add(, , ListView5.ListItems.count + 1)
           l.SubItems(1) = ListView5.ListItems.count & "."

        l.SubItems(2) = ![kode_brg]
        l.SubItems(3) = ![deskripsi]
            l.SubItems(5) = ![jumlah_brg2]

                l.SubItems(4) = ![satuan]
l.SubItems(6) = ![jumlah_beli]
    .MoveNext
    Loop
End With
End Sub
Private Sub supp()

  Dim I As Long
  Dim j As Long

text3.Clear
Set RS = New Recordset
sql = "select * from tblsupplier order by id_supplier"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text3.AddItem rsbarang!Supplier
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub

Private Sub sort()
If kolom.Text = "Semua" Then

dbgrid

Else


kolom2.Clear
sql = "select * from tblbarang order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
 If kolom.Text = "Deskripsi" Then
kolom2.AddItem rsbarang!deskripsi
Else

End If
rsbarang.MoveNext
 Loop
  End If
    With kolom2
    For I = 0 To .ListCount - 1
      For j = .ListCount To (I + 1) Step -1
         If .List(j) = .List(I) Then
           .RemoveItem j
         End If
      Next j
    Next I
  End With
End If

  End Sub
Private Sub kolom_Change()
sort
End Sub


Private Sub kolom2_Change()
If kolom.Text = "Semua" Then Exit Sub
Set rsbarang = New Recordset
sql = "select * from tblbarang where  " & kolom.Text & "  like '%" & kolom2.Text & "%'"
Set rsbarang = jual.Execute(sql)
Set dbgrid1.DataSource = rsbarang

End Sub
Private Sub kolom2_Click()
Set rsbarang = New Recordset
sql = "select * from tblbarang where  " & kolom.Text & "  ='" & kolom2.Text & "'"
Set rsbarang = jual.Execute(sql)
Set dbgrid1.DataSource = rsbarang

End Sub

Private Sub ThemedButton1_GotFocus()
ThemedButton1.BackColor = Red
End Sub

Private Sub Text6_LostFocus()
kete.Caption = ""
End Sub

Private Sub Text9_Change()
dbgridtrans2
End Sub

Private Sub tgl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
text6_KeyPress (13)

End If
End Sub

Private Sub txtcari_Change()
dbgridcari
End Sub

Private Sub txtcari_GotFocus()
ktr.Caption = "Tekan enter bila telah selesai"
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If lvbrg.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
lvbrg.SetFocus
End If

End Sub

Private Sub txtcrp_Change()
dbgridplg2
End Sub
Private Sub txtcrp_KeyPress(KeyAscii As Integer)
If lvplg.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
lvplg.SetFocus
End If
End Sub
Private Sub lvplg_DblClick()
On Error Resume Next
If Edit = False Then
text3.Text = lvplg.SelectedItem.SubItems(2)
Tab1.Tab = 0
text3_Click
Else
MsgBox "Tekan tombol baru dulu"

End If

End Sub
Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub

Private Sub txtppnp_Change()
If txtppnp.Text = "0" Then
txtppn.Text = "0"
Else
txtppn.Text = val(txtppnp.Text) * 0.01 * (val(text6.Text) * val(text5.Text) - val(XPText4.Text))
End If
XPText5.Text = val(text6.Text) * val(text5.Text) - val(XPText4.Text) + val(txtppn.Text)

End Sub

Private Sub txtppnp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text6_KeyPress (13)
End If
End Sub

Private Sub XPText3_Change()
If XPText3.Text = "0" Then
XPText4.Text = "0"
Else
XPText4.Text = val(XPText3.Text) * 0.01 * val(text6.Text) * val(text5.Text)
End If
XPText5.Text = val(text6.Text) * val(text5.Text) - val(XPText4.Text)
End Sub

Private Sub XPText3_GotFocus()
txtppnp.Text = ""
txtppn.Text = ""
End Sub

Private Sub XPText4_Change()

XPText5.Text = val(text6.Text) * val(text5.Text) - val(XPText4.Text)
End Sub

Private Sub XPText3_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 13 Then
text6_KeyPress (13)
End If
End Sub

Private Sub XPText4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text6_KeyPress (13)
End If
End Sub


