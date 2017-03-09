VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form transaksi2 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   10005
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16485
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   16485
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   9855
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   17383
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Penjualan"
      TabPicture(0)   =   "transaksi2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "jam"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SkinLabel13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SkinLabel11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "kasir"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "stok"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "SkinLabel1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "no"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "nama"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SkinLabel9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ket"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "total"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "LV1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MSComm1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "hapus"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "customer"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "hgc"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Timer1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "baru"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "command2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "command1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "command3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "DataGrid1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Skin1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "ket2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Data barang"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "kolom2"
      Tab(1).Control(1)=   "kolom"
      Tab(1).Control(2)=   "dbgrid1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Data pelanggan"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Cmdcari"
      Tab(2).Control(1)=   "Dbgrid2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Data transaksi"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dbgrid3"
      Tab(3).ControlCount=   1
      Begin ACTIVESKINLibCtl.SkinLabel ket2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi2.frx":001C
         TabIndex        =   59
         Top             =   9480
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   3840
         OleObjectBlob   =   "transaksi2.frx":0096
         Top             =   360
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3975
         Left            =   4560
         TabIndex        =   2
         Top             =   3600
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7011
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
      Begin penjualan.ThemedButton command3 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   8880
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Keluar"
         font            =   "transaksi2.frx":02CA
      End
      Begin penjualan.ThemedButton command1 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   8880
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         caption         =   "&Cetak"
         font            =   "transaksi2.frx":02F6
      End
      Begin penjualan.ThemedButton command2 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   8880
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         caption         =   "&Batal/F3"
         font            =   "transaksi2.frx":0322
      End
      Begin penjualan.ThemedButton baru 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   8880
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         caption         =   "&Baru/F2"
         font            =   "transaksi2.frx":034E
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   10440
         Top             =   600
      End
      Begin VB.CommandButton Cmdcari 
         Caption         =   "&Cari pelanggan"
         Height          =   450
         Left            =   -70680
         TabIndex        =   6
         Top             =   8220
         Width           =   2535
      End
      Begin VB.CheckBox hgc 
         Caption         =   "Harga grosir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   7
         Top             =   3360
         Width           =   1335
      End
      Begin XPControls.XPCombo customer 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   3000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Text            =   "Pelanggan bebas"
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
      Begin VB.ComboBox kolom2 
         Height          =   420
         Left            =   -72000
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   6300
         Width           =   2055
      End
      Begin VB.CommandButton hapus 
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   8880
         Width           =   855
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   15
         Top             =   3600
         Width           =   3975
         Begin VB.ComboBox text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   24
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   21
            Top             =   1200
            Width           =   1935
         End
         Begin VB.TextBox diskon 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1620
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":037A
            TabIndex        =   41
            Top             =   1680
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":03EA
            TabIndex        =   42
            Top             =   2040
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":045A
            TabIndex        =   43
            Top             =   1200
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":04C4
            TabIndex        =   44
            Top             =   720
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":053A
            TabIndex        =   45
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel stn 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "transaksi2.frx":05BC
            TabIndex        =   53
            Top             =   720
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   3120
            OleObjectBlob   =   "transaksi2.frx":061A
            TabIndex        =   54
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   9
         Top             =   6060
         Width           =   3975
         Begin VB.TextBox XPText6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox XPText2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox gtot 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            TabIndex        =   13
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   2040
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":067A
            TabIndex        =   46
            Top             =   2400
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":06EE
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":0760
            TabIndex        =   48
            Top             =   2040
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":07D0
            TabIndex        =   49
            Top             =   1680
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":0842
            TabIndex        =   50
            Top             =   600
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":08B6
            TabIndex        =   51
            Top             =   1320
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi2.frx":093C
            TabIndex        =   52
            Top             =   960
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker 
            Height          =   255
            Left            =   1920
            TabIndex        =   60
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   64552963
            CurrentDate     =   40299
         End
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   5520
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   8295
         Left            =   4440
         TabIndex        =   26
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   14631
         View            =   3
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
            Text            =   "Harga jual"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Sub total"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
      End
      Begin XPControls.XPText total 
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1720
         Text            =   ""
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   33.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Enabled         =   0   'False
      End
      Begin XPControls.XPCombo kolom 
         Height          =   315
         Left            =   -74880
         TabIndex        =   30
         Top             =   6300
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
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   31
         Top             =   900
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9128
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
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
      Begin MSDataGridLib.DataGrid Dbgrid2 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   17
         Top             =   900
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12726
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
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   18
         Top             =   960
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   12726
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
      Begin ACTIVESKINLibCtl.SkinLabel ket 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi2.frx":09B2
         TabIndex        =   34
         Top             =   900
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi2.frx":0A10
         TabIndex        =   35
         Top             =   3360
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel nama 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "transaksi2.frx":0A78
         TabIndex        =   36
         Top             =   2400
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel no 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "transaksi2.frx":0AD6
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi2.frx":0B34
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel st 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "transaksi2.frx":0BAE
         TabIndex        =   39
         Top             =   -2040
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel stok 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "transaksi2.frx":0C0C
         TabIndex        =   40
         Top             =   3360
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel kasir 
         Height          =   375
         Left            =   10560
         OleObjectBlob   =   "transaksi2.frx":0C6A
         TabIndex        =   55
         Top             =   8880
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Left            =   9720
         OleObjectBlob   =   "transaksi2.frx":0CE4
         TabIndex        =   56
         Top             =   8880
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel label10 
         Height          =   375
         Left            =   5520
         OleObjectBlob   =   "transaksi2.frx":0D4E
         TabIndex        =   57
         Top             =   9120
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "transaksi2.frx":0DAC
         TabIndex        =   58
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "F4--->Cari pelanggan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   9240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "F3--->Cari barang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   9240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "F5--->mengakhiri transaksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   9240
         Width           =   1935
      End
      Begin VB.Label jam 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   32
         Top             =   9120
         Width           =   1695
      End
   End
End
Attribute VB_Name = "transaksi2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode, nbrg As String

Private Sub baru_Click()
Edit = False
kosong
kosong2
no_oto
text4.Enabled = True
text4.SetFocus
command1.Enabled = False
Text2.Enabled = True
ket.Caption = ""
End Sub
Sub dbgridp()
Set rsplg = New Recordset

sql = "select * from pelanggan"
Set rsplg = jual.Execute(sql)

Set Dbgrid2.DataSource = rsplg


End Sub

Private Sub cmdcari_Click()
Dim kata As String

If Cmdcari.Caption = "&Cari pelanggan" Then
kata = InputBox("Masukkan id pelanggan atau nama pelanggan", "Cari...")
If kata = "" Then Exit Sub
sql = "select* from pelanggan where id_pelanggan='" & kata & "' or nama like '%" & kata & "%' "
Set rsplg = New Recordset
Set rsplg = jual.Execute(sql)

If Not rsplg.EOF Then
Set Dbgrid2.DataSource = rsplg
Cmdcari.Caption = "&Refresh"
Else
MsgBox "Tidak ada", vbOKOnly, judul
dbgridp
End If
Else
dbgridp
Cmdcari.Caption = "&Cari pelanggan"
End If

End Sub

Private Sub Command4_Click()
Tab1.Tab = 1
End Sub

Private Sub DataGrid1_DblClick()
On Error GoTo erol

text4.Text = DataGrid1.Columns(1)
text4_Click
erol:
 If err.Description <> vbNullString Then Exit Sub

End Sub

Private Sub dbgrid1_DblClick()
On Error Resume Next
If Edit = False Then
text4.Text = rsbarang!deskripsi
Tab1.Tab = 0
text4_Click
End If

End Sub

Private Sub dbgrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dbgrid1_DblClick
dbgrid1.Enabled = False
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If MsgBox("Simpan transaksi?", vbYesNo, "Tanya") = vbYes Then
simpandata
dbgridtrans
command1.Enabled = True
If MsgBox("Cetak struk belanja?", vbYesNo) = vbYes Then
Command1_Click
baru.SetFocus
Else
awal
command1.Enabled = True
baru.SetFocus
End If
End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
Unload Me
Else
If KeyCode = vbKeyF2 Then
baru_Click
Else
If KeyCode = vbKeyF5 Then
If LV1.ListItems.count <> 0 Then
text4_KeyPress (13)
Else
MsgBox "Belum ada transaksi", vbInformation, judul

End If
Else
If KeyCode = vbKeyF3 Then
Tab1.Tab = 1
kolom2.SetFocus
Else
If KeyCode = vbKeyF4 Then
Tab1.Tab = 2
cmdcari_Click
End If
End If
End If
End If
End If
End Sub



Private Sub Tab1_Click(PreviousTab As Integer)
dbgrid1.Enabled = True

End Sub

Private Sub Text2_GotFocus()
DataGrid1.Visible = False
End Sub

Private Sub text4_Change()
Set DataGrid1.DataSource = Nothing

Set rssupp = New Recordset

sql = "select  * from tblbarang where deskripsi like'%" & text4.Text & "%'"
Set rssupp = jual.Execute(sql)
Set DataGrid1.DataSource = rssupp

End Sub

Private Sub text4_Click()
Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari

If cari Is Nothing Then
sql = "select * from tblbarang where deskripsi='" & text4.Text & "' or kode_brg='" & text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
kode = rsbarang!kode_brg

nama.Caption = rsbarang!deskripsi
If hgc.Value = Checked Then
Text1.Text = rsbarang!Harga_jual_grosir
Else
Text1.Text = rsbarang!Harga_jual
End If
stok.Caption = rsbarang!stok
diskon.Text = Str(rsbarang!diskon)
Text2.SetFocus
Else
Tab1.Tab = 1
kolom2.Text = text4.Text
kolom2_KeyPress (13)
End If
Else
LV1.SelectedItem.SubItems(5) = val(LV1.SelectedItem.SubItems(5)) + 1
LV1.SelectedItem.SubItems(7) = val(LV1.SelectedItem.SubItems(5)) * val(LV1.SelectedItem.SubItems(4)) - val(LV1.SelectedItem.SubItems(6)) / 100 * val(LV1.SelectedItem.SubItems(5)) * val(LV1.SelectedItem.SubItems(4))
ttl
ttl_item
diskons
ttl2

text4.SetFocus
text4.Text = ""
End If
dbgrid1.Enabled = True

End Sub
Sub ttl_item()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(5))
Next i
Text7.Text = sum

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

Private Sub kolom2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set rsbarang = New Recordset
sql = "select * from tblbarang where  kode_brg ='" & kolom2.Text & "' or deskripsi like '%" & kolom2.Text & "%'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
Set dbgrid1.DataSource = rsbarang
dbgrid1.SetFocus

Else
MsgBox "Tak ada barang dengan nama ini", vbCritical, judul
End If

End If

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
Private Sub cust()

customer.Clear
sql = "select * from pelanggan order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
customer.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  End If
rsplg.Close


  End Sub

Private Sub text4_GotFocus()
DataGrid1.Visible = True
Set rssupp = New Recordset
pos = "2"
sql = "select  * from tblbarang order by deskripsi"
Set rssupp = jual.Execute(sql)
Set DataGrid1.DataSource = rssupp

If text4.Text = "" And val(Text7.Text) <> "0" Then
ket2.Caption = "Tekan enter bila telah selesai"
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" And val(Text7.Text) <> "0" Then
Text5.SetFocus
DataGrid1.Visible = False
Else
text4_Click
End If
End If
End Sub

Private Sub Command1_Click()
CetakData
baru.SetFocus
End Sub
Sub CetakData()
Dim mno, mhal, mbaris As Integer
Dim i, n As Integer
Dim mgrs, mgrss As String
If IsPrinterInstalled Then
Printer.Print "";
Printer.Print "";
Printer.Font = "Courier New"
Printer.ForeColor = &H0&
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = 0
Printer.CurrentY = 0
mno = 0
mhal = 0
i = 1
Do While i <= LV1.ListItems.count
    mbaris = 1
    mhal = mhal + 1
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print Tab(4); 'Form8.Caption;
    Printer.FontBold = False
    Printer.FontSize = 8
    Printer.Print Tab(3); "Nama perusahaan"
    Printer.Print Tab(3); "Alamat perusahaan"
    Printer.Print Tab(3); "";
    Printer.Print Tab(3); "Tanggal";
    Printer.Print Tab(23); ":"; label10.Caption;
    Printer.Print Tab(3); "Jam";
    Printer.Print Tab(23); ":"; Format(Now, "hh:mm:ss");
    Printer.Print Tab(3); "No Bukti";
        Printer.Print Tab(23); ":"; no.Caption;

    Printer.Print Tab(3); "Kasir";

        Printer.Print Tab(23); ":"; kasir.Caption;

    Printer.FontBold = False
    Printer.Print ; " "
    Printer.Print ; " "
mgrs = String$(45, "=")
mgrss = String$(40, "-")
Printer.Print Tab(3); mgrss
Printer.FontBold = False
Printer.Print Tab(3); "No";

Printer.Print Tab(8); "Nama Barang";
Printer.Print Tab(36); "Harga";


'Printer.Print Tab(65); "Jumlah";


Printer.FontBold = False
Printer.Print Tab(3); mgrss
mbaris = 0
Do While i <= LV1.ListItems.count And mbaris < 60
   Set itm = LV1.ListItems.Item(i)
    mno = mno + 1
        Printer.Print Tab(3); i; Space(3); itm.SubItems(3);

    Printer.Print Tab(8); itm.SubItems(5); " X "; Format(itm.SubItems(4), "###,###,###");
    Printer.Print Tab(30); RKanan(itm.SubItems(7), "###,###,###");
    sql = "select * from tblbarang where kode_brg='" & itm.SubItems(2) & "'"
    Set RS = New Recordset
    Set RS = jual.Execute(sql)
    If RS!diskon <> 0 Then
    Printer.Print Tab(8); "diskon" + itm.SubItems(6); "%";
    End If
RS.Close
    mbaris = mbaris + 1
    i = i + 1
Loop
Printer.Print Tab(3); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = fale
'Printer.Print Tab(4); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = False
'Printer.Print Tab(5); mgrss
Printer.Print Tab(3); "Total Bayar";
Printer.Print Tab(16); ":";
Printer.Print Tab(30); RKanan(XPText6.Text, "###,###,###");
Printer.Print Tab(3); "Tunai";
Printer.Print Tab(16); ":";

Printer.Print Tab(30); RKanan(Text5.Text, "###,###,###");
Printer.Print Tab(3); "Kembalian ";
Printer.Print Tab(16); ":";
Printer.Print Tab(30); RKanan(Text6.Text, "###,###,###");

Printer.FontBold = False
Printer.FontItalic = False
If mbaris >= 30 Then
Printer.NewPage
End If
Loop
Printer.EndDoc
Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

End Sub

Private Sub Command2_Click()
kosong
kosong2
awal
no.Caption = ""
baru.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set jual = New adodb.Connection
        jual.CursorLocation = adUseClient
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/Penjualan.mdb;Jet OLEDB:Database Password=tujuh;"
 Tab1.Tab = 0

 dbgrid

kbrg
dbgridp
dbgridtrans
cust
customer.Text = "Pelanggan bebas"
kasir.Caption = Mnutama.StatusBar1.Panels(2).Text

Ketengah Me
label10.Caption = Format(Now, "dd-mm-YYYY")
    SkinPath = App.Path & "\skin\B-Studio.skn"
    Skin1.LoadSkin SkinPath
    Skin1.ApplySkin Me.hWnd
    kolom.Clear
kolom.AddItem "Semua"
kolom.AddItem "Deskripsi"
kolom.AddItem "Kategori"
kolom.Text = "Semua"

awal
RekamKegiatan ("Masuk form penjualan")
End Sub
Private Sub kolom_Change()
sort
End Sub
Private Sub dbgrid2_DblClick()
On Error Resume Next
If Edit = False Then
customer.Text = Dbgrid2.Columns.Item(1)
Tab1.Tab = 0
text4.SetFocus
End If
End Sub

Private Sub sort()
If kolom.Text = "Semua" Then
kolom2.Clear

dbgrid

Else


kolom2.Clear
sql = "select * from tblbarang order by kode_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
 If kolom.Text = "Deskripsi" Then
kolom2.AddItem rsbarang!deskripsi
Else
If kolom.Text = "Kategori" Then
kolom2.AddItem rsbarang!Kategori
End If
End If
rsbarang.MoveNext
 Loop
  End If
    With kolom2
    For i = 0 To .ListCount - 1
      For j = .ListCount To (i + 1) Step -1
         If .List(j) = .List(i) Then
           .RemoveItem j
         End If
      Next j
    Next i
  End With
End If

  End Sub

Sub dbgrid()
Set rsbarang = New Recordset
sql = "select * from tblbarang"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang


End Sub
Sub dbgridtrans()
Set rstrans = New Recordset
sql = "select * from penjualan"
Set rstrans = jual.Execute(sql)

Set dbgrid3.DataSource = rstrans


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If


If KeyAscii = 13 And (val(stok.Caption) > 0) Then
If Text2.Text = "" And (val(stok.Caption) > 0) Then
Text2.Text = "1"
End If
text4.SetFocus
isigrid
ttl
ttl_item
diskons
ttl2

kosong

End If

End Sub
Private Sub MSComm1_OnComm()

Dim CheckMyScan As String

Dim CheckForCR As String

Dim MyText As String

Dim CountMe As Integer

Dim Counter As Integer

Dim Number As Integer

On Error GoTo Mscomm11:

If MSComm1.CommEvent = 2 And MSComm1.InBufferCount > 0 Then

CheckMyScan = MSComm1.Input

MyText = CheckMyScan

text4.Text = CheckMyScan

CountMe = Len(text4.Text)

Number = 0

Do Until Counter = CountMe

text4.SelStart = Number

text4.SelLength = Len(text4.Text)

CheckForCR = text4.SelText



If CheckForCR = vbCr Then

text4.Text = (MyText) - vbEnter

MSComm1.PortOpen = False

MSComm1.PortOpen = True

Text2.SetFocus

Exit Sub

End If

DoEvents

Number = Number + 1

Counter = Counter + 1

Loop

End If

Exit Sub

Mscomm11:

MsgBox "A error in reading this bar code", vbOKOnly, "POS"

End Sub

Sub ttl()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(5)) * val(LV1.ListItems(i).SubItems(4))
Next i
gtot = Format(sum, "###0")
total.Text = Format(sum, "#,##0")

End Sub
Private Sub Text2_Change()
If val(Text2) <= val(stok) Then

Text3 = val(Text1) * val(Text2) - rsbarang!diskon / 100 * val(Text1.Text) * val(Text2.Text)

Else
MsgBox "Lewat dari stok", vbCritical
Text2 = ""
Text3 = ""
Exit Sub

End If
ket.Caption = "TOTAL:"
End Sub
Sub isigrid()
    Set butir = LV1.ListItems.Add
    With butir
           .SubItems(1) = LV1.ListItems.count & "."

    .SubItems(2) = kode
    .SubItems(3) = nama.Caption
    .SubItems(4) = Text1.Text
    .SubItems(5) = Text2.Text
    .SubItems(6) = diskon.Text
    .SubItems(7) = Text3.Text
    End With
  
End Sub
Private Sub hapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
ttl_item
diskons
ttl2

For X = 1 To LV1.ListItems.count
LV1.ListItems(X).SubItems(1) = X
Next X
End Sub

Sub kosong()
text4.Text = ""
Text1.Text = ""
Text2 = ""
Text3.Text = ""
nama.Caption = ""
stok.Caption = ""
diskon.Text = ""

End Sub
Sub ttl2()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(7))
Next i
XPText6.Text = Format(sum, "###0")
total.Text = Format(sum, "#,##0")

End Sub
Sub diskons()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(6)) / 100 * val(LV1.ListItems(i).SubItems(5)) * val(LV1.ListItems(i).SubItems(4))
Next i
XPText2.Text = Format(sum, "###0")

End Sub

Sub kosong2()
total.Text = ""
gtot = ""
Text5 = ""
Text7.Text = ""
XPText6.Text = ""
XPText2.Text = ""
Text6 = ""
LV1.ListItems.Clear
End Sub

Private Sub text4_LostFocus()

ket2.Caption = ""
End Sub

Private Sub Text5_GotFocus()
ket2.Caption = "Masukkan nominal pembayaran konsumen"
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
If Text5 <> "" And gtot.Text <> "" Then

kembali = val(Text5) - val(XPText6)
Text6 = kembali
total.Text = Format(kembali, "#,##0")
ket.Caption = "KEMBALIAN:"
If val(Text5.Text) < val(XPText6.Text) Then
If customer.Text = "Pelanggan bebas" Then
MsgBox "Pelanggan bebas tidak dapat berhutang"
Text5.Text = ""
Text6.Text = ""
total.Text = ""
Text5.SetFocus
Exit Sub
End If
If MsgBox("Pembayaran kurang dari total penjualan,akan dimasukkan ke hutang pelanggan sebesar " & Format(-1 * Text6.Text, "#,##0") & "", vbOKCancel) = vbCancel Then
Text5.Text = ""
Text6.Text = ""
total.Text = ""
Text5.SetFocus
Exit Sub
Else
SkinLabel14.Visible = True
DTPicker1.Visible = True
DTPicker1.SetFocus
Exit Sub
End If
End If

If MsgBox("Simpan transaksi?", vbYesNo, "Tanya") = vbYes Then
simpandata
dbgridtrans
command1.Enabled = True
If MsgBox("Cetak struk belanja?", vbYesNo) = vbYes Then
Command1_Click
baru.SetFocus
Else
awal
command1.Enabled = True
baru.SetFocus
End If
End If
End If
End If

End Sub

Sub awal()
text4.Enabled = False
Text2.Enabled = False
command1.Enabled = False
Edit = True
End Sub
Sub simpandata()
Set rstrans = New Recordset
Set RS = New Recordset
sql = "select id_pelanggan from pelanggan where nama='" & customer.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
'id = RS!nama
idp = RS!id_pelanggan
RS.Close
Else
id = "Pelanggan bebas"
End If

sql = "insert into penjualan values('" & no & "','" & Format(Now, "dd mmmm YYYY") & "','" & gtot & "','" & XPText2.Text & "','" & XPText6.Text & "','" & kasir.Caption & "','" & idp & "')"

jual.Execute (sql)
For z = 1 To LV1.ListItems.count
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
st = rsbarang!stok
ns = st - LV1.ListItems(z).SubItems(5)
If st >= 0 Then
rsbarang.Fields("stok") = ns
End If
rsbarang.Update
rsbarang.Close

sql = "insert into detiljual values('" & no & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(5) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(7) & "','0','0','0')"
jual.Execute (sql)
    Next z
    If val(Text5.Text) >= val(XPText6.Text) Then
    byr = val(XPText6.Text)
    Else
    byr = val(Text5.Text)
    End If
jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan) values('" & Format(Now, "dd MMM yyyy") & "','Penjualan no faktur " & no.Caption & "','" & byr & "')"
sel = val(Text5.Text) - val(XPText6.Text)

If sel < 0 Then
jual.Execute "insert into piutang values('" & Format(Now, "dd MMM yyyy") & "','" & no & "','" & -1 * sel & "','0','" & Format(DTPicker1.Value, "dd MMM yyyy") & "','" & idp & "')"
SkinLabel14.Visible = False
DTPicker1.Visible = False
Set RS = New Recordset
RS.Open "select * from pelanggan where id_pelanggan='" & idp & "'", jual, adOpenStatic, adLockOptimistic
RS!jumlah_piutang = RS!jumlah_piutang + -1 * sel
RS.Update
RS.Close
End If
End Sub
Sub simpandata2()
Dim al As Double
Dim byr

For z = 1 To LV1.ListItems.count

Set brg = dbase.OpenRecordset("select * from tbbarang where kode= '" & LV1.ListItems(z).SubItems(3) & "'")
al = brg!jumlah

byr = al - LV1.ListItems(z).SubItems(5)
    
    dbase.Execute "insert into tbjual values ( '" & LV1.ListItems(z).SubItems(2) & "', '" & LV1.ListItems(z).SubItems(1) & "'," & _
        " '" & Text6 & "')"
    dbase.Execute "insert into tbdetjual values('" & LV1.ListItems(z).SubItems(2) & "','" & LV1.ListItems(z).SubItems(1) & " ', " & _
        " '" & LV1.ListItems(z).SubItems(3) & "','" & LV1.ListItems(z).SubItems(5) & "')"
 
'update data barang
   dbase.Execute "update tbbarang set jumlah = '" & byr & "' where kode = '" & LV1.ListItems(z).SubItems(3) & "' "
    Next z
    
End Sub

Sub no_oto()
Dim j As Integer
Dim br As String
Set rstrans = New Recordset
sql = "Select no_penjualan from penjualan order by no_penjualan Desc"
Set rstrans = jual.Execute(sql)
If rstrans.EOF = True Then
no = "P-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"

Else
j = val(Right(rstrans(0), 4))
br = "P-" + Format(Now, "YY-") + Format(Now, "MM") + "-" + Format(Str(j + 1), "0000")
no = br
nmr = Format(Mid(rstrans(0), 6, 2))
thn = Format(Mid(rstrans(0), 3, 2))
txt = Format(Mid(no, 6, 2))
thnn = Format(Mid(no, 3, 2))
If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
no = "P-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
End If
End If

End Sub

Private Sub text5_LostFocus()
ket2.Caption = ""
End Sub

Private Sub tgl_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Timer1_Timer()
jam.Caption = Format(Now, "hh:mm:ss")

End Sub



