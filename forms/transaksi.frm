VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form transaksi 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   8970
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab1 
      Height          =   9285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   16378
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
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
      TabPicture(0)   =   "transaksi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "jam"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "thpj"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Image1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SkinLabel24"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtpo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "SkinLabel27"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LV1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "tgll"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ket2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "label10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "total"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Skin1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "SkinLabel1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "nama"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "SkinLabel9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "stok"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ket"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "SkinLabel11"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "hapus"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "SkinLabel13"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "customer"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Timer1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "kasir"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "baru"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "command2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "command1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "command3"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "SkinLabel19"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "kete"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "notr"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Check1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "ThemedButton1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Check2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "CrystalReport1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "ThemedButton4"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Check3"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txttop"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "SkinLabel28"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "SkinLabel29"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "lvbrg2"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Option6"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "option5"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "option4"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cmdtutup"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Command7"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Check4"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Check5"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Frame1"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmbpkt"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Check6"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "SkinLabel25"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "SkinLabel26"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "SkinLabel31"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtkmsp"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtkms"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtiddok"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Command8"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "SkinLabel32"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).ControlCount=   65
      TabCaption(1)   =   "Data obat"
      TabPicture(1)   =   "transaksi.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvbrg"
      Tab(1).Control(1)=   "txtcari"
      Tab(1).Control(2)=   "SkinLabel30"
      Tab(1).Control(3)=   "ktr"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Data pelanggan"
      TabPicture(2)   =   "transaksi.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtcrp"
      Tab(2).Control(1)=   "SkinLabel21"
      Tab(2).Control(2)=   "lvplg"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Data transaksi"
      TabPicture(3)   =   "transaksi.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView2"
      Tab(3).Control(1)=   "ListView1"
      Tab(3).Control(2)=   "SkinLabel15"
      Tab(3).Control(3)=   "Text8"
      Tab(3).Control(4)=   "Command4"
      Tab(3).Control(5)=   "SkinLabel23"
      Tab(3).Control(6)=   "cmdhps"
      Tab(3).Control(7)=   "Command5"
      Tab(3).Control(8)=   "Command6"
      Tab(3).ControlCount=   9
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "transaksi.frx":0070
         TabIndex        =   116
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cari"
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
         Left            =   2760
         TabIndex        =   115
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtiddok 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtkms 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   113
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtkmsp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   112
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":00D0
         TabIndex        =   111
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":0148
         TabIndex        =   110
         Top             =   1200
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   255
         Left            =   7800
         OleObjectBlob   =   "transaksi.frx":01B8
         TabIndex        =   106
         Top             =   7560
         Width           =   1575
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Biarkan stok minus"
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
         Left            =   7920
         TabIndex        =   105
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbpkt 
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
         Left            =   960
         TabIndex        =   104
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Caption         =   "Jenis penjualan"
         Height          =   2175
         Left            =   12840
         TabIndex        =   97
         Top             =   360
         Width           =   2295
         Begin VB.OptionButton Option1 
            Caption         =   "Umum"
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
            TabIndex        =   100
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Grosir"
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
            TabIndex        =   99
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Resep"
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
            TabIndex        =   98
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Label6"
            Height          =   495
            Left            =   360
            TabIndex        =   109
            Top             =   1560
            Width           =   1695
         End
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Tampilkan gambar barang"
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
         Left            =   7920
         TabIndex        =   96
         Top             =   580
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Cetak langsung faktur (tanpa preview)"
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
         Left            =   11400
         TabIndex        =   95
         Top             =   6940
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Buka cash"
         Height          =   255
         Left            =   11400
         TabIndex        =   94
         Top             =   2280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdtutup 
         Caption         =   "&Tutup"
         Height          =   255
         Left            =   13920
         TabIndex        =   93
         Top             =   6340
         Width           =   1215
      End
      Begin VB.OptionButton option4 
         Caption         =   "Struk (usb)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12720
         TabIndex        =   89
         Top             =   6700
         Width           =   1335
      End
      Begin VB.OptionButton option5 
         Caption         =   "Faktur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14040
         TabIndex        =   88
         Top             =   6700
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Struk (lpt)"
         Height          =   255
         Left            =   11400
         TabIndex        =   87
         Top             =   6700
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvbrg2 
         Height          =   5175
         Left            =   240
         TabIndex        =   86
         Top             =   3120
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode barang"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   4940
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel ktr 
         Height          =   375
         Left            =   -68640
         OleObjectBlob   =   "transaksi.frx":023A
         TabIndex        =   82
         Top             =   6580
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
         Height          =   375
         Left            =   -74520
         OleObjectBlob   =   "transaksi.frx":0298
         TabIndex        =   81
         Top             =   6700
         Width           =   2295
      End
      Begin XPControls.XPText txtcari 
         Height          =   375
         Left            =   -72120
         TabIndex        =   80
         Top             =   6580
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
         Left            =   -74760
         TabIndex        =   79
         Top             =   700
         Width           =   14415
         _ExtentX        =   25426
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
         NumItems        =   8
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
            Text            =   "Nama bobat"
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
            Text            =   "HArga jual Grosir"
            Object.Width           =   2540
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
         Height          =   255
         Left            =   6480
         OleObjectBlob   =   "transaksi.frx":032A
         TabIndex        =   78
         Top             =   700
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "transaksi.frx":0390
         TabIndex        =   77
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox txttop 
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
         Left            =   5880
         TabIndex        =   76
         Top             =   700
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cetak S.Jalan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -71040
         TabIndex        =   75
         Top             =   3460
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "PPN"
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
         Left            =   11280
         TabIndex        =   72
         Top             =   6340
         Width           =   855
      End
      Begin apotekbaleendah.ThemedButton ThemedButton4 
         Height          =   255
         Left            =   12240
         TabIndex        =   70
         Top             =   6340
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         Caption         =   "Kalkulator"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":040C
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cetak &Faktur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72960
         TabIndex        =   69
         Top             =   3460
         Width           =   1695
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   14160
         Top             =   7420
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Aktifkan perintah cetak surat jalan"
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
         Left            =   9600
         TabIndex        =   68
         Top             =   8260
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton cmdhps 
         Caption         =   "&Hapus"
         Enabled         =   0   'False
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
         Left            =   -74640
         TabIndex        =   67
         Top             =   3460
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   615
         Left            =   -61680
         OleObjectBlob   =   "transaksi.frx":09A6
         TabIndex        =   66
         Top             =   3820
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -61680
         TabIndex        =   65
         Top             =   4900
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -61680
         TabIndex        =   63
         Top             =   4420
         Width           =   1815
      End
      Begin apotekbaleendah.ThemedButton ThemedButton1 
         Height          =   375
         Left            =   8400
         TabIndex        =   61
         Top             =   6340
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Cetak S.Jalan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":0A62
      End
      Begin VB.CheckBox Check1 
         Caption         =   "No faktur otomatis"
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
         Left            =   7920
         TabIndex        =   56
         Top             =   340
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   -74640
         OleObjectBlob   =   "transaksi.frx":0FFC
         TabIndex        =   55
         Top             =   7300
         Width           =   5055
      End
      Begin XPControls.XPText notr 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
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
      Begin ACTIVESKINLibCtl.SkinLabel kete 
         Height          =   615
         Left            =   4200
         OleObjectBlob   =   "transaksi.frx":10BA
         TabIndex        =   54
         Top             =   6840
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "transaksi.frx":1134
         TabIndex        =   53
         Top             =   340
         Width           =   1335
      End
      Begin apotekbaleendah.ThemedButton command3 
         Height          =   375
         Left            =   9960
         TabIndex        =   48
         Top             =   6340
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Keluar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":11B4
      End
      Begin apotekbaleendah.ThemedButton command1 
         Height          =   375
         Left            =   7440
         TabIndex        =   47
         Top             =   6340
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Cetak"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":174E
      End
      Begin apotekbaleendah.ThemedButton command2 
         Height          =   375
         Left            =   6360
         TabIndex        =   46
         Top             =   6340
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Bata&l"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":1CE8
      End
      Begin apotekbaleendah.ThemedButton baru 
         Height          =   375
         Left            =   5400
         TabIndex        =   0
         Top             =   6340
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "&Baru"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "transaksi.frx":2282
      End
      Begin ACTIVESKINLibCtl.SkinLabel kasir 
         Height          =   375
         Left            =   12480
         OleObjectBlob   =   "transaksi.frx":281C
         TabIndex        =   44
         Top             =   7180
         Width           =   2535
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   10440
         Top             =   580
      End
      Begin XPControls.XPCombo customer 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":2896
         TabIndex        =   37
         Top             =   720
         Width           =   1695
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
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   6340
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
         Height          =   2895
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   3735
         Begin VB.TextBox ppn 
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
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtppn 
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
            Left            =   1680
            TabIndex        =   27
            Top             =   2040
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel stnc2 
            Height          =   255
            Left            =   720
            OleObjectBlob   =   "transaksi.frx":2914
            TabIndex        =   60
            Top             =   1320
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel stnc 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "transaksi.frx":2972
            TabIndex        =   59
            Top             =   960
            Width           =   735
         End
         Begin VB.ComboBox satuan 
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
            TabIndex        =   58
            Top             =   600
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":29D0
            TabIndex        =   57
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox disp 
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
            TabIndex        =   26
            Top             =   1680
            Width           =   375
         End
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
            TabIndex        =   3
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
            TabIndex        =   24
            Top             =   960
            Width           =   1935
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
            TabIndex        =   23
            Top             =   2400
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
            TabIndex        =   25
            Top             =   1320
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
            Left            =   2040
            TabIndex        =   28
            Top             =   1680
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2A3A
            TabIndex        =   18
            Top             =   1680
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2AAA
            TabIndex        =   19
            Top             =   2400
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2B1A
            TabIndex        =   20
            Top             =   1320
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2B84
            TabIndex        =   21
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2BF6
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "transaksi.frx":2C68
            TabIndex        =   108
            Top             =   2040
            Width           =   1095
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
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   5520
         Width           =   3735
         Begin VB.TextBox txttuslah 
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
            TabIndex        =   118
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtdiskt 
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
            TabIndex        =   30
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox dtmbh 
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
            Left            =   2520
            TabIndex        =   31
            Top             =   1320
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2CD2
            TabIndex        =   45
            Top             =   3240
            Visible         =   0   'False
            Width           =   1695
         End
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
            TabIndex        =   41
            Top             =   1800
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
            TabIndex        =   40
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
            TabIndex        =   36
            Top             =   240
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2D46
            TabIndex        =   35
            Top             =   240
            Width           =   1455
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
            TabIndex        =   16
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
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   2520
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
            TabIndex        =   15
            Top             =   2880
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2DB8
            TabIndex        =   12
            Top             =   2880
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2E28
            TabIndex        =   13
            Top             =   2520
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2E9A
            TabIndex        =   14
            Top             =   600
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2F0E
            TabIndex        =   38
            Top             =   1800
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":2F94
            TabIndex        =   39
            Top             =   1080
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   49
            Top             =   3120
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   75104259
            CurrentDate     =   40299
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":300A
            TabIndex        =   50
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "transaksi.frx":3086
            TabIndex        =   117
            Top             =   2160
            Width           =   975
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   375
         Left            =   11640
         OleObjectBlob   =   "transaksi.frx":30F0
         TabIndex        =   5
         Top             =   7180
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "transaksi.frx":315A
         TabIndex        =   6
         Top             =   985
         Width           =   4935
      End
      Begin ACTIVESKINLibCtl.SkinLabel stok 
         Height          =   255
         Left            =   720
         OleObjectBlob   =   "transaksi.frx":31B8
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":3216
         TabIndex        =   8
         Top             =   2280
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel nama 
         Height          =   375
         Left            =   4080
         OleObjectBlob   =   "transaksi.frx":327E
         TabIndex        =   9
         Top             =   2260
         Width           =   5775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":32DC
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   9600
         OleObjectBlob   =   "transaksi.frx":3350
         Top             =   880
      End
      Begin XPControls.XPText total 
         Height          =   855
         Left            =   4200
         TabIndex        =   33
         Top             =   1300
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1508
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
      Begin ACTIVESKINLibCtl.SkinLabel label10 
         Height          =   375
         Left            =   9480
         OleObjectBlob   =   "transaksi.frx":3584
         TabIndex        =   34
         Top             =   7660
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket2 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":35E2
         TabIndex        =   43
         Top             =   9940
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker tgll 
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
         Left            =   5880
         TabIndex        =   52
         Top             =   340
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
         Format          =   151322627
         CurrentDate     =   37623
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   62
         Top             =   3820
         Width           =   12855
         _ExtentX        =   22675
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
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
            Text            =   "Harga jual"
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
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "thpj"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74640
         TabIndex        =   64
         Top             =   480
         Width           =   14295
         _ExtentX        =   25215
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
         NumItems        =   9
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
            Text            =   "ID Pelanggan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Pelanggan"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Diskon"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Kasir"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3615
         Left            =   4200
         TabIndex        =   71
         Top             =   2620
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6376
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   971
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Kode barang"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Harga jual"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jumlah"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Diskon"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "PPN"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Sub total"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "thpj"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "ketr"
            Object.Width           =   0
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
         Height          =   255
         Left            =   9240
         OleObjectBlob   =   "transaksi.frx":3640
         TabIndex        =   73
         Top             =   2140
         Visible         =   0   'False
         Width           =   855
      End
      Begin XPControls.XPText txtpo 
         Height          =   285
         Left            =   10200
         TabIndex        =   74
         Top             =   2140
         Visible         =   0   'False
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
      Begin MSComctlLib.ListView lvplg 
         Height          =   5775
         Left            =   -74760
         TabIndex        =   90
         Top             =   700
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
         NumItems        =   6
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
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Piutang"
            Object.Width           =   2540
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   375
         Left            =   -74520
         OleObjectBlob   =   "transaksi.frx":36A8
         TabIndex        =   91
         Top             =   6820
         Width           =   2295
      End
      Begin XPControls.XPText txtcrp 
         Height          =   375
         Left            =   -72000
         TabIndex        =   92
         Top             =   6700
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "transaksi.frx":373C
         TabIndex        =   103
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "F6--->Rubah jenis harga"
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
         Left            =   4200
         TabIndex        =   102
         Top             =   7560
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "F7--->Rubah satuan"
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
         Left            =   6240
         TabIndex        =   101
         Top             =   7560
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   345
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   7800
         TabIndex        =   85
         Top             =   7780
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
         Left            =   6240
         TabIndex        =   84
         Top             =   7780
         Width           =   2055
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
         Left            =   4200
         TabIndex        =   83
         Top             =   7780
         Width           =   1935
      End
      Begin VB.Label thpj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   12720
         TabIndex        =   51
         Top             =   5860
         Width           =   3495
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
         Left            =   11640
         TabIndex        =   42
         Top             =   7660
         Width           =   1695
      End
   End
End
Attribute VB_Name = "transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kode, nbrg, mt, ket1, kete2, idp, stn_baku, pos, kodebank, bilang, stri As String
Dim jumse, hasil As Double
Dim idsls, jnsjual As String
Dim hari As Integer
Dim nutupjual, letmin As Boolean
Dim c As Integer

Dim hpj, cosju, cosjug, cosjug2 As Currency
Sub tmpilgmbr()
On Error Resume Next
Set Image1.Picture = Nothing
If Check5.Value = Checked Then
Set Image1.Picture = LoadPicture(pss & "\gambar\" & kode & ".jpg")
End If

End Sub
Private Sub lispaket()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbpkt.Clear
sql = "select nama_pkt from paket order by nama_pkt"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbpkt.AddItem rsbarang!nama_pkt
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

  End Sub
Sub kidon()
If Option3.Value = True And txtiddok.Text = "" Then MsgBox "ID Dokter harus diisi", vbCritical: Exit Sub
If MsgBox("Simpan transaksi?", vbYesNo, "Tanya") = vbYes Then
pilih = ""
If val(text5.Text) > 0 Then
If val(text5.Text) <= (val(XPText6.Text) + val(ppn.Text) + val(txttuslah.Text)) Then
pembayaran = val(text5.Text)
hari = DTPicker1.Value - tgll.Value

Else
pembayaran = val(XPText6.Text)
End If
nutupjual = False
pilih = "KAS"
proses2
nutupjual = True

    'tanya.Show
    Else
    proses2
   End If
End If
End Sub

Private Sub Check1_Click()
    SaveSetting "apotekbaleendah", "transaksi", "Check1.value", Check1.Value
    

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

Private Sub baru_Click()
Dim tp As Currency
On Error Resume Next
dbgridplg
cust
Set Image1.Picture = Nothing
Frame1.Enabled = False
jual.Execute "Delete from stok"
    cmdtutup.Enabled = True

If baru.Caption = "&Baru" Then
idp = ""
idsls = ""
notr.Text = ""
Edit = False
kosong
kosong2
text4.Enabled = True
notr.SetFocus
Command1.Enabled = False
Text2.Enabled = True
ket.Caption = ""
tambah
If Check1.Value = Checked Then
GetNumber
pos = "1"
customer.SetFocus

End If

Else
hapus_faktur2
nutupjual = False

tanya.Show
End If
cust
Text2.Locked = False

End Sub

Private Sub tambah()
tgll.Value = Now
tgll.Enabled = True
notr.Enabled = True
customer.Enabled = True
text4.Enabled = True
Text2.Enabled = True
text5.Enabled = True
hapus.Enabled = True
notr.Enabled = True
dtmbh.Enabled = True
End Sub
Private Sub ubh()
tgll.Enabled = False
notr.Enabled = False
customer.Enabled = False
text4.Enabled = False
'Text2.Enabled = False
text5.Enabled = False
hapus.Enabled = False
dtmbh.Enabled = False
End Sub

Private Sub baru_GotFocus()
kete.Caption = "Klik tombol baru untuk memulai transaksi baru"
End Sub

Private Sub Check3_Click()
    SaveSetting "apotekbaleendah", "transaksi", "Check3.value", Check3.Value
If Check3.Value = Checked Then
txtppn.Enabled = True
Else
txtppn.Enabled = False
End If

End Sub


Private Sub Check4_Click()
    SaveSetting "apotekbaleendah", "transaksi", "Check4.value", Check4.Value

End Sub

Private Sub Check5_Click()
SaveSetting "apotekbaleendah", "transaksi", "Check5.value", Check5.Value
End Sub

Private Sub Check6_Click()
On Error Resume Next
If Check6.Value = Checked Then
letmin = True
Else
letmin = False
End If
If Text2.Enabled = True Then
Text2.SetFocus
End If
SaveSetting "apotekbaleendah", "transaksi", "Check6.value", Check6.Value
End Sub

Private Sub cmbpkt_Click()
'lv1.ListItems.Clear
Set RS = New Recordset
RS.Open "select d.*,deskripsi,satuan,harga_beli,stok from paket p,paket_detil d,tblbarang t where p.kode_pkt=d.kode_pkt and d.kode_brg=t.kode_brg and nama_pkt='" & cmbpkt.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub

With RS
.MoveFirst
    Do While Not .EOF
    If letmin = False Then
     If !jumlah <= !stok Then
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = LV1.ListItems.count & "."

        l.SubItems(2) = !kode_brg
        l.SubItems(3) = !deskripsi
        l.SubItems(4) = !satuan
        l.SubItems(5) = !harga_jual
        l.SubItems(6) = !jumlah
        l.SubItems(7) = !diskon
        l.SubItems(8) = 0
        l.SubItems(9) = !subttl
        l.SubItems(10) = !harga_beli
        l.SubItems(11) = cmbpkt.Text
    
    
    
        Else
        MsgBox "Maaf,barang " & !deskripsi & " kurang dari stok!", vbInformation, judul
        End If
    Else
    Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = LV1.ListItems.count & "."

        l.SubItems(2) = !kode_brg
        l.SubItems(3) = !deskripsi
        l.SubItems(4) = !satuan
        l.SubItems(5) = !harga_jual
        l.SubItems(6) = !jumlah
        l.SubItems(7) = !diskon
        l.SubItems(8) = 0
        l.SubItems(9) = !subttl
        l.SubItems(10) = !harga_beli
        l.SubItems(11) = cmbpkt.Text
    End If
    .MoveNext
    
    Loop
End With
ttl
ttl_item
diskons
ttl2
ttlb
kosong
text4.SetFocus
End Sub

Private Sub cmdhps_Click()
If ListView2.ListItems.count = 0 Then Exit Sub

If MsgBox("Yakin akan menghapus no faktur " & ListView2.SelectedItem.SubItems(1) & " dan merubah segala transakasi yang berhubungan dengan no faktur ini?", vbYesNo, jdul) = vbNo Then Exit Sub

hapus_faktur
dbgridtrans
ListView1.ListItems.Clear

End Sub
Sub hapus_faktur()
jual.Execute "delete from penjualan where no_penjualan='" & ListView2.SelectedItem.SubItems(1) & "'"

jual.Execute "delete from byr_piutang where no_penjualan='" & ListView2.SelectedItem.SubItems(1) & "'"

End Sub
Sub hapus_faktur2()


jual.Execute "delete from penjualan where no_penjualan='" & notr.Text & "'"
jual.Execute "delete from byr_piutang where no_penjualan='" & ListView2.SelectedItem.SubItems(1) & "'"

End Sub
Sub palid()
 kata = InputBox("Masukkan username", "Username")
If StrPtr(kata) = 0 Then Exit Sub

If jual.Execute("select* from pengguna where username='" & kata & "'").EOF Then
MsgBox "Username tidak terdaftar"
Command2_Click

Else

 
 Dim ret As String
  SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
  ret = InputBox("Enter Password")
  If StrPtr(ret) = 0 Then Exit Sub
  Set rspengguna = New Recordset
rspengguna.Open "select * from pengguna where username='" & kata & "' and password=md5('" & ret & "')", jual, adOpenDynamic, adLockPessimistic

If Not rspengguna.EOF Then

If rspengguna.Fields(18) = "1" Then
hapus_faktur2
Else
MsgBox "Anda tidak berhak merubah faktur"
Command2_Click
End If
rspengguna.Close
Else
MsgBox "Salah password"
End If
End If
 End Sub

Private Sub cmdtutup_Click()
If cmdtutup.Caption = "&Tutup" Then
If LV1.ListItems.count = 0 Then Exit Sub
jual.Execute "delete from penjualans"
jual.Execute "delete from detiljuals"
cmdtutup.Caption = "Bu&ka"

Set rstrans = New Recordset
sel = val(text5.Text) - val(XPText6.Text)

If val(text5.Text) = 0 Then
ket1 = "B"
kete2 = "BL"
Else
If val(text5.Text) >= val(XPText6.Text) + val(txttmbh2.Text) Then
ket1 = "C"
kete2 = "L"
Else
ket1 = "CB"
kete2 = "BL"
End If
End If
sql = "insert into penjualans(No_penjualan,tanggal,jumlah,total_diskon,total,kasir,id_pelanggan,harga_pokok_jual,keterangan1,keterangan2,ppn,no_po,hari) values('sementara','" & Format(tgll.Value, "YYYY-mm-dd") & "','" & gtot & "','" & val(XPText2.Text) + val(dtmbh.Text) & "','" & XPText6.Text & "','" & kasir.Caption & "','" & idp & "','" & val(thpj.Caption) & "','" & ket1 & "','" & kete2 & "','" & val(ppn.Text) & "','" & txtpo.Text & "','" & val(txttop.Text) & "')"
jual.Execute (sql)
For z = 1 To LV1.ListItems.count
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
st = rsbarang!stok
If LV1.ListItems(z).SubItems(4) <> rsbarang!satuan Then
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.ListItems(z).SubItems(2) & "' and satuan='" & LV1.ListItems(z).SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic
ns = st - LV1.ListItems(z).SubItems(6) * val(rsstn!konversi)
sql = "insert into detiljuals values('" & notr.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) / val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(6) * val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(9) & "','0','0','0','0',0,'" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "')"
jual.Execute (sql)

Else
ns = st - LV1.ListItems(z).SubItems(6)
sql = "insert into detiljuals values('sementara','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(9) & "','0','0','0','0',0,'" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "')"
jual.Execute (sql)

End If


    Next z
Command2_Click

    
    Else
    If cmdtutup.Caption = "Bu&ka" Then
    cmdtutup.Caption = "&Tutup"
    baru_Click

    
    Set rse1 = New Recordset
rse1.Open "Select * from penjualans where no_penjualan='sementara'", jual, adOpenStatic, adLockOptimistic
If rse1.EOF Then
hapus.Enabled = True
customer.SetFocus
Else
idp = rse1!id_pelanggan
Set rsplg = New Recordset
rsplg.Open "Select * from pelanggan where id_pelanggan='" & idp & "'", jual, adOpenStatic, adLockOptimistic
If Not rsplg.EOF Then
customer.Text = rsplg!nama
Else
customer.Text = ""
End If
ket1 = rse1!keterangan1
kete2 = rse1!keterangan2


kete.Caption = "Pilih di tabel barang mana yang akan diubah."
gtot.Text = rse1!jumlah
Set rse2 = New Recordset
rse2.Open "Select sum(diskon) as td from detiljuals where detiljuals.no_penjualan='sementara'", jual, adOpenStatic, adLockOptimistic
XPText2.Text = IIf(IsNull(rse2!td = True), "0", rse2!td)

dtmbh.Text = rse1!total_diskon - val(XPText2.Text)
XPText6.Text = rse1!total
rse2.Close
tgll.Value = rse1!tanggal
txttop.Text = rse1!hari

LV1.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select detiljuals.kode_brg,tblbarang.deskripsi,detiljuals.satuan,detiljuals.harga_beli,detiljuals.harga_jual,detiljuals.jumlah_brg,detiljuals.diskon,detiljuals.total from detiljuals,tblbarang where detiljuals.no_penjualan='sementara' and tblbarang.kode_brg=detiljuals.kode_brg", jual, adOpenStatic, adLockOptimistic
With rse3
.MoveFirst
    Do While Not .EOF
   Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
   l.SubItems(1) = LV1.ListItems.count & "."
        l.SubItems(2) = !kode_brg
              l.SubItems(3) = ![deskripsi]
                            l.SubItems(4) = ![satuan]
Set rse4 = New Recordset
rse4.Open "select* from satuan where kode_brg='" & l.SubItems(2) & "'and satuan='" & l.SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic


              l.SubItems(5) = ![harga_jual] * rse4!konversi
                l.SubItems(6) = ![jumlah_brg] / rse4!konversi
              l.SubItems(8) = ![diskon]
                      l.SubItems(9) = ![total]
                      l.SubItems(10) = !harga_beli


    .MoveNext
    Loop
End With



End If

    
    jual.Execute "delete from penjualans"
    jual.Execute "delete from detiljuals"
    
        GetNumber
End If
    End If


End Sub

Private Sub Command4_Click()
dbgridtrans
End Sub


Private Sub Command5_Click()

If ListView1.ListItems.count = 0 Then Exit Sub
If Option4.Value = True Or Option6.Value = True Then
notr_KeyPress (13)
End If
Command1_Click
End Sub

Private Sub Command6_Click()
ThemedButton1_Click
End Sub

Private Sub Command7_Click()
    Dim sPrinter   As String
    Dim sCodes     As String

    sPrinter = Printer.DeviceName
    sCodes = "27,112,0,64,240"
    Call openTillDrawerUsb(sPrinter, sCodes)


End Sub



Private Sub Command8_Click()
If text4.Enabled = False Then Exit Sub
frmcrdok.Show
End Sub

Private Sub customer_Click()
On Error Resume Next

sql = "select  * from pelanggan where nama ='" & customer.Text & "'"
Set RS = New Recordset
Set RS = jual.Execute(sql)
If RS.RecordCount > 1 Then
MsgBox "Nama pelanggan lebih dari 1,wajib di dobel klik di tabel kanan atas atau ketik id customer lalu enter!"
customer.Text = ""
customer.SetFocus
Exit Sub
Else
customer.Text = RS!id_pelanggan

customer_KeyPress (13)
text4.SetFocus
End If

End Sub

Private Sub customer_GotFocus()
If Edit = True Then
kete.Caption = "Rubah harga jual lalu tekan enter bila selesai atau isi diskon terlebih dahulu lalu enter"
Else
kete.Caption = "Tekan f4,ketik sebagian nama,enter,lalu pilih dengan tombol panah atas bawah lalu enter"
End If

End Sub

Private Sub customer_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
 
 
           

Set RS = New Recordset
sql = "select * from pelanggan where id_pelanggan='" & customer.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
'id = RS!nama
idp = RS!id_pelanggan
idsls = RS!id_sales
customer.Text = RS!nama

RS.Close
text4.SetFocus
Else
If customer.Text = "" Then
idp = "bebas"
idsls = ""
Else
idp = ""
idsls = ""
End If
text4.SetFocus
End If
End If

End Sub

Private Sub Dbgrid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dbgrid2_DblClick
text4.SetFocus
End If
End Sub

Private Sub Form_Deactivate()
If nutupjual = True Then
'Unload Me
End If

End Sub


Private Sub hgc_Click()

End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
ListView2_Click

End If

End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
If LV1.ListItems.count = 0 Then Exit Sub
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
If KeyPress = vbKeyDelete Then
End If
End Sub

Private Sub lvbrg_DblClick()
On Error Resume Next
'If Edit = False Then
text4.Text = lvbrg.SelectedItem.SubItems(1)
Tab1.Tab = 0
text4_Click
'Else
'MsgBox "Tekan tombol baru dulu"

'End If

End Sub

Private Sub lvbrg_GotFocus()
ktr.Caption = "Dobel klik atau enter untuk mengirimkan data"
End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub
Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub

Private Sub DTPicker1_Change()
txttop.Text = DTPicker1.Value - tgll.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "apotekbaleendah", "transaksi", "option4.value", Option4.Value
    SaveSetting "apotekbaleendah", "transaksi", "option5.value", option5.Value
    SaveSetting "apotekbaleendah", "transaksi", "Option1.value", Option1.Value
    SaveSetting "apotekbaleendah", "transaksi", "Option2.value", Option2.Value
    SaveSetting "apotekbaleendah", "transaksi", "Option3.value", Option3.Value
End Sub

Private Sub ListView2_Click()
On Error Resume Next
text4.Enabled = True
notr.Text = ListView2.SelectedItem.SubItems(1)
If ListView2.ListItems.count = 0 Then Exit Sub

If ListView2.ListItems.count <> 0 Then
cmdhps.Enabled = True
End If
ListView1.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select detiljual.kode_brg,tblbarang.deskripsi,detiljual.satuan,detiljual.harga_beli,detiljual.harga_jual,detiljual.jumlah_brg,detiljual.diskon,detiljual.total from detiljual,tblbarang where detiljual.no_penjualan='" & ListView2.SelectedItem.SubItems(1) & "' and tblbarang.kode_brg=detiljual.kode_brg", jual, adOpenStatic, adLockOptimistic
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


              l.SubItems(5) = ![harga_jual] * rse4!konversi
                l.SubItems(6) = ![jumlah_brg] / rse4!konversi
              l.SubItems(7) = ![diskon]
                      l.SubItems(9) = ![total]
                      l.SubItems(10) = !harga_beli


    .MoveNext
    Loop
End With

End If

End Sub

Private Sub ListView2_DblClick()
On Error Resume Next
Tab1.Tab = 0
notr.Text = ListView2.SelectedItem.SubItems(1)
Command1.Enabled = True
notr_KeyPress (13)

End Sub

Private Sub diskon_Change()
text3.Text = val(Text1.Text) * val(Text2.Text) - val(diskon.Text)
End Sub

Private Sub diskon_GotFocus()
disp.Text = ""
txtGotFocus
End Sub

Private Sub diskon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Edit = True Then
Text1_KeyPress (13)
Else
disp_KeyPress (13)
End If
End If
End Sub

Private Sub disp_Change()
If val(disp.Text) > 100 Then
MsgBox "ERORR", , judul
disp.Text = ""
disp.SetFocus
Exit Sub
End If

diskon.Text = val(disp.Text) / 100 * val(Text2.Text) * val(Text1.Text)
text3.Text = ""
text3.Text = val(Text2.Text) * val(Text1.Text) - val(diskon.Text)

End Sub
Public Function txtGotFocus()
Dim obj
Set obj = transaksi.ActiveControl
    If TypeOf obj Is TextBox Then
        obj.SelStart = 0
        obj.SelLength = Len(obj.Text)
    End If
    
    
    

End Function


Private Sub disp_GotFocus()
txtGotFocus
If Edit = True Then
kete.Caption = "Rubah harga jual lalu tekan enter bila selesai atau isi diskon terlebih dahulu lalu enter"
Else
kete.Caption = "Ketik diskon dalam persen,lalu enter"
End If

End Sub

Private Sub disp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Edit = True Then
Text1_KeyPress (13)
Else
text4.SetFocus
If hpj = val(Text1.Text) Then
MsgBox "Modal sama dengan harga juale", vbCritical, judul
Text1.SetFocus
Exit Sub
Else
If hpj > val(Text1.Text) Then
MsgBox "Modal lebih besar dari harga jualll", vbCritical, judul
Text1.SetFocus
Exit Sub
End If
End If

isigrid
ttl
ttl_item
diskons
ttl2
ttlb
kosong
End If
End If
End Sub

Private Sub dtmbh_Change()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(7))
Next I
XPText6.Text = val(gtot.Text) - (sum + val(dtmbh.Text))

total.Text = Format(val(gtot.Text) - (sum + val(dtmbh.Text)) + val(ppn.Text) + val(txttuslah.Text), "#,#0.#0")


End Sub

Private Sub dtmbh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then


If Check3.Value = Checked Then
txtppn.SetFocus
Else
text5.SetFocus
tmplbyr
End If
End If
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Option3.Value = True And txtiddok.Text = "" Then MsgBox "ID Dokter harus diisi", vbCritical: Exit Sub

If MsgBox("Simpan transaksi?", vbYesNo, "Tanya") = vbYes Then
pilih = ""
If val(text5.Text) > 0 Then
If val(text5.Text) <= (val(XPText6.Text) + val(ppn.Text) + val(txttuslah.Text)) Then
pembayaran = val(text5.Text)
hari = DTPicker1.Value - tgll.Value

Else
pembayaran = val(XPText6.Text)
End If
nutupjual = False
pilih = "KAS"
proses2
nutupjual = True

    'tanya.Show
    Else
    proses2
   End If
End If
End If
End Sub

Private Sub Form_Activate()
'kbrg
If pilih = "KAS" Or pilih = "BANK" Then
proses2
End If
If hpju = "" Then
hpju = "PJ"
End If
If hcus = "" Then
hcus = "CUS"
End If

pjgh = Len(hpju)
nutupjual = True

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
dtmbh_KeyPress (13)
Else
MsgBox "Belum ada transaksi", vbInformation, judul

End If
Else
If KeyCode = vbKeyF3 Then
Tab1.Tab = 1
txtcari.SetFocus
Else
If KeyCode = vbKeyF1 Then
If text4.Enabled = False Then Exit Sub
text4.SetFocus
Else
If KeyCode = vbKeyF6 Then
    'harga
    If Option1.Value = True Then
    Option2.Value = True
    Else
    If Option2.Value = True Then
    Option3.Value = True
    Else
    If Option3.Value = True Then
    Option1.Value = True
    End If
    End If
    End If

Else
If KeyCode = vbKeyF7 Then
If text4.Enabled = False Then Exit Sub
If satuan.ListCount = 0 Then Exit Sub
c = satuan.ListIndex
If (c + 1) >= satuan.ListCount Then c = -1
satuan.ListIndex = c + 1
c = satuan.ListIndex + 1
Else
If KeyCode = vbKeyF8 Then
LV1.SetFocus
Else
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\penjualan.doc" _
                 , vbNullString, vbNullString, 1
Else

If KeyCode = vbKeyF4 Then
Tab1.Tab = 2
txtcrp.SetFocus
Else
If KeyCode = vbKeyF11 Then
frmtutup.Show

Else
If KeyCode = vbKeyF9 And LV1.ListItems.count > 0 Then
rubah = InputBox("Masukan jumlah barang:")
  If StrPtr(rubah) = 0 Then Exit Sub
 If rubah = "" Or val(rubah) = 0 Then Exit Sub

Set rsbeli = New Recordset
rsbeli.Open "Select * from tblbarang where kode_brg='" & LV1.SelectedItem.SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
 stn_baku = rsbeli!satuan
 Set rspo = New Recordset
rspo.Open "select * from detiljual where kode_brg='" & LV1.SelectedItem.SubItems(2) & "' and satuan='" & LV1.SelectedItem.SubItems(4) & "' and no_penjualan='" & notr.Text & "' ", jual, adOpenStatic, adLockOptimistic

 
 
 Set rsbarang2 = New Recordset
rsbarang2.Open "select * from stok where kode_brg='" & LV1.SelectedItem.SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
If rsbarang2.EOF Then
jual.Execute "insert into stok values('" & LV1.SelectedItem.SubItems(2) & "','" & val(rubah) & "')"
Else
If rspo.EOF Then
rss = val(rubah)
Else
rss = val(rubah) - val(rspo!jumlah_brg)

End If
jual.Execute "update stok set stok=" & rss & " where kode_brg='" & LV1.SelectedItem.SubItems(2) & "'"
End If


 If LV1.SelectedItem.SubItems(4) <> stn_baku Then
     Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.SelectedItem.SubItems(2) & "' and satuan='" & LV1.SelectedItem.SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic
 stokjual = val(rubah) * rsstn!konversi
 Else
  stokjual = val(rubah)

 End If


If letmin = False Then
If Edit = False Then
    If val(stokjual) > val(rsbeli!stok) Then
    MsgBox "Lewat dari stok", vbCritical, judul
    Exit Sub
    End If
Else

    If Not rspo.EOF Then
        If (val(stokjual) - val(rspo!jumlah_brg)) > val(rsbeli!stok) Then
        MsgBox "Lewat dari stok", vbCritical, judul
        Exit Sub
        End If
    Else
        If val(stokjual) > val(rsbeli!stok) Then
        MsgBox "Lewat dari stok", vbCritical, judul
        Exit Sub
        End If
    End If

End If
End If
    LV1.SelectedItem.SubItems(6) = rubah
    LV1.SelectedItem.SubItems(7) = (val(LV1.SelectedItem.SubItems(7)) / (val(LV1.SelectedItem.SubItems(9)) + val(LV1.SelectedItem.SubItems(7)))) * (LV1.SelectedItem.SubItems(6) * LV1.SelectedItem.SubItems(5))
    LV1.SelectedItem.SubItems(9) = LV1.SelectedItem.SubItems(5) * LV1.SelectedItem.SubItems(6) - val(LV1.SelectedItem.SubItems(7))


ttl
ttl_item
diskons
ttl2
ttlb
kosong
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub



Private Sub sale_Click()
On Error GoTo erol
text4.SetFocus
erol:
 If err.Description <> vbNullString Then
 MsgBox "Klik baru dulu"
 baru.SetFocus
 Exit Sub
End If
End Sub

Private Sub no_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub


Private Sub lv1_Click()
On Error Resume Next
If Edit = True Then
text4.Text = LV1.SelectedItem.SubItems(3)
text4_Click
Text1.Text = LV1.SelectedItem.SubItems(5)
satuan.Text = LV1.SelectedItem.SubItems(4)
Text2.Text = LV1.SelectedItem.SubItems(6)
diskon.Text = LV1.SelectedItem.SubItems(7)
ppn.Text = LV1.SelectedItem.SubItems(8)
text3.Text = LV1.SelectedItem.SubItems(9)
ket.Caption = "Rubah harga jualnya di textt box harga jual"
Text2.Enabled = True
End If
End Sub

Private Sub lvbrg2_DblClick()
lvbrg2_KeyPress (13)
End Sub

Private Sub lvbrg2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lvbrg2.ListItems.count = 0 Then
lvbrg2.Visible = False
text4.Text = ""
text4.SetFocus
Else

text4.Text = lvbrg2.SelectedItem.SubItems(1)
lvbrg2.Visible = False
text4_Click
End If
End If
End Sub

Private Sub lvbrg2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = 8 Then
text4.SetFocus
End If
End Sub

Private Sub lvbrg2_LostFocus()
lvbrg2.Visible = False
End Sub

Private Sub lvplg_DblClick()
On Error Resume Next
If Edit = False Then
customer.Text = lvplg.SelectedItem.SubItems(2)
Tab1.Tab = 0
customer_Click
Else
MsgBox "Tekan tombol baru dulu"

End If

End Sub

Private Sub notr_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
If notr.Text = "" Then Exit Sub
Set rse1 = New Recordset
rse1.Open "Select * from penjualan where no_penjualan='" & notr.Text & "'", jual, adOpenStatic, adLockOptimistic
If rse1.EOF Then
hapus.Enabled = True
customer.SetFocus
Else
kosong2

'Text2.Locked = True
txttuslah.Text = rse1!tuslah

ppn.Text = rse1!ppn
Set rsbyr = New Recordset
rsbyr.Open "select sum(pemasukan) as pem from keuangan where keterangan like '%" & notr.Text & "' ", jual, adOpenStatic, adLockOptimistic
ttlm = IIf(IsNull(rsbyr!pem) = True, "", rsbyr!pem)
rsbyr.Close
If val(rse1!cash) >= val(rse1!total) Then
text5.Text = rse1!cash
Else
text5.Text = ttlm
End If
idp = rse1!id_pelanggan
idsls = rse1!id_sales
Set rsplg = New Recordset
rsplg.Open "Select * from pelanggan where id_pelanggan='" & idp & "'", jual, adOpenStatic, adLockOptimistic

customer.Text = rsplg!nama

ket1 = rse1!keterangan1
kete2 = rse1!keterangan2


kete.Caption = "Pilih di tabel barang mana yang akan diubah."
Edit = True
gtot.Text = rse1!jumlah
Set rse2 = New Recordset
rse2.Open "Select sum(diskon) as td from detiljual where detiljual.no_penjualan='" & notr.Text & "'", jual, adOpenStatic, adLockOptimistic
XPText2.Text = rse2!td

dtmbh.Text = rse1!total_diskon - rse2!td
XPText6.Text = rse1!total
rse2.Close
tgll.Value = rse1!tanggal
txttop.Text = rse1!hari

LV1.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select ketr,detiljual.kode_brg,tblbarang.deskripsi,detiljual.satuan,detiljual.harga_beli,detiljual.harga_jual,detiljual.jumlah_brg,detiljual.diskon,detiljual.total from detiljual,tblbarang where detiljual.no_penjualan='" & notr.Text & "' and tblbarang.kode_brg=detiljual.kode_brg", jual, adOpenStatic, adLockOptimistic
With rse3
.MoveFirst
    Do While Not .EOF
   Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
   l.SubItems(1) = LV1.ListItems.count & "."
        l.SubItems(2) = !kode_brg
              l.SubItems(3) = ![deskripsi]
                            l.SubItems(4) = ![satuan]
Set rse4 = New Recordset
rse4.Open "select* from satuan where kode_brg='" & l.SubItems(2) & "'and satuan='" & l.SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic


              l.SubItems(5) = ![harga_jual] * rse4!konversi
              l.SubItems(6) = ![jumlah_brg] / rse4!konversi
              l.SubItems(7) = ![diskon]
              l.SubItems(8) = ![ppn]
              l.SubItems(9) = ![total]
              l.SubItems(10) = !harga_beli
              l.SubItems(11) = !ketr


    .MoveNext
    Loop
End With

Text2.Enabled = True


End If


End If
End Sub

Private Sub Option1_Click()
SaveSetting "apotekbaleendah", "transaksi", "Option1.value", Option1.Value
txtiddok.Text = ""
Command8.Enabled = False
txtkmsp.Text = ""
txtkms.Text = ""
txtkmsp.Enabled = False
txtkms.Enabled = False
End Sub


Private Sub Option2_Click()
SaveSetting "apotekbaleendah", "transaksi", "Option2.value", Option2.Value
txtiddok.Text = ""
Command8.Enabled = False
txtkmsp.Text = ""
txtkms.Text = ""
txtkmsp.Enabled = False
txtkms.Enabled = False
End Sub

Private Sub Option3_Click()
SaveSetting "apotekbaleendah", "transaksi", "Option3.value", Option3.Value
txtiddok.Text = ""
Command8.Enabled = True
txtkmsp.Text = ""
txtkms.Text = ""
txtkmsp.Enabled = True
txtkms.Enabled = True
End Sub

Private Sub option4_Click()
    SaveSetting "apotekbaleendah", "transaksi", "option4.value", Option4.Value

End Sub

Private Sub option5_Click()
    SaveSetting "apotekbaleendah", "transaksi", "option5.value", option5.Value

End Sub

Private Sub Option6_Click()
    SaveSetting "apotekbaleendah", "transaksi", "Option6.value", Option6.Value

End Sub



Private Sub satuan_Click()

Text1.Text = ""
If satuan.Text = "" Then
Text1.Text = ""
Text2.Text = ""

stnc.Caption = ""
stnc2.Caption = ""

Else
stnc.Caption = "/" & satuan.Text
stnc2.Caption = satuan.Text

End If
If Text1.Enabled = True Then

Text1.SetFocus
Else
Text1_KeyPress (13)
Text2.SetFocus
End If
End Sub


Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 0 Then
baru.Caption = "&Baru"
Else
If Tab1.Tab = 1 Then
dbgrid

txtcari.Text = ""
txtcari.SetFocus
ktr.Caption = ""
Else
If Tab1.Tab = 2 Then
txtcrp.Text = ""
txtcrp.SetFocus
Else
If Tab1.Tab = 3 Then
dbgridtrans
End If

End If
End If
End If
End Sub

Private Sub Text1_GotFocus()
If Edit = True Then
kete.Caption = "Rubah harga jual lalu tekan enter bila selesai atau isi diskon terlebih dahulu lalu enter"
Else
kete.Caption = "Tekan enter untuk menyesuaikan harga dengan harga di master barang,bila beda isi harga baru kemudian tekan tab"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
    Text2.SetFocus
End If
   If KeyAscii = 13 And Text1.Text = "" Then
     If Option1.Value = True Then
        If satuan.Text <> stn_baku Then
        Set rsstn = New Recordset
        rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic
            If Not rsstn.EOF Then
            Text1.Text = rsstn!harga
            Else
            Text1.Text = cosju
            End If
         Else
        Text1.Text = cosju
        End If
      Else
      If Option2.Value = True Then
        If satuan.Text <> stn_baku Then
        Set rsstn = New Recordset
        rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic

        Text1.Text = rsstn!harga
        Else
        Text1.Text = cosjug
        End If
    Else
    If satuan.Text <> stn_baku Then
        Set rsstn = New Recordset
        rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic

        Text1.Text = rsstn!harga
        Else
        Text1.Text = cosjug2
        End If
     End If
     End If

    

    End If
End Sub

Private Sub Text1_LostFocus()
kete.Caption = ""
End Sub


Private Sub Text2_GotFocus()
If Edit = True Then
kete.Caption = "Rubah harga jual lalu tekan enter bila selesai atau isi diskon terlebih dahulu lalu enter"
Else
kete.Caption = "Ketik jumlah barang lalu enter,atau tekan tab bila ada diskon per barang"
End If

End Sub

Private Sub Text4_Change()
dbgridcari2


End Sub
Sub identiti()
Set rsbarang = New Recordset
text4.Text = Replace(text4.Text, "'", "''")

sql = "select * from tblbarang where deskripsi='" & text4.Text & "' or kode_brg='" & text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
nama.Caption = rsbarang!deskripsi

kode = rsbarang!kode_brg
tmpilgmbr
stn_baku = rsbarang!satuan
tampil_stn
satuan.Text = rsbarang.Fields("satuan")

cosju = rsbarang!harga_jual
cosjug = rsbarang!Harga_jual2
cosjug2 = rsbarang!Harga_jual3

hpj = rsbarang!harga_beli
Else
lvbrg2.Visible = True
lvbrg2.SetFocus

End If
End Sub
Private Sub text4_Click()
On Error Resume Next
Text1.Text = ""
Text2.Text = ""

identiti
'If cari Is Nothing Then
Set rsbarang2 = New Recordset
rsbarang2.Open "select * from stok where kode_brg='" & kode & "'", jual, adOpenStatic, adLockOptimistic
If rsbarang2.EOF Then


stok.Caption = rsbarang!stok & " " & rsbarang!satuan

Else
stok.Caption = rsbarang!stok - rsbarang2!stok
End If

disp.Text = Str(rsbarang!diskon)
If Text1.Enabled = False Then
Text1_KeyPress (13)
Else
Text1.SetFocus
End If
'Else
'LV1.SelectedItem.SubItems(6) = val(LV1.SelectedItem.SubItems(6)) + 1
'LV1.SelectedItem.SubItems(8) = val(LV1.SelectedItem.SubItems(6)) * val(LV1.SelectedItem.SubItems(5)) - val(LV1.SelectedItem.SubItems(7))
'ttl
'ttl_item
'diskons
'ttl2
'ttlb
'text4.SetFocus
'text4.Text = ""
'End If

End Sub
Sub ttl_item()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(6))
Next I
text7.Text = sum

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

sql = "select * from pelanggan where nama <>'' order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst
 Do While Not rsplg.EOF
customer.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  rsplg.MoveFirst
  End If

rsplg.Close

  End Sub

Private Sub text4_GotFocus()
If Edit = True Then
kete.Caption = "Rubah harga jual lalu tekan enter bila selesai atau isi diskon terlebih dahulu lalu enter"
Else
If LV1.ListItems.count = 0 Then
kete.Caption = "Tekan f3 untuk mencari barang.Atau ketik sebagian huruf(jangan semua) lalu enter lalu pilih dengan panah atas bawah terus enter atau dobel klik"
Else
kete.Caption = "Tekan f3 untuk mencari barang.Atau ketik sebagian huruf(jangan semua) lalu enter lalu pilih dengan panah atas bawah terus enter atau dobel klik.Tekan f5 bila transaksi selesai,atau enter bila ada diskon tambahan."

End If
End If

If text4.Text = "" And val(text7.Text) <> "0" Then
ket2.Caption = "Tekan enter bila telah selesai"
End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" And val(text7.Text) <> "0" Then
txtdiskt.SetFocus

Else
Dim stri As String

Set rse1 = New Recordset
stri = Replace(text4.Text, "'", "''")

rse1.Open "select * from tblbarang where deskripsi='" & stri & "' or kode_brg='" & stri & "'", jual, adOpenStatic, adLockOptimistic
If Not rse1.EOF Then
If letmin = False Then
If val(rse1!stok) <= 0 Then
MsgBox "Stok habis", vbCritical, judul
Exit Sub
End If
End If
text4_Click
Text1_KeyPress (13)
disp_KeyPress (13)
Set cari = LV1.FindItem(kode, 1, , 1)
LV1.SelectedItem = cari
Else
lvbrg2.Visible = True
lvbrg2.SetFocus
End If


End If
End If
End Sub

 Sub Command1_Click()
On Error Resume Next
nutupjual = False

Set rsbarang = New Recordset

   bilang = TerbilangDesimal(val(XPText6.Text) + val(ppn.Text)) + "rupiah"

If Option4.Value = True Then
Command1.Enabled = True
'If Check4.Value = Checked Then
'CetakData3
strukracikan
'Else
'Cetakprev
'End If
Else
If option5.Value = True Then

CetakData2
Else
If Option6.Value = True Then

CetakData4
End If
End If
End If
'baru.SetFocus

End Sub


Sub CetakData()
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
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
I = 1
mbaris = 0

'Do While i <= LV1.ListItems.count

    mhal = mhal + 1
    Printer.Print ; " "
    Printer.Print Tab(4); 'Form8.Caption;
    Printer.FontBold = False
    Printer.FontSize = 14
    Printer.Font = "Times New Roman"
    Printer.FontBold = True

    Printer.Print Tab(10); nama_toko2
        Printer.FontSize = 12
    Printer.Font = "Times New Roman"
    Printer.FontBold = True
Printer.FontItalic = True
    Printer.Print Tab(12); "GOOD YEAR";
    Printer.FontItalic = False

        Printer.FontSize = 12
    Printer.Font = "Times New Roman"
   'Printer.FontBold = True

    Printer.Print Tab(27); "Sentraservis "
        Printer.Print ; " "

Printer.FontItalic = False


    Printer.FontBold = False

    Printer.FontSize = 10
    Printer.Font = "Courier New"
    Printer.Print Tab(10); "Jl.Raya Rancaekek KM 22 ";
     Printer.Print Tab(55); "Tanggal";
         Printer.Print Tab(67); ": "; Format(tgll.Value, "dd MMMM yyyy");

    Printer.Print Tab(10); "Telp  : 022-7782780";

       

        Printer.Print Tab(55); "Kepada YTH ";

    Printer.Print ; Tab(55); customer.Text;

    


 '   Printer.Print ; Tab(52); "Alamat    : "; almt.Text;


    Printer.Print ; " "

        Printer.Print ; " "

    
    Printer.Print Tab(10); "No. Faktur";
    Printer.Print Tab(23); ": "; notr.Text;

    

    Printer.FontBold = False
mgrs = String$(76, "=")
mgrss = String$(76, "-")
    Printer.Print ; " "

Printer.Print Tab(10); mgrs
Printer.FontBold = False
Printer.Print Tab(10); "Jumlah";
Printer.Print Tab(19); "Nama Barang";
Printer.Print Tab(55); "Harga";
Printer.Print Tab(75); "Sub Total";
Printer.FontBold = False
Printer.Print Tab(10); mgrss
Do While I <= LV1.ListItems.count
   Set itm = LV1.ListItems.Item(I)
    mno = mno + 1
   Printer.Print Tab(3); RKanan(itm.SubItems(6), "###,###,###");
Printer.Print Tab(19); itm.SubItems(3);
Printer.Print Tab(52); RKanan(itm.SubItems(5), "###,###,###");

   Printer.Print Tab(73); RKanan(itm.SubItems(9), "###,###,###");
    mbaris = mbaris + 1
    I = I + 1
Loop
For j = 1 To (9 - LV1.ListItems.count)
      Printer.Print ; ""
   
      Next j

Printer.Print Tab(10); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = fale
'Printer.Print Tab(4); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = False
'Printer.Print Tab(5); mgrss
 Printer.FontSize = 10

Printer.Print Tab(10); "Tanda Terima";
 Printer.FontSize = 8
 Printer.Font = "Times New Roman"

Printer.Print Tab(52); "Perhatian:";
Printer.FontSize = 10
Printer.Font = "Courier New"

Printer.Print Tab(50); "Grand Total";
Printer.Print Tab(63); ":";
Printer.Print Tab(73); RKanan(XPText6.Text, "###,###,###");

If val(text5.Text) < val(XPText6.Text) Then
 Printer.FontSize = 8
Printer.Font = "Times New Roman"
Printer.Print Tab(52); "-Barang yang sudah dibeli/dipesan  ";
Printer.Print Tab(52); " tidak dapat ditukar/dikembalikan  ";
 Printer.FontSize = 10
Printer.Font = "Courier New"
Printer.Print Tab(50); "DP";
Printer.Print Tab(63); ":";

Printer.Print Tab(73); RKanan(val(text5.Text), "###,###,###");
 Printer.FontSize = 8
Printer.Font = "Times New Roman"
Printer.Print Tab(52); "-Ban vulkanisir pecah tidak diganti";
Printer.Print Tab(52); "-50% kembang copot/lepas ";
Printer.Print Tab(52); " diganti sesuai prosentasi pemakaian ";
 Printer.FontSize = 10
Printer.Font = "Courier New"
Printer.Print Tab(50); "Sisa ";
Printer.Print Tab(63); ":";
Printer.Print Tab(73); RKanan(val(XPText6.Text) - val(text5.Text), "###,###,###");

Else
 Printer.FontSize = 8
Printer.Font = "Times New Roman"
Printer.Print Tab(52); "-Barang yang sudah dibeli/dipesan  ";
Printer.Print Tab(52); " tidak dapat ditukar/dikembalikan  ";
 Printer.FontSize = 10
Printer.Font = "Courier New"
Printer.Print Tab(50); "Pembayaran";
Printer.Print Tab(63); ":";

Printer.Print Tab(73); RKanan(val(text5.Text), "###,###,###");
 Printer.FontSize = 8
Printer.Font = "Times New Roman"
Printer.Print Tab(52); "-Ban vulkanisir pecah tidak diganti";
Printer.Print Tab(52); "-50% kembang copot/lepas ";
Printer.Print Tab(52); " diganti sesuai prosentasi pemakaian ";
 Printer.FontSize = 10
Printer.Font = "Courier New"
If val(text5.Text) > val(XPText6.Text) Then
Printer.Print Tab(50); "Kembalian ";
Printer.Print Tab(63); ":";
Printer.Print Tab(73); RKanan(val(text5.Text) - val(XPText6.Text), "###,###,###");
End If
End If


Printer.Print ; " "
Printer.Print ; " "

 Printer.FontSize = 8

Printer.Print Tab(38); "Syarat pembayaran:";

 Printer.FontSize = 10

Printer.Print Tab(10); "_______________";
     Printer.FontSize = 8


Printer.FontBold = False
Printer.FontItalic = False
If mbaris >= 9 Then
Printer.NewPage

End If

'Loop
Printer.EndDoc
Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If
baru.SetFocus

End Sub
Sub CetakData2()
On Error Resume Next
Dim pj As Integer
pj = 56 - Len(almt)
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\invoice.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{penjualan.no_penjualan}='" & notr.Text & "'"
    .Formulas(0) = "almt='" & almt & "' + Chr(13) +'" & almt2 & "'"
    .Formulas(1) = "nama='" & nama_toko & "'"
    .Formulas(2) = "nama2='" & nama_toko2 & "'"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
        If Check4.Value = Checked Then
        .Destination = crptToPrinter
        End If


  .Action = 1
End With
'Me.Hide


End Sub
Sub cetakjalan()
On Error Resume Next
Dim pj As Integer
pj = 31 - Len(almt)
With CrystalReport1
.Reset

  .ReportFileName = serperreport & "\jalan2.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.Formulas(0) = "almt='" & almt & "' + Chr(13) +'" & almt2 & "'"
    .Formulas(1) = "nama='" & nama_toko & "'"
    .SelectionFormula = "{penjualan.no_penjualan}='" & notr.Text & "'"
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
                '.Destination = crptToPrinter

  .Action = 1
End With
Me.Hide

End Sub
Sub CetakData3()
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
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
I = 1
Do While I <= LV1.ListItems.count
    mbaris = 1
    mhal = mhal + 1
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print Tab(4); 'Form8.Caption;
    Printer.FontBold = False
    Printer.FontSize = 10
        Printer.FontSize = 8
Printer.Print Tab(3); nama_toko;
    Printer.Print Tab(3); almt;
        Printer.Print Tab(3); almt2;

    Printer.Print Tab(3); "Tanggal";
    Printer.Print Tab(15); ": "; Label10.Caption;
    Printer.Print Tab(3); "Jam";
    Printer.Print Tab(15); ": "; Format(Now, "hh:mm:ss");
    Printer.Print Tab(3); "No Bukti";
        Printer.Print Tab(15); ": "; notr.Text;
    Printer.Print Tab(3); "Pelanggan";
    Printer.Print Tab(15); ": "; customer.Text;

    Printer.Print Tab(3); "Kasir";

        Printer.Print Tab(15); ": "; kasir.Caption;

    Printer.FontBold = False
    Printer.Print ; " "
    Printer.Print ; " "
mgrs = String$(45, "=")
mgrss = String$(40, "-")
Printer.Print Tab(3); mgrss
Printer.FontBold = False
Printer.Print Tab(3); "No";

Printer.Print Tab(8); "Nama Barang";
Printer.Print Tab(31); "Harga";


'Printer.Print Tab(65); "Jumlah";


Printer.FontBold = False
Printer.Print Tab(3); mgrss
mbaris = 0
Do While I <= LV1.ListItems.count And mbaris < 60
   Set itm = LV1.ListItems.Item(I)
    mno = mno + 1
        Printer.Print Tab(3); I; Space(3); itm.SubItems(3);

    Printer.Print Tab(8); itm.SubItems(6) & ""; itm.SubItems(4); " X "; Format(itm.SubItems(5), "###,###,###");
    Printer.Print Tab(25); RKanan(itm.SubItems(9), "###,###,###");
    sql = "select * from tblbarang where kode_brg='" & itm.SubItems(2) & "'"
    Set RS = New Recordset
    Set RS = jual.Execute(sql)
    If val(itm.SubItems(7)) <> 0 Then
    Printer.Print Tab(8); "diskon " + itm.SubItems(7);
    End If
RS.Close
    mbaris = mbaris + 1
    I = I + 1
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
If val(dtmbh.Text) > 0 Then
Printer.Print Tab(3); "Total";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(gtot.Text) - val(XPText2.Text), "###,###,###");

Printer.Print Tab(3); "Diskon";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(dtmbh.Text, "###,###,###");
End If
If val(ppn.Text) > 0 Then
Printer.Print Tab(3); "PPN";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(ppn.Text, "###,###,###");

End If
Printer.Print Tab(3); "Grand Total";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(XPText6.Text) + val(ppn.Text), "###,###,###");
Printer.Print Tab(3); "Tunai";
Printer.Print Tab(16); ":";

Printer.Print Tab(25); RKanan(text5.Text, "###,###,###");
Printer.Print Tab(3); mgrss

Printer.Print Tab(3); "Kembalian ";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(text5.Text) - (val(XPText6.Text) + val(ppn.Text)), "###,###,###");
    Printer.Print Tab(5); "";
    Printer.Print Tab(5); "";
    Printer.Print Tab(5); "";
    Printer.Print ; " "
Printer.Print Tab(3); "Barang yang sudah dibeli";
Printer.Print Tab(3); "tidak dapat ditukar/dikembalikan.";
Printer.Print Tab(10); "********Terima Kasih********";
Printer.Print ; " "
Printer.Print ; " "
Printer.Print Tab(3); cttn1;
Printer.Print Tab(3); cttn2;
Printer.Print Tab(3); cttn3;

Printer.Print Tab(3); cttn4;
Printer.Print Tab(3); cttn5;
Printer.Print Tab(3); cttn6;
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
Sub strukracikan()
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
Dim mgrs, mgrss As String
If IsPrinterInstalled Then
Printer.Print "";
Printer.Print "";
'Printer.Font = "Courier New"
Printer.ForeColor = &H0&
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = 0
Printer.CurrentY = 0
mno = 0
mhal = 0
I = 1
    mbaris = 1
    mhal = mhal + 1
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print Tab(4); 'Form8.Caption;
    Printer.FontBold = False
    Printer.FontSize = 10
        Printer.FontSize = 8
Printer.Print Tab(3); nama_toko;
    Printer.Print Tab(3); almt;
        Printer.Print Tab(3); almt2;

    Printer.Print Tab(3); "Tanggal";
    Printer.Print Tab(15); ": "; Label10.Caption;
    Printer.Print Tab(3); "Jam";
    Printer.Print Tab(15); ": "; Format(Now, "hh:mm:ss");
    Printer.Print Tab(3); "No Bukti";
        Printer.Print Tab(15); ": "; notr.Text;
    Printer.Print Tab(3); "Pelanggan";
    Printer.Print Tab(15); ": "; customer.Text;

    Printer.Print Tab(3); "Kasir";

        Printer.Print Tab(15); ": "; kasir.Caption;

    Printer.FontBold = False
    Printer.Print ; " "
    Printer.Print ; " "
mgrs = String$(45, "=")
mgrss = String$(40, "-")
Printer.Print Tab(3); mgrss
Printer.FontBold = False
Printer.Print Tab(3); "No";

Printer.Print Tab(8); "Nama Barang";
Printer.Print Tab(31); "Harga";


'Printer.Print Tab(65); "Jumlah";


Printer.FontBold = False
Printer.Print Tab(3); mgrss
mbaris = 0
'no racikan
Set rsd = New Recordset
rsd.Open "select d.*,deskripsi,p.jenis from detiljual d,tblbarang b,penjualan p where b.kode_brg=d.kode_brg and p.no_penjualan=d.no_penjualan and d.no_penjualan='" & notr.Text & "' and ketr='umum'", jual, adOpenStatic, adLockOptimistic
   If Not rsd.EOF Then
   rsd.MoveFirst
   End If
Do While I <= rsd.RecordCount And mbaris < 60
    mno = mno + 1
    Printer.Print Tab(3); I; Space(3); rsd!deskripsi;
    If rsd!jenis <> "resep" Then
    Printer.Print Tab(8); rsd!jumlah_brg & ""; rsd!satuan; " X "; Format(rsd!harga_jual + (rsd!ppn / rsd!jumlah_brg), "###,###,###");
    End If
    Printer.Print Tab(25); RKanan(rsd!total, "###,###,###");
    
    If val(rsd!diskon) <> 0 Then
    Printer.Print Tab(8); "diskon " & rsd!diskon;
    End If

    mbaris = mbaris + 1
    I = I + 1
    rsd.MoveNext
Loop
'racikan
Set rsd = New Recordset
rsd.Open "select ketr,sum(total) total,coalesce(sum(diskon),0) diskon from detiljual  where  no_penjualan='" & notr.Text & "' and ketr!='umum' group by ketr", jual, adOpenStatic, adLockOptimistic
  If Not rsd.EOF Then
   rsd.MoveFirst
   End If
Do While Not rsd.EOF And mbaris < 60
    mno = mno + 1
    Printer.Print Tab(3); I; Space(3); rsd!ketr;

    Printer.Print Tab(8); "1 X "; Format(rsd!total, "###,###,###");
    Printer.Print Tab(25); RKanan(rsd!total, "###,###,###");
    
    If val(rsd!diskon) <> 0 Then
    Printer.Print Tab(8); "diskon " + rsd!diskon;
    End If

    mbaris = mbaris + 1
    I = I + 1
    rsd.MoveNext
Loop
Printer.Print Tab(3); mgrss
If val(dtmbh.Text) > 0 Then
Printer.Print Tab(3); "Total";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(gtot.Text) - val(XPText2.Text), "###,###,###");

Printer.Print Tab(3); "Diskon";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(dtmbh.Text, "###,###,###");
End If

Printer.Print Tab(3); "Grand Total";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(XPText6.Text) + val(txttuslah.Text), "###,###,###");


Printer.Print Tab(3); "Tunai";
Printer.Print Tab(16); ":";

Printer.Print Tab(25); RKanan(text5.Text, "###,###,###");
Printer.Print Tab(3); mgrss

Printer.Print Tab(3); "Kembalian ";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(val(text5.Text) - (val(XPText6.Text) + val(txttuslah.Text)), "###,###,###");
    Printer.Print Tab(5); "";
    Printer.Print Tab(5); "";
    Printer.Print Tab(5); "";
    If val(ppn.Text) > 0 Then
Printer.Print Tab(3); "(PPN)";
Printer.Print Tab(16); ":";
Printer.Print Tab(25); RKanan(ppn.Text, "###,###,###");

End If

    Printer.Print ; " "
Printer.Print Tab(3); "Barang yang sudah dibeli";
Printer.Print Tab(3); "tidak dapat ditukar/dikembalikan.";
Printer.Print Tab(10); "********Terima Kasih********";
Printer.Print ; " "
Printer.Print ; " "
Printer.Print Tab(3); cttn1;
Printer.Print Tab(3); cttn2;
Printer.Print Tab(3); cttn3;

Printer.Print Tab(3); cttn4;
Printer.Print Tab(3); cttn5;
Printer.Print Tab(3); cttn6;
Printer.FontBold = False
Printer.FontItalic = False
If mbaris >= 30 Then
Printer.NewPage
End If
Printer.EndDoc

Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

End Sub

Sub Cetakprev()
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
Dim mgrs, mgrss As String
frmprev.Show
If IsPrinterInstalled Then
frmprev.Print "";
frmprev.Print "";
frmprev.Font = "Courier New"
frmprev.ForeColor = &H0&
frmprev.FontSize = 6
frmprev.FontBold = False
frmprev.CurrentX = 0
frmprev.CurrentY = 0
mno = 0
mhal = 0
I = 1
Do While I <= LV1.ListItems.count
    mbaris = 1
    mhal = mhal + 1
    frmprev.Print ; " "
    frmprev.Print ; " "
    frmprev.Print ; " "
    frmprev.Print Tab(4); 'Form8.Caption;
    frmprev.FontBold = False
    frmprev.FontSize = 10
        frmprev.FontSize = 8
frmprev.Print Tab(3); nama_toko;
    frmprev.Print Tab(3); almt;
        frmprev.Print Tab(3); almt2;

    frmprev.Print Tab(3); "Tanggal";
    frmprev.Print Tab(15); ": "; Label10.Caption;
    frmprev.Print Tab(3); "Jam";
    frmprev.Print Tab(15); ": "; Format(Now, "hh:mm:ss");
    frmprev.Print Tab(3); "No Bukti";
        frmprev.Print Tab(15); ": "; notr.Text;
    frmprev.Print Tab(3); "Pelanggan";
    frmprev.Print Tab(15); ": "; customer.Text;

    frmprev.Print Tab(3); "Kasir";

        frmprev.Print Tab(15); ": "; kasir.Caption;

    frmprev.FontBold = False
    frmprev.Print ; " "
    frmprev.Print ; " "
mgrs = String$(45, "=")
mgrss = String$(40, "-")
frmprev.Print Tab(3); mgrss
frmprev.FontBold = False
frmprev.Print Tab(3); "No";

frmprev.Print Tab(8); "Nama Barang";
frmprev.Print Tab(31); "Harga";


'frmprev.print Tab(65); "Jumlah";


frmprev.FontBold = False
frmprev.Print Tab(3); mgrss
mbaris = 0
'no racikan
Set rsd = New Recordset
rsd.Open "select d.*,deskripsi from detiljual d,tblbarang b where b.kode_brg=d.kode_brg and no_penjualan='" & notr.Text & "' and ketr='umum'", jual, adOpenStatic, adLockOptimistic
   If Not rsd.EOF Then
   rsd.MoveFirst
   End If
Do While I <= rsd.RecordCount And mbaris < 60
    mno = mno + 1
    frmprev.Print Tab(3); I; Space(3); rsd!deskripsi;

   frmprev.Print Tab(8); rsd!jumlah_brg & ""; rsd!satuan; " X "; Format(rsd!harga_jual, "###,###,###");
    frmprev.Print Tab(25); RKanan(rsd!total, "###,###,###");
    
    If val(rsd!diskon) <> 0 Then
    Printer.Print Tab(8); "diskon " + rsd!diskon;
    End If

    mbaris = mbaris + 1
    I = I + 1
    rsd.MoveNext
Loop
'racikan
Set rsd = New Recordset
rsd.Open "select ketr,sum(total) total,coalesce(sum(diskon),0) diskon from detiljual  where  no_penjualan='" & notr.Text & "' and ketr!='umum' group by ketr", jual, adOpenStatic, adLockOptimistic
  If Not rsd.EOF Then
   rsd.MoveFirst
   End If
Do While Not rsd.EOF And mbaris < 60
    mno = mno + 1
    frmprev.Print Tab(3); I; Space(3); rsd!ketr;

    frmprev.Print Tab(8); "1 X "; Format(rsd!total, "###,###,###");
    frmprev.Print Tab(25); RKanan(rsd!total, "###,###,###");
    
    If val(rsd!diskon) <> 0 Then
    frmprev.Print Tab(8); "diskon " + rsd!diskon;
    End If

    mbaris = mbaris + 1
    I = I + 1
    rsd.MoveNext
Loop
frmprev.Print Tab(3); mgrss
'frmprev.print ; " "
'frmprev.FontSize = 8
'frmprev.FontBold = fale
'frmprev.print Tab(4); mgrss
'frmprev.print ; " "
'frmprev.FontSize = 8
'frmprev.FontBold = False
'frmprev.print Tab(5); mgrss
If val(dtmbh.Text) > 0 Then
frmprev.Print Tab(3); "Total";
frmprev.Print Tab(16); ":";
frmprev.Print Tab(25); RKanan(val(gtot.Text) - val(XPText2.Text), "###,###,###");

frmprev.Print Tab(3); "Diskon";
frmprev.Print Tab(16); ":";
frmprev.Print Tab(25); RKanan(dtmbh.Text, "###,###,###");
End If
If val(ppn.Text) > 0 Then
frmprev.Print Tab(3); "PPN";
frmprev.Print Tab(16); ":";
frmprev.Print Tab(25); RKanan(ppn.Text, "###,###,###");

End If
frmprev.Print Tab(3); "Grand Total";
frmprev.Print Tab(16); ":";
frmprev.Print Tab(25); RKanan(val(XPText6.Text) + val(ppn.Text), "###,###,###");
frmprev.Print Tab(3); "Tunai";
frmprev.Print Tab(16); ":";

frmprev.Print Tab(25); RKanan(text5.Text, "###,###,###");
frmprev.Print Tab(3); mgrss

frmprev.Print Tab(3); "Kembalian ";
frmprev.Print Tab(16); ":";
frmprev.Print Tab(25); RKanan(val(text5.Text) - (val(XPText6.Text) + val(ppn.Text)), "###,###,###");
    frmprev.Print Tab(5); "";
    frmprev.Print Tab(5); "";
    frmprev.Print Tab(5); "";
    frmprev.Print ; " "
frmprev.Print Tab(3); "Barang yang sudah dibeli";
frmprev.Print Tab(3); "tidak dapat ditukar/dikembalikan.";
frmprev.Print Tab(10); "********Terima Kasih********";
frmprev.Print ; " "
frmprev.Print ; " "
frmprev.Print Tab(3); cttn1;
frmprev.Print Tab(3); cttn2;
frmprev.Print Tab(3); cttn3;

frmprev.Print Tab(3); cttn4;
frmprev.Print Tab(3); cttn5;
frmprev.Print Tab(3); cttn6;
frmprev.FontBold = False
frmprev.FontItalic = False
If mbaris >= 30 Then
'frmprev.NewPage
End If
Loop
'frmprev.EndDoc

Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

End Sub

Sub CetakData4()
On Error GoTo erol
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
Dim mgrs, mgrss As String
Open "LPT1" For Output As #1

If IsPrinterInstalled Then
Print #1, "";
Print #1, "";
Printer.Font = "Courier New"
Printer.ForeColor = &H0&
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = 0
Printer.CurrentY = 0
mno = 0
mhal = 0
Print #1, Chr(27) & Chr(33) & Chr(1);
I = 1
    mbaris = 1
    mhal = mhal + 1
    Print #1, ; " "
    Print #1, ; " "
    Print #1, ; " "
    Print #1, Tab(4); 'Form8.Caption;
    Printer.FontBold = False
    Printer.FontSize = 10
        Printer.FontSize = 8
Print #1, Tab(3); nama_toko;
    Print #1, Tab(3); almt;
        Print #1, Tab(3); almt2;

    Print #1, Tab(3); "Tanggal";
    Print #1, Tab(15); ": "; Label10.Caption;
    Print #1, Tab(3); "Jam";
    Print #1, Tab(15); ": "; Format(Now, "hh:mm:ss");
    Print #1, Tab(3); "No Bukti";
        Print #1, Tab(15); ": "; notr.Text;
    Print #1, Tab(3); "Pelanggan";
    Print #1, Tab(15); ": "; customer.Text;

    Print #1, Tab(3); "Kasir";

        Print #1, Tab(15); ": "; kasir.Caption;

    Printer.FontBold = False
    Print #1, ; " "
    Print #1, ; " "
mgrs = String$(45, "=")
mgrss = String$(40, "-")
Print #1, Tab(3); mgrss
Printer.FontBold = False
Print #1, Tab(3); "No";

Print #1, Tab(8); "Nama Barang";
Print #1, Tab(31); "Harga";


'Print #1, Tab(65); "Jumlah";


Printer.FontBold = False
Print #1, Tab(3); mgrss

Set rsd = New Recordset
rsd.Open "Select * from detiljual,tblbarang where detiljual.kode_brg=tblbarang.kode_brg and detiljual.no_penjualan='" & notr.Text & "' order by tblbarang.deskripsi", jual, adOpenStatic, adLockOptimistic

mbaris = 0
Do While I <= rsd.RecordCount And mbaris < 60
   Set itm = LV1.ListItems.Item(I)
    mno = mno + 1
        Print #1, Tab(3); I; Space(3); rsd![deskripsi];

    Print #1, Tab(8); rsd![jumlah_brg] & " "; rsd![detiljual.satuan]; " X "; Format(rsd![detiljual.harga_jual], "###,###,###");
    Print #1, Tab(25); RKanan(rsd![total], "###,###,###");
    
    If val(rsd![detiljual.diskon]) <> 0 Then
    Print #1, Tab(8); "diskon " + Format(rsd![detiljual.diskon], "#,#");
    End If
    mbaris = mbaris + 1
    I = I + 1
    rsd.MoveNext
Loop
Print #1, Tab(3); mgrss
If val(dtmbh.Text) > 0 Then
Print #1, Tab(3); "Total";
Print #1, Tab(16); ":";
Print #1, Tab(25); RKanan(val(gtot.Text) - val(XPText2.Text), "###,###,###");

Print #1, Tab(3); "Diskon";
Print #1, Tab(16); ":";
Print #1, Tab(25); RKanan(dtmbh.Text, "###,###,###");
End If
If val(ppn.Text) > 0 Then
Print #1, Tab(3); "PPN";
Print #1, Tab(16); ":";
Print #1, Tab(25); RKanan(ppn.Text, "###,###,###");

End If
Print #1, Tab(3); "Grand Total";
Print #1, Tab(16); ":";
Print #1, Tab(25); RKanan(val(XPText6.Text) + val(ppn.Text), "###,###,###");
Print #1, Tab(3); "Tunai";
Print #1, Tab(16); ":";

Print #1, Tab(25); RKanan(text5.Text, "###,###,###");
Print #1, Tab(3); mgrss

Print #1, Tab(3); "Kembalian ";
Print #1, Tab(16); ":";
Print #1, Tab(25); RKanan(val(text5.Text) - (val(XPText6.Text) + val(ppn.Text)), "###,###,###");
    Print #1, Tab(5); "";
    Print #1, Tab(5); "";
    Print #1, Tab(5); "";
    Print #1, ; " "
Print #1, ; " "

Print #1, Tab(11); "********Terima Kasih********";
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, ; " "
Print #1, Chr(27) & Chr(33) & Chr(0);
   Close #1
Printer.FontBold = False
Printer.FontItalic = False
If mbaris >= 30 Then
Printer.NewPage
End If
Printer.EndDoc

Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

erol:
 If err.Description <> vbNullString Then
 MsgBox "Bukan printer port paralel"
 Exit Sub
End If

End Sub
Sub BukaDrawer()
On Error Resume Next
Open "LPT1" For Output As #1
Print #1, Chr$(27); Chr$(112); Chr$(0)
Close #1
End Sub
Private Sub Command2_Click()
kosong
kosong2
awal
notr.Text = ""
baru.Caption = "&Baru"
baru.SetFocus
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

cekaktip
 Check1.Value = GetSetting("apotekbaleendah", "transaksi", "Check1.value", Checked)
 Check2.Value = GetSetting("apotekbaleendah", "transaksi", "Check2.value", Unchecked)
Check3.Value = GetSetting("apotekbaleendah", "transaksi", "Check3.value", Unchecked)
Check4.Value = GetSetting("apotekbaleendah", "transaksi", "Check4.value", Unchecked)
Check5.Value = GetSetting("apotekbaleendah", "transaksi", "Check5.value", Checked)
Check6.Value = GetSetting("apotekbaleendah", "transaksi", "Check6.value", Unchecked)
Check6_Click
hpju = GetSetting("apotekbaleendah", "huruff", "text2.text", "")
txttop.Text = GetSetting("apotekbaleendah", "transaksi", "txttop.text", "30")
isipromosi = GetSetting("apotekbaleendah", "frmpromosi", "txtisi.text", "")
Option4.Value = True
  If hpju = "" Then
hpju = "PJ"
End If
pjgh = Len(hpju)

If Check3.Value = Checked Then
txtppn.Enabled = True
Else
txtppn.Enabled = False
End If
hari = 0
 Option6.Value = GetSetting("apotekbaleendah", "transaksi", "Option6.value", False)
Option4.Value = GetSetting("apotekbaleendah", "transaksi", "option4.value", False)
option5.Value = GetSetting("apotekbaleendah", "transaksi", "option5.value", True)
 Option1.Value = GetSetting("apotekbaleendah", "transaksi", "option1.value", False)
Option2.Value = GetSetting("apotekbaleendah", "transaksi", "option2.value", False)
Option3.Value = GetSetting("apotekbaleendah", "transaksi", "option3.value", True)
lispaket
 Tab1.Tab = 0
pilih = ""
 tgll.Value = Now
'kbrg
dbgridplg
cust
Text1.Enabled = rhj
kasir.Caption = Mnutama.StatusBar1.Panels(8).Text
pilih = ""
Text1.Enabled = IIf(rshusus.Fields(19) = False, False, True)
cmdhps.Visible = IIf(rshusus.Fields(18) = False, False, True)
'Ketengah Me
Label10.Caption = Format(Now, "dd-mm-YYYY")
    Skinpath = App.Path & "\skin\B-Studio.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
awal
'Option4.Value = True

End Sub
Sub cekaktip()
If aktipasi = True Then Exit Sub
Set RS = New Recordset
RS.Open "select coalesce(count(no_penjualan),0) as jum from penjualan", jual, adOpenStatic, adLockOptimistic
If RS!jum >= 10 Then
MsgBox "Masih demo,maksimal 10x transaksi " & Chr(13) & " Silahkan menghubungi 082116969006 (Pak Mudiman) ", vbCritical, judul
End
Exit Sub
End If
End Sub
Private Sub dbgrid2_DblClick()
On Error Resume Next
If Edit = False Then
customer.Text = Dbgrid2.Columns.Item(1)
Tab1.Tab = 0
text4.SetFocus
End If
End Sub


Sub dbgrid()
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
Sub dbgridplg()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from pelanggan order by nama"
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
        l.SubItems(5) = ![jumlah_piutang]

    .MoveNext
    Loop
End With


End Sub
Sub dbgridplg2()
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
        l.SubItems(5) = ![jumlah_piutang]

    .MoveNext
    Loop
End With


End Sub

Sub dbgridtrans()
On Error Resume Next

Set rstrans = New Recordset


sql = "select penjualan.no_penjualan,penjualan.tanggal,penjualan.jumlah,penjualan.total_diskon,penjualan.total,penjualan.kasir,penjualan.id_pelanggan,pelanggan.nama from penjualan,pelanggan where penjualan.id_pelanggan=pelanggan.id_pelanggan order by no_penjualan desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_penjualan]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![id_pelanggan]
                              
                l.SubItems(4) = ![nama]
                l.SubItems(5) = ![jumlah]
                l.SubItems(6) = ![total_diskon]
  l.SubItems(7) = ![total]
    l.SubItems(8) = ![kasir]

    .MoveNext
    Loop
End With

End Sub

Sub dbgridtrans2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select penjualan.no_penjualan,penjualan.tanggal,penjualan.jumlah,penjualan.total_diskon,penjualan.total,penjualan.kasir,penjualan.id_pelanggan,pelanggan.nama from penjualan,pelanggan where (penjualan.no_penjualan like'" & text8.Text & "%' or penjualan.id_pelanggan like'" & text8.Text & "%' or pelanggan.nama like'" & text8.Text & "%')and penjualan.id_pelanggan=pelanggan.id_pelanggan order by no_penjualan desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_penjualan]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![id_pelanggan]

                l.SubItems(4) = ![nama]
                l.SubItems(5) = ![jumlah]
                l.SubItems(6) = ![total_diskon]
  l.SubItems(7) = ![total]
    l.SubItems(8) = ![kasir]

    .MoveNext
    Loop
End With

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

            If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
            End If

    'If KeyAscii = 13 And (val(stok.Caption) > 0) Then
        'If Text2.Text = "" And (val(stok.Caption) > 0) Then
        If KeyAscii = 13 Then
        If Text2.Text = "" Then
        Text2.Text = "1"
        End If
    disp.SetFocus
    text4.SetFocus

    isigrid
    ttl
    ttl_item
    diskons
    ttl2
    ttlb
    kosong
    Else
    If KeyAscii = 13 Then

        If Edit = True Then
        Text1_KeyPress (13)
        disp.Text = ""
        diskon.Text = ""
        End If

    End If



End If


End Sub
Private Sub MSComm1_OnComm()

Dim CheckMyScan As String

Dim CheckForCR As String

Dim MyText As String

Dim CountMe As Integer

Dim counter As Integer

Dim Number As Integer

On Error GoTo Mscomm11:

If MSComm1.CommEvent = 2 And MSComm1.InBufferCount > 0 Then

CheckMyScan = MSComm1.Input

MyText = CheckMyScan

text4.Text = CheckMyScan

CountMe = Len(text4.Text)

Number = 0

Do Until counter = CountMe

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

counter = counter + 1

Loop

End If

Exit Sub

Mscomm11:

MsgBox "A error in reading this bar code", vbOKOnly, "POS"

End Sub

Sub ttl()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(6)) * val(LV1.ListItems(I).SubItems(5))
Next I
gtot = sum
XPText6.Text = val(gtot.Text) - (sum + val(dtmbh.Text))

'total.Text = Format(val(XPText6.Text) + val(ppn.Text), "#,#0.#0")
End Sub
Sub ttlstok()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(6))
gtot = sum
Next I
'total.Text = Format(sum + val(ppn.Text), "#,#0.#0")

End Sub


Sub ttlb()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(6)) * val(LV1.ListItems(I).SubItems(10))
Next I
thpj.Caption = sum
End Sub

Private Sub Text2_Change()
On Error Resume Next
'If Edit = True Then Exit Sub
If satuan.Text <> stn_baku Then
        Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic

stokjual = val(Text2.Text) * rsstn!konversi
Else
stokjual = val(Text2.Text)
End If
  If letmin = False Then
    If val(stokjual) <= val(stok) Then
    diskon.Text = val(rsbarang!diskon) / 100 * val(Text1.Text) * val(Text2.Text)
    text3.Text = val(Text1.Text) * val(Text2.Text) - val(diskon.Text)
    Else
        If Edit = False And Text2.Text <> "" Then
        MsgBox "Lewat dari stok", vbCritical
        Text2 = ""
        text3 = ""
        End If

    Exit Sub
    End If
  Else
  diskon.Text = val(rsbarang!diskon) / 100 * val(Text1.Text) * val(Text2.Text)
  text3.Text = val(Text1.Text) * val(Text2.Text) - val(diskon.Text)
  End If
ket.Caption = "TOTAL:"
End Sub
Private Sub info_Click()
MsgBox "Pengisian kurs terakhir 1 US$=Rp" & matu & ""
End Sub
Sub isigrid()
On Error GoTo erol
If Text2.Text = "" Then
Text2.Text = "1"
End If

     If satuan.Text <> stn_baku Then
      Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic
hasil = val(Text2.Text) * rsstn!konversi
Else
hasil = val(Text2.Text)
End If
Set rsbarang2 = New Recordset
rsbarang2.Open "select * from stok where kode_brg='" & kode & "'", jual, adOpenStatic, adLockOptimistic
If rsbarang2.EOF Then
jual.Execute "insert into stok values('" & kode & "','" & hasil & "')"
Else
rss = rsbarang2!stok + hasil

jual.Execute "update stok set stok=" & rss & " where kode_brg='" & kode & "'"
End If



Set cari = LV1.FindItem(kode, 1, , 1)
LV1.SelectedItem = cari

'And LV1.SelectedItem.SubItems(4) <> Satuan.Text Then
If cari Is Nothing Then
isi
    Else
    If LV1.SelectedItem.SubItems(4) <> satuan.Text Then
    isi
    Else
    LV1.SelectedItem.SubItems(6) = LV1.SelectedItem.SubItems(6) + val(Text2.Text)
    LV1.SelectedItem.SubItems(7) = (val(diskon.Text) / (val(Text2.Text) * val(Text1.Text))) * (LV1.SelectedItem.SubItems(6) * LV1.SelectedItem.SubItems(5))
    LV1.SelectedItem.SubItems(9) = LV1.SelectedItem.SubItems(5) * LV1.SelectedItem.SubItems(6) - LV1.SelectedItem.SubItems(7)
    End If
    
    End If
    Set cari = LV1.FindItem(kode, 1, , 1)
LV1.SelectedItem = cari


    
erol:
 If err.Description <> vbNullString Then
 lvbrg2.Visible = True
 lvbrg2.SetFocus
 'MsgBox "Barang tidak terdaftar"
 'Text4.SetFocus
 Exit Sub
End If

    

End Sub
Sub isi()
    Set butir = LV1.ListItems.Add
    With butir
           .SubItems(1) = LV1.ListItems.count & "."

    .SubItems(2) = kode
    .SubItems(3) = nama.Caption
    .SubItems(4) = satuan.Text

    .SubItems(5) = Text1.Text
    .SubItems(6) = Text2.Text
    .SubItems(7) = diskon.Text
    .SubItems(8) = ppn.Text
    .SubItems(9) = text3.Text
    .SubItems(10) = hpj
    .SubItems(11) = "umum"
    
    End With

End Sub
Private Sub hapus_Click()
If LV1.SelectedItem.SubItems(11) <> "umum" Then
Command2_Click
Exit Sub
End If
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
ttl_item
ttlb
diskons
ttl2

For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x
kosong
End Sub

Sub kosong()
text4.Text = ""
Text1.Text = ""
Text2 = ""
text3.Text = ""
nama.Caption = ""
stok.Caption = ""
diskon.Text = ""
dtmbh.Text = ""
satuan.Text = ""
stnc.Caption = ""
stnc2.Caption = ""
txtppn.Text = ""

ppn.Text = ""
Set Image1.Picture = Nothing

End Sub
Sub ttl2()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(9))
Next I
XPText6.Text = sum

total.Text = Format(sum, "#,#0.#0")
End Sub
Sub diskons()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(7))
Next I
XPText2.Text = sum

End Sub

Sub kosong2()
total.Text = ""
gtot = ""
text5 = ""
text7.Text = ""
txtppn.Text = ""
ppn.Text = ""
txtpo.Text = ""
txttuslah.Text = ""
disp.Text = ""
XPText6.Text = ""
XPText2.Text = ""
thpj.Caption = ""
satuan.Text = ""
text6.Text = ""
txtkmsp.Text = ""
txtkms.Text = ""
txtiddok.Text = ""
LV1.ListItems.Clear
End Sub

Private Sub Text4_LostFocus()

ket2.Caption = ""
End Sub

Private Sub Text5_Click()
If LV1.ListItems.count = 0 Then
MsgBox "Belum ada transaksi", vbCritical, judul
text4.SetFocus

Exit Sub
End If

bayar.Text1.Text = ""
tmplbyr
bayar.Text1.SetFocus

End Sub

Private Sub Text5_GotFocus()
'Text5_Click

ket2.Caption = "Masukkan nominal pembayaran konsumen"
End Sub
Sub tmplbyr()
nutupjual = False
bayar.Text1.Text = ""

bayar.Show
bayar.Text1.SetFocus

End Sub
 Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If text5.Text = "" And KeyAscii <> 13 Then
tmplbyr
End If
If KeyAscii = 13 Then
If Edit = True Then
palid
'hapus_faktur2
End If
proses
End If
End Sub
Sub proses()
Set RS = New Recordset
sql = "select id_pelanggan,id_sales from pelanggan where nama='" & customer.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
'id = RS!nama
idp = RS!id_pelanggan
idsls = IIf(IsNull(RS!id_sales) = True, "", RS!id_sales)
RS.Close
Else
If customer.Text = "" Then

idp = ""
idsls = ""
End If
End If

If val(text5.Text) <= val(XPText6.Text) Then
pembayaran = val(text5.Text)
Else
pembayaran = val(XPText6.Text)
End If
If gtot.Text <> "" Then
ket.Caption = "KEMBALIAN:"

kembali = val(text5) - (val(XPText6) + val(ppn.Text) + val(txttuslah.Text))
text6 = kembali

total.Text = Format(kembali, "#,#0.#0")



If val(text5.Text) < val(XPText6.Text) Then

Set RS = New Recordset
RS.Open "select * from pelanggan where id_pelanggan='" & idp & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
If customer.Text = "" Then
MsgBox "Pelanggan bebas tidak dapat berhutang"
'bayar.Show
text5.Text = ""
text6.Text = ""
total.Text = Format(val(XPText6.Text) + val(ppn.Text) + val(txttuslah.Text), "#,#0.#0")
text5.SetFocus
tmplbyr
Exit Sub
End If
Else
If RS!id_pelanggan = "bebas" Then
MsgBox "Pelanggan bebas tidak dapat berhutang"
text5.Text = ""
text6.Text = ""
total.Text = Format(val(XPText6.Text) + val(ppn.Text) + val(txttuslah.Text), "#,#0.#0")
text5.SetFocus
tmplbyr
Exit Sub
End If
End If
If MsgBox("Pembayaran kurang dari total penjualan,akan dimasukkan ke piutang sebesar  " & Format(-1 * text6.Text, "#,#0.#0") & "", vbOKCancel) = vbCancel Then
text5.Text = ""
text6.Text = ""
total.Text = ""
text5.SetFocus
tmplbyr
'Exit Sub
Else
'SkinLabel14.Visible = True
'DTPicker1.Visible = True
'DTPicker1.Value = tgll.Value + val(txttop.Text)
'DTPicker1.SetFocus
kidon
Exit Sub
End If
End If
If Option3.Value = True And txtiddok.Text = "" Then MsgBox "ID Dokter harus diisi", vbCritical: Exit Sub

If MsgBox("Simpan transaksi?", vbYesNo, "Tanya") = vbYes Then
pilih = ""
If val(text5.Text) > 0 Then
nutupjual = False
'tidak tampil tanya
pilih = "KAS"
proses2
nutupjual = True
'sudah tidak tampil tanya


    'tanya.Show
    Else
  proses2
  End If
End If
End If

End Sub
Sub proses2()
If Option1.Value = True Then
jnsjual = "umum"
Else
If Option2.Value = True Then
jnsjual = "dokter"
Else
jnsjual = "resep"
End If
End If
simpandata
dbgridplg
cust
Command1.Enabled = True
If Check2.Value = Checked Then

If MsgBox("Cetak surat jalan ?", vbYesNo) = vbYes Then
ThemedButton1_Click
End If
End If
If MsgBox("Cetak faktur ?", vbYesNo) = vbYes Then
Command1_Click
Else

awal
Command1.Enabled = True
MsgBox "Data penjualan berhasil disimpan"
baru.SetFocus
End If
End Sub
Sub awal()
notr.Enabled = False
text4.Enabled = False
Text2.Enabled = False
Command1.Enabled = False
Frame1.Enabled = True

Edit = True
End Sub
Sub simpandata()
On Error GoTo erol
If idp = "" And customer.Text <> "" Then
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
sql = "Select id_pelanggan from pelanggan order by id_pelanggan Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = hcus & "0001"
Else
j = val(Right(rsplg(0), 4))
idc = hcus + Format(Str(j + 1), "0000")

End If
If customer.Text <> "" Then
jual.Execute "insert into pelanggan(id_pelanggan,nama,jumlah_piutang) values('" & idc & "','" & customer.Text & "',0)"
Else
jual.Execute "insert into pelanggan(id_pelanggan,nama,jumlah_piutang) values('" & idc & "','Pelanggan bebas',0)"

End If
idp = idc
idsls = ""
Else
Set rse1 = New Recordset
rse1.Open "Select * from pelanggan where nama='Pelanggan bebas'", jual, adOpenStatic, adLockOptimistic

If rse1.EOF And customer.Text = "" Then
jual.Execute "insert into pelanggan(id_pelanggan,nama) values('bebas','Pelanggan bebas')"
idp = "bebas"
idsls = ""
Else
If Not rse1.EOF And customer.Text = "" Then
idp = "bebas"
idsls = ""
End If
End If
rse1.Close


End If



Set rstrans = New Recordset
sel = val(text5.Text) - val(XPText6.Text)

If val(text5.Text) = 0 Then
ket1 = "B"
kete2 = "BL"
Else
If val(text5.Text) >= val(XPText6.Text) Then
ket1 = "C"
kete2 = "L"
Else
ket1 = "CB"
kete2 = "BL"
End If
End If
If Option3.Value = True Then
sql = "insert into penjualan(No_penjualan,tanggal,total_diskon,kasir,id_pelanggan,id_sales,harga_pokok_jual,keterangan1,keterangan2,ppn,no_po,hari,cash,id_shift,jenis,id_dokter,komisi,tuslah) values('" & notr.Text & "','" & Format(tgll.Value, "YYYY-mm-dd") & "','" & val(dtmbh.Text) & "','" & kasir.Caption & "','" & idp & "','" & idsls & "','" & val(thpj.Caption) & "','" & ket1 & "','" & kete2 & "','" & val(ppn.Text) & "','" & txtpo.Text & "','" & val(txttop.Text) & "','" & _
val(text5.Text) & "','" & kodesip & "','" & jnsjual & "','" & txtiddok.Text & "','" & Format(txtkms.Text, Number) & "','" & Format(txttuslah.Text, Number) & "')"
'jual.Execute "insert into jurnal values(null,'" & Format(tgll.Value, "yyyy-mm-dd") & "','1.1','Kas','" & Format(txtkms.Text, Number) & "',0,'" & notr.Text & "','')"
'jual.Execute "insert into jurnal values(null,'" & Format(tgll.Value, "yyyy-mm-dd") & "','5.9','Komisi dokter',0,'" & Format(txtkms.Text, Number) & "','" & notr.Text & "','')"

Else
sql = "insert into penjualan(No_penjualan,tanggal,total_diskon,kasir,id_pelanggan,id_sales,harga_pokok_jual,keterangan1,keterangan2,ppn,no_po,hari,cash,id_shift,jenis) values('" & notr.Text & "','" & Format(tgll.Value, "YYYY-mm-dd") & "','" & val(dtmbh.Text) & "','" & kasir.Caption & "','" & idp & "','" & idsls & "','" & val(thpj.Caption) & "','" & ket1 & "','" & kete2 & "','" & val(ppn.Text) & "','" & txtpo.Text & "','" & val(txttop.Text) & "','" & _
val(text5.Text) & "','" & kodesip & "','" & jnsjual & "')"

End If

jual.Execute (sql)
For z = 1 To LV1.ListItems.count
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where kode_brg='" & LV1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
st = rsbarang!stok
If LV1.ListItems(z).SubItems(4) <> rsbarang!satuan Then
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & LV1.ListItems(z).SubItems(2) & "' and satuan='" & LV1.ListItems(z).SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic
ns = st - LV1.ListItems(z).SubItems(6) * val(rsstn!konversi)
sql = "insert into detiljual values('" & notr.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) / val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(6) * val(rsstn!konversi) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(8) & "','" & LV1.ListItems(z).SubItems(9) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "','utama','" & LV1.ListItems(z).SubItems(11) & "')"
jual.Execute (sql)

Else
ns = st - LV1.ListItems(z).SubItems(6)
sql = "insert into detiljual values('" & notr.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(5) & "','" & LV1.ListItems(z).SubItems(10) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(7) & "','" & LV1.ListItems(z).SubItems(8) & "','" & LV1.ListItems(z).SubItems(9) & "','" & LV1.ListItems(z).SubItems(6) & "','" & LV1.ListItems(z).SubItems(4) & "','utama','" & LV1.ListItems(z).SubItems(11) & "')"
jual.Execute (sql)

End If

    Next z
    If val(text5.Text) >= val(XPText6.Text) Then
    byr = val(XPText6.Text)
    Else
    byr = val(text5.Text)
    End If

  awal
pilih = ""
cmdtutup.Enabled = False
jual.Execute "Delete from stok"
Exit Sub
erol:
If err.Number = -2147217900 Then
MsgBox "Nomor faktur ganda", vbCritical, judul
End If
End Sub

Sub no_oto()
Dim j As Integer
Dim br As String
Set rstrans = New Recordset
sql = "Select no_penjualan from penjualan where no_penjualan like 'P-%'order by no_penjualan Desc"
Set rstrans = jual.Execute(sql)
If rstrans.EOF = True Then
notr.Text = "P-" + Format(Now, "YY-") + Format(Now, "MM-") + Format(Now, "dd-") + "001"

Else
j = val(Right(rstrans(0), 3))
br = "P-" + Format(Now, "YY-") + Format(Now, "MM") + "-" + Format(Now, "dd-") + Format(Str(j + 1), "000")
notr.Text = br
nmr = Str(Format(Mid(rstrans(0), 7, 2)))
thn = Str(Format(Mid(rstrans(0), 4, 2)))
txt = Format(Mid(No, 7, 2))
thnn = Format(Mid(No, 4, 2))
If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
notr.Text = "P-" + Format(Now, "YY-") + Format(Now, "MM-") + Format(Now, "dd-") + "001"
End If
End If

End Sub
Sub GetNumber()
On Error GoTo salah
    Dim counter As String * 11
    Dim Hitung As Integer
    Dim tgl As String
        tgl = Format(Now, "dd/mm/yyyy")

    A = hpju + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "%"
sql = "Select no_penjualan from penjualan where no_penjualan like '" & A & "' order by no_penjualan"
    Set rstrans = jual.Execute(sql)

    With rstrans
        If .RecordCount = 0 Then
            counter = hpju + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "001"
        Else
           .MoveLast
            If Left(![no_penjualan], pjgh + 6) <> hpju + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = hpju + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "001"
            Else
                Hitung = val(Right(!no_penjualan, 3)) + 1
               counter = hpju + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("000" & Hitung, 3)
            End If
        End If
        notr.Text = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description
End Sub

Private Sub Text5_LostFocus()
ket2.Caption = ""
End Sub

Private Sub Text8_Change()
dbgridtrans2
End Sub

Private Sub ThemedButton1_Click()
On Error Resume Next
nutupjual = False

cetakjalan
baru.SetFocus

End Sub



Private Sub ThemedButton4_Click()
calc.Show
End Sub

Private Sub Timer1_Timer()
jam.Caption = Format(Now, "hh:mm:ss")

End Sub

Private Sub total_Change()
If Option3.Value = True And ket.Caption <> "KEMBALIAN:" Then
txtkms.Text = Format(val(total.Text), Number) * (val(txtkmsp.Text)) * 0.01
End If
End Sub

Private Sub txtcari_Change()
dbgridcari

End Sub
Sub dbgridcari()
On Error Resume Next

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
  l.SubItems(7) = Format(![Harga_jual2], "#,#")
 l.SubItems(8) = Format(![Harga_jual3], "#,#")
    .MoveNext
    Loop
End With

End Sub
Sub dbgridcari2()
On Error Resume Next

Set rstrans = New Recordset

stri = Replace(text4.Text, "'", "''")

sql = "select * from tblbarang where kode_brg like '" & stri & "%' or deskripsi like '%" & stri & "%' order by deskripsi"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvbrg2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg2.ListItems.Add(, , lvbrg2.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                               
               
              
  
    .MoveNext
    Loop
End With

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

Private Sub txtdiskt_Change()
If txtdiskt.Text = "" Then
dtmbh.Text = ""
Exit Sub
End If
dtmbh.Text = (val(gtot.Text) - (sum + val(XPText2.Text))) * 0.01 * val(txtdiskt.Text)

End Sub

Private Sub txtdiskt_GotFocus()
kete.Caption = "Ketik diskon tambahan dalam diskon,lalu enter."

End Sub

Private Sub txtdiskt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Option3.Value = True Then
txttuslah.SetFocus
Else
text5.SetFocus
tmplbyr
End If
End If

End Sub

Private Sub txtkmsp_Change()
txtkms.Text = val(txtkmsp.Text) * 0.01 * val(Format(total.Text, Number))
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
customer.SetFocus
End If
End Sub

Private Sub txtppn_Change()
If val(txtppn.Text) > 100 Then
MsgBox "ERORR", , judul
txtppn.Text = ""
txtppn.SetFocus
Exit Sub
End If

ppn.Text = val(txtppn.Text) * 0.01 * (val(Text2.Text) * val(Text1.Text) - val(diskon.Text))
text3.Text = ""
text3.Text = val(Text2.Text) * val(Text1.Text) - val(diskon.Text) + val(ppn.Text)


End Sub

Private Sub txtppn_GotFocus()
kete.Caption = "Ketik PPN tambahan dalam diskon,lalu enter."

End Sub

Private Sub txtppn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Edit = True Then
Text1_KeyPress (13)
Else
text4.SetFocus
If hpj = val(Text1.Text) Then
MsgBox "Modal sama dengan harga jual", vbCritical, judul
Text1.SetFocus
Exit Sub
Else

End If

isigrid
ttl
ttl_item
diskons
ttl2
ttlb
kosong
End If
End If

End Sub

Private Sub txttop_Change()
SaveSetting "apotekbaleendah", "transaksi", "txttop.text", txttop.Text

End Sub

Private Sub txttuslah_Change()
total.Text = Format(val(XPText6.Text) + val(txttuslah.Text), "#,#0.#0")

End Sub

Private Sub txttuslah_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text5.SetFocus
tmplbyr
End If
End Sub

Private Sub XPText6_Change()
total.Text = Format(val(XPText6.Text) + val(ppn.Text) + val(txttuslah.Text), "#,#0.#0")

End Sub


