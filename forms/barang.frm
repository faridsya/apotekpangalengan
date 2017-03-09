VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Barang 
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   ClipControls    =   0   'False
   Icon            =   "barang.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   13365
   Begin VB.CommandButton Command4 
      Caption         =   "&Info stok gudang"
      Height          =   255
      Left            =   6360
      TabIndex        =   61
      Top             =   5400
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel hrfc 
      Height          =   255
      Left            =   3960
      OleObjectBlob   =   "barang.frx":324A
      TabIndex        =   51
      Top             =   6000
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   3000
      OleObjectBlob   =   "barang.frx":32A8
      TabIndex        =   50
      Top             =   6000
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   8040
      OleObjectBlob   =   "barang.frx":331C
      TabIndex        =   49
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Frame Frame 
      Height          =   5295
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   13095
      Begin VB.ComboBox cmbgol 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox cmbtipe 
         Height          =   315
         ItemData        =   "barang.frx":33AA
         Left            =   1440
         List            =   "barang.frx":33B4
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin XPControls.XPText txtbatch 
         Height          =   285
         Left            =   5280
         TabIndex        =   21
         Top             =   2520
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":33CA
         TabIndex        =   66
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox cmblok 
         Height          =   315
         Left            =   5280
         TabIndex        =   20
         Top             =   2160
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":3438
         TabIndex        =   65
         Top             =   2160
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   375
         Left            =   4080
         OleObjectBlob   =   "barang.frx":34AC
         TabIndex        =   64
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "barang.frx":3520
         TabIndex        =   60
         Top             =   3600
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "barang.frx":3580
         TabIndex        =   59
         Top             =   3960
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":35E0
         TabIndex        =   58
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":3654
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":36C2
         TabIndex        =   56
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cari"
         Height          =   255
         Left            =   7200
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel harga 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":3740
         TabIndex        =   53
         Top             =   3240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel label10 
         Height          =   255
         Left            =   7200
         OleObjectBlob   =   "barang.frx":379E
         TabIndex        =   52
         Top             =   2880
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":37FC
         TabIndex        =   48
         Top             =   600
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "barang.frx":3872
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":38E2
         TabIndex        =   46
         Top             =   3960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3958
         TabIndex        =   45
         Top             =   3600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":39CA
         TabIndex        =   44
         Top             =   3240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3A3C
         TabIndex        =   43
         Top             =   2880
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3AB0
         TabIndex        =   42
         Top             =   1560
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3B1A
         TabIndex        =   41
         Top             =   1200
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3B88
         TabIndex        =   40
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":3BF8
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox text3 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin XPControls.XPText ps2 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   3960
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
      Begin XPControls.XPText ps3 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   3600
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
      Begin XPControls.XPText text2 
         Height          =   525
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSComctlLib.ListView LV1 
         Height          =   3135
         Left            =   8520
         TabIndex        =   36
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Satuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Harga"
            Object.Width           =   2540
         EndProperty
      End
      Begin XPControls.XPText txtharga 
         Height          =   285
         Left            =   5280
         TabIndex        =   24
         Top             =   3240
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
      Begin XPControls.XPText jum 
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         Top             =   2880
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
      Begin XPControls.XPText txtstn 
         Height          =   285
         Left            =   5280
         TabIndex        =   22
         Top             =   2880
         Width           =   1095
         _ExtentX        =   1931
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
      Begin apotekbaleendah.ThemedButton Tcmdhapus 
         Height          =   375
         Left            =   11160
         TabIndex        =   37
         Top             =   3480
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Hapus"
         font            =   "barang.frx":3C68
         forecolor       =   0
         mouseicon       =   "barang.frx":3C94
      End
      Begin XPControls.XPText text9 
         Height          =   285
         Left            =   5280
         TabIndex        =   16
         Top             =   600
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
      Begin XPControls.XPText text7 
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   5280
         TabIndex        =   15
         Top             =   240
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
      Begin XPControls.XPText text6 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   1440
         TabIndex        =   8
         Top             =   3240
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
      Begin XPControls.XPText text4 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   2880
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
      Begin XPControls.XPText txtsupp 
         Height          =   285
         Left            =   5280
         TabIndex        =   55
         Top             =   1320
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
         Locked          =   -1  'True
      End
      Begin XPControls.XPText txtids 
         Height          =   285
         Left            =   5280
         TabIndex        =   17
         Top             =   960
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "barang.frx":422E
         TabIndex        =   62
         Top             =   4320
         Width           =   255
      End
      Begin XPControls.XPText psh3 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   4320
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
      Begin XPControls.XPText txtharga3 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":428E
         TabIndex        =   63
         Top             =   4320
         Width           =   1215
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
         Left            =   5280
         TabIndex        =   19
         Top             =   1680
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
         Format          =   119275523
         CurrentDate     =   37623
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":4302
         TabIndex        =   67
         Top             =   2040
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel Kepemilikan 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "barang.frx":4370
         TabIndex        =   68
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "="
         Height          =   255
         Left            =   6360
         TabIndex        =   38
         Top             =   2880
         Width           =   255
      End
   End
   Begin apotekbaleendah.ThemedButton ThemedButton1 
      Height          =   255
      Left            =   4920
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
      _extentx        =   1931
      _extenty        =   450
      caption         =   "Ubah"
      font            =   "barang.frx":43E4
      mouseicon       =   "barang.frx":4410
   End
   Begin VB.TextBox txtcari 
      Height          =   375
      Left            =   10680
      TabIndex        =   33
      Top             =   5520
      Width           =   2415
   End
   Begin apotekbaleendah.ThemedButton Cmdkeluar 
      Height          =   375
      Left            =   5280
      TabIndex        =   27
      Top             =   5400
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "&Keluar"
      font            =   "barang.frx":49AA
      mouseicon       =   "barang.frx":49D6
   End
   Begin apotekbaleendah.ThemedButton Cmdsimpan 
      Height          =   375
      Left            =   4440
      TabIndex        =   25
      Top             =   5400
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Simpan"
      font            =   "barang.frx":4F70
      mouseicon       =   "barang.frx":4F9C
   End
   Begin apotekbaleendah.ThemedButton Cmdbatal 
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   5400
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "Bata&l"
      font            =   "barang.frx":5536
      mouseicon       =   "barang.frx":5562
   End
   Begin apotekbaleendah.ThemedButton Cmdhapus 
      Height          =   375
      Left            =   2520
      TabIndex        =   29
      Top             =   5400
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "&Hapus"
      font            =   "barang.frx":5AFC
      mouseicon       =   "barang.frx":5B28
   End
   Begin apotekbaleendah.ThemedButton Cmdedit 
      Height          =   375
      Left            =   1560
      TabIndex        =   28
      Top             =   5400
      Width           =   975
      _extentx        =   1720
      _extenty        =   661
      caption         =   "&Ubah"
      font            =   "barang.frx":60C2
      mouseicon       =   "barang.frx":60EE
   End
   Begin apotekbaleendah.ThemedButton cmdtambah 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   5400
      Width           =   855
      _extentx        =   1508
      _extenty        =   661
      caption         =   "&Baru"
      font            =   "barang.frx":6688
      mouseicon       =   "barang.frx":66B4
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Kode barang otomatis"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5880
      Width           =   2055
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6840
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CheckBox XPCheck1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Visible         =   0   'False
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "barang.frx":6C4E
      TabIndex        =   30
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "barang.frx":6CC4
      Top             =   6120
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   2895
      Left            =   0
      TabIndex        =   54
      Top             =   6360
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
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
End
Attribute VB_Name = "Barang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim status As Byte
Dim flag, sumber As String
Dim ttext5, ttext6, ttext7 As Double
Dim frmResize As New ControlResizer
        Dim fso As New FileSystemObject


Private Sub Command2_Click()
Dim cls As New clsDlg
sumber = cls.OpenFlDlg(Me.hwnd, "File gambar(jpg,gif,bmp)|*.jpg;*.gif;*.bmp|JPEGS|*.jpg|GIFS|*.gif|Bitmaps|*.bmp", "Open File", vbNullString, False)



End Sub

Private Sub cmbgol_Click()
cmbtipe.SetFocus
End Sub

Private Sub cmblok_Click()
txtbatch.SetFocus
End Sub

Private Sub cmbtipe_Click()
text4.SetFocus
End Sub

Private Sub cmbtipe_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command4_Click()
If Text1.Text = "" Then Exit Sub
Set RS = New Recordset
RS.Open "select kode_brg from tblbarang where kode_brg='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub

If Text1.Text <> "" Then
frmstok.Show
End If

End Sub

Private Sub Form_Activate()
If hbrg = "" Then
hbrg = "Br"
End If
hrfc.Caption = hbrg

End Sub
Sub dbgrid1_Click()
On Error Resume Next
teks

Frame.Enabled = False
'Cmdedit.SetFocus
Cmdsimpan.Enabled = False

End Sub
Private Sub Form_Resize()
  Call frmResize.FormResized(Me)
    
End Sub


Private Sub Check1_Click()
    SaveSetting "apotekbaleendah", "Barang", "Check1.value", Check1.Value

End Sub
Sub ttl()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(5)) * val(LV1.ListItems(I).SubItems(4))
Next I
text4.Text = sum
End Sub

Private Sub Cmdbatal_Click()
awal
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub lvbrg_Click()
'On Error Resume Next
teks

Frame.Enabled = False
'Cmdedit.SetFocus
Cmdsimpan.Enabled = False

End Sub
Private Sub lvbrg_GotFocus()
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub

Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
teks

Frame.Enabled = True
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
Text1.Enabled = False

End Sub

Sub teks()
'On Error Resume Next
dbgridk
cmdtambah.Enabled = True
Text1.Text = dbgrid1.Columns(0).Text
text2.Text = dbgrid1.Columns(1).Text
Combo1.Text = dbgrid1.Columns(2).Text

text3.Text = dbgrid1.Columns(3).Text
cmbgol.Text = dbgrid1.Columns(4).Text
cmbtipe.Text = dbgrid1.Columns(5).Text
text4.Text = dbgrid1.Columns(6).Text
text5.Text = dbgrid1.Columns(7).Text
text6.Text = dbgrid1.Columns(8).Text
text7.Text = dbgrid1.Columns(9).Text
txtharga3.Text = dbgrid1.Columns(10).Text

text8.Text = dbgrid1.Columns(11).Text
text9.Text = dbgrid1.Columns(12).Text
txtids.Text = dbgrid1.Columns(13).Text
tgl.Value = dbgrid1.Columns(14).Text
cmblok.Text = dbgrid1.Columns(15).Text
txtbatch.Text = dbgrid1.Columns(16).Text

txtids_KeyPress (13)

End Sub
Sub teks2()
'On Error Resume Next
Set RS = New Recordset
RS.Open "select * from tblbarang where kode_brg='" & dbgrid1.Columns(0).Text & "'", jual, adOpenStatic, adLockOptimistic
With RS
Text1.Text = .Fields(0)
text2.Text = .Fields(1)
Combo1.Text = .Fields(2)

text3.Text = .Fields(3)
cmbgol.Text = .Fields(4)
cmbtipe.Text = .Fields(5)
text4.Text = .Fields(6)
text5.Text = .Fields(7)
text6.Text = .Fields(8)
text7.Text = .Fields(9)
txtharga3.Text = .Fields(10)

text8.Text = .Fields(11)
text9.Text = .Fields(12)
txtids.Text = .Fields(13)
tgl.Value = .Fields(14)
cmblok.Text = .Fields(1153)
txtbatch.Text = .Fields(16)

txtids_KeyPress (13)

End With
End Sub
Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
Set rsd = New Recordset
rsd.Open "Select kode_brg from detiljual where kode_brg='" & Text1.Text & "' limit 1", jual, adOpenStatic, adLockOptimistic
If Not rsd.EOF Then
MsgBox "Tidak dapat dihapus,sudah ada transaksi dengan kode barang ini", vbCritical, judul
Exit Sub
End If
sql = "delete from tblbarang where kode_brg='" & Text1.Text & "'"

jual.Execute (sql)
dbgrid
teks
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub


Private Sub Cmdsimpan_Click()
'On Error GoTo erol
text2.Text = Replace(text2.Text, "'", "''")
If Text1.Text = "" Then
MsgBox "Kode barang harus diisi", vbCritical, judul
Text1.SetFocus
Exit Sub
Else

If text2.Text = "" Then
MsgBox "Nama barang harus diisi", vbCritical, judul
text2.SetFocus
Exit Sub
Else
If text6.Text = "" Then
MsgBox "Harga barang harus diisi", vbCritical, judul
text6.SetFocus
Exit Sub
Else
If text7.Text = "" Then
MsgBox "Harga grosir barang harus diisi", vbCritical, judul
text7.SetFocus
Exit Sub
Else
If text5.Text = "" Then
MsgBox "Harga beli barang harus diisi", vbCritical, judul
text5.SetFocus
Exit Sub
Else
End If
End If
End If
End If
End If
If Edit = False Then
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where deskripsi='" & text2.Text & "' ", jual, adOpenStatic, adLockOptimistic
If Not rsbarang.EOF Then
MsgBox "Nama barang sudah terdaftar", vbCritical
text2.SetFocus
Exit Sub
End If
End If

If Edit = True Then
If val(text4.Text) <> val(dbgrid1.Columns(6).Text) Then
MsgBox "Perubahan stok harus di form penyesuaian stok! ", vbCritical, judul
Exit Sub
End If
Set rsbarang = New Recordset
rsbarang.Open "select * from tblbarang where deskripsi='" & text2.Text & "' and kode_brg<> '" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rsbarang.EOF Then
MsgBox "Nama barang sudah terdaftar", vbCritical
text2.SetFocus
Exit Sub
End If
End If


Set cari = LV1.FindItem(text3.Text, 1, , 1)
LV1.SelectedItem = cari

If Not cari Is Nothing Then
MsgBox "Satuan sudah terdaftar"
text3.SetFocus
Exit Sub
End If

Set rsbarang = New Recordset
A = Text1.Text
rsbarang.Open "select * from tblbarang where kode_brg='" & A & "' ", jual, adOpenStatic, adLockOptimistic
If Edit = False Then
ubah

jual.Execute "insert into tblbarang values('" & Text1.Text & "','" & text2.Text & "','" & Combo1.Text & "','" & text3.Text & "','" & cmbgol.Text & "','" & cmbtipe.Text & "'," & val(text4.Text) & "," & ttext5 & "," & ttext6 & "," & ttext7 & ",'" & Format(txtharga3.Text, Number) & "','" & text8.Text & "'," & val(text9.Text) & ",'" & txtids.Text & "','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & cmblok.Text & "','" & txtbatch.Text & "')"

text4.Locked = False
text5.Locked = False
If isibeli = True Then
pembelian.Show
pembelian.cmdtambah_Click
pembelian.text4.Text = Text1.Text
pembelian.text4_Click
Exit Sub
End If

If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
dbgrid

cmdtambah_Click
Exit Sub
Else
awal
End If
Else
If text3.Text <> rsbarang.Fields(3) Then
jual.Execute "delete from tblbarang where kode_brg='" & Text1.Text & "'"
jual.Execute "delete from stokgudang where kode_brg='" & Text1.Text & "'"

ubah
jual.Execute "insert into tblbarang values('" & Text1.Text & "','" & text2.Text & "','" & Combo1.Text & "','" & text3.Text & "','" & cmbgol.Text & "','" & cmbtipe.Text & "'," & val(text4.Text) & "," & ttext5 & "," & ttext6 & "," & ttext7 & ",'" & Format(txtharga3.Text, Number) & "','" & text8.Text & "'," & val(text9.Text) & ",'" & txtids.Text & "','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & cmblok.Text & "','" & txtbatch.Text & "')"

MsgBox "Data berhasil diubah", vbInformation
txtcari_Change
Exit Sub
Else

ubah
jual.Execute "update tblbarang set deskripsi='" & text2.Text & "',kategori='" & Combo1.Text & "',satuan='" & text3.Text & "',golongan='" & cmbgol.Text & "',tipe='" & cmbtipe.Text & "',stok=" & val(text4.Text) & ",harga_beli=" & ttext5 & ",harga_jual=" & ttext6 & ",harga_jual2=" & ttext7 & ",harga_jual3='" & Format(txtharga3.Text, Number) & "',diskon=" & val(text8.Text) & ",stok_minimal=" & val(text9.Text) & ",id_supplier='" & txtids.Text & "',tgl_expire='" & Format(tgl.Value, "yyyy-mm-dd") & "',lokasi='" & cmblok.Text & "',batch='" & txtbatch.Text & "'  where kode_brg='" & Text1.Text & "'"


MsgBox "Data berhasil diubah", vbInformation
txtcari_Change
Exit Sub

End If
End If
dbgrid
Exit Sub
erol:
If err.Number = -2147217900 Then
MsgBox "Kode barang sudah terdaftar!", vbCritical, judul
Else
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
Frame.Enabled = True
End If
End If
End Sub

Sub ubah()
Dim dis As String
If text5.Text = "" Or text5.Text = "0" Then
ttext5 = 0
Else
ttext5 = Format(text5.Text, Number)

End If
If text6.Text = "" Or text6.Text = "0" Then
ttext6 = 0
Else
ttext6 = Format(text6.Text, Number)

End If
If text7.Text = "" Or text7.Text = "0" Then
ttext7 = 0
Else
ttext7 = Format(text7.Text, Number)

End If


jual.Execute "Delete from satuan where kode_brg='" & Text1.Text & "' and keterangan<>'utama'"
If LV1.ListItems.count <> 0 Then

For z = 1 To LV1.ListItems.count

sql = "insert into satuan values('" & Text1.Text & "','" & LV1.ListItems(z).SubItems(1) & "','" & LV1.ListItems(z).SubItems(2) & "','','" & LV1.ListItems(z).SubItems(3) & "')"
jual.Execute (sql)
    Next z
End If
If sumber = "" Then Exit Sub
FileName = App.Path & "\gambar\" & Text1.Text & "" + ".jpg"
     Set fso = New FileSystemObject

fso.CopyFile sumber, FileName
Set fso = Nothing

Frame.Enabled = False


End Sub
Sub cmdtambah_Click()
isibeli = False

Edit = False
tambah
kosong
LV1.ListItems.Clear

Text1.SetFocus
If Check1.Value = Checked Then
kode_oto
text2.SetFocus

End If
ktgr
stn
datagolongan
daba = 0
End Sub
Sub kode_oto()
Dim j As Integer
Dim No As String
Set rsbarang = New Recordset
A = hbrg & "%"
sql = "Select kode_brg from tblbarang where kode_brg like '" & A & "' order by kode_brg "
Set rsbarang = jual.Execute(sql)
If rsbarang.EOF = True Then
Text1.Text = hbrg & "0001"
Else
rsbarang.MoveLast
j = val(Right(rsbarang(0), 4))
No = hbrg + Format(Str(j + 1), "0000")
Text1.Text = No

End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub ktgr()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
sql = "select * from tblbarang group by kategori order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo1.AddItem rsbarang!kategori
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
cmdtambah_Click
Else
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\ISI DATA BARANG.doc" _
                 , vbNullString, vbNullString, 1
End If

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If status_isi = "pemesanan aktif" Then
pemesanan.Show
End If

End Sub

Private Sub jum_Change()
If val(text7.Text) = 0 Then
txtharga.Text = val(jum.Text) * val(Format(text6.Text, Number))
Else
txtharga.Text = val(jum.Text) * val(Format(text7.Text, Number))

End If
End Sub

Private Sub jum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 13 Then
If txtstn.Text = "" Or jum.Text = "" Then Exit Sub

If txtstn.Text = text3.Text Then
MsgBox "Jangan sama dengan satuan utama"
jum.Text = ""
txtstn.Text = ""
txtstn.SetFocus
End If

isigrid
jum.Text = ""
txtstn.Text = ""
txtstn.SetFocus
harga.Caption = ""
txtharga.Text = ""
End If
End Sub
Sub isigrid()
Dim ttxtharga As Double
ttxtharga = val(Format(txtharga.Text, Number))
On Error Resume Next
   '     Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        'l.SubItems(1) = ![komposisi]
Set cari = LV1.FindItem(txtstn.Text, 1, , 1)
LV1.SelectedItem = cari

If Not cari Is Nothing Then
MsgBox "Satuan sudah terdaftar"
text9.SetFocus
Exit Sub
End If

    Set butir = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
    With butir

    .SubItems(1) = txtstn.Text
    .SubItems(2) = jum.Text
       .SubItems(3) = ttxtharga

 End With
  LV1.Enabled = True
End Sub


Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
Frame.Enabled = True
Cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
dbgrid1.Enabled = False
Text1.Enabled = True
cmbtipe.ListIndex = 0
End Sub
Sub kosong()
Text1.Text = ""
text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""
text7.Text = ""
txtharga3.Text = ""
cmblok.Text = ""
text8.Text = "0"
text9.Text = ""
label10.Caption = ""
harga.Caption = ""
txtids.Text = ""
txtbatch.Text = ""
txtsupp.Text = ""
cmbgol.Text = ""
cmbtipe.Text = "sendiri"
End Sub

Private Sub Command1_Click()
frmsupp.Show

End Sub


Private Sub Command3_Click()
On Error GoTo erol
If Edit = False Then Exit Sub
Kill App.Path & "\gambar\" & Text1.Text & ".jpg"
sumber = ""
MsgBox "Foto telah berhasil dihapus"
'End If
erol:
If err.Number = 53 Then Exit Sub


End Sub


Private Sub dbgrid1_GotFocus()
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Refresh
XPCheck1.Value = GetSetting("apotekbaleendah", "Barang", "XPCheck1.value", Checked)
Check1.Value = GetSetting("apotekbaleendah", "Barang", "Check1.value", Checked)
If hbrg = "" Then
hbrg = "Br"
End If
tgl.Value = Now
hrfc.Caption = hbrg
awal
Ketengah Me
dbgrid
  frmResize.KeepRatio = True
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
    Me.WindowState = 2
  Form_Resize

Dim Arq As String
    Skinpath = App.Path & "\skin\galaxy.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
ktgr
stn
lislok

End Sub
Sub dbgridk()
On Error Resume Next
Dim l As ListItem
LV1.ListItems.Clear
LV1.ColumnHeaders.Item(1).Text = "No"
LV1.ColumnHeaders.Item(2).Text = "Satuan"

Set RS = New Recordset
RS.Open "Select * from satuan where kode_brg='" & dbgrid1.Columns(0).Text & "' and keterangan=''", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
LV1.ColumnHeaders.Item(3).Text = "Konversi"
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![satuan]
              l.SubItems(2) = ![konversi]
  l.SubItems(3) = ![harga]
    .MoveNext
    Loop
End With
RS.Close
End Sub

Private Sub stn()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

text3.Clear
sql = "select distinct satuan from tblbarang where satuan is not null and satuan!='' order by kode_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text3.AddItem rsbarang!satuan
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

  End Sub
Private Sub datagolongan()
'On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbgol.Clear
sql = "select distinct golongan from tblbarang where golongan is not null and golongan!='' order by golongan"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbgol.AddItem rsbarang!golongan
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
End Sub
Private Sub awal()
Frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
dbgrid1.Enabled = True

End Sub
Private Sub lislok()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmblok.Clear
sql = "select distinct lokasi from tblbarang where lokasi is not null and lokasi!='' order by kode_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmblok.AddItem rsbarang!Lokasi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

  End Sub


Sub dbgrid()
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang
dbgrid1.Columns(0).Width = 1500
dbgrid1.Columns(1).Width = 4000
dbgrid1.Columns(2).Width = 1400
dbgrid1.Columns(3).Width = 1400
dbgrid1.Columns(4).Width = 1400
dbgrid1.Columns(5).Width = 1500
dbgrid1.Columns(6).Width = 1400
dbgrid1.Columns(7).Width = 1400
dbgrid1.Columns(8).Width = 1400
dbgrid1.Columns(9).Width = 1400
dbgrid1.Columns(10).Width = 1400

dbgrid1.Columns(1).Caption = "Nama obat"
dbgrid1.Columns(8).Caption = "Harga Umum"
dbgrid1.Columns(9).Caption = "Harga Dokter"
dbgrid1.Columns(10).Caption = "Harga Resep"


End Sub
Sub Dbgrid2()
Dim stri As String

stri = Replace(txtcari.Text, "'", "''")

sql = "select * from tblbarang where deskripsi like '%" & stri & "%' or kode_brg like '" & stri & "%' order by deskripsi"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang
dbgrid1.Columns(0).Width = 1500
dbgrid1.Columns(1).Width = 4000
dbgrid1.Columns(2).Width = 1400
dbgrid1.Columns(3).Width = 1400
dbgrid1.Columns(4).Width = 1400
dbgrid1.Columns(5).Width = 1500
dbgrid1.Columns(6).Width = 1400
dbgrid1.Columns(7).Width = 1400
dbgrid1.Columns(8).Width = 1400
dbgrid1.Columns(9).Width = 1400
dbgrid1.Columns(10).Width = 1400
dbgrid1.Columns(1).Caption = "Nama obat"
dbgrid1.Columns(6).Caption = "Harga Umum"
dbgrid1.Columns(7).Caption = "Harga Dokter"
dbgrid1.Columns(8).Caption = "Harga Resep"


End Sub



Private Sub m_Click()

End Sub

Private Sub ps2_Change()
If text5.Text = "" Then Exit Sub

text7.Text = (val(ps2.Text) + 100) * 0.01 * Format(text5.Text, Number)

End Sub

Private Sub ps2_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = validasiAngka2(KeyAscii)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub ps3_Change()
If text5.Text = "" Then Exit Sub
text6.Text = (val(ps3.Text) + 100) * 0.01 * Format(text5.Text, Number)
End Sub

Private Sub ps3_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = validasiAngka2(KeyAscii)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub psh3_Change()
If text5.Text = "" Then Exit Sub

txtharga3.Text = (val(psh3.Text) + 100) * 0.01 * Format(text5.Text, Number)

End Sub

Private Sub TCmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
For I = 1 To LV1.ListItems.count
    LV1.ListItems(I).Text = I
Next I

End Sub

Private Sub KODE_Change()
kode = A

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If
End Sub

Private Sub text3_Click()
On Error Resume Next
            SendKeys "{tab}"     ' Set the focus to the next control.

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmbgol.SetFocus
End If
End Sub

Private Sub text4_GotFocus()
If isibeli = True Then
text4.Text = ""
text4.Locked = True
Else
text4.Locked = False

End If
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text5_Change()
On Error Resume Next
text5.Text = Format(text5.Text, "#,#"): SendKeys "{end}"
If text5.Text = "" Then Exit Sub
text6.Text = (val(ps3.Text) + 100) * 0.01 * Format(text5.Text, Number)
text7.Text = (val(ps2.Text) + 100) * 0.01 * Format(text5.Text, Number)
txtharga3.Text = (val(psh3.Text) + 100) * 0.01 * Format(text5.Text, Number)

End Sub
Private Sub Text6_Change()
On Error Resume Next
text5.Text = Format(text5.Text, "#,#"): SendKeys "{end}"

text6.Text = Format(text6.Text, "#,#"): SendKeys "{end}"

End Sub

Private Sub tgl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtstn.SetFocus
End If
End Sub

Private Sub txtharga3_Change()
On Error Resume Next

txtharga3.Text = Format(txtharga3.Text, "#,#"): SendKeys "{end}"

End Sub

Private Sub Text7_Change()
On Error Resume Next

text7.Text = Format(text7.Text, "#,#"): SendKeys "{end}"

End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
ps3.SetFocus
End If

End Sub
Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub
Private Sub hjg_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
text7.SetFocus
End If

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If
End Sub

Private Sub text9_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If
End Sub

Private Sub ThemedButton1_Click()
huruff.Show
End Sub

Private Sub txtcari_Change()
Dbgrid2
End Sub


Private Sub txtharga_Change()
On Error Resume Next

txtharga.Text = Format(txtharga.Text, "#,#"): SendKeys "{end}"

End Sub

Private Sub txtharga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
jum_KeyPress (13)
End If
End Sub

Sub txtids_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtids.Text = "" Then Exit Sub
txtsupp.Text = ""
If txtids.Text = "" Then
txtstn.SetFocus
Exit Sub
End If
Set RS = New Recordset
RS.Open "select * from tblsupplier where id_supplier='" & txtids.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Supplier tidak terdaftar", vbInformation, judul
txtids.Text = ""
txtsupp.Text = ""
txtids.SetFocus
Else
            txtsupp.Text = RS!Supplier
            'txtstn.SetFocus
End If
End If

End Sub

Private Sub txtstn_Change()
If txtstn.Text = "" Then
harga.Caption = ""
Else
harga.Caption = "Harga/" & txtstn.Text
End If
End Sub

Private Sub txtstn_GotFocus()
label10.Caption = text3.Text
End Sub

Private Sub txtstn_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key
 If Not txtstn.Text = "" Then
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
            Else
            Cmdsimpan.SetFocus
         End If
End If
End Sub

Private Sub XPCheck1_Click()
    
    SaveSetting "apotekbaleendah", "Barang", "XPCheck1.value", XPCheck1.Value

End Sub
