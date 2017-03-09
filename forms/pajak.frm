VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form pajak 
   Caption         =   "Pajak"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin penjualan.ThemedButton Cmdhapus 
      Height          =   375
      Left            =   8880
      TabIndex        =   56
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "pajak.frx":0000
   End
   Begin penjualan.ThemedButton cmdsimpan 
      Height          =   255
      Left            =   5520
      TabIndex        =   55
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "ThemedButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "pajak.frx":059A
   End
   Begin penjualan.ThemedButton command1 
      Height          =   375
      Left            =   4320
      TabIndex        =   54
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
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
      MouseIcon       =   "pajak.frx":0B34
   End
   Begin penjualan.ThemedButton Cmdkeluar 
      Height          =   375
      Left            =   2880
      TabIndex        =   53
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
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
      MouseIcon       =   "pajak.frx":10CE
   End
   Begin penjualan.ThemedButton cmdtambah 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "B&aru"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "pajak.frx":1668
   End
   Begin penjualan.ThemedButton cmdbatal 
      Height          =   375
      Left            =   1560
      TabIndex        =   52
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "pajak.frx":1C02
   End
   Begin ACTIVESKINLibCtl.SkinLabel ket 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "pajak.frx":219C
      TabIndex        =   51
      Top             =   9000
      Width           =   4095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "pajak.frx":21FA
      TabIndex        =   50
      Top             =   9240
      Width           =   4575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2280
      OleObjectBlob   =   "pajak.frx":22B6
      Top             =   8400
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10080
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin VB.Frame Frame 
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   9975
      Begin XPControls.XPCombo nama 
         Height          =   315
         Left            =   6720
         TabIndex        =   25
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
      Begin XPControls.XPText almt 
         Height          =   285
         Left            =   6720
         TabIndex        =   26
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
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
      Begin XPControls.XPText telp 
         Height          =   285
         Left            =   6720
         TabIndex        =   27
         Top             =   1200
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "pajak.frx":24EA
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Index           =   0
         Left            =   5280
         OleObjectBlob   =   "pajak.frx":2554
         TabIndex        =   47
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   5280
         OleObjectBlob   =   "pajak.frx":25CA
         TabIndex        =   48
         Top             =   1200
         Width           =   1335
      End
      Begin XPControls.XPCombo nama2 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
      Begin XPControls.XPText almt2 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
      Begin XPControls.XPText telp2 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "pajak.frx":2630
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "pajak.frx":269A
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "pajak.frx":2714
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   6600
      Width           =   4095
      Begin VB.ComboBox text4 
         Height          =   315
         Left            =   1800
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin XPControls.XPText stn 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
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
      Begin XPControls.XPText harga 
         Height          =   285
         Left            =   1800
         TabIndex        =   14
         Top             =   600
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
      Begin XPControls.XPText jum 
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
      Begin XPControls.XPText subttl 
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   2040
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
      Begin XPControls.XPText ttlh 
         Height          =   285
         Left            =   1800
         TabIndex        =   37
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "pajak.frx":277A
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "pajak.frx":27E4
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "pajak.frx":284E
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "pajak.frx":28BE
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "pajak.frx":2928
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   4440
      TabIndex        =   1
      Top             =   6720
      Width           =   4575
      Begin XPControls.XPText gttl 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   720
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
      Begin XPControls.XPText dp 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
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
      Begin XPControls.XPText disk 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   360
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
      Begin XPControls.XPText disp 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   360
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
      Begin XPControls.XPText sisa 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
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
      Begin XPControls.XPText dasar 
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   1800
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":299C
         TabIndex        =   30
         Top             =   1800
         Width           =   1815
      End
      Begin XPControls.XPText ppn 
         Height          =   285
         Left            =   2040
         TabIndex        =   31
         Top             =   2160
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":2A24
         TabIndex        =   33
         Top             =   2160
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":2A88
         TabIndex        =   34
         Top             =   720
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":2B0E
         TabIndex        =   35
         Top             =   1080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":2B86
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "pajak.frx":2BFE
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2415
      Left            =   120
      TabIndex        =   29
      Top             =   3720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4260
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keterangan"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Harga "
         Object.Width           =   2646
      EndProperty
   End
   Begin XPControls.XPText total 
      Height          =   975
      Left            =   4560
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
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
   Begin ACTIVESKINLibCtl.SkinLabel label10 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "pajak.frx":2C64
      TabIndex        =   42
      Top             =   360
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel jam 
      Height          =   375
      Left            =   2760
      OleObjectBlob   =   "pajak.frx":2CDE
      TabIndex        =   43
      Top             =   360
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel kasir 
      Height          =   375
      Left            =   1080
      OleObjectBlob   =   "pajak.frx":2D58
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "pajak.frx":2DD2
      TabIndex        =   45
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin XPControls.XPText notr 
      Height          =   285
      Left            =   7680
      TabIndex        =   4
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
   Begin MSComCtl2.DTPicker tgl 
      Height          =   375
      Left            =   7680
      TabIndex        =   49
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   59572227
      CurrentDate     =   40299
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "pajak.frx":2E3C
      TabIndex        =   40
      Top             =   960
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel Nomor 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "pajak.frx":2EAA
      TabIndex        =   41
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "pajak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim kode, nbrg, pos, KETE As String
Dim st, cosbli, cosju, sel As Currency
Dim afk As Integer


Private Sub almt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
telp.SetFocus
End If

End Sub
Private Sub kbrg()
On Error Resume Next
  Dim i As Long
  Dim j As Long

text4.Clear
sql = "select * from detiljual order by nama_barang"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text4.AddItem rsbarang!nama_barang
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    With text4
    For i = 0 To .ListCount - 1
      For j = .ListCount To (i + 1) Step -1
         If .List(j) = .List(i) Then
           .RemoveItem j
         End If
      Next j
    Next i
  End With


  End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
kurir.Enabled = True
kurir.SetFocus
Else
kurir.Enabled = False
kurir.Text = ""
End If

End Sub

Private Sub almt2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
telp2.SetFocus
End If

End Sub

Private Sub cmdbatal_Click()
awal
kosong
cmdtambah.SetFocus
End Sub
Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
Frame.Enabled = True
Frame2.Enabled = True

Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True

End Sub




Private Sub Command1_Click()
On Error Resume Next

With CrystalReport1
.Reset
  '.Password = Chr(10) & "tujuh"

  .ReportFileName = App.Path & "\pajak.rpt"
  '.RetrieveDataFiles

  .WindowTitle = "invoice"

      .Formulas(0) = "tgl= '" & Format(tgl.Value, "dd-MMM-YYYY") & "'"
.Formulas(1) = "kode= '" & notr.Text & "'"
.Formulas(2) = "nama= '" & nama.Text & "'"
.Formulas(3) = "almt= '" & almt.Text & "'"
.Formulas(4) = "npwp= '" & telp.Text & "'"
.Formulas(5) = "namau= '" & nama2.Text & "'"
.Formulas(6) = "almtu= '" & almt2.Text & "'"
.Formulas(7) = "npwpu= '" & telp2.Text & "'"

.Formulas(8) = "nama1= '" & LV1.ListItems(1).SubItems(2) & "'"
.Formulas(9) = "nama2= '" & LV1.ListItems(2).SubItems(2) & "'"
.Formulas(10) = "nama3= '" & LV1.ListItems(3).SubItems(2) & "'"
.Formulas(11) = "nama4= '" & LV1.ListItems(4).SubItems(2) & "'"
.Formulas(12) = "harga1= '" & Format(LV1.ListItems(1).SubItems(3), "#,#0.#0") & "'"
.Formulas(13) = "harga2= '" & Format(LV1.ListItems(2).SubItems(3), "#,#0.#0") & "'"
.Formulas(14) = "harga3= '" & Format(LV1.ListItems(3).SubItems(3), "#,#0.#0") & "'"
.Formulas(15) = "harga4= '" & Format(LV1.ListItems(4).SubItems(3), "#,#0.#0") & "'"
.Formulas(16) = "no1= '" & LV1.ListItems(1).SubItems(1) & "'"
.Formulas(17) = "no2= '" & LV1.ListItems(2).SubItems(1) & "'"
.Formulas(18) = "no3= '" & LV1.ListItems(3).SubItems(1) & "'"
.Formulas(19) = "no4= '" & LV1.ListItems(4).SubItems(1) & "'"
.Formulas(20) = "ttl= '" & Format(val(ttlh.Text), "#,#0.#0") & "'"

.Formulas(23) = "dasar= '" & Format(val(dasar.Text), "#,#0.#0") & "'"
.Formulas(24) = "ppn= '" & Format(val(ppn.Text), "#,#0.#0") & "'"

.Formulas(25) = "ter= '" & terbilang.Caption & "'"


        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hWnd

        .WindowState = crptMaximized
                .Destination = crptToPrinter

  .Action = 1
End With
cmdtambah.Enabled = True
Set RS = New Recordset
sql = "select id_pengusaha from pengusaha where nama_pengusaha='" & nama2.Text & "' and alamat='" & almt2.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
idu = RS!id_pengusaha
RS.Close
Else
idu = ""
End If
Set RS = New Recordset
sql = "select id_pembeli from pembeli where nama_pembeli='" & nama.Text & "' and alamat='" & almt.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
idb = RS!id_pembeli
RS.Close
Else
idb = ""
End If


If idu = "" Then
Dim j As Integer
Dim no As String
Set rsplg = New Recordset
sql = "Select id_pengusaha from pengusaha order by id_pengusaha Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = "psu0001"
Else
j = val(Right(rsplg(0), 4))
idc = "psu" + Format(Str(j + 1), "0000")

End If
jual.Execute "insert into pengusaha(id_pengusaha,nama_pengusaha,alamat,npwp) values('" & idc & "','" & nama2.Text & "','" & almt2.Text & "','" & telp2.Text & "')"

End If


If idb = "" Then
Dim n As Integer
Dim Noo As String
Set rsplg = New Recordset
sql = "Select id_pembeli from pembeli order by id_pembeli Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = "pbl0001"
Else
no = val(Right(rsplg(0), 4))
idc = "pbl" + Format(Str(j + 1), "0000")

End If
jual.Execute "insert into pembeli(id_pembeli,nama_pembeli,alamat,npwp) values('" & idc & "','" & nama.Text & "','" & almt.Text & "','" & telp.Text & "')"

End If
cmdtambah.Enabled = True



End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub




Sub simpandata()
'On Error GoTo erol
Set RS = New Recordset
sql = "select id_konsumen from konsumen where nama_konsumen='" & nama.Text & "' and alamat='" & almt.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
id = RS!id_konsumen
RS.Close
Else
id = ""
End If
If val(dp.Text) >= val(gttl.Text) Then
byr = val(gttl.Text)
Else
byr = val(dp.Text)
End If

sql = "insert into penjualan values('" & notr.Text & "','" & Format(tgl, "dd MMM yyyy") & "','" & nama.Text & "','" & ttlh.Text & "','" & val(disk.Text) & "','" & val(gttl.Text) & "','" & kasir.Caption & "')"
jual.Execute (sql)

Set RS = New Recordset
For z = 1 To LV1.ListItems.count

sql = "insert into detiljual values('" & notr.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(3) & "','" & LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(6) & "')"
jual.Execute (sql)
    Next z
    
If id = "" Then
Dim j As Integer
Dim no As String
Set rsplg = New Recordset
sql = "Select id_konsumen from konsumen order by id_konsumen Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = "cus0001"
Else
j = val(Right(rsplg(0), 4))
idc = "cus" + Format(Str(j + 1), "0000")

End If
jual.Execute "insert into konsumen(id_konsumen,nama_konsumen,alamat,no_telp) values('" & idc & "','" & nama.Text & "','" & almt.Text & "','" & telp.Text & "')"

End If
Set RS = New Recordset
If afk > 0 Then
jual.Execute "insert into ayam(tanggal,ayam_afkir) values('" & tgl.Value & "','" & afk & "')"
End If
    MsgBox "Data sudah tersimpan", vbInformation, judul
      awal
 cmdtambah.SetFocus

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, judul
    Exit Sub
Frame.Enabled = True
Exit Sub

End If


End Sub

Private Sub Cmdsimpan_Click()
If MsgBox("Simpan data?", vbYesNo, judul) = vbYes And nama.Text <> "" Then
simpandata
Command1.Enabled = True

If MsgBox("Cetak struk?", vbYesNo, judul) = vbYes Then
Command1_Click
Else

Command1.SetFocus
End If

Else
MsgBox "Data belum lengkap", , judul
End If

End Sub

Sub cmdtambah_Click()
Edit = False
tambah
kosong
kosong2
'supp

notr.SetFocus
'supp
'suppp
End Sub

Sub tambah()
Edit = False
cmdtambah.Enabled = False
Frame.Enabled = True
Cmdbatal.Enabled = True
Frame2.Enabled = True
End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
sql = "Select no_penjualan from penjualan order by no_penjualan Desc"
Set RS = jual.Execute(sql)
If RS.EOF = True Then
notr.Text = "000001"

Else
j = val(Right(RS(0), 4))
br = Format(Str(j + 1), "000000")
notr.Text = br
no = br

nmr = Format(Mid(RS(0), 8, 2))
thn = Format(Mid(RS(0), 5, 2))
txt = Format(Mid(no, 8, 2))
thnn = Format(Mid(no, 5, 2))

'If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
'notr.Text = "PJ-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
'End If
End If
End Sub
Sub kosong()
notr.Text = ""
nama.Text = ""
almt.Text = ""
telp.Text = ""
nama2.Text = ""
almt2.Text = ""
telp2.Text = ""

harga.Text = ""
jum.Text = ""
subttl.Text = ""
ttlh.Text = ""
disp.Text = ""
disk.Text = ""
dp.Text = ""
gttl.Text = ""
sisa.Text = ""
LV1.ListItems.Clear
dasar.Text = ""
ppn.Text = ""
End Sub




Private Sub DataGrid1_DblClick()
On Error GoTo erol
 If pos = "1" Then
 nama.Text = DataGrid1.Columns(1)
 nama_Click
 text4.SetFocus
 Else
 If pos = "2" Then
 text4.Text = DataGrid1.Columns(1)
 text4_Click
 Else
 If pos = "3" Then
 remark.Text = DataGrid1.Columns(1)
 remark_KeyPress (13)
  Else
  If pos = "4" Then
kurir.Text = DataGrid1.Columns(1)
 kurir_KeyPress (13)
End If

 End If
 End If
 End If
erol:
 If err.Description <> vbNullString Then Exit Sub



End Sub






Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Cmdsimpan.SetFocus
End If

End Sub

Private Sub DTPicker3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan.SetFocus
End If
End Sub

Private Sub dasar_Change()
ppn.Text = 0.1 * val(dasar.Text)
End Sub

Private Sub dasar_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
If MsgBox("Cetak?", vbYesNo, judul) = vbYes Then
Command1_Click
Else

Command1.SetFocus
End If
End If
End Sub

Private Sub disk_Change()
gttl.Text = val(ttlh.Text) - val(disk.Text)

End Sub

Private Sub disk_GotFocus()
disp.Text = ""
End Sub

Private Sub disk_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
dp.SetFocus
End If

End Sub

Private Sub disp_Change()
If val(disp.Text) > 100 Then
MsgBox "ERORR"
disp.Text = "0"
Exit Sub
End If
disk.Text = val(disp.Text) / 100 * val(ttlh.Text)
gttl.Text = val(ttlh.Text) - val(disk.Text)
End Sub

Private Sub disp_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
gttl.Text = val(ttlh.Text) - val(disk.Text)
dp.SetFocus
End If
End Sub

Private Sub dp_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
If val(dp.Text) > val(gttl.Text) Then
kmbali = val(dp.Text) - val(gttl.Text)
sisa.Text = "0"
Else
sisa.Text = val(gttl.Text) - val(dp.Text)
End If



dasar.SetFocus
End If

End Sub





Sub kosong2()
text4.Text = ""
jum.Text = ""
subttl.Text = ""
harga.Text = ""
stn.Text = ""

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


Private Sub Form_Load()
    Set jual = New adodb.Connection
        jual.CursorLocation = adUseClient
jual.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/penjualan.mdb;Jet OLEDB:Database Password=tujuh;"
Label10.Caption = Format(Now, "dd-mm-YYYY")

tgl.Value = Now
Edit = True
'supp
'suppp

'kbrg
awal
kasir.Caption = kasirr
Ketengah Me
 DTPicker1 = Format(Now)
  DTPicker2 = Format(Now)
  DTPicker3 = Format(Now + 14)

     SkinPath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin SkinPath
    Skin1.ApplySkin Me.hWnd


End Sub
Private Sub supp()

  Dim i As Long
  Dim j As Long

nama.Clear
sql = "select * from pembeli order by nama_pembeli"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
nama.AddItem rsbarang!nama_pembeli
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub
Private Sub suppp()

  Dim i As Long
  Dim j As Long

nama2.Clear
sql = "select * from pengusaha order by nama_pengusaha"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
nama2.AddItem rsbarang!nama_pengusaha
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub

Private Sub dbgrid2_DblClick()
On Error Resume Next

If Edit = False Then
text3.Text = dbgrid2.Columns(1)
Tab1.Tab = 0
text4.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus

End If
End Sub

Sub dbgrid()
Set rsbarang = New Recordset
sql = "select * from tblbarang"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang


End Sub
Sub dbgrids()
Set rssupp = New Recordset

sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set dbgrid2.DataSource = rssupp


End Sub


Private Sub awal()
Frame.Enabled = False
Frame2.Enabled = False
Edit = True
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
End Sub

Private Sub hb_Click()
text5.Text = hb.Caption
text6.SetFocus
End Sub

Sub isigrid()
    Set butir = LV1.ListItems.Add
    With butir
    .SubItems(1) = LV1.ListItems.count & "."

    .SubItems(2) = text4.Text
    .SubItems(3) = harga.Text

    End With
  LV1.Enabled = True
  kosong2
End Sub







Private Sub harga_Change()
jum.Text = ""
End Sub

Private Sub harga_KeyPress(KeyAscii As Integer)
On Error Resume Next

KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

If KeyAscii = 13 Then

isigrid
kosong2
ttl
ttl2
text4.SetFocus
kosong2
text4.SetFocus

End If

End Sub

Private Sub jum_Change()
subttl.Text = val(jum.Text) * val(harga.Text)
End Sub


Private Sub kurir_Change()
Set rssupp = New Recordset

sql = "select  * from kurir where namakurir like'%" & kurir.Text & "%'"
Set rspo = jual.Execute(sql)
Set DataGrid1.DataSource = rspo

End Sub

Private Sub kurir_GotFocus()
Set rssupp = New Recordset
pos = "4"
sql = "select  * from kurir"
Set rspo = jual.Execute(sql)
Set DataGrid1.DataSource = rspo


End Sub

Private Sub kurir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
nama.SetFocus
End If

End Sub

Private Sub layanan_Click()
nama.SetFocus
End Sub

Private Sub layanan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
nama.SetFocus
End If
End Sub


Private Sub nama_Click()
nama_KeyPress (13)
End Sub
Private Sub nama2_Click()
nama2_KeyPress (13)
End Sub



Private Sub nama_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
sql = "select  * from pembeli where nama_pembeli ='" & nama.Text & "'"
Set RS = jual.Execute(sql)
If RS.EOF Then
almt.Text = ""
telp.Text = ""

almt.SetFocus
Else
almt.Text = RS!alamat
telp.Text = RS!npwp
text4.SetFocus
RS.Close
End If
End If
End Sub

Private Sub nama2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
sql = "select  * from pengusaha where nama_pengusaha ='" & nama2.Text & "'"
Set RS = jual.Execute(sql)
If RS.EOF Then
almt2.Text = ""
telp2.Text = ""

almt2.SetFocus
Else
almt2.Text = RS!alamat
telp2.Text = RS!npwp
nama.SetFocus
RS.Close
End If
End If
End Sub


Private Sub cmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
For X = 1 To LV1.ListItems.count
LV1.ListItems(X).SubItems(1) = X
Next X

End Sub

Private Sub remark_Change()

End Sub

Private Sub remark_KeyPress(KeyAscii As Integer)

End Sub

Private Sub notr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub stn_Change()
stn2.Caption = stn.Text
End Sub

Private Sub stn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If


End Sub

Private Sub telp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4.SetFocus
End If
End Sub


Private Sub telp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
nama.SetFocus
End If

End Sub

Private Sub text4_Change()
If text4.Text <> "" Then
ket.Caption = ""
Else
ket.Caption = "tekan enter bila barang sudah lengkap"

End If

End Sub
Private Sub text4_LostFocus()
ket.Caption = ""
End Sub

Private Sub text4_Click()
Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari
harga.SetFocus

If cari Is Nothing Then
harga.SetFocus
Else
harga.Text = LV1.SelectedItem.SubItems(3)
harga.SetFocus
End If
End Sub





Sub ttl()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(3))
Next i
ttlh.Text = sum
total.Text = Format(sum, "#,##0")
SkinLabel6.Visible = True

End Sub
Sub ttl2()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(i).SubItems(7)
Next i
afk = sum

End Sub
Sub diskon()
sum = 0
For i = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(i).SubItems(6)) / 100 * val(LV1.ListItems(i).SubItems(5)) * val(LV1.ListItems(i).SubItems(4))
Next i
XPText2.Text = Format(sum, "###0")

End Sub


Private Sub text4_GotFocus()
stn.Text = ""
If LV1.ListItems.count = 0 Then
ket.Caption = ""
Else
If text4.Text = "" Then
ket.Caption = "tekan enter bila barang sudah lengkap"
End If
End If

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" And LV1.ListItems.count <> 0 Then
disp.SetFocus
Else
text4_Click
End If
End If

End Sub

Private Sub ThemedButton1_Click()

End Sub

Private Sub Timer1_Timer()
jam.Caption = Format(Now, "hh:mm:ss")

End Sub
