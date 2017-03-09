VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreju 
   Caption         =   "RETUR JUAL"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10230
   Icon            =   "frmreju.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10230
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10320
      OleObjectBlob   =   "frmreju.frx":0CCA
      Top             =   6120
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11668
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Input data"
      TabPicture(0)   =   "frmreju.frx":0EFE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TCmdhapus"
      Tab(0).Control(1)=   "cmdkeluar"
      Tab(0).Control(2)=   "cmdsimpan"
      Tab(0).Control(3)=   "cmdbatal"
      Tab(0).Control(4)=   "cmdtambah"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(7)=   "Frame"
      Tab(0).Control(8)=   "LV1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Data penjualan"
      TabPicture(1)   =   "frmreju.frx":0F1A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SkinLabel7"
      Tab(1).Control(1)=   "SkinLabel4"
      Tab(1).Control(2)=   "txtcari"
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(4)=   "ListView1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Data Retur"
      TabPicture(2)   =   "frmreju.frx":0F36
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lv3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lv2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtcari2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "SkinLabel10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton TCmdhapus 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   -67320
         TabIndex        =   41
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdkeluar 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   -70560
         TabIndex        =   40
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   -72000
         TabIndex        =   39
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   -73080
         TabIndex        =   38
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   -74400
         TabIndex        =   37
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Bata&lkan"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   3360
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   5520
         OleObjectBlob   =   "frmreju.frx":0F52
         TabIndex        =   33
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtcari2 
         Height          =   285
         Left            =   6720
         TabIndex        =   32
         Top             =   3480
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   -74760
         OleObjectBlob   =   "frmreju.frx":0FCA
         TabIndex        =   31
         Top             =   3480
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   -70920
         OleObjectBlob   =   "frmreju.frx":105C
         TabIndex        =   28
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   -68160
         TabIndex        =   27
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   -69000
         TabIndex        =   20
         Top             =   4560
         Width           =   3135
         Begin ACTIVESKINLibCtl.SkinLabel lbldi 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmreju.frx":10FA
            TabIndex        =   36
            Top             =   600
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":1158
            TabIndex        =   35
            Top             =   600
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":11C4
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":1244
            TabIndex        =   22
            Top             =   1080
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel hb 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmreju.frx":12BE
            TabIndex        =   23
            Top             =   1080
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmreju.frx":131C
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   -74400
         TabIndex        =   16
         Top             =   4440
         Width           =   5295
         Begin XPControls.XPText alasan 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   960
            Width           =   3375
            _ExtentX        =   5953
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
            Top             =   600
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
            Left            =   1680
            TabIndex        =   1
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":137A
            TabIndex        =   17
            Top             =   960
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":13E4
            TabIndex        =   18
            Top             =   600
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmreju.frx":144E
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame 
         Height          =   1575
         Left            =   -74520
         TabIndex        =   5
         Top             =   360
         Width           =   8655
         Begin XPControls.XPCombo text3 
            Height          =   315
            Left            =   6360
            TabIndex        =   6
            Top             =   240
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
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   7
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   119209987
            CurrentDate     =   40299
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   0
            Top             =   720
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
         Begin XPControls.XPText text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   8
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
         Begin XPControls.XPText tbrg 
            Height          =   285
            Left            =   6360
            TabIndex        =   9
            Top             =   720
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   375
            Left            =   4680
            OleObjectBlob   =   "frmreju.frx":14C6
            TabIndex        =   10
            Top             =   720
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmreju.frx":153C
            TabIndex        =   11
            Top             =   720
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "frmreju.frx":15AC
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmreju.frx":1620
            TabIndex        =   13
            Top             =   1080
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel No 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "frmreju.frx":1698
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   1935
         Left            =   -74520
         TabIndex        =   15
         Top             =   1920
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3413
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
         NumItems        =   5
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
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Jumlah"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Alasan"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   25
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
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
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   26
         Top             =   4200
         Width           =   9495
         _ExtentX        =   16748
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
            Weight          =   400
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
            Text            =   "Sub total"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   3015
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
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
            Text            =   "No Retur"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Jumlah brg"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv3 
         Height          =   2535
         Left            =   240
         TabIndex        =   30
         Top             =   3840
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4471
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
            Weight          =   400
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
            Text            =   "No Faktur"
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
      End
   End
End
Attribute VB_Name = "frmreju"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim idp, kode, kbr, stts As String
Dim st, cosju As Currency
Dim pr As Double
Sub dbgridtrans()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select penjualan.no_penjualan,penjualan.tanggal,penjualan.jumlah,penjualan.total_diskon,penjualan.total,penjualan.id_pelanggan,pelanggan.nama from penjualan,pelanggan where penjualan.id_pelanggan=pelanggan.id_pelanggan order by no_penjualan desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_penjualan]
        l.SubItems(2) = Format(![Tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![id_pelanggan]
                              
                l.SubItems(4) = ![nama]
                l.SubItems(5) = Format(![jumlah], "#,#")
                l.SubItems(6) = Format(![total_diskon], "#,#")

  l.SubItems(7) = Format(![total], "#,#")

    .MoveNext
    Loop
End With

End Sub
Sub dbgridtrans2()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select penjualan.no_penjualan,penjualan.tanggal,penjualan.jumlah,penjualan.total_diskon,penjualan.total,penjualan.id_pelanggan,pelanggan.nama from penjualan,pelanggan where penjualan.id_pelanggan=pelanggan.id_pelanggan and (no_penjualan like '%" & txtcari.Text & "%' or nama like '%" & txtcari.Text & "%') order by no_penjualan desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![no_penjualan]
        l.SubItems(2) = Format(![Tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![id_pelanggan]
                              
                l.SubItems(4) = ![nama]
                l.SubItems(5) = ![jumlah]
                l.SubItems(6) = Format(![total_diskon], "#,#")

  l.SubItems(7) = Format(![total], "#,#")


    .MoveNext
    Loop
End With

End Sub
Sub dbgridreju()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select * from retur_jual order by no_retur desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lv2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
        l.SubItems(1) = ![no_retur]
        l.SubItems(2) = Format(![Tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![total_brg]
                              l.SubItems(4) = Format(![total], "#,#")

    .MoveNext
    Loop
End With

End Sub
Sub dbgridreju2()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select * from retur_jual where no_retur like '%" & txtcari2.Text & "%' order by no_retur desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lv2.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
        l.SubItems(1) = ![no_retur]
        l.SubItems(2) = Format(![Tanggal], "dd MMM yyyy")
                                l.SubItems(3) = ![total_brg]
                              l.SubItems(4) = Format(![total], "#,#")

    .MoveNext
    Loop
End With

End Sub



Private Sub satuan_Click()
text6.SetFocus
End Sub

Private Sub Command1_Click()
If MsgBox("Yakin batalkan retur ini?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from retur_jual where no_retur='" & lv2.SelectedItem.SubItems(1) & "'"
dbgridreju
lv3.ListItems.Clear
MsgBox "Retur jual berhasil dibatalkan.", vbInformation, judul


End Sub

Private Sub ListView2_Click()
'On Error Resume Next
text4.Enabled = True
If ListView2.ListItems.count = 0 Then Exit Sub

ListView1.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select detiljual.kode_brg,tblbarang.deskripsi,detiljual.harga_beli,detiljual.harga_jual,detiljual.jumlah_brg,detiljual.total from detiljual,tblbarang where detiljual.no_penjualan='" & ListView2.SelectedItem.SubItems(1) & "' and tblbarang.kode_brg=detiljual.kode_brg", jual, adOpenStatic, adLockOptimistic
If Not rse3.EOF Then

With rse3
.MoveFirst
    Do While Not .EOF
   Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
   l.SubItems(1) = ListView1.ListItems.count & "."
        l.SubItems(2) = !kode_brg
              l.SubItems(3) = ![deskripsi]
                           

              l.SubItems(4) = Format(![harga_jual], "#,#")

                l.SubItems(5) = ![jumlah_brg]
              

                l.SubItems(6) = Format(![total], "#,#")


    .MoveNext
    Loop
End With

End If

End Sub

Private Sub ListView2_DblClick()
If ListView2.ListItems.count = 0 Then Exit Sub
If cmdtambah.Enabled = True Then
cmdtambah_Click
End If

Text2.Text = ListView2.SelectedItem.SubItems(1)
Tab1.Tab = 0
Text2_KeyPress (13)

End Sub

Private Sub LV2_Click()
'On Error Resume Next
text4.Enabled = True
If lv2.ListItems.count = 0 Then Exit Sub

lv3.ListItems.Clear

Set rse3 = New Recordset
rse3.Open "Select * from detilreturjual,tblbarang where detilreturjual.no_retur='" & lv2.SelectedItem.SubItems(1) & "' and tblbarang.kode_brg=detilreturjual.kode_brg", jual, adOpenStatic, adLockOptimistic
If Not rse3.EOF Then

With rse3
.MoveFirst
    Do While Not .EOF
   Set l = lv3.ListItems.Add(, , lv3.ListItems.count + 1)
   l.SubItems(1) = lv3.ListItems.count & "."
        l.SubItems(2) = !kode_brg
        l.SubItems(3) = ![deskripsi]
        l.SubItems(4) = ![no_penjualan]
        l.SubItems(5) = Format(![harga_jual], "#,#")
        l.SubItems(6) = ![jumlah]
        
        l.SubItems(7) = Format(![diskon], "#,#")
        l.SubItems(8) = Format(![total], "#,#")


    .MoveNext
    Loop
End With

End If

End Sub

Private Sub lvbrg_DblClick()
If lvbrg.ListItems.count = 0 Then Exit Sub
If cmdtambah.Enabled = True Then
cmdtambah_Click
End If

Tab1.Tab = 0
text4.Text = lvbrg.SelectedItem.SubItems(1)
text4_Click
text6.SetFocus
End Sub

Private Sub lvbrg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvbrg_DblClick
End If
End Sub


Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 3 Then
txtcrp.SetFocus
Else
If Tab1.Tab = 4 Then
txtcrb.SetFocus
End If

End If
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
  l.SubItems(7) = Format(![Harga_grosir], "#,#")

    .MoveNext
    Loop
End With

End Sub
Sub dbgridbrg2()
On Error Resume Next

Set rstrans = New Recordset

stri = Replace(txtcrb.Text, "'", "''")

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
  l.SubItems(7) = Format(![Harga_grosir], "#,#")

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
        l.SubItems(5) = ![disk_faktur]

    .MoveNext
    Loop
End With


End Sub
Sub dbgridplg2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from pelanggan where nama like '%" & txtcrp.Text & "%' or id_pelanggan like '" & txtcrp.Text & "%' order by nama"
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
        l.SubItems(5) = ![disk_faktur]

    .MoveNext
    Loop
End With


End Sub

Private Sub TCmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New Recordset
RS.Open "select pelanggan.id_pelanggan,pelanggan.nama from penjualan,pelanggan where penjualan.id_pelanggan=pelanggan.id_pelanggan and penjualan.no_penjualan='" & Text2.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Tidak ada data nomor faktur ini", vbInformation
Else
text3.Text = RS!nama
idp = RS!id_pelanggan
kbrg
End If
RS.Close
End If
End Sub

Private Sub text4_Click()
sql = "select * from tblbarang where deskripsi='" & text4.Text & "' or kode_brg='" & text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
kode = rsbarang!kode_brg
End If

Set RS2 = New Recordset
RS2.Open "select jumlah_brg,tanggal from penjualan,detiljual where penjualan.no_penjualan=detiljual.no_penjualan and penjualan.no_penjualan='" & Text2.Text & "' and kode_brg='" & kode & "'", jual, adOpenStatic, adLockOptimistic
If RS2.EOF Then Exit Sub
stok.Caption = RS2!jumlah_brg
hb.Caption = Format(RS2!Tanggal, "dd mmm yyyy")
Set RS2 = New Recordset
RS2.Open "select coalesce(sum(detilreturjual.jumlah),0) as jum from detilreturjual,penjualan where penjualan.no_penjualan=detilreturjual.no_penjualan and detilreturjual.no_penjualan='" & Text2.Text & "' and kode_brg='" & kode & "'", jual, adOpenStatic, adLockOptimistic
lbldi.Caption = RS2!jum
text6.SetFocus




End Sub



Private Sub alasan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If val(text6.Text) <= 0 Then Exit Sub
sql = "select * from tblbarang where kode_brg='" & kode & "'"


Set RS = jual.Execute(sql)
If Not RS.EOF Then
isigrid

ttl
kosong2
text4.SetFocus
End If
End If
End Sub
Sub isigrid()

    Set butir = LV1.ListItems.Add
    With butir
           .SubItems(1) = LV1.ListItems.count & "."
    
    

    .SubItems(2) = kode
    .SubItems(3) = text6.Text
    .SubItems(4) = alasan.Text
 
    End With
  LV1.Enabled = True
End Sub




Private Sub Cmdbatal_Click()
awal
kosong
cmdtambah.SetFocus
End Sub


Private Sub cmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If
ttl
For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdsimpan_Click()
'On Error GoTo erol
If MsgBox("Simpan data retur?", vbYesNo, judul) = vbNo Then Exit Sub
If LV1.ListItems.count = 0 Then Exit Sub
simpandata
awal
MsgBox "Data retur jual berhasil disimpan"
dbgridreju
awal
cmdtambah.SetFocus
erol:
If err.Description <> vbNullString Then
MsgBox "Data belum lengkap", vbCritical, "Penjualan"
Exit Sub
End If

End Sub
Sub simpandata()
Set RS = New Recordset
sql = "insert into retur_jual(no_retur,tanggal,no_penjualan,id_pelanggan) values('" & text1.Text & "','" & Format(DTPicker1, "YYYY-mm-dd") & "','" & Text2.Text & "','" & text3.Text & "')"

jual.Execute (sql)
For z = 1 To LV1.ListItems.count
'cek lagi
sql = "insert into detilreturjual(no_retur,no_penjualan,kode_brg,jumlah,alasan) values('" & text1.Text & "','" & Text2.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(3) & "','" & LV1.ListItems(z).SubItems(4) & "')"
jual.Execute (sql)
 Next z


End Sub
Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
kosong2
no_oto
Text2.SetFocus


End Sub
Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
Cmdbatal.Enabled = True
LV1.Enabled = False
Frame2.Enabled = True

End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
sql = "Select no_retur from retur_jual order by no_retur Desc"
Set RS = jual.Execute(sql)
If RS.EOF = True Then
text1.Text = "RJ-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
Else
j = val(Right(RS(0), 4))
br = "RJ-" + Format(Now, "YY-") + Format(Now, "MM") + "-" + Format(Str(j + 1), "0000")
text1.Text = br
nmr = Format(Mid(RS(0), 7, 2))
thn = Format(Mid(RS(0), 4, 2))
txt = Format(Mid(No, 7, 2))
thnn = Format(Mid(No, 4, 2))

If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
text1.Text = "RJ-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
End If
End If
End Sub
Sub kosong()
text1.Text = ""
Text2.Text = ""
text3.Text = ""
text4.Text = ""
text6.Text = ""
tbrg.Text = ""
alasan.Text = ""

LV1.ListItems.Clear

End Sub
Sub kosong2()
text4.Text = ""
text6.Text = ""
stok = ""
teretur = ""
lbldi.Caption = ""
hb = ""
hj = ""
alasan.Text = ""
End Sub
Private Sub Form_Load()
Tab1.Tab = 0
pilih = ""
dbgridtrans
dbgridreju
Edit = True
awal
Ketengah Me
 DTPicker1 = Format(Now)
  DTPicker2 = Format(Now)

    Skinpath = App.Path & "\skin\B-Studio.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub

Private Sub lvplg_DblClick()
If lvplg.ListItems.count = 0 Then Exit Sub
If cmdtambah.Enabled = True Then
cmdtambah_Click
End If
Tab1.Tab = 0
text3.Text = lvplg.SelectedItem.SubItems(1)
text4.SetFocus
End Sub

Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub


Private Sub kbrg()

text4.Clear
sql = "select * from tblbarang,detiljual,penjualan where penjualan.no_penjualan='" & Text2.Text & "' and penjualan.no_penjualan=detiljual.no_penjualan and tblbarang.kode_brg=detiljual.kode_brg"
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

Sub dbgrid()
Set rsbarang = New Recordset
sql = "select * from tblbarang"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang


End Sub
Private Sub awal()
frame.Enabled = False
Frame2.Enabled = False
Edit = True
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
LV1.Enabled = True
End Sub

Private Sub hb_Click()
Text5.Text = hb.Caption
text6.SetFocus
End Sub








Sub databaru()
Barang.Show
Barang.tambah
Barang.Text2.Text = text4.Text
Barang.text4.Text = "0"
Barang.kode_oto
Barang.text3.SetFocus
End Sub


Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" Then
Cmdsimpan.SetFocus
Else
text4_Click
End If
End If

End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.Text = cosju
text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
If Edit = True Then Exit Sub

'If val(Text6.Text) > (val(stok) - val(lbldi.Caption)) Then
'MsgBox "Melebihi stok beli penjualan", vbInformation
'Text6.Text = ""
'Text6.SetFocus
'End If

End Sub

Sub ttl()
sum2 = 0
For I = 1 To LV1.ListItems.count
sum2 = sum2 + val(LV1.ListItems(I).SubItems(4))

Next I
tbrg.Text = sum2
End Sub



Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 13 Then
alasan.SetFocus
End If
End Sub



Private Sub txtcari_Change()
dbgridtrans2

End Sub

Private Sub txtcari2_Change()
dbgridreju2
End Sub

Private Sub txtcrb_Change()
dbgridbrg2
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
