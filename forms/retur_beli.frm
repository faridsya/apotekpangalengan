VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form retur_beli 
   Caption         =   "Retur Pembelian Barang"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9975
   Icon            =   "retur_beli.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   7215
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Retur Pembelian"
      TabPicture(0)   =   "retur_beli.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TCmdhapus"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdtambah"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Cmdkeluar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Cmdsimpan"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdbatal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LV1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "text7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Cmdcari"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ket"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Data barang"
      TabPicture(1)   =   "retur_beli.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrid1"
      Tab(1).Control(1)=   "kolom"
      Tab(1).Control(2)=   "kolom2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Data Supplier"
      TabPicture(2)   =   "retur_beli.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Skin1"
      Tab(2).Control(1)=   "dbgrid2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Data Pembelian"
      TabPicture(3)   =   "retur_beli.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dbgrid3"
      Tab(3).Control(1)=   "Command1"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Kembali barang/uang"
      TabPicture(4)   =   "retur_beli.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "keter"
      Tab(4).Control(1)=   "SkinLabel16"
      Tab(4).Control(2)=   "ttlr"
      Tab(4).Control(3)=   "SkinLabel13"
      Tab(4).Control(4)=   "balik"
      Tab(4).Control(5)=   "balu"
      Tab(4).Control(6)=   "dbgrid5"
      Tab(4).Control(7)=   "lv2"
      Tab(4).Control(8)=   "SkinLabel15"
      Tab(4).Control(9)=   "SkinLabel14"
      Tab(4).Control(10)=   "DTPicker2"
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "Data Retur Beli"
      TabPicture(5)   =   "retur_beli.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SkinLabel17"
      Tab(5).Control(1)=   "dbgrid4"
      Tab(5).ControlCount=   2
      Begin ACTIVESKINLibCtl.SkinLabel keter 
         Height          =   255
         Left            =   -68760
         OleObjectBlob   =   "retur_beli.frx":0D72
         TabIndex        =   58
         Top             =   4080
         Width           =   3615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   -74880
         OleObjectBlob   =   "retur_beli.frx":0DF8
         TabIndex        =   57
         Top             =   5760
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.SkinLabel ket 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "retur_beli.frx":0E82
         TabIndex        =   56
         Top             =   6480
         Width           =   8415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   -73320
         OleObjectBlob   =   "retur_beli.frx":0EE0
         TabIndex        =   54
         Top             =   600
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel ttlr 
         Height          =   255
         Left            =   -67560
         OleObjectBlob   =   "retur_beli.frx":0F4C
         TabIndex        =   47
         Top             =   6480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   -69960
         OleObjectBlob   =   "retur_beli.frx":0FAA
         TabIndex        =   46
         Top             =   6480
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   45
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8493
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
      Begin VB.CommandButton balik 
         Caption         =   "Kembali barang"
         Height          =   255
         Left            =   -74040
         TabIndex        =   43
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton balu 
         Caption         =   "Kembali &uang"
         Height          =   255
         Left            =   -72360
         TabIndex        =   42
         Top             =   6480
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   1815
         Left            =   -74040
         TabIndex        =   41
         Top             =   4440
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3201
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
      Begin VB.CommandButton Command1 
         Caption         =   "&Cari "
         Height          =   255
         Left            =   -72000
         TabIndex        =   40
         Top             =   6120
         Width           =   2535
      End
      Begin VB.ComboBox kolom2 
         Height          =   315
         Left            =   -72000
         TabIndex        =   34
         Top             =   6120
         Width           =   2055
      End
      Begin XPControls.XPCombo kolom 
         Height          =   315
         Left            =   -74880
         TabIndex        =   33
         Top             =   6120
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
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   6000
         TabIndex        =   25
         Top             =   4440
         Width           =   3135
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   252
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":104A
            TabIndex        =   26
            Top             =   240
            Width           =   1452
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":10CA
            TabIndex        =   27
            Top             =   1200
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel hb 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "retur_beli.frx":1144
            TabIndex        =   28
            Top             =   1200
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel stok 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "retur_beli.frx":11A2
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":1200
            TabIndex        =   36
            Top             =   720
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel teretur 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "retur_beli.frx":127C
            TabIndex        =   37
            Top             =   720
            Width           =   855
         End
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -71640
         OleObjectBlob   =   "retur_beli.frx":12DA
         Top             =   6240
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   840
         TabIndex        =   20
         Top             =   4320
         Width           =   4815
         Begin ACTIVESKINLibCtl.SkinLabel stnc 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "retur_beli.frx":150E
            TabIndex        =   61
            Top             =   960
            Width           =   1095
         End
         Begin XPControls.XPText alasan 
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   1680
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":156C
            TabIndex        =   35
            Top             =   1680
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":15D6
            TabIndex        =   31
            Top             =   1320
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":1640
            TabIndex        =   22
            Top             =   960
            Width           =   975
         End
         Begin XPControls.XPText text6 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
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
         Begin XPControls.XPText text5 
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   960
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
            Left            =   1200
            TabIndex        =   4
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":16B2
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin XPControls.XPCombo satuan 
            Height          =   315
            Left            =   1200
            TabIndex        =   59
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "retur_beli.frx":1726
            TabIndex        =   60
            Top             =   600
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   17
         Top             =   840
         Width           =   8175
         _ExtentX        =   14420
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
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   16
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
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
      Begin VB.CommandButton Cmdcari 
         Caption         =   "&Cari"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   4080
         Width           =   735
      End
      Begin VB.Frame Frame 
         Height          =   1575
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   8655
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   375
            Left            =   4680
            OleObjectBlob   =   "retur_beli.frx":1790
            TabIndex        =   38
            Top             =   840
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "retur_beli.frx":1806
            TabIndex        =   32
            Top             =   1200
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "retur_beli.frx":186E
            TabIndex        =   30
            Top             =   720
            Width           =   1335
         End
         Begin XPControls.XPText text8 
            Height          =   285
            Left            =   6360
            TabIndex        =   23
            Top             =   1200
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
         Begin XPControls.XPCombo text3 
            Height          =   315
            Left            =   6360
            TabIndex        =   3
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
            TabIndex        =   2
            Top             =   1080
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   128581635
            CurrentDate     =   40299
         End
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   1920
            TabIndex        =   1
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
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   1920
            TabIndex        =   9
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
            Left            =   4680
            OleObjectBlob   =   "retur_beli.frx":18E4
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "retur_beli.frx":195C
            TabIndex        =   13
            Top             =   1200
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel No 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "retur_beli.frx":19CA
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin XPControls.XPText tbrg 
            Height          =   285
            Left            =   6360
            TabIndex        =   39
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
      End
      Begin XPControls.XPText text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   6840
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
      Begin MSComctlLib.ListView LV1 
         Height          =   1695
         Left            =   480
         TabIndex        =   24
         Top             =   2160
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2990
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
         NumItems        =   9
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
            Text            =   "No faktur"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kode barang"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Barang"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Harga Beli"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Jumlah"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Sub Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Alasan"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   2535
         Left            =   -73320
         TabIndex        =   44
         ToolTipText     =   "Double Click to edit or Hapus data obat"
         Top             =   1440
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "No."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "No Pembelian"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tanggal Pembelian"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Supplier"
            Object.Width           =   3175
         EndProperty
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   -74400
         OleObjectBlob   =   "retur_beli.frx":1A42
         TabIndex        =   48
         Top             =   4080
         Width           =   3495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   -73320
         OleObjectBlob   =   "retur_beli.frx":1B00
         TabIndex        =   49
         Top             =   1080
         Width           =   3015
      End
      Begin apotekku.ThemedButton cmdbatal 
         Height          =   375
         Left            =   1920
         TabIndex        =   50
         Top             =   3960
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
         ForeColor       =   0
         MouseIcon       =   "retur_beli.frx":1BA6
      End
      Begin apotekku.ThemedButton Cmdsimpan 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Simpan"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MouseIcon       =   "retur_beli.frx":2140
      End
      Begin apotekku.ThemedButton Cmdkeluar 
         Height          =   375
         Left            =   4680
         TabIndex        =   51
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
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
         ForeColor       =   0
         MouseIcon       =   "retur_beli.frx":26DA
      End
      Begin apotekku.ThemedButton cmdtambah 
         Height          =   375
         Left            =   600
         TabIndex        =   0
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Tambah"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         MouseIcon       =   "retur_beli.frx":2C74
      End
      Begin apotekku.ThemedButton TCmdhapus 
         Height          =   375
         Left            =   7920
         TabIndex        =   52
         Top             =   3960
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
         ForeColor       =   0
         MouseIcon       =   "retur_beli.frx":320E
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   -72480
         TabIndex        =   53
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   128581635
         CurrentDate     =   40299
      End
      Begin MSDataGridLib.DataGrid dbgrid4 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   55
         Top             =   840
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   8493
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
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   6840
         Width           =   855
      End
   End
End
Attribute VB_Name = "retur_beli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ids, kode, kbr As String
Dim st, cosbli, cosju As Currency
Dim pr As Double

Private Sub alasan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sql = "select * from tblbarang where kode_brg='" & kode & "'"


Set RS = jual.Execute(sql)
If Not RS.EOF Then

If satuan.Text <> RS!satuan Then
 Set rsstn = New Recordset
 rsstn.Open "select * from satuan where kode_brg='" & kode & "' and satuan='" & satuan.Text & "'", jual, adOpenStatic, adLockOptimistic
 satuan.Text = rsstn!satuan
Text6.Text = val(Text6.Text) * val(rsstn!konversi)
End If
st = val(Text5.Text) * val(Text6.Text)



isigrid

ttl
kosong2
Text4.SetFocus
End If
End If

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


Private Sub balik_Click()
On Error GoTo erol
Dim jumbar As String

 jumbar = InputBox("Masukkan jumlah barang yang dikembalikan:")
  If StrPtr(jumbar) = 0 Then Exit Sub

Set RS = New Recordset
RS.Open "select * from tblbarang where kode_brg='" & dbgrid5.Columns(0) & "'", jual, adOpenStatic, adLockOptimistic
st = RS!stok
ns = st + val(jumbar)
jual.Execute "update tblbarang set stok=" & ns & " where kode_brg='" & dbgrid5.Columns(0) & "'"
Set RS = New Recordset
RS.Open "select * from detilbeli where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
If val(jumbar) > val(RS!teretur) Then
MsgBox "Jangan melebihi yang teretur"
Exit Sub
End If
ret = RS!kembali_brg
kemb = val(ret) + val(jumbar)
ter = RS!teretur - val(jumbar)
jual.Execute "update detilbeli set kembali_brg=" & kemb & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
jual.Execute "update detilbeli set teretur=" & ter & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"


jual.Execute "Insert into balik_brg1 values('" & Format(DTPicker2, "YYYY-mm-dd") & "','" & lv2.SelectedItem.SubItems(2) & "','" & dbgrid5.Columns(0) & "','" & val(jumbar) & "')"
LV2_Click
keter.Caption = "*Pilih di tabel atas"
erol:
If err.Description <> vbNullString Then

MsgBox "Pilih dulu di tabel tabel atas/faktur penjualan"
End If
End Sub

Private Sub balu_Click()

On Error GoTo erol
Dim hb, krg, krg2 As Currency

 jumbar = InputBox("Masukkan jumlah barang yang dikembalikan dengan uang:")
 If StrPtr(jumbar) = 0 Then Exit Sub
Set RS = New Recordset
Set RS = New Recordset
RS.Open "select * from detilbeli where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
If val(jumbar) > val(RS!teretur) Then
MsgBox "Jangan melebihi yang teretur"
Exit Sub
End If
pilih = ""
byr = val(jumbar) * val(dbgrid5.Columns(7).Text)
pembayaran = byr
If MsgBox("Potong faktur dan hutang terhadap supplier?", vbYesNo) = vbYes Then
Set rssupp = New Recordset
 rssupp.Open "select * from tblsupplier where supplier='" & lv2.SelectedItem.SubItems(4) & "'", jual, adOpenStatic, adLockOptimistic
 If rssupp!jumlah_hutang > 0 Then

   If rssupp!jumlah_hutang < byr Then

    If MsgBox("Jumlah hutang lebih kecil dari jumlah retur,sisanya akan dimasukkan ke keuangan.", vbInformation) = vbNo Then Exit Sub
     
     pembayaran = byr - rssupp!jumlah_hutang
     rssupp!jumlah_hutang = 0
     jual.Execute "update tblsupplier set jumlah_hutang=0 where supplier='" & lv2.SelectedItem.SubItems(4) & "'"
     ret = RS!kembali_uang

     valid
     tanya.Show
     Exit Sub
    Else
   If rssupp!jumlah_hutang = byr Then
     jual.Execute "update tblsupplier set jumlah_hutang=0 where supplier='" & lv2.SelectedItem.SubItems(4) & "'"
ret = RS!kembali_uang

valid
   Else

        juhu = rssupp!jumlah_hutang - byr
             jual.Execute "update tblsupplier set jumlah_hutang=" & juhu & " where supplier='" & lv2.SelectedItem.SubItems(4) & "'"

ret = RS!kembali_uang

valid
    End If
    End If
    
    
    
    
 Else
 MsgBox "Tidak punya hutang terhadap supplier ini", vbInformation
 'Exit Sub
End If
LV2_Click

 dbgridsret
dbgridss
keter.Caption = "*Pilih di tabel atas"

Else
tanya.Show
End If
erol:
If err.Description <> vbNullString Then

MsgBox "Pilih dulu di tabel daftar detil penjualan atau tabel pembelian atas"
End If
End Sub
Private Sub valid()
kembu = RS.Fields("kembali_uang") + val(jumbar)
trt = RS!teretur - val(jumbar)
hali = RS!harga_beli

totur = RS!total_retur + val(jumbar) * hali

jual.Execute "update detilbeli set kembali_uang=" & kembu & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
jual.Execute "update detilbeli set total_retur=" & totur & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
jual.Execute "update detilbeli set teretur=" & trt & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"

krg2 = val(jumbar) * RS!harga_beli

 
 
 Set RS = New Recordset
RS.Open "select * from hutang where no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
 If Not RS.EOF Then
 
 juhu = RS!jumlah_hutang - krg2
 jual.Execute "update hutang set jumlah_hutang=" & juhu & " where no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
End If
End Sub
Private Sub valid2()
kembu = RS.Fields("kembali_uang2") + val(jumbar)
trt = RS!teretur - val(jumbar)
hali = RS!harga_beli

totur = RS!total_retur + val(jumbar) * hali

jual.Execute "update detilbeli set kembali_uang2=" & kembu & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
jual.Execute "update detilbeli set total_retur=" & totur & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"
jual.Execute "update detilbeli set teretur=" & trt & " where kode_brg='" & dbgrid5.Columns(0) & "' and no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'"


 
 
End Sub

Private Sub cmdbatal_Click()
awal
kosong
cmdtambah.SetFocus
End Sub

Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
frame.Enabled = True
Frame2.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True

End Sub

Private Sub cmdhapus_Click()
If Not lv1.SelectedItem Is Nothing Then
lv1.ListItems.Remove lv1.SelectedItem.Index
End If
ttl
For x = 1 To lv1.ListItems.count
lv1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
If Edit = True Then
jual.Execute "update detilreturbeli set jumlah=" & val(Text6.Text) & " where no_retur='" & Text1.Text & "'"
jual.Execute "update detilbeli set teretur=" & val(Text6.Text) & " where no_pembelian='" & Text2.Text & "' and kode_brg='" & kbr & "'"

Set rse1 = New Recordset
rse1.Open "Select * from tblbarang where kode_brg='" & kbr & "'", jual, adOpenStatic, adLockOptimistic
stk = rse1!stok + (pr - val(Text6.Text))
jual.Execute "update tblbarang set stok=" & stk & " where kode_brg='" & kbr & "'"
awal
MsgBox "Data berhasil diubah"
dbgridr

Else
simpandata
dbgridr
dbgridsret
awal
MsgBox "Data retur beli berhasil disimpan"
End If
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
sql = "insert into retur_beli values('" & Text1.Text & "','" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ids & "','" & tbrg.Text & "','" & text8.Text & "')"

jual.Execute (sql)
For z = 1 To lv1.ListItems.count
Set RS = New Recordset
RS.Open "select * from tblbarang where kode_brg='" & lv1.ListItems(z).SubItems(3) & "'", jual, adOpenStatic, adLockOptimistic
st = RS!stok
ns = st - lv1.ListItems(z).SubItems(6)
jual.Execute "update tblbarang set stok=" & ns & " where kode_brg='" & lv1.ListItems(z).SubItems(3) & "'"

sql = "insert into detilreturbeli values('" & Text1.Text & "','" & lv1.ListItems(z).SubItems(2) & "','" & _
lv1.ListItems(z).SubItems(3) & "','" & lv1.ListItems(z).SubItems(4) & "','" & lv1.ListItems(z).SubItems(6) & "','" & _
lv1.ListItems(z).SubItems(5) & "','" & lv1.ListItems(z).SubItems(7) & "','" & lv1.ListItems(z).SubItems(8) & "')"

jual.Execute (sql)
Set RS = New Recordset
RS.Open "select * from detilbeli where kode_brg='" & lv1.ListItems(z).SubItems(3) & "' and no_pembelian='" & lv1.ListItems(z).SubItems(2) & "'", jual, adOpenStatic, adLockOptimistic
ret = RS!teretur
ttr = val(ret) + lv1.ListItems(z).SubItems(6)

jual.Execute "update detilbeli set teretur=" & ttr & " where kode_brg='" & lv1.ListItems(z).SubItems(3) & "' and no_pembelian='" & lv1.ListItems(z).SubItems(2) & "'"


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
cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
cmdbatal.Enabled = True
Cmdcari.Enabled = False
lv1.Enabled = False
Frame2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
sql = "Select no_retur from retur_beli order by no_retur Desc"
Set RS = jual.Execute(sql)
If RS.EOF = True Then
Text1.Text = "RB-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
Else
j = val(Right(RS(0), 4))
br = "RB-" + Format(Now, "YY-") + Format(Now, "MM") + "-" + Format(Str(j + 1), "0000")
Text1.Text = br
nmr = Format(Mid(RS(0), 7, 2))
thn = Format(Mid(RS(0), 4, 2))
txt = Format(Mid(No, 7, 2))
thnn = Format(Mid(No, 4, 2))

If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
Text1.Text = "RB-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
End If
End If
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
text7.Text = ""
text8.Text = ""
tbrg.Text = ""
alasan.Text = ""
satuan.Text = ""
stnc.Caption = ""
lv1.ListItems.Clear

End Sub
Sub kosong2()
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
stok = ""
teretur = ""
hb = ""
hj = ""
alasan.Text = ""
satuan.Text = ""

stnc.Caption = ""
End Sub

Private Sub Command1_Click()
Dim kata As String

If Cmdcari.Caption = "&Cari" Then
kata = InputBox("Masukkan kata kunci", "Cari...")
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

Private Sub Command2_Click()
retur.Show
End Sub

Private Sub dbgrid3_DblClick()
On Error Resume Next

If Edit = False Then
Text2.Text = dbgrid3.Columns(0)
Text2_KeyPress (13)
Tab1.Tab = 0
Text4.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus

End If

End Sub

Private Sub dbgrid4_Click()
Set RS = New Recordset
RS.Open "select detilbeli.kode_brg,tblbarang.deskripsi,detilbeli.jumlah_brg,detilbeli.teretur,detilbeli.kembali_brg,detilbeli.kembali_uang,detilbeli.kembali_uang2  from detilbeli,tblbarang where detilbeli.kode_brg=tblbarang.kode_brg and detilbeli.no_pembelian='" & dbgrid4.Columns(0).Text & "'and detilbeli.teretur<>0", jual, adOpenStatic, adLockOptimistic
Set dbgrid5.DataSource = RS

End Sub


Private Sub dbgrid5_Click()
On Error Resume Next
ttlr.Caption = val(dbgrid5.Columns(3)) + val(dbgrid5.Columns(4)) + val(dbgrid5.Columns(5)) + val(dbgrid5.Columns(6))
keter.Caption = "*Lalu pilih kembali barang atau potong faktur"
End Sub

Private Sub Form_Activate()
If pilih = "KAS" Or pilih = "BANK" Then
byr = val(jumbar) * val(dbgrid5.Columns(7).Text)


valid2

    If pilih = "KAS" Then

jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan) values('" & Format(DTPicker2, "YYYY-mm-dd") & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "','" & pembayaran & "')"
 Else
 
 If pilih = "BANK" Then
   If junai = pembayaran Then
jual.Execute "Insert into keuangan2(Tanggal2,keterangan2,pemasukan2,kode_bank,bentuk) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "','" & pembayaran & "','" & idb & "','Tunai')"
   Else
   If junai = 0 Then
jual.Execute "Insert into keuangan2(Tanggal2,keterangan2,pemasukan2,kode_bank,bentuk) values('" & Format(gtgl, "YYYY-mm-dd") & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "','" & pembayaran & "','" & idb & "','Giro')"
jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_masuk,no_faktur,keterangan) values('" & Format(DTPicker2, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & gnom & "','" & Dbgrid2.Columns(0).Text & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "')"

Else
jual.Execute "Insert into keuangan2(Tanggal2,keterangan2,pemasukan2,kode_bank,bentuk) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "','" & junai & "','" & idb & "','Tunai')"
jual.Execute "Insert into keuangan2(Tanggal2,keterangan2,pemasukan2,kode_bank,bentuk) values('" & Format(gtgl, "YYYY-mm-dd") & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "','" & jugir & "','" & idb & "','Giro')"
  jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_masuk,no_faktur,keterangan) values('" & Format(DTPicker2, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & jugir & "','" & Dbgrid2.Columns(0).Text & "','Retur beli " & lv2.SelectedItem.SubItems(2) & "')"

  End If
End If
End If

 
 
End If
 LV2_Click

 dbgridsret
keter.Caption = "*Pilih di tabel atas"

End If
End Sub

Private Sub Form_Load()
Tab1.Tab = 0

Edit = True
supp
awal
Ketengah Me
kolom.Text = "Semua"
pilih = ""

 dbgrids
 dbgridss
 dbgridsret
 dbgridr
 DTPicker1 = Format(Now)
  DTPicker2 = Format(Now)

     Skinpath = App.Path & "\skin\green.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
kolom.Clear
kolom.AddItem "Semua"
kolom.AddItem "Deskripsi"
kolom.AddItem "Kategori"
kolom.Text = "Semua"
End Sub
Private Sub dbgridss()
Set rsbeli = New Recordset
sql = "select pembelian.no_pembelian as No_faktur,tblsupplier.supplier,pembelian.tanggal_pembelian ,pembelian.total_stlh_diskon as Total from pembelian,tblsupplier where pembelian.id_supplier=tblsupplier.id_supplier order by pembelian.tanggal_pembelian"
Set rsbeli = jual.Execute(sql)

Set dbgrid3.DataSource = rsbeli

End Sub
Private Sub dbgridr()
Set rsr = New Recordset
sql = "Select detilreturbeli.no_retur,pembelian.no_pembelian as No_Faktur,tblsupplier.supplier,detilreturbeli.kode_brg,detilreturbeli.nama_barang,detilreturbeli.jumlah,detilreturbeli.harga_beli from pembelian,detilreturbeli,tblsupplier where detilreturbeli.no_pembelian=pembelian.no_pembelian and pembelian.id_supplier=tblsupplier.id_supplier"
Set rsr = jual.Execute(sql)
Set dbgrid4.DataSource = rsr
End Sub
Private Sub dbgrid4_DblClick()
On Error Resume Next
Set rse2 = New Recordset
rse2.Open "Select* from detilbeli where no_pembelian='" & dbgrid4.Columns(1).Text & "' and kode_brg='" & dbgrid4.Columns(3).Text & "'", jual, adOpenStatic, adLockOptimistic
If rse2!kembali_brg <> 0 Or rse2!kembali_uang <> 0 Then
MsgBox "Tidak bisa dirubah,sudah pernah dikembalikan", vbInformation
Exit Sub
End If
Edit = True
pr = val(dbgrid4.Columns(5).Text)
kbr = dbgrid4.Columns(3).Text
cmdsimpan.Enabled = True
Text1.Text = dbgrid4.Columns(0).Text
Text2.Text = dbgrid4.Columns(1).Text
Text3.Text = dbgrid4.Columns(2).Text
Text4.Text = dbgrid4.Columns(4).Text
Text5.Text = dbgrid4.Columns(6).Text
Text6.Text = dbgrid4.Columns(5).Text
frame.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = True
Frame2.Enabled = True
Tab1.Tab = 0
End Sub

Private Sub dbgridsret()
Set RS = New Recordset
    kata = "select pembelian.no_pembelian,pembelian.tanggal_pembelian,tblsupplier.supplier from pembelian,detilbeli,tblsupplier where pembelian.id_supplier=tblsupplier.id_supplier and pembelian.no_pembelian=detilbeli.no_pembelian and detilbeli.teretur<>0 "
    Set RS = New Recordset
        RS.Open kata, jual, adOpenStatic, adLockOptimistic
        lv2.ListItems.Clear
        If Not RS.EOF Then
          RS.MoveFirst
            I = 1
            While Not RS.EOF
                Set butir = lv2.ListItems.Add
                With butir
                .SubItems(1) = lv2.ListItems.count & "."

    .SubItems(2) = RS![no_pembelian]
    .SubItems(3) = RS![tanggal_pembelian]
    .SubItems(4) = RS![Supplier]
                    RS.MoveNext
                    I = I + 1
                   
    End With
                    
                Wend
            End If
           RS.Close
            Set RS = Nothing
            Me.MousePointer = 1
            
                With lv2
    For I = 1 To .ListItems.count
      For j = .ListItems.count To (I + 1) Step -1
         If .ListItems(j).SubItems(2) = .ListItems(I).SubItems(2) Then
         .ListItems.Remove (I)

         End If
      Next j
    Next I
  End With

For x = 1 To lv2.ListItems.count
lv2.ListItems(x).SubItems(1) = x & "."
Next x

   Exit Sub
salah:
    MsgBox err.Description

End Sub


Private Sub kbrg()

Text4.Clear
sql = "select * from tblbarang,detilbeli,pembelian where pembelian.no_pembelian='" & Text2.Text & "' and pembelian.no_pembelian=detilbeli.no_pembelian and tblbarang.kode_brg=detilbeli.kode_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text4.AddItem rsbarang!deskripsi
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
Sub dbgrids()
Set rssupp = New Recordset

sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set Dbgrid2.DataSource = rssupp


End Sub

Private Sub supp()

  Dim I As Long
  Dim j As Long

Text3.Clear
sql = "select * from tblsupplier order by id_supplier"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text3.AddItem rsbarang!Supplier
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub
Private Sub awal()
frame.Enabled = False
Frame2.Enabled = False
Edit = True
cmdsimpan.Enabled = False
cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True
lv1.Enabled = True
End Sub

Private Sub hb_Click()
Text5.Text = hb.Caption
Text6.SetFocus
End Sub




Private Sub LV2_Click()
If lv2.ListItems.count <> 0 Then

Set RS = New Recordset
RS.Open "select detilbeli.kode_brg,tblbarang.deskripsi,detilbeli.jumlah_brg,detilbeli.teretur,detilbeli.kembali_brg,detilbeli.kembali_uang,detilbeli.kembali_uang2,detilbeli.harga_beli from detilbeli,tblbarang where detilbeli.kode_brg=tblbarang.kode_brg and detilbeli.no_pembelian='" & lv2.SelectedItem.SubItems(2) & "'and detilbeli.teretur<>0", jual, adOpenStatic, adLockOptimistic
Set dbgrid5.DataSource = RS
End If
keter.Caption = "*Lalu pilih di tabel bawah"
End Sub

Private Sub satuan_Click()
Text6.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New Recordset
RS.Open "select tblsupplier.id_supplier,tblsupplier.supplier from pembelian,tblsupplier where pembelian.id_supplier=tblsupplier.id_supplier and no_pembelian='" & Text2.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Tidak ada data nomor faktur ini", vbInformation
Else
Text3.Text = RS!Supplier
ids = RS!id_supplier
kbrg
End If
RS.Close
End If
End Sub

Private Sub text3_Click()
Text4.SetFocus
End Sub

Private Sub text4_Click()
On Error Resume Next
Set cari = lv1.FindItem(Text4.Text, 1, , 1)
Text4.Text = Replace(Text4.Text, "'", "''")

sql = "select * from tblbarang where deskripsi='" & Text4.Text & "' or kode_brg='" & Text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
kode = rsbarang!kode_brg
stn_baku = rsbarang!satuan
stnc.Caption = "/" & rsbarang!satuan
tampil_stn
satuan.Text = rsbarang!satuan

End If

sql = "select detilbeli.jumlah_brg,detilbeli.teretur,detilbeli.harga_beli,pembelian.tanggal_pembelian from detilbeli,pembelian,tblbarang where pembelian.no_pembelian=detilbeli.no_pembelian and detilbeli.kode_brg=tblbarang.kode_brg and tblbarang.deskripsi='" & Text4.Text & "' and pembelian.no_pembelian='" & Text2.Text & "'"
Set RS = New Recordset
Set RS = jual.Execute(sql)
If Not RS.EOF Then
Text5.Text = RS.Fields("harga_beli")

stok.Caption = RS!jumlah_brg
teretur.Caption = RS!teretur
hb = RS!tanggal_pembelian
End If
Text6.SetFocus



End Sub
Sub databaru()
Barang.Show
Barang.tambah
Barang.Text2.Text = Text4.Text
Barang.Text4.Text = "0"
Barang.kode_oto
Barang.Text3.SetFocus
End Sub


Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text4.Text = "" Then
cmdsimpan.SetFocus
Else
text4_Click
End If
End If

End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.Text = cosbli
Text6.SetFocus
End If
End Sub

Private Sub Text6_Change()
If Edit = True Then Exit Sub
If val(Text6.Text) > (val(stok) - val(teretur)) Then
MsgBox "Melebihi stok pembelian", vbInformation
Text6.Text = ""
End If

End Sub

Sub ttl()
sum = 0
sum2 = 0
For I = 1 To lv1.ListItems.count
sum = sum + val(lv1.ListItems(I).SubItems(7))
sum2 = sum2 + val(lv1.ListItems(I).SubItems(6))

Next I
text8.Text = Format(sum, "###0")
tbrg.Text = sum2
End Sub

Sub isigrid()
    Set butir = lv1.ListItems.Add
    With butir
           .SubItems(1) = lv1.ListItems.count & "."
    .SubItems(2) = Text2.Text

    .SubItems(3) = kode
    .SubItems(4) = Text4.Text
    .SubItems(5) = Text5.Text
    .SubItems(6) = Text6.Text
    .SubItems(7) = st
        .SubItems(8) = alasan.Text

    End With
  lv1.Enabled = True
   With butir
End With
End Sub

Private Sub sort()
If kolom.Text = "Semua" Then
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
kolom2.AddItem rsbarang!kategori
End If
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

Private Sub text6_GotFocus()
If Edit = True Then
ket.Caption = "Rubah jumlah yang diretur lalu tekan simpan"
Else
ket.Caption = "Masukkan jumlah/berat barang yang diretur"
End If
End Sub

Private Sub text6_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka2(KeyAscii)
If KeyAscii = 13 Then
alasan.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
ket.Caption = ""
End Sub
