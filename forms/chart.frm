VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form chart 
   Caption         =   "Form2"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14265
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   6120
      OleObjectBlob   =   "chart.frx":0000
      TabIndex        =   9
      Top             =   4080
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "chart.frx":0064
      TabIndex        =   8
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Lihat Grafik"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "chart.frx":00E0
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MMM yyyy"
      Format          =   67305475
      CurrentDate     =   37623
   End
   Begin MSComCtl2.DTPicker DTPicker1 
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
      Left            =   4080
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "MMM yyyy"
      Format          =   67305475
      CurrentDate     =   37623
   End
End
Attribute VB_Name = "chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With chart
        MSChart1.chartType = VtChChartType2dBar
        MSChart1.AllowSelections = False
        MSChart1.ColumnCount = 1
        MSChart1.RowCount = 3
        
        MSChart1.Row = 1
        MSChart1.RowLabel = "Januari"
        MSChart1.Data = val(Text1.Text)
        MSChart1.Row = 2
        MSChart1.RowLabel = "Februari"
        MSChart1.Data = val(Text2.Text)
        MSChart1.Row = 3
        MSChart1.RowLabel = "Maret"
        MSChart1.Data = val(text3.Text)
        
       
    End With
End Sub

Private Sub Command2_Click()
End
End Sub


