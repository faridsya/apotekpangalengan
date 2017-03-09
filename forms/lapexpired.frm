VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapexpired 
   Caption         =   "Laporan Expired"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1440
      OleObjectBlob   =   "lapexpired.frx":0000
      Top             =   1200
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "lapexpired.frx":0234
      TabIndex        =   0
      Top             =   720
      Width           =   2055
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
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   125632515
      CurrentDate     =   37623
   End
End
Attribute VB_Name = "lapexpired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo Pesan

With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\expired.rpt"
  .RetrieveDataFiles
'.Formulas(1) = "nama='" & nama_toko & "'"

  .WindowTitle = "laporan"
  .SelectionFormula = "{tblbarang.tgl_expire}<#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "#"
.Formulas(0) = "waktu='Expired < '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
.Formulas(2) = "tgjwb='" & tgjwb & "'"

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
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If

End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
