VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapkomisi 
   Caption         =   "Laporan Komisi"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   ClipControls    =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbdok 
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapkomisi.frx":0000
      TabIndex        =   8
      Top             =   960
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapkomisi.frx":007A
      Left            =   2640
      List            =   "lapkomisi.frx":00A2
      TabIndex        =   7
      Text            =   "bln"
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapkomisi.frx":00CD
      Left            =   3360
      List            =   "lapkomisi.frx":00F5
      TabIndex        =   6
      Text            =   "tahun"
      Top             =   1680
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4320
      OleObjectBlob   =   "lapkomisi.frx":0141
      Top             =   2040
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lapkomisi.frx":0375
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "Per bulan"
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
   Begin XPControls.XPOption Option1 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "Per tanggal"
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
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
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   117309443
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
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   117309443
      CurrentDate     =   37623
   End
End
Attribute VB_Name = "lapkomisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Option2.Value = True
End Sub
Private Sub Combo2_Change()
Option2.Value = True
End Sub

Private Sub Command1_Click()
'On Error GoTo Pesan
If cmbdok.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {dokter.nama}='" & cmbdok.Text & "'"
End If

With CrystalReport1
  .Reset
  .ReportFileName = serperreport & "\lapkomisi.rpt"
  .RetrieveDataFiles
.Formulas(1) = "dok='" & cmbdok.Text & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{penjualan.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {penjualan.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# " & tmbh & ""
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode : '+'" & Format(DTPicker1.Value, "dd MMM yyyy") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM yyyy") & "'"
Else
.Formulas(0) = "waktu='Periode : '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If
Else
q = "1 / Combo1.Text / 2000"
A = MonthName(Combo1.Text)
b = Combo2.Text
.SelectionFormula = "month({penjualan.tanggal})=" & Combo1.Text & " and year({penjualan.tanggal})=" & Combo2.Text & " " & tmbh & ""
.Formulas(0) = "waktu='Periode: '+'" & A & "'+'-'+'" & b & "'"

End If
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

Private Sub DTPicker1_Change()
Option1.Value = True
End Sub
Private Sub DTPicker2_Change()
Option1.Value = True
End Sub
Private Sub lisdok()

cmbdok.Clear
cmbdok.AddItem "Semua"

sql = "select * from dokter order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbdok.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  End If
cmbdok.Text = "Semua"

rsplg.Close

  End Sub

Private Sub Form_Load()
Ketengah Me
lisdok
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")
Option1.Value = True
Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub

