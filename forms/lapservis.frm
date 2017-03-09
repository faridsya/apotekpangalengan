VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapservis 
   Caption         =   "Laporan Servis"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   ClipControls    =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "lapservis.frx":0000
      TabIndex        =   15
      Top             =   3840
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "lapservis.frx":006A
      Top             =   4800
   End
   Begin VB.ComboBox cmbstts2 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ComboBox cmbteknisi 
      Height          =   315
      Left            =   2400
      TabIndex        =   13
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1920
      TabIndex        =   10
      Top             =   0
      Width           =   2775
      Begin VB.OptionButton Option3 
         Caption         =   "Tanggal Masuk"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tanggal diambil"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2760
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
      TabIndex        =   5
      Top             =   1200
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Text            =   "bln"
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Text            =   "tahun"
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   4800
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   108527619
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
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   108527619
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lapservis.frx":029E
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapservis.frx":0318
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "lapservis.frx":038E
      TabIndex        =   14
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "lapservis"
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
On Error GoTo Pesan
If cmbteknisi.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {teknisi.nama_teknisi}='" & cmbteknisi.Text & "'"
End If
If cmbstts2.Text = "Semua" Then
tmbh2 = ""
Else
tmbh2 = " and {tservis.status_servis}='" & cmbstts2.Text & "'"
End If
If Option3.Value = True Then
jns = "tgl_msk"
jns2 = " tgl masuk :"
Else
jns = "tgl_out"
jns2 = " tgl diambil: "
End If
With CrystalReport1
  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\servis.rpt"
  .RetrieveDataFiles
.Formulas(1) = "stts='" & cmbstts2.Text & "'"
.Formulas(2) = "teknisi='" & cmbteknisi.Text & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{tservis." & jns & "}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {tservis." & jns & "}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# " & tmbh & " " & tmbh2 & ""
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode'+'" & jns2 & "'+'" & Format(DTPicker1.Value, "dd MMM yyyy") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM yyyy") & "'"
Else
.Formulas(0) = "waktu='Periode'+'" & jns2 & "'+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If
Else
q = "1 / Combo1.Text / 2000"
A = MonthName(Combo1.Text)
b = Combo2.Text
.SelectionFormula = "month({tservis." & jns & "})=" & Combo1.Text & " and year({tservis." & jns & "})=" & Combo2.Text & " " & tmbh & " " & tmbh2 & ""
.Formulas(0) = "waktu='Periode'+'" & jns2 & "'+'" & A & "'+'-'+'" & b & "'"

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
Private Sub listeknisi()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbteknisi.Clear
cmbteknisi.AddItem "Semua"
sql = "select nama_teknisi from teknisi order by nama_teknisi"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbteknisi.AddItem rsbarang!nama_teknisi
rsbarang.MoveNext
 Loop
  End If
  cmbteknisi.Text = "Semua"
rsbarang.Close

  End Sub

Private Sub Form_Load()
Ketengah Me
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")
Option1.Value = True
listeknisi
cmbstts2.AddItem "Semua"
cmbstts2.AddItem "daftar"
cmbstts2.AddItem "tunggu konfirmasi"
cmbstts2.AddItem "setuju"
cmbstts2.AddItem "perbaikan"
cmbstts2.AddItem "selesai"
cmbstts2.AddItem "diambil"
cmbstts2.AddItem "batal"

cmbstts2.Text = "Semua"
Option3.Value = True
Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub

