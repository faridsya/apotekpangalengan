VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapbyru 
   Caption         =   "Laporan pembayaran hutang"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   ClipControls    =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel Supplier 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "lapbyru.frx":0000
      TabIndex        =   12
      Top             =   2520
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "lapbyru.frx":006E
      Top             =   3120
   End
   Begin VB.ComboBox text1 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapbyru.frx":02A2
      Left            =   3000
      List            =   "lapbyru.frx":02CA
      TabIndex        =   6
      Text            =   "tahun"
      Top             =   1680
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapbyru.frx":0316
      Left            =   2280
      List            =   "lapbyru.frx":033E
      TabIndex        =   5
      Text            =   "bln"
      Top             =   1680
      Width           =   615
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
      Left            =   2640
      TabIndex        =   0
      Top             =   3120
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
      Format          =   123994115
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
      Format          =   123994115
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lapbyru.frx":0369
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapbyru.frx":03E3
      TabIndex        =   11
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Bayar dari bank"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "lapbyru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ids As String

Private Sub Combo1_Change()
Option2.Value = True
End Sub
Private Sub sup()

Text1.Clear
Text1.AddItem "Semua"

sql = "select * from tblsupplier order by supplier"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
Text1.AddItem rsplg!Supplier
rsplg.MoveNext
 Loop
  End If
Text1.Text = "Semua"

rsplg.Close

  End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub text1_Click()
If Text1.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select* from tblsupplier where supplier='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
ids = RS!id_supplier
Else
MsgBox "Pelanggan tidak terdaftar"
End If
End Sub


Private Sub Combo1_GotFocus()
Option2.Value = True

End Sub

Private Sub Combo2_Change()
Option2.Value = True
End Sub

Private Sub Combo2_GotFocus()
Option2.Value = True

End Sub

Private Sub Command1_Click()
On Error GoTo Pesan

With CrystalReport1
  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\bayar_hutang.rpt"
  .RetrieveDataFiles
.Formulas(2) = "nama='" & nama_toko & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then
   If Combo3.Text <> "Semua" And Text1.Text = "Semua" Then

  .SelectionFormula = "{byr_hutang.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {byr_hutang.Tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
Else
 If Combo3.Text = "Semua" And Text1.Text = "Semua" Then
  
  .SelectionFormula = "{byr_hutang.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {byr_hutang.Tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
  Else
  If Combo3.Text <> "Semua" And Text1.Text <> "Semua" Then
    .SelectionFormula = "{byr_hutang.id_supplier}='" & ids & "'  and {byr_hutang.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {byr_hutang.Tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
Else
  If Combo3.Text = "Semua" And Text1.Text <> "Semua" Then
  
      .SelectionFormula = "{byr_hutang.id_supplier}='" & ids & "' and {byr_hutang.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {byr_hutang.Tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"

End If
End If
  End If
  End If
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode: '+ '" & Format(DTPicker1.Value, "dd MMMM YYYY") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMMM YYYY") & "'"
      Else
      .Formulas(0) = "waktu='Periode: '+ '" & Format(DTPicker1.Value, "dd MMMM YYYY") & "'"
End If
Else
q = "1 / Combo1.Text / 2000"
A = MonthName(Combo1.Text)
b = Combo2.Text
    If Combo3.Text <> "Semua" And Text1.Text = "Semua" Then
     
.SelectionFormula = "month({byr_hutang.tanggal})=" & Combo1.Text & " and year({byr_hutang.tanggal})=" & Combo2.Text & ""
Else
 If Combo3.Text = "Semua" And Text1.Text = "Semua" Then
.SelectionFormula = " month({byr_hutang.tanggal})=" & Combo1.Text & " and year({byr_hutang.tanggal})=" & Combo2.Text & ""
Else
  If Combo3.Text <> "Semua" And Text1.Text <> "Semua" Then
  .SelectionFormula = "{byr_hutang.id_supplier}='" & ids & "' and month({byr_hutang.tanggal})=" & Combo1.Text & " and year({byr_hutang.tanggal})=" & Combo2.Text & ""
Else
  If Combo3.Text = "Semua" And Text1.Text <> "Semua" Then

.SelectionFormula = "{byr_hutang.id_supplier}='" & ids & "' and month({byr_hutang.tanggal})=" & Combo1.Text & " and year({byr_hutang.tanggal})=" & Combo2.Text & ""
End If
End If
End If
End If
.Formulas(0) = "waktu='Periode: '+ '" & A & "'+'-'+'" & b & "'"

End If
.Formulas(1) = "bank='" & Combo3.Text & "'"
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

Private Sub DTPicker1_Change()
Option1.Value = True
End Sub

Private Sub DTPicker1_GotFocus()
Option1.Value = True
End Sub

Private Sub DTPicker2_Change()
Option1.Value = True
End Sub

Private Sub DTPicker2_GotFocus()
Option1.Value = True

End Sub

Private Sub Form_Load()
Ketengah Me
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")
sup
Option1.Value = True
Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

