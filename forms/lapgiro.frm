VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapgiro 
   Caption         =   "Laporan arus giro"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5940
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5280
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Item"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox text1 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
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
      TabIndex        =   7
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapgiro.frx":0000
      Left            =   2760
      List            =   "lapgiro.frx":0028
      TabIndex        =   2
      Text            =   "bln"
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapgiro.frx":0053
      Left            =   3480
      List            =   "lapgiro.frx":007B
      TabIndex        =   1
      Text            =   "tahun"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   3240
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
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   116719619
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
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   116719619
      CurrentDate     =   37623
   End
   Begin VB.Label Label5 
      Caption         =   "Bank"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Dari tanggal :"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai tanggal :"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "lapgiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idc, idi As String

Private Sub Combo1_Change()
Option2.Value = True
End Sub
Private Sub Combo2_Change()
Option2.Value = True
End Sub

Private Sub cust()

Text1.Clear
Text1.AddItem "Semua"

sql = "select * from bank order by nama_bank"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
Text1.AddItem rsplg!nama_bank
rsplg.MoveNext
 Loop
  End If
Text1.Text = "Semua"

rsplg.Close

  End Sub
Private Sub itm()

Item.Clear
Set rsplg = New Recordset
sql = "select * from tblbarang order by deskripsi"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
Item.AddItem rsplg!deskripsi
rsplg.MoveNext
 Loop
  End If

rsplg.Close

  End Sub

Private Sub Command1_Click()
'On Error GoTo Pesan
With CrystalReport1
  .Password = Chr(10) & "tujuh"
    If Not Check1.Value = Checked Then
If (Option1.Value = True) Or Option2.Value = True Then
  .ReportFileName = serperreport & "\giro.rpt"
  .RetrieveDataFiles
  .Formulas(1) = "nama='" & nama_toko & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then

   If Text1.Text = "Semua" Then
  .SelectionFormula = "{giro.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {giro.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
Else
  .SelectionFormula = "{giro.kode_bank}='" & idc & "' and {giro.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {giro.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"

End If
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode : '+'" & Format(DTPicker1.Value, "dd MMM yyyy") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM yyyy") & "'"
Else
.Formulas(0) = "waktu='Periode : '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If

Else

q = "1 / Combo1.Text / 2000"
a = MonthName(Combo1.Text)
b = Combo2.Text
   If Text1.Text = "Semua" Then

.SelectionFormula = "month({giro.tanggal})=" & Combo1.Text & " and year({giro.tanggal})=" & Combo2.Text & ""
Else
.SelectionFormula = "{giro.kode_bank}='" & idc & "' and month({giro.tanggal})=" & Combo1.Text & " and year({giro.tanggal})=" & Combo2.Text & ""
End If


.Formulas(0) = "waktu='Periode: '+ '" & a & "'+'-'+'" & b & "'"

End If
End If
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
Me.Hide
Else
 .ReportFileName = serperreport & "\detil_jual.rpt"
  .RetrieveDataFiles

  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{penjualan.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {penjualan.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
.Formulas(0) = "waktu='Periode: '+ '" & Format(DTPicker1.Value, "dd MMMM YYYY") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMMM YYYY") & "'"
Else
q = "1 / Combo1.Text / 2000"
a = MonthName(Combo1.Text)
b = Combo2.Text
.SelectionFormula = "month({penjualan.tanggal})=" & Combo1.Text & " and year({penjualan.tanggal})=" & Combo2.Text & ""

.Formulas(0) = "waktu='Periode: '+ '" & a & "'+'-'+'" & b & "'"

End If
        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
  End If

End With
Me.Hide
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

Private Sub Form_Load()
Ketengah Me
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")
cust
Option1.Value = True

'text1_Click
End Sub


Private Sub item_Click()
If Item.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select* from tblbarang where deskripsi='" & Item.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
idi = RS!kode_brg
MsgBox idi
Else
MsgBox "Barang tidak terdaftar"
End If

End Sub

Private Sub text1_Click()
If Text1.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select* from bank where nama_bank='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
idc = RS!kode_bank
Else
MsgBox "Bank tidak terdaftar"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub
