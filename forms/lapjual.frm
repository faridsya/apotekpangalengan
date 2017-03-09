VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapjual 
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbmilik 
      Height          =   315
      Left            =   2640
      TabIndex        =   20
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox cmbsls 
      Height          =   315
      Left            =   2640
      TabIndex        =   18
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ComboBox cmbbrg 
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblbrg 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lapjual.frx":0000
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox text1 
      Height          =   315
      Left            =   2640
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox txtksr 
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox brg 
      Caption         =   "Barang"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   1080
      OleObjectBlob   =   "lapjual.frx":006A
      TabIndex        =   9
      Top             =   960
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   1080
      OleObjectBlob   =   "lapjual.frx":00E4
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "lapjual.frx":015A
      Top             =   4320
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapjual.frx":038E
      Left            =   3480
      List            =   "lapjual.frx":03B6
      TabIndex        =   7
      Text            =   "tahun"
      Top             =   1440
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapjual.frx":0402
      Left            =   2760
      List            =   "lapjual.frx":042A
      TabIndex        =   6
      Text            =   "bln"
      Top             =   1440
      Width           =   615
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5280
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Item"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1440
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
      Top             =   4080
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
      Format          =   164495363
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
      Format          =   164495363
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblksr 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapjual.frx":0455
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblplg 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapjual.frx":04D1
      TabIndex        =   12
      Top             =   1920
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel slsc 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lapjual.frx":0541
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lapjual.frx":05A9
      TabIndex        =   19
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "lapjual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idc, idi As String

Private Sub brg_Click()
If brg.Value = Checked Then
lblbrg.Visible = True
cmbbrg.Visible = True
lblplg.Visible = False
lblksr.Visible = False
Text1.Visible = False
txtksr.Visible = False
cmbsls.Visible = False
cmbmilik.Visible = False
slsc.Visible = False
Else
lblbrg.Visible = False
cmbbrg.Visible = False
lblplg.Visible = True
lblksr.Visible = True
Text1.Visible = True
txtksr.Visible = True
cmbsls.Visible = True
cmbmilik.Visible = True
slsc.Visible = True

End If
End Sub

Private Sub cmbsls_Click()
If cmbsls.Text <> "Semua" Then
Text1.Text = "Semua"
txtksr.Text = "Semua"
End If

End Sub

Private Sub Combo1_Change()
Option2.Value = True
End Sub
Private Sub Combo2_Change()
Option2.Value = True
End Sub

Private Sub cust()

Text1.Clear
Text1.AddItem "Semua"

sql = "select * from pelanggan order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
Text1.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  End If
Text1.Text = "Semua"

rsplg.Close

  End Sub
  Private Sub tbrg()

cmbbrg.Clear
cmbbrg.AddItem "Semua"

sql = "select deskripsi from tblbarang order by deskripsi"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbbrg.AddItem rsplg!deskripsi
rsplg.MoveNext
 Loop
  End If
cmbbrg.Text = "Semua"

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
On Error GoTo Pesan
Me.MousePointer = 11
If brg.Value = Checked Then
ctk2
Else
ctk1
End If
Me.MousePointer = 1

'Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If
End Sub
Sub ctk2()
If cmbbrg.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {tblbarang.deskripsi}='" & cmbbrg.Text & "'"
End If

With CrystalReport1
.Reset

  .ReportFileName = serperreport & "\brgjual.rpt"
  .RetrieveDataFiles

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




.Formulas(0) = "waktu='Periode: '+ '" & A & "'+'-'+'" & b & "'"

End If
.Formulas(2) = "tgjwb='" & tgjwb & "'"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
'Me.Hide

End With

End Sub
Private Sub sls()

cmbsls.Clear
cmbsls.AddItem "Semua"

sql = "select nama from pengguna order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbsls.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  End If
cmbsls.Text = "Semua"

rsplg.Close

  End Sub

Sub ctk1()
If Text1.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {pelanggan.nama}='" & Text1.Text & "'"
End If
If cmbsls.Text = "Semua" Then
tmbh2 = ""
Else
tmbh2 = " and {penjualan.kasir}='" & cmbsls.Text & "'"
End If
If txtksr.Text = "Semua" Then
tmbh3 = ""
Else
tmbh3 = " and {penjualan.jenis}='" & txtksr.Text & "'"
End If
If cmbmilik.Text = "semua" Then
tmbh4 = ""
Else
tmbh4 = " and {tblbarang.tipe}='" & cmbmilik.Text & "'"
End If
With CrystalReport1
.Reset

  .ReportFileName = serperreport & "\penjualan.rpt"
  .RetrieveDataFiles
  .Formulas(1) = "nama='" & nama_toko & "'"
    .Formulas(2) = "jns='" & txtksr.Text & "'"
.Formulas(3) = "sls='" & cmbsls.Text & "'"
  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{penjualan.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {penjualan.tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# " & tmbh & " " & tmbh2 & " " & tmbh3 & " " & tmbh4 & ""


  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode : '+'" & Format(DTPicker1.Value, "dd MMM yyyy") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM yyyy") & "'"
Else
.Formulas(0) = "waktu='Periode : '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If

Else

q = "1 / Combo1.Text / 2000"
A = MonthName(Combo1.Text)
b = Combo2.Text
.SelectionFormula = "month({penjualan.tanggal})=" & Combo1.Text & " and year({penjualan.tanggal})=" & Combo2.Text & " " & tmbh & " " & tmbh2 & " " & tmbh3 & " " & tmbh4 & ""


.Formulas(0) = "waktu='Periode: '+ '" & A & "'+'-'+'" & b & "'"

End If
.Formulas(2) = "tgjwb='" & tgjwb & "'"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
'Me.Hide

End With

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
tbrg
sls
cmbmilik.AddItem "semua"
cmbmilik.AddItem "sendiri"
cmbmilik.AddItem "titipan"
cmbmilik.ListIndex = 0

Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
txtksr.AddItem "Semua"
txtksr.AddItem "umum"
txtksr.AddItem "dokter"
txtksr.AddItem "resep"
txtksr.ListIndex = 0
Option1.Value = True
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

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
If Text1.Text <> "Semua" Then
cmbsls.Text = "Semua"
txtksr.Text = "Semua"

End If

If Text1.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select p.id_pelanggan,nama_sales from pelanggan p left join sales s on p.id_sales=s.id_sales where nama='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
idc = RS!id_pelanggan
cmbsls.Text = IIf(IsNull(RS!nama_sales) = True, "Semua", RS!nama_sales)
Else
MsgBox "Pelanggan tidak terdaftar"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub txtksr_Click()
If txtksr.Text <> "Semua" Then
cmbsls.Text = "Semua"
Text1.Text = "Semua"
End If
End Sub
