VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form cash 
   Caption         =   "CASHFLOW"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8115
   Icon            =   "cash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2520
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   13800
      TabIndex        =   28
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ce&tak"
      Height          =   255
      Left            =   8280
      TabIndex        =   27
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4680
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "cash.frx":324A
      Left            =   6240
      List            =   "cash.frx":3272
      TabIndex        =   2
      Text            =   "bln"
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "cash.frx":329D
      Left            =   6960
      List            =   "cash.frx":32C5
      TabIndex        =   1
      Text            =   "tahun"
      Top             =   480
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
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
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Tanggal "
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keterangan"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Debit(Rp)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Kredit(Rp)"
         Object.Width           =   2187
      EndProperty
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
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
      Left            =   840
      TabIndex        =   4
      Top             =   240
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
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   124911619
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
      Left            =   2520
      TabIndex        =   6
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   124911619
      CurrentDate     =   37623
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3135
      Left            =   8160
      TabIndex        =   10
      Top             =   1440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
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
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Tanggal "
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keterangan"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Debit(Rp)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Kredit(Rp)"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.Label Label11 
      Caption         =   "Bank"
      Height          =   255
      Left            =   12600
      TabIndex        =   29
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Total"
      Height          =   255
      Left            =   10680
      TabIndex        =   26
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label sal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9360
      TabIndex        =   23
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "-------------------------------------------------+"
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Total bank"
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label sal2 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9720
      TabIndex        =   20
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label bankk 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   14280
      TabIndex        =   19
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label bankm 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   12720
      TabIndex        =   18
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label sal1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Total kas"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label kask 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label kasm 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "BANK"
      Height          =   255
      Left            =   8160
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "KAS"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Dari tanggal :"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai tanggal :"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Periode"
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum1, sum2, sum3, sum4, sel1, sel2, d2, k2, k1, d1, d1d, k1d, k2d, d2d As Currency
Private Sub dbgrid()

On Error Resume Next
Dim sql As String
Dim l As ListItem
kosong
d1 = 0
      k1 = 0
      d1d = 0
      k1d = 0

Set RS = New Recordset
ListView1.ListItems.Clear

If Option1.Value = True Then
If databes = "Akses" Then
sql = "select * from keuangan where tanggal between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & "#  and  #" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#  order by tanggal"
Else
sql = "select * from keuangan where tanggal between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "'  and  '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "'  order by tanggal"

End If
Else
sql = "select * from keuangan where (month(keuangan.tanggal)='" & Combo1.Text & "' and year(keuangan.tanggal)='" & Combo2.Text & "') order by tanggal"
End If
Set RS = jual.Execute(sql)

If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
         


        Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
        l.SubItems(1) = Format(![tanggal], "dd MMM yyyy")
        l.SubItems(2) = ![keterangan]
              l.SubItems(3) = ![pemasukan]
              l.SubItems(4) = ![pengeluaran]

d1 = d1 + val(l.SubItems(3))
k1 = k1 + val(l.SubItems(4))
              l.SubItems(3) = Format(![pemasukan], "#,#0.#0")
              l.SubItems(4) = Format(![pengeluaran], "#,#0.#0")
   
    .MoveNext
    Loop
End With

kask.Caption = Format(k1, "#,#0.#0")
kasm.Caption = Format(d1, "#,#0.#0")

sal1.Caption = Format(d1 - k1, "#,#0.#0")

ListView1.Refresh

End Sub

Private Sub Dbgrid2()
On Error Resume Next
Dim l2 As ListItem
ListView2.ListItems.Clear
If Option1.Value = True Then
If Combo3.Text = "Semua" Or idb = "" Then
If databes = "Akses" Then
sql = "select * from keuangan2 where tanggal2 between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " # order by tanggal2 "
Else

sql = "select * from keuangan2 where tanggal2 between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ' order by tanggal2 "

End If

Else
If databes = "Akses" Then

sql = "select * from keuangan2 where tanggal2 between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " # and kode_bank='" & idb & "'  order by tanggal2 "
Else
sql = "select * from keuangan2 where tanggal2 between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ' and kode_bank='" & idb & "'  order by tanggal2 "

End If
End If
Else
If Combo3.Text = "Semua" Or idb = "" Then


sql = "select * from keuangan2 where (month(keuangan2.tanggal2)='" & Combo1.Text & "' and year(keuangan2.tanggal2)='" & Combo2.Text & "') order by tanggal2"
Else
sql = "select * from keuangan2 where (month(keuangan2.tanggal2)='" & Combo1.Text & "' and year(keuangan2.tanggal2)='" & Combo2.Text & "')and kode_bank='" & idb & "' order by tanggal2"

End If
End If

Set RS = jual.Execute(sql)

If RS.RecordCount = 0 Then Exit Sub
d2 = 0
k2 = 0
k2d = 0
d2d = 0
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l2 = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l2.SubItems(1) = Format(![Tanggal2], "dd MMM yyyy")
        l2.SubItems(2) = ![keterangan2]
                l2.SubItems(3) = ![pemasukan2]
                l2.SubItems(4) = ![pengeluaran2]
                               l2.SubItems(5) = ![pemasukan2d]
                l2.SubItems(6) = ![pengeluaran2d]
    
                  
                  d2 = d2 + val(l2.SubItems(3))
                  k2 = k2 + val(l2.SubItems(4))
                
                l2.SubItems(3) = Format(![pemasukan2], "#,#0.#0")
                l2.SubItems(4) = Format(![pengeluaran2], "#,#0.#0")

    .MoveNext
    Loop
End With
bankk.Caption = Format(k2, "#,#0.#0")
bankm.Caption = Format(d2, "#,#0.#0")

sal2.Caption = Format(d2 - k2, "#,#0.#0")

ListView2.Refresh

End Sub
Sub ttl()

sal.Caption = "Rp" & Format(val(d1) + val(d2) - val(k1) - val(k2), "#,#0.#0")

End Sub

Private Sub Combo1_Change()
dbgrid
Dbgrid2
ttl


End Sub

Private Sub Combo1_Click()
dbgrid
Dbgrid2
ttl

End Sub

Private Sub Combo1_GotFocus()
Option2.Value = True
End Sub

Private Sub Combo2_Change()
dbgrid
Dbgrid2
ttl

End Sub

Private Sub Combo2_Click()
dbgrid
Dbgrid2

ttl

End Sub

Private Sub Combo2_GotFocus()
Option2.Value = True

End Sub

Private Sub Combo3_Change()
dbgrid
Dbgrid2
ttl

End Sub

Private Sub Combo3_Click()
bankk.Caption = ""
bankm.Caption = ""
If Combo3.Text = "Semua" Then
idb = ""
dbgrid
Dbgrid2
ttl

Exit Sub
End If
Set rsbank = New Recordset
rsbank.Open "Select* from bank where nama_bank='" & Combo3.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rsbank.EOF Then
idb = rsbank!kode_bank
rsbank.Close
Else
MsgBox "Bank tidak terdaftar"
End If
dbgrid
Dbgrid2
ttl

End Sub

Private Sub Command1_Click()
If ListView1.ListItems.count = 0 Then
    MsgBox "Tidak ada data yang akan dcetak !", vbCritical, judul
    Exit Sub
End If
If MsgBox("Apakah data akan dicetak ?", vbQuestion + vbYesNo, judul) = vbNo Then Exit Sub
cetakdataa

End Sub
Sub cetakdataa()
On Error GoTo Pesan

With CrystalReport1
.Reset

  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\labarugi.rpt"
  .RetrieveDataFiles
.Formulas(1) = "nama='" & nama_toko & "'"

  .WindowTitle = "laporan"
  If Option1.Value = True Then
  .SelectionFormula = "{keuangan.tanggal}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {keuangan.Tanggal}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#"
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode: '+ '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "'+'-'+'" & Format(DTPicker2.Value, "YYYY-mm-dd") & "'"
      Else
      .Formulas(0) = "waktu='Periode: '+ '" & Format(DTPicker1.Value, "dd MMM yyyy") & "'"
End If
Else
q = "1 / Combo1.Text / 2000"
A = MonthName(Combo1.Text)
b = Combo2.Text
.SelectionFormula = "month({keuangan.tanggal})=" & Combo1.Text & " and year({keuangan.tanggal})=" & Combo2.Text & ""
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
End With
'Me.Hide

Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener", , judul
End If

End Sub
Public Sub CetakData()
On Error Resume Next
Printer.Orientation = vbPRORPortrait
Printer.PaperSize = vbPRPSLegal
Printer.PrintQuality = vbPRPQDraft
Printer.Font = "Courier New"

Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print
Printer.Print
'cetak head
Printer.FontSize = 12

Printer.Print Tab(34); "LAPORAN ARUS KAS "
Printer.Print Tab(34); ""
Printer.FontSize = 10
Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(5); "Periode";
Printer.Print Tab(22); ":";
If Option1.Value = True Then
If Not Format(DTPicker1.Value, "dd MMM yyyy") = Format(DTPicker2.Value, "dd MMM yyyy") Then
Printer.Print Tab(24); Format(DTPicker1.Value, "dd MMM yyyy"); "-"; Format(DTPicker1.Value, "dd MMM yyyy");
Else

Printer.Print Tab(24); Format(DTPicker1.Value, "dd MMM yyyy");
End If
Else
If Option2.Value = True Then
b = Format("" & Combo1.Text & " " & Combo2.Text & "")
c = Format(b, "mmmm yyyy")

Printer.Print Tab(29); c;
Else
Printer.Print Tab(29); Combo3.Text
End If
End If
Printer.Print ; ""
Printer.Print ; ""
 Printer.Print Tab(5); "Total debet kas(Rp)";
Printer.Print Tab(26); ":";
Printer.Print Tab(27); RKanan(d1, "###,###,##0.#0");

Printer.Print Tab(5); "Total kredit kas(Rp)";
Printer.Print Tab(26); ":";
Printer.Print Tab(27); RKanan(k1, "###,###,##0.#0");







    Printer.Print ; " "
    
    Printer.Print ; " "

Printer.Print Tab(5); "KAS";
Printer.FontSize = 9

Printer.Print Tab(5); String$(95, "=");

'Printer.Print Tab(90); "BANK";
Printer.Print Tab(5); "No.";
Printer.Print Tab(9); "Tanggal";
Printer.Print Tab(18); "Keterangan";
Printer.Print Tab(62); "Debet(Rp)";
Printer.Print Tab(81); "Kredit(Rp)";

Printer.Print Tab(5); String$(95, "=");
t = ListView1.ListItems.count
For I = 1 To t
'cetak isi
    Printer.Print Tab(5); ListView1.ListItems(I).Text;
    Printer.Print Tab(9); Format(ListView1.ListItems(I).SubItems(1), "dd/mm/yy");
    Printer.Print Tab(18); ListView1.ListItems(I).SubItems(2);
    Printer.Print Tab(57); RKanan(ListView1.ListItems(I).SubItems(3), "###,###,##0.#0");
    Printer.Print Tab(76); RKanan(ListView1.ListItems(I).SubItems(4), "###,###,##0.#0");

Next I
For I = 1 To ListView1.ListItems.count
'cetak isi

Next I

'Printer.Print Tab(10); String$(85, "=")

'keluarkan kertas
Printer.NewPage

Printer.EndDoc
End Sub
Public Sub CetakData2()
On Error Resume Next
Printer.Orientation = vbPRORPortrait
Printer.PaperSize = vbPRPSLegal
Printer.PrintQuality = vbPRPQDraft
Printer.Font = "Courier New"

Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print
Printer.Print
'cetak head
Printer.FontSize = 12

Printer.Print Tab(34); "LAPORAN ARUS BANK "
Printer.Print Tab(34); ""

Printer.FontSize = 10
Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(5); "Periode";
Printer.Print Tab(22); ":";
If Option1.Value = True Then
If Not Format(DTPicker1.Value, "dd MMM yyyy") = Format(DTPicker2.Value, "dd MMM yyyy") Then
Printer.Print Tab(24); Format(DTPicker1.Value, "dd MMM yyyy"); "-"; Format(DTPicker1.Value, "dd MMM yyyy");
Else

Printer.Print Tab(24); Format(DTPicker1.Value, "dd MMM yyyy");
End If
Else
If Option2.Value = True Then
b = Format("" & Combo1.Text & " " & Combo2.Text & "")
c = Format(b, "mmmm yyyy")

Printer.Print Tab(29); c;
Else
Printer.Print Tab(29); Combo3.Text
End If
End If
Printer.Print ; ""
Printer.Print ; ""
 Printer.Print Tab(5); "Total debet bank(Rp)";
Printer.Print Tab(26); ":";
Printer.Print Tab(29); RKanan(d2, "###,###,##0.#0");

Printer.Print Tab(5); "Total kredit bank(Rp)";
Printer.Print Tab(26); ":";
Printer.Print Tab(29); RKanan(k2, "###,###,##0.#0");

Printer.Print Tab(55); "Bank";
Printer.Print Tab(60); ":";
Printer.Print Tab(62); Combo3.Text;






    Printer.Print ; " "
    
    Printer.Print ; " "

Printer.Print Tab(5); "BANK";
Printer.FontSize = 9

Printer.Print Tab(5); String$(103, "=");

'Printer.Print Tab(90); "BANK";
Printer.Print Tab(5); "No.";
Printer.Print Tab(9); "Tanggal";
Printer.Print Tab(18); "Keterangan";
Printer.Print Tab(62); "Debet(Rp)";
Printer.Print Tab(81); "Kredit(Rp)";

Printer.Print Tab(5); String$(103, "=");
t = ListView2.ListItems.count
For I = 1 To t
'cetak isi
    Printer.Print Tab(5); ListView2.ListItems(I).Text;
    Printer.Print Tab(9); Format(ListView2.ListItems(I).SubItems(1), "dd/mm/yy");
    Printer.Print Tab(18); ListView2.ListItems(I).SubItems(2);
    Printer.Print Tab(57); RKanan(ListView2.ListItems(I).SubItems(3), "###,###,##0.#0");
    Printer.Print Tab(76); RKanan(ListView2.ListItems(I).SubItems(4), "###,###,##0.#0");

Next I
For I = 1 To ListView2.ListItems.count
'cetak isi

Next I

'Printer.Print Tab(10); String$(85, "=")

'keluarkan kertas
Printer.NewPage

Printer.EndDoc
End Sub


Private Sub Command2_Click()
dbgrid
Dbgrid2
ttl
ListView1.Refresh
End Sub

Private Sub Command3_Click()
If ListView2.ListItems.count = 0 Then
    MsgBox "Tidak ada data yang akan dcetak !", vbCritical, judul
    Exit Sub
End If
If MsgBox("Apakah data akan dicetak ?", vbQuestion + vbYesNo, judul) = vbNo Then Exit Sub
CetakData2

End Sub

Private Sub DTPicker1_Change()
dbgrid
Dbgrid2
ttl

End Sub

Private Sub DTPicker2_Change()
dbgrid
Dbgrid2
ttl




End Sub

Private Sub Form_Load()
Ketengah Me
Combo1.Text = Format(Now, "m")
Combo2.Text = Format(Now, "yyyy")
DTPicker1.Value = Now
DTPicker2.Value = Now
Option1.Value = True
dbgrid
Dbgrid2
ttl
idb = ""
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Option1_Change()
dbgrid
Dbgrid2
ttl

End Sub

Private Sub Option2_Change()
dbgrid
Dbgrid2
ttl

End Sub
Private Sub kosong()
d1 = 0
      k1 = 0
      d1d = 0
      k1d = 0
d2 = 0
k2 = 0
k2d = 0
d2d = 0

sal1.Caption = ""
sal2.Caption = ""
sal.Caption = ""
kasm.Caption = ""
kask.Caption = ""
kask.Caption = ""
bankm.Caption = ""
bankk.Caption = ""
End Sub
