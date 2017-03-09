VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kasir 
   Caption         =   "LAPORAN PENERIMAAN KASIR"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   Icon            =   "kasir.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtksr 
      Height          =   315
      Left            =   2280
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "proses"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin XPControls.XPOption Option1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
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
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   126418947
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
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   126418947
      CurrentDate     =   37623
   End
   Begin VB.Label Label4 
      Caption         =   "Kasir"
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "-"
      Height          =   375
      Left            =   8640
      TabIndex        =   13
      Top             =   3000
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   8520
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label ttldiskonc 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label tljual2 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "Total diskon"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Total penjualan"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Total Uang masuk"
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label ttlc 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai tanggal :"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Dari tanggal :"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "kasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttlm, ttlm2, ttlk, ttlk2, sum, sum2, tljual, tljuall, tljualll, tljuald, ttldiskon, ttldiskon2, pm, pmd, trj, ttlwrg As Currency
Sub itung1()
Set RS = New Recordset


Set RS = New Recordset

sql = "select sum(detiljual.total_retur) as pem from detiljual,penjualan where (penjualan.tanggal  between # " & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " #) and penjualan.no_penjualan=detiljual.no_penjualan "
Set RS = jual.Execute(sql)

trj = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset

sql = "select sum(pemasukan) as pem from keuangan where jenis='Lain-lain' and (Tanggal between # " & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " #)"
Set RS = jual.Execute(sql)

pm1 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close
Set RS = New Recordset

sql = "select sum(pemasukan2) as pem from keuangan2 where jenis2='Lain-lain' and (Tanggal2 between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " #)"
Set RS = jual.Execute(sql)

pm2 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close
pm = val(pm1) + val(pm2)
Set RS = New Recordset

sql = "select sum(pemasukan) as pem from keuangan where jenis='Lain-lain' and (Tanggal between # " & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " #)"
Set RS = jual.Execute(sql)

pm1d = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

'kepake
Set RS = New Recordset
sql = "select sum(detiljual.total) as pem from detiljual,penjualan,tblbarang where penjualan.no_penjualan=detiljual.no_penjualan and detiljual.kode_brg=tblbarang.kode_brg and tanggal  between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# and tblbarang.jenis='Umum'"
Set RS = jual.Execute(sql)
tljual = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select sum(total_diskon) as pem from penjualan where tanggal  between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & " # and #" & Format(DTPicker2.Value, "YYYY-mm-dd") & " # "
Set RS = jual.Execute(sql)

ttldiskon = IIf(IsNull(RS!pem) = True, 0, RS!pem)
ttlm = tljual
RS.Close

Set RS = New Recordset
sql = "select sum(detiljual.harga_beli*(detiljual.jumlah_brg-(detiljual.kembali_uang+detiljual.kembali_uang2))) as pem from detiljual,penjualan where penjualan.tanggal  between #" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# and  #" & Format(DTPicker2.Value, "YYYY-mm-dd") & "#  and penjualan.no_penjualan=detiljual.no_penjualan"
Set RS = jual.Execute(sql)

ttlk = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

End Sub
Sub itung2()



Set RS = New Recordset
sql = "select sum(detiljual.total) as pem from detiljual,penjualan,tblbarang where detiljual.kode_brg=tblbarang.kode_brg and penjualan.no_penjualan=detiljual.no_penjualan  and tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "'  and penjualan.kasir='" & txtksr.Text & "'"
Set RS = jual.Execute(sql)
tljual = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close


Set RS = New Recordset
sql = "select sum(total_diskon) as pem from penjualan where tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "' and penjualan.kasir='" & txtksr.Text & "'"
Set RS = jual.Execute(sql)

ttldiskon = IIf(IsNull(RS!pem) = True, 0, RS!pem)

ttlm = tljual + -ttldiskon

End Sub
Public Sub CetakData()
On Error Resume Next
'vbPRORPortrait  = 1
'vbPRORLandscape = 2
Printer.Orientation = vbPRORPortrait
Printer.PaperSize = vbPRPSLegal
Printer.PrintQuality = vbPRPQDraft
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print
Printer.Print
'cetak head
Printer.FontSize = 18
Printer.Print Tab(14); "Laporan penerimaan kasir " + txtksr.Text
Printer.FontSize = 10
Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(10); "Periode";
Printer.Print Tab(25); ":";
If Option1.Value = True Then
If Not Format(DTPicker1.Value, "dd MMM yyyy") = Format(DTPicker2.Value, "dd MMM yyyy") Then
Printer.Print Tab(27); Format(DTPicker1.Value, "dd MMM yyyy"); "-"; Format(DTPicker2.Value, "dd MMM yyyy");
Else

Printer.Print Tab(27); Format(DTPicker1.Value, "dd MMM yyyy");
End If
Else
If Option2.Value = True Then
b = Format("" & Combo1.Text & " " & Combo2.Text & "")
c = Format(b, "mmmm yyyy")

Printer.Print Tab(27); c;
Else
Printer.Print Tab(27); Combo3.Text
End If


End If
Printer.Print ; ""
Printer.Print ; ""

Printer.Print Tab(10); String$(85, "=");

Printer.Print ; ""
Printer.Print ; ""
Printer.Font = "Courier New"

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""


Printer.Print Tab(10); "TOTAL PENERIMAAN PENJUALAN";
Printer.Print Tab(50); RKanan(val(tljual), "###,###,##0.#0");



Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""



Printer.Print Tab(10); "TOTAL DISKON TAMBAHAN";
Printer.Print Tab(50); RKanan(val(ttldiskon), "###,###,##0.#0");
Printer.Print Tab(46); "------------------ -";

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""

Printer.Print Tab(10); "TOTAL UANG MASUK";
Printer.Print Tab(50); RKanan(val(ttlm), "###,###,##0.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""





Printer.Print Tab(10); String$(85, "=")

'keluarkan kertas
Printer.NewPage

Printer.EndDoc
End Sub

Private Sub Combo1_GotFocus()
Option2.Value = True
End Sub

Private Sub Combo3_GotFocus()
Option3.Value = True
End Sub

Private Sub Command1_Click()
If MsgBox("Apakah data akan dicetak ?", vbQuestion + vbYesNo, judul) = vbNo Then Exit Sub
CetakData

End Sub

Private Sub Combo2_GotFocus()
Option2.Value = True

End Sub

Sub Command2_Click()
On Error Resume Next
trjc.Caption = ""
ttldiskonc.Caption = ""
ttl_jual.Caption = ""

ttl_beli.Caption = ""

margin1.Caption = ""

margin2.Caption = ""
lain.Caption = ""
trj = 0
ttlm = 0
ttlm2 = 0
ttlk = 0
ttlk2 = 0

tljual = 0
pm = 0
ttldiskon = 0
ttldiskon2 = 0

sum = 0

If databes = "Akses" Then
itung1
Else
itung2
End If

tljual2.Caption = Format(val(tljual), "#,#0.#0")

ttldiskonc.Caption = Format(val(ttldiskon), "#,#0.#0")

ttlc.Caption = Format(ttlm, "#,#0.#0")




Exit Sub
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
End If

End Sub

Private Sub DTPicker1_GotFocus()
Option1.Value = True

End Sub


Private Sub Form_Load()
 DTPicker1.Value = Now
  DTPicker2.Value = Now
Option1.Value = True
ksr
End Sub

Private Sub ksr()
On Error Resume Next
txtksr.Clear

sql = "select * from penjualan where kasir <>'' order by kasir"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
txtksr.AddItem rsplg!kasir
rsplg.MoveNext
 Loop
  End If
txtksr.Text = transaksi.kasir.Caption

    With txtksr
    For I = 0 To .ListCount - 1
      For j = .ListCount To (I + 1) Step -1
         If .List(j) = .List(I) Then
         .RemoveItem j
         End If
      Next j
    Next I
  End With


  End Sub

Private Sub txtksr_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub
