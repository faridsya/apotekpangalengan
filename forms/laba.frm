VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form laba 
   BackColor       =   &H00FFFFC0&
   Caption         =   "LABA RUGI"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "laba.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Catatan"
      Height          =   375
      Left            =   10080
      TabIndex        =   40
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "laba.frx":324A
      Left            =   2280
      List            =   "laba.frx":3272
      TabIndex        =   24
      Text            =   "tahun"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "proses"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   6240
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "laba.frx":32BE
      Left            =   3240
      List            =   "laba.frx":32E6
      TabIndex        =   5
      Text            =   "tahun"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "laba.frx":3332
      Left            =   2400
      List            =   "laba.frx":335A
      TabIndex        =   4
      Text            =   "bln"
      Top             =   1320
      Width           =   615
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
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
      Format          =   128385027
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
      Format          =   128385027
      CurrentDate     =   37623
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   9960
      TabIndex        =   14
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9975
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Keterangan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Jumlah(Rp)"
         Object.Width           =   2540
      EndProperty
   End
   Begin XPControls.XPOption option3 
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "Per Tahun"
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
   Begin VB.Label zkt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7080
      TabIndex        =   39
      Top             =   5640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label18 
      Caption         =   "ZAKAT"
      Height          =   375
      Left            =   4320
      TabIndex        =   38
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   9480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   375
      Left            =   9600
      TabIndex        =   37
      Top             =   4680
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   6720
      X2              =   9480
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      Height          =   375
      Left            =   9600
      TabIndex        =   36
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label slsh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7680
      TabIndex        =   35
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   375
      Left            =   8640
      TabIndex        =   34
      Top             =   2880
      Width           =   255
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   8520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   375
      Left            =   8640
      TabIndex        =   33
      Top             =   1920
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   8520
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label trjc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5880
      TabIndex        =   32
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Retur Jual"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label tml 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6960
      TabIndex        =   30
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemasukan lain2"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ttldiskonc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6000
      TabIndex        =   28
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label tljual2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total diskon"
      Height          =   255
      Left            =   4320
      TabIndex        =   26
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Penjualan Total"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Laba operasional"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Biaya OPerasional"
      Height          =   615
      Left            =   4320
      TabIndex        =   21
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Laba/Rugi"
      Height          =   615
      Left            =   4320
      TabIndex        =   20
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga Pokok Jual"
      Height          =   615
      Left            =   4320
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total penjualan"
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Biaya Operasional"
      Height          =   375
      Left            =   9960
      TabIndex        =   17
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label margin2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label lain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label margin1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label ttl_beli 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label ttl_jual 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   615
      Left            =   5880
      TabIndex        =   11
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Periode"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai tanggal :"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dari tanggal :"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "laba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttlm, ttlm2, ttlk, ttlk2, sum, sum2, tljual, tljuald, ttldiskon, ttldiskon2, pm, pmd, trj As Currency
Sub itung2()
Set RS = New Recordset

sql = "select coalesce(sum(total),0) as pem from penjualan where tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ' "
Set RS = jual.Execute(sql)
ttlm = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
'sql = "select coalesce(sum(keuangan.pengeluaran),0) as pem from keuangan where (keuangan.tanggal  between ' " & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ') and keterangan like 'Retur jual'"
sql = "select sum(total) as pem from retur_jual where tanggal  between ' " & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ' "
Set RS = jual.Execute(sql)

trj = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset

sql = "select coalesce(sum(pemasukan),0) as pem from keuangan where jenis='Lain-lain' and (Tanggal between ' " & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ')"
Set RS = jual.Execute(sql)

pm1 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

pm = val(pm1)


Set RS = New Recordset
sql = "select coalesce(sum(jumlah),0) as pem from penjualan where tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "' "
Set RS = jual.Execute(sql)
tljual = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(total_diskon),0) as pem from penjualan where tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ' "
Set RS = jual.Execute(sql)

ttldiskon = IIf(IsNull(RS!pem) = True, 0, RS!pem)

RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(detiljual.harga_beli*detiljual.jumlah_brg),0) as pem from detiljual,penjualan where penjualan.tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' and  '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "'  and penjualan.no_penjualan=detiljual.no_penjualan"
Set RS = jual.Execute(sql)

ttlk = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(d.harga_beli*j.jumlah),0) as pem from detiljual d,detilreturjual j,retur_jual r where r.tanggal  between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' and  '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "'  and r.no_retur=j.no_retur and d.no_penjualan=j.no_penjualan and d.kode_brg=j.kode_brg"
Set RS = jual.Execute(sql)

ttlk2 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close
ttlk = ttlk - ttlk2
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
Printer.Print Tab(21); "LAPORAN LABA RUGI"
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

Printer.Print Tab(10); "PEMASUKAN LAIN2";
Printer.Print Tab(60); RKanan(val(pm), "###,###,##0.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(10); "PENJUALAN TOTAL";
Printer.Print Tab(50); RKanan(val(tljual), "###,###,###.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""

Printer.Print Tab(10); "TOTAL RETUR JUAL";
Printer.Print Tab(50); RKanan(val(trj), "###,###,##0.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""



Printer.Print Tab(10); "TOTAL DISKON";
Printer.Print Tab(50); RKanan(val(ttldiskon), "###,###,##0.#0");
Printer.Print Tab(46); "------------------ -";

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""

Printer.Print Tab(10); "TOTAL PENJUALAN";
Printer.Print Tab(50); RKanan(val(ttlm - trj), "###,###,##0.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""


Printer.Print Tab(10); "TOTAL HARGA POKOK PENJUALAN";
Printer.Print Tab(50); RKanan(val(ttlk), "###,###,##0.#0");
Printer.Print Tab(46); "------------------ -";
Printer.Print Tab(60); RKanan(val(ttlm - trj - ttlk), "###,###,##0.#0");
Printer.Print Tab(56); "------------------ +";

Printer.Print ; ""

Printer.Print ; ""
Printer.Print ; ""
If val(ttlm) >= val(ttlk) Then
Printer.Print Tab(10); "LABA";

Else
Printer.Print Tab(10); "RUGI";

End If
Printer.Print Tab(60); RKanan(val(ttlm - trj) + val(pm) - val(ttlk), "###,###,##0.#0");

Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(10); "BIAYA OPERASIONAL";
Printer.Print ; ""
Printer.Print ; ""

For I = 1 To ListView1.ListItems.count
'cetak isi
    Printer.Print Tab(10); ListView1.ListItems(I).SubItems(1);
    Printer.Print Tab(30); RKanan(ListView1.ListItems(I).SubItems(2), "###,###,##0.#0");
    'Printer.Print Tab(63); RKanan(ListView1.ListItems(i).SubItems(3), "###,###,##0.#0");
 
 Printer.Print ; ""
   
Next I
 Printer.Print ; ""
Printer.Print Tab(10); "Total";
Printer.Print Tab(60); RKanan(sum, "###,###,##0.#0");

Printer.Print ; ""
Printer.Print Tab(56); "------------------ -";

Printer.Print ; ""
Printer.Print ; ""
Printer.Print Tab(10); "LABA/RUGI OPERASIONAL";
Printer.Print Tab(60); RKanan(val(ttlm - trj) + val(pm) - val(ttlk) - val(sum), "###,###,##0.#0");
Printer.Print ; ""
Printer.Print ; ""
'Printer.Print Tab(10); "ZAKAT";

'Printer.Print Tab(60); RKanan((val(ttlm - trj) + val(pm) - val(ttlk) - val(sum)) * 0.025, "###,###,##0.#0");

'Printer.Print ; ""
'Printer.Print ; ""

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

Private Sub Command2_Click()
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
If Option1.Value = True Then

itung2
Else
If Option2.Value = True Then
Set RS = New Recordset

sql = "select coalesce(sum(pemasukan),0) as pem from keuangan where jenis='Lain-lain' and (month(keuangan.tanggal)='" & Combo1.Text & "' and year(keuangan.tanggal)='" & Combo2.Text & "')"
Set RS = jual.Execute(sql)

pm1 = val(RS!pem)
RS.Close
pm = val(pm1)

Set RS = New Recordset

sql = "select coalesce(sum(total),0) as pem from penjualan where (month(penjualan.tanggal)='" & Combo1.Text & "' and year(penjualan.tanggal)='" & Combo2.Text & "')"
Set RS = jual.Execute(sql)

ttlm = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset

'sql = "select coalesce(sum(keuangan.pengeluaran),0) as pem from keuangan where (month(keuangan.tanggal)='" & Combo1.Text & "' and year(keuangan.tanggal)='" & Combo2.Text & "') and keterangan like 'Retur jual'"
sql = "select sum(total) as pem from retur_jual where month(tanggal)='" & Combo1.Text & "' and year(tanggal)='" & Combo2.Text & "'  "

Set RS = jual.Execute(sql)

trj = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset

sql = "select coalesce(sum(jumlah),0) as pem from penjualan where (month(penjualan.tanggal)='" & Combo1.Text & "' and year(penjualan.tanggal)='" & Combo2.Text & "')"
Set RS = jual.Execute(sql)

tljual = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close


Set RS = New Recordset

sql = "select coalesce(sum(total_diskon),0)as pem from penjualan where (month(penjualan.tanggal)='" & Combo1.Text & "' and year(penjualan.tanggal)='" & Combo2.Text & "')"
Set RS = jual.Execute(sql)

ttldiskon = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(detiljual.harga_beli*detiljual.jumlah_brg),0) as pem from detiljual,penjualan where ((month(penjualan.tanggal)='" & Combo1.Text & "' and year(penjualan.tanggal)='" & Combo2.Text & "')) and penjualan.no_penjualan=detiljual.no_penjualan"
Set RS = jual.Execute(sql)

ttlk = IIf(IsNull(RS!pem) = True, 0, RS!pem)
Set RS = New Recordset
sql = "select coalesce(sum(d.harga_beli*j.jumlah),0)as pem from detiljual d,retur_jual r,detilreturjual j where ((month(r.tanggal)='" & Combo1.Text & "' and year(r.tanggal)='" & Combo2.Text & "')) and r.no_retur=j.no_retur and d.no_penjualan=j.no_penjualan and d.kode_brg=j.kode_brg"
Set RS = jual.Execute(sql)

ttlk2 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
ttlk = ttlk - ttlk2
Else
Set RS = New Recordset

sql = "select coalesce(sum(pemasukan),0) as pem from keuangan where jenis='Lain-lain' and (year(keuangan.tanggal)='" & Combo3.Text & "')"
Set RS = jual.Execute(sql)

pm1 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close
pm = val(pm1)
Set RS = New Recordset

sql = "select coalesce(sum(total),0)as pem from penjualan where year(penjualan.tanggal)='" & Combo3.Text & "'"
Set RS = jual.Execute(sql)

ttlm = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset

sql = "select coalesce(sum(keuangan.pengeluaran),0) as pem from keuangan where year(keuangan.tanggal)='" & Combo3.Text & "' and keterangan like 'Retur jual'"
sql = "select sum(total) as pem from retur_jual where year(tanggal)='" & Combo3.Text & "'  "

Set RS = jual.Execute(sql)

trj = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

sql = "select coalesce(sum(jumlah),0)as pem from penjualan where year(penjualan.tanggal)='" & Combo3.Text & "'"
Set RS = jual.Execute(sql)

tljual = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

sql = "select coalesce(sum(total_diskon),0)as pem from penjualan where year(penjualan.tanggal)='" & Combo3.Text & "'"
Set RS = jual.Execute(sql)

ttldiskon = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(detiljual.harga_beli*detiljual.jumlah_brg),0) as pem from detiljual,penjualan where year(penjualan.tanggal)='" & Combo3.Text & "' and penjualan.no_penjualan=detiljual.no_penjualan"
Set RS = jual.Execute(sql)

ttlk = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close

Set RS = New Recordset
sql = "select coalesce(sum(d.harga_beli*j.jumlah),0)as pem from detiljual d,retur_jual r,detilreturjual j where year(r.tanggal)='" & Combo3.Text & "' and r.no_retur=j.no_retur and d.no_penjualan=j.no_penjualan and d.kode_brg=j.kode_brg"
Set RS = jual.Execute(sql)

ttlk2 = IIf(IsNull(RS!pem) = True, 0, RS!pem)
RS.Close
ttlk = ttlk - ttlk2
End If
End If
tml.Caption = Format(val(pm), "#,#0.#0")
trjc.Caption = Format(val(trj), "#,#0.#0")
ttl_jual.Caption = Format(val(ttlm - trj), "#,#0.#0")

tljual2.Caption = Format(val(tljual), "#,#0.#0")

ttldiskonc.Caption = Format(val(ttldiskon), "#,#0.#0")

ttl_beli.Caption = Format(val(ttlk), "#,##0.#0")
slsh.Caption = Format(ttlm - trj - ttlk, "#,##0.#0")
margin1.Caption = Format(val(ttlm - trj) + val(pm) - val(ttlk), "#,#0.#0")



Set RS = New Recordset
If Option1.Value = True Then
sql = "select keterangan,sum(pengeluaran) as jum from keuangan where pengeluaran > 0  and jenis='Lain-lain' and (Tanggal between '" & Format(DTPicker1.Value, "YYYY-mm-dd") & " ' and  '" & Format(DTPicker2.Value, "YYYY-mm-dd") & " ') group by keterangan order by keterangan"

Else
If Option2.Value = True Then
sql = "select keterangan,sum(pengeluaran) as jum from keuangan where pengeluaran>0  and jenis='Lain-lain' and (month(keuangan.tanggal)='" & Combo1.Text & "' and year(keuangan.tanggal)='" & Combo2.Text & "') group by keterangan order by keterangan"

Else
sql = "select keterangan,sum(pengeluaran) as jum from keuangan where pengeluaran>0  and jenis='Lain-lain' and year(keuangan.tanggal)='" & Combo3.Text & "' group by keterangan order by keterangan"

End If
End If
Set RS = jual.Execute(sql)

Dim l As ListItem
ListView1.ListItems.Clear

If RS.RecordCount <> 0 Then
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
        
        l.SubItems(1) = ![keterangan]
       
l.SubItems(2) = Format(RS!jum, "#,#")

    .MoveNext
    Loop
End With

End If
ttl
margin2.Caption = Format(val(ttlm - trj) + val(pm) - val(ttlk) - val(sum), "#,#0.#0")
zkt.Caption = Format((val(ttlm - trj) + val(pm) - val(ttlk) - val(sum)) * 0.025, "#,#0.#0")
For I = 1 To ListView1.ListItems.count
    ListView1.ListItems(I).Text = I
Next I

Exit Sub
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
End If

End Sub
Sub ttl()
sum = 0
For I = 1 To ListView1.ListItems.count
sum = sum + val(Format(ListView1.ListItems(I).SubItems(2), Number))
Next I
lain.Caption = Format(sum, "#,#0.#0")
End Sub

Private Sub Command3_Click()
MsgBox "Jika terdapat tanggal retur jual dan penjualan berbeda,maka periode laba rugi harus dalam perode yang sama untuk mendapatkan data laba rugi yang valid.Contoh:Penjualan tgl 3,retur tgl 4.Maka periode laba rugi tgl 3 atau sblmnya sampai tgl 4 atau sesudahnya.", vbInformation
End Sub

Private Sub DTPicker1_GotFocus()
Option1.Value = True

End Sub

Private Sub DTPicker2_GotFocus()
Option1.Value = True

End Sub

Private Sub Form_Load()
 DTPicker1.Value = Now
  DTPicker2.Value = Now
Option1.Value = True
Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
Combo3.Text = Format(Now, "yyyy")
End Sub

Private Sub Label20_Click()

End Sub

