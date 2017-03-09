VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kluar_masuk 
   Caption         =   "Keluar masuk barang"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel pf2c 
      Height          =   255
      Left            =   11520
      OleObjectBlob   =   "kluar_masuk.frx":0000
      TabIndex        =   28
      Top             =   2280
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel pf1c 
      Height          =   255
      Left            =   4920
      OleObjectBlob   =   "kluar_masuk.frx":005E
      TabIndex        =   26
      Top             =   2280
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "kluar_masuk.frx":00BC
      TabIndex        =   25
      Top             =   2280
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel stoka 
      Height          =   255
      Left            =   6360
      OleObjectBlob   =   "kluar_masuk.frx":0138
      TabIndex        =   24
      Top             =   240
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel stok 
      Height          =   255
      Left            =   7080
      OleObjectBlob   =   "kluar_masuk.frx":0196
      TabIndex        =   23
      Top             =   4440
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel stokk 
      Height          =   375
      Left            =   8640
      OleObjectBlob   =   "kluar_masuk.frx":01F4
      TabIndex        =   22
      Top             =   3360
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel stokt 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "kluar_masuk.frx":0252
      TabIndex        =   21
      Top             =   3480
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":02B0
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
   End
   Begin penjualan.ThemedButton cmdproses 
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Proses"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "kluar_masuk.frx":0334
   End
   Begin ACTIVESKINLibCtl.SkinLabel karejuc 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "kluar_masuk.frx":08CE
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel rejuc 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "kluar_masuk.frx":092C
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel jujuc 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "kluar_masuk.frx":098A
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel karelic 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "kluar_masuk.frx":09E8
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel relic 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "kluar_masuk.frx":0A46
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel jubelc 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "kluar_masuk.frx":0AA4
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   6600
      OleObjectBlob   =   "kluar_masuk.frx":0B02
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":0B80
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":0C0C
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   5520
      OleObjectBlob   =   "kluar_masuk.frx":0C7E
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   255
      Left            =   4920
      OleObjectBlob   =   "kluar_masuk.frx":0CF0
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   6600
      OleObjectBlob   =   "kluar_masuk.frx":0D60
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":0DD2
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox text4 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":0E50
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "kluar_masuk.frx":0EBA
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker tgll 
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
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMM yyyy"
      Format          =   60030979
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   6600
      OleObjectBlob   =   "kluar_masuk.frx":0F26
      TabIndex        =   11
      Top             =   2880
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   6600
      OleObjectBlob   =   "kluar_masuk.frx":0FB2
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   10080
      OleObjectBlob   =   "kluar_masuk.frx":1036
      TabIndex        =   27
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "kluar_masuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim jumbel, reli, kareli, juju, reju, kareju, plus, min, pf1, pf2, akhir As Single
Dim kode As String
Dim bener As Double

Private Sub Form_Load()
kbrg
Ketengah Me
tgll.Value = Now
End Sub

Private Sub pf1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub text4_Click()
bener = 0
sql = "select * from tblbarang where deskripsi='" & Text4.Text & "' or kode_brg='" & Text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then

kode = rsbarang!kode_brg
bener = rsbarang!stok
stok.Caption = Format(rsbarang!stok, "#,#0.#0")
rsbarang.Close
End If

End Sub
Private Sub kbrg()

Text4.Clear
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text4.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

Private Sub cmdproses_Click()
On Error Resume Next
jumbel = 0
reli = 0
kareli = 0
juju = 0
kareju = 0
reju = 0
plus = 0
min = 0
pf1 = 0
pf2 = 0

Set RS = New Recordset
RS.Open "select sum(detilbeli.jumlah_brg) as pem from detilbeli,pembelian where pembelian.tanggal_pembelian  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detilbeli.kode_brg='" & kode & "' and pembelian.no_pembelian=detilbeli.no_pembelian ", jual, adOpenStatic, adLockOptimistic
jumbel = val(RS!pem)
RS.Close
jubelc.Caption = Format(jumbel, "#,#0.#0")

Set RS = New Recordset
RS.Open "select sum(detilbeli.kembali_uang) as pem from detilbeli,pembelian where pembelian.tanggal_pembelian  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detilbeli.kode_brg='" & kode & "' and pembelian.no_pembelian=detilbeli.no_pembelian ", jual, adOpenStatic, adLockOptimistic
pf1 = val(RS!pem)
RS.Close
pf1c.Caption = Format(pf1, "#,#0.#0")


Set RS = New Recordset

RS.Open "select sum(detiljual.jumlah_brg) as pem from detiljual,penjualan where penjualan.tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detiljual.kode_brg='" & kode & "' and penjualan.no_penjualan=detiljual.no_penjualan", jual, adOpenStatic, adLockOptimistic
juju = val(RS!pem)
jujuc.Caption = Format(juju, "#,#0.#0")
RS.Close

Set RS = New Recordset

RS.Open "select sum(detiljual.kembali_uang) as pem from detiljual,penjualan where penjualan.tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detiljual.kode_brg='" & kode & "' and penjualan.no_penjualan=detiljual.no_penjualan", jual, adOpenStatic, adLockOptimistic
pf2 = val(RS!pem)
pf2c.Caption = Format(pf2, "#,#0.#0")
RS.Close

Set RS = New Recordset

RS.Open "select sum(detil_returbeli.jumlah) as pem from retur_beli,detil_returbeli where retur_beli.tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detil_returbeli.kode_brg='" & kode & "' and retur_beli.no_retur=detil_returbeli.no_retur", jual, adOpenStatic, adLockOptimistic
reli = val(RS!pem)
relic.Caption = Format(reli, "#,#0.#0")
RS.Close



Set RS = New Recordset



RS.Open "select sum(detil_returjual.jumlah) as pem from retur_jual,detil_returjual where retur_jual.tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and detil_returjual.kode_brg='" & kode & "' and retur_jual.no_retur=detil_returjual.no_retur", jual, adOpenStatic, adLockOptimistic
reju = val(RS!pem)
rejuc.Caption = Format(reju, "#,#0.#0")
RS.Close
Set RS = New Recordset

RS.Open "select sum(jumlah) as pem from balik_brg1 where tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and kode_brg='" & kode & "' ", jual, adOpenStatic, adLockOptimistic
kareli = val(RS!pem)
karelic.Caption = Format(kareli, "#,#0.#0")
RS.Close
Set RS = New Recordset

RS.Open "select sum(jumlah) as pem from balik_brg2 where tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and kode_brg='" & kode & "' ", jual, adOpenStatic, adLockOptimistic
kareju = val(RS!pem)
karejuc.Caption = Format(kareju, "#,#0.#0")
RS.Close
Set RS = New Recordset

RS.Open "select sum(jumlah) as pem from sesuai where tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and kode_brg='" & kode & "' and jenis='Menambah'", jual, adOpenStatic, adLockOptimistic
plus = val(RS!pem)
stokt.Caption = Format(plus, "#,#0.#0")
RS.Close
Set RS = New Recordset

RS.Open "select sum(jumlah) as pem from sesuai where tanggal  = # " & Format(tgll.Value, "YYYY-mm-dd") & " # and kode_brg='" & kode & "' and jenis='Mengurang'", jual, adOpenStatic, adLockOptimistic
min = val(RS!pem)
stokk.Caption = Format(min, "#,#0.#0")
RS.Close
akhir = jumbel + kareli + reju + plus - juju - kareju - reli - min
stoka.Caption = Format(bener - akhir, "#,#0.#0")

End Sub
