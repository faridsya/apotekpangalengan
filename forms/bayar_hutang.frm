VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form bayar_hutang 
   Caption         =   "Bayar Hutang"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11340
   Icon            =   "bayar_hutang.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pembayaran"
      TabPicture(0)   =   "bayar_hutang.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "sisa1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "sisa2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ListView2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DTPicker2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Skinner1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtcari2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Data Pembayaran"
      TabPicture(1)   =   "bayar_hutang.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "ListView3"
      Tab(1).Control(2)=   "cmdhapus"
      Tab(1).Control(3)=   "txtcari"
      Tab(1).Control(4)=   "cmdcetak"
      Tab(1).Control(5)=   "CrystalReport1"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtcari2 
         Height          =   375
         Left            =   7200
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   -73560
         Top             =   5520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdcetak 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   -73080
         TabIndex        =   15
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   -69360
         TabIndex        =   13
         Top             =   4800
         Width           =   2535
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -74520
         TabIndex        =   12
         Top             =   4800
         Width           =   1215
      End
      Begin apotekbaleendah.ThemedButton command2 
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   5940
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Ba&yar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "bayar_hutang.frx":3282
      End
      Begin apotekbaleendah.ThemedButton Command1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   5820
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&Bayar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "bayar_hutang.frx":381C
      End
      Begin vbskpro.Skinner Skinner1 
         Left            =   6240
         Top             =   7620
         _ExtentX        =   1270
         _ExtentY        =   1270
         BorderStyleViejo=   2
         NombreForm_ParaBorderStyleViejo=   "bayar_hutang"
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   540
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   151388163
         CurrentDate     =   40299
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   3420
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4048
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "No Faktur"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal Faktur"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total (Rp)"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Jumlah Hutang(Rp)"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Jumlah bayar(Rp)"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   1140
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   3836
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Id Supplier"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Alamat"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Telepon"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Jumlah hutang(Rp)"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3375
         Left            =   -74520
         TabIndex        =   11
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5953
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No bayar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Jumlah bayar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "No Faktur"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Supplier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "id supp"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cari supplier atau nomor faktur"
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Cari supplier atau nomor faktur"
         Height          =   375
         Left            =   -71880
         TabIndex        =   14
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal pembayaran"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label sisa2 
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   5880
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Sisa hutang dollar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   5580
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label sisa1 
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   5700
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Sisa hutang rupiah:"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   5580
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "bayar_hutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sis1, sis2, sipi, sipiu, sipiu2, sipiu3 As Currency
Dim mu2, ket2, idc, nobp As String
Dim mulai As Boolean

Sub dbgridpiu()
On Error Resume Next

Set rstrans = New Recordset


sql = "select byr_hutang.id_supplier,byr_hutang.no_bayar,byr_hutang.tanggal,byr_hutang.jumlah_byr,byr_hutang.no_pembelian,tblsupplier.supplier from byr_hutang,tblsupplier where byr_hutang.id_supplier=tblsupplier.id_supplier order by byr_hutang.no_bayar desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_bayar]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![jumlah_byr], "#,#")

                l.SubItems(4) = ![no_pembelian]
l.SubItems(5) = ![Supplier]
l.SubItems(6) = ![id_supplier]

    .MoveNext
    Loop
End With

End Sub
Sub cetak()
'On Error Resume Next
Dim pj As Integer
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\fakturbayar2.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{byr_hutang.no_bayar}='" & nobp & "'"
    '.Formulas(0) = "almt='" & almt & "' + Chr(13) +'" & almt2 & "'"
    .Formulas(1) = "nama='" & nama_toko & "'"
    .Formulas(2) = "nama2='" & nama_toko2 & "'"

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


End Sub

Sub dbgridpiu2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select byr_hutang.id_supplier,byr_hutang.no_bayar,byr_hutang.tanggal,byr_hutang.jumlah_byr,byr_hutang.no_pembelian,tblsupplier.supplier from byr_hutang,tblsupplier where (tblsupplier.supplier like '" & txtcari.Text & "%' or byr_hutang.no_pembelian like '" & txtcari.Text & "%') and byr_hutang.id_supplier=tblsupplier.id_supplier order by byr_hutang.no_bayar desc"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_bayar]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![jumlah_byr], "#,#")

                l.SubItems(4) = ![no_pembelian]
l.SubItems(5) = ![Supplier]
l.SubItems(6) = ![id_supplier]

    .MoveNext
    Loop
End With

End Sub

Sub GetNumber()

    On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl, A, sql As String
sql = "Select no_bayar from byr_hutang where no_bayar like 'BP%' order by no_bayar"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "BP" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_bayar], 8) <> "BP" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "BP" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_bayar, 2)) + 1
               counter = "BP" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        nobp = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description

End Sub

Private Sub cmdcetak_Click()
If ListView3.ListItems.count = 0 Then Exit Sub
If MsgBox("Cetak bukti?", vbYesNo, judul) = vbNo Then Exit Sub
nobp = ListView3.SelectedItem.SubItems(1)
cetak

End Sub

Private Sub cmdhapus_Click()
If ListView3.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin akan membatalkan pembayaran ini?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from byr_hutang where no_bayar='" & ListView3.SelectedItem.SubItems(1) & "'"
dbgridpiu
supp

End Sub

Private Sub Command1_Click()
'On Error GoTo erol
 jumbay = InputBox("Masukkan jumlah hutang rupiah yang dibayar:")
  If StrPtr(jumbay) = 0 Then Exit Sub
    If jumbay = "" Then Exit Sub
    If jumbay > val(ListView2.SelectedItem.SubItems(5)) Then
    MsgBox "Kelebihan untuk tblsupplier ini"
    Exit Sub
    End If
    jual.Execute "delete from hutang where jumlah_hutang=0 and id_supplier='" & ListView2.SelectedItem.SubItems(1) & "'"
pembayaran = jumbay
sipi = jumbay
sipiu = jumbay
sipiu2 = jumbay
sipiu3 = jumbay
mulai = False

pilih = ""
GetNumber
tanya.Show
erol:
If err.Description <> vbNullString Then
    MsgBox "Error", vbCritical, "pembelian"
End If

End Sub



Private Sub dbgrid1_Click()
On Error Resume Next

ListView1.ListItems.Clear
Set RS = New Recordset
RS.Open "select hutang.no_pembelian,pembelian.tanggal_pembelian,pembelian.total_stlh_diskon,hutang.jumlah_hutang,hutang.jumlah_byr from pembelian,hutang where hutang.id_supplier='" & dbgrid1.Columns(0).Text & "' and hutang.no_pembelian=pembelian.no_pembelian and hutang.jumlah_hutang and hutang.jumlah_byr<hutang.jumlah_hutang  order by jatuh_tempo", jual, adOpenStatic, adLockOptimistic

If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
         


        Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
        l.SubItems(1) = ![no_pembelian]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
              l.SubItems(3) = ![total]
              l.SubItems(4) = ![jumlah_hutang]
                 l.SubItems(5) = ![jumlah_byr]

    .MoveNext
    Loop
End With





End Sub

Private Sub dbgrid2_Click()
On Error Resume Next
sis1 = val(Dbgrid2.Columns(4).Text) - val(Dbgrid2.Columns(6).Text)
sisa1.Caption = Format(sis1, "#,#0.#0")
sis2 = val(Dbgrid2.Columns(5).Text) - val(Dbgrid2.Columns(7).Text)
sisa2.Caption = Format(sis2, "#,#0.#0")
End Sub

Private Sub dbgrid2_DblClick()

Command1_Click

End Sub

Private Sub Form_Activate()
On Error Resume Next


keu
'jual.Execute "Insert into byr_hutang(Tanggal,jumlah_byr,no_pembelian,id_supplier,kode_bank,bentuk) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','" & val(jumbay) & "','" & ListView1.SelectedItem.SubItems(1) & "','" & ListView2.SelectedItem.SubItems(1) & "','" & idb & "','Tunai Kas')"
supp
Set cari = ListView2.FindItem(idc, 1, , 1)
ListView2.SelectedItem = cari

ListView2_Click


sis1 = 0
DTPicker2.Value = Format(Now)
pilih = ""
dbgridpiu

If mulai = False Then
If MsgBox("Cetak bukti?", vbYesNo, judul) = vbNo Then Exit Sub
cetak
End If


End Sub

Sub keu()
For I = 1 To ListView1.ListItems.count

If sipiu <= 0 Then Exit Sub
Set RS = New Recordset

RS.Open "select * from hutang where no_pembelian='" & ListView1.ListItems(I).SubItems(1) & "' ", jual, adOpenStatic, adLockOptimistic
dipren = RS!jumlah_hutang - RS!jumlah_byr

If sipiu >= val(dipren) Then
msk = dipren

sipiu = sipiu - (dipren)
jual.Execute "Insert into byr_hutang(no_bayar,Tanggal,jumlah_byr,no_pembelian,id_supplier) values('" & nobp & "','" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','" & val(msk) & "','" & ListView1.ListItems(I).SubItems(1) & "','" & ListView2.SelectedItem.SubItems(1) & "')"



Else

jual.Execute "Insert into byr_hutang(no_bayar,Tanggal,jumlah_byr,no_pembelian,id_supplier) values('" & nobp & "','" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','" & val(sipiu) & "','" & ListView1.ListItems(I).SubItems(1) & "','" & ListView2.SelectedItem.SubItems(1) & "')"

sipiu = 0
End If
RS.Close

Next I
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\byrhutang.doc" _
                 , vbNullString, vbNullString, 1
End If

End Sub

Private Sub Form_Load()
Ketengah Me
DTPicker2.Value = Format(Now)
pilih = ""
mulai = True

supp
dbgridpiu
Tab1.Tab = 0
End Sub

Private Sub supp()
On Error Resume Next
ListView2.ListItems.Clear

Set RS = New Recordset
sql = "select"
RS.Open "select s.* from tblsupplier s join hutang h on s.id_supplier=h.id_supplier where (s.jumlah_hutang > 0 and Supplier like '%" & txtcari2.Text & "%') or (h.Jumlah_byr<h.Jumlah_hutang and h.no_pembelian='" & txtcari2.Text & "') group by s.id_supplier order by Supplier", jual, adOpenStatic, adLockOptimistic

If RS.RecordCount = 0 Then Exit Sub


With RS
.MoveFirst
    Do While Not .EOF
         


        Set l = ListView2.ListItems.Add(, , ListView2.ListItems.count + 1)
        l.SubItems(1) = ![id_supplier]
        l.SubItems(2) = Format(![Supplier], "dd MMM yyyy")
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![no_telp]
        l.SubItems(5) = ![jumlah_hutang]

    .MoveNext
    Loop
End With



End Sub

Private Sub ListView1_DblClick()
'On Error GoTo erol
 jumbay = InputBox("Masukkan jumlah hutang rupiah yang dibayar:")
  If StrPtr(jumbay) = 0 Then Exit Sub
    If jumbay = "" Then Exit Sub
    If jumbay > (val(ListView1.SelectedItem.SubItems(4)) - val(ListView1.SelectedItem.SubItems(5))) Then
    MsgBox "Kelebihan untuk tblsupplier ini"
    Exit Sub
    End If
    jual.Execute "delete from hutang where jumlah_hutang=0 and id_supplier='" & ListView2.SelectedItem.SubItems(1) & "'"
pembayaran = jumbay
sipi = jumbay
sipiu = jumbay
sipiu2 = jumbay
sipiu3 = jumbay
mulai = False

pilih = ""
GetNumber
jual.Execute "Insert into byr_hutang(no_bayar,Tanggal,jumlah_byr,no_pembelian,id_supplier) values('" & nobp & "','" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','" & val(jumbay) & "','" & ListView1.SelectedItem.SubItems(1) & "','" & ListView2.SelectedItem.SubItems(1) & "')"
supp
Set cari = ListView2.FindItem(idc, 1, , 1)
ListView2.SelectedItem = cari

ListView2_Click


sis1 = 0
DTPicker2.Value = Format(Now)
pilih = ""
dbgridpiu

If mulai = False Then
If MsgBox("Cetak bukti?", vbYesNo, judul) = vbNo Then Exit Sub
cetak
End If



erol:
If err.Description <> vbNullString Then
    MsgBox "Error", vbCritical, "pembelian"
End If
End Sub

Private Sub ListView2_Click()
On Error Resume Next
If ListView2.ListItems.count = 0 Then Exit Sub

ListView1.ListItems.Clear
idc = ListView2.SelectedItem.SubItems(1)

Set RS = New Recordset
RS.Open "select hutang.no_pembelian,pembelian.tanggal_pembelian,pembelian.total_stlh_diskon,hutang.jumlah_hutang,hutang.jumlah_byr from pembelian,hutang where hutang.id_supplier='" & ListView2.SelectedItem.SubItems(1) & "' and hutang.no_pembelian=pembelian.no_pembelian and hutang.jumlah_byr<hutang.jumlah_hutang  order by jatuh_tempo,no_pembelian", jual, adOpenStatic, adLockOptimistic

If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
         


        Set l = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)
        l.SubItems(1) = ![no_pembelian]
        l.SubItems(2) = Format(![tanggal_pembelian], "dd MMM yyyy")
              l.SubItems(3) = ![total_stlh_diskon]
              l.SubItems(4) = ![jumlah_hutang]
                 l.SubItems(5) = ![jumlah_byr]

    .MoveNext
    Loop
End With




End Sub



Private Sub txtcari_Change()
dbgridpiu2
End Sub


Private Sub txtcari2_Change()
ListView1.ListItems.Clear
supp
End Sub
