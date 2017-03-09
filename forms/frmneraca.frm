VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmneraca 
   Caption         =   "Neraca Awal"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1440
      OleObjectBlob   =   "frmneraca.frx":0000
      Top             =   5520
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   840
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   6840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   13560
      TabIndex        =   17
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox txtsaldo 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Input"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.ComboBox cmbakun 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.ComboBox cmbthn 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5655
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "kode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Akun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lv2 
      Height          =   5655
      Left            =   9720
      TabIndex        =   4
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nmr Akun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama akun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
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
      Left            =   1920
      TabIndex        =   7
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
      Format          =   120848387
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "frmneraca.frx":0234
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "frmneraca.frx":02AC
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel ttlc2 
      Height          =   375
      Left            =   2160
      OleObjectBlob   =   "frmneraca.frx":0322
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel ttlc 
      Height          =   375
      Left            =   2160
      OleObjectBlob   =   "frmneraca.frx":0380
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel Saldo 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmneraca.frx":03DE
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmneraca.frx":0446
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "frmneraca.frx":04BE
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmneraca.frx":0524
      TabIndex        =   19
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmneraca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nmr, jns1, jns2 As String
Dim total1, total2 As Double
Private Sub cmbakun_Click()
Set RS = New Recordset
RS.Open "select no_akun,jns,jns2 from akun where nama_akun='" & cmbakun.Text & "'", jual, adOpenStatic, adLockOptimistic
nmr = RS!no_akun
jns1 = RS!jns
jns2 = RS!jns2
txtsaldo.SetFocus
End Sub

Private Sub cmbakun_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub cmbthn_Click()
cmbthn_KeyPress (13)
End Sub

Private Sub cmbthn_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

If KeyAscii = 13 Then
lv1.ListItems.Clear
lv2.ListItems.Clear
Set RS = New Recordset
RS.Open "select j.*,a.nama_akun from jurnal j,akun a where debet>0 and j.no_akun=a.no_akun and no_transaksi='" & cmbthn.Text & "' AND year(j.tanggal)='" & cmbthn.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub

With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![no_akun]
        l.SubItems(2) = ![nama_akun]
        l.SubItems(3) = Format(![debet], "#,#")

                

    .MoveNext
    Loop
    End With

  Set RS = New Recordset
RS.Open "select j.*,a.nama_akun from jurnal j,akun a where kredit>0 and j.no_akun=a.no_akun and no_transaksi='" & cmbthn.Text & "' AND year(j.tanggal)='" & cmbthn.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
  With RS

    .MoveFirst
    Do While Not .EOF
     
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
        l.SubItems(1) = ![no_akun]
        l.SubItems(2) = ![nama_akun]
        l.SubItems(3) = Format(![kredit], "#,#")

                

    .MoveNext
    Loop

End With
End If
ttl
End Sub

Private Sub Command1_Click()
Set RS = New Recordset
RS.Open "select no_transaksi from jurnal where no_transaksi='" & cmbthn.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Sudah dibuat", vbInformation, judul
Exit Sub
End If
If lv2.ListItems.count = 0 And lv1.ListItems.count = 0 Then Exit Sub
If total1 <> total2 Then
MsgBox "Tidak seimbang"
Exit Sub
End If
For I = 1 To lv1.ListItems.count
jual.Execute "insert into jurnal values('','" & Format(tgll.Value, "yyyy-mm-dd") & "','" & lv1.ListItems(I).SubItems(1) & "','" & lv1.ListItems(I).SubItems(2) & "','" & Format(lv1.ListItems(I).SubItems(3), Number) & "',0,'1','" & cmbthn.Text & "','neraca awal')"

Next
For j = 1 To lv2.ListItems.count
jual.Execute "insert into jurnal values('','" & Format(tgll.Value, "yyyy-mm-dd") & "','" & lv2.ListItems(j).SubItems(1) & "','" & lv2.ListItems(j).SubItems(2) & "',0,'" & Format(lv2.ListItems(j).SubItems(3), Number) & "','1','" & cmbthn.Text & "','neraca awal')"

Next
MsgBox "Data berhasil disimpan", vbInformation, judul
End Sub

Private Sub Command2_Click()
If jns1 = "1" Then
isi
Else
isi2
End If
cmbakun.Text = ""
txtsaldo.Text = ""
nmr = ""
jns1 = ""
jns2 = ""
cmbakun.SetFocus
ttl
End Sub

Private Sub Command3_Click()
If Not lv1.SelectedItem Is Nothing Then
lv1.ListItems.Remove lv1.SelectedItem.Index
End If



End Sub
Sub ttl()
sum1 = 0
sum2 = 0
For I = 1 To lv1.ListItems.count
sum1 = sum1 + Format(lv1.ListItems(I).SubItems(3), Number)
Next
For I = 1 To lv2.ListItems.count
sum2 = sum2 + Format(lv2.ListItems(I).SubItems(3), Number)
Next
ttlc.Caption = Format(sum1, "#,#")
ttlc2.Caption = Format(sum2, "#,#")
total1 = sum1
total2 = sum2
End Sub

Private Sub Command4_Click()
If Not lv2.SelectedItem Is Nothing Then
lv2.ListItems.Remove lv2.SelectedItem.Index
End If

End Sub

Private Sub Command5_Click()
Set RS = New Recordset
RS.Open "select no_transaksi from jurnal where year(tanggal)='" & cmbthn.Text & "' and keterangan2!='neraca awal'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Tidak dapat dihapus,sudah ada transaksi", vbCritical, judul
Exit Sub
End If
jual.Execute "delete from jurnal where no_transaksi='" & cmbthn.Text & "' and keterangan2='neraca awal'"
MsgBox "data berhasil dihapus", vbInformation, judul
lv1.ListItems.Clear
lv2.ListItems.Clear
End Sub

Private Sub Command6_Click()
If MsgBox("Cetak neraca awal?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "drop view if exists vneracaawal"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `kopsumedang`.`vneracaawal` AS select `a`.`no_akun` AS `no_akun`,`a`.`nama_akun` AS `nama_akun`,`a`.`jns` AS `jns`,`a`.`jns2` AS `jns2`,coalesce(sum((case when (`a`.`jns` = _utf8'1') then (`j`.`debet` - `j`.`kredit`) else (`j`.`kredit` - `j`.`debet`) end)),0) AS `saldo`,`j`.`no_transaksi` AS `no_transaksi` " _
& " from (`kopsumedang`.`akun` `a` left join `kopsumedang`.`jurnal` `j` on(((`a`.`no_akun` = `j`.`no_akun`) and (`j`.`keterangan2` = _utf8'neraca awal') and (year(`j`.`tanggal`) = '" & cmbthn.Text & "')))) group by `a`.`no_akun`;"

With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\neracaawal2.rpt"
  .RetrieveDataFiles
.SelectionFormula = "{vneraca.no_transaksi}='" & cmbthn.Text & "'"
.Formulas(0) = "waktu='" & cmbthn.Text & "'"
  .WindowTitle = "Laporan Simpanan "
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
''Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If

End Sub

Private Sub Form_Load()
For I = 2014 To 2100
cmbthn.AddItem I
Next
cmbthn.Text = Format(Now, "yyyy")
dakun
cmbthn_KeyPress (13)
tgll.Value = Now
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub dakun()

cmbakun.Clear

sql = "select nama_akun from akun order by nama_akun"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbakun.AddItem rsplg!nama_akun
rsplg.MoveNext
 Loop
  End If

rsplg.Close

  End Sub

Sub isi()
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = nmr
        l.SubItems(2) = cmbakun.Text
        l.SubItems(3) = txtsaldo.Text

End Sub
Sub isi2()
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
        l.SubItems(1) = nmr
        l.SubItems(2) = cmbakun.Text
        l.SubItems(3) = txtsaldo.Text

End Sub

Private Sub txtsaldo_Change()
On Error Resume Next
txtsaldo.Text = Format(txtsaldo.Text, "#,#"): SendKeys "{end}"

End Sub

Private Sub txtsaldo_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
Command2_Click
End If
End Sub
