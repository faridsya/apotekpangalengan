VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapneraca 
   Caption         =   "Laporan Neraca"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapneraca.frx":0000
      Left            =   3000
      List            =   "lapneraca.frx":0028
      TabIndex        =   6
      Text            =   "tahun"
      Top             =   480
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapneraca.frx":0074
      Left            =   2280
      List            =   "lapneraca.frx":009C
      TabIndex        =   5
      Text            =   "bln"
      Top             =   480
      Width           =   615
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3840
      OleObjectBlob   =   "lapneraca.frx":00C7
      Top             =   1920
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Tahun"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bulan"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cmbthn 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "lapneraca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

    Private Const SYNCHRONIZE       As Long = &H100000
    Private Const INFINITE          As Long = &HFFFF
                Dim fso As New FileSystemObject

Private Sub Command1_Click()
Dim min, max As Integer
jual.Execute "delete from tblneracasaldo"
thn2 = cmbthn.Text
thn1 = val(cmbthn.Text) - 1
Set RS = New Recordset
RS.Open "select no_akun,nama_akun,coalesce(sum(case when tahun='" & thn1 & "' then saldo end),0) saldo1,coalesce(sum(case when tahun='" & thn2 & "' then saldo end),0) saldo2 from vneracabaru where jns=1 group by no_akun order by no_akun", jual, adOpenStatic, adLockOptimistic

Set RS2 = New Recordset
RS2.Open "select no_akun,nama_akun,coalesce(sum(case when tahun='" & thn1 & "' then saldo end),0) saldo1,coalesce(sum(case when tahun='" & thn2 & "' then saldo end),0) saldo2 from vneracabaru where jns=2 group by no_akun order by no_akun", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount >= RS2.RecordCount Then
min = RS2.RecordCount
max = RS.RecordCount
Else
max = RS2.RecordCount
min = RS.RecordCount
End If
With RS
RS.MoveFirst
RS2.MoveFirst

For I = 1 To min
jual.Execute "insert into tblneracasaldo values('" & !no_akun & "','" & !nama_akun & "','" & !saldo1 & "','" & !saldo2 & "','" & RS2!no_akun & "','" & RS2!nama_akun & "','" & RS2!saldo1 & "','" & RS2!saldo2 & "')"
RS2.MoveNext
RS.MoveNext
Next I
End With

If RS.RecordCount > RS2.RecordCount Then
With RS
For I = (min + 1) To max
jual.Execute "insert into tblneracasaldo values('" & !no_akun & "','" & !nama_akun & "','" & !saldo1 & "','" & !saldo2 & "','','',0,0)"
Next I
End With
Else
With RS2
For I = (min + 1) To max
jual.Execute "insert into tblneracasaldo values('','',0,0,'" & !no_akun & "','" & !nama_akun & "','" & !saldo1 & "','" & !saldo2 & "')"
Next I
End With

End If

End Sub

Private Sub Command2_Click()
If MsgBox("Cetak Neraca?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "drop PROCEDURE IF EXISTS apotekbaleendah.`simpan`"
jual.Execute "drop view if exists apotekbaleendah.`vuntung`,apotekbaleendah.vneracaakhir"
If Option2.Value = True Then
thn2 = cmbthn.Text
thn1 = val(cmbthn.Text) - 1

thn2 = cmbthn.Text
thn1 = val(cmbthn.Text) - 1

jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `apotekbaleendah`.`vneracaakhir` AS select `vneracabaruakhir`.`no_akun` AS `no_akun`,`vneracabaruakhir`.`jns` AS `jns`,`vneracabaruakhir`.`jns2` AS `jns2`,`vneracabaruakhir`.`urut` AS `urut`,`vneracabaruakhir`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (year(`vneracabaruakhir`.`tanggal`) = _utf8'" & thn1 & "') then `vneracabaruakhir`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (year(`vneracabaruakhir`.`tanggal`) = _utf8'" & thn2 & "') then `vneracabaruakhir`.`saldo` end)),0) AS `saldo2` from `apotekbaleendah`.`vneracabaruakhir` group by `vneracabaruakhir`.`no_akun`"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `apotekbaleendah`.`vuntung` AS select coalesce(sum((case when ((year(`j`.`tanggal`) = _utf8'" & thn1 & "') and ((`a`.`jns2` = _utf8'pendapatan') or (`a`.`jns2` = _utf8'beban'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung1`,coalesce(sum((case when ((year(`j`.`tanggal`) = _utf8'" & thn2 & "') and ((`a`.`jns2` = _utf8'pendapatan') or (`a`.`jns2` = _utf8'beban'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung2` from (`apotekbaleendah`.`jurnal` `j` join `apotekbaleendah`.`akun` `a` on((`a`.`no_akun` = `j`.`no_akun`)))"
Else
thn2 = Combo2.Text
If val(Combo1.Text) = 1 Then
thn1 = val(Combo2.Text) - 1
Else
thn1 = thn2
End If
bln2 = Combo1.Text
If val(Combo1.Text) = 1 Then
bln1 = "12"
Else

bln1 = val(Combo1.Text) - 1
End If
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `apotekbaleendah`.`vneracaakhir` AS select `vneracabaruakhir`.`no_akun` AS `no_akun`,`vneracabaruakhir`.`jns` AS `jns`,`vneracabaruakhir`.`jns2` AS `jns2`,`vneracabaruakhir`.`urut` AS `urut`,`vneracabaruakhir`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (month(`vneracabaruakhir`.`tanggal`) = _utf8'" & bln1 & "' and year(`vneracabaruakhir`.`tanggal`) = _utf8'" & thn1 & "') then `vneracabaruakhir`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (month(`vneracabaruakhir`.`tanggal`) = _utf8'" & bln2 & "' and year(`vneracabaruakhir`.`tanggal`) = _utf8'" & thn2 & "') then `vneracabaruakhir`.`saldo` end)),0) AS `saldo2` from `apotekbaleendah`.`vneracabaruakhir` group by `vneracabaruakhir`.`no_akun`"
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `apotekbaleendah`.`vuntung` AS select coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8'" & bln1 & "' and year(`j`.`tanggal`) = _utf8'" & thn1 & "') and ((`a`.`jns2` = _utf8'pendapatan') or (`a`.`jns2` = _utf8'beban'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung1`,coalesce(sum((case when ((month(`j`.`tanggal`) = _utf8'" & bln2 & "' and year(`j`.`tanggal`) = _utf8'" & thn2 & "') and ((`a`.`jns2` = _utf8'pendapatan') or (`a`.`jns2` = _utf8'beban'))) then (`j`.`kredit` - `j`.`debet`) end)),0) AS `untung2` from (`apotekbaleendah`.`jurnal` `j` join `apotekbaleendah`.`akun` `a` on((`a`.`no_akun` = `j`.`no_akun`)))"
End If

jual.Execute "CREATE DEFINER=`root`@`localhost` PROCEDURE `simpan`() BEGIN declare utg1 double(19,2);declare utg2 double(19,2);delete from tblneraca2;insert into tblneraca2 select * from vneracaakhir;set utg1=(select untung1 from vuntung);set utg2=(select untung2 from vuntung);update tblneraca2 set saldo1=utg1,saldo2=utg2 where no_akun='7.8';END;"

jual.Execute "call simpan"
With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\neracaakhir.rpt"
  .RetrieveDataFiles
    If Option2.Value = True Then

.Formulas(1) = "thn1='" & thn1 & "'"
.Formulas(2) = "thn2='" & thn2 & "'"
Else
A = MonthName(bln1)
b = MonthName(bln2)

  .Formulas(1) = "thn1='" & A & "-" & thn1 & "'"
.Formulas(2) = "thn2='" & b & "-" & Combo2.Text & "'"

End If

.Formulas(3) = "nama='" & nama_toko & "'"
.Formulas(4) = "judul2='Pengurus KPRI " & namkop & "'"
.Formulas(5) = "ketua='" & ketkop & "'"
.Formulas(6) = "sekretaris='" & sekkop & "'"
.Formulas(7) = "bendahara='" & benkop & "'"
.Formulas(8) = "tgjwb='" & tgjwb & "'"

  .WindowTitle = "Laporan Simpanan "
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
Combo1.Text = Format(Now, "mm")
Combo2.Text = Format(Now, "yyyy")
Option1.Value = True

cmbthn.Text = Format(Now, "yyyy")
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
