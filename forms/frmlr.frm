VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form laprl 
   Caption         =   "PHU"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmlr.frx":0000
      Left            =   2880
      List            =   "frmlr.frx":0028
      TabIndex        =   5
      Text            =   "tahun"
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmlr.frx":0074
      Left            =   2160
      List            =   "frmlr.frx":009C
      TabIndex        =   4
      Text            =   "bln"
      Top             =   360
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Tahun"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bulan"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cmbthn 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmlr.frx":00C7
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak PHU"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "laprl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Cetak laba rugi?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "drop view if exists vlabarugi"
If Option2.Value = True Then
thn2 = cmbthn.Text
thn1 = val(cmbthn.Text) - 1
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vlabarugi` AS select `vlaba`.`no_akun` AS `no_akun`,`vlaba`.`jns` AS `jns`,`vlaba`.`jns2` AS `jns2`,`vlaba`.`urut` AS `urut`,`vlaba`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8'" & thn1 & "') then `vlaba`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (year(`vlaba`.`tanggal`) = _utf8'" & thn2 & "') then `vlaba`.`saldo` end)),0) AS `saldo2` from `vlaba` group by `vlaba`.`no_akun`;"
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
jual.Execute "CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `vlabarugi` AS select `vlaba`.`no_akun` AS `no_akun`,`vlaba`.`jns` AS `jns`,`vlaba`.`jns2` AS `jns2`,`vlaba`.`urut` AS `urut`,`vlaba`.`nama_akun` AS `nama_akun`,coalesce(sum((case when (month(`vlaba`.`tanggal`) = _utf8'" & bln1 & "' and year(`vlaba`.`tanggal`) = _utf8'" & thn1 & "') then `vlaba`.`saldo` end)),0) AS `saldo1`,coalesce(sum((case when (month(`vlaba`.`tanggal`) = _utf8'" & bln2 & "' and year(`vlaba`.`tanggal`) = _utf8'" & thn2 & "') then `vlaba`.`saldo` end)),0) AS `saldo2` from `vlaba` group by `vlaba`.`no_akun`;"

End If

With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\labaakhir.rpt"
  .RetrieveDataFiles
  .WindowTitle = "Laporan Simpanan "
  If Option2.Value = True Then
  .Formulas(1) = "thn1='" & thn1 & "'"
.Formulas(2) = "thn2='" & thn2 & "'"
Else
A = MonthName(bln1)
b = MonthName(bln2)

  .Formulas(1) = "thn1='" & A & "-" & thn1 & "'"
.Formulas(2) = "thn2='" & b & "-" & Combo2.Text & "'"

End If
.Formulas(3) = "judul1='KPRI " & namkop & "'"
.Formulas(4) = "judul2='Pengurus KPRI " & namkop & "'"
.Formulas(5) = "ketua='" & ketkop & "'"
.Formulas(6) = "sekretaris='" & sekkop & "'"
.Formulas(7) = "bendahara='" & benkop & "'"

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
