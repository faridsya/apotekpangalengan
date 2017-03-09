VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtransaksi 
   Caption         =   "Transaksi tambahan"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "frmtransaksi.frx":0000
      Top             =   4320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cmbakun2 
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtjum 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox cmbakun 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
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
      TabIndex        =   1
      Top             =   720
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
      Format          =   125108227
      CurrentDate     =   37623
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5655
      Left            =   6000
      TabIndex        =   8
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "No.Akun"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Akun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Debet"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kredit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "notr"
         Object.Width           =   0
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "frmtransaksi.frx":0234
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmtransaksi.frx":02A8
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "frmtransaksi.frx":031E
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "frmtransaksi.frx":0390
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmtransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nmr, jns1, jns2, nmrtr As String

Private Sub cmbakun_Click()
txtjum.SetFocus
End Sub

Private Sub Command1_Click()
If cmbakun2.Text = cmbakun.Text Then
MsgBox "Akun tak boleh sama", vbCritical, judul
Exit Sub
End If
If val(txtjum.Text) = 0 Then Exit Sub
Set RS = New Recordset
RS.Open "select no_akun from akun where nama_akun='" & cmbakun.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
jns1 = RS!no_akun
Set RS = New Recordset
RS.Open "select no_akun from akun where nama_akun='" & cmbakun2.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
jns2 = RS!no_akun
If MsgBox("Simpan data?", vbYesNo, judul) = vbNo Then Exit Sub
GetNumber1
jual.Execute "insert into jurnal value('','" & Format(tgll.Value, "yyyy-mm-dd") & "','" & jns2 & "','   " & cmbakun2.Text & "',0,'" & Format(txtjum.Text, Number) & "','1','" & nmrtr & "','transaksi')"

jual.Execute "insert into jurnal value('','" & Format(tgll.Value, "yyyy-mm-dd") & "','" & jns1 & "','" & cmbakun.Text & "','" & Format(txtjum.Text, Number) & "',0,'1','" & nmrtr & "','transaksi')"
MsgBox "Data berhasil disimpan", vbInformation, judul
dbgrid
End Sub

Private Sub Command2_Click()
If lv1.ListItems.count = 0 Then Exit Sub
If MsgBox("Hapus data?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from jurnal where no_transaksi='" & lv1.SelectedItem.SubItems(6) & "'"
MsgBox "data berhasil dihapus", vbInformation, judul
dbgrid
End Sub

Private Sub Form_Load()
tgll.Value = Now
dakun
dakun2
dbgrid
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Sub dbgrid()
lv1.ListItems.Clear

Set RS = New Recordset
RS.Open "select j.*,a.nama_akun from akun a,jurnal j where j.no_akun=a.no_akun and keterangan2='transaksi' order by nmr desc", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![nmr]
        l.SubItems(2) = ![no_akun]
        l.SubItems(3) = ![nama_akun]
        l.SubItems(4) = Format(![debet], "#,#")
        l.SubItems(5) = Format(![kredit], "#,#")
        l.SubItems(6) = ![no_transaksi]

    .MoveNext
    Loop
    End With

End Sub
Sub GetNumber1()

    On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl, A, sql As String
sql = "Select no_transaksi from jurnal where no_transaksi like 'Tx%' and keterangan2='transaksi' order by no_transaksi"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "Tx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_transaksi], 8) <> "Tx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "Tx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_transaksi, 2)) + 1
               counter = "Tx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        nmrtr = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description

End Sub

Private Sub dakun()

cmbakun.Clear
sql = "select nama_akun from akun where nama_akun!='Sisa hasil usaha' order by nama_akun"
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
Private Sub dakun2()

cmbakun2.Clear

sql = "select nama_akun from akun where nama_akun!='Sisa hasil usaha' order by nama_akun"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbakun2.AddItem rsplg!nama_akun
rsplg.MoveNext
 Loop
  End If

rsplg.Close

  End Sub

Private Sub txtjum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub
