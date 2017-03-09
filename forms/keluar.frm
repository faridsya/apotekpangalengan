VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form keluarr 
   Caption         =   "Pengeluaran"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "keluar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   4935
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Input data"
      TabPicture(0)   =   "keluar.frx":324A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SkinLabel3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SkinLabel2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SkinLabel1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPicker1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "notr"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbakun"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "KET"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "jum"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SkinLabel4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Data"
      TabPicture(1)   =   "keluar.frx":3266
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "ListView3"
      Tab(1).Control(2)=   "Option1"
      Tab(1).Control(3)=   "Option2"
      Tab(1).Control(4)=   "cmdhapus"
      Tab(1).Control(5)=   "txtcari"
      Tab(1).ControlCount=   6
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   1200
         OleObjectBlob   =   "keluar.frx":3282
         TabIndex        =   19
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   -71040
         TabIndex        =   18
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   -74760
         TabIndex        =   16
         Top             =   4440
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "BANK"
         Height          =   255
         Left            =   -68040
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "KAS"
         Height          =   255
         Left            =   -69000
         TabIndex        =   14
         Top             =   4440
         Width           =   855
      End
      Begin VB.TextBox jum 
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Baru"
         Height          =   375
         Left            =   3120
         TabIndex        =   0
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox KET 
         Height          =   315
         Left            =   3360
         TabIndex        =   1
         Top             =   1740
         Width           =   3615
      End
      Begin VB.ComboBox cmbakun 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   2880
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox notr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   149487619
         CurrentDate     =   40299
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "keluar.frx":32E8
         TabIndex        =   9
         Top             =   1020
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   1200
         OleObjectBlob   =   "keluar.frx":3354
         TabIndex        =   10
         Top             =   1740
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   1200
         OleObjectBlob   =   "keluar.frx":33C6
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3855
         Left            =   -74520
         TabIndex        =   13
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6800
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No transaksi"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tanggal "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Jumlah "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Keterangan"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cari nama biaya"
         Height          =   255
         Left            =   -72840
         TabIndex        =   17
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "No.transaksi"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "keluarr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mu2, ketr, ketr2, nmrtr As String
Sub dbgridpiu()
On Error Resume Next
Dim l As ListItem

Set rstrans = New Recordset

If Option1.Value = True Then
sql = "select * from keuangan where no_transaksi<>'' and pengeluaran <>0 and jenis='Lain-lain' order by no_transaksi desc"

Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pengeluaran], "#,#")

                l.SubItems(4) = ![keterangan]

    .MoveNext
    Loop
End With
Else
sql = "select * from keuangan2 where no_transaksi<>'' and pengeluaran2 <>0 order by no_transaksi"
Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![Tanggal2], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pengeluaran2], "#,#")

                l.SubItems(4) = ![keterangan2]

    .MoveNext
    Loop
End With

End If
End Sub
Sub dbgridpiu2()
On Error Resume Next
Dim l As ListItem

Set rstrans = New Recordset

If Option1.Value = True Then
sql = "select * from keuangan where keterangan like '" & txtcari.Text & "%' and  no_transaksi<>'' and pengeluaran <>0 and jenis='Lain-lain' order by no_transaksi desc"

Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pengeluaran], "#,#")

                l.SubItems(4) = ![keterangan]

    .MoveNext
    Loop
End With
Else
sql = "select * from keuangan2 where keterangan2 like '" & txtcari.Text & "%' and  no_transaksi<>'' and pengeluaran2 <>0 order by no_transaksi"
Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![Tanggal2], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pengeluaran2], "#,#")

                l.SubItems(4) = ![keterangan2]

    .MoveNext
    Loop
End With

End If
End Sub

Sub GetNumber1()

    On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl, A, sql As String
sql = "Select no_transaksi from keuangan where no_transaksi like 'Bx%' order by no_transaksi"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "Bx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_transaksi], 8) <> "Bx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "Bx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_transaksi, 2)) + 1
               counter = "Bx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        nmrtr = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description

End Sub
Sub GetNumber2()

    On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl, A, sql As String
sql = "Select no_transaksi from keuangan2 where no_transaksi like 'By%' order by no_transaksi"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "By" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_transaksi], 8) <> "By" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "By" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_transaksi, 2)) + 1
               counter = "By" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        nmrtr = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description

End Sub

Private Sub cmdhapus_Click()
If ListView3.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin akan membatalkan biaya ini?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from keuangan where no_transaksi='" & ListView3.SelectedItem.SubItems(1) & "'"
dbgridpiu
End Sub

Private Sub Combo1_Click()
If Not Combo1.Text = "" Then
mu2 = Combo1.Text

End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command1_Click()
On Error GoTo erol
cmbakun.ListIndex = 0
Set RS = New Recordset
RS.Open "select no_akun from akun where nama_akun='" & cmbakun.Text & "'", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then
MsgBox "Nomor akun tidak ada", vbCritical, judul
Exit Sub
End If
no_akun = RS!no_akun
Set rse1 = New Recordset
rse1.Open "Select * from keuangan where keterangan='" & ket.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rse1.EOF Then
rse1.MoveFirst
ketr = rse1!keterangan
Else
ketr = ket.Text
End If

If ket.Text = "" Or jum.Text = "" Then Exit Sub



pembayaran = val(jum.Text)
GetNumber1
notr.Text = nmrtr

jual.Execute "insert into keuangan(Tanggal,Keterangan,pengeluaran,jenis,no_transaksi,no_akun) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ketr & "','" & jum.Text & "','Lain-lain','" & nmrtr & "','" & no_akun & "')"
MsgBox "Data berhasli disimpan", vbInformation, judul
dbgridpiu
Exit Sub
erol:
If err.Description <> vbNullString Then
MsgBox "Data belum lengkap", vbCritical, "Penjualan"
If ket.Text = "" Then
ket.SetFocus
Else
If jum.Text = "" Then
jum.SetFocus
End If
End If
End If

End Sub
Private Sub muang()
On Error Resume Next
Combo1.Clear
Combo1.AddItem "Dollar"
Combo1.AddItem "Rupiah"
  End Sub


Private Sub kete()
  
Set rsbarang = New Recordset
ket.Clear
sql = "select * from keuangan where pengeluaran>0 and jenis='Lain-lain' group by keterangan order by keterangan"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
ket.AddItem rsbarang!keterangan
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
   
End Sub
Private Sub lisbeban()
  
Set rsbarang = New Recordset
cmbakun.Clear
sql = "select nama_akun from akun where (jns2='Beban' and urut=10) or nama_akun='Pengambilan pribadi' order by nama_akun"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbakun.AddItem rsbarang!nama_akun
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
ket.Text = ""
jum.Text = ""
notr.Text = ""
ket.SetFocus

End Sub

Private Sub Form_Activate()
dbgridpiu
If pilih = "KAS" Or pilih = "BANK" Then

    If pilih = "KAS" Then
GetNumber1
notr.Text = nmrtr

jual.Execute "insert into keuangan(Tanggal,Keterangan,pengeluaran,jenis,no_transaksi) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ketr & "','" & jum.Text & "','Lain-lain','" & nmrtr & "')"
Else
GetNumber2
notr.Text = nmrtr
If pilih = "BANK" Then
   If junai = pembayaran Then
    jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pengeluaran2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ketr2 & "','" & pembayaran & "','Lain-lain','" & idb & "','Tunai','" & nmrtr & "')"
   Else
   If junai = 0 Then
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pengeluaran2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ketr2 & "','" & pembayaran & "','Lain-lain','" & idb & "','Giro','" & nmrtr & "')"
jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_keluar,keterangan) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & gnom & "','" & ketr2 & " ')"

Else

       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pengeluaran2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ketr2 & "','" & junai & "','Lain-lain','" & idb & "','Tunai','" & nmrtr & "')"
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pengeluaran2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ketr2 & "','" & jugir & "','Lain-lain','" & idb & "','Giro','" & nmrtr & "')"
  jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_keluar,keterangan) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & jugir & "','" & ketr2 & " ')"

  End If
End If
End If



End If
MsgBox "Data berhasil disimpan", vbInformation, judul
dbgridpiu

'Command3_Click
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pengeluaran.doc" _
                 , vbNullString, vbNullString, 1
End If

End Sub

Private Sub Form_Load()
DTPicker1.Value = Now
Ketengah Me
kete
pilih = ""
muang
Option1.Value = True
Tab1.Tab = 0
dbgridpiu
lisbeban
End Sub

Private Sub jum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub ket_Click()
jum.SetFocus
End Sub

Private Sub ket_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub Option1_Click()
dbgridpiu
End Sub

Private Sub Option2_Click()
dbgridpiu
End Sub

Private Sub txtcari_Change()
dbgridpiu2
End Sub
