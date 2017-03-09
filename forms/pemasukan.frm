VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form masuk 
   Caption         =   "Pemasukan lain-lain"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8400
   Icon            =   "pemasukan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Input"
      TabPicture(0)   =   "pemasukan.frx":324A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "notr"
      Tab(0).Control(1)=   "jum"
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "ket"
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(9)=   "SkinLabel3"
      Tab(0).Control(10)=   "SkinLabel2"
      Tab(0).Control(11)=   "SkinLabel1"
      Tab(0).Control(12)=   "DTPicker1"
      Tab(0).Control(13)=   "Label1"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Data"
      TabPicture(1)   =   "pemasukan.frx":3266
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdhapus"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtcari"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   3360
         TabIndex        =   20
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox notr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72720
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   5280
         TabIndex        =   15
         Top             =   4200
         Width           =   2655
         Begin VB.OptionButton Option4 
            Caption         =   "KAS"
            Height          =   255
            Left            =   720
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "BANK"
            Height          =   255
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox jum 
         Height          =   375
         Left            =   -72720
         TabIndex        =   2
         Top             =   3180
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   -71640
         TabIndex        =   3
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Keluar"
         Height          =   375
         Left            =   -70200
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Baru"
         Height          =   375
         Left            =   -72840
         TabIndex        =   0
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox ket 
         Height          =   375
         Left            =   -72720
         TabIndex        =   1
         Top             =   2460
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -70920
         TabIndex        =   7
         Top             =   3180
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Modal"
         Height          =   255
         Left            =   -74400
         TabIndex        =   6
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pemasukan"
         Height          =   255
         Left            =   -72120
         TabIndex        =   5
         Top             =   540
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   -74400
         OleObjectBlob   =   "pemasukan.frx":3282
         TabIndex        =   8
         Top             =   3180
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   -74400
         OleObjectBlob   =   "pemasukan.frx":32EC
         TabIndex        =   9
         Top             =   2460
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   -74400
         OleObjectBlob   =   "pemasukan.frx":335E
         TabIndex        =   10
         Top             =   1740
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72720
         TabIndex        =   12
         Top             =   1740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108855299
         CurrentDate     =   40299
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3855
         Left            =   120
         TabIndex        =   14
         Top             =   360
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
         Left            =   1920
         TabIndex        =   21
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "No.transaksi"
         Height          =   255
         Left            =   -74400
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "masuk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mu2, nmrtr As String
Sub GetNumber1()

    On Error GoTo salah
    Dim counter As String * 10
    Dim Hitung As Integer
    Dim tgl, A, sql As String
sql = "Select no_transaksi from keuangan where no_transaksi like 'Mx%' order by no_transaksi"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "Mx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_transaksi], 8) <> "Mx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "Mx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_transaksi, 2)) + 1
               counter = "Mx" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
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
sql = "Select no_transaksi from keuangan2 where no_transaksi like 'My%' order by no_transaksi"
    Set rstrans = jual.Execute(sql)

    tgl = Format(Now, "dd/mm/yyyy")
    With rstrans
        If .RecordCount = 0 Then
            counter = "My" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
        Else
           .MoveLast
            If Left(![no_transaksi], 8) <> "My" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) Then
            counter = "My" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + "01"
            Else
                Hitung = val(Right(!no_transaksi, 2)) + 1
               counter = "My" + Right(tgl, 2) + Mid(tgl, 4, 2) + Left(tgl, 2) + Right("00" & Hitung, 2)
            End If
        End If
        nmrtr = counter
    End With
    Exit Sub
salah:
    MsgBox err.Description

End Sub

Sub dbgridpiu()
On Error Resume Next
Dim l As ListItem

Set rstrans = New Recordset

If Option4.Value = True Then
sql = "select * from keuangan where no_transaksi<>'' and pemasukan <>0 and jenis='Lain-lain' order by no_transaksi desc"

Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pemasukan], "#,#")

                l.SubItems(4) = ![keterangan]

    .MoveNext
    Loop
End With
Else
If Option3.Value = True Then
sql = "select * from keuangan2 where no_transaksi<>''and pemasukan2 <>0 order by no_transaksi"
Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![Tanggal2], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pemasukan2], "#,#")

                l.SubItems(4) = ![keterangan2]

    .MoveNext
    Loop
End With
End If
End If
End Sub
Sub dbgridpiu2()
On Error Resume Next
Dim l As ListItem

Set rstrans = New Recordset

If Option4.Value = True Then
sql = "select * from keuangan where keterangan like '" & txtcari.Text & "%' and no_transaksi<>'' and pemasukan <>0 and jenis='Lain-lain' order by no_transaksi desc"

Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![tanggal], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pemasukan], "#,#")

                l.SubItems(4) = ![keterangan]

    .MoveNext
    Loop
End With
Else
If Option3.Value = True Then
sql = "select * from keuangan2 where keterangan2 like '" & txtcari.Text & "%' and no_transaksi<>'' and pemasukan2 <>0 order by no_transaksi"
Set rstrans = jual.Execute(sql)
ListView3.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = ListView3.ListItems.Add(, , ListView3.ListItems.count + 1)
        l.SubItems(1) = ![no_transaksi]
        l.SubItems(2) = Format(![Tanggal2], "dd MMM yyyy")
                                l.SubItems(3) = Format(![pemasukan2], "#,#")

                l.SubItems(4) = ![keterangan2]

    .MoveNext
    Loop
End With
End If
End If
End Sub

Private Sub cmdhapus_Click()
If ListView3.ListItems.count = 0 Then Exit Sub
If MsgBox("Yakin akan membatalkan pemasukan ini?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from keuangan where no_transaksi='" & ListView3.SelectedItem.SubItems(1) & "'"
jual.Execute "delete from keuangan2 where no_transaksi='" & ListView3.SelectedItem.SubItems(1) & "'"
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
If ket.Text = "" Or jum.Text = "" Then Exit Sub
pembayaran = val(jum.Text)
    GetNumber1
    notr.Text = nmrtr
   If Option1.Value = True Then
   jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan,jenis,no_transaksi) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ket.Text & "','" & jum.Text & "','Modal','" & nmrtr & "')"
   Else
   jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan,jenis,no_transaksi) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ket.Text & "','" & jum.Text & "','Lain-lain','" & nmrtr & "')"
   End If
   MsgBox "Data berhasil disimpan", vbInformation

dbgridpiu
Exit Sub
'tanya.Show
erol:
If err.Description <> vbNullString Then
MsgBox "Data belum lengkap", vbCritical, judul
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

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
ket.Text = ""
jum.Text = ""
ket.SetFocus

End Sub

Private Sub Form_Activate()
If pilih = "KAS" Or pilih = "BANK" Then

    If pilih = "KAS" Then
    GetNumber1
    notr.Text = nmrtr
   If Option1.Value = True Then
   jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan,jenis,no_transaksi) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ket.Text & "','" & jum.Text & "','Modal','" & nmrtr & "')"
   Else
   jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan,jenis,no_transaksi) values('" & Format(DTPicker1.Value, "YYYY-mm-dd") & "','" & ket.Text & "','" & jum.Text & "','Lain-lain','" & nmrtr & "')"
   End If
Else
GetNumber2
    notr.Text = nmrtr

If pilih = "BANK" Then
 If Option1.Value = True Then
   If junai = pembayaran Then
    jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ket.Text & " ','" & pembayaran & "','Modal','" & idb & "','Tunai','" & nmrtr & "')"
   Else
   If junai = 0 Then
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ket.Text & " ','" & pembayaran & "','Modal','" & idb & "','Giro','" & nmrtr & "')"
jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_masuk,keterangan) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & gnom & "','" & ket.Text & " ')"

Else
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ket.Text & " ','" & junai & "','Modal','" & idb & "','Tunai','" & nmrtr & "')"
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ket.Text & " ','" & jugir & "','Modal','" & idb & "','Giro','" & nmrtr & "')"
  jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_masuk,keterangan) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & gnom & "','" & ket.Text & " ')"

  End If
End If
Else
 If junai = pembayaran Then
    jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ket.Text & " ','" & pembayaran & "','Lain-lain','" & idb & "','Tunai','" & nmrtr & "')"

   Else
   If junai = 0 Then
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ket.Text & " ','" & pembayaran & "','Lain-lain','" & idb & "','Giro','" & nmrtr & "')"
Else
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & ket.Text & " ','" & junai & "','Lain-lain','" & idb & "','Tunai','" & nmrtr & "')"
       jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,pemasukan2,jenis2,kode_bank,bentuk,no_transaksi) values('" & Format(gtgl, "YYYY-mm-dd") & "','" & ket.Text & " ','" & jugir & "','Lain-lain','" & idb & "','Giro','" & nmrtr & "')"
    jual.Execute "insert into giro(tanggal,no_giro,tgl_jt,kode_bank,giro_masuk,keterangan) values('" & Format(DTPicker1, "YYYY-mm-dd") & "','" & gno & "','" & Format(gtgl, "YYYY-mm-dd") & "','" & idb & "','" & gnom & "','" & ket.Text & " ')"

  End If
End If
End If
End If





End If
dbgridpiu

MsgBox "Data berhasil disimpan", vbInformation
'Command3_Click
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF10 Then
ShellExecute Me.hwnd, "open", App.Path & "\panduan\pemasukan.doc" _
                 , vbNullString, vbNullString, 1
End If

End Sub

Private Sub Form_Load()
pilih = ""
Ketengah Me
muang
DTPicker1.Value = Now
Option2.Value = True
Option4.Value = True
dbgridpiu
Tab1.Tab = 0
End Sub

Private Sub jum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub ket_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub Option3_Click()
dbgridpiu
End Sub

Private Sub option4_Click()
dbgridpiu

End Sub

Private Sub txtcari_Change()
dbgridpiu2
End Sub
