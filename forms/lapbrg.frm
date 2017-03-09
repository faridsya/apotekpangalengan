VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapbrg 
   Caption         =   "Laporan data produk"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbgol 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox cmbmilik 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   2520
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "lapbrg.frx":0000
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo3"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Berdasar supplier"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "lapbrg.frx":0074
      Top             =   3120
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Berdasar Kategori"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Berdasar Produk"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "lapbrg.frx":02A8
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "lapbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub supp()
On Error Resume Next

  Dim I As Long
  Dim j As Long

Combo1.Clear
Combo1.AddItem "Semua"
Set rssupp = New Recordset
sql = "select * from tblsupplier order by id_supplier"
Set rssupp = obat.Execute(sql)
If Not rssupp.EOF Then
rssupp.MoveFirst
 Do While Not rssupp.EOF
Combo1.AddItem rssupp!deskripsi
rssupp.MoveNext
 Loop
  End If
rssupp.Close
  End Sub

Private Sub Combo1_GotFocus()
Option1.Value = True

End Sub



 Sub Combo2_GotFocus()
Option2.Value = True

End Sub
 Sub Combo3_GotFocus()
Option3.Value = True

End Sub

Private Sub Command1_Click()
'On Error GoTo Pesan
Combo1.Text = Replace(Combo1.Text, "'", "''")

If Option1.Value = True And Combo1.Text = "" Then
MsgBox "Pilih dulu", vbInformation
Combo1.SetFocus
Exit Sub
Else
If Option3.Value = True And Combo3.Text = "" Then
MsgBox "Pilih dulu", vbInformation

Combo2.SetFocus
Exit Sub

If Option2.Value = True And Combo2.Text = "" Then
MsgBox "Pilih dulu", vbInformation

Combo2.SetFocus
Exit Sub
End If
End If

End If

If cmbmilik.ListIndex = 0 Then
sql = ""
Else
sql = "  and {tblbarang.tipe}='" & cmbmilik.Text & "'"
End If
If cmbgol.ListIndex = 0 Then
sql2 = ""
Else
sql2 = "  and {tblbarang.golongan}='" & cmbgol.Text & "'"
End If
With CrystalReport1
.Reset
  .ReportFileName = serperreport & "\barang.rpt"
  .RetrieveDataFiles
.Formulas(1) = "nama='" & nama_toko & "'"

  .WindowTitle = "Laporan Data Barang"
  
  If Option1.Value = True Then
  
    If Combo1.Text <> "Semua" Then
    
        .SelectionFormula = "{tblbarang.deskripsi}='" & Combo1.Text & "' " & sql & " " & sql2 & ""
        Else
        .SelectionFormula = "{tblbarang.deskripsi}<>'' " & sql & " " & sql2 & ""

End If
.Formulas(0) = "kat='" & Combo1.Text & "'"

Else
If Option2.Value = True Then
    If Combo2.Text <> "Semua" Then
    
        .SelectionFormula = "{tblbarang.kategori}='" & Combo2.Text & "' " & sql & " " & sql2 & ""
Else
.SelectionFormula = "{tblbarang.deskripsi}<>'' " & sql & " " & sql2 & ""
End If
.Formulas(0) = "kat='" & Combo2.Text & "'"
Else
    If Combo3.Text <> "Semua" Then
    
        .SelectionFormula = "{tblsupplier.supplier}='" & Combo3.Text & "' " & sql & " " & sql2 & ""
 Else
 .SelectionFormula = "{tblbarang.deskripsi}<>'' " & sql & " " & sql2 & ""
End If
.Formulas(0) = "kat='" & Combo3.Text & "'"
End If
End If
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
'Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If

End Sub



Private Sub Form_Load()
Ketengah Me
cmbmilik.AddItem "semua"
cmbmilik.AddItem "sendiri"
cmbmilik.AddItem "titipan"
cmbmilik.ListIndex = 0
kbrg
kat
datagolongan
tsup
Option1.Value = True
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub datagolongan()
'On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbgol.Clear
cmbgol.AddItem "semua"
sql = "select distinct golongan from tblbarang where golongan is not null and golongan!='' order by golongan"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbgol.AddItem rsbarang!golongan
rsbarang.MoveNext
 Loop
  End If
  cmbgol.ListIndex = 0
rsbarang.Close
End Sub

Private Sub kat()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo2.Clear
Combo2.AddItem "Semua"

sql = "select * from tblbarang order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo2.AddItem rsbarang!kategori
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    With Combo2
    For I = 0 To .ListCount - 1
      For j = .ListCount To (I + 1) Step -1
         If .List(j) = .List(I) Then
           .RemoveItem j
         End If
      Next j
    Next I
  End With


  End Sub

Private Sub kbrg()

Combo1.Clear

Combo1.AddItem "Semua"
sql = "select deskripsi from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF

Combo1.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

Private Sub tsup()

Combo3.Clear

Combo3.AddItem "Semua"
sql = "select supplier from tblsupplier order by supplier"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF

Combo3.AddItem rsbarang!Supplier
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

