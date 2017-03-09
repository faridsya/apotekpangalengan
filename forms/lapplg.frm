VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapplg 
   Caption         =   "Laporan data pelanggan"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3690
   Icon            =   "lapplg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Pilih kota"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "lapplg"
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
sql = "select kota from pelanggan where kota!='' group by kota order by id_pelanggan"
Set rssupp = jual.Execute(sql)
If Not rssupp.EOF Then
rssupp.MoveFirst
 Do While Not rssupp.EOF
Combo1.AddItem rssupp!kota
rssupp.MoveNext
 Loop
  End If
rssupp.Close


  End Sub

Private Sub Command1_Click()
On Error GoTo Pesan
With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\pelanggan.rpt"
  .RetrieveDataFiles
.Formulas(0) = "nama='" & nama_toko & "'"

  .WindowTitle = "Laporan Data Barang"
    If Combo1.Text <> "Semua" Then
        .SelectionFormula = "{pelanggan.kota}='" & Combo1.Text & "'"

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
Exit Sub
Pesan:
If err.Description <> vbNullString Then
MsgBox "Lum pilih tanggal yg bener"
End If

End Sub

Private Sub Form_Activate()
supp
End Sub

Private Sub Form_Load()
Ketengah Me
supp
End Sub
