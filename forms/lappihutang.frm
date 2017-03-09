VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lappiutang 
   Caption         =   "Laporan piutang"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5490
   LinkTopic       =   "Form2"
   ScaleHeight     =   4710
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbsls 
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel label3 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lappihutang.frx":0000
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   840
      OleObjectBlob   =   "lappihutang.frx":0074
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "lappihutang.frx":00E4
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "lappihutang.frx":014E
      Top             =   3960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox text1 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   175898627
      CurrentDate     =   37623
   End
   Begin MSComCtl2.DTPicker DTPicker2 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   175898627
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lappihutang.frx":0382
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel label4 
      Height          =   375
      Left            =   960
      OleObjectBlob   =   "lappihutang.frx":03EA
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel label2 
      Height          =   255
      Left            =   960
      OleObjectBlob   =   "lappihutang.frx":0464
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel slsc 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "lappihutang.frx":04DA
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "lappiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idc As String

Private Sub cmbsls_Click()
If cmbsls.Text <> "Semua" Then
Text1.Text = "Semua"
End If

End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Belum Lunas" Then
DTPicker1.Visible = True
DTPicker2.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Option1.Visible = True
Option2.Visible = True
SkinLabel1.Visible = True
Option1.Value = True

Else
DTPicker1.Visible = False
DTPicker2.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Option1.Visible = False
Option2.Visible = False
SkinLabel1.Visible = False

End If
End Sub
Private Sub cust()

Text1.Clear
Text1.AddItem "Semua"

sql = "select * from pelanggan order by nama"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
Text1.AddItem rsplg!nama
rsplg.MoveNext
 Loop
  End If
Text1.Text = "Semua"

rsplg.Close

  End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command1_Click()
On Error GoTo Pesan
jual.Execute "delete from piutang where jumlah_piutang=0"
If cmbsls.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {sales.nama_sales}='" & cmbsls.Text & "'"
End If
tmbh2 = " and {piutang.jumlah_piutang}<>0"
With CrystalReport1
.Reset
  .ReportFileName = serperreport & "\pihutang.rpt"
  .RetrieveDataFiles
  .Formulas(0) = "nama='" & nama_toko & "'"
  .Formulas(1) = "sls='" & cmbsls.Text & "'"
    .Formulas(2) = "cust='" & Text1.Text & "'"
SelectionFormula = "{piutang.no_bayar}<>'' " & tmbh & " " & tmbh2 & ""
  .WindowTitle = "laporan"
   If Combo1.Text <> "Semua" And Text1.Text = "Semua" Then
    If Combo1.Text = "Lunas" Then
   
    .SelectionFormula = "{piutang.jumlah_piutang}={piutang.jumlah_byr} " & tmbh & " " & tmbh2 & ""
   Else
     

If Combo1.Text = "Belum Lunas" Then

                    If Option1.Value = True Then

      
     .SelectionFormula = "{Piutang.jumlah_piutang}>{piutang.jumlah_byr}and ({piutang.jatuh_tempo}>=#" & Format(DTPicker1.Value, "YYYY-mm-dd") & "# And {piutang.jatuh_tempo}<=#" & Format(DTPicker2.Value, "YYYY-mm-dd") & "# ) " & tmbh & " " & tmbh2 & ""
                  Else
                      .SelectionFormula = "{piutang.jumlah_piutang}>{piutang.jumlah_byr} " & tmbh & " " & tmbh2 & ""

                   End If
           End If
           End If
  Else
  If Text1.Text <> "Semua" And Combo1.Text = "Semua" Then
      .SelectionFormula = "{piutang.id_pelanggan}='" & idc & "' " & tmbh & " " & tmbh2 & ""
Else
If Text1.Text <> "Semua" And Combo1.Text <> "Semua" Then
        If Combo1.Text = "Lunas" Then
           .SelectionFormula = "{piutang.id_pelanggan}='" & idc & "' and {piutang.jumlah_piutang}<={piutang.jumlah_byr} " & tmbh & " " & tmbh2 & ""
        Else
            If Combo1.Text = "Belum Lunas" Then
              If Option1.Value = True Then

         .SelectionFormula = "{piutang.id_pelanggan}='" & idc & "' and {piutang.jumlah_piutang}>{piutang.jumlah_byr} and {piutang.jatuh_tempo}>=#" & DTPicker1.Value & "# And {piutang.jatuh_tempo}<=#" & DTPicker2.Value & "# " & tmbh & " " & tmbh2 & ""
              Else
                       .SelectionFormula = "{piutang.id_pelanggan}='" & idc & "' and {piutang.jumlah_piutang}>{piutang.jumlah_byr} " & tmbh & " " & tmbh2 & ""

              End If
           End If
       End If
End If
End If
End If


.Formulas(3) = "tgjwb='" & tgjwb & "'"


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
Combo1.AddItem "Semua"
Combo1.AddItem "Lunas"
Combo1.AddItem "Belum Lunas"
Ketengah Me
DTPicker1.Value = Format(Now, "YYYY-mm-dd")
DTPicker2.Value = Format(Now, "YYYY-mm-dd")
cust
sls
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub sls()

cmbsls.Clear
cmbsls.AddItem "Semua"

sql = "select nama_sales from sales order by nama_sales"
Set rsplg = jual.Execute(sql)
If Not rsplg.EOF Then
rsplg.MoveFirst

 Do While Not rsplg.EOF
cmbsls.AddItem rsplg!nama_sales
rsplg.MoveNext
 Loop
  End If
cmbsls.Text = "Semua"

rsplg.Close

  End Sub

Private Sub label1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub text1_Click()
If Text1.Text <> "Semua" Then
cmbsls.Text = "Semua"

End If

If Text1.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select p.id_pelanggan,nama_sales from pelanggan p left join sales s on p.id_sales=s.id_sales where nama='" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
idc = RS!id_pelanggan
cmbsls.Text = IIf(IsNull(RS!nama_sales) = True, "Semua", RS!nama_sales)
Else
MsgBox "Pelanggan tidak terdaftar"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)
End Sub
