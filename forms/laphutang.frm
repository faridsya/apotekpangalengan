VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form laphutang 
   Caption         =   "Laporan hutang"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbsup 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "laphutang.frx":0000
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "laphutang.frx":006A
      Top             =   2520
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   2520
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
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   360
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
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   150142979
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   150142979
      CurrentDate     =   37623
   End
   Begin ACTIVESKINLibCtl.SkinLabel label4 
      Height          =   375
      Left            =   840
      OleObjectBlob   =   "laphutang.frx":029E
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel label2 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "laphutang.frx":0318
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel label3 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "laphutang.frx":038E
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "laphutang.frx":0406
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "laphutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.Text = "Belum Lunas" Then
DTPicker1.Visible = True
DTPicker2.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
End If
End Sub

Private Sub Command1_Click()
jual.Execute "delete from hutang where jumlah_hutang=0"

If cmbsup.Text = "semua" Then
sql = ""
Else
sql = "and {tblsupplier.Supplier}='" & cmbsup.Text & "'"
End If
With CrystalReport1
  .Password = Chr(10) & "tujuh"
  .ReportFileName = serperreport & "\hutang.rpt"
  .RetrieveDataFiles
  .WindowTitle = "laporan"
   If Combo1.Text <> "Semua" Then
    If Combo1.Text = "Lunas" Then
   
  .SelectionFormula = "{Hutang.jumlah_hutang}={Hutang.jumlah_byr} " & sql & ""
Else
  .SelectionFormula = "{Hutang.jumlah_hutang}>{Hutang.jumlah_byr} and {hutang.jatuh_tempo}>=#" & Format(DTPicker1.Value, "dd MMM YYYY") & "# And {hutang.jatuh_tempo}<=#" & Format(DTPicker2.Value, "dd MMM YYYY") & "# " & sql & ""
  End If
  Else
  .SelectionFormula = "{tblsupplier.Id_supplier}<>'' " & sql & ""
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
Private Sub datasup()
cmbsup.Clear
cmbsup.AddItem "semua"
Set RS = New Recordset
sql = "select `Supplier` from tblsupplier order by `Supplier`"
Set RS = jual.Execute(sql)
If RS.EOF Then Exit Sub
RS.MoveFirst
Do While Not RS.EOF
cmbsup.AddItem RS.Fields(0)
RS.MoveNext
Loop
cmbsup.ListIndex = 0
End Sub
Private Sub Form_Load()
Combo1.AddItem "Semua"
Combo1.AddItem "Lunas"
Combo1.AddItem "Belum Lunas"
datasup
Ketengah Me
DTPicker1.Value = Format(Now, "dd MMM yyyy")
DTPicker2.Value = Format(Now, "dd MMM yyyy")
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
