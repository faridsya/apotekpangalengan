VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmout 
   Caption         =   "Pengambilan barang"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbbrg 
      Height          =   315
      Left            =   3120
      TabIndex        =   16
      Top             =   4440
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmout.frx":0000
      TabIndex        =   14
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   5040
      Width           =   1695
   End
   Begin XPControls.XPText txtbrg 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   3960
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPText txtkomisi 
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblttl 
      Height          =   375
      Left            =   3120
      OleObjectBlob   =   "frmout.frx":0084
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblbiaya2 
      Height          =   375
      Left            =   3120
      OleObjectBlob   =   "frmout.frx":00E2
      TabIndex        =   8
      Top             =   2280
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblbiaya 
      Height          =   375
      Left            =   3120
      OleObjectBlob   =   "frmout.frx":0140
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblnmr 
      Height          =   255
      Left            =   3120
      OleObjectBlob   =   "frmout.frx":019E
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmout.frx":01FC
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmout.frx":027E
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmout.frx":02F8
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "frmout.frx":0360
      Top             =   4920
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5160
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmout.frx":0594
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmout.frx":060A
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmout.frx":0690
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker tgl 
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
      Left            =   3000
      TabIndex        =   15
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
      Format          =   124256259
      CurrentDate     =   37623
   End
End
Attribute VB_Name = "frmout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbbrg_Click()
If cmbbrg.Text = "" Then Exit Sub
If txtbrg.Text = "" Then
txtbrg.Text = cmbbrg.Text
Else
txtbrg.Text = txtbrg.Text + "," + txtbrg.Text
End If
End Sub

Private Sub Command1_Click()
If MsgBox("Proses pengambilan oleh konsumen?", vbYesNo, judul) = vbNo Then Exit Sub

jual.Execute "update tservis set tgl_out='" & Format(tgl.Value, "yyyy-mm-dd") & "',status_servis='diambil',komisi_teknisi='" & Format(txtkomisi.Text, Number) & "',brgservis='" & txtbrg.Text & "' where no_servis='" & lblnmr.Caption & "'"
jual.Execute "update tservis_dtl2 set stts_brg='out' where no_servis='" & lblnmr.Caption & "'"
If MsgBox("Update status berhasil,cetak faktur?", vbYesNo, judul) = vbYes Then
cetakfaktur
End If
frmpaket.Dbgrid2
End Sub
Private Sub lisbrg()
On Error Resume Next

  Dim I As Long
  Dim j As Long
Set rsbarang = New Recordset

cmbbrg.Clear
sql = "select nama_brg from tservis t,tservis_dtl1 d where t.no_servis=d.no_servis and t.no_servis='" & lblnmr.Caption & "' order by nama_brg"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbbrg.AddItem rsbarang!nama_brg
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

  End Sub

Private Sub Form_Load()
tgl.Value = Now
lisbrg
Skinpath = App.Path & "\skin\green.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
End Sub
Sub cetakfaktur()
On Error Resume Next
With CrystalReport1
.Reset
 
  .ReportFileName = serperreport & "\fakturambil.rpt"
  .RetrieveDataFiles
.CopiesToPrinter = 1
  .WindowTitle = "invoice"
.SelectionFormula = "{tservis.no_servis}='" & lblnmr.Caption & "'"
    .Formulas(0) = "nama='" & nama_toko & "'"

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

End Sub

