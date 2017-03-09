VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPcontrols.OCX"
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form stokbrg 
   Caption         =   "Stok obat"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10200
   Begin VB.ComboBox text4 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   5295
   End
   Begin XPControls.XPOption option3 
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "<= Stok minimal"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   480
      Top             =   3960
      _ExtentX        =   1270
      _ExtentY        =   1270
      BorderStyleViejo=   2
      NombreForm_ParaBorderStyleViejo=   "stokbrg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin XPControls.XPText Text1 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
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
   Begin XPControls.XPOption Option2 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Stok kurang dari :"
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
   Begin XPControls.XPOption Option1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Stok habis"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin XPControls.XPOption option4 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Nama barang"
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
   Begin ACTIVESKINLibCtl.SkinLabel sumhj3 
      Height          =   255
      Left            =   7920
      OleObjectBlob   =   "stokbrg.frx":0000
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label sumhj2 
      Caption         =   "Label1"
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label sumhj1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7920
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label sumhb 
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "stokbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
text4.Text = Replace(text4.Text, "'", "''")

With CrystalReport1
.Reset
  

  .ReportFileName = serperreport & "\stok2.rpt"
  .RetrieveDataFiles
  .Formulas(0) = "nama='" & nama_toko & "'"

  If Option2.Value = True Then
  .SelectionFormula = "{tblbarang.stok}<" & Text1.Text & ""
Else
If Option1.Value = True Then
  .SelectionFormula = "{tblbarang.stok}=0"
  Else
  If Option3.Value = True Then
    .SelectionFormula = "{tblbarang.stok}<={tblbarang.stok_minimal}"
    Else
    If Option4.Value = True Then
    If text4.Text <> "Semua" Then
        .SelectionFormula = "{tblbarang.deskripsi}='" & text4.Text & "'"

End If
    End If
  End If
  End If
End If
  .Formulas(1) = "sumhb='" & sumhb.Caption & "'"
  .Formulas(2) = "sumhj1='" & sumhj1.Caption & "'"
  .Formulas(3) = "sumhj2='" & sumhj2.Caption & "'"
.Formulas(4) = "sumhj3='" & sumhj3.Caption & "'"
.Formulas(5) = "tgjwb='" & tgjwb & "'"

  .WindowTitle = "Laporan Data Barang"

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
MsgBox "Data kosong"
End If

End Sub

Private Sub Command2_Click()

lap.Show
End Sub

Private Sub Form_Activate()
kbrg
Set RS = New Recordset
sql = "select sum(stok*harga_beli) as pem from tblbarang "
Set RS = jual.Execute(sql)
sumhb.Caption = Format(RS!pem, "#,#")
Set RS = New Recordset
sql = "select sum(stok*harga_jual) as pem from tblbarang "
Set RS = jual.Execute(sql)
sumhj1.Caption = Format(RS!pem, "#,#")
Set RS = New Recordset
sql = "select coalesce(sum(stok*harga_jual2),0) as pem from tblbarang "
Set RS = jual.Execute(sql)
sumhj2.Caption = Format(RS!pem, "#,#")
Set RS = New Recordset
sql = "select coalesce(sum(stok*harga_jual3),0) as pem from tblbarang "
Set RS = jual.Execute(sql)
sumhj3.Caption = Format(RS!pem, "#,#")

End Sub

Private Sub Form_Load()
Ketengah Me
kbrg

Set RS = New Recordset
'RS.Open "select sum(stok*harga_beli) as pem from tblbarang ", jual, adOpenStatic, adLockOptimistic
'sumhb.Caption = Format(RS!pem, "Rp#,#0.#0")
End Sub

Private Sub Option1_Change()
sql = "select * from tblbarang where stok=0"
Set DataGrid1.DataSource = jual.Execute(sql)
End Sub
Private Sub Option3_Change()
sql = "select * from tblbarang where stok <= stok_minimal"
Set DataGrid1.DataSource = jual.Execute(sql)
End Sub

Private Sub Option2_Change()
Text1.SetFocus
End Sub
Private Sub kbrg()

text4.Clear
text4.AddItem "Semua"
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
text4.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close


  End Sub

Private Sub Text1_Change()
Option2.Value = True
If Text1.Text <> "" Then
sql = "select * from tblbarang where stok < " & Text1.Text & ""
Set DataGrid1.DataSource = jual.Execute(sql)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub text4_GotFocus()
Option4.Value = True
End Sub

