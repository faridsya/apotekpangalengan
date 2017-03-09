VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmstokbrg 
   Caption         =   "Stok produk"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4590
   Begin VB.ComboBox cmbgudang 
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox text4 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin XPControls.XPOption option3 
      Height          =   255
      Left            =   9720
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
      NombreForm_ParaBorderStyleViejo=   "frmstokbrg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin XPControls.XPText Text1 
      Height          =   285
      Left            =   8760
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
      Left            =   7080
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
      Left            =   5520
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
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8070
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
   Begin VB.Label Label2 
      Caption         =   "Gudang"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nama barang"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmstokbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Text4.Text = Replace(Text4.Text, "'", "''")
If Text4.Text = "Semua" Then
tmbh = ""
Else
tmbh = " and {tblbarang.deskripsi}='" & Text4.Text & "'"
End If
If cmbgudang.Text = "Semua" Then
tmbh2 = ""
Else
tmbh2 = " and {gudang.nama_gudang}='" & cmbgudang.Text & "'"
End If
With CrystalReport1
.Reset

  .ReportFileName = serperreport & "\stokbrg.rpt"
  .RetrieveDataFiles
  '.Formulas(0) = "nama='" & nama_toko & "'"

  .SelectionFormula = "{tblbarang.kode_brg}<>'' " & tmbh & " " & tmbh2 & ""

  .WindowTitle = "Laporan Data Barang"
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
MsgBox "Data kosong"
End If

End Sub



Private Sub Form_Load()
Ketengah Me
kbrg
kgdg
Set RS = New Recordset
'RS.Open "select sum(stok*harga_beli) as pem from tblbarang ", jual, adOpenStatic, adLockOptimistic
'sumhb.Caption = Format(RS!pem, "Rp#,#0.#0")
End Sub

Private Sub kbrg()

Text4.Clear
Text4.AddItem "Semua"
sql = "select * from tblbarang order by deskripsi"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Text4.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
Text4.Text = "Semua"

  End Sub
Private Sub kgdg()

cmbgudang.Clear
cmbgudang.AddItem "Semua"
sql = "select * from gudang order by nama_gudang"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
cmbgudang.AddItem rsbarang!nama_gudang
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close

cmbgudang.Text = "Semua"
  End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

