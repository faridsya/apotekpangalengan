VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form stokbrg2 
   Caption         =   "Stok bahan baku"
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
      Width           =   1935
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
      NombreForm_ParaBorderStyleViejo=   "stokbrg2"
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
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   7455
      _ExtentX        =   13150
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
End
Attribute VB_Name = "stokbrg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With CrystalReport1
.Reset
  .Password = Chr(10) & "tujuh"

  .ReportFileName = App.Path & "\stok2.rpt"
  .RetrieveDataFiles
  If Option2.Value = True Then
  .SelectionFormula = "{tblbarang2.stok}<" & Text1.Text & ""
Else
If Option1.Value = True Then
  .SelectionFormula = "{tblbarang2.stok}=0"
  Else
  If option3.Value = True Then
    .SelectionFormula = "{tblbarang2.stok}<={tblbarang2.stok_minimal}"
    Else
    If option4.Value = True Then
    If text4.Text <> "Semua" Then
        .SelectionFormula = "{tblbarang2.deskripsi}='" & text4.Text & "'"

End If
    End If
  End If
  End If
End If

  .WindowTitle = "Laporan Data Barang"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowParentHandle = Mnutama.hWnd

        .WindowState = crptMaximized
        
  .Action = 1
End With
Me.Hide
Pesan:
If err.Description <> vbNullString Then
MsgBox "Data kosong"
End If

End Sub

Private Sub Form_Activate()
kbrg

End Sub

Private Sub Form_Load()
Ketengah Me
kbrg
End Sub

Private Sub Option1_Change()
sql = "select * from tblbarang2 where stok=0"
Set DataGrid1.DataSource = jual.Execute(sql)
End Sub
Private Sub Option3_Change()
sql = "select * from tblbarang2 where stok <= stok_minimal"
Set DataGrid1.DataSource = jual.Execute(sql)
End Sub

Private Sub Option2_Change()
Text1.SetFocus
End Sub
Private Sub kbrg()

text4.Clear
text4.AddItem "Semua"
sql = "select * from tblbarang2 order by deskripsi"
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

Private Sub text1_Change()
Option2.Value = True
If Text1.Text <> "" Then
sql = "select * from tblbarang2 where stok < " & Text1.Text & ""
Set DataGrid1.DataSource = jual.Execute(sql)
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub text4_GotFocus()
option4.Value = True
End Sub
