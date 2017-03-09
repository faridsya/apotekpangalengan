VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form lapam 
   Caption         =   "Laporan Pengambilan Bahan Baku"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "lapam.frx":0000
      Left            =   3360
      List            =   "lapam.frx":0028
      TabIndex        =   11
      Text            =   "tahun"
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lapam.frx":0074
      Left            =   2640
      List            =   "lapam.frx":009C
      TabIndex        =   10
      Text            =   "bln"
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin XPControls.XPOption Option2 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Caption         =   "Per bulan"
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
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "Per tanggal"
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak Laporan"
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   16449539
      CurrentDate     =   37623
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
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   16449539
      CurrentDate     =   37623
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileODBCSource=   "latih"
      PrintFileODBCUser=   "admin"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      Caption         =   "Bahan Baku"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Dari tanggal :"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sampai tanggal :"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Periode"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "lapam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idc As String
Private Sub Combo1_Change()
Option2.Value = True
End Sub
Private Sub Combo2_Change()
Option2.Value = True
End Sub
Private Sub kbrg()

Combo3.Clear
Combo3.AddItem "Semua"

sql = "select * from tblbarang order by deskripsi"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo3.AddItem rsbarang!deskripsi
rsbarang.MoveNext
 Loop
  End If
  Combo3.Text = "Semua"

rsbarang.Close


  End Sub

Private Sub combo3_Click()
If Combo3.Text = "Semua" Then Exit Sub
Set RS = New Recordset
RS.Open "Select* from tblbarang where deskripsi='" & Combo3.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
idc = RS!kode_brg

Else
MsgBox "Bahan baku tidak terdaftar"
End If
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
combo3_Click
End If
End Sub

Private Sub Command1_Click()
With CrystalReport1
  .Password = Chr(10) & "tujuh"
  .ReportFileName = App.Path & "\ambil_baku.rpt"
  .RetrieveDataFiles
  .WindowTitle = "laporan"
  If Option1.Value = True Then
   If Combo3.Text = "Semua" Then
  .SelectionFormula = "{ambil_baku.tanggal}>=#" & Format(DTPicker1.Value, "dd MM YYYY") & "# And {ambil_baku.tanggal}<=#" & Format(DTPicker2.Value, "dd MM YYYY") & "#"
Else
  .SelectionFormula = "{ambil_baku.kode_bahan_baku}='" & idc & "' and {ambil_baku.tanggal}>=#" & Format(DTPicker1.Value, "dd MM YYYY") & "# And {ambil_baku.tanggal}<=#" & Format(DTPicker2.Value, "dd MM YYYY") & "#"
End If
  If Not DTPicker1.Value = DTPicker2.Value Then
.Formulas(0) = "waktu='Periode : '+'" & Format(DTPicker1.Value, "dd MMM YYYY") & "'+'-'+'" & Format(DTPicker2.Value, "dd MMM YYYY") & "'"
Else
.Formulas(0) = "waktu='Periode : '+ '" & Format(DTPicker1.Value, "dd MMM YYYY") & "'"
End If

Else

q = "1 / Combo1.Text / 2000"
a = MonthName(Combo1.Text)
b = Combo2.Text
   If Combo3.Text = "Semua" Then

.SelectionFormula = "month({ambil_baku.tanggal})=" & Combo1.Text & " and year({ambil_baku.tanggal})=" & Combo2.Text & ""
Else
.SelectionFormula = "{ambil_baku.kode_bahan_baku}='" & idc & "' and month({ambil_baku.tanggal})=" & Combo1.Text & " and year({ambil_baku.tanggal})=" & Combo2.Text & ""
End If
.Formulas(0) = "waktu='" & a & "'+'-'+'" & b & "'"

End If
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
MsgBox "Lum pilih tanggal yg bener"
End If
End Sub

Private Sub DTPicker1_Change()
Option1.Value = True
End Sub
Private Sub DTPicker2_Change()
Option1.Value = True
End Sub

Private Sub Form_Load()
Ketengah Me
DTPicker1.Value = Format(Now, "dd MMM yyyy")
DTPicker2.Value = Format(Now, "dd MMM yyyy")
kbrg
End Sub

