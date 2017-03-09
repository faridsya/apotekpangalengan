VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form jalan 
   Caption         =   "Surat Jalan"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel ket 
      Height          =   375
      Left            =   4560
      OleObjectBlob   =   "jalan.frx":0000
      TabIndex        =   35
      Top             =   5880
      Width           =   5295
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   8520
      OleObjectBlob   =   "jalan.frx":005E
      Top             =   120
   End
   Begin apotekbaleendah.ThemedButton cmdhapus 
      Height          =   375
      Left            =   8760
      TabIndex        =   20
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":0292
   End
   Begin apotekbaleendah.ThemedButton cmdsimpan 
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":082C
   End
   Begin apotekbaleendah.ThemedButton command1 
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   4800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Cetak"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":0DC6
   End
   Begin apotekbaleendah.ThemedButton cmdkeluar 
      Height          =   375
      Left            =   2400
      TabIndex        =   34
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Keluar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":1360
   End
   Begin apotekbaleendah.ThemedButton cmdbatal 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   33
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":18FA
   End
   Begin apotekbaleendah.ThemedButton cmdtambah 
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "jalan.frx":1E94
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   10080
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   0
   End
   Begin VB.Frame Frame 
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   9975
      Begin XPControls.XPText notr 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSComCtl2.DTPicker tgl 
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   115474435
         CurrentDate     =   40299
      End
      Begin XPControls.XPCombo nama 
         Height          =   315
         Left            =   5760
         TabIndex        =   19
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
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
      Begin XPControls.XPText almt 
         Height          =   285
         Left            =   5760
         TabIndex        =   5
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
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
      Begin XPControls.XPText telp 
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "jalan.frx":242E
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "jalan.frx":249A
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Index           =   0
         Left            =   4080
         OleObjectBlob   =   "jalan.frx":2504
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   4080
         OleObjectBlob   =   "jalan.frx":257E
         TabIndex        =   26
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   4095
      Begin VB.ComboBox satuan 
         Height          =   315
         Left            =   1800
         TabIndex        =   37
         Top             =   600
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "jalan.frx":25EA
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox text4 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin XPControls.XPText harga 
         Height          =   285
         Left            =   3480
         TabIndex        =   12
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin XPControls.XPText jum 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
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
      Begin XPControls.XPText subttl 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
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
      Begin XPControls.XPText ttlh 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         Locked          =   -1  'True
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "jalan.frx":2654
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "jalan.frx":26C8
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   6000
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   3735
      Begin XPControls.XPText gttl 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
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
         Locked          =   -1  'True
      End
      Begin XPControls.XPText dp 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
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
      Begin XPControls.XPText disk 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
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
      Begin XPControls.XPText disp 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   375
         _ExtentX        =   661
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
      Begin XPControls.XPText sisa 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
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
         Locked          =   -1  'True
      End
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Keterangan"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Satuan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jumlah"
         Object.Width           =   1764
      EndProperty
   End
   Begin XPControls.XPText total 
      Height          =   975
      Left            =   4080
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1720
      Text            =   ""
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Enabled         =   0   'False
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   255
      Left            =   4440
      OleObjectBlob   =   "jalan.frx":2732
      TabIndex        =   29
      Top             =   5520
      Width           =   4575
   End
   Begin ACTIVESKINLibCtl.SkinLabel label10 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "jalan.frx":27EE
      TabIndex        =   30
      Top             =   0
      Width           =   2295
   End
   Begin ACTIVESKINLibCtl.SkinLabel jam 
      Height          =   375
      Left            =   2640
      OleObjectBlob   =   "jalan.frx":2868
      TabIndex        =   31
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "jalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim kode, nbrg, pos, kete As String
Dim st, cosbli, cosju, sel As Currency
Dim afk As Integer

  Private Sub tampil_stn()
On Error Resume Next
satuan.Clear
sql = "select * from satuan where kode_brg='" & kode & "' order by satuan"
Set rsstn = jual.Execute(sql)
If Not rsbarang.EOF Then
rsstn.MoveFirst
 Do While Not rsstn.EOF
satuan.AddItem rsstn!satuan
rsstn.MoveNext
 Loop
  End If
rsstn.Close


  End Sub

Private Sub almt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
telp.SetFocus
End If

End Sub
Private Sub kbrg()
On Error Resume Next
  Dim I As Long
  Dim j As Long

text4.Clear
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
    With text4
    For I = 0 To .ListCount - 1
      For j = .ListCount To (I + 1) Step -1
         If .List(j) = .List(I) Then
           .RemoveItem j
         End If
      Next j
    Next I
  End With


  End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
kurir.Enabled = True
kurir.SetFocus
Else
kurir.Enabled = False
kurir.Text = ""
End If

End Sub

Private Sub Cmdbatal_Click()
awal
kosong
cmdtambah.SetFocus
End Sub
Private Sub cmdedit_Click()
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
Frame.Enabled = True
Frame2.Enabled = True

cmdsimpan.Enabled = True
cmdbatal.Enabled = True

End Sub




Private Sub Command1_Click()
On Error Resume Next
'cetak2
cetak
'tmbh_plg
cmdtambah.Enabled = True
End Sub
Sub cetak()
Dim mno, mhal, mbaris As Integer
Dim I, n As Integer
Dim mgrs, mgrss As String
If IsPrinterInstalled Then
Printer.Print "";
Printer.Print "";
Printer.Font = "Courier New"
Printer.ForeColor = &H0&
Printer.FontSize = 6
Printer.FontBold = False
Printer.CurrentX = 0
Printer.CurrentY = 0
mno = 0
mhal = 0
I = 1
mbaris = 0

Do While I <= LV1.ListItems.count
    mhal = mhal + 1
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print ; " "
    Printer.Print Tab(4); 'Form8.Caption;
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Font = "Times New Roman"
    Printer.Print Tab(10); "SURAT JALAN"
        Printer.FontBold = False
    Printer.Print ; " "

     Printer.Print Tab(11); ""
        Printer.FontSize = 12

    Printer.FontSize = 10

    Printer.Print ; Tab(55); "Kepada YTH: ";
    Printer.Print ; " "

Printer.Print Tab(10); "Telp  :087779577761";
    Printer.Print ; Tab(55); nama.Text

 '   Printer.Print ; Tab(52); "Alamat    : "; almt.Text;


    Printer.Print ; " "

    Printer.Print Tab(10); "Tanggal";
    Printer.Print Tab(29); ": "; Format(tgl.Value, "dd MMM yyyy")
    
        Printer.Print ; " "


    Printer.Print Tab(10); "Kami kirimkan barang-barang dibawah ini:";

'    Printer.Print Tab(29); ": "; kasir.Caption;

    Printer.FontBold = False
    Printer.Print ; " "
mgrs = String$(76, "=")
mgrss = String$(76, "-")
    Printer.Print ; " "

Printer.Print Tab(10); mgrs
Printer.FontBold = False
Printer.Print Tab(10); "No";

Printer.Print Tab(15); "Nama Barang";
Printer.Print Tab(50); "Jumlah";
Printer.Print Tab(58); "Satuan";
Printer.FontBold = False
Printer.Print Tab(10); mgrss
Do While I <= LV1.ListItems.count
   Set itm = LV1.ListItems.Item(I)
    mno = mno + 1
        Printer.Print Tab(10); I;
 Printer.Print Tab(15); itm.SubItems(2);
    Printer.Print Tab(51); itm.SubItems(4);
   Printer.Print Tab(58); itm.SubItems(3);
    mbaris = mbaris + 1
    I = I + 1
Loop
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""
Printer.Print ; ""

Printer.Print Tab(10); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = fale
'Printer.Print Tab(4); mgrss
'Printer.Print ; " "
'Printer.FontSize = 8
'Printer.FontBold = False
'Printer.Print Tab(5); mgrss
Printer.Print Tab(10); "Telah Diterima";
Printer.Print Tab(67); "Hormat Kami,";

Printer.Print ; " "

Printer.Print ; " "
Printer.Print ; " "
Printer.Print ; " "
Printer.Print ; " "
Printer.Print Tab(10); "_________________";
Printer.Print Tab(65); "_________________";

    Printer.Print ; " "

Printer.FontBold = False
Printer.FontItalic = False
If mbaris >= 15 Then
Printer.NewPage
End If
Loop
Printer.EndDoc
Else
   MsgBox "Printer belum terinstall di PC Anda!", _
           vbCritical, "Belum Terinstall"
End If

End Sub
Private Sub Cmdkeluar_Click()
Unload Me
End Sub
Sub cetak2()
With CrystalReport1
.Reset
  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\jalan.rpt"
  '.RetrieveDataFiles

  .WindowTitle = "invoice"

      .Formulas(0) = "tgl= '" & Format(tgl.Value, "dd-MMM-YYYY") & "'"
.Formulas(1) = "nama= '" & nama.Text & "'"
.Formulas(2) = "almt= '" & almt.Text & "'"
.Formulas(3) = "telp= '" & telp.Text & "'"
.Formulas(4) = "nama1= '" & LV1.ListItems(1).SubItems(2) & "'"
.Formulas(5) = "nama2= '" & LV1.ListItems(2).SubItems(2) & "'"
.Formulas(6) = "nama3= '" & LV1.ListItems(3).SubItems(2) & "'"
.Formulas(7) = "nama4= '" & LV1.ListItems(4).SubItems(2) & "'"
.Formulas(8) = "nama5= '" & LV1.ListItems(5).SubItems(2) & "'"

.Formulas(9) = "jum1= '" & LV1.ListItems(1).SubItems(3) & "'"
.Formulas(10) = "jum2= '" & LV1.ListItems(2).SubItems(3) & "'"
.Formulas(11) = "jum3= '" & LV1.ListItems(3).SubItems(3) & "'"
.Formulas(12) = "jum4= '" & LV1.ListItems(4).SubItems(3) & "'"
.Formulas(13) = "jum5= '" & LV1.ListItems(5).SubItems(3) & "'"


        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        ''.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
                '.Destination = crptToPrinter

  .Action = 1
End With

End Sub
Sub tmbh_plg()
If id = "" Then
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
sql = "Select id_pelanggan from pelanggan order by id_pelanggan Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = "cus0001"
Else
j = val(Right(rsplg(0), 4))
idc = "cus" + Format(Str(j + 1), "0000")

End If
jual.Execute "insert into konsumen(id_pelanggan,nama,alamat,telepon) values('" & idc & "','" & nama.Text & "','" & almt.Text & "','" & telp.Text & "')"

End If

End Sub


Sub simpandata()
On Error GoTo erol
Set RS = New Recordset
sql = "select id_konsumen from konsumen where nama_konsumen='" & nama.Text & "' and alamat='" & almt.Text & "'"
Set RS = jual.Execute(sql)
If Not RS.EOF Then
id = RS!id_konsumen
RS.Close
Else
id = ""
End If
If val(dp.Text) >= val(gttl.Text) Then
byr = val(gttl.Text)
Else
byr = val(dp.Text)
End If
sql = "insert into penjualan values('" & notr.Text & "','" & Format(tgl.Value, "YYYY-MM-dd") & "','" & nama.Text & "','" & ttlh.Text & "','" & val(disk.Text) & "','" & val(gttl.Text) & "','" & kasir.Caption & "')"
jual.Execute (sql)

Set RS = New Recordset
For z = 1 To LV1.ListItems.count

sql = "insert into detiljual values('" & notr.Text & "','" & LV1.ListItems(z).SubItems(2) & "','" & _
LV1.ListItems(z).SubItems(3) & "','" & LV1.ListItems(z).SubItems(4) & "','" & LV1.ListItems(z).SubItems(6) & "')"
jual.Execute (sql)
    Next z
    
If id = "" Then
Dim j As Integer
Dim No As String
Set rsplg = New Recordset
sql = "Select id_konsumen from konsumen order by id_konsumen Desc"
Set rsplg = jual.Execute(sql)
If rsplg.EOF = True Then
idc = "cus0001"
Else
j = val(Right(rsplg(0), 4))
idc = "cus" + Format(Str(j + 1), "0000")

End If
jual.Execute "insert into konsumen(id_konsumen,nama_konsumen,alamat,no_telp) values('" & idc & "','" & nama.Text & "','" & almt.Text & "','" & telp.Text & "')"

End If
Set RS = New Recordset
If afk > 0 Then
jual.Execute "insert into ayam(tanggal,ayam_afkir) values('" & tgl.Value & "','" & afk & "')"
End If
    MsgBox "Data sudah tersimpan", vbInformation, judul
      awal
 cmdtambah.SetFocus

erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, judul
    Exit Sub
Frame.Enabled = True
Exit Sub

End If


End Sub

Private Sub Cmdsimpan_Click()
If MsgBox("Simpan data?", vbYesNo, judul) = vbYes And nama.Text <> "" Then
simpandata
Command1.Enabled = True

If MsgBox("Cetak struk?", vbYesNo, judul) = vbYes Then
Command1_Click
Else

Command1.SetFocus
End If

Else
MsgBox "Data belum lengkap", , judul
End If

End Sub

Sub cmdtambah_Click()
Edit = False
tambah
kosong
kosong2
no_oto
supp

nama.SetFocus
End Sub

Sub tambah()
Edit = False
cmdtambah.Enabled = False
Frame.Enabled = True
cmdbatal.Enabled = True
Frame2.Enabled = True
End Sub
Sub no_oto()
Dim j As Integer
Dim br As String
Set RS = New Recordset
sql = "Select no_penjualan from penjualan order by no_penjualan Desc"
Set RS = jual.Execute(sql)
If RS.EOF = True Then
notr.Text = "000001"

Else
j = val(Right(RS(0), 4))
br = Format(Str(j + 1), "000000")
notr.Text = br
No = br

nmr = Format(Mid(RS(0), 8, 2))
thn = Format(Mid(RS(0), 5, 2))
txt = Format(Mid(No, 8, 2))
thnn = Format(Mid(No, 5, 2))

'If (val(txt) = val(nmr) + 1) Or (val(thnn) = val(thn) + 1) Then
'notr.Text = "PJ-" + Format(Now, "YY-") + Format(Now, "MM-") + "0001"
'End If
End If
End Sub
Sub kosong()
notr.Text = ""
nama.Text = ""
almt.Text = ""
telp.Text = ""
harga.Text = ""
jum.Text = ""
subttl.Text = ""
ttlh.Text = ""
disp.Text = ""
disk.Text = ""
dp.Text = ""
gttl.Text = ""
sisa.Text = ""
LV1.ListItems.Clear

End Sub


Private Sub DataGrid1_DblClick()
On Error GoTo erol
 If pos = "1" Then
 nama.Text = DataGrid1.Columns(1)
 nama_Click
 text4.SetFocus
 Else
 If pos = "2" Then
 text4.Text = DataGrid1.Columns(1)
 text4_Click
 Else
 If pos = "3" Then
 remark.Text = DataGrid1.Columns(1)
 remark_KeyPress (13)
  Else
  If pos = "4" Then
kurir.Text = DataGrid1.Columns(1)
 kurir_KeyPress (13)
End If

 End If
 End If
 End If
erol:
 If err.Description <> vbNullString Then Exit Sub



End Sub






Private Sub DTPicker3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
cmdsimpan.SetFocus
End If

End Sub

Private Sub DTPicker3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsimpan.SetFocus
End If
End Sub

Private Sub disk_Change()
gttl.Text = val(ttlh.Text) - val(disk.Text)

End Sub

Private Sub disk_GotFocus()
disp.Text = ""
End Sub

Private Sub disk_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
dp.SetFocus
End If

End Sub

Private Sub disp_Change()
If val(disp.Text) > 100 Then
MsgBox "ERORR"
disp.Text = "0"
Exit Sub
End If
disk.Text = val(disp.Text) / 100 * val(ttlh.Text)
gttl.Text = val(ttlh.Text) - val(disk.Text)
End Sub

Private Sub disp_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
gttl.Text = val(ttlh.Text) - val(disk.Text)
dp.SetFocus
End If
End Sub

Private Sub dp_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
If val(dp.Text) > val(gttl.Text) Then
kmbali = val(dp.Text) - val(gttl.Text)
MsgBox "Kembalian " & Format(kmbali, "#,###") & ""
sisa.Text = "0"
Else
sisa.Text = val(gttl.Text) - val(dp.Text)
End If

If MsgBox("Simpan data?", vbYesNo, judul) = vbYes And nama.Text <> "" Then
simpandata
Command1.Enabled = True

If MsgBox("Cetak struk?", vbYesNo, judul) = vbYes Then
Command1_Click
Else

Command1.SetFocus
End If

Else
MsgBox "Data belum lengkap", , judul
End If
End If
End Sub





Sub kosong2()
text4.Text = ""
jum.Text = ""
subttl.Text = ""
harga.Text = ""
satuan.Text = ""

End Sub
Private Sub dbgrid1_DblClick()
On Error Resume Next
If Edit = False Then
text4.Text = dbgrid1.Columns(1)
Tab1.Tab = 0
text4_Click
text5.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus
End If
End Sub

Private Sub ekor_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari

If cari Is Nothing Then
isigrid
Else
LV1.SelectedItem.SubItems(4) = LV1.SelectedItem.SubItems(4) + val(jum.Text)
LV1.SelectedItem.SubItems(6) = val(LV1.SelectedItem.SubItems(3)) * val(LV1.SelectedItem.SubItems(4))

End If
kosong2
ttl
ttl2
text4.SetFocus
End If
End Sub

Private Sub Form_Load()
Label10.Caption = Format(Now, "dd-mm-YYYY")

tgl.Value = Now
Edit = True
supp
kbrg
awal
'kasir.Caption = kasirr
Ketengah Me
 DTPicker1 = Format(Now)
  DTPicker2 = Format(Now)
  DTPicker3 = Format(Now + 14)

     Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd


End Sub
Private Sub supp()

  Dim I As Long
  Dim j As Long

nama.Clear
sql = "select * from pelanggan order by nama"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
nama.AddItem rsbarang!nama
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
  End Sub

Private Sub dbgrid2_DblClick()
On Error Resume Next

If Edit = False Then
text3.Text = Dbgrid2.Columns(1)
Tab1.Tab = 0
text4.SetFocus
Else
MsgBox "Klik tombol tambah dulu", vbInformation
Tab1.Tab = 0
cmdtambah.SetFocus

End If
End Sub

Sub dbgrid()
Set rsbarang = New Recordset
sql = "select * from tblbarang"
Set rsbarang = jual.Execute(sql)

Set dbgrid1.DataSource = rsbarang


End Sub
Sub dbgrids()
Set rssupp = New Recordset

sql = "select * from tblsupplier"
Set rssupp = jual.Execute(sql)

Set Dbgrid2.DataSource = rssupp


End Sub


Private Sub awal()
Frame.Enabled = False
Frame2.Enabled = False
Edit = True
cmdbatal.Enabled = False
cmdtambah.Enabled = True
End Sub

Private Sub hb_Click()
text5.Text = hb.Caption
text6.SetFocus
End Sub

Sub isigrid()
    Set butir = LV1.ListItems.Add
    With butir
    .SubItems(1) = LV1.ListItems.count & "."

    .SubItems(2) = text4.Text
            .SubItems(3) = satuan.Text

        .SubItems(4) = jum.Text

    End With
  LV1.Enabled = True
  kosong2
End Sub







Private Sub harga_Change()
jum.Text = ""
End Sub

Private Sub harga_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If

End Sub

Private Sub jum_Change()
subttl.Text = val(jum.Text) * val(harga.Text)
End Sub

Private Sub jum_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 44 Then  ' The ENTER key.
            KeyAscii = 46
         End If

If KeyAscii = 13 Then
If jum.Text = "" Then
jum.Text = "1"
End If

Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari

If cari Is Nothing Then
isigrid
Else
LV1.SelectedItem.SubItems(4) = LV1.SelectedItem.SubItems(4) + val(jum.Text)
LV1.SelectedItem.SubItems(6) = val(LV1.SelectedItem.SubItems(3)) * val(LV1.SelectedItem.SubItems(4))

End If
kosong2
ttl
ttl2
text4.SetFocus
kosong2
text4.SetFocus

End If
End Sub

Private Sub kurir_Change()
Set rssupp = New Recordset

sql = "select  * from kurir where namakurir like'%" & kurir.Text & "%'"
Set rspo = jual.Execute(sql)
Set DataGrid1.DataSource = rspo

End Sub

Private Sub kurir_GotFocus()
Set rssupp = New Recordset
pos = "4"
sql = "select  * from kurir"
Set rspo = jual.Execute(sql)
Set DataGrid1.DataSource = rspo


End Sub

Private Sub kurir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
nama.SetFocus
End If

End Sub

Private Sub layanan_Click()
nama.SetFocus
End Sub

Private Sub layanan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
nama.SetFocus
End If
End Sub


Private Sub nama_Click()
nama_KeyPress (13)
End Sub



Private Sub nama_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
sql = "select  * from konsumen where nama_konsumen ='" & nama.Text & "'"
Set RS = jual.Execute(sql)
If RS.EOF Then
almt.Text = ""
telp.Text = ""
almt.SetFocus
Else
almt.Text = RS!alamat
telp.Text = RS!no_telp
text4.SetFocus
RS.Close
End If
End If
End Sub



Private Sub cmdhapus_Click()
If Not LV1.SelectedItem Is Nothing Then
LV1.ListItems.Remove LV1.SelectedItem.Index
End If

For x = 1 To LV1.ListItems.count
LV1.ListItems(x).SubItems(1) = x
Next x

End Sub

Private Sub remark_Change()

End Sub

Private Sub remark_KeyPress(KeyAscii As Integer)

End Sub

Private Sub stn_Change()
stn2.Caption = stn.Text
End Sub

Private Sub stn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"    ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
End If


End Sub

Private Sub satuan_Click()
jum.SetFocus
End Sub

Private Sub telp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4.SetFocus
End If
End Sub


Private Sub Text4_Change()
If text4.Text <> "" Then
ket.Caption = ""
Else
ket.Caption = "tekan enter bila barang sudah lengkap"

End If

End Sub

Private Sub text4_Click()

Set cari = LV1.FindItem(text4.Text, 1, , 1)
LV1.SelectedItem = cari
text4.Text = Replace(text4.Text, "'", "''")

sql = "select * from tblbarang where deskripsi='" & text4.Text & "' or kode_brg='" & text4.Text & "'"
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then

kode = rsbarang!kode_brg
stn_baku = rsbarang!satuan
tampil_stn
satuan.Text = rsbarang.Fields("satuan")
jum.SetFocus
End If
If cari Is Nothing Then
jum.SetFocus
Else
jum.SetFocus
End If
End Sub





Sub ttl()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(3)) * LV1.ListItems(I).SubItems(4)
Next I
ttlh.Text = sum
total.Text = Format(sum, "#,##0")

End Sub
Sub ttl2()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + LV1.ListItems(I).SubItems(7)
Next I
afk = sum

End Sub
Sub diskon()
sum = 0
For I = 1 To LV1.ListItems.count
sum = sum + val(LV1.ListItems(I).SubItems(6)) / 100 * val(LV1.ListItems(I).SubItems(5)) * val(LV1.ListItems(I).SubItems(4))
Next I
XPText2.Text = Format(sum, "###0")

End Sub


Private Sub text4_GotFocus()
satuan.Text = ""
If LV1.ListItems.count = 0 Then
ket.Caption = ""
Else
If text4.Text = "" Then
ket.Caption = "tekan enter bila barang sudah lengkap"
End If
End If

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If text4.Text = "" And LV1.ListItems.count <> 0 Then
Command1.SetFocus
Else
text4_Click
End If
End If

End Sub

Private Sub ThemedButton1_Click()

End Sub

Private Sub Text4_LostFocus()
ket.Caption = ""
End Sub

Private Sub Timer1_Timer()
jam.Caption = Format(Now, "hh:mm:ss")

End Sub
