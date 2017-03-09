VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistem Informasi Kriminalitas (File Kriminalitas)"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4710
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3240
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Height          =   1065
      Left            =   720
      TabIndex        =   17
      Top             =   3090
      Width           =   3855
      Begin VB.CommandButton command4 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2310
         Picture         =   "Form3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton command2 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   810
         Picture         =   "Form3.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton command1 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   60
         Picture         =   "Form3.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton command3 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         Picture         =   "Form3.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton command5 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3060
         Picture         =   "Form3.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4950
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2130
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   3075
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3855
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         Top             =   2670
         Width           =   2085
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "Form3.frx":15E4
         Left            =   1680
         List            =   "Form3.frx":15E6
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   0
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   930
         Width           =   1755
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   1680
         TabIndex        =   8
         Top             =   1590
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   345
         Left            =   1680
         TabIndex        =   7
         Top             =   1950
         Width           =   2085
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "Form3.frx":15E8
         Left            =   1680
         List            =   "Form3.frx":15EA
         TabIndex        =   2
         Top             =   1260
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   2310
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   16761024
         CalendarTitleBackColor=   12640511
         CalendarTitleForeColor=   12582912
         CalendarTrailingForeColor=   -2147483624
         CustomFormat    =   "MM/dd/yy"
         Format          =   16515075
         CurrentDate     =   38140
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   16
         Top             =   1950
         Width           =   1245
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nama Pelaku"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   390
         TabIndex        =   15
         Top             =   930
         Width           =   1275
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kode Pelaku"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   270
         TabIndex        =   14
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nomor Pasal"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   1260
         Width           =   1185
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nomor Berkas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   510
         TabIndex        =   12
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tempat Kejadian"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   300
         TabIndex        =   11
         Top             =   2670
         Width           =   1365
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Kejadian"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   210
         TabIndex        =   10
         Top             =   2310
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tindak Pidana"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         TabIndex        =   9
         Top             =   1590
         Width           =   1245
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5370
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4170
      Width           =   1185
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset
Const putih = &HFFC0FF
Const abu = &HFFC0C0
Dim boleh As Boolean

Private Sub Combo1_Change()
If Trim(Combo1.Text) = "" Then
    text2.Text = ""
    Combo1.Text = ""
Else
    Combo1_Click
End If
End Sub

Private Sub Combo1_Click()
'On Error Resume Next
Set rs2 = New Recordset
rs2.Open ("select nm_pelaku from tpelaku where kd_pelaku='" & Combo1.Text & "'"), db, adOpenStatic, adLockOptimistic
text2.Text = rs2(0)
Set rs2 = Nothing
End Sub

Private Sub Combo2_Change()
If Trim(Combo2.Text) = "" Then
    Text3.Text = ""
    Text4.Text = ""
Else
    Combo2_Click
End If
End Sub

Private Sub Combo2_Click()
'On Error Resume Next
Set rs3 = New Recordset

rs3.Open ("select * from tpasal where no_pas='" & Combo2.Text & "'"), db, adOpenStatic, adLockOptimistic
Text3.Text = rs3(1)
Text4.Text = rs3(2)
Set rs3 = Nothing
End Sub

Private Sub Command7_Click()
kode_oto
text1_KeyPress (13)

End Sub
Sub kode_oto()
Dim j As Integer
Dim no As String
Set rsbarang = New Recordset
sql = "Select no_berkas from tkriminal order by no_berkas Desc"
Set rsbarang = db.Execute(sql)
If rsbarang.EOF = True Then
Text1.Text = "BS-00001"
Else
j = val(Right(rsbarang(0), 5))
no = "BS-" + Format(Str(j + 1), "00000")
Text1.Text = no

End If
End Sub


Private Sub Form_Load()
Set rs = New Recordset
 rs.Open ("select * from tpasal order by no_pas"), db, adOpenStatic, adLockOptimistic
Set rs2 = New Recordset

rs2.Open ("select * from tpelaku  order by  kd_pelaku"), db, adOpenStatic, adLockOptimistic
If rs.RecordCount <> 0 Then
    rs.MoveFirst
End If
Do While Not rs.EOF
  Combo2.AddItem rs(0)
  rs.MoveNext
Loop
If rs2.RecordCount <> 0 Then
    rs2.MoveFirst
End If
Do While Not rs2.EOF
  Combo1.AddItem rs2(0)
  rs2.MoveNext
Loop
kosong
End Sub

Private Sub tombol(A As Boolean, B As Boolean, c As Boolean)
  Command1.Enabled = A
  Command2.Enabled = B
  Command3.Enabled = c
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
kosong
 If Trim(Text1.Text) = "" Then
    MsgBox "Isi Data Dulu"
Else
Set rs = New Recordset
rs.Open ("select * from tkriminal where no_berkas='" & Text1.Text & "' "), db, adOpenStatic, adLockOptimistic
 If rs.EOF Then
  aktif True
  tombol True, False, False
  Combo1.SetFocus
 Else
  cari
 End If
'Set rs = Nothing
End If
End If
End Sub

Private Sub cari()
  'aktif True

 tgl.Value = rs(3)
 Combo1.Text = rs(1)
 Combo2.Text = rs.Fields(2)
 Text5.Text = rs(4)
 If MsgBox("Data Sudah Ada..........,Update Data??", vbQuestion + vbYesNo, "Redundandy Input") = vbYes Then
  aktif True
  tombol False, True, False
 Else
  tombol False, False, True
 End If
End Sub

Private Sub kosong()

Text5.Text = ""
Combo1.Text = ""
Combo2.Text = ""
tgl.Value = Format(Date, "MM/dd/yy")
tombol False, False, False
aktif False

End Sub

Private Sub aktif(p As Boolean)
Text5.Enabled = p
Combo1.Enabled = p
Combo2.Enabled = p
tgl.Enabled = p
If p = False Then
  warna = abu
Else
  warna = putih
End If
Text5.BackColor = warna
Combo1.BackColor = warna
Combo2.BackColor = warna
tgl.CalendarBackColor = warna
End Sub

Private Sub cek()
If Trim(Text1.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
  boleh = False
Else
  boleh = True
End If
End Sub

Private Sub Command1_Click()
cek
If boleh Then
 If MsgBox("Simpan Data?????", vbYesNo + vbQuestion, "Replace") = vbYes Then
  db.Execute ("insert into tkriminal values ('" & Text1.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & tgl.Value & "','" & Text5.Text & "')")
  Command4_Click
 End If
Else
  MsgBox "isi dulu", vbInformation + vbOKOnly, "informasi"
  Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
cek
If boleh Then
 If MsgBox("Perbaharui Data?????", vbYesNo + vbQuestion, "Update") = vbYes Then
  db.Execute ("update tkriminal set kd_pelaku='" & Combo1.Text & "',no_pas='" & Combo2.Text & "',tgl_kejadian= '" & tgl.Value & "',tkp='" & Text5.Text & "' where no_berkas='" & Text1.Text & "'")
  Command4_Click
 End If
Else
  MsgBox "isi dulu", vbInformation + vbOKOnly, "informasi"
  Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
If MsgBox("Hapus Data?????", vbYesNo + vbQuestion, "Delete") = vbYes Then
 db.Execute ("delete * from tkriminal where no_berkas ='" & Text1.Text & "'")
 Command4_Click
End If
End Sub

Private Sub Command4_Click()
Text1.Text = ""
kosong
tombol False, False, False
aktif False
Text1.SetFocus
End Sub

Private Sub Command5_Click()
rs.Close
rs.Close
'reck.Close
Unload Me
End Sub


