VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistem Informasi Kriminalitas (File Pelaku Kejahatan)"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4725
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
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   735
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
      Left            =   3900
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3750
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
      Left            =   2400
      Picture         =   "Form1.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3750
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
      Left            =   900
      Picture         =   "Form1.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3750
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
      Left            =   1650
      Picture         =   "Form1.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3750
      Width           =   735
   End
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
      Left            =   3150
      Picture         =   "Form1.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3750
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Height          =   3795
      Left            =   360
      TabIndex        =   13
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "Form1.frx":15E4
         Left            =   1680
         List            =   "Form1.frx":15EE
         TabIndex        =   6
         Top             =   2100
         Width           =   1455
      End
      Begin VB.TextBox Text6 
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   3210
         Width           =   1455
      End
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
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2820
         Width           =   2115
      End
      Begin VB.TextBox Text4 
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
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   2430
         Width           =   2085
      End
      Begin VB.TextBox Text3 
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
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   990
         Width           =   2055
      End
      Begin VB.TextBox Text2 
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
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   2115
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
         TabIndex        =   1
         Top             =   210
         Width           =   825
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Height          =   315
         ItemData        =   "Form1.frx":1608
         Left            =   1680
         List            =   "Form1.frx":161B
         TabIndex        =   5
         Top             =   1770
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1380
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MM/dd/yy"
         Format          =   16580611
         CurrentDate     =   38140
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agama Pelaku"
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
         Left            =   420
         TabIndex        =   24
         Top             =   1770
         Width           =   1245
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kewarganegaraan"
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
         Left            =   240
         TabIndex        =   23
         Top             =   3210
         Width           =   1425
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pekerjaan"
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
         Left            =   660
         TabIndex        =   22
         Top             =   2430
         Width           =   1005
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alamat "
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
         Left            =   960
         TabIndex        =   21
         Top             =   2820
         Width           =   705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tempat Lahir"
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
         TabIndex        =   20
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal lahir"
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
         Left            =   480
         TabIndex        =   19
         Top             =   1380
         Width           =   1185
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
         Height          =   375
         Left            =   270
         TabIndex        =   18
         Top             =   210
         Width           =   1395
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
         Height          =   375
         Left            =   390
         TabIndex        =   17
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jenis Kelamin"
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
         Left            =   420
         TabIndex        =   16
         Top             =   2100
         Width           =   1245
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
      Left            =   2460
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2070
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As Recordset
Const putih = &HFFC0FF
Const abu = &HFFC0C0
Dim boleh As Boolean

Private Sub Combo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Command7_Click()
kode_oto
text1_KeyPress (13)
End Sub

Private Sub Form_Load()
'Set rs = db.OpenRecordset("select * from pelaku order by kd_pelaku where kd_pelaku='" & Text1.Text & "'")
tombol False, False, False
aktif False
End Sub

Private Sub tombol(A As Boolean, B As Boolean, c As Boolean)
  Command1.Enabled = A
  Command2.Enabled = B
  Command3.Enabled = c
End Sub
Sub kode_oto()
Dim j As Integer
Dim no As String
Set rsbarang = New Recordset
sql = "Select kd_pelaku from tpelaku order by kd_pelaku Desc"
Set rsbarang = db.Execute(sql)
If rsbarang.EOF = True Then
Text1.Text = "P-00001"
Else
j = val(Right(rsbarang(0), 5))
no = "P-" + Format(Str(j + 1), "00000")
Text1.Text = no

End If
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Text1.Text) = "" Then
    MsgBox "Isi Data Dulu"
Else
Set rs = New Recordset
 rs.Open "select * from tpelaku  where kd_pelaku='" & Text1.Text & "'", db, adOpenStatic, adLockOptimistic
 If rs.EOF Then
  aktif True
  tombol True, False, False
  text2.SetFocus
 Else
  cari
 End If
Set rs = Nothing
End If
End If
End Sub

Private Sub cari()
aktif True
 text2.Text = rs(1)
 Text3.Text = rs(2)
 tgl.Value = rs(3)
 Combo1.Text = rs(4)
 Combo2.Text = IIf(Left(rs(5), 1) = "L", "Laki-Laki", "Perempuan")
 Text4.Text = rs(6)
 Text5.Text = rs(7)
 Text6.Text = rs(8)
 If MsgBox("Data Sudah Ada..........,Update Data??", vbQuestion + vbYesNo, "Redundandy Input") = vbYes Then
  tombol False, True, False
 Else
  tombol False, False, True
 End If
End Sub

Private Sub kosong()
Text1.Text = ""
text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
tgl.Value = Format(Date, "MM/dd/yy")
Text1.SetFocus
End Sub

Private Sub aktif(p As Boolean)
text2.Enabled = p
Text3.Enabled = p
Text4.Enabled = p
Text5.Enabled = p
Text6.Enabled = p
Combo1.Enabled = p
Combo2.Enabled = p
tgl.Enabled = p
If p = False Then
  warna = abu
Else
  warna = putih
End If
text2.BackColor = warna
Text3.BackColor = warna
Text4.BackColor = warna
Text5.BackColor = warna
Text6.BackColor = warna
Combo1.BackColor = warna
Combo2.BackColor = warna
End Sub

Private Sub cek()
If Trim(text2.Text) = "" Or Trim(Text3.Text) = "" Or Trim(Text4.Text) = "" Or Trim(Text5.Text) = "" Or Trim(Text6.Text) = "" Or Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then
  boleh = False
Else
  boleh = True
End If
End Sub

Private Sub Command1_Click()
cek
If boleh Then
 If MsgBox("Simpan Data?????", vbYesNo + vbQuestion, "Replace") = vbYes Then
  db.Execute ("insert into tpelaku values ('" & Text1.Text & "','" & text2.Text & "','" & Text3.Text & "','" & tgl.Value & "','" & Combo1.Text & "','" & Left(Combo2.Text, 1) & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')")
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
  db.Execute ("update tpelaku set tgl_pelaku= '" & tgl.Value & "',nm_pelaku='" & text2.Text & "', t_lahir='" & Text3.Text & "',j_kelamin='" & Left(Combo2.Text, 1) & "',agama='" & Combo1.Text & "',kerja='" & Text4.Text & "',alm_pelaku='" & Text5.Text & "',kewarganegaraan='" & Text6.Text & "' where kd_pelaku='" & Text1.Text & "'")
  Command4_Click
 End If
Else
  MsgBox "isi dulu", vbInformation + vbOKOnly, "informasi"
  Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
If MsgBox("Hapus Data?????", vbYesNo + vbQuestion, "Delete") = vbYes Then
 db.Execute ("delete * from tpelaku where kd_pelaku ='" & Text1.Text & "'")
 Command4_Click
End If
End Sub

Private Sub Command4_Click()
kosong
aktif False
tombol False, False, False
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub

Private Sub tgl_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then  ' The ENTER key.
            SendKeys "{tab}"     ' Set the focus to the next control.
            KeyAscii = 0        ' Ignore this key.
         End If

End Sub
