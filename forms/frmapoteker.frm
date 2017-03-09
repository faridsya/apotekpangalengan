VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmapoteker 
   Caption         =   "Apoteker"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6915
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "frmapoteker.frx":0000
      Top             =   5160
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdcari 
      Caption         =   "&Cari"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&Edit"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Cmdkeluar 
      Caption         =   "&Keluar"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   11
      Top             =   0
      Width           =   8175
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data apoteker"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Isi data"
      TabPicture(0)   =   "frmapoteker.frx":0234
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frame"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data apoteker"
      TabPicture(1)   =   "frmapoteker.frx":0250
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LV1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtcari"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SkinLabel1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin apotekbaleendah.xFrame frame 
         Height          =   3495
         Left            =   -74640
         TabIndex        =   16
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6165
         Caption         =   ""
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontSize        =   8.25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeaderGradientBottom=   12611136
         Begin XPControls.XPText Text1 
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Top             =   840
            Width           =   3495
            _ExtentX        =   6165
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
         Begin XPControls.XPText Text2 
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   1320
            Width           =   3495
            _ExtentX        =   6165
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
         Begin XPControls.XPText Text3 
            Height          =   285
            Left            =   2040
            TabIndex        =   2
            Top             =   1800
            Width           =   3495
            _ExtentX        =   6165
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
         Begin XPControls.XPText Text4 
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   2280
            Width           =   3495
            _ExtentX        =   6165
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
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ID apoteker"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   21
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   3
            Left            =   360
            TabIndex        =   20
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Telepon"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   7
            Left            =   360
            TabIndex        =   19
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   18
            Top             =   1800
            Width           =   1815
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   360
         OleObjectBlob   =   "frmapoteker.frx":026C
         TabIndex        =   15
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   3720
         Width           =   3375
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3135
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
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
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id apoteker"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama apoteker"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Alamat "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "No.Telp"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmapoteker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdbatal_Click()
awal
dbgrid
kosong
cmdtambah.SetFocus
End Sub

Private Sub Cmdcari_Click()
Tab1.Tab = 1
txtcari.SetFocus

End Sub

Private Sub Cmdkeluar_Click()
Unload Me

End Sub

Sub teks()

cmdtambah.Enabled = True
With LV1.SelectedItem
Text1.Text = .SubItems(1)
Text2.Text = .SubItems(2)
text3.Text = .SubItems(3)
text4.Text = .SubItems(4)
End With
End Sub
Private Sub cmdedit_Click()
Tab1.Tab = 0
Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
frame.Enabled = True

Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data") = vbYes Then
sql = "delete from apoteker where id_apoteker='" & Text1.Text & "'"
jual.Execute (sql)
dbgrid
kosong
MsgBox "Data telah berhasil dihapus", vbYesOnly, "Penjualan"
End If

End Sub

Private Sub Cmdsimpan_Click()
On Error GoTo erol
stri = Replace(Text2.Text, "'", "''")

Set rssupp = New Recordset
A = Text1.Text
rssupp.Open "select * from apoteker where id_apoteker='" & A & "'", jual, adOpenStatic, adLockOptimistic



If Edit = False Then
Set RS = New Recordset
RS.Open "select nama from apoteker where nama='" & stri & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama apoteker sudah ada,bedakan walau 1 huruf", vbInformation, judul
Text2.SetFocus
Exit Sub
End If
jual.Execute "insert into apoteker values('" & Text1.Text & "','" & stri & "','" & text3.Text & "','" & text4.Text & "')"
If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal
End If
Else
Set RS = New Recordset
RS.Open "select nama from apoteker where nama='" & stri & "' and id_apoteker<>'" & Text1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama apoteker sudah ada,bedakan walai 1 huruf", vbInformation, judul
Text1.SetFocus
Exit Sub
End If

jual.Execute "update apoteker set nama='" & stri & "',alamat='" & text3.Text & "',no_telp='" & text4.Text & "' where id_apoteker='" & Text1.Text & "'"
frame.Enabled = False
MsgBox "Data berhasil diubah", vbInformation, judul
End If
dbgrid
erol:
If err.Description <> vbNullString Then
    MsgBox "Duplikasi data atau data belum lengkap", vbCritical, "Penjualan"
frame.Enabled = True
End If

End Sub

Private Sub cmdtambah_Click()
Edit = False
tambah
kosong
idoto
Tab1.Tab = 0
Text2.SetFocus

End Sub
Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
Cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
End Sub
Sub kosong()
Text1.Text = ""
Text2.Text = ""
text3.Text = ""
text4.Text = ""
End Sub
Sub idoto()
Dim j As Integer
Dim No As String
Set rssupp = New Recordset
sql = "Select id_apoteker from apoteker order by id_apoteker Desc"
Set rssupp = jual.Execute(sql)
If rssupp.EOF = True Then
Text1.Text = "Apt001"
Else
j = val(Right(rssupp(0), 3))
No = "Apt" + Format(Str(j + 1), "000")
Text1.Text = No
End If
End Sub




Private Sub Command1_Click()
Tab1.Tab = 2
End Sub

Private Sub Form_Load()

awal

Ketengah Me
dbgrid
Tab1.Tab = 0
Dim Arq As String
    Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

Arq = ReadINI(App.Path & "\Backup.ini", "Backup_Dir", "Backup_Dir")
End Sub
Private Sub awal()
frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True

End Sub
Sub dbgrid()
On Error Resume Next

Dim l As ListItem
LV1.ListItems.Clear
Set RS = New Recordset
RS.Open "Select * from apoteker order by nama", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![id_apoteker]
        l.SubItems(2) = ![nama]
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub
Sub Dbgrid2()
On Error Resume Next

Dim l As ListItem
LV1.ListItems.Clear

Set RS = New Recordset
RS.Open "Select * from apoteker where nama like '%" & txtcari.Text & "%' order by nama", jual, adOpenStatic, adLockOptimistic

If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![id_apoteker]
        l.SubItems(2) = ![nama]
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub

Private Sub lv1_Click()
If LV1.ListItems.count = 0 Then Exit Sub
teks
Cmdedit.Enabled = True
Cmdhapus.Enabled = True

End Sub

Private Sub LV1_DblClick()
If LV1.ListItems.count = 0 Then Exit Sub
Tab1.Tab = 0
teks
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
If LV1.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
LV1_DblClick
End If
End Sub

Private Sub lv2_DblClick()
lv2_KeyPress (13)
End Sub

Private Sub lv2_KeyPress(KeyAscii As Integer)
If lv2.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
Tab1.Tab = 0
cmbpoli.Text = lv2.SelectedItem.SubItems(1)
End If
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
txtcari.SetFocus
Cmdsimpan.Enabled = False
txtcari.Text = ""
Cmdcari.Enabled = True
Else
If Tab1.Tab = 2 Then
txtcari2.SetFocus
txtcari2.Text = ""
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
text4.SetFocus
End If
End Sub
Private Sub text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Cmdsimpan.SetFocus
End If

End Sub


Private Sub txtcari_Change()
Dbgrid2
End Sub


