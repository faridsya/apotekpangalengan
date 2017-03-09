VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmjadwal 
   Caption         =   "Jadwal Shift"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10035
   Icon            =   "frmjadwal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbhari 
      Height          =   315
      Left            =   8040
      TabIndex        =   24
      Top             =   4320
      Width           =   1815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   6840
      OleObjectBlob   =   "frmjadwal.frx":324A
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "&Tambah"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "&Ubah"
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame frame 
      Caption         =   "Isi data"
      Height          =   2535
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "frmjadwal.frx":32BA
         TabIndex        =   16
         Top             =   2040
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmjadwal.frx":331E
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmjadwal.frx":338C
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin MSComCtl2.DTPicker jam2 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125698050
         CurrentDate     =   41906
      End
      Begin MSComCtl2.DTPicker jam1 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125698050
         CurrentDate     =   41906
      End
      Begin VB.ComboBox cmbuser 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CheckBox cek 
         Caption         =   "Sabtu"
         Height          =   255
         Index           =   6
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox cek 
         Caption         =   "Kamis"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox cek 
         Caption         =   "Rabu"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox cek 
         Caption         =   "Jumat"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox cek 
         Caption         =   "Selasa"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox cek 
         Caption         =   "Senin"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox cek 
         Caption         =   "Minggu"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Semua hari"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      Picture         =   "frmjadwal.frx":33F2
      ScaleHeight     =   975
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   2280
         OleObjectBlob   =   "frmjadwal.frx":663C
         Top             =   720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Shift"
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
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2895
      Left            =   4800
      TabIndex        =   20
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Jam mulai"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jam akhir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HAri"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmjadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id As Integer
Private Sub Check1_Click()
If Check1.Value = Checked Then
For I = 0 To 6
cek(I).Value = Checked
Next I
Else
For I = 0 To 6

cek(I).Value = Unchecked
Next I

End If
End Sub
Private Sub awal()
frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True

End Sub

Private Sub cmbhari_Click()
dbgridcari
End Sub

Private Sub cmdbatal_Click()
awal
dbgrid
kosong
cmdtambah.SetFocus


End Sub
Sub teks()
idi = lv1.SelectedItem.SubItems(1)
cmbuser.Text = lv1.SelectedItem.SubItems(5)
jam1.Value = lv1.SelectedItem.SubItems(2)
jam2.Value = lv1.SelectedItem.SubItems(3)
s = lv1.SelectedItem.SubItems(4)
  s = StrConv(s, vbUnicode)
  pormat = Split(s, vbNullChar)
For I = 0 To (UBound(pormat) - 1)
cek(I).Value = IIf(pormat(I) = "1", Checked, Unchecked)
Next I
End Sub

Private Sub cmdedit_Click()
If lv1.ListItems.count = 0 Then Exit Sub

Edit = True
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, "Penjualan"
teks
frame.Enabled = True
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End Sub

Private Sub cmdhapus_Click()
If MsgBox("Hapus data?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from jadwalshift where id='" & lv1.SelectedItem.SubItems(1) & "'"
MsgBox "Data berhasil dihapus!", vbInformation, judul
dbgrid

End Sub

Private Sub Cmdsimpan_Click()
Dim nilai As String
If cmbuser.Text = "" Then
MsgBox "Data belum lengkap", vbCritical, judul
Exit Sub
End If
nilai = ""
For j = 0 To 6
If cek(j).Value = Checked Then
nil = "1"
Else
nil = "0"
End If
nilai = nilai + nil

Next j
MsgBox nilai
If Edit = False Then
If MsgBox("Simpan data?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "insert into jadwalshift values('','" & Format(jam1.Value, "hh:mm:ss") & "','" & Format(jam2.Value, "hh:mm:ss") & "','" & nilai & "','" & cmbuser & "')"
dbgrid
If MsgBox("Data berhasil disimpan,tambah data?", vbYesNo, judul) = vbYes Then
cmdtambah_Click
Else
awal
End If

Else
    If MsgBox("Rubah data?", vbYesNo, judul) = vbNo Then Exit Sub


    
    dbkonek.Execute "update jadwalshift set mulai='" & Format(jam1.Value, "hh:mm:ss") & "',akhir='" & Format(jam2.Value, "hh:mm:ss") & "',hari='" & nilai & "',username='" & cmbuser.Text & "' where id='" & idi & "'"
    MsgBox "Data berhasil dirubah", vbInformation, judul
    dbgrid
    awal
    
End If
End Sub
Sub dbgrid()
On Error Resume Next

Set RS2 = New Recordset


sql = "select * from jadwalshift order by id"
Set RS2 = jual.Execute(sql)
Dim l As ListItem
lv1.ListItems.Clear
If RS2.RecordCount = 0 Then Exit Sub
With RS2
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![id]
        
        l.SubItems(2) = Format(![mulai], "hh:mm:ss")
        l.SubItems(3) = Format(![akhir], "hh:mm:ss")
        l.SubItems(4) = ![hari]
        
        l.SubItems(5) = ![UserName]
        
               

    .MoveNext
    Loop
End With


End Sub
Sub dbgridcari()
On Error Resume Next

Set RS2 = New Recordset

If cmbhari.Text = "Semua" Then
dbgrid
Exit Sub
Else
sql = "select * from jadwalshift where substring(hari," & cmbhari.ListIndex & ",1)='1' order by id"
End If
Set RS2 = jual.Execute(sql)
Dim l As ListItem
lv1.ListItems.Clear
If RS2.RecordCount = 0 Then Exit Sub
With RS2
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![id]
        
        l.SubItems(2) = Format(![mulai], "hh:mm:ss")
        l.SubItems(3) = Format(![akhir], "hh:mm:ss")
        l.SubItems(4) = ![hari]
        
        l.SubItems(5) = ![UserName]
        
               

    .MoveNext
    Loop
End With


End Sub

Private Sub cmdtambah_Click()
On Error Resume Next
Edit = False
tambah
kosong

Check1.SetFocus
End Sub
Sub kosong()
For I = 0 To 6

cek(I).Value = Unchecked
Next I
cmbuser.Text = ""
End Sub
Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
Cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False

End Sub

Private Sub Form_Load()
lisuser
awal
dbgrid
cmbhari.AddItem "Semua"

cmbhari.AddItem "Minggu"
cmbhari.AddItem "Senin"
cmbhari.AddItem "Selasa"
cmbhari.AddItem "Rabu"
cmbhari.AddItem "Kamis"
cmbhari.AddItem "Jumat"
cmbhari.AddItem "Sabtu"
End Sub
Sub lisuser()
Set RS = New Recordset
RS.Open "select username from pengguna order by username", jual, adOpenStatic, adLockOptimistic
If RS.EOF Then Exit Sub
RS.MoveFirst
Do While Not RS.EOF
cmbuser.AddItem RS!UserName
RS.MoveNext
Loop
End Sub

Private Sub jam1_Change()
If jam1.Value >= jam2.Value Then
jam2.Value = jam1.Value
End If
End Sub

Private Sub lv1_Click()
teks
    
frame.Enabled = False
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
'Cmdedit.SetFocus
Cmdsimpan.Enabled = False
End Sub
