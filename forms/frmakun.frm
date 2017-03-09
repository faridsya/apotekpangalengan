VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmakun 
   Caption         =   "Nomor akun tambahan"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtakun 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtnmr 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox cmbjns 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5655
      Left            =   5280
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9975
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No.Akun"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Akun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jenis"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "frmakun.frx":0000
      Top             =   3840
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "frmakun.frx":0234
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "frmakun.frx":02A6
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "frmakun.frx":0316
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmakun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbjns_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command1_Click()
If cmbjns.Text <> "Beban" Then
Set RS = New Recordset
RS.Open "select jns,urut from akun where jns2='" & cmbjns.Text & "'", jual, adOpenStatic, adLockOptimistic
jns = RS!jns
urut = RS!urut
Else
Set RS = New Recordset
RS.Open "select jns,urut from akun where jns2='" & cmbjns.Text & "'", jual, adOpenStatic, adLockOptimistic
jns = RS!jns
urut = "10"

End If
Set RS = New Recordset
RS.Open "select no_akun from akun where no_akun='" & txtnmr.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nomor akun telah terdaftar", vbCritical, judul
Exit Sub
End If
Set RS = New Recordset
RS.Open "select nama_akun from akun where nama_akun='" & txtakun.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Nama akun telah terdaftar", vbCritical, judul
Exit Sub
End If
If MsgBox("Simpan akun?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "insert into akun values('" & txtnmr.Text & "','" & txtakun.Text & "','" & jns & "','" & cmbjns.Text & "','tambahan','" & urut & "')"
MsgBox "Data berhasil disimpan", vbInformation, judul
dbgrid

End Sub

Private Sub Command2_Click()
Set RS = New Recordset
RS.Open "select no_akun from jurnal where no_akun='" & lv1.SelectedItem.SubItems(1) & "'", jual, adOpenStatic, adLockOptimistic
If Not RS.EOF Then
MsgBox "Tidak dapat dihapus,sudah ada transaksi!", vbCritical, judul
Exit Sub
End If
If MsgBox("Hapus akun", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from akun where no_akun='" & lv1.SelectedItem.SubItems(1) & "'"
dbgrid

MsgBox "Data berhasil dihapus", vbInformation, judul

End Sub

Private Sub Command3_Click()
dbgrid
End Sub

Private Sub Form_Load()
dbgrid
Set RS = New Recordset
RS.Open "select jns2 from akun where urut!=0 group by jns2 order by jns2", jual, adOpenStatic, adLockOptimistic
RS.MoveFirst
Do While Not RS.EOF
cmbjns.AddItem RS!jns2
RS.MoveNext
Loop
Skinpath = App.Path & "\skin\triton.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Sub dbgrid()
lv1.ListItems.Clear
  Set RS = New Recordset
  RS.Open "Select * from akun where ktr!='awal' order by no_akun", jual, adOpenStatic, adLockOptimistic
  
If RS.EOF Then Exit Sub
  With RS

    .MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![no_akun]
        l.SubItems(2) = ![nama_akun]
        l.SubItems(3) = !jns2

                

    .MoveNext
    Loop

End With

End Sub


