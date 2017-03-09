VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbuka 
   Caption         =   "BUKA SHIFT"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5880
      OleObjectBlob   =   "frmbuka.frx":0000
      Top             =   2760
   End
   Begin VB.ComboBox cmbhari 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdbuka 
      Caption         =   "&Buka Shift"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   5880
      OleObjectBlob   =   "frmbuka.frx":0234
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtsaldo 
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmbuka.frx":02A6
      TabIndex        =   1
      Top             =   3240
      Width           =   2175
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "frmbuka.frx":0336
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "frmbuka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim usernem As String

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

Private Sub cmbhari_Click()
dbgridcari
End Sub

Private Sub cmdbuka_Click()
If MsgBox("Buka shift untuk user " & usernem & "?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "insert into transshift(mulai,saldoawal,username,status) values('" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Format(txtsaldo.Text, Number) & "','" & usernem & "','y')"
Set RS = New Recordset
RS.Open "select max(id) kodesip from transshift", jual, adOpenStatic, adLockOptimistic
kodesip = RS!kodesip
sip = True
transaksi.Show
Unload Me
End Sub

Private Sub Form_Load()
dbgrid
cmbhari.AddItem "Semua"

cmbhari.AddItem "Minggu"
cmbhari.AddItem "Senin"
cmbhari.AddItem "Selasa"
cmbhari.AddItem "Rabu"
cmbhari.AddItem "Kamis"
cmbhari.AddItem "Jumat"
cmbhari.AddItem "Sabtu"
Skinpath = App.Path & "\skin\jade2.skn"
Skin1.LoadSkin Skinpath
Skin1.ApplySkin Me.hwnd
End Sub

Private Sub LV1_DblClick()
Dim ret As String
If lv1.ListItems.count = 0 Then Exit Sub
  SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
ret = InputBox("Masukkan Password")
  If StrPtr(ret) = 0 Then Exit Sub
  Set rspengguna = New Recordset
rspengguna.Open "select * from pengguna where username='" & lv1.SelectedItem.SubItems(5) & "' and password=md5('" & ret & "')", jual, adOpenDynamic, adLockPessimistic
If rspengguna.EOF Then
MsgBox "Password salah!", vbCritical, judul
cmdbuka.Enabled = False
Exit Sub
Else
txtsaldo.Text = "0"
usernem = lv1.SelectedItem.SubItems(5)
cmdbuka.Enabled = True
End If
End Sub
