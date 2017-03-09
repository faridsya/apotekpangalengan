VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcrdok 
   Caption         =   "Pencarian Dokter"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8820
   Icon            =   "frmcrdok.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcari 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3735
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6588
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Id dokter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama dokter"
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
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Spesialis"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmcrdok.frx":324A
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmcrdok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbgrid
End Sub
Sub dbgrid()
On Error Resume Next
Dim l As ListItem
LV1.ListItems.Clear
Set RS = New Recordset
RS.Open "Select * from dokter order by nama", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![id_dokter]
        l.SubItems(2) = ![nama]
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![no_telp]
        l.SubItems(5) = ![spesialis]

    .MoveNext
    Loop
End With


End Sub
Sub Dbgrid2()
On Error Resume Next

Dim l As ListItem
LV1.ListItems.Clear

Set RS = New Recordset
RS.Open "Select * from dokter where nama like '%" & txtcari.Text & "%' order by nama", jual, adOpenStatic, adLockOptimistic

If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = LV1.ListItems.Add(, , LV1.ListItems.count + 1)
        l.SubItems(1) = ![id_dokter]
        l.SubItems(2) = ![nama]
        l.SubItems(3) = ![alamat]
        l.SubItems(4) = ![no_telp]
        l.SubItems(5) = ![spesialis]

    .MoveNext
    Loop
End With


End Sub

Private Sub LV1_DblClick()
If LV1.ListItems.count = 0 Then Exit Sub
transaksi.txtiddok.Text = LV1.SelectedItem.SubItems(1)
transaksi.txtkmsp.SetFocus
Unload Me
End Sub

Private Sub LV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
LV1_DblClick
End If
End Sub

Private Sub txtcari_Change()
Dbgrid2
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If LV1.ListItems.count = 0 Then Exit Sub
LV1.SetFocus
End If
End Sub
