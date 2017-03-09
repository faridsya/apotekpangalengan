VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstok 
   Caption         =   "Stok gudang"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmstok.frx":0000
      TabIndex        =   5
      Top             =   4800
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Daftarkan semua"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "frmstok.frx":0116
      TabIndex        =   3
      Top             =   5040
      Width           =   6375
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblket 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmstok.frx":020C
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7223
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
         Text            =   "Kode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gudang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kode brg"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jumlah"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3598
      EndProperty
   End
   Begin VB.Menu mnstokgdg 
      Caption         =   "Pilihan"
      Visible         =   0   'False
      Begin VB.Menu mninput 
         Caption         =   "&Daftarkan"
      End
   End
End
Attribute VB_Name = "frmstok"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If MsgBox("Daftarkan Semua", vbYesNo, judul) = vbNo Then Exit Sub
For I = 1 To lv1.ListItems.count
If lv1.ListItems(I).SubItems(5) = "Tidak terdaftar" Then
jual.Execute "insert into stokgudang values('" & lv1.ListItems(I).SubItems(1) & "','" & Barang.Text1.Text & "',0)"
End If
Next
dbgrid

End Sub

Private Sub Form_Load()
dbgrid
lblket.Caption = "Barang " & Barang.Text2.Text
End Sub
Sub dbgrid()

On Error Resume Next
Dim l As ListItem
lv1.ListItems.Clear

Set RS = New Recordset
RS.Open "Select g.kode_gudang,nama_gudang,coalesce(s.kode_brg,0) as kode_brg,jumlah from gudang g left join stokgudang s on g.kode_gudang=s.kode_gudang and s.kode_brg='" & Barang.Text1.Text & "'  order by nama_gudang", jual, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
With RS
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![kode_gudang]
              l.SubItems(2) = ![nama_gudang]
  l.SubItems(3) = ![kode_brg]
    l.SubItems(4) = IIf(IsNull(![jumlah]) = True, 0, ![jumlah])
      l.SubItems(5) = IIf(IsNull(![jumlah]) = True, "Tidak terdaftar", "Terdaftar")
  

    .MoveNext
    Loop
End With
End Sub

Private Sub lv1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If lv1.SelectedItem.SubItems(1) = "" Then Exit Sub
If lv1.ListItems.count = 0 Then Exit Sub
If Button = vbRightButton Then
    Me.PopupMenu mnstokgdg
   Barang.dbgrid1_Click
    End If

End Sub

Private Sub mninput_Click()
If lv1.SelectedItem.SubItems(5) = "Terdaftar" Then Exit Sub
 With lv1.SelectedItem
jual.Execute "insert into stokgudang values('" & .SubItems(1) & "','" & Barang.Text1.Text & "',0)"
If .SubItems(1) = "utama" Then
Barang.dbgrid
End If
End With
dbgrid


End Sub
