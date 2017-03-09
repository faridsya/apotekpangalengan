VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmsupp 
   Caption         =   "Data suppplier"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7440
      OleObjectBlob   =   "frmsupp.frx":0000
      Top             =   6720
   End
   Begin MSComctlLib.ListView lvplg 
      Height          =   6015
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10610
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Id supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama "
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact person"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Alamat"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "No.Telp"
         Object.Width           =   2540
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmsupp.frx":0234
      TabIndex        =   2
      Top             =   6360
      Width           =   2295
   End
   Begin XPControls.XPText txtcrp 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   6360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmsupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbgridplg
    Skinpath = App.Path & "\skin\mac.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Sub dbgridplg()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblsupplier order by supplier"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_supplier]
        l.SubItems(2) = ![Supplier]
                                l.SubItems(3) = ![Kontak_person]

                l.SubItems(4) = ![alamat]
                
                l.SubItems(5) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub
Sub dbgridplg2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblsupplier where id_supplier like '" & txtcrp.Text & "%' or supplier like '%" & txtcrp.Text & "%' order by supplier"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lvplg.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvplg.ListItems.Add(, , lvplg.ListItems.count + 1)
        l.SubItems(1) = ![id_supplier]
        l.SubItems(2) = ![Supplier]
                                l.SubItems(3) = ![Kontak_person]

                l.SubItems(4) = ![alamat]
                
                l.SubItems(5) = ![no_telp]

    .MoveNext
    Loop
End With


End Sub
Private Sub txtcrp_KeyPress(KeyAscii As Integer)
If lvplg.ListItems.count = 0 Then Exit Sub
If KeyAscii = 13 Then
lvplg.SetFocus
End If
End Sub
Private Sub lvplg_DblClick()
On Error Resume Next
Barang.txtids.Text = lvplg.SelectedItem.SubItems(1)
Barang.txtids_KeyPress (13)
Barang.Show
Unload Me
End Sub
Private Sub lvplg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lvplg_DblClick
End If
End Sub

Private Sub txtcrp_Change()
dbgridplg2

End Sub
