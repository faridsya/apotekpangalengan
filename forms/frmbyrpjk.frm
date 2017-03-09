VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbyrpjk 
   Caption         =   "Pembayaran Pajak"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab Tab1 
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pembayaran"
      TabPicture(0)   =   "frmbyrpjk.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tgl"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lv1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "SkinLabel1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Skin1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SkinLabel2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblttl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SkinLabel3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Check1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbsumber"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Data pembayaran"
      TabPicture(1)   =   "frmbyrpjk.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "lv2"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command2 
         Caption         =   "Ba&talkan"
         Height          =   375
         Left            =   -74280
         TabIndex        =   11
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Bayar"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   4560
         Width           =   1215
      End
      Begin VB.ComboBox cmbsumber 
         Height          =   315
         Left            =   5520
         TabIndex        =   6
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ceklis semua"
         Height          =   315
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   840
         OleObjectBlob   =   "frmbyrpjk.frx":0038
         TabIndex        =   1
         Top             =   4080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblttl 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "frmbyrpjk.frx":00B0
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3720
         OleObjectBlob   =   "frmbyrpjk.frx":010E
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   3720
         OleObjectBlob   =   "frmbyrpjk.frx":0176
         Top             =   4380
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4680
         OleObjectBlob   =   "frmbyrpjk.frx":03AA
         TabIndex        =   5
         Top             =   4080
         Width           =   855
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   3255
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No Faktur"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tgl faktur"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "PPN"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComCtl2.DTPicker tgl 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dddd, d MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   4080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   125632515
         CurrentDate     =   37623
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   3735
         Left            =   -74280
         TabIndex        =   10
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6588
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Id bayar"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tgl bayar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sumber"
            Object.Width           =   2469
         EndProperty
      End
   End
End
Attribute VB_Name = "frmbyrpjk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub dbgriddata()
On Error Resume Next

Set RS2 = New Recordset


sql = "select * from bayar_pajak  order by id_byrpjk desc"
Set RS2 = jual.Execute(sql)
Dim l As ListItem
lv2.ListItems.Clear
If RS2.RecordCount = 0 Then Exit Sub
With RS2
.MoveFirst
    Do While Not .EOF
     
        Set l = lv2.ListItems.Add(, , lv2.ListItems.count + 1)
        l.SubItems(1) = ![id_byrpjk]
        
        l.SubItems(2) = Format(![tgl_byr], "dd mmm yyyy")
        l.SubItems(3) = Format(![jum_byr], "#,#")
        l.SubItems(4) = ![byrdari]
        
       
               

    .MoveNext
    Loop
End With


End Sub

Sub dbgrid()
On Error Resume Next

Set RS2 = New Recordset


sql = "select * from penjualan where ppn>0 and id_byrpjk is null order by id_byrpjk"
Set RS2 = jual.Execute(sql)
Dim l As ListItem
lv1.ListItems.Clear
If RS2.RecordCount = 0 Then Exit Sub
With RS2
.MoveFirst
    Do While Not .EOF
     
        Set l = lv1.ListItems.Add(, , lv1.ListItems.count + 1)
        l.SubItems(1) = ![no_penjualan]
        
        l.SubItems(2) = Format(![tanggal], "dd mmm yyyy")
        l.SubItems(3) = Format(![total], "#,#")
        l.SubItems(4) = Format(![ppn], "#,#")
        
       
               

    .MoveNext
    Loop
End With


End Sub

Private Sub Check1_Click()
If Check1.Value = Checked Then
For I = 1 To lv1.ListItems.count
lv1.ListItems(I).Checked = True
Next I
Else
For I = 1 To lv1.ListItems.count
lv1.ListItems(I).Checked = False
Next I
End If
ttl
End Sub
Sub ttl()
lblttl.Caption = ""
sum = 0
For I = 1 To lv1.ListItems.count
If lv1.ListItems(I).Checked = True Then
sum = sum + Format(lv1.ListItems(I).SubItems(4), Number)
End If
Next
lblttl.Caption = Format(sum, "#,#")
End Sub

Private Sub Command1_Click()
If Format(lblttl.Caption, Number) <= 0 Then Exit Sub
If MsgBox("Bayar PPN pada faktur yang terpilih?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "insert into bayar_pajak values('','" & Format(tgl.Value, "yyyy-mm-dd") & "','" & Format(lblttl.Caption, Number) & "','" & cmbsumber.Text & "')"
Set RS = New Recordset
RS.Open "select max(id_byrpjk) id from bayar_pajak", jual, adOpenStatic, adLockOptimistic
id = RS!id
For j = 1 To lv1.ListItems.count
If lv1.ListItems(j).Checked = True Then
jual.Execute "update penjualan set id_byrpjk=" & id & " where no_penjualan='" & lv1.ListItems(j).SubItems(1) & "'"
End If
Next j
dbgrid
dbgriddata
MsgBox "Berhasil"
End Sub

Private Sub Command2_Click()
If lv2.ListItems.count = 0 Then Exit Sub
If MsgBox("Batalkan pembayaran?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "delete from bayar_pajak where id_byrpjk=" & lv2.SelectedItem.SubItems(1) & ""
MsgBox "Pembayaran berhasil dibatalkan", vbInformation, judul
dbgriddata
dbgrid
End Sub

Private Sub Form_Load()
cmbsumber.AddItem "Kas"
cmbsumber.AddItem "Bank"
cmbsumber.ListIndex = 0
dbgrid
dbgriddata
Tab1.Tab = 0
tgl.Value = Now
Skinpath = App.Path & "\skin\jade2.skn"
Skin1.LoadSkin Skinpath
Skin1.ApplySkin Me.hwnd
End Sub

Private Sub lv1_Click()
ttl
End Sub

