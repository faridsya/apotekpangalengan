VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmbarcode 
   Caption         =   "Cetak Barcode"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   13170
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Buat Barcode"
      TabPicture(0)   =   "frmbarcode.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lvbrg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtjum"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Skin1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CrystalReport1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtmulai"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Picture1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Data barang"
      TabPicture(1)   =   "frmbarcode.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "lv"
      Tab(1).Control(2)=   "txtcari"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command2 
         Caption         =   "&Cetak"
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   840
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   720
         TabIndex        =   14
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtcari 
         Height          =   285
         Left            =   -72240
         TabIndex        =   9
         Top             =   6360
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Hapus"
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Bersihkan"
         Height          =   495
         Left            =   3480
         TabIndex        =   3
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak"
         Height          =   375
         Left            =   7320
         TabIndex        =   2
         Top             =   6120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtmulai 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   600
         Top             =   5040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   480
         OleObjectBlob   =   "frmbarcode.frx":0038
         Top             =   3120
      End
      Begin XPControls.XPText txtjum 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
      Begin MSComctlLib.ListView lvbrg 
         Height          =   5655
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Jumlah"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv 
         Height          =   5655
         Left            =   -74280
         TabIndex        =   8
         Top             =   600
         Width           =   11775
         _ExtentX        =   20770
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nama barang"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Kategori"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Satuan"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Stok"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Harga Jual"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "HArga jual Grosir"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   600
         Top             =   5640
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Nama barang"
         Height          =   255
         Left            =   -74280
         TabIndex        =   13
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Mulai dari"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Kode barang"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmbarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public Function SavePictureToDB(RS As adodb.Recordset, _
    sFileName As String)

    On Error GoTo errSimpan
    Dim oPict As StdPicture
    Set oPict = LoadPicture(sFileName)
    'jika gambar tida ditemukan

    Set adostream = New adodb.Stream
    adostream.Type = adTypeBinary
    adostream.Open
    adostream.LoadFromFile sFileName
      RS!barcod = adostream.Read
   ' Image1.Picture = LoadPicture(sFileName)
    adostream.Close
    SavePictureToDB = True

Exit Function
errSimpan:
    SavePictureToDB = False
End Function
Sub DrawBarcode2(ByVal bc_string As String, obj As Object)
'Thanks to someone on PSC to give me information about BarCode
Dim xpos!
Dim Y1!
Dim Y2!
Dim dw%
Dim Th!
Dim tw
Dim new_string$
    If bc_string = "" Then obj.cls: Exit Sub

Dim bc(90) As String
    bc(1) = "1 1221"

    bc(2) = "1 1221"
    bc(48) = "11 221"
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
    bc(65) = "211 12"
    bc(66) = "121 12"
    bc(67) = "221 11"
    bc(68) = "112 12"
    bc(69) = "212 11"
    bc(70) = "122 11"
    bc(71) = "111 22"
    bc(72) = "211 21"
    bc(73) = "121 21"
    bc(74) = "112 21"
    bc(75) = "2111 2"
    bc(76) = "1211 2"
    bc(77) = "2211 1"
    bc(78) = "1121 2"
    bc(79) = "2121 1"
    bc(80) = "1221 1"
    bc(81) = "1112 2"
    bc(82) = "2112 1"
    bc(83) = "1212 1"
    bc(84) = "1122 1"
    bc(85) = "2 1112"
    bc(86) = "1 2112"
    bc(87) = "2 2111"
    bc(88) = "1 1212"
    bc(89) = "2 1211"
    bc(90) = "1 2211"
    bc(32) = "1 2121"
    bc(35) = ""
    bc(36) = "1 1 1 11"
    bc(37) = "11 1 1 1"
    bc(43) = "1 11 1 1"
    bc(45) = "1 1122"
    bc(47) = "1 1 11 1"
    bc(46) = "2 1121"
    bc(64) = ""
    bc(42) = "1 1221"
    bc_string = UCase(bc_string)
    obj.ScaleMode = 3
    obj.cls
    obj.Picture = Nothing
    dw = CInt(obj.ScaleHeight / 40)
    If dw < 1 Then dw = 1
    Th = obj.TextHeight(bc_string)
    tw = obj.TextWidth(bc_string)
    new_string = Chr$(1) & bc_string & Chr$(2)
    Y1 = obj.ScaleTop + 2
    Y2 = obj.ScaleTop + obj.ScaleHeight * Th
    obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
    xpos = obj.ScaleLeft
    
    For n = 1 To Len(new_string)
        c = Asc(Mid$(new_string, n, 1))
        If c > 90 Then c = 0
        bc_pattern$ = bc(c)
        For I = 1 To Len(bc_pattern$)
            Select Case Mid$(bc_pattern$, I, 1)
                Case " "
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                Case "1"
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
                xpos = xpos + dw
                Case "2"
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
                xpos = xpos + 2 * dw
            End Select
        Next
    Next
    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
    xpos = xpos + dw
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = (obj.ScaleWidth - tw) / 2
    obj.CurrentY = Y2 + 0.1 * Th
    'obj.Print bc_string
    
    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
    xpos = xpos + dw
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = 0 '(obj.ScaleWidth - tw) / 2
    obj.CurrentY = 0 'Y2 - 3.25 * Th
    'obj.Print Text3.Text

End Sub


Sub dbgridbrg()
'On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblbarang order by deskripsi"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lv.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lv.ListItems.Add(, , lv.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![kategori]

                l.SubItems(4) = ![satuan]
                l.SubItems(5) = ![stok]
                l.SubItems(6) = Format(![harga_jual], "#,#")
  l.SubItems(7) = Format(![Harga_jual2], "#,#")

    .MoveNext
    Loop
End With


End Sub

Private Sub Command2_Click()
'On Error Resume Next
jual.Execute "delete from barcode"

For m = 1 To (val(txtmulai.Text) - 1)
    jual.Execute "insert into barcode values(null,null)"
Next
With lvbrg

 For I = 1 To lvbrg.ListItems.count
Text3.Text = lvbrg.ListItems(I).SubItems(2)
idp = 0
Text2.Text = lvbrg.ListItems(I).SubItems(1)
 SavePicture Picture1.Image, App.Path & "\" & lvbrg.ListItems(I).SubItems(1) & ".jpg"

    For j = 1 To .ListItems(I).SubItems(3)
    idp = j
    jual.Execute "insert into barcode(kode_brg,deskripsi,id) values('" & .ListItems(I).SubItems(1) & "','" & .ListItems(I).SubItems(2) & "'," & idp & ")"
    Set RS = New Recordset
RS.Open "select * from barcode where kode_brg='" & lvbrg.ListItems(I).SubItems(1) & "' and id=" & idp & "", jual, adOpenStatic, adLockOptimistic
If SavePictureToDB(RS, App.Path & "\" & lvbrg.ListItems(I).SubItems(1) & ".jpg") = True Then
       'MsgBox "simpan mahasiswabarberhasil"

    End If

    RS.Update

    Next j


   ' frmPrint2.Picture1(i - 1).Picture = Me.Picture1.Image
Next I
End With
Kill App.Path & "\*.jpg"
With CrystalReport1
  .Reset

  .ReportFileName = serperreport & "\barcode.rpt"
  .RetrieveDataFiles

  .WindowTitle = "Cetak Barcode"

        .WindowMinButton = False
        .WindowShowCancelBtn = True
        .WindowShowCloseBtn = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        '.WindowParentHandle = Mnutama.hwnd

        .WindowState = crptMaximized
  .Action = 1
  'Me.Hide
End With
Pesan:
If err.Description <> vbNullString Then
End If

End Sub

Private Sub Command3_Click()
lvbrg.ListItems.Clear

End Sub

Private Sub Command4_Click()
If Not lvbrg.SelectedItem Is Nothing Then
jual.Execute "delete from harga where kode_brg='" & lvbrg.SelectedItem.SubItems(1) & "'"

lvbrg.ListItems.Remove lvbrg.SelectedItem.Index
End If
End Sub

Private Sub Form_Load()
txtmulai.Text = "1"
    Skinpath = App.Path & "\skin\galaxy.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd
dbgridbrg
Tab1.Tab = 0
End Sub

Private Sub lv_DblClick()
If lv.ListItems.count = 0 Then Exit Sub
Tab1.Tab = 0
Text1.Text = lv.SelectedItem.SubItems(1)
txtjum.SetFocus
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lv_DblClick
End If
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
If Tab1.Tab = 1 Then
txtcari.Text = ""
txtcari.SetFocus
End If
End Sub
Sub dbgridcari()
On Error Resume Next

Set rstrans = New Recordset

stri = Replace(txtcari.Text, "'", "''")

sql = "select * from tblbarang where kode_brg like '" & stri & "%' or deskripsi like '%" & stri & "%' order by deskripsi"
Set rstrans = jual.Execute(sql)
Dim l As ListItem
lv.ListItems.Clear
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lv.ListItems.Add(, , lv.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![kategori]

                l.SubItems(4) = ![satuan]
                l.SubItems(5) = ![stok]
                l.SubItems(6) = Format(![harga_jual], "#,#")
  l.SubItems(7) = Format(![Harga_jual2], "#,#")

    .MoveNext
    Loop
End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtjum.SetFocus
End If
End Sub
Sub dbgrid()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblbarang where kode_brg='" & Text1.Text & "'"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , lvbrg.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
        l.SubItems(3) = txtjum.Text

                

    .MoveNext
    Loop
End With


 


End Sub

Private Sub Text2_Change()
    Call DrawBarcode2(Text2, Picture1)

End Sub

Private Sub txtcari_Change()
dbgridcari
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If lv.ListItems.count = 0 Then Exit Sub
lv.SetFocus
End If
End Sub

Private Sub txtjum_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

If KeyAscii = 13 Then
dbgrid
Text1.Text = ""
txtjum.Text = ""
Text1.SetFocus
End If
End Sub
