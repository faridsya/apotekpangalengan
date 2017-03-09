VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form harga 
   Caption         =   "Cetak Label harga"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1920
      OleObjectBlob   =   "harga.frx":0000
      Top             =   6840
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   11760
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   240
      TabIndex        =   11
      Text            =   "E3108115"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cetak Barcode"
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
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
      Height          =   855
      Left            =   120
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Bersihkan"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Tambah"
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Kode barang"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Kategori"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cetak"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvbrg 
      Height          =   5655
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
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
         Text            =   "Harga"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "harga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub simpan()
jual.Execute "delete from harga"
For z = 1 To lvbrg.ListItems.count
sql = "insert into harga values('" & lvbrg.ListItems(z).SubItems(1) & "','" & Replace(lvbrg.ListItems(z).SubItems(2), "'", "''") & "','" & lvbrg.ListItems(z).SubItems(3) & "')"
jual.Execute (sql)



    Next z

End Sub
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
    Y1 = obj.ScaleTop + 13
    Y2 = obj.ScaleTop + obj.ScaleHeight - 1 * Th
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
    obj.Print bc_string
    
    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
    xpos = xpos + dw
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = 0 '(obj.ScaleWidth - tw) / 2
    obj.CurrentY = 0 'Y2 - 3.25 * Th
    obj.Print Text3.Text

End Sub

Private Sub cmdPrint_Click()
On Error Resume Next

For I = 1 To lvbrg.ListItems.count
Text3.Text = lvbrg.ListItems(I).SubItems(2)

Text2.Text = lvbrg.ListItems(I).SubItems(1)

    frmPrint2.Picture1(I - 1).Picture = Me.Picture1.Image
Next I
    Printer.CurrentX = 200
    Printer.CurrentY = 200
    frmPrint2.PrintForm
    Printer.PaperSize = vbPRPSLegal
    Printer.Copies = Text4.Text
    Printer.EndDoc
    Unload frmPrint2

End Sub

Private Sub Command1_Click()

With CrystalReport1
  .Password = Chr(10) & "tujuh"

  .ReportFileName = serperreport & "\harga.rpt"
  .RetrieveDataFiles

  .WindowTitle = "Laporan Data Barang"

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

Private Sub Command2_Click()
If Option1.Value = True Then
dbgrid
Else
Dbgrid2

End If
simpan

lvbrg.Refresh

End Sub
Sub dbgrid()
On Error Resume Next

Set rstrans = New Recordset

If Combo1.Text = "Semua" Then
sql = "select * from tblbarang order by deskripsi"

Else
sql = "select * from tblbarang where kategori='" & Combo1.Text & "' order by deskripsi"

End If
Set rstrans = jual.Execute(sql)
Dim l As ListItem
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , lvbrg.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![harga_jual]

                

    .MoveNext
    Loop
End With


 


End Sub
Sub Dbgrid2()
On Error Resume Next

Set rstrans = New Recordset


sql = "select * from tblbarang where kode_brg='" & Text1.Text & "' order by deskripsi"

Set rstrans = jual.Execute(sql)
Dim l As ListItem
If rstrans.RecordCount = 0 Then Exit Sub
With rstrans
.MoveFirst
    Do While Not .EOF
     
        Set l = lvbrg.ListItems.Add(, , lvbrg.ListItems.count + 1)
        l.SubItems(1) = ![kode_brg]
        l.SubItems(2) = ![deskripsi]
                                l.SubItems(3) = ![harga_jual]

                

    .MoveNext
    Loop
End With


 


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

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
ktgr
Ketengah Me
Option1.Value = True
Text2.Text = ""
    Skinpath = App.Path & "\skin\galaxy.skn"
    Skin1.LoadSkin Skinpath
    Skin1.ApplySkin Me.hwnd

End Sub
Private Sub ktgr()
On Error Resume Next
  Dim I As Long
  Dim j As Long

Combo1.Clear
Combo1.AddItem "Semua"
sql = "select kategori from tblbarang  group by kategori order by kode_brg"
Set rsbarang = New Recordset
Set rsbarang = jual.Execute(sql)
If Not rsbarang.EOF Then
rsbarang.MoveFirst
 Do While Not rsbarang.EOF
Combo1.AddItem rsbarang!kategori
rsbarang.MoveNext
 Loop
  End If
rsbarang.Close
    Combo1.Text = "Semua"

  End Sub

Private Sub Option2_Click()
Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
Option2.Value = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2_Click
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Text2_Change()
    Call DrawBarcode2(Text2, Picture1)

End Sub

