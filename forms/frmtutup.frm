VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmtutup 
   Caption         =   "Tutup Shift"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel lblttl 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "frmtutup.frx":0000
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":005E
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   615
      Left            =   120
      OleObjectBlob   =   "frmtutup.frx":00C8
      TabIndex        =   12
      Top             =   4080
      Width           =   5415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":024A
      Top             =   3600
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblterima 
      Height          =   255
      Left            =   2520
      OleObjectBlob   =   "frmtutup.frx":047E
      TabIndex        =   11
      Top             =   1440
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblsaldo 
      Height          =   495
      Left            =   2520
      OleObjectBlob   =   "frmtutup.frx":04DC
      TabIndex        =   10
      Top             =   840
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblbeda 
      Height          =   255
      Left            =   2640
      OleObjectBlob   =   "frmtutup.frx":053A
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtril 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel lbluser 
      Height          =   375
      Left            =   2520
      OleObjectBlob   =   "frmtutup.frx":0598
      TabIndex        =   7
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tutup Shift"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":05F6
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":0662
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":06D6
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":075C
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "frmtutup.frx":07CE
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmtutup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim terima As Currency
Dim id As Integer
Private Sub proses()
Set RS = New Recordset
rsd.Open "select id,username u,saldoawal from transshift where status='y'", jual, adOpenStatic, adLockOptimistic
With rsd
lbluser.Caption = !u
lblsaldo.Caption = Format(!saldoawal, "#,#")
id = !id
End With
Set RS = New Recordset
RS.Open "select coalesce(sum(case when cash<=total then cash else total end),0) jum from penjualan where id_shift=" & id & "", jual, adOpenStatic, adLockPessimistic
lblterima.Caption = Format(RS!jum, "#,#")
terima = rsd!saldoawal + RS!jum
lblttl.Caption = Format(terima, "#,#")
End Sub

Private Sub Command1_Click()
If val(txtril.Text) = 0 Then
MsgBox "Isi jumlah uang riil", vbCritical
Exit Sub
End If
If MsgBox("Tutup shift?", vbYesNo, judul) = vbNo Then Exit Sub
jual.Execute "update transshift set status='n',akhir='" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',penerimaan=" & RS!jum & ",adanya=" & Format(txtril.Text, Number) & ",selisih=" & Format(val(lblbeda.Caption), Number) & " where id=" & id & ""
sip = False
MsgBox "Berhasil", vbInformation, judul
End Sub

Private Sub Form_Load()
proses
Skinpath = App.Path & "\skin\jade2.skn"
Skin1.LoadSkin Skinpath
Skin1.ApplySkin Me.hwnd
End Sub

Private Sub txtril_Change()
On Error Resume Next
lblbeda.Caption = Format(val(terima) - Format(txtril.Text, Number), "#,#")
End Sub

Private Sub txtril_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub
