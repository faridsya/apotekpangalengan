VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tanya 
   Caption         =   "Tanya"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4920
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   2280
      Width           =   1575
   End
   Begin apotekbaleendah.xFrame frame1 
      Height          =   3255
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      Caption         =   "Data pembayaran"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8.25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Begin XPControls.XPText txttotal 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin XPControls.XPText txttunai 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin XPControls.XPText txtgiro 
         Height          =   285
         Left            =   1680
         TabIndex        =   19
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin MSComCtl2.DTPicker tgl 
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   115146753
         CurrentDate     =   40735
      End
      Begin XPControls.XPText txtno 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin XPControls.XPText txtbank 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin XPControls.XPText txtkode 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":0000
         TabIndex        =   13
         Top             =   2880
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":0068
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":00D0
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":0146
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "tanya.frx":01BA
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":0226
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "tanya.frx":028C
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Confirm"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BANK"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KAS"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "BANK:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Masuk ke catatan?"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "tanya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "" Then Exit Sub
Set rsbank = New Recordset
rsbank.Open "Select* from bank where nama_bank='" & Combo1.Text & "'", jual, adOpenStatic, adLockOptimistic
If Not rsbank.EOF Then
idb = rsbank!kode_bank
txtkode.Text = idb
txtbank.Text = Combo1.Text
rsbank.Close
Else
MsgBox "Bank tidak terdaftar"
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka3(KeyAscii)

End Sub

Private Sub Command1_Click()
pilih = "KAS"
idb = "Kas"
Unload Me

End Sub

Private Sub Command2_Click()
Combo1.Visible = True
Label2.Visible = True
Command3.Visible = True
Frame1.Visible = True
tanya.Height = 7020
txtgiro.Text = pembayaran
txttotal.Text = pembayaran
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Then
MsgBox "Pilih Bank dulu"
Combo1.SetFocus
Exit Sub
Else
If val(txtgiro.Text) <> 0 And txtno.Text = "" Then
MsgBox "Masukkan no giro"
txtno.SetFocus
Else
jugir = val(txtgiro.Text)
junai = val(txttunai.Text)
pilih = "BANK"
gno = txtno.Text
gnom = txtgiro.Text
gtgl = tgl.Value
Unload Me
End If
End If

pilih = "BANK"

End Sub

Private Sub Command4_Click()
pilih = ""
Unload Me
End Sub

Private Sub Form_Load()
Ketengah Me
pilih = ""
tgl.Value = Now + 7
End Sub



Private Sub txtgiro_Change()
If val(txtgiro.Text) <= pembayaran Then
txttunai.Text = pembayaran - val(txtgiro.Text)
Else
MsgBox "Jangan melebihi pembayaran"
txtgiro.Text = ""
txtgiro.SetFocus
Exit Sub
End If

End Sub

Private Sub txtgiro_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub

Private Sub txttunai_Change()
If val(txttunai.Text) <= pembayaran Then
txtgiro.Text = pembayaran - val(txttunai.Text)
Else
MsgBox "Jangan melebihi pembayaran"
txttunai.Text = ""
txttunai.SetFocus
Exit Sub
End If

End Sub

Private Sub txttunai_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)

End Sub
