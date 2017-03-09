VERSION 5.00
Begin VB.Form huruff 
   Caption         =   "Huruf awal "
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "huruf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin apotekbaleendah.ThemedButton ThemedButton2 
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&Keluar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "huruf.frx":0CCA
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin apotekbaleendah.ThemedButton ThemedButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Simpan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "huruf.frx":1264
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Masing-masing harus 2 huruf atau kosongkan"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "Huruf awal  no.pemesanan"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Huruf awal id supplier"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Huruf awal id pelanggan"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Huruf awal no.pembelian"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Huruf awal no.penjualan"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Huruf awal kode barang"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "huruff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Ketengah Me
Text1.Text = GetSetting("apotekbaleendah", "huruff", "text1.text", "")
Text2.Text = GetSetting("apotekbaleendah", "huruff", "text2.text", "")
text3.Text = GetSetting("apotekbaleendah", "huruff", "text3.text", "")
text4.Text = GetSetting("apotekbaleendah", "huruff", "text4.text", "")
text5.Text = GetSetting("apotekbaleendah", "huruff", "text5.text", "")
text6.Text = GetSetting("apotekbaleendah", "huruff", "text6.text", "")

End Sub





Private Sub ThemedButton1_Click()
If Text1.Text <> "" And Len(Text1.Text) <> 2 Then Exit Sub
If Text2.Text <> "" And Len(Text2.Text) <> 2 Then Exit Sub
If text3.Text <> "" And Len(text3.Text) <> 2 Then Exit Sub
If text4.Text <> "" And Len(text4.Text) <> 2 Then Exit Sub
If text5.Text <> "" And Len(text5.Text) <> 2 Then Exit Sub
If text6.Text <> "" And Len(text6.Text) <> 2 Then Exit Sub

SaveSetting "apotekbaleendah", "huruff", "text1.text", Text1.Text
SaveSetting "apotekbaleendah", "huruff", "text2.text", Text2.Text
SaveSetting "apotekbaleendah", "huruff", "text3.text", text3.Text
SaveSetting "apotekbaleendah", "huruff", "text4.text", text4.Text
SaveSetting "apotekbaleendah", "huruff", "text5.text", text5.Text
SaveSetting "apotekbaleendah", "huruff", "text6.text", text6.Text


hbrg = GetSetting("apotekbaleendah", "huruff", "text1.text", "")
hpju = GetSetting("apotekbaleendah", "huruff", "text2.text", "")
hpb = GetSetting("apotekbaleendah", "huruff", "text3.text", "")
hcus = GetSetting("apotekbaleendah", "huruff", "text4.text", "")
hsup = GetSetting("apotekbaleendah", "huruff", "text5.text", "")
hpo = GetSetting("apotekbaleendah", "huruff", "text6.text", "")

MsgBox "Huruf awal berhasil diubah"
End Sub

Private Sub ThemedButton2_Click()
Unload Me
End Sub
