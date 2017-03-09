VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmfoot 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Setting Promosi"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5850
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   14
      Top             =   0
      Width           =   8175
      Begin ACTIVESKINLibCtl.Skin Skin2 
         Left            =   6240
         OleObjectBlob   =   "frmfoot.frx":0000
         Top             =   120
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Promosi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   3015
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1680
      OleObjectBlob   =   "frmfoot.frx":0234
      Top             =   4200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Baris satu sampai 6 akan muncul dibawah struk"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 2"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 3"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 4"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 5"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 6"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Baris 1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmfoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


SaveSetting "apotekbaleendah", "frmfoot", "text1.text", Text1.Text
SaveSetting "apotekbaleendah", "frmfoot", "text2.text", Text2.Text
SaveSetting "apotekbaleendah", "frmfoot", "text3.text", text3.Text
SaveSetting "apotekbaleendah", "frmfoot", "text4.text", text4.Text
SaveSetting "apotekbaleendah", "frmfoot", "text5.text", text5.Text
SaveSetting "apotekbaleendah", "frmfoot", "text6.text", text6.Text


cttn1 = GetSetting("apotekbaleendah", "frmfoot", "text1.text", "")
cttn2 = GetSetting("apotekbaleendah", "frmfoot", "text2.text", "")
cttn3 = GetSetting("apotekbaleendah", "frmfoot", "text3.text", "")
cttn4 = GetSetting("apotekbaleendah", "frmfoot", "text4.text", "")
cttn5 = GetSetting("apotekbaleendah", "frmfoot", "text5.text", "")
cttn6 = GetSetting("apotekbaleendah", "frmfoot", "text6.text", "")

End Sub


Private Sub Form_Load()
Ketengah Me
Text1.Text = GetSetting("apotekbaleendah", "frmfoot", "text1.text", "")
Text2.Text = GetSetting("apotekbaleendah", "frmfoot", "text2.text", "")
text3.Text = GetSetting("apotekbaleendah", "frmfoot", "text3.text", "")
text4.Text = GetSetting("apotekbaleendah", "frmfoot", "text4.text", "")
text5.Text = GetSetting("apotekbaleendah", "frmfoot", "text5.text", "")
text6.Text = GetSetting("apotekbaleendah", "frmfoot", "text6.text", "")


End Sub
