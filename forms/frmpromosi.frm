VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmpromosi 
   Caption         =   "Setting Promosi"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "frmpromosi.frx":0000
      Top             =   2280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "frmpromosi.frx":0234
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtisi 
      Height          =   1335
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmpromosi.frx":02BE
      Top             =   600
      Width           =   3615
   End
End
Attribute VB_Name = "frmpromosi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If MsgBox("Simpan isi promosi?", vbYesNo, judul) = vbNo Then Exit Sub
SaveSetting "apotekbaleendah", "frmpromosi", "txtisi.text", txtisi.Text
End Sub

Private Sub Command2_Click()
MsgBox Len(txtisi.Text)
End Sub

Private Sub Form_Load()
isipromosi = GetSetting("apotekbaleendah", "frmpromosi", "txtisi.text", "")
txtisi.Text = isipromosi
Skinpath = App.Path & "\skin\galaxy.skn"
Skin1.LoadSkin Skinpath
Skin1.ApplySkin Me.hwnd
End Sub

