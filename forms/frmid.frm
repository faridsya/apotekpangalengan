VERSION 5.00
Begin VB.Form frmid 
   Caption         =   "Aktivasi Program"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "frmid.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buka"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Sensitive Case"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Kata Kunci"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ID PC"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Hash As New MD5Hash
Private strFile As String
Private bytBlock() As Byte
Dim emde As String


Private Sub Command1_Click()
    bytBlock = StrConv(Text1.Text, vbFromUnicode)
    emde = LCase(Hash.HashBytes(bytBlock))
If Text2.Text = Right(emde, 6) + Mid(emde, 5, 4) + Mid(emde, 10, 3) Then
text3.Text = "True"
aktipasi = True
SaveSetting "apotekbaleendah", "frmid", "Text3.text", text3.Text

MsgBox "Benar,silahkan restart program", vbInformation, judul
End
Else
aktipasi = False
text3.Text = "False"
SaveSetting "apotekbaleendah", "frmid", "Text3.text", text3.Text

MsgBox "Kata kunci salah,silahkan mengulang", vbInformation, judul
Text2.SetFocus
End If
  text4.Visible = False
  Command2.Visible = False

End Sub

Private Sub Command2_Click()
'idpc = "bfebfbff000006f2"
    bytBlock = StrConv(idpc, vbFromUnicode)
    
    emde = LCase(Hash.HashBytes(bytBlock))

text4.Text = Right(emde, 6) + Mid(emde, 5, 4) + Mid(emde, 10, 3)
End Sub

Private Sub Command3_Click()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
 Dim ret As String
  SetTimer hwnd, NV_INPUTBOX, 10, AddressOf TimerProc
  ret = InputBox("Enter Password")
  If StrPtr(ret) = 0 Then Exit Sub
  If ret = "balesawala" Then
  text4.Visible = True
  Command2.Visible = True
End If
End If
End Sub

Private Sub Form_Load()
Text1.Text = idpc
text4.Visible = False
Command2.Visible = False
End Sub
