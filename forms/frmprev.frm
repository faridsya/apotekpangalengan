VERSION 5.00
Begin VB.Form frmprev 
   Caption         =   "Preview Struk"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4230
   Icon            =   "frmprev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cetak"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
End
Attribute VB_Name = "frmprev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
transaksi.CetakData3
Unload Me
transaksi.SetFocus
transaksi.baru.SetFocus
End Sub

