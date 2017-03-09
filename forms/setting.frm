VERSION 5.00
Begin VB.Form setting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Default Printer"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option6 
      Caption         =   "Struk (lpt)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton option5 
      Caption         =   "Faktur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton option4 
      Caption         =   "Struk (usb)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Default"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis struk:"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Select a printer:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'To verify that the printer has been changed to the default, open your "Printers" folder
'in your control panel and select a printer and press the right mouse button to access
'the context menu. Take notice of "Set As Default" check mark for the selected printer
'when your run this app...

Private cSetPrinter As New cSetDfltPrinter

Private Sub Command1_Click()
    Dim sMsg As String
    Dim DeviceName As String
    
    If List1.SelCount = 1 Then
        DeviceName = List1.List(List1.ListIndex)
        If cSetPrinter.SetPrinterAsDefault(DeviceName) Then
            sMsg = DeviceName & " berhasil jadi printer utama."
        Else
            sMsg = DeviceName & " gagal jadi printer utama."
        End If
        MsgBox sMsg, vbExclamation, judul
    Else
        MsgBox "Please select a printer from the list.", vbInformation, App.Title
    End If
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Ketengah Me
    For I = 0 To Printers.count - 1
        List1.AddItem Printers(I).DeviceName
    Next I

End Sub

Private Sub option4_Click()
    SaveSetting "apotekbaleendah", "setting", "Option4.value", Option4.Value

End Sub

Private Sub option5_Click()
    SaveSetting "apotekbaleendah", "setting", "Option5.value", option5.Value

End Sub

Private Sub Option6_Click()
    SaveSetting "apotekbaleendah", "setting", "Option6.value", Option6.Value

End Sub
