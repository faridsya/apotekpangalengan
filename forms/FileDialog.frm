VERSION 5.00
Begin VB.Form frmFileDialog 
   Caption         =   "File Dialog"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Folder tujuan:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bLoading As Boolean
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    gsFileName = ""
    Unload Me
End Sub


Private Sub Dir1_Change()
    On Error GoTo Dir1_Change_Error
    
    txtPath = Dir1.Path
    
    
    If bLoading = False Then
    End If
Exit_Dir1_Change:
        
   Exit Sub
    
Dir1_Change_Error:
    
    #If gnDebug Then
        Stop
        Resume
    #End If

    Resume Exit_Dir1_Change
    
End Sub

Private Sub Drive1_Change()
    On Error GoTo Drive1_Change_Error
    
    Dir1.Path = Drive1.Drive
    gsDrive = Drive1.Drive
    
Exit_Drive1_Change:
        
   Exit Sub
    
Drive1_Change_Error:
    
    #If gnDebug Then
        Stop
        Resume
    #End If
    HandleError "Drive1_Change", err.Description, err.Number, gErrFormName

    Resume Exit_Drive1_Change

End Sub


Private Sub Form_Load()
Ketengah Me
    bLoading = True
    If gsDrive <> "" Then
        Drive1.Drive = gsDrive
        Dir1.Path = Drive1.Drive
    End If

    If gsPath <> "" Then
        Dir1.Path = gsPath
    End If
    bLoading = False
End Sub
