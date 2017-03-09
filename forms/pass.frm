VERSION 5.00
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Begin VB.Form pass 
   Caption         =   "Ganti password"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   Icon            =   "pass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin vbskpro.Skinner Skinner1 
      Left            =   960
      Top             =   2520
      _ExtentX        =   1270
      _ExtentY        =   1270
      BorderStyleViejo=   2
      NombreForm_ParaBorderStyleViejo=   "pass"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Batal"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Password lama:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Password baru:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Ulangi password baru:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Len(Text2.Text) <= 4 Then
MsgBox "Password harus lebih dari 4 karakter", , judul
Text2.SetFocus
Exit Sub
End If

user = Mnutama.StatusBar1.Panels(2).Text
Set rspengguna = New Recordset
rspengguna.Open "select * from pengguna where username='" & user & "' and password=md5('" & Text1.Text & "')", jual, adOpenDynamic, adLockPessimistic
If rspengguna.EOF Then
MsgBox "Password salah!", vbCritical, judul
Exit Sub
End If
    If Text3.Text = Text2.Text Then
    If databes = "Akses" Then
    DataString = Text2.Text
            Translate
jual.Execute "update pengguna set password=('" & Temp$ & "') where username='" & user & "'"

Else
    
jual.Execute "update pengguna set password=md5('" & Text2.Text & "') where username='" & user & "'"
End If

MsgBox "Password berhasil diubah", vbInformation, judul
Unload Me
Else
MsgBox "Passsword ulangan tidak sesuai"
End If
rspengguna.Close

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Ketengah Me

End Sub

Private Sub Text2_GotFocus()
Text2.SelStart = 0
  Text2.SelLength = Len(Text2)

End Sub

