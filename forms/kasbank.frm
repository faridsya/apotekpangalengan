VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form kasbank 
   Caption         =   "Deposit n transit"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Bank ke kas"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Kas ke bank"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox bayar 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   61997059
      CurrentDate     =   40299
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Ju 
      Caption         =   "Jumlah"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "kasbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bayar_KeyPress(KeyAscii As Integer)
KeyAscii = validasiAngka(KeyAscii)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
    jual.Execute "insert into keuangan(Tanggal,Keterangan,Pengeluaran) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Deposit and transit','" & val(bayar.Text) & "')"
    jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pemasukan2) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Deposit and transit','" & val(bayar.Text) & "')"
 Else
 If Option2.Value = True Then
    jual.Execute "insert into keuangan(Tanggal,Keterangan,Pemasukan) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Deposit and transit','" & val(bayar.Text) & "')"
    jual.Execute "insert into keuangan2(Tanggal2,Keterangan2,Pengeluaran2) values('" & Format(DTPicker2.Value, "YYYY-mm-dd") & "','Deposit and transit','" & val(bayar.Text) & "')"
End If
End If
    
    MsgBox "Berhasil", vbOKOnly, "Berhasil"
    bayar.Text = ""
    bayar.SetFocus

End Sub

Private Sub Form_Load()
Ketengah Me
DTPicker2.Value = Now
Option1.Value = True
End Sub
