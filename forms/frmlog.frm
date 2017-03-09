VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lihat File Log"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Hapus Log"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Kembali"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fungsi API ini untuk membantu membuat scroll-bar
'horizontal di ListBox jika ada data yang lebar-nya
'melebihi lebar ListBox yang sudah fix
Private Declare Function SendMessageByNum Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
wParam As Long, ByVal lparam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194  'Ini penentunya...

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
If MsgBox("Yakin menghapus semua log?", vbYesNo + vbQuestion, judul) = vbNo Then Exit Sub
SetAttr serper & "\Loglogin.log", vbNormal

Kill serper & "\Loglogin.log"
      List1.Clear
End Sub

Private Sub Form_Load()
    Dim sNamaFile As String
    Dim NextLine As String
    Dim FileNum As Integer
    Dim i As Integer
    On Error GoTo Pesan
    'Bersihkan dulu Listbox mula-mula
    List1.Clear
    'Tampung nama file dan lokasinya
    sNamaFile = serper & "\LogLogin.log"
    'Fungsi FreeFile meng-assign nomor yang unik ke
    'variabel Filenum, untuk menghindari "bentrokan"
    'dengan file yang sudah terbuka. Jadi, Windows
    'akan otomatis mengatasinya... hebat, bukan?
    FileNum = FreeFile
    'Buka file log dengan mode Input untuk menampilkan
    'isinya ke suatu variabel NextLine......
    Open sNamaFile For Input As FileNum
    'Ulangi sampai mencapai akhir file teks
    Do Until EOF(FileNum)
    'Baca satu baris dari file log ke variabel NextLine
    Line Input #FileNum, NextLine
      'Tambahkan ke dalam List1
      List1.AddItem NextLine
    Loop
    Close  'Tutup file log jika sudah selesai
    'Jika ada data di dalam file log, sorot item teratas
    If List1.ListCount > 0 Then List1.Selected(0) = True
    'Berikut di bawah ini untuk menambahkan scrollbar
    'horizontal di List1 jika lebar data melebihi dari
    'lebar List yang fix
    Static x As Long
    For i = 0 To List1.ListCount - 1
      'Jika nilai x masih lebih kecil dari lebar
      'listbox yang sedang aktif... (tentu saja
      'x selalu lebih kecil, karena ditambahkan
      'string kosong " " di belakangnya...!
      If x < TextWidth(List1.List(i) & " ") Then
         'Tambahkan spasi di belakangya agar lebih
         'leluasa lagi dilihat
         x = TextWidth(List1.List(i) & "    ")
         'Jika property ScaleMode = vbTwips
         If ScaleMode = vbTwips Then
            'Lakukan perhitungan yang sesuai
            x = x / Screen.TwipsPerPixelX
         End If
         'Set lebar paling maksimal di List1
         SendMessageByNum List1.hwnd, _
                          LB_SETHORIZONTALEXTENT, _
                          x, 0
      End If
    Next i 'Counter bertambah satu
    Exit Sub
Pesan:  'Jika error, tampilkan pesan........
      List1.Clear

End Sub
