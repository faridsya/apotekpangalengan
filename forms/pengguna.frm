VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{C5743C1F-5CAB-11D6-82C2-000021B74250}#23.0#0"; "vbskpro.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{42A5133D-AF48-42E1-904C-D9C4E9F82ED5}#1.0#0"; "button.ocx"
Begin VB.Form pengguna 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Data Pengguna"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Lihat Lo&g"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Lacak Password"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   1270
      _ExtentY        =   1270
      BorderStyleViejo=   1
      NombreForm_ParaBorderStyleViejo=   "pengguna"
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDBCtls.DBList List1 
      Bindings        =   "pengguna.frx":0000
      Height          =   1320
      Left            =   3690
      TabIndex        =   0
      Top             =   11575
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   2328
      _Version        =   393216
      ForeColor       =   12582912
      ListField       =   "kdpengguna"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   8.25
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6660
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4950
      Width           =   1140
   End
   Begin XPControls.XPFrame Frame 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7011
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check1 
         Caption         =   "Semua"
         Height          =   210
         Left            =   1440
         TabIndex        =   38
         Top             =   1920
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Entry"
         Height          =   2115
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   5505
         Begin VB.CheckBox cek 
            Caption         =   "Laporan akunting"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   20
            Left            =   2040
            TabIndex        =   43
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox cek 
            Caption         =   "Proses teknisi"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   19
            Left            =   3840
            TabIndex        =   42
            Top             =   1800
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Pendaftaran servis"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   18
            Left            =   3960
            TabIndex        =   41
            Top             =   1920
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry Teknisi"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   17
            Left            =   3720
            TabIndex        =   40
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry Master servis"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   16
            Left            =   3960
            TabIndex        =   39
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Rubah data jual"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   15
            Left            =   3840
            TabIndex        =   37
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Hapus data"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   14
            Left            =   3840
            TabIndex        =   36
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Ambil database"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   13
            Left            =   3840
            TabIndex        =   35
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Laporan-laporan"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   10
            Left            =   3840
            TabIndex        =   34
            Top             =   120
            Width           =   1590
         End
         Begin VB.CheckBox cek 
            Caption         =   "Back up database"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   12
            Left            =   3840
            TabIndex        =   33
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Pengguna"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   11
            Left            =   3840
            TabIndex        =   32
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Transaksi Pembelian"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   5
            Left            =   2040
            TabIndex        =   31
            Top             =   120
            Width           =   1935
         End
         Begin VB.CheckBox cek 
            Caption         =   "Keuangan"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   9
            Left            =   2040
            TabIndex        =   30
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Retur Jual"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   8
            Left            =   2040
            TabIndex        =   17
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Transaksi Penjualan"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   6
            Left            =   2040
            TabIndex        =   16
            Top             =   360
            Width           =   2130
         End
         Begin VB.CheckBox cek 
            Caption         =   "Retur Beli"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   7
            Left            =   2040
            TabIndex        =   15
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry Supplier"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   4
            Left            =   0
            TabIndex        =   14
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry obat"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   30
            Width           =   1590
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry pelanggan"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Top             =   300
            Width           =   1590
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry Mutasi gudang"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   570
            Width           =   1815
         End
         Begin VB.CheckBox cek 
            Caption         =   "Entry Penyesuain stok"
            ForeColor       =   &H00FF0000&
            Height          =   330
            Index           =   3
            Left            =   0
            TabIndex        =   10
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.ComboBox bagian 
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   1440
         Width           =   2130
      End
      Begin VB.TextBox nama 
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   2100
      End
      Begin VB.TextBox pass 
         ForeColor       =   &H00C00000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   2085
      End
      Begin VB.TextBox id 
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   2085
      End
      Begin VB.Label Label5 
         Caption         =   "Hak akses :"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bagian"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1125
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pengguna"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Top             =   405
         Width           =   1455
      End
   End
   Begin Dacara_dcButton.dcButton cmdtambah 
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   397
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdedit 
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdhapus 
      Height          =   255
      Left            =   1920
      TabIndex        =   25
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdsimpan 
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BackColor       =   10591645
      ButtonStyle     =   2
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
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdbatal 
      Height          =   225
      Left            =   3720
      TabIndex        =   27
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   397
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Batal"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdcari 
      Height          =   255
      Left            =   4560
      TabIndex        =   28
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   10591645
      ButtonStyle     =   2
      Caption         =   "&Cari"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210816
   End
   Begin Dacara_dcButton.dcButton Cmdkeluar 
      Height          =   255
      Left            =   5400
      TabIndex        =   29
      Top             =   4200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BackColor       =   10591645
      ButtonStyle     =   2
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
      ForeColor       =   4210816
   End
End
Attribute VB_Name = "pengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As String

Private Sub bag()
On Error Resume Next
  Dim I As Long
  Dim j As Long

bagian.Clear
sql = "select * from pengguna group by bagian order by bagian"
Set rspengguna = jual.Execute(sql)
If Not rspengguna.EOF Then
rspengguna.MoveFirst
 Do While Not rspengguna.EOF
bagian.AddItem rspengguna!bagian
rspengguna.MoveNext
 Loop
  End If
  rspengguna.Close

End Sub



Private Sub Check1_Click()
If Check1.Value = Checked Then
For I = 0 To 20
cek(I).Value = Checked
Next I
Else
For I = 0 To 20

cek(I).Value = Unchecked
Next I

End If
End Sub

Private Sub cmdbatal_Click()
awal
End Sub

Private Sub Cmdcari_Click()
Dim kata As String

frame.Enabled = False
kata = InputBox("Masukkan username", "Cari...")
If StrPtr(kata) = 0 Then Exit Sub
If kata = "" Then Exit Sub
Cmdsimpan.Enabled = False
sql = "select* from pengguna where username='" & kata & "'"
Set RS = New Recordset
Set RS = jual.Execute(sql)
If Not RS.EOF Then
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
teks
dbgrid

Else
MsgBox "Tidak ada", vbOKOnly, judul
dbgrid
awal
End If

End Sub

Private Sub cmdedit_Click()
Edit = True
A = id.Text
MsgBox "Tinggal rubah saja Data yang akan diedit lalu tekan simpan", vbYesOnly, judul
frame.Enabled = True
Cmdsimpan.Enabled = True
Cmdbatal.Enabled = True
End Sub
Sub teks()

cmdtambah.Enabled = True
Set RS = New Recordset
sql = "select * from pengguna where username='" & dbgrid1.Columns(0).Text & "'"
Set RS = jual.Execute(sql)
id.Text = RS!UserName
DataString = RS!Password
            Translate
            pass.Text = Temp$
nama.Text = RS!nama
bagian.Text = RS!bagian
For I = 0 To 20
cek(I).Value = IIf(RS.Fields(I + 4) = True, vbChecked, vbUnchecked)
Next I
End Sub

Private Sub cmdhapus_Click()
If Not Mnutama.StatusBar1.Panels(2).Text = id.Text Then
If MsgBox("Are You Sure?", vbYesNo, "Hapus Data ") = vbYes Then
sql = "delete from pengguna where username='" & id.Text & "'"
jual.Execute (sql)
dbgrid
teks
MsgBox "Data telah berhasil dihapus", vbYesOnly, judul
End If
Else
MsgBox "Tidak dapat menghapus pengguna yang aktif", vbInformation
End If
End Sub

Private Sub cmdout_Click()
Unload Me

End Sub

Private Sub Cmdkeluar_Click()
Unload Me
End Sub

Private Sub Cmdsimpan_Click()
    If Not Edit Then
    If Len(pass.Text) <= 4 Then
MsgBox "Password harus lebih dari 4 karakter", , judul
pass.SetFocus
Exit Sub
End If
Set rspengguna = New Recordset
    rspengguna.Open "select * from pengguna where username='" & id.Text & "'", jual, adOpenDynamic, adLockPessimistic

        If rspengguna.EOF Then
        If databes = "Akses" Then
                      DataString = pass.Text
      Translate
    pass.Text = Temp$

jual.Execute "insert into pengguna values('" & id.Text & "','" & nama.Text & "','" & bagian.Text & "',('" & pass.Text & "'),'" & cek(0).Value & "','" & cek(1).Value & "','" & cek(2).Value & "','" & cek(3).Value & "','" & cek(4).Value & "','" & cek(5).Value & "','" & cek(6).Value & "','" & cek(7).Value & "','" & cek(8).Value & "','" & cek(9).Value & "','" & cek(10).Value & "','" & cek(11).Value & "','" & cek(12).Value & "','" & cek(13).Value & "','" & cek(14).Value & "','" & cek(15).Value & "','" & cek(15).Value & "')"
Else
jual.Execute "insert into pengguna values('" & id.Text & "','" & nama.Text & "','" & bagian.Text & "',md5('" & pass.Text & "'),'" & cek(0).Value & "','" & cek(1).Value & "','" & cek(2).Value & "','" & cek(3).Value & "','" & cek(4).Value & "','" & cek(5).Value & "','" & cek(6).Value & "','" & cek(7).Value & "','" & cek(8).Value & "','" & cek(9).Value & "','" & cek(10).Value & "','" & cek(11).Value & "','" & cek(12).Value & "','" & cek(13).Value & "','" & cek(14).Value & "','" & cek(15).Value & "','" & cek(16).Value & "','" & cek(17).Value & "','" & _
cek(18).Value & "','" & cek(19).Value & "','" & cek(20).Value & "')"

End If
      frame.Enabled = False
      rspengguna.Close
If MsgBox("Tambah data Lagi?", vbYesNo, "Tanya") = vbYes Then
cmdtambah_Click
Else
awal

End If
    Else
    MsgBox "Duplikasi data ", vbCritical, judul
End If
   
    Else
    jual.Execute "update pengguna set nama='" & nama.Text & "',bagian='" & bagian.Text & "',cek0='" & cek(0).Value & "',cek1='" & cek(1).Value & "',cek2='" & cek(2).Value & "',cek3='" & cek(3).Value & "',cek4='" & cek(4).Value & "',cek5='" & cek(5).Value & "',cek6='" & cek(6).Value & "',cek7='" & cek(7).Value & "',cek8='" & cek(8).Value & "',cek9='" & cek(9).Value & "',cek10='" & cek(10).Value & "',cek11='" & cek(11).Value & "',cek12='" & cek(12).Value & "',cek13='" & cek(13).Value & "',cek14='" & cek(14).Value & "',cek15='" & _
    cek(15).Value & "',cek16='" & cek(16).Value & "',cek17='" & cek(17).Value & "',cek18='" & cek(18).Value & "',cek19='" & cek(19).Value & "',cek20='" & cek(20).Value & "' where username='" & id.Text & "'"
        
        
        frame.Enabled = False

    End If
          dbgrid
End Sub

Private Sub Command1_Click()
  'Decrypt password dari tabel T_User
  MsgBox "UserID = " & id.Text & " " & Chr(13) & _
         "Password = " & pass.Text & "", vbInformation, "Konfirmasi Password"
MsgBox cek(0).Value
End Sub

Private Sub Command2_Click()
frmLog.Show 1
End Sub

Private Sub dbgrid1_Click()
teks
frame.Enabled = False
Cmdsimpan.Enabled = False
End Sub
Private Sub dbgrid1_GotFocus()
Cmdedit.Enabled = True
Cmdhapus.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)

    nutup Me
    
End Sub


Private Sub cmdtambah_Click()
tambah
kosong
bag

id.SetFocus

End Sub
Sub tambah()
Edit = False
Cmdsimpan.Enabled = True
cmdtambah.Enabled = False
frame.Enabled = True
Cmdbatal.Enabled = True
Cmdhapus.Enabled = False
Cmdedit.Enabled = False
Cmdcari.Enabled = False
dbgrid1.Enabled = False

End Sub
Sub kosong()
  Dim ctrl As Control

id.Text = ""
pass.Text = ""
nama.Text = ""
bagian.Text = ""
  For Each ctrl In pengguna.Controls
    If TypeOf ctrl Is CheckBox Then
      With ctrl
      .Value = False
      End With
    End If
  Next

End Sub
Sub dbgrid()
sql = "select username,nama,bagian from pengguna"
Set RS = jual.Execute(sql)
Set dbgrid1.DataSource = RS


End Sub

Private Sub Form_Load()

bag
awal
Ketengah Me
dbgrid

End Sub
Private Sub awal()
frame.Enabled = False
Cmdedit.Enabled = False
Cmdhapus.Enabled = False
Cmdsimpan.Enabled = False
Cmdbatal.Enabled = False
cmdtambah.Enabled = True
Cmdcari.Enabled = True
dbgrid1.Enabled = True

End Sub







Private Sub id_Change()
  Dim posisi As Integer
  posisi = id.SelStart
id.Text = AwalKataKapital(id.Text)
id.SelStart = posisi

End Sub

Private Sub pass_GotFocus()
pass.SelStart = 0
pass.SelLength = Len(pass)

End Sub

Private Sub pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
nama.SetFocus
End If
End Sub
Private Sub nama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
bagian.SetFocus
End If
End Sub





