VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Datauser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MASTER DATA USER AKSES ::."
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDatauser 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton btnAdd 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   16
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton btnHapus 
         Caption         =   "HAPUS"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   15
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton btnTutup 
         Caption         =   "TUTUP"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton btnSimpan 
         Caption         =   "PERBAHARUI"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   13
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton btnTambah 
         Caption         =   "TAMBAH USER"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton btnCari 
         Caption         =   "CARI USERNAME"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   10
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox tcari 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   4080
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid griddatauser 
         Height          =   3375
         Left            =   240
         TabIndex        =   11
         Top             =   4560
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5953
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
               LCID            =   1033
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
               LCID            =   1033
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
      Begin VB.ComboBox cmblvl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form_Datauser.frx":0000
         Left            =   1800
         List            =   "Form_Datauser.frx":0002
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox tpass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox tnamauser 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox tusername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "LEVEL AKSES"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "NAMA USER"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1030
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   400
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form_Datauser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
Dim levl As Integer
If cmblvl.Text = "Administrator" Then
    levl = 1
Else
    levl = 2
End If
Call koneksi_db
datauser.Open "SELECT username FROM tbl_user WHERE username = '" & tusername.Text & "' ", konn
If Not datauser.EOF Then
    MsgBox "Username " & tusername.Text & " telah ada didatabase, harap gunakan username lain.", vbCritical, "Information"
    tusername.Text = ""
    tusername.SetFocus
ElseIf cmblvl.Text = "-- PILIH --" Then
    MsgBox "Harap pilih level akses user.", vbCritical, "Information"
    cmblvl.SetFocus
ElseIf tusername.Text = "" Then
    MsgBox "Username tidak boleh kosong, harap isi username.", vbCritical, "Information"
    tusername.SetFocus
ElseIf tnamauser.Text = "" Then
    MsgBox "Nama user tidak boleh kosong, harap isi nama user.", vbCritical, "Information"
    tnamauser.SetFocus
ElseIf tpass.Text = "" Then
    MsgBox "Password tidak boleh kosong, harap isi password.", vbCritical, "Information"
    tpass.SetFocus
Else
    Call koneksi_db
    datauser.Open "INSERT INTO tbl_user (namauser,username,katasandi,level) VALUES ('" & tnamauser.Text & "','" & tusername.Text & "','" & tpass.Text & "','" & levl & "')", konn
    MsgBox "Tambah user akses baru berhasil.", vbInformation, "Information"
    Call clear_box
    Call nonaktif_load
    'Call koneksi_db
    Call tampil_grid
End If

End Sub

Private Sub btncancel_Click()
Call nonaktif_load
btntutup.Visible = True
End Sub

Private Sub btncari_Click()
Call koneksi_db
datauser.Open "SELECT namauser, username, IF(level = 1,'Administrator','User Pengelola') AS level, katasandi FROM tbl_user WHERE username = '" & tcari.Text & "'", konn
If datauser.EOF Then
    MsgBox "Data user tidak ditemukan.", vbCritical, "Information"
    tcari.Text = ""
    tcari.SetFocus
    Call clear_box
Else
    With datauser
        tnamauser.Text = .Fields("namauser")
        tusername.Text = .Fields("username")
        tpass.Text = .Fields("katasandi")
        cmblvl.Text = .Fields("level")
        btnAdd.Visible = False
        btntutup.Visible = False
        btnhapus.Visible = True
        btnhapus.Enabled = True
        btnsimpan.Enabled = True
        btncancel.Enabled = True
        btntambah.Enabled = False
    End With
End If
End Sub

Private Sub btnHapus_Click()
msgdel = MsgBox("Anda akan menghapus user " + tusername.Text + " ?", vbCritical + vbYesNo, "Information")
If msgdel = vbYes Then
    Call koneksi_db
    datamobil.Open "DELETE FROM tbl_user WHERE username = '" & tusername.Text & "' ", konn
    MsgBox "Username " + tusername.Text + " berhasil dihapus.", vbInformation, "Information"
    Call nonaktif_load
    Call tampil_grid
End If
End Sub

Private Sub btnsimpan_Click()
Dim levl As Integer
If cmblvl.Text = "Administrator" Then
    levl = 1
Else
    levl = 2
End If
    If tnamauser.Text = "" Then
        MsgBox "Nama user tidak boleh kosong.", vbCritical, "Information"
    ElseIf tpass.Text = "" Then
        MsgBox "Password tidak boleh kosong.", vbCritical, "Information"
    ElseIf cmblvl.Text = "" Then
        MsgBox "Level Akses tidak boleh kosong.", vbCritical, "Information"
    Else
        Call koneksi_db
        datauser.Open "UPDATE tbl_user SET namauser = '" & tnamauser.Text & "', katasandi = '" & tpass.Text & "', level = '" & levl & "' WHERE username = '" & tusername.Text & "' ", konn
        MsgBox "Data user telah berhasil diperbaharui.", vbInformation, "Information"
        Call tampil_grid
        Call nonaktif_load
    End If
End Sub

Private Sub btntambah_Click()
Call aktif_add
btncancel.Enabled = True
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call nonaktif_load
Call tampil_grid
End Sub

Sub nonaktif_load()
tnamauser.Enabled = False
tusername.Enabled = False
tpass.Enabled = False
cmblvl.Enabled = False
btnsimpan.Enabled = False
btnhapus.Visible = False
tnamauser.BackColor = &H80000003
tusername.BackColor = &H80000003
tpass.BackColor = &H80000003
cmblvl.BackColor = &H80000003
btntambah.Visible = True
btntambah.Enabled = True
btncancel.Enabled = False
btncancel.Visible = True
btncari.Enabled = True
tcari.Enabled = True
btntutup.Visible = True
Call clear_box
tcari.Text = ""
cmblvl.Clear
btntambah.TabIndex = 1
btntutup.TabIndex = 2
btnAdd.Visible = False
End Sub

Sub aktif_add()
tnamauser.Enabled = True
tusername.Enabled = True
tpass.Enabled = True
cmblvl.Enabled = True
btnsimpan.Enabled = True
btnhapus.Visible = False
btncancel.Visible = True
btntambah.Enabled = False
btnAdd.Visible = True
tnamauser.BackColor = &H80000005
tusername.BackColor = &H80000005
tpass.BackColor = &H80000005
cmblvl.BackColor = &H80000005
tusername.SetFocus
cmblvl.Clear
cmblvl.AddItem ("Administrator")
cmblvl.AddItem ("User Pengelola")
cmblvl.Text = "-- PILIH --"
tpass.MaxLength = 10
tusername.MaxLength = 10
End Sub

Sub tampil_grid()
Call koneksi_db
datauser.CursorLocation = adUseClient
datauser.CursorType = adOpenKeyset
datauser.LockType = adLockOptimistic
datauser.Open "SELECT NAMAUSER, USERNAME, IF(level = 1,'Administrator','User Pengelola') AS USER_AKSES FROM tbl_user ORDER BY id_user DESC", konn
Set griddatauser.DataSource = datauser
griddatauser.Columns(0).Width = 1550
griddatauser.Columns(1).Width = 1550
griddatauser.Columns(2).Width = 1620
griddatauser.AllowDelete = False
griddatauser.AllowUpdate = False
griddatauser.Refresh
End Sub

Sub clear_box()
tusername.Text = ""
tnamauser.Text = ""
tpass.Text = ""
cmblvl.Text = ""
End Sub

Sub aktif_btncari()
btntambah.Enabled = False
btnsimpan.Visible = True
btnhapus.Enabled = True
btnhapus.Visible = True
btntutup.Visible = False
btncancel.Enabled = True
btncancel.Visible = True
btnAdd.Visible = False
btncari.Enabled = False
tcari.Enabled = False
tcari.BackColor = &H80000005
End Sub

Sub aktif_txtcari()
tnamauser.Enabled = True
tusername.Enabled = False
tpass.Enabled = True
cmblvl.Enabled = True
btnsimpan.Enabled = True
btnhapus.Visible = False
btncancel.Visible = True
btntambah.Enabled = False
btnAdd.Visible = True
tnamauser.BackColor = &H80000005
tusername.BackColor = &H8000000F
tpass.BackColor = &H80000005
cmblvl.BackColor = &H80000005
cmblvl.AddItem ("Administrator")
cmblvl.AddItem ("User Pengelola")
tnamauser.SetFocus
tpass.MaxLength = 10
tusername.MaxLength = 10
End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call btncari_Click
End If
End Sub
