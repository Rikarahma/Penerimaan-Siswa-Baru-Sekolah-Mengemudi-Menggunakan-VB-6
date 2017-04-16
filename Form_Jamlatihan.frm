VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Jamlatihan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MASTER JAM LATIHAN ::."
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton btnCari 
         Caption         =   "CARI KODE"
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
         Left            =   3600
         TabIndex        =   16
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox tcari 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   3480
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid gridlatihan 
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   3960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5318
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
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   2520
         Width           =   855
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
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   2520
         Width           =   855
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
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton btnUpdate 
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
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton btnSimpan 
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
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton btnTambah 
         Caption         =   "TAMBAH"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox cmbdurasi 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox cmbhari 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox tjamlatihan 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Text            =   "JAM"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox tkode 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblcoa 
         Caption         =   "#CARIORADD"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "DURASI LATIHAN"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "HARI LATIHAN"
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
         Top             =   1150
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "KODE JAM LATIHAN"
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
         Top             =   530
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form_Jamlatihan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call combohari
Call combodurasi
Call nonaktif_afterloadawal
Call tampil_grid
End Sub

Private Sub btnHapus_Click()
msgdel = MsgBox("Anda akan menghapus data ini ?", vbExclamation + vbYesNo, "Information")
If msgdel = vbYes Then
    Call koneksi_db
        datadurasilatihan.Open "DELETE FROM tbl_jam_latihan WHERE kode_jam_latihan = '" & tkode.Text & "'", konn
        MsgBox "Data berhasil dihapus.", vbInformation, "Information"
        Call tampil_grid
        Call nonaktif_aftercancelcari
End If
End Sub

Private Sub btnUpdate_Click()
If cmbhari.Text = "PILIH" Or cmbhari.Text = "" Then
    MsgBox "Harap pilih kolom hari latihan.", vbExclamation, "Information"
ElseIf cmbdurasi.Text = "PILIH" Or cmbdurasi.Text = "" Then
    MsgBox "Harap pilih kolom durasi latihan.", vbExclamation, "Information"
Else
    Call koneksi_db
    datadurasilatihan.Open "UPDATE tbl_jam_latihan SET hari = '" & cmbhari.Text & "', durasi = '" & cmbdurasi.Text & "' WHERE kode_jam_latihan = '" & tkode.Text & "'", konn
    MsgBox "Data berhasil diperbaharui.", vbInformation, "Information"
    Call tampil_grid
End If
End Sub

Private Sub btnCari_Click()
If tcari.Text = "" Then
    MsgBox "Maaf, kolom pencarian tidak boleh kosong.", vbCritical, "Information"
Else
    Call koneksi_db
    datadurasilatihan.Open "SELECT kode_jam_latihan, hari, durasi FROM tbl_jam_latihan WHERE kode_jam_latihan = '" & tcari.Text & "' ", konn
    If datadurasilatihan.EOF Then
        MsgBox "Maaf, data tidak ditemukan.", vbCritical, "Information"
    Else
        With datadurasilatihan
        cmbdurasi.Text = .Fields("durasi")
        cmbhari.Text = .Fields("hari")
        tkode.Text = .Fields("kode_jam_latihan")
        Call aktif_aftercarikode
        lblcoa.Caption = 2
        End With
    End If
End If
End Sub

Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub cmbdurasi_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub cmbhari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If Not (KeyAscii >= Asc("a") & Chr(13) _
        And KeyAscii <= Asc("z") & Chr(13) _
        Or (KeyAscii >= Asc("A") & Chr(13) _
            And KeyAscii <= Asc("Z") & Chr(13) _
            Or KeyAscii = vbKeyBack _
            Or KeyAscii = vbKeyDelete _
            Or KeyAscii = vbKeySpace)) Then
    Beep
    KeyAscii = 0
End If
End Sub

Private Sub btnTambah_Click()
Call aktif_afteradd
Call koneksi_db
datadurasilatihan.Open "SELECT kode_jam_latihan FROM tbl_jam_latihan ORDER BY id_jam_latihan DESC", konn
With datadurasilatihan
    If .BOF And .EOF Then
      tkode.Text = "JLSC" + "001"
    Else
       tkode.Text = "JLSC" + Right(Str(Val(Right(.Fields("kode_jam_latihan"), 3)) + 1001), 3)
    End If
End With
lblcoa.Caption = 1
End Sub

Private Sub btnSimpan_Click()
If cmbhari.Text = "PILIH" Then
    MsgBox "Hari latihan harus dipilih.", vbCritical, "Information"
ElseIf cmbdurasi.Text = "PILIH" Then
    MsgBox "Hari latihan harus dipilih.", vbCritical, "Information"
Else
    Call koneksi_db
    datadurasilatihan.Open "INSERT INTO tbl_jam_latihan (kode_jam_latihan, hari, durasi) VALUES ('" & tkode.Text & "','" & cmbhari.Text & "','" & cmbdurasi.Text & "')", konn
    MsgBox "Jam latihan berhasil ditambah.", vbInformation, "Information"
    Call nonaktif_aftercanceladd
    Call tampil_grid
End If
End Sub

Private Sub btnCancel_Click()
If lblcoa.Caption = 1 Then
    Call nonaktif_aftercanceladd
Else
    Call nonaktif_aftercancelcari
End If
End Sub

Sub nonaktif_afterhapus()
btnTambah.Enabled = True
btnUpdate.Enabled = False
btnCancel.Enabled = False
btnHapus.Visible = False
btnTutup.Visible = True
tcari.Text = ""
tkode.Text = ""
cmbhari.Text = "PILIH"
cmbdurasi.Text = "PILIH"
End Sub

Sub nonaktif_aftercancelcari()
tkode.Enabled = False
cmbhari.Enabled = False
cmbdurasi.Enabled = False
btnTambah.Enabled = True
btnUpdate.Enabled = False
btnTutup.Visible = True
btnCancel.Enabled = False
btnCancel.Visible = True
btnTutup.Enabled = True
cmbhari.BackColor = &H80000003
cmbdurasi.BackColor = &H80000003
tkode.Text = ""
cmbhari.Text = "PILIH"
cmbdurasi.Text = "PILIH"
tcari.Text = ""
lblcoa.Caption = ""
End Sub

Sub aktif_aftercarikode()
tkode.Enabled = False
cmbhari.Enabled = True
cmbdurasi.Enabled = True
btnTambah.Enabled = False
btnUpdate.Enabled = True
btnHapus.Visible = True
btnHapus.Enabled = True
btnTutup.Visible = False
btnCancel.Visible = True
btnCancel.Enabled = True
cmbhari.BackColor = &H80000005
cmbdurasi.BackColor = &H80000005
End Sub

Sub tampil_grid()
Call koneksi_db
datadurasilatihan.CursorLocation = adUseClient
datadurasilatihan.CursorType = adOpenKeyset
datadurasilatihan.LockType = adLockOptimistic
datadurasilatihan.Open "SELECT KODE_JAM_LATIHAN, HARI ,CONCAT(durasi,' JAM') AS DURASI FROM tbl_jam_latihan ORDER BY id_jam_latihan DESC", konn
Set gridlatihan.DataSource = datadurasilatihan
gridlatihan.Columns(0).Width = 2100
gridlatihan.Columns(1).Width = 1200
gridlatihan.Columns(2).Width = 1200
gridlatihan.AllowDelete = False
gridlatihan.AllowUpdate = False
gridlatihan.Refresh
End Sub

Sub nonaktif_aftercanceladd()
tkode.BackColor = &H80000003
cmbhari.BackColor = &H80000003
cmbdurasi.BackColor = &H80000003
btnTambah.Enabled = True
btnUpdate.Enabled = False
btnUpdate.Visible = True
btnSimpan.Visible = False
btnHapus.Enabled = False
btnCancel.Enabled = False
tkode.Text = ""
cmbhari.Text = "PILIH"
cmbdurasi.Text = "PILIH"
tcari.Text = ""
lblcoa.Caption = ""
tkode.Enabled = False
cmbhari.Enabled = False
cmbdurasi.Enabled = False
End Sub

Sub aktif_afteradd()
tkode.BackColor = &H80000003
cmbhari.BackColor = &H80000005
cmbdurasi.BackColor = &H80000005
btnTambah.Enabled = False
btnUpdate.Visible = False
btnSimpan.Visible = True
btnCancel.Enabled = True
tkode.Enabled = False
cmbhari.Enabled = True
cmbdurasi.Enabled = True
End Sub

Sub nonaktif_afterloadawal()
tkode.BackColor = &H80000003
cmbhari.BackColor = &H80000003
cmbdurasi.BackColor = &H80000003
btnUpdate.Enabled = False
btnCancel.Enabled = False
tkode.Text = ""
tkode.Enabled = False
cmbhari.Enabled = False
cmbdurasi.Enabled = False
End Sub

Sub combodurasi()
cmbdurasi.Text = "PILIH"
cmbdurasi.AddItem ("3")
cmbdurasi.AddItem ("6")
cmbdurasi.AddItem ("9")
cmbdurasi.AddItem ("12")
End Sub

Sub combohari()
cmbhari.Text = "PILIH"
cmbhari.AddItem ("SENIN")
cmbhari.AddItem ("SELASA")
cmbhari.AddItem ("RABU")
cmbhari.AddItem ("KAMIS")
cmbhari.AddItem ("JUMAT")
cmbhari.AddItem ("SABTU")
cmbhari.AddItem ("MINGGU")
End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call btnCari_Click
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
