VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Datasiswa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MASTER DATA SISWA ::."
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "PENCARIAN DATA"
      BeginProperty Font 
         Name            =   "Ubuntu Mono"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   6120
      TabIndex        =   1
      Top             =   0
      Width           =   8655
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
         Left            =   4440
         TabIndex        =   25
         Top             =   330
         Width           =   1935
      End
      Begin VB.CommandButton btncari 
         Caption         =   "CARI NO INDUK"
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
         Left            =   6600
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid gridsiswa 
         Height          =   3615
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6376
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
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton btnbatal 
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
         Left            =   3120
         TabIndex        =   22
         Top             =   4080
         Visible         =   0   'False
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
         Left            =   4680
         TabIndex        =   21
         Top             =   4080
         Visible         =   0   'False
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
         Height          =   495
         Left            =   4680
         TabIndex        =   20
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton btncancel 
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
         Left            =   3120
         TabIndex        =   19
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton btnsimpan 
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
         Left            =   1440
         TabIndex        =   18
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton btnupdate 
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
         Left            =   1440
         TabIndex        =   17
         Top             =   4080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton btntambah 
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
         TabIndex        =   16
         Top             =   4080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox tjob 
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
         Left            =   2160
         TabIndex        =   15
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox temail 
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
         Left            =   2160
         TabIndex        =   14
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox ttelpon 
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
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2520
         Width           =   3495
      End
      Begin VB.TextBox tktp 
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
         Left            =   2160
         MaxLength       =   16
         TabIndex        =   12
         Top             =   2040
         Width           =   3495
      End
      Begin VB.TextBox talamat 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox tnama 
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
         Left            =   2160
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox tnis 
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
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label7 
         Caption         =   "PEKERJAAN SISWA"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3530
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "EMAIL SISWA"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "TELEPON SISWA"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "NO KTP SISWA"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2050
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "ALAMAT SISWA"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1400
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "NAMA SISWA"
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
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "NO INDUK SISWA"
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
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Datasiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancel_Click()
Call nonaktif_aftercanceladd
End Sub

Private Sub btncari_Click()
If tcari.Text = "" Then
    MsgBox "Kolom pencarian tidak boleh kosong.", vbCritical, "Information"
Else
   Call koneksi_db
    datasiswa.Open "SELECT noinduk_siswa, nama_siswa, alamat_siswa, ktp_siswa, telpon_siswa, email_siswa, pekerjaan_siswa FROM tbl_siswa WHERE noinduk_siswa = '" & tcari.Text & "' ", konn
    If datasiswa.EOF Then
        MsgBox "No induk siswa " + tcari.Text + " tidak ditemukan.", vbCritical, "Information"
    Else
        With datasiswa
            tnis.Text = .Fields("noinduk_siswa")
            tnama.Text = .Fields("nama_siswa")
            talamat.Text = .Fields("alamat_siswa")
            tktp.Text = .Fields("ktp_siswa")
            ttelpon.Text = .Fields("telpon_siswa")
            temail.Text = .Fields("email_siswa")
            tjob.Text = .Fields("pekerjaan_siswa")
            Call aktif_aftercari
        End With
    End If
End If
End Sub

Private Sub btnHapus_Click()
msgdel = MsgBox("Apa anda yakin ingin menghapus data ini ?", vbCritical + vbYesNo, "Information")
If msgdel = vbYes Then
    Call koneksi_db
    datasiswa.Open "DELETE FROM tbl_siswa WHERE noinduk_siswa = '" & tnis.Text & "' ", konn
    MsgBox "Data siswa " + tnis.Text + " berhasil dihapus.", vbInformation, "Information"
    Call tampil_grid
    Call nonaktif_aftercanceladd
End If
End Sub

Private Sub btnsimpan_Click()
If tnama.Text = "" Then
    MsgBox "Nama siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf talamat.Text = "" Then
    MsgBox "Alamat siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf tktp.Text = "" Then
    MsgBox "KTP siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf ttelpon.Text = "" Then
    MsgBox "Telepon siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf temail.Text = "" Then
    MsgBox "Email siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf tjob.Text = "" Then
    MsgBox "Pekerjaan siswa jika kosong harap beri tanda (-).", vbExclamation, "Information"
Else
    Call koneksi_db
    datasiswa.Open "INSERT INTO tbl_siswa (noinduk_siswa, nama_siswa, alamat_siswa, ktp_siswa, telpon_siswa, email_siswa, pekerjaan_siswa) VALUES ('" & tnis.Text & "','" & tnama.Text & "','" & talamat.Text & "','" & tktp.Text & "','" & ttelpon.Text & "','" & temail.Text & "','" & tjob.Text & "')", konn
    MsgBox "Data siswa berhasil ditambah.", vbInformation, "Information"
    Call tampil_grid
    Call nonaktif_aftercanceladd
End If
End Sub

Private Sub btntambah_Click()
Call tampil_grid
Call aktif_afteradd
Call koneksi_db
datasiswa.Open "SELECT noinduk_siswa FROM tbl_siswa ORDER BY noinduk_siswa DESC", konn
With datasiswa
    If .BOF And .EOF Then
      tnis.Text = "NSC" + "001"
    Else
       tnis.Text = "NSC" + Right(Str(Val(Right(.Fields("noinduk_siswa"), 3)) + 1001), 3)
    End If
End With
End Sub

Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub btnupdate_Click()
If tnama.Text = "" Then
    MsgBox "Nama siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf talamat.Text = "" Then
    MsgBox "Alamat siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf tktp.Text = "" Then
    MsgBox "KTP siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf ttelpon.Text = "" Then
    MsgBox "Telepon siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf temail.Text = "" Then
    MsgBox "Email siswa tidak boleh kosong.", vbExclamation, "Information"
ElseIf tjob.Text = "" Then
    MsgBox "Pekerjaan siswa jika kosong harap beri tanda (-).", vbExclamation, "Information"
Else
    Call koneksi_db
    datasiswa.Open "UPDATE tbl_siswa SET nama_siswa = '" & tnama.Text & "', alamat_siswa = '" & talamat.Text & "', ktp_siswa = '" & tktp.Text & "', telpon_siswa = '" & ttelpon.Text & "', email_siswa = '" & temail.Text & "', pekerjaan_siswa = '" & tjob.Text & "' WHERE noinduk_siswa = '" & tnis.Text & "' ", konn
    MsgBox "Data siswa berhasil diperbaharui.", vbInformation, "Information"
    Call tampil_grid
    Call nonaktif_aftercanceladd
End If
End Sub

Private Sub Form_Load()
Call nonaktif_formloaded
Call tampil_grid
End Sub

Sub aktif_aftercari()
tnis.Enabled = False
tnama.Enabled = True
talamat.Enabled = True
tktp.Enabled = True
ttelpon.Enabled = True
temail.Enabled = True
tjob.Enabled = True
tnis.BackColor = &H80000003
tnama.BackColor = &H80000005
talamat.BackColor = &H80000005
tktp.BackColor = &H80000005
ttelpon.BackColor = &H80000005
temail.BackColor = &H80000005
tjob.BackColor = &H80000005
btntambah.Visible = True
btnupdate.Visible = True
btncancel.Visible = True
btnHapus.Visible = True
btntambah.Enabled = False
btnupdate.Enabled = True
btncancel.Enabled = True
btnHapus.Enabled = True
End Sub

Sub nonaktif_formloaded()
tnis.Enabled = False
tnama.Enabled = False
talamat.Enabled = False
tktp.Enabled = False
ttelpon.Enabled = False
temail.Enabled = False
tjob.Enabled = False
tnis.BackColor = &H80000003
tnama.BackColor = &H80000003
talamat.BackColor = &H80000003
tktp.BackColor = &H80000003
ttelpon.BackColor = &H80000003
temail.BackColor = &H80000003
tjob.BackColor = &H80000003
btntambah.Visible = True
btnupdate.Visible = True
btncancel.Visible = True
btnTutup.Visible = True
btntambah.Enabled = True
btnupdate.Enabled = False
btncancel.Enabled = False
btnTutup.Enabled = True
End Sub

Sub aktif_afteradd()
tnis.Enabled = False
tnama.Enabled = True
talamat.Enabled = True
tktp.Enabled = True
ttelpon.Enabled = True
temail.Enabled = True
tjob.Enabled = True
tnis.BackColor = &H80000003
tnama.BackColor = &H80000005
talamat.BackColor = &H80000005
tktp.BackColor = &H80000005
ttelpon.BackColor = &H80000005
temail.BackColor = &H80000005
tjob.BackColor = &H80000005
btntambah.Visible = True
btnsimpan.Visible = True
btncancel.Visible = True
btnTutup.Visible = True
btntambah.Enabled = False
btnsimpan.Enabled = True
btncancel.Enabled = True
btnTutup.Enabled = True
End Sub

Sub nonaktif_aftercanceladd()
tnis.Enabled = False
tnama.Enabled = False
talamat.Enabled = False
tktp.Enabled = False
ttelpon.Enabled = False
temail.Enabled = False
tjob.Enabled = False
tnis.Text = ""
tnama.Text = ""
talamat.Text = ""
tktp.Text = ""
ttelpon.Text = ""
temail.Text = ""
tjob.Text = ""
tcari.Text = ""
tnis.BackColor = &H80000003
tnama.BackColor = &H80000003
talamat.BackColor = &H80000003
tktp.BackColor = &H80000003
ttelpon.BackColor = &H80000003
temail.BackColor = &H80000003
tjob.BackColor = &H80000003
btntambah.Visible = True
btnupdate.Visible = True
btncancel.Visible = True
btnTutup.Visible = True
btnHapus.Visible = False
btntambah.Enabled = True
btnupdate.Enabled = False
btncancel.Enabled = False
btnTutup.Enabled = True
btnsimpan.Visible = False
btntambah.TabIndex = 1
btnTutup.TabIndex = 2
btntambah.SetFocus
End Sub

Sub tampil_grid()
Call koneksi_db
datasiswa.CursorLocation = adUseClient
datasiswa.CursorType = adOpenKeyset
datasiswa.LockType = adLockOptimistic
datasiswa.Open "SELECT noinduk_siswa AS NO_INDUK, nama_siswa AS NAMA, ktp_siswa AS KTP, telpon_siswa AS TELEPON, email_siswa AS EMAIL FROM tbl_siswa ORDER BY id_siswa DESC", konn
Set gridsiswa.DataSource = datasiswa
gridsiswa.Columns(0).Width = 1200
gridsiswa.Columns(1).Width = 1800
gridsiswa.Columns(2).Width = 1600
gridsiswa.Columns(3).Width = 1700
gridsiswa.Columns(4).Width = 1800
gridsiswa.AllowAddNew = False
gridsiswa.AllowDelete = False
gridsiswa.AllowUpdate = False
End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tjob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnsimpan_Click
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tktp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnsimpan_Click
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub tnama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnsimpan_Click
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ttelpon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnsimpan_Click
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub
