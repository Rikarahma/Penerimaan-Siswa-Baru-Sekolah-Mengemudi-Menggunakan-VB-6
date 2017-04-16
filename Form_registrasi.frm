VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Registrasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: FORM REGISTRASI  ::."
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "DATA TERSIMPAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   8160
      TabIndex        =   27
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton btntutup 
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
         Left            =   7200
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid gridregistrasi 
         Height          =   7695
         Left            =   3240
         TabIndex        =   28
         Top             =   840
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   13573
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
   Begin VB.Frame ttotalbayar 
      Caption         =   "FORM TRANSAKSI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox tbiayasim 
         BackColor       =   &H80000003&
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
         TabIndex        =   30
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox tkembali 
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
         TabIndex        =   26
         Top             =   6960
         Width           =   1935
      End
      Begin VB.TextBox tbayar 
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
         TabIndex        =   5
         Top             =   6360
         Width           =   1935
      End
      Begin VB.TextBox ttotal 
         BackColor       =   &H80000003&
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
         Left            =   2520
         TabIndex        =   4
         Top             =   5760
         Width           =   1935
      End
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
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   7920
         Visible         =   0   'False
         Width           =   1095
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
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   7920
         Visible         =   0   'False
         Width           =   1215
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
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   7920
         Width           =   1095
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   8400
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
               MinWidth        =   2822
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
               MinWidth        =   2822
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
               MinWidth        =   2822
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbkelas 
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox tbiayadaftar 
         BackColor       =   &H80000003&
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
         TabIndex        =   20
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox tbiayapaket 
         BackColor       =   &H80000003&
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
         TabIndex        =   19
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox ttemu 
         BackColor       =   &H80000003&
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
         TabIndex        =   18
         Top             =   3360
         Width           =   495
      End
      Begin VB.ComboBox cmbpaket 
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
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ComboBox cmbnis 
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox tnama 
         BackColor       =   &H80000003&
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
         TabIndex        =   17
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox tnotrans 
         BackColor       =   &H80000003&
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
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "BIAYA SIM"
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
         TabIndex        =   31
         Top             =   5200
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "UANG KEMBALI"
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
         TabIndex        =   25
         Top             =   7005
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "UANG BAYAR"
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
         TabIndex        =   24
         Top             =   6405
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL BAYAR"
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
         TabIndex        =   23
         Top             =   5800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "BIAYA DAFTAR"
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
         TabIndex        =   15
         Top             =   4605
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "BIAYA PAKET"
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
         TabIndex        =   14
         Top             =   4005
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "JUMLAH PERTEMUAN"
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
         TabIndex        =   13
         Top             =   3405
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "TIPE PAKET"
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
         TabIndex        =   12
         Top             =   2800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "KELAS"
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
         TabIndex        =   11
         Top             =   2200
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   10
         Top             =   1600
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   9
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "NO REGISTRASI"
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
         TabIndex        =   8
         Top             =   400
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Registrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim closebutton As New Misc

Private Sub btnbatal_Click()
Call get_disable_form
End Sub

Private Sub btnsimpan_Click()
If cmbnis.Text = "" Then
    MsgBox "Harap pilih induk siswa.", vbCritical, "Information"
ElseIf cmbkelas.Text = "" Then
    MsgBox "Harap pilih kelas.", vbCritical, "Information"
ElseIf cmbpaket.Text = "" Then
    MsgBox "Harap pilih tipe paket.", vbCritical, "Information"
ElseIf tbayar.Text = "" Then
    MsgBox "Nilai bayar tidak boleh kosong.", vbCritical, "Information"
    tbayar.SetFocus
ElseIf tbayar.Text < ttotal.Text Then
    MsgBox "Nilai bayar tidak boleh kurang dari total bayar.", vbCritical, "Information"
    tbayar.SetFocus
Else
    Call koneksi_db
    dataregistrasi.Open "INSERT INTO tbl_registrasi (no_registrasi, kode_siswa, nama_siswa, kelas, kode_paket, jumlah_temu, biaya_paket, biaya_daftar, total_bayar) VALUES ('" & tnotrans.Text & "','" & cmbnis.Text & "','" & tnama.Text & "','" & cmbkelas.Text & "','" & cmbpaket.Text & "','" & ttemu.Text & "','" & tbiayapaket.Text & "','" & tbiayadaftar.Text & "','" & ttotal.Text & "')", konn
    Call get_grid
    Call get_disable_form
    MsgBox "Data berhasil disimpan.", vbInformation, "Information"
End If
End Sub

Private Sub btntambah_Click()
Call get_enable_form
Call koneksi_db
dataregistrasi.Open "SELECT no_registrasi, kode_siswa, nama_siswa, kelas, kode_paket, nama_paket, jumlah_temu, biaya_paket, biaya_daftar, total_bayar FROM tbl_registrasi ORDER BY id_registrasi DESC", konn
With dataregistrasi
    If .BOF And .EOF Then
      tnotrans.Text = "TRSC" + Format(Date, "YYMM") + "001"
    Else
       tnotrans.Text = "TRSC" + Format(Date, "YYMM") + Right(Str(Val(Right(.Fields("no_registrasi"), 3)) + 1001), 3)
    End If
End With
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub

Private Sub cmbnis_Click()
If cmbnis.Text = cmbnis Then
    tnama.Text = ""
        Call koneksi_db
        datasiswa.Open "SELECT noinduk_siswa, nama_siswa FROM tbl_siswa WHERE noinduk_siswa = '" & cmbnis.Text & "' ", konn
        datasiswa.Requery
        Do While Not datasiswa.EOF
            tnama.Text = datasiswa!nama_siswa
            datasiswa.MoveNext
        Loop
End If
End Sub

Sub get_grid()
Call koneksi_db
dataregistrasi.CursorLocation = adUseClient
dataregistrasi.CursorType = adOpenKeyset
dataregistrasi.LockType = adLockOptimistic
dataregistrasi.Open "SELECT no_registrasi, kode_siswa, nama_siswa, kelas, kode_paket, total_bayar FROM tbl_registrasi ORDER BY id_registrasi DESC", konn
Set gridregistrasi.DataSource = dataregistrasi
gridregistrasi.AllowAddNew = False
gridregistrasi.AllowDelete = False
gridregistrasi.AllowUpdate = False
End Sub

Sub get_disable_form()
tnotrans.Enabled = False
cmbnis.Enabled = False
tnama.Enabled = False
cmbkelas.Enabled = False
cmbpaket.Enabled = False
tbiayasim.Enabled = False
ttemu.Enabled = False
tbiayapaket.Enabled = False
tbiayadaftar.Enabled = False
ttotal.Enabled = False
tbayar.Enabled = False
tkembali.Enabled = False
tbiayasim.BackColor = &H80000003
ttotal.BackColor = &H80000003
tbayar.BackColor = &H80000003
tkembali.BackColor = &H80000003
cmbkelas.BackColor = &H80000003
cmbpaket.BackColor = &H80000003
cmbnis.BackColor = &H80000003
btntambah.Enabled = True
btntambah.Visible = True
btnsimpan.Enabled = False
btnsimpan.Visible = True
btnbatal.Enabled = False
btnbatal.Visible = True
tnotrans.Text = ""
cmbpaket.ListIndex = -1
cmbnis.ListIndex = -1
tnama.Text = ""
cmbkelas.ListIndex = -1
tbiayasim.Text = ""
ttemu.Text = ""
tbiayapaket.Text = ""
tbiayadaftar.Text = ""
ttotal.Text = ""
tbayar.Text = ""
tkembali.Text = ""
End Sub

Sub get_enable_form()
tnotrans.Enabled = False
cmbnis.TabIndex = 1
cmbnis.Enabled = True
tnama.Enabled = False
cmbkelas.Enabled = True
cmbpaket.Enabled = True
tbiayasim.Enabled = False
ttemu.Enabled = False
tbiayapaket.Enabled = False
tbiayadaftar.Enabled = False
ttotal.Enabled = False
tbayar.Enabled = True
tkembali.Enabled = False
tbiayasim.BackColor = &H80000003
ttotal.BackColor = &H80000003
tbayar.BackColor = &H80000005
tkembali.BackColor = &H80000005
cmbkelas.BackColor = &H80000005
cmbpaket.BackColor = &H80000005
cmbnis.BackColor = &H80000005
btntambah.Enabled = False
btnsimpan.Enabled = True
btnsimpan.Visible = True
btnbatal.Enabled = True
btnbatal.Visible = True
End Sub

Sub get_nis()
cmbnis.ListIndex = -1
Call koneksi_db
datasiswa.Open "SELECT noinduk_siswa, nama_siswa FROM tbl_siswa", konn
datasiswa.Requery
    Do While Not datasiswa.EOF
        cmbnis.AddItem datasiswa!noinduk_siswa
        datasiswa.MoveNext
    Loop
End Sub

Sub get_paket()
'cmbpaket.Text = "PILIH"
Call koneksi_db
databiaya.Open "SELECT kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar FROM tbl_biaya_paket", konn
databiaya.Requery
    Do While Not databiaya.EOF
        cmbpaket.AddItem databiaya!kode_paket
        databiaya.MoveNext
    Loop
End Sub

Sub get_kelas()
'cmbkelas.Text = "PILIH"
cmbkelas.AddItem ("PAGI")
cmbkelas.AddItem ("SIANG")
cmbkelas.AddItem ("MALAM")
End Sub

Sub get_clear_paket()
tbiayasim.Text = ""
ttemu.Text = ""
tbiayapaket.Text = ""
tbiayadaftar.Text = ""
End Sub

Private Sub cmbpaket_Click()
If cmbpaket.Text = cmbpaket Then
    Call get_clear_paket
    Call koneksi_db
    databiaya.Open "SELECT kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar,biaya_sim,(biaya_paket+biaya_daftar+biaya_sim) AS total FROM tbl_biaya_paket WHERE kode_paket = '" & cmbpaket.Text & "' ", konn
    databiaya.Requery
    Do While Not databiaya.EOF
        'tnmpaket.Text = databiaya!nama_paket
        ttemu.Text = databiaya!paket_pertemuan
        tbiayapaket.Text = databiaya!biaya_paket
        tbiayadaftar.Text = databiaya!biaya_daftar
        tbiayasim.Text = databiaya!biaya_sim
        ttotal.Text = databiaya!total
        databiaya.MoveNext
    Loop
End If
End Sub

Private Sub Form_Load()
closebutton.DisableCloseButton Me
Call get_nis
Call get_kelas
Call get_paket
Call get_disable_form
Call get_grid
End Sub

Private Sub tbayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    tkembali.Text = tbayar.Text - ttotal.Text
End If
End Sub
