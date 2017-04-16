VERSION 5.00
Begin VB.Form Form_Register 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: FORM REGISTRASI ::."
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "FORM SISWA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton btnupdate 
         Caption         =   "UPDATE"
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
         TabIndex        =   35
         Top             =   4560
         Width           =   1095
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
         Height          =   495
         Left            =   2640
         TabIndex        =   34
         Top             =   4560
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
         Height          =   495
         Left            =   1440
         TabIndex        =   33
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton btnhapus 
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
         TabIndex        =   32
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
      End
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
         Left            =   4680
         TabIndex        =   31
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   240
         TabIndex        =   30
         Top             =   4560
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   2640
         TabIndex        =   28
         Top             =   1560
         Width           =   2535
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
         Height          =   855
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3360
         Width           =   2535
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
         Left            =   2640
         TabIndex        =   21
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox tnohp 
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
         Left            =   2640
         TabIndex        =   20
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox tnis 
         BackColor       =   &H80000003&
         Enabled         =   0   'False
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
         Left            =   2640
         TabIndex        =   19
         Top             =   360
         Width           =   2055
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
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblkdnis 
         Caption         =   "#kdnis"
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "NO KTP"
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
         Left            =   360
         TabIndex        =   29
         Top             =   1600
         Width           =   1455
      End
      Begin VB.Label Label13 
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
         Left            =   360
         TabIndex        =   27
         Top             =   3405
         Width           =   1695
      End
      Begin VB.Label Label12 
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
         Left            =   360
         TabIndex        =   26
         Top             =   2805
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Left            =   360
         TabIndex        =   25
         Top             =   2205
         Width           =   1815
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   24
         Top             =   980
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   23
         Top             =   400
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FORM REGISTRASI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox tjumlahbayar 
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
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox ttotal 
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
         Left            =   2160
         TabIndex        =   8
         Top             =   3960
         Width           =   1815
      End
      Begin VB.TextBox tbiayasim 
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
         Left            =   2160
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox tbiayadaftar 
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
         Left            =   2160
         TabIndex        =   6
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox tbiayapaket 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   2160
         Width           =   1815
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
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
         ItemData        =   "Form_Register.frx":0000
         Left            =   2160
         List            =   "Form_Register.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox tnoregist 
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
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblkdpaket 
         Caption         =   "#kdpaket"
         Height          =   375
         Left            =   960
         TabIndex        =   36
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "JUMLAH BAYAR"
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
         TabIndex        =   16
         Top             =   4600
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         TabIndex        =   15
         Top             =   4000
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         TabIndex        =   14
         Top             =   3400
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   2850
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         TabIndex        =   12
         Top             =   2205
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "PAKET"
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
         Top             =   1600
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         TabIndex        =   10
         Top             =   1000
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         TabIndex        =   2
         Top             =   400
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form_Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, e As String
Sub nonaktif_textbox()
tnoregist.Text = ""
tnis.Text = ""
tjumlahbayar.Text = ""
tnama.Text = ""
tktp.Text = ""
tnohp.Text = ""
temail.Text = ""
talamat.Text = ""
tnoregist.Enabled = False
tjumlahbayar.Enabled = False
tnama.Enabled = False
tktp.Enabled = False
tnohp.Enabled = False
temail.Enabled = False
talamat.Enabled = False
cmbkelas.Enabled = False
cmbpaket.Enabled = False
btntambah.Enabled = True
btntambah.Visible = True
btntutup.Enabled = True
btntutup.Visible = True
btnsimpan.Enabled = False
btnsimpan.Visible = False
btnupdate.Enabled = False
btnupdate.Visible = False
btnbatal.Visible = False
btnbatal.Visible = False
Call clearcmbkelas
Call clearcmbpaket
Call get_clear_paket
btntambah.TabIndex = 1
btntutup.TabIndex = 2
End Sub

Sub aktif_textbox()
tnoregist.Enabled = False
tjumlahbayar.Enabled = True
tnama.Enabled = True
tktp.Enabled = True
tnohp.Enabled = True
temail.Enabled = True
talamat.Enabled = True
cmbkelas.Enabled = True
cmbpaket.Enabled = True
btntambah.Enabled = False
btntambah.Visible = True
btnsimpan.Enabled = True
btnsimpan.Visible = True
btnbatal.Enabled = True
btnbatal.Visible = True
btntutup.Enabled = True
btntutup.Visible = True
Call isicmbkelas
Call isicmbpaket
cmbkelas.SetFocus
tktp.MaxLength = 16
tnohp.MaxLength = 20
End Sub

Sub isicmbkelas()
cmbkelas.AddItem ("PAGI")
cmbkelas.AddItem ("SIANG")
cmbkelas.AddItem ("MALAM")
End Sub

Sub clearcmbkelas()
cmbkelas.Clear
cmbkelas.ListIndex = -1
End Sub

Sub isicmbpaket()
Call koneksi_db
databiaya.Open "SELECT kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar FROM tbl_biaya_paket", konn
databiaya.Requery
    Do While Not databiaya.EOF
        cmbpaket.AddItem databiaya!kode_paket
        databiaya.MoveNext
    Loop
End Sub

Sub clearcmbpaket()
cmbpaket.Clear
cmbpaket.ListIndex = -1
End Sub

Sub getindexadd()
cmbkelas.TabIndex = 1
cmbpaket.TabIndex = 2
tjumlahbayar.TabIndex = 3
tnama.TabIndex = 4
tktp.TabIndex = 5
tnohp.TabIndex = 6
temail.TabIndex = 7
talamat.TabIndex = 8
btnsimpan.TabIndex = 9
btnbatal.TabIndex = 10
btntutup.TabIndex = 11
End Sub

Sub get_clear_paket()
tbiayasim.Text = ""
ttotal.Text = ""
tbiayapaket.Text = ""
tbiayadaftar.Text = ""
End Sub

Private Sub btnsimpan_Click()
Dim noregist As String
a = Val(ttotal.Text)
b = Val(tjumlahbayar.Text)
If cmbkelas.Text = "" Then
    MsgBox "Harap pilih kelas.", vbCritical, "Information"
    cmbkelas.SetFocus
ElseIf cmbpaket.Text = "" Then
    MsgBox "Harap pilih paket.", vbCritical, "Information"
    cmbpaket.SetFocus
ElseIf tjumlahbayar.Text = "" Then
    MsgBox "Jumlah bayar tidak boleh kosong.", vbCritical, "Information"
    tjumlahbayar.SetFocus
ElseIf tnama.Text = "" Then
    MsgBox "Kolom nama tidak boleh kosong.", vbCritical, "Information"
    tnama.SetFocus
ElseIf tktp.Text = "" Then
    MsgBox "Kolom ktp tidak boleh kosong.", vbCritical, "Information"
    tktp.SetFocus
'ElseIf tktp.Text = e Then
'    MsgBox "No KTP " + tktp.Text + " telah ada disystem.", vbCritical, "Information"
'    tktp.Text = ""
'    tktp.SetFocus
ElseIf tnohp.Text = "" Then
    MsgBox "Kolom telepon tidak boleh kosong.", vbCritical, "Information"
    tnohp.SetFocus
ElseIf temail.Text = "" Then
    MsgBox "Kolom email tidak boleh kosong.", vbCritical, "Information"
    temail.SetFocus
ElseIf temail.Text = "" Then
    MsgBox "Kolom alamat tidak boleh kosong.", vbCritical, "Information"
    temail.SetFocus
ElseIf tjumlahbayar.Text < ttotal.Text Then
    MsgBox "Kolom jumlah bayar tidak boleh kurang dari total bayar.", vbCritical, "Information"
    tjumlahbayar.SetFocus
ElseIf tjumlahbayar.Text > ttotal.Text Then
    c = b - a
    msgoke = MsgBox("Uang kembali siswa " + tnama.Text + " sebesar " + "Rp " & Format(Str(c), "###,###,###") + "." + vbNewLine + "Tekan oke untuk menyimpan data registrasi", vbInformation + vbYesNo, "Information")
        If msgoke = vbYes Then
            Call koneksi_db
            dataregistrasi.Open "INSERT INTO tbl_registrasi (no_registrasi, kelas, kode_paket, total_bayar, noinduk_siswa, nama_siswa, ktp, telepon, email, alamat ) VALUES ('" & tnoregist.Text & "','" & cmbkelas.Text & "','" & lblkdpaket.Caption & "','" & tjumlahbayar.Text & "', '" & tnis.Text & "','" & tnama.Text & "','" & tktp.Text & "','" & tnohp.Text & "','" & temail.Text & "','" & talamat.Text & "')", konn
            'datasiswa.Open "INSERT INTO tbl_siswa (noinduk_siswa, nama_siswa, alamat_siswa, ktp_siswa, telpon_siswa, email_siswa) VALUES ('" & tnis.Text & "','" & tnama.Text & "','" & talamat.Text & "','" & tktp.Text & "','" & tnohp.Text & "','" & temail.Text & "')", konn
            MsgBox "Registrasi dengan nomor " + tnoregist.Text + " berhasil tersimpan.", vbInformation, "Information"
            Call nonaktif_textbox
        End If
Else
    Call koneksi_db
    dataregistrasi.Open "INSERT INTO tbl_registrasi (no_registrasi, kelas, kode_paket, total_bayar, noinduk_siswa, nama_siswa, ktp, telepon, email, alamat ) VALUES ('" & tnoregist.Text & "','" & cmbkelas.Text & "','" & lblkdpaket.Caption & "','" & tjumlahbayar.Text & "', '" & tnis.Text & "','" & tnama.Text & "','" & tktp.Text & "','" & tnohp.Text & "','" & temail.Text & "','" & talamat.Text & "')", konn
    'datasiswa.Open "INSERT INTO tbl_siswa (noinduk_siswa, nama_siswa, alamat_siswa, ktp_siswa, telpon_siswa, email_siswa) VALUES ('" & tnis.Text & "','" & tnama.Text & "','" & talamat.Text & "','" & tktp.Text & "','" & tnohp.Text & "','" & temail.Text & "')", konn
    MsgBox "Registrasi dengan nomor " + tnoregist.Text + " berhasil tersimpan.", vbInformation, "Information"
    Call nonaktif_textbox
End If
End Sub

Private Sub btntambah_Click()
Call aktif_textbox
Call koneksi_db
dataregistrasi.Open "SELECT no_registrasi FROM tbl_registrasi ORDER BY id_registrasi DESC", konn
With dataregistrasi
    If .BOF And .EOF Then
      tnoregist.Text = "TRSC" + Format(Date, "YYMM") + "001"
    Else
       tnoregist.Text = "TRSC" + Format(Date, "YYMM") + Right(Str(Val(Right(.Fields("no_registrasi"), 3)) + 1001), 3)
    End If
End With
datasiswa.Open "SELECT noinduk_siswa FROM tbl_registrasi ORDER BY id_registrasi DESC", konn
With datasiswa
    If .BOF And .EOF Then
        tnis.Text = "NSC" + "001"
        Call getindexadd
    Else
        tnis.Text = "NSC" + Right(Str(Val(Right(.Fields("noinduk_siswa"), 3)) + 1001), 3)
        Call getindexadd
    End If
End With
End Sub

Private Sub btnbatal_Click()
Call nonaktif_textbox
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub

Private Sub cmbpaket_Click()
If cmbpaket.Text = cmbpaket Then
    Call get_clear_paket
    Call koneksi_db
    databiaya.Open "SELECT id_biaya_paket,kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar,biaya_sim,(biaya_paket+biaya_daftar+biaya_sim) AS total FROM tbl_biaya_paket WHERE kode_paket = '" & cmbpaket.Text & "' ", konn
    databiaya.Requery
    Do While Not databiaya.EOF
        tbiayapaket.Text = databiaya!biaya_paket
        tbiayadaftar.Text = databiaya!biaya_daftar
        tbiayasim.Text = databiaya!biaya_sim
        ttotal.Text = databiaya!total
        lblkdpaket.Caption = databiaya!id_biaya_paket
        databiaya.MoveNext
        tjumlahbayar.SetFocus
    Loop
End If
End Sub

Private Sub Form_Load()
Call nonaktif_textbox
End Sub

Private Sub tjumlahbayar_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub tktp_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub tnama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tnohp_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub
