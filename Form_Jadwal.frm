VERSION 5.00
Begin VB.Form Form_Jadwal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: FORM JADWAL SISWA ::."
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tmerek 
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
      Left            =   3000
      TabIndex        =   32
      Top             =   6720
      Width           =   2655
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
      Left            =   1680
      TabIndex        =   29
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox ttipe 
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
      Left            =   3000
      TabIndex        =   28
      Top             =   6240
      Width           =   2655
   End
   Begin VB.ComboBox cmbkdmobil 
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
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox cmbkdjam 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ComboBox cmbregist 
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   960
         Width           =   2655
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
         Left            =   4560
         TabIndex        =   24
         Top             =   7560
         Width           =   1095
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
         Left            =   2880
         TabIndex        =   23
         Top             =   7560
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
         TabIndex        =   22
         Top             =   7560
         Width           =   1095
      End
      Begin VB.TextBox tplat 
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
         Left            =   2880
         TabIndex        =   21
         Top             =   5760
         Width           =   2655
      End
      Begin VB.TextBox tdurasi 
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
         Left            =   2880
         TabIndex        =   20
         Top             =   4800
         Width           =   2655
      End
      Begin VB.TextBox thari 
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
         Left            =   2880
         TabIndex        =   19
         Top             =   4320
         Width           =   2655
      End
      Begin VB.TextBox ttemu 
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
         Left            =   2880
         TabIndex        =   18
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox tnmpaket 
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
         Left            =   2880
         TabIndex        =   17
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox tkelas 
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
         Left            =   2880
         TabIndex        =   16
         Top             =   2400
         Width           =   2655
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
         Left            =   2880
         TabIndex        =   15
         Top             =   1920
         Width           =   2655
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
         Left            =   2880
         TabIndex        =   14
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox kdjadwal 
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
         Left            =   2880
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "MEREK MOBIL"
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
         Top             =   6800
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "TIPE MOBIL"
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
         TabIndex        =   30
         Top             =   6300
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "PLAT MOBIL"
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
         Top             =   5800
         Width           =   1555
      End
      Begin VB.Label Label11 
         Caption         =   "KODE MOBIL"
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
         Top             =   5350
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4860
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         TabIndex        =   9
         Top             =   4370
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         TabIndex        =   8
         Top             =   3880
         Width           =   1935
      End
      Begin VB.Label Label7 
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
         TabIndex        =   7
         Top             =   3400
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "NAMA PAKET"
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
         TabIndex        =   6
         Top             =   2950
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "KELAS SISWA"
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
         TabIndex        =   5
         Top             =   2450
         Width           =   1935
      End
      Begin VB.Label Label4 
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
         TabIndex        =   4
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "NO REGISTRASI SISWA"
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
         Top             =   1000
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "KODE JADWAL"
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
         Top             =   500
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form_Jadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancel_Click()
Call nonaktif_tb
btntambah.Enabled = True
btnsimpan.Enabled = False
btncancel.Enabled = False
End Sub

Private Sub btnsimpan_Click()
If cmbregist.Text = "" Then
    MsgBox "No registrasi harap dipilih dahulu", vbCritical, "Information"
    cmbregist.SetFocus
ElseIf cmbkdjam.Text = "" Then
    MsgBox "kode jam harap dipilih dahulu", vbCritical, "Information"
    cmbkdjam.SetFocus
ElseIf cmbkdmobil.Text = "" Then
    MsgBox "kode mobil harap dipilih dahulu", vbCritical, "Information"
    cmbkdmobil.SetFocus
Else
    Call koneksi_db
    datajadwal.Open "INSERT INTO tbl_jadwal (kode_jadwal, no_registrasi, kode_jam, kode_mobil) VALUES ('" & kdjadwal.Text & "','" & cmbregist.Text & "','" & cmbkdjam.Text & "','" & cmbkdmobil.Text & "') ", konn
    MsgBox "JADWAL SISWA " + tnama.Text + " BERHASIL DISIMPAN", vbInformation, "Information"
    Call nonaktif_tb
    btntambah.Enabled = True
    btnsimpan.Enabled = False
    btncancel.Enabled = False
End If
End Sub

Private Sub btntambah_Click()
Call aktif_tb
Call kode_jadwal
Call isicmbnoregistrasi
Call isicmbkdjam
Call isicmbkdmobil
cmbregist.SetFocus
cmbregist.TabIndex = 1
cmbkdjam.TabIndex = 2
cmbkdmobil.TabIndex = 3
btntambah.Enabled = False
btnsimpan.Enabled = True
btncancel.Enabled = True
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub

Sub kode_jadwal()
Call koneksi_db
datasiswa.Open "SELECT kode_jadwal FROM tbl_jadwal ORDER BY id_jadwal DESC", konn
With datasiswa
    If .BOF And .EOF Then
      kdjadwal.Text = "JSC" + "001"
    Else
       kdjadwal.Text = "JSC" + Right(Str(Val(Right(.Fields("kode_jadwal"), 3)) + 1001), 3)
    End If
End With
End Sub

Private Sub cmbregist_Click()
Call koneksi_db
dataregistrasi.Open "SELECT a.no_registrasi, a.noinduk_siswa, a.nama_siswa, a.kelas, b.nama_paket, b.paket_pertemuan FROM tbl_registrasi a JOIN tbl_biaya_paket b ON a.kode_paket = b.id_biaya_paket WHERE a.no_registrasi = '" & cmbregist.Text & "' ", konn
dataregistrasi.Requery
    Do While Not dataregistrasi.EOF
        tnis.Text = dataregistrasi!noinduk_siswa
        tnama.Text = dataregistrasi!nama_siswa
        tkelas.Text = dataregistrasi!kelas
        tnmpaket.Text = dataregistrasi!nama_paket
        ttemu.Text = dataregistrasi!paket_pertemuan
        dataregistrasi.MoveNext
    Loop
End Sub

Sub isicmbnoregistrasi()
Call koneksi_db
dataregistrasi.Open "SELECT no_registrasi FROM tbl_registrasi WHERE no_registrasi NOT IN (SELECT no_registrasi FROM tbl_jadwal)", konn
dataregistrasi.Requery
    Do While Not dataregistrasi.EOF
        cmbregist.AddItem dataregistrasi!no_registrasi
        dataregistrasi.MoveNext
    Loop
End Sub

Private Sub cmbkdjam_Click()
Call koneksi_db
datadurasilatihan.Open "SELECT kode_jam_latihan, hari, durasi FROM tbl_jam_latihan WHERE kode_jam_latihan = '" & cmbkdjam.Text & "' ", konn
datadurasilatihan.Requery
    Do While Not datadurasilatihan.EOF
        thari.Text = datadurasilatihan!hari
        tdurasi.Text = datadurasilatihan!durasi & " Jam"
        datadurasilatihan.MoveNext
    Loop
End Sub

Sub isicmbkdjam()
Call koneksi_db
datadurasilatihan.Open "SELECT kode_jam_latihan FROM tbl_jam_latihan", konn
datadurasilatihan.Requery
    Do While Not datadurasilatihan.EOF
        cmbkdjam.AddItem datadurasilatihan!kode_jam_latihan
        datadurasilatihan.MoveNext
    Loop
End Sub

Private Sub cmbkdmobil_Click()
Call koneksi_db
datamobil.Open "SELECT kode_mobil, plat_mobil, merek_mobil,tipe_mobil FROM tbl_mobil WHERE kode_mobil = '" & cmbkdmobil.Text & "' ", konn
datamobil.Requery
    Do While Not datamobil.EOF
        tplat.Text = datamobil!plat_mobil
        tmerek.Text = datamobil!merek_mobil
        ttipe.Text = datamobil!tipe_mobil
        datamobil.MoveNext
    Loop
End Sub

Sub isicmbkdmobil()
Call koneksi_db
datamobil.Open "SELECT kode_mobil FROM tbl_mobil", konn
datamobil.Requery
    Do While Not datamobil.EOF
        cmbkdmobil.AddItem datamobil!kode_mobil
        datamobil.MoveNext
    Loop
End Sub

Sub clearcmbregist()
cmbregist.Clear
cmbregist.ListIndex = -1
End Sub

Sub clearcmbkdjam()
cmbkdjam.Clear
cmbkdjam.ListIndex = -1
End Sub

Sub clearcmbkdmobil()
cmbkdmobil.Clear
cmbkdmobil.ListIndex = -1
End Sub

Sub nonaktif_tb()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
If TypeOf kontrol Is TextBox Then kontrol.BackColor = &H80000003
If TypeOf kontrol Is TextBox Then kontrol.Text = ""
Next
cmbregist.Enabled = False
cmbkdjam.Enabled = False
cmbkdmobil.Enabled = False
Call clearcmbregist
Call clearcmbkdjam
Call clearcmbkdmobil
End Sub

Sub aktif_tb()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
cmbregist.Enabled = True
cmbkdjam.Enabled = True
cmbkdmobil.Enabled = True
End Sub

Private Sub Form_Load()
Call nonaktif_tb
btnsimpan.Enabled = False
btncancel.Enabled = False
'cmbregist.AddItem "NO REGISTRASI SISWA"
'cmbregist.ListIndex = cmbregist.NewIndex
'cmbkdjam.AddItem "JAM LATIHAN"
'cmbkdjam.ListIndex = cmbkdjam.NewIndex
'cmbkdmobil.AddItem "KODE MOBIL"
'cmbkdmobil.ListIndex = cmbkdmobil.NewIndex
End Sub
