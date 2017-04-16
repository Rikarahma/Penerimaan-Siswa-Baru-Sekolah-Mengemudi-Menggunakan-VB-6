VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Form_utama 
   BackColor       =   &H8000000C&
   Caption         =   ".:: Halaman Utama ::."
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14460
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbmenu 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   6150
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Ubuntu Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu menuutama 
      Caption         =   "&Menu"
      Begin VB.Menu menuhistorylogin 
         Caption         =   "Riwayat Login"
         Shortcut        =   ^L
      End
      Begin VB.Menu menuchpas 
         Caption         =   "Ganti Password"
         Shortcut        =   ^G
      End
      Begin VB.Menu menulogout 
         Caption         =   "Log Out"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menumaster 
      Caption         =   "Master &Data"
      Begin VB.Menu menudatauser 
         Caption         =   "Master Data User Akses"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menudatasiswa 
         Caption         =   "Master Data Siswa"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu menudatabiaya 
         Caption         =   "Master Data Biaya"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menudatamobil 
         Caption         =   "Master Data Mobil"
         Shortcut        =   {F4}
      End
      Begin VB.Menu menujamlatihan 
         Caption         =   "Master Data Jam Latihan"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu menutrx 
      Caption         =   "&Registrasi"
      Begin VB.Menu menudaftar 
         Caption         =   "Registrasi"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu menudataregistrasi 
         Caption         =   "Data Registrasi"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu menusima 
         Caption         =   "Pembuatan SIM A"
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menujadwal 
      Caption         =   "&Jadwal"
      Begin VB.Menu menuformjadwal 
         Caption         =   "Form Jadwal"
      End
      Begin VB.Menu menudatajadwal 
         Caption         =   "Data Jadwal"
      End
   End
   Begin VB.Menu menulaporan 
      Caption         =   "&Laporan"
   End
   Begin VB.Menu menuinfo 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "Form_utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim closebutton As New Misc

Private Sub MDIForm_Load()
closebutton.DisableCloseButton Me
End Sub

Private Sub menuchpas_Click()
Load Form_chpass
Form_chpass.Show 1, Form_utama
End Sub

Private Sub menudaftar_Click()
Load Form_Register
Form_Register.Show 1, Form_utama
End Sub

Private Sub menudatabiaya_Click()
Load Form_Databiaya
Form_Databiaya.Show 1, Form_utama
End Sub

Private Sub menudatajadwal_Click()
Load Form_datajadwal
Form_datajadwal.Show 1, Form_utama
End Sub

Private Sub menudatamobil_Click()
Load Form_Datamobil
Form_Datamobil.Show 1, Form_utama
End Sub

Private Sub menudataregistrasi_Click()
Load Form_DataRegistrasi
Form_DataRegistrasi.Show 1, Form_utama
End Sub

Private Sub menudatasiswa_Click()
Load Form_Datasiswa
Form_Datasiswa.Show 1, Form_utama
End Sub

Private Sub menudatauser_Click()
Load Form_Datauser
Form_Datauser.Show 1, Form_utama
End Sub

Private Sub menuformjadwal_Click()
Load Form_Jadwal
Form_Jadwal.Show 1, Form_utama
End Sub

Private Sub menuhistorylogin_Click()
Load Form_history
Form_history.Show 1, Form_utama
End Sub

Private Sub menuinfo_Click()
Load Form_Info
Form_Info.Show 1, Form_utama
End Sub

Private Sub menujamlatihan_Click()
Load Form_Jamlatihan
Form_Jamlatihan.Show 1, Form_utama
End Sub

Private Sub menulaporan_Click()
Load Form_Laporan
Form_Laporan.Show 1, Form_utama
End Sub

Private Sub menulogout_Click()
If MsgBox("Anda akan keluar dari aplikasi ?", vbQuestion + vbYesNo, "Information") = vbYes Then
Unload Me
Form_login.Show
End If
End Sub
