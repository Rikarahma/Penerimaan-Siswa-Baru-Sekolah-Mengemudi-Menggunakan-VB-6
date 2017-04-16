VERSION 5.00
Begin VB.Form Form_login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: Menu Login  ::."
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton btnLogin 
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton btnTutup 
         Caption         =   "TUTUP"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox tpass 
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
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   2535
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
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "SEKOLAH KURSUS MENGEMUDI"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   720
         Width           =   2535
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         TabIndex        =   7
         Top             =   1750
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "SISTEM INFORMASI"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "PENERIMAAN SISWA BARU"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   2295
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
      End
   End
End
Attribute VB_Name = "Form_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lepel As New Misc
Dim closebutton As New Misc


Sub menu_admin()
Form_utama.menuutama.Enabled = True
Form_utama.menuhistorylogin.Enabled = True
Form_utama.menulogout.Enabled = True
End Sub

Sub menu_user()
Form_utama.menuutama.Enabled = True
Form_utama.menuhistorylogin.Enabled = False
Form_utama.menulogout.Enabled = True
End Sub


Private Sub btnLogin_Click()
Dim waktu, tanggal, nmuser, waktuskrng, cekpass As String
Dim tlvl, iduser, logged As Integer
waktu = Time$()
waktuskrng = Time
tanggal = Format(Date, "yyyy-m-d")
logged = 1

    Call koneksi_db
    datauser.Open "SELECT id_user,username,katasandi,level FROM tbl_user WHERE username = '" & tnama.Text & "' AND katasandi = '" & tpass.Text & "' ", konn
    cekpass = datauser.Fields("katasandi")
    If datauser.EOF Then
        MsgBox "Login Gagal, Silahkan cek username dan password anda kembali", vbCritical, "Information"
        tnama.Text = ""
        tpass.Text = ""
        tnama.SetFocus
    ElseIf cekpass <> tpass.Text Then
        MsgBox "Login Gagal, Silahkan cek username dan password anda kembali", vbCritical, "Information"
        tnama.Text = ""
        tpass.Text = ""
        tnama.SetFocus
    Else
        tlvl = datauser!Level
        iduser = datauser!id_user
        nmuser = datauser!UserName
        historylogin.Open "INSERT INTO tbl_history (id_user,id_level,jam,tanggal) VALUES('" & iduser & "', '" & tlvl & "', '" & waktu & "', '" & tanggal & "')", konn
    
        If tlvl = 1 Then 'admin
            Unload Me
            Call menu_admin
            Form_utama.Show
            MsgBox "Halo, " & nmuser & " Anda Login Sebagai Administrator", vbInformation, "Information"
            Form_utama.sbmenu.Panels(1) = "Username : " & UCase(nmuser)
            Form_utama.sbmenu.Panels(2) = logged
            Form_utama.sbmenu.Panels(3) = "User Akses : " & lepel.nmlevel(Val(tlvl))
            Form_utama.sbmenu.Panels(4) = "Waktu : " & waktuskrng
            Form_utama.sbmenu.Panels(5) = iduser
            Form_utama.sbmenu.Panels(6) = nmuser
        ElseIf tlvl = 2 Then 'user
            Unload Me
            Call menu_user
            Form_utama.Show
            MsgBox "Halo, " & nmuser & " Anda Login Sebagai User Pengelola", vbInformation, "Information"
            Form_utama.sbmenu.Panels(1) = "Username : " & UCase(nmuser)
            Form_utama.sbmenu.Panels(2) = logged
            Form_utama.sbmenu.Panels(3) = "User Akses : " & lepel.nmlevel(Val(tlvl))
            Form_utama.sbmenu.Panels(4) = "Waktu : " & waktuskrng
            Form_utama.sbmenu.Panels(5) = iduser
            Form_utama.sbmenu.Panels(6) = nmuser
            Form_utama.menumaster.Visible = False
            Form_utama.menulaporan.Visible = False
        End If
    End If
End Sub

Private Sub btntutup_Click()
End
End Sub

Private Sub Form_Load()
closebutton.DisableCloseButton Me
End Sub

Private Sub tnama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If tnama.Text = "" Then
        MsgBox "Harap isi kolom username", vbInformation, "Information"
        tnama.SetFocus
    ElseIf tpass.Text = "" Then
        MsgBox "Harap isi kolom password", vbInformation, "Information"
        tpass.SetFocus
    Else
        Call btnLogin_Click
    End If
End If
End Sub

Private Sub tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If tnama.Text = "" Then
        MsgBox "Harap isi kolom username", vbInformation, "Information"
        tnama.SetFocus
    ElseIf tpass.Text = "" Then
        MsgBox "Harap isi kolom password", vbInformation, "Information"
        tpass.SetFocus
    Else
        Call btnLogin_Click
    End If
End If
End Sub



