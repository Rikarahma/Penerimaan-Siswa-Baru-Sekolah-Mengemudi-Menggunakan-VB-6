VERSION 5.00
Begin VB.Form Form_chpass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: RUBAH PASSWORD ::."
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton btnTutup 
         Caption         =   "BATAL"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton btnSimpan 
         Caption         =   "PERBAHARUI"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox tpasskonf 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox tpassnew 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox tpassold 
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2400
         TabIndex        =   1
         Top             =   960
         Width           =   2295
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
         Height          =   405
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblmatching 
         Caption         =   "#MATCHPASS"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbliduser 
         Caption         =   "#IDUSER"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblpassold 
         Caption         =   "#PASSOLD"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "KOFNRIMASI PASSWORD"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2265
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "PASSWORD BARU"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "PASSWORD LAMA"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "USERNAME"
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
         Left            =   240
         TabIndex        =   6
         Top             =   470
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form_chpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSimpan_Click()
Dim passold, passnew, passkonf As String
passold = tpassold.Text
passnew = tpassnew.Text
passkonf = tpasskonf.Text

Call koneksi_db
datauser.Open "SELECT katasandi FROM tbl_user WHERE id_user = '" & lbliduser & "' ", konn
    If datauser.EOF Then
        lblpassold.Caption = ""
        MsgBox "Password lama salah.", vbCritical, "Information"
    Else
        lblpassold.Caption = datauser.Fields("katasandi")
        If passnew <> passkonf Then
            MsgBox "Password baru dan konfirmasi password anda salah, harap perhatikan huruf besar dan kecil.", vbCritical, "Information"
            tpassnew.Text = ""
            tpasskonf.Text = ""
            tpassnew.SetFocus
        ElseIf passold <> lblpassold.Caption Then
            MsgBox "Password lama anda salah, harap perhatikan huruf besar dan kecil.", vbCritical, "Information"
            tpassold.Text = ""
            tpassold.SetFocus
        ElseIf passkonf = passold Then
            MsgBox "Password tidak berubah, karena password sama.", vbCritical, "Information"
            tpassold.Text = ""
            tpassnew.Text = ""
            tpasskonf.Text = ""
            tpassold.SetFocus
        Else
            Call koneksi_db
            datauser.Open " UPDATE tbl_user SET katasandi = '" & passkonf & "' WHERE id_user = '" & lbliduser & "' ", konn
            MsgBox "Password anda telah berhasil dirubah.", vbInformation, "Information"
            Unload Me
        End If
    End If
End Sub

Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
tnama.Text = Form_utama.sbmenu.Panels(6)
lbliduser.Caption = Form_utama.sbmenu.Panels(5)
tnama.Enabled = False
tnama.BackColor = &H8000000A
lbliduser.Visible = False
lblpassold.Visible = False
tpassold.PasswordChar = "*"
tpassnew.PasswordChar = "*"
tpasskonf.PasswordChar = "*"
End Sub

Private Sub tpasskonf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnSimpan_Click
End If
End Sub

Private Sub tpassnew_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call btnSimpan_Click
End If
End Sub
