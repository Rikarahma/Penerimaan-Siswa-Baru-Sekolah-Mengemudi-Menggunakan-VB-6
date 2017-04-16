VERSION 5.00
Begin VB.Form Form_Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tentang Aplikasi"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Label Label16 
         Caption         =   "Lisensi GNU GPL-3.0"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label lbltxt 
         Caption         =   "Source Code :"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblsrc 
         Caption         =   "#NA"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   4320
         Width           =   3615
      End
      Begin VB.Label Label15 
         Caption         =   "Aplikasi ini bersifat Open Source."
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4920
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label14 
         Caption         =   "11223344"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   3300
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "12155917"
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
         TabIndex        =   13
         Top             =   2960
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "12156134"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   2590
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "12155596"
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
         Left            =   2520
         TabIndex        =   11
         Top             =   2240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "12150143"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   1880
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "?????"
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
         Left            =   480
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "TOMMY MOFU"
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
         Left            =   480
         TabIndex        =   8
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "CHANDRA PRAKOSO"
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
         Left            =   480
         TabIndex        =   7
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "SAIFUL RAMADHANI"
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
         Left            =   480
         TabIndex        =   6
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "POPO SUWONDO"
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
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "TEAM PENGEMBANG :"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4920
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Sekolah Kursus Mengemudi SideCar"
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
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Penerimaan Siswa Baru"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Sistem Informasi"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Load()
'lblsrc.Caption = "Source Code : Https://www.github.com/justpoypoy"
'End Sub
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
With lblsrc
    .Caption = "http://s.id/psbvb6"
    .ForeColor = vbBlue
    .Font.Underline = True
End With
End Sub

Private Sub lblsrc_Click()
With lblsrc
' Call ShellExecute(0&, vbNullString, "Mailto:" & .Caption, vbNullString, vbNullString, vbNormalFocus)
  Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
End With
End Sub
