VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_history 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: HISTORY LOGIN ::."
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
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
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid gridHistory 
         Height          =   5655
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   9975
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
      Begin VB.Label Label1 
         Caption         =   "RIWAYAT LOGIN USER AKSES"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "Form_history"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call koneksi_db
historylogin.CursorLocation = adUseClient
historylogin.CursorType = adOpenKeyset
historylogin.LockType = adLockOptimistic
historylogin.Open "SELECT b.USERNAME, IF(a.id_level = 1,'Administrator','User Pengelola') AS USER_AKSES, a.JAM  FROM tbl_history AS a LEFT JOIN tbl_user AS b ON a.id_user = b.id_user ORDER BY a.id_history DESC", konn
Set gridHistory.DataSource = historylogin
gridHistory.AllowDelete = False
gridHistory.AllowUpdate = False

End Sub
