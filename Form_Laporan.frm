VERSION 5.00
Begin VB.Form Form_Laporan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MENU LAPORAN ::."
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btntampilregist 
      Caption         =   "TAMPIL DATA"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton btntampiljadwal 
      Caption         =   "TAMPIL DATA"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton btncetakjadwal 
      Caption         =   "CETAK"
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
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton btncetakregist 
         Caption         =   "CETAK"
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
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton btncari 
         Caption         =   "CARI"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cmblaporan 
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
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "PILIH LAPORAN"
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
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form_Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub isicmb()
cmblaporan.AddItem "REGISTER"
cmblaporan.AddItem "JADWAL"
End Sub

Sub hiddenbtn()
btncetakjadwal.Visible = False
btncetakregist.Visible = False
btntampiljadwal.Visible = False
btntampilregist.Visible = False
End Sub

Private Sub btncari_Click()
If cmblaporan.Text = "REGISTER" Then
    btncetakregist.Visible = True
    btntampilregist.Visible = True
ElseIf cmblaporan.Text = "JADWAL" Then
    btncetakjadwal.Visible = True
    btntampiljadwal.Visible = True
Else
    MsgBox "Harap pilih laporan dahulu.", vbCritical, "Information"
End If
End Sub

Private Sub btntampiljadwal_Click()
Load Form_datajadwal
Form_datajadwal.Show 1, Form_Laporan
End Sub

Private Sub btntampilregist_Click()
Load Form_DataRegistrasi
Form_DataRegistrasi.Show 1, Form_Laporan
End Sub

Private Sub cmblaporan_Click()
Call hiddenbtn
End Sub

Private Sub Form_Load()
Call isicmb
Call hiddenbtn
End Sub
