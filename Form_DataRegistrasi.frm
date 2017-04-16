VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_DataRegistrasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: FORM DATA REGISTRASI ::."
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin MSDataGridLib.DataGrid gridregistrasi 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4471
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
End
Attribute VB_Name = "Form_DataRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tampilgrid()
Call koneksi_db
dataregistrasi.CursorLocation = adUseClient
dataregistrasi.CursorType = adOpenKeyset
dataregistrasi.LockType = adLockOptimistic
dataregistrasi.Open "SELECT a.no_registrasi, a.nama_siswa, a.ktp, a.noinduk_siswa, b.nama_paket FROM tbl_registrasi a JOIN tbl_biaya_paket b ON a.kode_paket = b.id_biaya_paket WHERE a.flag_jadwal = 1 ORDER BY a.id_registrasi DESC", konn
Set gridregistrasi.DataSource = dataregistrasi
gridregistrasi.Columns(0).Width = 1500
gridregistrasi.Columns(1).Width = 1500
gridregistrasi.Columns(2).Width = 1600
gridregistrasi.Columns(3).Width = 1500
gridregistrasi.Columns(4).Width = 1500
gridregistrasi.AllowAddNew = False
gridregistrasi.AllowDelete = False
gridregistrasi.AllowUpdate = False
gridregistrasi.Refresh
End Sub

Sub hidden_all()
'cmdtambah.Visible = False
'lblnoregist.Visible = False
'lblnosiswa.Visible = False
'tcari.SetFocus
End Sub

Sub visib_all()
cmdtambah.Visible = True
lblnoregist.Visible = True
lblnosiswa.Visible = True
End Sub

Private Sub btncari_Click()
If tcari.Text = "" Then
    MsgBox "Kolom pencarian tidak boleh kosong.", vbCritical, "Information"
    tcari.SetFocus
Else
    Call koneksi_db
    dataregistrasi.Open "SELECT no_registrasi, noinduk_siswa FROM tbl_registrasi WHERE no_registrasi = '" & tcari.Text & "'", konn
    If dataregistrasi.EOF Then
        MsgBox "Nomor Registrasi " + tcari.Text + " tidak ada, Harap pastikan nomor registrasi ada.", vbCritical, "Information"
    Call hidden_all
    Else
        lblnoregist.Caption = dataregistrasi!no_registrasi
        lblnosiswa.Caption = dataregistrasi!noinduk_siswa
        Call visib_all
    End If
End If
End Sub

Private Sub Form_Load()
Call tampilgrid
Call hidden_all
'tcari.SetFocus
End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
