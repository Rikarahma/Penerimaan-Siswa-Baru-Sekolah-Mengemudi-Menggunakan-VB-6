VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_datajadwal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: DATA JADWAL SISWA MENGEMUDI ::."
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   13440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13215
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
         Left            =   11520
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid gridjadwal 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   0   'False
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
Attribute VB_Name = "Form_datajadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tampildata()
Call koneksi_db
datajadwal.CursorLocation = adUseClient
datajadwal.CursorType = adOpenKeyset
datajadwal.LockType = adLockOptimistic
datajadwal.Open "SELECT a.kode_jadwal, a.kode_jam, a.kode_mobil,a.no_registrasi, b.hari, CONCAT(b.durasi,' JAM') as durasi, c.merek_mobil, c.plat_mobil, c.tipe_mobil, d.nama_siswa, d.telepon, e.nama_paket, e.paket_pertemuan as pertemuan FROM tbl_jadwal a JOIN tbl_jam_latihan b ON a.kode_jam = b.kode_jam_latihan JOIN tbl_mobil c ON a.kode_mobil = c.kode_mobil join tbl_registrasi d ON a.no_registrasi = d.no_registrasi JOIN tbl_biaya_paket e ON d.kode_paket = e.id_biaya_paket", konn
Set gridjadwal.DataSource = datajadwal
gridjadwal.Columns(0).Width = 1000
gridjadwal.Columns(1).Width = 800
gridjadwal.Columns(2).Width = 1000
gridjadwal.Columns(3).Width = 1200
gridjadwal.Columns(4).Width = 1000
gridjadwal.Columns(5).Width = 700
gridjadwal.Columns(6).Width = 1000
gridjadwal.Columns(7).Width = 1000
gridjadwal.Columns(8).Width = 900
gridjadwal.Columns(9).Width = 1000
gridjadwal.Columns(10).Width = 1100
gridjadwal.Columns(11).Width = 1000
gridjadwal.Refresh
End Sub

Private Sub btntutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call tampildata
End Sub
