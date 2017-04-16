VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Datamobil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MASTER DATA MOBIL ::."
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton btnCancel 
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
         Height          =   615
         Left            =   3000
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton btnHapus 
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
         Height          =   615
         Left            =   4800
         TabIndex        =   16
         Top             =   2760
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
         Height          =   615
         Left            =   1440
         TabIndex        =   15
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton btnupdate 
         Caption         =   "PERBAHARUI"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton btnTutup 
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
         Height          =   615
         Left            =   4800
         TabIndex        =   13
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton btnTambah 
         Caption         =   "TAMBAH MOBIL"
         BeginProperty Font 
            Name            =   "Ubuntu Mono"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton btncari 
         Caption         =   "CARI KODE MOBIL"
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
         Left            =   2280
         TabIndex        =   11
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox tcari 
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
         Left            =   240
         TabIndex        =   10
         Top             =   3720
         Width           =   1815
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
         Left            =   2040
         TabIndex        =   9
         Top             =   2280
         Width           =   3015
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
         Left            =   2040
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
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
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox tkode 
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
         Left            =   2040
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid gridmobil 
         Height          =   3495
         Left            =   240
         TabIndex        =   1
         Top             =   4200
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6165
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
      Begin VB.Label lblplat 
         Caption         =   "#PLAT"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2320
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1720
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   3
         Top             =   1150
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   2
         Top             =   500
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form_Datamobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tampil_grid()
Call koneksi_db
datamobil.CursorLocation = adUseClient
datamobil.CursorType = adOpenKeyset
datamobil.LockType = adLockOptimistic
datamobil.Open "SELECT KODE_MOBIL, MEREK_MOBIL, TIPE_MOBIL, PLAT_MOBIL FROM tbl_mobil ORDER BY id_mobil DESC", konn
Set gridmobil.DataSource = datamobil
gridmobil.Columns(0).Width = 1200
gridmobil.Columns(1).Width = 1300
gridmobil.Columns(2).Width = 1200
gridmobil.Columns(3).Width = 1500
gridmobil.AllowAddNew = False
gridmobil.AllowDelete = False
gridmobil.AllowUpdate = False
gridmobil.Refresh
End Sub

Sub nonaktif_all()
tkode.Enabled = False
tmerek.Enabled = False
ttipe.Enabled = False
tplat.Enabled = False
btnupdate.Visible = False
btnSimpan.Visible = True
btnCancel.Enabled = False
btnSimpan.Enabled = False
btnHapus.Visible = False
btnTutup.Visible = True
btnTambah.Enabled = True
btnTambah.TabIndex = 1
tcari.TabIndex = 2
btnCari.TabIndex = 3
btnTutup.TabIndex = 4
tkode.BackColor = &H80000003
tmerek.BackColor = &H80000003
ttipe.BackColor = &H80000003
tplat.BackColor = &H80000003
tkode.Text = ""
tmerek.Text = ""
ttipe.Text = ""
tplat.Text = ""
tcari.Text = ""
lblplat.Caption = ""
End Sub

Sub aktif_add()
tkode.Enabled = False
tmerek.Enabled = True
ttipe.Enabled = True
tplat.Enabled = True
btnCancel.Enabled = True
btnSimpan.Enabled = True
btnHapus.Enabled = False
btnTambah.Enabled = False
tkode.BackColor = &H80000003
tmerek.BackColor = &H80000005
ttipe.BackColor = &H80000005
tplat.BackColor = &H80000005
End Sub

Sub aktif_cari()
tkode.Enabled = False
tmerek.Enabled = True
ttipe.Enabled = True
tplat.Enabled = True
btnCancel.Enabled = True
btnSimpan.Visible = False
btnupdate.Visible = True
btnHapus.Visible = True
btnHapus.Enabled = True
btnTutup.Visible = False
btnTambah.Enabled = False
tkode.BackColor = &H80000003
tmerek.BackColor = &H80000005
ttipe.BackColor = &H80000005
tplat.BackColor = &H80000005
End Sub

Private Sub btnCancel_Click()
Call nonaktif_all
End Sub

Private Sub btncari_Click()
If tcari.Text = "" Then
    MsgBox "Kolom pencarian tidak boleh kosong.", vbCritical, "Information"
    tcari.SetFocus
Else
    Call koneksi_db
    datamobil.Open "SELECT kode_mobil, merek_mobil, tipe_mobil, plat_mobil FROM tbl_mobil WHERE kode_mobil = '" & Replace(tcari.Text, " ", "") & "' ", konn
    If datamobil.EOF Then
       MsgBox "Kode mobil " + tcari.Text + " tidak ditemukan, harap masukan kode mobil dengan benar.", vbCritical, "Information"
    Else
        With datamobil
            tkode.Text = .Fields("kode_mobil")
            tmerek.Text = .Fields("merek_mobil")
            ttipe.Text = .Fields("tipe_mobil")
            tplat.Text = .Fields("plat_mobil")
            lblplat.Caption = .Fields("plat_mobil")
            Call aktif_cari
        End With
    End If
End If
End Sub

Private Sub btnHapus_Click()
msgdel = MsgBox("Anda akan menghapus kode mobil " + tkode.Text + " ?", vbCritical + vbYesNo, "Information")
If msgdel = vbYes Then
    Call koneksi_db
    datamobil.Open "DELETE FROM tbl_mobil WHERE kode_mobil = '" & tkode.Text & "' ", konn
    MsgBox "kode mobil " + tkode.Text + " berhasil dihapus.", vbInformation, "Information"
    Call nonaktif_all
    Call tampil_grid
End If

End Sub

Private Sub btnsimpan_Click()
Call koneksi_db
datamobil.Open "SELECT plat_mobil FROM tbl_mobil WHERE plat_mobil = '" & tplat.Text & "' ", konn
If Not datamobil.EOF Then
    MsgBox "Plat mobil " + tplat.Text + " telah ada didatabase, silahkan masukan plat mobil baru.", vbCritical, "Information"
Else
    If tmerek.Text = "" Then
        MsgBox "Merek mobil tidak boleh kosong", vbCritical, "Information"
    ElseIf ttipe.Text = "" Then
        MsgBox "Tipe mobil tidak boleh kosong", vbCritical, "Information"
    ElseIf tplat.Text = "" Then
        MsgBox "Merek mobil tidak boleh kosong", vbCritical, "Information"
    Else
        Call koneksi_db
        datamobil.Open "INSERT INTO tbl_mobil (kode_mobil, merek_mobil, tipe_mobil, plat_mobil) VALUES ('" & tkode.Text & "','" & tmerek.Text & "','" & ttipe.Text & "','" & Replace(tplat.Text, " ", "") & "')", konn
        MsgBox "Data mobil telah tersimpan.", vbInformation, "Information"
        Call nonaktif_all
        Call tampil_grid
    End If
End If




End Sub

Private Sub btnTambah_Click()
Call aktif_add
Call koneksi_db
datamobil.Open "SELECT KODE_MOBIL FROM tbl_mobil ORDER BY id_mobil DESC", konn
With datamobil
    If .BOF And .EOF Then
      tkode.Text = "SC" + "001"
    Else
       tkode.Text = "SC" + Right(Str(Val(Right(.Fields("KODE_MOBIL"), 3)) + 1001), 3)
    End If
End With
End Sub

Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub btnupdate_Click()
'Call koneksi_db
'datamobil.Open "SELECT plat_mobil FROM tbl_mobil WHERE plat_mobil = '" & tplat.Text & "'", konn
'If Not datamobil.EOF Then
'    MsgBox "Plat mobil " + tplat.Text + " telah ada didatabase, silahkan masukan plat mobil lain.", vbCritical, "Information"
'Else
    If tmerek.Text = "" Then
        MsgBox "Merek mobil tidak boleh kosong.", vbCritical, "Information"
    ElseIf ttipe.Text = "" Then
        MsgBox "Tipe mobil tidak boleh kosong.", vbCritical, "Information"
    ElseIf tplat.Text = "" Then
        MsgBox "Plat mobil tidak boleh kosong.", vbCritical, "Information"
    Else
        Call koneksi_db
        datamobil.Open "UPDATE tbl_mobil SET merek_mobil = '" & tmerek.Text & "', tipe_mobil = '" & ttipe.Text & "', plat_mobil = '" & tplat.Text & "' WHERE kode_mobil = '" & tkode.Text & "' ", konn
        MsgBox "Data mobil telah tersimpan.", vbInformation, "Information"
        Call nonaktif_all
        Call tampil_grid
    End If
'End If
End Sub

Private Sub Form_Load()
Call tampil_grid
Call nonaktif_all
End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub tmerek_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call btnsimpan_Click
End If
End Sub

Private Sub tplat_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call btnsimpan_Click
End If
End Sub

Private Sub ttipe_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call btnsimpan_Click
End If
End Sub
