VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_Databiaya 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".:: MASTER DATA BIAYA ::."
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "DATA BIAYA"
      BeginProperty Font 
         Name            =   "Ubuntu Mono"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5880
      TabIndex        =   19
      Top             =   0
      Width           =   7575
      Begin MSDataGridLib.DataGrid gridbiaya 
         Height          =   3615
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   6376
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
      Begin VB.CommandButton btnCari 
         Caption         =   "CARI KODE BIAYA"
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
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Width           =   2175
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
         Left            =   3000
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
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
         Left            =   4440
         TabIndex        =   18
         Top             =   3960
         Width           =   1095
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
         Left            =   4440
         TabIndex        =   17
         Top             =   3960
         Width           =   1095
      End
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
         Left            =   2760
         TabIndex        =   16
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton btnUpdate 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton BtnSimpan 
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
         Left            =   1200
         TabIndex        =   14
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton btnTambah 
         Caption         =   "TAMBAH BIAYA"
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
         TabIndex        =   13
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox tsima 
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
         Left            =   2520
         TabIndex        =   12
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox tdaftar 
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
         Left            =   2520
         TabIndex        =   11
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox tbiaya 
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
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
         Width           =   2895
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
         Left            =   2520
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox tnmtingkat 
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
         Left            =   2520
         TabIndex        =   8
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox tkdbiaya 
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
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lbluporadd 
         Caption         =   "#UPORADD"
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   4200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "PEMBUATAN SIM A"
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
         Top             =   3400
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "PENDAFTARAN"
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
         Top             =   2800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "BIAYA PAKET"
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
         Top             =   2200
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   1600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "NAMA TINGKATAN"
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
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "KODE BIAYA"
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
         Top             =   400
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form_Databiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub nonaktif_whenloaded()
tkdbiaya.Enabled = False
tnmtingkat.Enabled = False
ttemu.Enabled = False
tbiaya.Enabled = False
tdaftar.Enabled = False
tsima.Enabled = False
btnUpdate.Enabled = False
btnHapus.Enabled = False
btnCancel.Enabled = False
tkdbiaya.BackColor = &H80000003
tnmtingkat.BackColor = &H80000003
ttemu.BackColor = &H80000003
tbiaya.BackColor = &H80000003
tdaftar.BackColor = &H80000003
tsima.BackColor = &H80000003
End Sub

Sub aktif_aftertambah()
btnTambah.Enabled = False
'btnUpdate.Enabled = True
btnUpdate.Visible = False
btnHapus.Visible = False
BtnSimpan.Visible = True
btnCancel.Visible = True
btnCancel.Enabled = True
tkdbiaya.Enabled = False
tkdbiaya.BackColor = &H80000003
tnmtingkat.BackColor = &H80000005
ttemu.BackColor = &H80000005
tbiaya.BackColor = &H80000005
tdaftar.BackColor = &H80000005
tsima.BackColor = &H80000005
tkdbiaya.Enabled = True
tnmtingkat.Enabled = True
ttemu.Enabled = True
tbiaya.Enabled = True
tdaftar.Enabled = True
tsima.Enabled = True
End Sub

Sub nonaktif_aftercancel()
btnTambah.Enabled = True
btnUpdate.Enabled = False
btnUpdate.Visible = True
btnHapus.Visible = True
BtnSimpan.Visible = False
btnHapus.Visible = False
btnCancel.Enabled = False
tkdbiaya.Enabled = False
tkdbiaya.BackColor = &H80000003
tnmtingkat.BackColor = &H80000003
ttemu.BackColor = &H80000003
tbiaya.BackColor = &H80000003
tdaftar.BackColor = &H80000003
tsima.BackColor = &H80000003
tkdbiaya.Text = ""
tnmtingkat.Text = ""
ttemu.Text = ""
tbiaya.Text = ""
tdaftar.Text = ""
tsima.Text = ""
tcari.Text = ""
If lbluporadd.Caption = 2 Then
    btnTutup.Visible = True
End If
End Sub

Sub nonaktif_afterupdate()
btnTambah.Enabled = True
btnUpdate.Visible = True
btnUpdate.Enabled = False
btnHapus.Visible = True
btnHapus.Enabled = False
BtnSimpan.Visible = False
btnCancel.Visible = False
tkdbiaya.Enabled = False
tkdbiaya.BackColor = &H80000003
tnmtingkat.BackColor = &H80000003
ttemu.BackColor = &H80000003
tbiaya.BackColor = &H80000003
tdaftar.BackColor = &H80000003
tsima.BackColor = &H80000003
tkdbiaya.Text = ""
tnmtingkat.Text = ""
ttemu.Text = ""
tbiaya.Text = ""
tdaftar.Text = ""
tsima.Text = ""
tcari.Text = ""
lbluporadd.Caption = ""
End Sub

Sub nonaktif_afterhapus()
btnTambah.Enabled = True
btnUpdate.Visible = True
btnUpdate.Enabled = False
btnHapus.Visible = True
btnHapus.Enabled = False
BtnSimpan.Visible = False
btnCancel.Visible = True
btnCancel.Enabled = False
btnTutup.Visible = True
tkdbiaya.Enabled = False
tkdbiaya.BackColor = &H80000003
tnmtingkat.BackColor = &H80000003
ttemu.BackColor = &H80000003
tbiaya.BackColor = &H80000003
tdaftar.BackColor = &H80000003
tsima.BackColor = &H80000003
tkdbiaya.Text = ""
tnmtingkat.Text = ""
ttemu.Text = ""
tbiaya.Text = ""
tdaftar.Text = ""
tsima.Text = ""
tcari.Text = ""
lbluporadd.Caption = ""
End Sub


Sub tampil_grid()
Call koneksi_db
databiaya.CursorLocation = adUseClient
databiaya.CursorType = adOpenKeyset
databiaya.LockType = adLockOptimistic
databiaya.Open "SELECT KODE_PAKET, nama_paket AS PAKET, paket_pertemuan AS PERTEMUAN, biaya_paket AS BIAYA_PAKET, biaya_daftar AS BIAYA_DAFTAR, biaya_sim AS BIAYA_SIM FROM tbl_biaya_paket ORDER BY id_biaya_paket DESC", konn
Set gridbiaya.DataSource = databiaya

gridbiaya.Columns(0).Width = 1150
gridbiaya.Columns(1).Width = 1100
gridbiaya.Columns(2).Width = 1100
gridbiaya.Columns(3).Width = 1200
gridbiaya.Columns(4).Width = 1270
gridbiaya.Columns(5).Width = 1200
gridbiaya.AllowDelete = False
gridbiaya.AllowUpdate = False
gridbiaya.Refresh
End Sub

Sub aktif_aftercari()
tnmtingkat.BackColor = &H80000005
ttemu.BackColor = &H80000005
tbiaya.BackColor = &H80000005
tdaftar.BackColor = &H80000005
tsima.BackColor = &H80000005
btnTambah.Enabled = False
btnUpdate.Enabled = True
btnHapus.Enabled = True
btnHapus.Visible = True
btnTutup.Visible = False
btnCancel.Enabled = True
tnmtingkat.Enabled = True
ttemu.Enabled = True
tbiaya.Enabled = True
tdaftar.Enabled = True
tsima.Enabled = True
End Sub

Private Sub btnCari_Click()
If tcari.Text = "" Then
    MsgBox "Kolom pencarian tidak boleh kosong.", vbCritical, "Information"
Else
    Call koneksi_db
    databiaya.Open "SELECT kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar, biaya_sim FROM tbl_biaya_paket WHERE kode_paket = '" & tcari.Text & "' LIMIT 1", konn
    If databiaya.EOF Then
        MsgBox "Kode paket " + tcari.Text + " tidak ditemukan.", vbCritical, "Information"
    Else
        With databiaya
            tkdbiaya.Text = .Fields("kode_paket")
            tnmtingkat.Text = .Fields("nama_paket")
            ttemu.Text = .Fields("paket_pertemuan")
            tbiaya.Text = .Fields("biaya_paket")
            tdaftar.Text = .Fields("biaya_daftar")
            tsima.Text = .Fields("biaya_sim")
            Call aktif_aftercari
        End With
    End If
End If
lbluporadd.Caption = 2
End Sub

Private Sub btnHapus_Click()
msgdel = MsgBox("Anda yakin ingin menghapus kode paket " + tkdbiaya.Text + " ?", vbCritical + vbYesNo, "Information")
If msgdel = vbYes Then
    Call koneksi_db
    databiaya.Open "DELETE FROM tbl_biaya_paket WHERE kode_paket = '" & tkdbiaya.Text & "' ", konn
    MsgBox "Kode paket " + tkdbiaya.Text + " telah berhasil dihapus", vbInformation, "Information"
    Call tampil_grid
    Call nonaktif_afterhapus
End If
End Sub

Private Sub BtnSimpan_Click()
If tnmtingkat.Text = "" Then
    MsgBox "Nama tingkatan paket tidak boleh kosong", vbCritical, "Information"
    tnmtingkat.SetFocus
ElseIf ttemu.Text = "" Then
    MsgBox "Jumlah pertemuan paket tidak boleh kosong", vbCritical, "Information"
    ttemu.SetFocus
ElseIf tbiaya.Text = "" Then
    MsgBox "Biaya paket tidak boleh kosong", vbCritical, "Information"
    tbiaya.SetFocus
ElseIf tdaftar.Text = "" Then
    MsgBox "Biaya pendaftaran paket tidak boleh kosong", vbCritical, "Information"
    tdaftar.SetFocus
ElseIf tsima.Text = "" Then
    MsgBox "Biaya pembuatan SIM A tidak boleh kosong", vbCritical, "Information"
    tsima.SetFocus
Else
    Call koneksi_db
    databiaya.Open "INSERT INTO tbl_biaya_paket (kode_paket, nama_paket, paket_pertemuan, biaya_paket, biaya_daftar, biaya_sim) VALUES ('" & tkdbiaya.Text & "','" & tnmtingkat.Text & "','" & ttemu.Text & "','" & tbiaya.Text & "','" & tdaftar.Text & "','" & tsima.Text & "')", konn
    MsgBox "Paket " + tnmtingkat.Text + " telah berhasil disimpan.", vbInformation, "Information"
    Call nonaktif_aftercancel
    Call tampil_grid
End If
End Sub

Private Sub btnUpdate_Click()
If tnmtingkat.Text = "" Then
    MsgBox "Nama tingkatan paket tidak boleh kosong", vbCritical, "Information"
    tnmtingkat.SetFocus
ElseIf ttemu.Text = "" Then
    MsgBox "Jumlah pertemuan paket tidak boleh kosong", vbCritical, "Information"
    ttemu.SetFocus
ElseIf tbiaya.Text = "" Then
    MsgBox "Biaya paket tidak boleh kosong", vbCritical, "Information"
    tbiaya.SetFocus
ElseIf tdaftar.Text = "" Then
    MsgBox "Biaya pendaftaran paket tidak boleh kosong", vbCritical, "Information"
    tdaftar.SetFocus
ElseIf tsima.Text = "" Then
    MsgBox "Biaya pembuatan SIM A tidak boleh kosong", vbCritical, "Information"
    tsima.SetFocus
Else
    Call koneksi_db
    databiaya.Open "UPDATE tbl_biaya_paket SET nama_paket = '" & tnmtingkat.Text & "', paket_pertemuan = '" & ttemu.Text & "', biaya_paket = '" & tbiaya.Text & "', biaya_daftar = '" & tdaftar.Text & "', biaya_sim = '" & tsima.Text & "' WHERE kode_paket = '" & tkdbiaya.Text & "'", konn
    MsgBox "Paket " + tnmtingkat.Text + " telah berhasil dirubah.", vbInformation, "Information"
    Call nonaktif_afterupdate
    Call tampil_grid
End If
End Sub

Private Sub btnTambah_Click()
Call aktif_aftertambah
Call koneksi_db
datamobil.Open "SELECT kode_paket FROM tbl_biaya_paket ORDER BY id_biaya_paket DESC", konn
With datamobil
    If .BOF And .EOF Then
      tkdbiaya.Text = "PSC" + "001"
    Else
       tkdbiaya.Text = "PSC" + Right(Str(Val(Right(.Fields("kode_paket"), 3)) + 1001), 3)
    End If
End With
lbluporadd.Caption = 1
End Sub

Private Sub btnCancel_Click()
Call nonaktif_aftercancel
lbluporadd.Caption = ""
End Sub

Private Sub btnTutup_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call nonaktif_whenloaded
Call tampil_grid
End Sub

Private Sub tbiaya_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lbluporadd.Caption = 1 Then
       Call BtnSimpan_Click
    Else
        Call btnUpdate_Click
    End If
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If

End Sub

Private Sub tcari_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tdaftar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lbluporadd.Caption = 1 Then
       Call BtnSimpan_Click
    Else
        Call btnUpdate_Click
    End If
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub

Private Sub tnmtingkat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lbluporadd.Caption = 1 Then
       Call BtnSimpan_Click
    Else
        Call btnUpdate_Click
    End If
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tsima_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lbluporadd.Caption = 1 Then
       Call BtnSimpan_Click
    Else
        Call btnUpdate_Click
    End If
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If

End Sub

Private Sub ttemu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lbluporadd.Caption = 1 Then
       Call BtnSimpan_Click
    Else
        Call btnUpdate_Click
    End If
End If
If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
    KeyAscii = 0
End If
End Sub
