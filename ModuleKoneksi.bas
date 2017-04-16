Attribute VB_Name = "ModuleKoneksi"
Public konn As New ADODB.Connection
Public datauser As New ADODB.Recordset
Public historylogin As New ADODB.Recordset
Public datamobil As New ADODB.Recordset
Public databiaya As New ADODB.Recordset
Public datadurasilatihan As New ADODB.Recordset
Public datasiswa As New ADODB.Recordset
Public dataregistrasi As New ADODB.Recordset
Public datajadwal As New ADODB.Recordset

Sub koneksi_db()
'On Error Resume Next
Set konn = New ADODB.Connection
Set datauser = New ADODB.Recordset
Set historylogin = New ADODB.Recordset
Set datamobil = New ADODB.Recordset
Set databiaya = New ADODB.Recordset
Set datadurasilatihan = New ADODB.Recordset
Set datasiswa = New ADODB.Recordset
Set dataregistrasi = New ADODB.Recordset
Set datajadwal = New ADODB.Recordset
konn.ConnectionString = "DSN=mobil"
konn.Open
'If err <> 0 Then
'    MsgBox ("Harap cek koneksi Database MySql anda.")
'End If
'On Error GoTo 0
End Sub
