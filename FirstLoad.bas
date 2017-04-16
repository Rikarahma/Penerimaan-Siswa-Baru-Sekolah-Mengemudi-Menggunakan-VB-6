Attribute VB_Name = "FirstLoad"
Sub main()
'Dim cek As Integer
If Form_utama.sbmenu.Panels(2) = "" Then
    Load Form_login
    Form_login.Show
Else
    Load Form_utama
    Form_utama.Show
End If
End Sub
