VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Sub DisableCloseButton(ByVal Frm As Form)
    Dim h As Long
    h = GetSystemMenu(Frm.hwnd, 0)
    RemoveMenu h, 6, &H400
    RemoveMenu h, 5, &H400
End Sub

Public Function nmlevel(ByVal s As Variant)
If s = 1 Then
    nmlevel = "Administrator"
ElseIf s = 2 Then
    nmlevel = "User Pengelola"
End If
End Function

Public Function lvlstrtoint(ByVal a As Variant)
If a = "Administrator" Then
    lvlstrtoint = 1
Else
    lvlstrtoint = 2
End If
End Function
