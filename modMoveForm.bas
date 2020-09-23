Attribute VB_Name = "modMoveForm"
Option Explicit

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub EasyMove(frm As Form)
  If frm.WindowState <> vbMaximized Then
    ReleaseCapture
    SendMessage frm.hwnd, &HA1, 2, 0&
  End If
End Sub
