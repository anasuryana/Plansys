Attribute VB_Name = "FinsDevDrag"
'Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam _
'As Long, lParam As Any) As Long

Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

