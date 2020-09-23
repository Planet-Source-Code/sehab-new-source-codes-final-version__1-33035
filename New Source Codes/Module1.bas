Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
'Move
Public Const HTMOVE = 2

