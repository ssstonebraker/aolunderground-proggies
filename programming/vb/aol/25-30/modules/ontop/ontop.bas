Attribute VB_Name = "Module2"
'*********************************************************
'|                                                       |
'|                The OnTop.BAS File                     |
'|          Original Code By: Garkon 667@aol.com         |
'|                 This is Freeware                      |
'|                     ~ Garkon ~                        |
'*********************************************************

Declare Sub SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%)

Sub StayOnTop(f As Form)
SetWindowPos f.hWnd, HWND_TOPMOST, -1, -1, -1, -1, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

