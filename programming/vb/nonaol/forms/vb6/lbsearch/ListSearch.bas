Attribute VB_Name = "ListSearch"
Option Explicit

Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Global Const LB_FINDSTRINGEXACT = &H1A2&

