Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Function ExHandle(Text As String) As String
    ' This function edits the class names of
    ' windows so we can use them as variables
    Dim i As Integer, i2 As Integer, t$
    i = Len(Text)
    Text = LCase$(Trim(Text))
    For i2 = 1 To i
        Select Case Mid(Text, i2, 1)
        Case Chr$(Asc("a")) To Chr$(Asc("z"))
            t$ = t$ & Mid(Text, i2, 1)
        End Select
    Next
    ExHandle = t$
End Function

Public Sub StayOnTop(Frm As Form)
Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
