Attribute VB_Name = "OnTop"
'-=OnTop.bas=-
'-=By: EcCo=-

'-=Declartions=-
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'-=Constants=-
Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'-=Subs=-
Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
