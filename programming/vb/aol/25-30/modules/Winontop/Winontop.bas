'Declarations for SetWindowPos

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

'This API Function will put the Window (hWnd)
'on top of all others. Similar to Windows Clock
'See frmPalette's Resize event for usage

Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer

