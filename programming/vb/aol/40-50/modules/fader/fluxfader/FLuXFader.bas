Attribute VB_Name = "FLuXFader"

' ________________________________________________'
'|    FLuX Fader Bas [Version 1] By BaD & 007    |'
'|      For 32-Bit VB/API Programming            |'
'|       This BAS Was Made For AOL 40            |'
'|      This May Be Freely Distributed           |'
'|  Any Questions/Comments/Problems Email:       |'
'|           fluxaol@hotmail.com                 |'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
'All Faders Work with the SendChat Sub
'Call SendChat ("Font codes <B><I><S> etc." & Fader Name(text))





'** Windows 95 API Public Function Declarations **'
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'** Windows 95 API Public Functions Substitutes **'
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal Dest As Long, ByVal nCount&)
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal Dest&, ByVal nCount&)
Declare Sub ReleaseCapture Lib "user32" ()

'Windows 95 API Private Function & Sub Declarations'
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellUse Lib "shell32.dll Alias (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long" ()
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)


'  ** Public Windows 95 API Constant Functions **  '

'WindowsMessage()
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CLEAR = &H303
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203

'ListBox()
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_GETITEMDATA = &H199
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_DELETE = &H2E
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_CONTROL = &H11
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_CREATE = &H3

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SPI_SCREENSAVERRUNNING = 97

'GetWindow()
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5


Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'sndSoundPlay()
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

'ExitWindows()
Public Const EW_RESTARTWINDOWS = &H42
Public Const EW_REBOOTSYSTEM = &H43

'ShowWindow()
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_APPEND = &H100&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

'ErrorHandling()
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_INVALID_FUNCTION = 1&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_BAD_NETPATH = 53&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_INVALID_PASSWORDNAME = 1216&

Public Const GWL_STYLE = (-16)

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)

Public Const PROCESS_VM_READ = &H10

'OpenFile()
Private Const OF_READ = &H0
Private Const OF_WRITE = &H1
Private Const OF_READWRITE = &H2
Private Const OF_SHARE_COMBAT = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_PARSE = &H100
Private Const OF_DELETE = &H200
Private Const OF_VERIFY = &H400
Private Const OF_CANCEL = &H800
Private Const OF_CREATE = &H1000
Private Const OF_PROMPT = &H2000
Private Const OF_EXIST = &H4000
Private Const OF_REOPEN = &H8000

'SystemMetrics()
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CMENTRICS = 44


Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type OFSTRUCT
      cBytes As Byte
      fFixedByte As Byte
      nErrCode As Integer
      Reserved1 As Integer
      Reserved2 As Integer
      szPathName(128) As Byte
End Type


Global giBeepBox As Integer
Global r&
Global entry$
Global iniPath$
Function Fa_BlackBlue(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 0, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlackBlue = Msg

End Function



Function Fa_BlackGreen(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, F, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlackGreen = Msg

End Function

Function Fa_BlackPurple(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 0, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlackPurple = Msg

End Function


Function Fa_BlackRed2(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 0, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<B><Font Color=#" & h & ">" & d
Next b
Fa_BlackRed2 = Msg

End Function
Function Fa_BlackRed(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 0, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlackRed = Msg

End Function



Function Fa_BlackYellow(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, F, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlackYellow = Msg

End Function


Function Fa_BlueBlack(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, 0, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlueBlack = Msg

End Function

Function Fa_BlueGreen(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, F, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlueGreen = Msg

End Function



Function Fa_BluePurple(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255, 0, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BluePurple = Msg

End Function

Function Fa_BlueRed(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, 0, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_BlueRed = Msg

End Function

Function Fa_BlueYellow(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, F, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<b><Font Color=#" & h & ">" & d
Next b
Fa_BlueYellow = Msg

End Function

Function Fa_GreenBlack(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255 - F, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_GreenBlack = Msg

End Function



Function Fa_GreenBlue(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 255 - F, 0)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_GreenBlue = Msg

End Function
Function Fa_GreenPurple(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 255 - F, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_GreenPurple = Msg

End Function


Function Fa_GreenRed(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255 - F, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_GreenRed = Msg

End Function

Function Fa_GreenYellow(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_GreenYellow = Msg

End Function

Function Fa_Grey(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 220 / A
F = e * b
g = RGB(F, F, F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_Grey = Msg
End Function


Function Fa_PurpleBlack(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, 0, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_PurpleBlack = Msg

End Function

Function Fa_PurpleBlue(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255, 0, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_PurpleBlue = Msg
End Function

Function Fa_PurpleGreen(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, F, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_PurpleGreen = Msg
End Function


Function Fa_PurpleRed(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, 0, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_PurpleRed = Msg
End Function


Function Fa_PurpleYellow(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(255 - F, F, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_PurpleYellow = Msg

End Function

Function Fa_RedBlack(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 0, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_RedBlack = Msg

End Function


Function Fa_RedBlue(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 0, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_RedBlue = Msg

End Function

Function Fa_RedGreen(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, F, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_RedGreen = Msg

End Function


Function Fa_RedPurple(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 0, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_RedPurple = Msg

End Function

Function Fa_RedYellow(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, F, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_RedYellow = Msg

End Function

Function Fa_YellowBlack(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255 - F, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_YellowBlack = Msg

End Function



Function Fa_YellowBlue(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 255 - F, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_YellowBlue = Msg

End Function


Function Fa_YellowGreen(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255, 255 - F)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_YellowGreen = Msg

End Function


Function Fa_YellowPurple(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(F, 255 - F, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_YellowPurple = Msg

End Function

Function Fa_YellowRed(txt)


A = Len(txt)
For b = 1 To A
c = Left(txt, b)
d = Right(c, 1)
e = 255 / A
F = e * b
g = RGB(0, 255 - F, 255)
h = Hex(g)
i = Len(h)
If i = 5 Then h = "0" & h
If i = 4 Then h = "00" & h
If i = 3 Then h = "000" & h
If i = 2 Then h = "0000" & h
If i = 1 Then h = "00000" & h
Msg = Msg & "<Font Color=#" & h & ">" & d
Next b
Fa_YellowRed = Msg

End Function

Sub SendChat(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
Call SetFocusAPI(AORich%)
DoEvents
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "")
DoEvents
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)

End Sub

Function WavyFaderBlackRed1(txt)
A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#190000><sup>" & ab$ & "<FONT COLOR=#260000></sup>" & u$ & "<FONT COLOR=#3f0000><sub>" & S$ & "<FONT COLOR=#580000></sub>" & t$ & "<FONT COLOR=#720000><sup>" & Y$ & "<FONT COLOR=#8b0000></sup>" & L$ & "<FONT COLOR=#a50000><sub>" & F$ & "<FONT COLOR=#be0000></sub>" & b$ & "<FONT COLOR=#d70000><sup>" & c$ & "<FONT COLOR=#f10000></sup>" & d$ & "<FONT COLOR=#d70000><sub>" & h$ & "<FONT COLOR=#be0000></sub>" & j$ & "<FONT COLOR=#a50000><sup>" & k$ & "<FONT COLOR=#8b0000></sup>" & m$ & "<FONT COLOR=#720000><sub>" & n$ & "<FONT COLOR=#580000></sub>" & q$ & "<FONT COLOR=#3f0000><sup>" & V$ & "<FONT COLOR=#260000></sup>" & Z$
Next w
WavyFaderBlackRed1 = p$

End Function
Function WavyFaderBlackRed2(txt)
A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#190000><sup>" & ab$ & "<FONT COLOR=#260000></sup>" & u$ & "<FONT COLOR=#3f0000><sub>" & S$ & "<FONT COLOR=#580000></sub>" & t$ & "<FONT COLOR=#720000><sup>" & Y$ & "<FONT COLOR=#8b0000></sup>" & L$ & "<FONT COLOR=#a50000><sub>" & F$ & "<FONT COLOR=#be0000></sub>" & b$ & "<FONT COLOR=#d70000><sup>" & c$ & "<FONT COLOR=#f10000></sup>" & d$ & "<FONT COLOR=#d70000><sub>" & h$ & "<FONT COLOR=#be0000></sub>" & j$ & "<FONT COLOR=#a50000><sup>" & k$ & "<FONT COLOR=#8b0000></sup>" & m$ & "<FONT COLOR=#720000><sub>" & n$ & "<FONT COLOR=#580000></sub>" & q$ & "<FONT COLOR=#3f0000><sup>" & V$ & "<FONT COLOR=#260000></sup>" & Z$
Next w
WavyFaderBlackRed2 = p$

End Function

Function WavyFaderBlackYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#191900><sup>" & ab$ & "<FONT COLOR=#262600></sup>" & u$ & "<FONT COLOR=#3f3f00><sub>" & S$ & "<FONT COLOR=#585800></sub>" & t$ & "<FONT COLOR=#727200><sup>" & Y$ & "<FONT COLOR=#8b8b00></sup>" & L$ & "<FONT COLOR=#a5a500><sub>" & F$ & "<FONT COLOR=#bebe00></sub>" & b$ & "<FONT COLOR=#d7d700><sup>" & c$ & "<FONT COLOR=#f1f100></sup>" & d$ & "<FONT COLOR=#d7d700><sub>" & h$ & "<FONT COLOR=#bebe00></sub>" & j$ & "<FONT COLOR=#a5a500><sup>" & k$ & "<FONT COLOR=#8b8b00></sup>" & m$ & "<FONT COLOR=#727200><sub>" & n$ & "<FONT COLOR=#585800></sub>" & q$ & "<FONT COLOR=#3f3f00><sup>" & V$ & "<FONT COLOR=#262600></sup>" & Z$
Next w
WavyFaderBlackYellow1 = p$

End Function
Function WavyFaderBlackYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#191900><sup>" & ab$ & "<FONT COLOR=#262600></sup>" & u$ & "<FONT COLOR=#3f3f00><sub>" & S$ & "<FONT COLOR=#585800></sub>" & t$ & "<FONT COLOR=#727200><sup>" & Y$ & "<FONT COLOR=#8b8b00></sup>" & L$ & "<FONT COLOR=#a5a500><sub>" & F$ & "<FONT COLOR=#bebe00></sub>" & b$ & "<FONT COLOR=#d7d700><sup>" & c$ & "<FONT COLOR=#f1f100></sup>" & d$ & "<FONT COLOR=#d7d700><sub>" & h$ & "<FONT COLOR=#bebe00></sub>" & j$ & "<FONT COLOR=#a5a500><sup>" & k$ & "<FONT COLOR=#8b8b00></sup>" & m$ & "<FONT COLOR=#727200><sub>" & n$ & "<FONT COLOR=#585800></sub>" & q$ & "<FONT COLOR=#3f3f00><sup>" & V$ & "<FONT COLOR=#262600></sup>" & Z$
Next w
WavyFaderBlackYellow2 = p$

End Function

Function WavyFaderBlackPurple1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#190019><sup>" & ab$ & "<FONT COLOR=#260026></sup>" & u$ & "<FONT COLOR=#3f003f><sub>" & S$ & "<FONT COLOR=#580058></sub>" & t$ & "<FONT COLOR=#720072><sup>" & Y$ & "<FONT COLOR=#8b008b></sup>" & L$ & "<FONT COLOR=#a500a5><sub>" & F$ & "<FONT COLOR=#be00be></sub>" & b$ & "<FONT COLOR=#d700d7><sup>" & c$ & "<FONT COLOR=#f100f1></sup>" & d$ & "<FONT COLOR=#d700d7><sub>" & h$ & "<FONT COLOR=#be00be></sub>" & j$ & "<FONT COLOR=#a500a5><sup>" & k$ & "<FONT COLOR=#8b008b></sup>" & m$ & "<FONT COLOR=#720072><sub>" & n$ & "<FONT COLOR=#580058></sub>" & q$ & "<FONT COLOR=#3f003f><sup>" & V$ & "<FONT COLOR=#260026></sup>" & Z$
Next w
WavyFaderBlackPurple1 = p$

End Function
Function WavyFaderBlackPurple2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#190019><sup>" & ab$ & "<FONT COLOR=#260026></sup>" & u$ & "<FONT COLOR=#3f003f><sub>" & S$ & "<FONT COLOR=#580058></sub>" & t$ & "<FONT COLOR=#720072><sup>" & Y$ & "<FONT COLOR=#8b008b></sup>" & L$ & "<FONT COLOR=#a500a5><sub>" & F$ & "<FONT COLOR=#be00be></sub>" & b$ & "<FONT COLOR=#d700d7><sup>" & c$ & "<FONT COLOR=#f100f1></sup>" & d$ & "<FONT COLOR=#d700d7><sub>" & h$ & "<FONT COLOR=#be00be></sub>" & j$ & "<FONT COLOR=#a500a5><sup>" & k$ & "<FONT COLOR=#8b008b></sup>" & m$ & "<FONT COLOR=#720072><sub>" & n$ & "<FONT COLOR=#580058></sub>" & q$ & "<FONT COLOR=#3f003f><sup>" & V$ & "<FONT COLOR=#260026></sup>" & Z$
Next w
WavyFaderBlackPurple2 = p$

End Function

Function WavyFaderBlackBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000019><sup>" & ab$ & "<FONT COLOR=#000026></sup>" & u$ & "<FONT COLOR=#00003F><sub>" & S$ & "<FONT COLOR=#000058></sub>" & t$ & "<FONT COLOR=#000072><sup>" & Y$ & "<FONT COLOR=#00008B></sup>" & L$ & "<FONT COLOR=#0000A5><sub>" & F$ & "<FONT COLOR=#0000BE></sub>" & b$ & "<FONT COLOR=#0000D7><sup>" & c$ & "<FONT COLOR=#0000F1></sup>" & d$ & "<FONT COLOR=#0000D7><sub>" & h$ & "<FONT COLOR=#0000BE></sub>" & j$ & "<FONT COLOR=#0000A5><sup>" & k$ & "<FONT COLOR=#00008B></sup>" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & Z$
Next w
WavyFaderBlackBlue1 = p$

End Function
Function WavyFaderBlackBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000019><sup>" & ab$ & "<FONT COLOR=#000026></sup>" & u$ & "<FONT COLOR=#00003F><sub>" & S$ & "<FONT COLOR=#000058></sub>" & t$ & "<FONT COLOR=#000072><sup>" & Y$ & "<FONT COLOR=#00008B></sup>" & L$ & "<FONT COLOR=#0000A5><sub>" & F$ & "<FONT COLOR=#0000BE></sub>" & b$ & "<FONT COLOR=#0000D7><sup>" & c$ & "<FONT COLOR=#0000F1></sup>" & d$ & "<FONT COLOR=#0000D7><sub>" & h$ & "<FONT COLOR=#0000BE></sub>" & j$ & "<FONT COLOR=#0000A5><sup>" & k$ & "<FONT COLOR=#00008B></sup>" & m$ & "<FONT COLOR=#000072><sub>" & n$ & "<FONT COLOR=#000058></sub>" & q$ & "<FONT COLOR=#00003F><sup>" & V$ & "<FONT COLOR=#000026></sup>" & Z$
Next w
WavyFaderBlackBlue2 = p$

End Function




Function WavyFaderBlack1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000000></sup>" & u$ & "<FONT COLOR=#000000><sub>" & S$ & "<FONT COLOR=#000000></sub>" & t$ & "<FONT COLOR=#000000><sup>" & Y$ & "<FONT COLOR=#000000></sup>" & L$ & "<FONT COLOR=#000000><sub>" & F$ & "<FONT COLOR=#000000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000000><sub>" & h$ & "<FONT COLOR=#000000></sub>" & j$ & "<FONT COLOR=#000000><sup>" & k$ & "<FONT COLOR=#000000></sup>" & m$ & "<FONT COLOR=#000000><sub>" & n$ & "<FONT COLOR=#000000></sub>" & q$ & "<FONT COLOR=#000000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBlack1 = p$

End Function
Function WavyFaderBlack2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000000></sup>" & u$ & "<FONT COLOR=#000000><sub>" & S$ & "<FONT COLOR=#000000></sub>" & t$ & "<FONT COLOR=#000000><sup>" & Y$ & "<FONT COLOR=#000000></sup>" & L$ & "<FONT COLOR=#000000><sub>" & F$ & "<FONT COLOR=#000000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000000><sub>" & h$ & "<FONT COLOR=#000000></sub>" & j$ & "<FONT COLOR=#000000><sup>" & k$ & "<FONT COLOR=#000000></sup>" & m$ & "<FONT COLOR=#000000><sub>" & n$ & "<FONT COLOR=#000000></sub>" & q$ & "<FONT COLOR=#000000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBlack2 = p$

End Function
Function WavyFaderBlackGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#001900><sup>" & ab$ & "<FONT COLOR=#002600></sup>" & u$ & "<FONT COLOR=#003F00><sub>" & S$ & "<FONT COLOR=#005800></sub>" & t$ & "<FONT COLOR=#007200><sup>" & Y$ & "<FONT COLOR=#008B00></sup>" & L$ & "<FONT COLOR=#00A500><sub>" & F$ & "<FONT COLOR=#00BE00></sub>" & b$ & "<FONT COLOR=#00D700><sup>" & c$ & "<FONT COLOR=#00F100></sup>" & d$ & "<FONT COLOR=#00D700><sub>" & h$ & "<FONT COLOR=#00BE00></sub>" & j$ & "<FONT COLOR=#00A500><sup>" & k$ & "<FONT COLOR=#008B00></sup>" & m$ & "<FONT COLOR=#007200><sub>" & n$ & "<FONT COLOR=#005800></sub>" & q$ & "<FONT COLOR=#003F00><sup>" & V$ & "<FONT COLOR=#002600></sup>" & Z$
Next w
WavyFaderBlackGreen1 = p$

End Function

Function WavyFaderBlackGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#001900><sup>" & ab$ & "<FONT COLOR=#002600></sup>" & u$ & "<FONT COLOR=#003F00><sub>" & S$ & "<FONT COLOR=#005800></sub>" & t$ & "<FONT COLOR=#007200><sup>" & Y$ & "<FONT COLOR=#008B00></sup>" & L$ & "<FONT COLOR=#00A500><sub>" & F$ & "<FONT COLOR=#00BE00></sub>" & b$ & "<FONT COLOR=#00D700><sup>" & c$ & "<FONT COLOR=#00F100></sup>" & d$ & "<FONT COLOR=#00D700><sub>" & h$ & "<FONT COLOR=#00BE00></sub>" & j$ & "<FONT COLOR=#00A500><sup>" & k$ & "<FONT COLOR=#008B00></sup>" & m$ & "<FONT COLOR=#007200><sub>" & n$ & "<FONT COLOR=#005800></sub>" & q$ & "<FONT COLOR=#003F00><sup>" & V$ & "<FONT COLOR=#002600></sup>" & Z$
Next w
WavyFaderBlackGreen2 = p$

End Function




Function WavyFaderGreenBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#004400></sup>" & u$ & "<FONT COLOR=#008800><sub>" & S$ & "<FONT COLOR=#00cc00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00cc00></sup>" & L$ & "<FONT COLOR=#008800><sub>" & F$ & "<FONT COLOR=#004400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000044><sub>" & h$ & "<FONT COLOR=#000088></sub>" & j$ & "<FONT COLOR=#0000cc><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000cc><sub>" & n$ & "<FONT COLOR=#000088></sub>" & q$ & "<FONT COLOR=#000044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderGreenBlue1 = p$

End Function
Function WavyFaderGreenBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#00ee11></sup>" & u$ & "<FONT COLOR=#00cc33><sub>" & S$ & "<FONT COLOR=#009966></sub>" & t$ & "<FONT COLOR=#006699><sup>" & Y$ & "<FONT COLOR=#0033cc></sup>" & L$ & "<FONT COLOR=#0022dd><sub>" & F$ & "<FONT COLOR=#0011ee></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#0011ee><sub>" & h$ & "<FONT COLOR=#0022dd></sub>" & j$ & "<FONT COLOR=#0033cc><sup>" & k$ & "<FONT COLOR=#006699></sup>" & m$ & "<FONT COLOR=#009966><sub>" & n$ & "<FONT COLOR=#00cc33></sub>" & q$ & "<FONT COLOR=#00ee11><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreenBlue2 = p$

End Function
Function WavyFaderGreenPurple2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#11ee11></sup>" & u$ & "<FONT COLOR=#33cc33><sub>" & S$ & "<FONT COLOR=#669966></sub>" & t$ & "<FONT COLOR=#996699><sup>" & Y$ & "<FONT COLOR=#cc33cc></sup>" & L$ & "<FONT COLOR=#dd22dd><sub>" & F$ & "<FONT COLOR=#ee11ee></sub>" & b$ & "<FONT COLOR=#ff00ff><sup>" & c$ & "<FONT COLOR=#ff00ff></sup>" & d$ & "<FONT COLOR=#ee11ee><sub>" & h$ & "<FONT COLOR=#dd22dd></sub>" & j$ & "<FONT COLOR=#cc33cc><sup>" & k$ & "<FONT COLOR=#996699></sup>" & m$ & "<FONT COLOR=#669966><sub>" & n$ & "<FONT COLOR=#33cc33></sub>" & q$ & "<FONT COLOR=#11ee11><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreenPurple2 = p$

End Function

Function WavyFaderGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#00ff00></sup>" & u$ & "<FONT COLOR=#00ff00><sub>" & S$ & "<FONT COLOR=#00ff00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00ff00></sup>" & L$ & "<FONT COLOR=#00ff00><sub>" & F$ & "<FONT COLOR=#00ff00></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#00ff00><sub>" & h$ & "<FONT COLOR=#00ff00></sub>" & j$ & "<FONT COLOR=#00ff00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00ff00><sub>" & n$ & "<FONT COLOR=#00ff00></sub>" & q$ & "<FONT COLOR=#00ff00><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreen1 = p$

End Function
Function WavyFaderGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#00ff00></sup>" & u$ & "<FONT COLOR=#00ff00><sub>" & S$ & "<FONT COLOR=#00ff00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00ff00></sup>" & L$ & "<FONT COLOR=#00ff00><sub>" & F$ & "<FONT COLOR=#00ff00></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#00ff00><sub>" & h$ & "<FONT COLOR=#00ff00></sub>" & j$ & "<FONT COLOR=#00ff00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00ff00><sub>" & n$ & "<FONT COLOR=#00ff00></sub>" & q$ & "<FONT COLOR=#00ff00><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreen2 = p$

End Function

Function WavyFaderGreenRed1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#004400></sup>" & u$ & "<FONT COLOR=#008800><sub>" & S$ & "<FONT COLOR=#00cc00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00cc00></sup>" & L$ & "<FONT COLOR=#008800><sub>" & F$ & "<FONT COLOR=#004400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440000><sub>" & h$ & "<FONT COLOR=#880000></sub>" & j$ & "<FONT COLOR=#cc0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#cc0000><sub>" & n$ & "<FONT COLOR=#880000></sub>" & q$ & "<FONT COLOR=#440000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderGreenRed1 = p$

End Function
Function WavyFaderGreenRed2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#11ee00></sup>" & u$ & "<FONT COLOR=#33cc00><sub>" & S$ & "<FONT COLOR=#669900></sub>" & t$ & "<FONT COLOR=#996600><sup>" & Y$ & "<FONT COLOR=#cc3300></sup>" & L$ & "<FONT COLOR=#dd2200><sub>" & F$ & "<FONT COLOR=#ee1100></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ee1100><sub>" & h$ & "<FONT COLOR=#dd2200></sub>" & j$ & "<FONT COLOR=#cc3300><sup>" & k$ & "<FONT COLOR=#996600></sup>" & m$ & "<FONT COLOR=#669900><sub>" & n$ & "<FONT COLOR=#33cc00></sub>" & q$ & "<FONT COLOR=#11ee00><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreenRed2 = p$

End Function

Function WavyFaderGreenPurple1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#004400></sup>" & u$ & "<FONT COLOR=#008800><sub>" & S$ & "<FONT COLOR=#00cc00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00cc00></sup>" & L$ & "<FONT COLOR=#008800><sub>" & F$ & "<FONT COLOR=#004400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440044><sub>" & h$ & "<FONT COLOR=#880088></sub>" & j$ & "<FONT COLOR=#cc00cc><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#cc00cc><sub>" & n$ & "<FONT COLOR=#880088></sub>" & q$ & "<FONT COLOR=#440044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderGreenPurple1 = p$

End Function

Function WavyFaderGreenYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#004400></sup>" & u$ & "<FONT COLOR=#008800><sub>" & S$ & "<FONT COLOR=#00cc00></sub>" & t$ & "<FONT COLOR=#00ff00><sup>" & Y$ & "<FONT COLOR=#00cc00></sup>" & L$ & "<FONT COLOR=#008800><sub>" & F$ & "<FONT COLOR=#004400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#444400><sub>" & h$ & "<FONT COLOR=#888800></sub>" & j$ & "<FONT COLOR=#cccc00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#cccc00><sub>" & n$ & "<FONT COLOR=#888800></sub>" & q$ & "<FONT COLOR=#444400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderGreenYellow1 = p$

End Function
Function WavyFaderGreenYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#00ff00><sup>" & ab$ & "<FONT COLOR=#11ee00></sup>" & u$ & "<FONT COLOR=#22dd00><sub>" & S$ & "<FONT COLOR=#33cc00></sub>" & t$ & "<FONT COLOR=#44bb00><sup>" & Y$ & "<FONT COLOR=#55aa00></sup>" & L$ & "<FONT COLOR=#669900><sub>" & F$ & "<FONT COLOR=#778800></sub>" & b$ & "<FONT COLOR=#888800><sup>" & c$ & "<FONT COLOR=#888800></sup>" & d$ & "<FONT COLOR=#778800><sub>" & h$ & "<FONT COLOR=#669900></sub>" & j$ & "<FONT COLOR=#55aa00><sup>" & k$ & "<FONT COLOR=#44bb00></sup>" & m$ & "<FONT COLOR=#33cc00><sub>" & n$ & "<FONT COLOR=#22dd00></sub>" & q$ & "<FONT COLOR=#11ee00><sup>" & V$ & "<FONT COLOR=#00ff00></sup>" & Z$
Next w
WavyFaderGreenYellow2 = p$

End Function

Function WavyFaderPurpleRed1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440044></sup>" & u$ & "<FONT COLOR=#880088><sub>" & S$ & "<FONT COLOR=#cc00cc></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#cc00cc></sup>" & L$ & "<FONT COLOR=#880088><sub>" & F$ & "<FONT COLOR=#440044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440000><sub>" & h$ & "<FONT COLOR=#880000></sub>" & j$ & "<FONT COLOR=#cc0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#cc0000><sub>" & n$ & "<FONT COLOR=#880000></sub>" & q$ & "<FONT COLOR=#440000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderPurpleRed1 = p$

End Function
Function WavyFaderPurpleRed2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#880088><sup>" & ab$ & "<FONT COLOR=#880077></sup>" & u$ & "<FONT COLOR=#990066><sub>" & S$ & "<FONT COLOR=#aa0055></sub>" & t$ & "<FONT COLOR=#bb0044><sup>" & Y$ & "<FONT COLOR=#cc0033></sup>" & L$ & "<FONT COLOR=#dd0022><sub>" & F$ & "<FONT COLOR=#ee0011></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ee0011><sub>" & h$ & "<FONT COLOR=#dd0022></sub>" & j$ & "<FONT COLOR=#cc0033><sup>" & k$ & "<FONT COLOR=#bb0044></sup>" & m$ & "<FONT COLOR=#aa0055><sub>" & n$ & "<FONT COLOR=#990066></sub>" & q$ & "<FONT COLOR=#880077><sup>" & V$ & "<FONT COLOR=#880088></sup>" & Z$
Next w
WavyFaderPurpleRed2 = p$

End Function

Function WavyFaderPurpleBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440044></sup>" & u$ & "<FONT COLOR=#880088><sub>" & S$ & "<FONT COLOR=#cc00cc></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#cc00cc></sup>" & L$ & "<FONT COLOR=#880088><sub>" & F$ & "<FONT COLOR=#440044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000044><sub>" & h$ & "<FONT COLOR=#000088></sub>" & j$ & "<FONT COLOR=#0000cc><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000cc><sub>" & n$ & "<FONT COLOR=#000088></sub>" & q$ & "<FONT COLOR=#000044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderPurpleBlue1 = p$

End Function
Function WavyFaderPurpleBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#880088><sup>" & ab$ & "<FONT COLOR=#770088></sup>" & u$ & "<FONT COLOR=#660099><sub>" & S$ & "<FONT COLOR=#5500aa></sub>" & t$ & "<FONT COLOR=#4400bb><sup>" & Y$ & "<FONT COLOR=#3300cc></sup>" & L$ & "<FONT COLOR=#2200dd><sub>" & F$ & "<FONT COLOR=#1100ee></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#1100ee><sub>" & h$ & "<FONT COLOR=#2200dd></sub>" & j$ & "<FONT COLOR=#3300cc><sup>" & k$ & "<FONT COLOR=#4400bb></sup>" & m$ & "<FONT COLOR=#5500aa><sub>" & n$ & "<FONT COLOR=#660099></sub>" & q$ & "<FONT COLOR=#770088><sup>" & V$ & "<FONT COLOR=#880088></sup>" & Z$
Next w
WavyFaderPurpleBlue2 = p$

End Function


Function WavyFaderPurple1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff00ff><sup>" & ab$ & "<FONT COLOR=#ff00ff></sup>" & u$ & "<FONT COLOR=#ff00ff><sub>" & S$ & "<FONT COLOR=#ff00ff></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#ff00ff></sup>" & L$ & "<FONT COLOR=#ff00ff><sub>" & F$ & "<FONT COLOR=#ff00ff></sub>" & b$ & "<FONT COLOR=#ff00ff><sup>" & c$ & "<FONT COLOR=#ff00ff></sup>" & d$ & "<FONT COLOR=#ff00ff><sub>" & h$ & "<FONT COLOR=#ff00ff></sub>" & j$ & "<FONT COLOR=#ff00ff><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#ff00ff><sub>" & n$ & "<FONT COLOR=#ff00ff></sub>" & q$ & "<FONT COLOR=#ff00ff><sup>" & V$ & "<FONT COLOR=#ff00ff></sup>" & Z$
Next w
WavyFaderPurple1 = p$

End Function
Function WavyFaderPurple2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff00ff><sup>" & ab$ & "<FONT COLOR=#ff00ff></sup>" & u$ & "<FONT COLOR=#ff00ff><sub>" & S$ & "<FONT COLOR=#ff00ff></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#ff00ff></sup>" & L$ & "<FONT COLOR=#ff00ff><sub>" & F$ & "<FONT COLOR=#ff00ff></sub>" & b$ & "<FONT COLOR=#ff00ff><sup>" & c$ & "<FONT COLOR=#ff00ff></sup>" & d$ & "<FONT COLOR=#ff00ff><sub>" & h$ & "<FONT COLOR=#ff00ff></sub>" & j$ & "<FONT COLOR=#ff00ff><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#ff00ff><sub>" & n$ & "<FONT COLOR=#ff00ff></sub>" & q$ & "<FONT COLOR=#ff00ff><sup>" & V$ & "<FONT COLOR=#ff00ff></sup>" & Z$
Next w
WavyFaderPurple2 = p$

End Function

Function WavyFaderPurpleYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440044></sup>" & u$ & "<FONT COLOR=#880088><sub>" & S$ & "<FONT COLOR=#cc00cc></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#cc00cc></sup>" & L$ & "<FONT COLOR=#880088><sub>" & F$ & "<FONT COLOR=#440044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#444400><sub>" & h$ & "<FONT COLOR=#888800></sub>" & j$ & "<FONT COLOR=#cccc00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#cccc00><sub>" & n$ & "<FONT COLOR=#888800></sub>" & q$ & "<FONT COLOR=#444400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderPurpleYellow1 = p$

End Function
Function WavyFaderPurpleYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#880088><sup>" & ab$ & "<FONT COLOR=#881177></sup>" & u$ & "<FONT COLOR=#882266><sub>" & S$ & "<FONT COLOR=#883355></sub>" & t$ & "<FONT COLOR=#884444><sup>" & Y$ & "<FONT COLOR=#885533></sup>" & L$ & "<FONT COLOR=#886622><sub>" & F$ & "<FONT COLOR=#887711></sub>" & b$ & "<FONT COLOR=#888800><sup>" & c$ & "<FONT COLOR=#888800></sup>" & d$ & "<FONT COLOR=#887711><sub>" & h$ & "<FONT COLOR=#886622></sub>" & j$ & "<FONT COLOR=#885533><sup>" & k$ & "<FONT COLOR=#884444></sup>" & m$ & "<FONT COLOR=#883355><sub>" & n$ & "<FONT COLOR=#882266></sub>" & q$ & "<FONT COLOR=#881177><sup>" & V$ & "<FONT COLOR=#880088></sup>" & Z$
Next w
WavyFaderYellow2 = p$

End Function

Function WavyFaderPurpleGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440044></sup>" & u$ & "<FONT COLOR=#880088><sub>" & S$ & "<FONT COLOR=#cc00cc></sub>" & t$ & "<FONT COLOR=#ff00ff><sup>" & Y$ & "<FONT COLOR=#cc00cc></sup>" & L$ & "<FONT COLOR=#880088><sub>" & F$ & "<FONT COLOR=#440044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#004400><sub>" & h$ & "<FONT COLOR=#008800></sub>" & j$ & "<FONT COLOR=#00cc00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00cc00><sub>" & n$ & "<FONT COLOR=#008800></sub>" & q$ & "<FONT COLOR=#004400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderPurpleGreen1 = p$

End Function
Function WavyFaderPurpleGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff00ff><sup>" & ab$ & "<FONT COLOR=#ee11ee></sup>" & u$ & "<FONT COLOR=#cc33cc><sub>" & S$ & "<FONT COLOR=#996699></sub>" & t$ & "<FONT COLOR=#669966><sup>" & Y$ & "<FONT COLOR=#33cc33></sup>" & L$ & "<FONT COLOR=#22dd22><sub>" & F$ & "<FONT COLOR=#11ee11></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#11ee11><sub>" & h$ & "<FONT COLOR=#22dd22></sub>" & j$ & "<FONT COLOR=#33cc33><sup>" & k$ & "<FONT COLOR=#669966></sup>" & m$ & "<FONT COLOR=#996699><sub>" & n$ & "<FONT COLOR=#cc33cc></sub>" & q$ & "<FONT COLOR=#ee11ee><sup>" & V$ & "<FONT COLOR=#ff00ff></sup>" & Z$
Next w
WavyFaderPurpleGreen2 = p$

End Function

Function WavyFaderYellowBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#444400></sup>" & u$ & "<FONT COLOR=#888800><sub>" & S$ & "<FONT COLOR=#cccc00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#cccc00></sup>" & L$ & "<FONT COLOR=#888800><sub>" & F$ & "<FONT COLOR=#444400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000044><sub>" & h$ & "<FONT COLOR=#000088></sub>" & j$ & "<FONT COLOR=#0000cc><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000cc><sub>" & n$ & "<FONT COLOR=#000088></sub>" & q$ & "<FONT COLOR=#000044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderYellowBlue1 = p$

End Function
Function WavyFaderYellowBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#eeee11></sup>" & u$ & "<FONT COLOR=#cccc33><sub>" & S$ & "<FONT COLOR=#999966></sub>" & t$ & "<FONT COLOR=#666699><sup>" & Y$ & "<FONT COLOR=#3333cc></sup>" & L$ & "<FONT COLOR=#2222dd><sub>" & F$ & "<FONT COLOR=#1111ee></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#1111ee><sub>" & h$ & "<FONT COLOR=#2222dd></sub>" & j$ & "<FONT COLOR=#3333cc><sup>" & k$ & "<FONT COLOR=#666699></sup>" & m$ & "<FONT COLOR=#999966><sub>" & n$ & "<FONT COLOR=#cccc33></sub>" & q$ & "<FONT COLOR=#eeee11><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellowBlue2 = p$

End Function

Function WavyFaderYellowGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#444400></sup>" & u$ & "<FONT COLOR=#888800><sub>" & S$ & "<FONT COLOR=#cccc00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#cccc00></sup>" & L$ & "<FONT COLOR=#888800><sub>" & F$ & "<FONT COLOR=#444400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#004400><sub>" & h$ & "<FONT COLOR=#008800></sub>" & j$ & "<FONT COLOR=#00cc00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00cc00><sub>" & n$ & "<FONT COLOR=#008800></sub>" & q$ & "<FONT COLOR=#004400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderYellowGreen1 = p$

End Function
Function WavyFaderYellowGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#eeff00></sup>" & u$ & "<FONT COLOR=#ccff00><sub>" & S$ & "<FONT COLOR=#99ff00></sub>" & t$ & "<FONT COLOR=#66ff00><sup>" & Y$ & "<FONT COLOR=#33ff00></sup>" & L$ & "<FONT COLOR=#22ff00><sub>" & F$ & "<FONT COLOR=#11ff00></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#11ff00><sub>" & h$ & "<FONT COLOR=#22ff00></sub>" & j$ & "<FONT COLOR=#33ff00><sup>" & k$ & "<FONT COLOR=#66ff00></sup>" & m$ & "<FONT COLOR=#99ff00><sub>" & n$ & "<FONT COLOR=#ccff00></sub>" & q$ & "<FONT COLOR=#eeff00><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellowGreen2 = p$

End Function


Function WavyFaderYellowRed1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#444400></sup>" & u$ & "<FONT COLOR=#888800><sub>" & S$ & "<FONT COLOR=#cccc00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#cccc00></sup>" & L$ & "<FONT COLOR=#888800><sub>" & F$ & "<FONT COLOR=#444400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440000><sub>" & h$ & "<FONT COLOR=#880000></sub>" & j$ & "<FONT COLOR=#cc0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#cc0000><sub>" & n$ & "<FONT COLOR=#880000></sub>" & q$ & "<FONT COLOR=#440000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderYellowRed1 = p$

End Function
Function WavyFaderYellowRed2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#ffee00></sup>" & u$ & "<FONT COLOR=#ffcc00><sub>" & S$ & "<FONT COLOR=#ff9900></sub>" & t$ & "<FONT COLOR=#ff6600><sup>" & Y$ & "<FONT COLOR=#ff3300></sup>" & L$ & "<FONT COLOR=#ff2200><sub>" & F$ & "<FONT COLOR=#ff1100></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ff1100><sub>" & h$ & "<FONT COLOR=#ff2200></sub>" & j$ & "<FONT COLOR=#ff3300><sup>" & k$ & "<FONT COLOR=#ff6600></sup>" & m$ & "<FONT COLOR=#ff9900><sub>" & n$ & "<FONT COLOR=#ffcc00></sub>" & q$ & "<FONT COLOR=#ffee00><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellowRed2 = p$

End Function

Function WavyFaderYellowPurple1(txt)


A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#444400></sup>" & u$ & "<FONT COLOR=#888800><sub>" & S$ & "<FONT COLOR=#cccc00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#cccc00></sup>" & L$ & "<FONT COLOR=#888800><sub>" & F$ & "<FONT COLOR=#444400></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440044><sub>" & h$ & "<FONT COLOR=#880088></sub>" & j$ & "<FONT COLOR=#cc00cc><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#cc00cc><sub>" & n$ & "<FONT COLOR=#880088></sub>" & q$ & "<FONT COLOR=#440044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderYellowPurple1 = p$

End Function
Function WavyFaderYellowPurple2(txt)


A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#eeee11></sup>" & u$ & "<FONT COLOR=#ddcc33><sub>" & S$ & "<FONT COLOR=#cc9966></sub>" & t$ & "<FONT COLOR=#bb6699><sup>" & Y$ & "<FONT COLOR=#aa3399></sup>" & L$ & "<FONT COLOR=#992299><sub>" & F$ & "<FONT COLOR=#991199></sub>" & b$ & "<FONT COLOR=#990099><sup>" & c$ & "<FONT COLOR=#990099></sup>" & d$ & "<FONT COLOR=#991199><sub>" & h$ & "<FONT COLOR=#992299></sub>" & j$ & "<FONT COLOR=#aa3399><sup>" & k$ & "<FONT COLOR=#bb6699></sup>" & m$ & "<FONT COLOR=#cc9966><sub>" & n$ & "<FONT COLOR=#ddcc33></sub>" & q$ & "<FONT COLOR=#eeee11><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellowPurple2 = p$

End Function

Function WavyFaderYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#ffff00></sup>" & u$ & "<FONT COLOR=#ffff00><sub>" & S$ & "<FONT COLOR=#ffff00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#ffff00></sup>" & L$ & "<FONT COLOR=#ffff00><sub>" & F$ & "<FONT COLOR=#ffff00></sub>" & b$ & "<FONT COLOR=#ffff00><sup>" & c$ & "<FONT COLOR=#ffff00></sup>" & d$ & "<FONT COLOR=#ffff00><sub>" & h$ & "<FONT COLOR=#ffff00></sub>" & j$ & "<FONT COLOR=#ffff00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#ffff00><sub>" & n$ & "<FONT COLOR=#ffff00></sub>" & q$ & "<FONT COLOR=#ffff00><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellow1 = p$

End Function
Function WavyFaderYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ffff00><sup>" & ab$ & "<FONT COLOR=#ffff00></sup>" & u$ & "<FONT COLOR=#ffff00><sub>" & S$ & "<FONT COLOR=#ffff00></sub>" & t$ & "<FONT COLOR=#ffff00><sup>" & Y$ & "<FONT COLOR=#ffff00></sup>" & L$ & "<FONT COLOR=#ffff00><sub>" & F$ & "<FONT COLOR=#ffff00></sub>" & b$ & "<FONT COLOR=#ffff00><sup>" & c$ & "<FONT COLOR=#ffff00></sup>" & d$ & "<FONT COLOR=#ffff00><sub>" & h$ & "<FONT COLOR=#ffff00></sub>" & j$ & "<FONT COLOR=#ffff00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#ffff00><sub>" & n$ & "<FONT COLOR=#ffff00></sub>" & q$ & "<FONT COLOR=#ffff00><sup>" & V$ & "<FONT COLOR=#ffff00></sup>" & Z$
Next w
WavyFaderYellow2 = p$

End Function

Function WavyFaderBlueRed1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000044></sup>" & u$ & "<FONT COLOR=#000088><sub>" & S$ & "<FONT COLOR=#0000cc></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000cc></sup>" & L$ & "<FONT COLOR=#000088><sub>" & F$ & "<FONT COLOR=#000044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440000><sub>" & h$ & "<FONT COLOR=#880000></sub>" & j$ & "<FONT COLOR=#cc0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#cc0000><sub>" & n$ & "<FONT COLOR=#880000></sub>" & q$ & "<FONT COLOR=#440000><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBlueRed1 = p$

End Function
Function WavyFaderBlueRed2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#1100ee></sup>" & u$ & "<FONT COLOR=#3300cc><sub>" & S$ & "<FONT COLOR=#660099></sub>" & t$ & "<FONT COLOR=#990066><sup>" & Y$ & "<FONT COLOR=#cc0033></sup>" & L$ & "<FONT COLOR=#dd0022><sub>" & F$ & "<FONT COLOR=#ee0011></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ee0011><sub>" & h$ & "<FONT COLOR=#dd0022></sub>" & j$ & "<FONT COLOR=#cc0033><sup>" & k$ & "<FONT COLOR=#990066></sup>" & m$ & "<FONT COLOR=#660099><sub>" & n$ & "<FONT COLOR=#3300cc></sub>" & q$ & "<FONT COLOR=#1100ee><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBlueRed2 = p$

End Function

Function WavyFaderBluePurple1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000044></sup>" & u$ & "<FONT COLOR=#000088><sub>" & S$ & "<FONT COLOR=#0000cc></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000cc></sup>" & L$ & "<FONT COLOR=#000088><sub>" & F$ & "<FONT COLOR=#000044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440044><sub>" & h$ & "<FONT COLOR=#880088></sub>" & j$ & "<FONT COLOR=#cc00cc><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#cc00cc><sub>" & n$ & "<FONT COLOR=#880088></sub>" & q$ & "<FONT COLOR=#440044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBluePurple1 = p$

End Function
Function WavyFaderBluePurple2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#1100ee></sup>" & u$ & "<FONT COLOR=#2200dd><sub>" & S$ & "<FONT COLOR=#3300cc></sub>" & t$ & "<FONT COLOR=#4400bb><sup>" & Y$ & "<FONT COLOR=#5500aa></sup>" & L$ & "<FONT COLOR=#660099><sub>" & F$ & "<FONT COLOR=#770088></sub>" & b$ & "<FONT COLOR=#880088><sup>" & c$ & "<FONT COLOR=#880088></sup>" & d$ & "<FONT COLOR=#770088><sub>" & h$ & "<FONT COLOR=#660099></sub>" & j$ & "<FONT COLOR=#5500aa><sup>" & k$ & "<FONT COLOR=#4400bb></sup>" & m$ & "<FONT COLOR=#3300cc><sub>" & n$ & "<FONT COLOR=#2200dd></sub>" & q$ & "<FONT COLOR=#1100ee><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBluePurple2 = p$

End Function

Function WavyFaderBlueYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000044></sup>" & u$ & "<FONT COLOR=#000088><sub>" & S$ & "<FONT COLOR=#0000cc></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000cc></sup>" & L$ & "<FONT COLOR=#000088><sub>" & F$ & "<FONT COLOR=#000044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#444400><sub>" & h$ & "<FONT COLOR=#888800></sub>" & j$ & "<FONT COLOR=#cccc00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#cccc00><sub>" & n$ & "<FONT COLOR=#888800></sub>" & q$ & "<FONT COLOR=#444400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBlueYellow1 = p$

End Function
Function WavyFaderBlueYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#1111ee></sup>" & u$ & "<FONT COLOR=#3333cc><sub>" & S$ & "<FONT COLOR=#666699></sub>" & t$ & "<FONT COLOR=#999966><sup>" & Y$ & "<FONT COLOR=#cccc33></sup>" & L$ & "<FONT COLOR=#dddd22><sub>" & F$ & "<FONT COLOR=#eeee11></sub>" & b$ & "<FONT COLOR=#ffff00><sup>" & c$ & "<FONT COLOR=#ffff00></sup>" & d$ & "<FONT COLOR=#eeee11><sub>" & h$ & "<FONT COLOR=#dddd22></sub>" & j$ & "<FONT COLOR=#cccc33><sup>" & k$ & "<FONT COLOR=#999966></sup>" & m$ & "<FONT COLOR=#666699><sub>" & n$ & "<FONT COLOR=#3333cc></sub>" & q$ & "<FONT COLOR=#1111ee><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBlueYellow2 = p$

End Function

Function WavyFaderBlueGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#000044></sup>" & u$ & "<FONT COLOR=#000088><sub>" & S$ & "<FONT COLOR=#0000cc></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000cc></sup>" & L$ & "<FONT COLOR=#000088><sub>" & F$ & "<FONT COLOR=#000044></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#004400><sub>" & h$ & "<FONT COLOR=#008800></sub>" & j$ & "<FONT COLOR=#00cc00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00cc00><sub>" & n$ & "<FONT COLOR=#008800></sub>" & q$ & "<FONT COLOR=#004400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderBlueGreen1 = p$

End Function
Function WavyFaderBlueGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#1100ee></sup>" & u$ & "<FONT COLOR=#0033cc><sub>" & S$ & "<FONT COLOR=#006699></sub>" & t$ & "<FONT COLOR=#009966><sup>" & Y$ & "<FONT COLOR=#00cc33></sup>" & L$ & "<FONT COLOR=#00dd22><sub>" & F$ & "<FONT COLOR=#00ee11></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#00ee11><sub>" & h$ & "<FONT COLOR=#00dd22></sub>" & j$ & "<FONT COLOR=#00cc33><sup>" & k$ & "<FONT COLOR=#009966></sup>" & m$ & "<FONT COLOR=#006699><sub>" & n$ & "<FONT COLOR=#0033cc></sub>" & q$ & "<FONT COLOR=#0011ee><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBlueGreen2 = p$

End Function

Function WavyFaderBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#0000ff></sup>" & u$ & "<FONT COLOR=#0000ff><sub>" & S$ & "<FONT COLOR=#0000ff></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000ff></sup>" & L$ & "<FONT COLOR=#0000ff><sub>" & F$ & "<FONT COLOR=#0000ff></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#0000ff><sub>" & h$ & "<FONT COLOR=#0000ff></sub>" & j$ & "<FONT COLOR=#0000ff><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000ff><sub>" & n$ & "<FONT COLOR=#0000ff></sub>" & q$ & "<FONT COLOR=#0000ff><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBlue1 = p$

End Function
Function WavyFaderBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#0000ff><sup>" & ab$ & "<FONT COLOR=#0000ff></sup>" & u$ & "<FONT COLOR=#0000ff><sub>" & S$ & "<FONT COLOR=#0000ff></sub>" & t$ & "<FONT COLOR=#0000ff><sup>" & Y$ & "<FONT COLOR=#0000ff></sup>" & L$ & "<FONT COLOR=#0000ff><sub>" & F$ & "<FONT COLOR=#0000ff></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#0000ff><sub>" & h$ & "<FONT COLOR=#0000ff></sub>" & j$ & "<FONT COLOR=#0000ff><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000ff><sub>" & n$ & "<FONT COLOR=#0000ff></sub>" & q$ & "<FONT COLOR=#0000ff><sup>" & V$ & "<FONT COLOR=#0000ff></sup>" & Z$
Next w
WavyFaderBlue2 = p$

End Function


Function WavyFaderRedBlue1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440000></sup>" & u$ & "<FONT COLOR=#880000><sub>" & S$ & "<FONT COLOR=#cc0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#cc0000></sup>" & L$ & "<FONT COLOR=#880000><sub>" & F$ & "<FONT COLOR=#440000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#000044><sub>" & h$ & "<FONT COLOR=#000088></sub>" & j$ & "<FONT COLOR=#0000cc><sup>" & k$ & "<FONT COLOR=#0000ff></sup>" & m$ & "<FONT COLOR=#0000cc><sub>" & n$ & "<FONT COLOR=#000088></sub>" & q$ & "<FONT COLOR=#000044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderRedBlue1 = p$

End Function



Sub WavyFaderRedBlue2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ee0011></sup>" & u$ & "<FONT COLOR=#cc0033><sub>" & S$ & "<FONT COLOR=#990066></sub>" & t$ & "<FONT COLOR=#660099><sup>" & Y$ & "<FONT COLOR=#3300cc></sup>" & L$ & "<FONT COLOR=#2200dd><sub>" & F$ & "<FONT COLOR=#1100ee></sub>" & b$ & "<FONT COLOR=#0000ff><sup>" & c$ & "<FONT COLOR=#0000ff></sup>" & d$ & "<FONT COLOR=#1100ee><sub>" & h$ & "<FONT COLOR=#2200dd></sub>" & j$ & "<FONT COLOR=#3300cc><sup>" & k$ & "<FONT COLOR=#660099></sup>" & m$ & "<FONT COLOR=#990066><sub>" & n$ & "<FONT COLOR=#cc0033></sub>" & q$ & "<FONT COLOR=#ee0011><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRedBlue2 = p$

End Sub

Function WavyFaderRed1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ff0000></sup>" & u$ & "<FONT COLOR=#ff0000><sub>" & S$ & "<FONT COLOR=#ff0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#ff0000></sup>" & L$ & "<FONT COLOR=#ff0000><sub>" & F$ & "<FONT COLOR=#ff0000></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ff0000><sub>" & h$ & "<FONT COLOR=#ff0000></sub>" & j$ & "<FONT COLOR=#ff0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#ff0000><sub>" & n$ & "<FONT COLOR=#ff0000></sub>" & q$ & "<FONT COLOR=#ff0000><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRed1 = p$

End Function
Function WavyFaderRed2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ff0000></sup>" & u$ & "<FONT COLOR=#ff0000><sub>" & S$ & "<FONT COLOR=#ff0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#ff0000></sup>" & L$ & "<FONT COLOR=#ff0000><sub>" & F$ & "<FONT COLOR=#ff0000></sub>" & b$ & "<FONT COLOR=#ff0000><sup>" & c$ & "<FONT COLOR=#ff0000></sup>" & d$ & "<FONT COLOR=#ff0000><sub>" & h$ & "<FONT COLOR=#ff0000></sub>" & j$ & "<FONT COLOR=#ff0000><sup>" & k$ & "<FONT COLOR=#ff0000></sup>" & m$ & "<FONT COLOR=#ff0000><sub>" & n$ & "<FONT COLOR=#ff0000></sub>" & q$ & "<FONT COLOR=#ff0000><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRed2 = p$

End Function

Function WavyFaderRedGreen1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440000></sup>" & u$ & "<FONT COLOR=#880000><sub>" & S$ & "<FONT COLOR=#cc0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#cc0000></sup>" & L$ & "<FONT COLOR=#880000><sub>" & F$ & "<FONT COLOR=#440000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#004400><sub>" & h$ & "<FONT COLOR=#008800></sub>" & j$ & "<FONT COLOR=#00cc00><sup>" & k$ & "<FONT COLOR=#00ff00></sup>" & m$ & "<FONT COLOR=#00cc00><sub>" & n$ & "<FONT COLOR=#008800></sub>" & q$ & "<FONT COLOR=#004400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderGreen1 = p$

End Function
Function WavyFaderRedGreen2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ee1100></sup>" & u$ & "<FONT COLOR=#cc3300><sub>" & S$ & "<FONT COLOR=#996600></sub>" & t$ & "<FONT COLOR=#669900><sup>" & Y$ & "<FONT COLOR=#33cc00></sup>" & L$ & "<FONT COLOR=#22dd00><sub>" & F$ & "<FONT COLOR=#11ee00></sub>" & b$ & "<FONT COLOR=#00ff00><sup>" & c$ & "<FONT COLOR=#00ff00></sup>" & d$ & "<FONT COLOR=#11ee00><sub>" & h$ & "<FONT COLOR=#22dd00></sub>" & j$ & "<FONT COLOR=#33cc00><sup>" & k$ & "<FONT COLOR=#669900></sup>" & m$ & "<FONT COLOR=#996600><sub>" & n$ & "<FONT COLOR=#cc3300></sub>" & q$ & "<FONT COLOR=#ee1100><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRedGreen2 = p$

End Function

Function WavyFaderRedPurple2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ee0011></sup>" & u$ & "<FONT COLOR=#dd0022><sub>" & S$ & "<FONT COLOR=#cc0033></sub>" & t$ & "<FONT COLOR=#bb0044><sup>" & Y$ & "<FONT COLOR=#aa0055></sup>" & L$ & "<FONT COLOR=#990066><sub>" & F$ & "<FONT COLOR=#880077></sub>" & b$ & "<FONT COLOR=#880088><sup>" & c$ & "<FONT COLOR=#880088></sup>" & d$ & "<FONT COLOR=#880077><sub>" & h$ & "<FONT COLOR=#99066></sub>" & j$ & "<FONT COLOR=#aa0055><sup>" & k$ & "<FONT COLOR=#bb0044></sup>" & m$ & "<FONT COLOR=#cc0033><sub>" & n$ & "<FONT COLOR=#dd0022></sub>" & q$ & "<FONT COLOR=#ee0011><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRedPurple2 = p$

End Function

Sub WavyFaderRedPurple1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440000></sup>" & u$ & "<FONT COLOR=#880000><sub>" & S$ & "<FONT COLOR=#cc0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#cc0000></sup>" & L$ & "<FONT COLOR=#880000><sub>" & F$ & "<FONT COLOR=#440000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#440044><sub>" & h$ & "<FONT COLOR=#880088></sub>" & j$ & "<FONT COLOR=#cc00cc><sup>" & k$ & "<FONT COLOR=#ff00ff></sup>" & m$ & "<FONT COLOR=#cc00cc><sub>" & n$ & "<FONT COLOR=#880088></sub>" & q$ & "<FONT COLOR=#440044><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderRedPurple1 = p$

End Sub
Function WavyFaderRedYellow1(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#000000><sup>" & ab$ & "<FONT COLOR=#440000></sup>" & u$ & "<FONT COLOR=#880000><sub>" & S$ & "<FONT COLOR=#cc0000></sub>" & t$ & "<FONT COLOR=#ff0000><sup>" & Y$ & "<FONT COLOR=#cc0000></sup>" & L$ & "<FONT COLOR=#880000><sub>" & F$ & "<FONT COLOR=#440000></sub>" & b$ & "<FONT COLOR=#000000><sup>" & c$ & "<FONT COLOR=#000000></sup>" & d$ & "<FONT COLOR=#444400><sub>" & h$ & "<FONT COLOR=#888800></sub>" & j$ & "<FONT COLOR=#cccc00><sup>" & k$ & "<FONT COLOR=#ffff00></sup>" & m$ & "<FONT COLOR=#cccc00><sub>" & n$ & "<FONT COLOR=#888800></sub>" & q$ & "<FONT COLOR=#444400><sup>" & V$ & "<FONT COLOR=#000000></sup>" & Z$
Next w
WavyFaderRedYellow1 = p$

End Function
Function WavyFaderRedYellow2(txt)

A = Len(txt)
For w = 1 To A Step 18
    ab$ = Mid$(txt, w, 1)
    u$ = Mid$(txt, w + 1, 1)
    S$ = Mid$(txt, w + 2, 1)
    t$ = Mid$(txt, w + 3, 1)
    Y$ = Mid$(txt, w + 4, 1)
    L$ = Mid$(txt, w + 5, 1)
    F$ = Mid$(txt, w + 6, 1)
    b$ = Mid$(txt, w + 7, 1)
    c$ = Mid$(txt, w + 8, 1)
    d$ = Mid$(txt, w + 9, 1)
    h$ = Mid$(txt, w + 10, 1)
    j$ = Mid$(txt, w + 11, 1)
    k$ = Mid$(txt, w + 12, 1)
    m$ = Mid$(txt, w + 13, 1)
    n$ = Mid$(txt, w + 14, 1)
    q$ = Mid$(txt, w + 15, 1)
    V$ = Mid$(txt, w + 16, 1)
    Z$ = Mid$(txt, w + 17, 1)
    p$ = p$ & "<FONT COLOR=#ff0000><sup>" & ab$ & "<FONT COLOR=#ee1100></sup>" & u$ & "<FONT COLOR=#dd2200><sub>" & S$ & "<FONT COLOR=#cc3300></sub>" & t$ & "<FONT COLOR=#bb4400><sup>" & Y$ & "<FONT COLOR=#aa5500></sup>" & L$ & "<FONT COLOR=#996600><sub>" & F$ & "<FONT COLOR=#887700></sub>" & b$ & "<FONT COLOR=#888800><sup>" & c$ & "<FONT COLOR=#888800></sup>" & d$ & "<FONT COLOR=#887700><sub>" & h$ & "<FONT COLOR=#996600></sub>" & j$ & "<FONT COLOR=#aa5500><sup>" & k$ & "<FONT COLOR=#bb4400></sup>" & m$ & "<FONT COLOR=#cc3300><sub>" & n$ & "<FONT COLOR=#dd2200></sub>" & q$ & "<FONT COLOR=#ee1100><sup>" & V$ & "<FONT COLOR=#ff0000></sup>" & Z$
Next w
WavyFaderRedYellow2 = p$

End Function

