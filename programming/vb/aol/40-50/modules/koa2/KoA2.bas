Attribute VB_Name = "KoA2"
'Sup? This .bas was made by KoA.
'This .bas is my first .bas, KoA.bas, combined with
'Cryofade.bas. Don't give me any credit for the fader
'part. Thanks Cryo!
'You can e-mail me at: koa@programmingworld.com
'You can visit my site at: http://www.programmingworld.com
'Copyright ©1998 Programming World™
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

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
Function BlackBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlue = Msg
End Function

Function BlackGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreen = Msg
End Function

Function BlackGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGrey = Msg
End Function

Function BlackPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurple = Msg
End Function

Function BlackRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRed = Msg
End Function

Function BlackYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellow = Msg
End Function

Function BlueBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlack = Msg
End Function

Function BlueGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreen = Msg
End Function

Function BluePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurple = Msg
End Function

Function BlueRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRed = Msg
End Function

Function BlueYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellow = Msg
End Function

Function GreenBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlack = Msg
End Function

Function GreenBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlue = Msg
End Function

Function GreenPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurple = Msg
End Function

Function GreenRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRed = Msg
End Function

Function GreenYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellow = Msg
End Function

Function GreyBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlack = Msg
End Function

Function GreyBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlue = Msg
End Function

Function GreyGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreen = Msg
End Function

Function GreyPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurple = Msg
End Function

Function GreyRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRed = Msg
End Function

Function GreyYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellow = Msg
End Function

Function PurpleBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlack = Msg
End Function

Function PurpleBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlue = Msg
End Function

Function PurpleGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreen = Msg
End Function

Function PurpleRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRed = Msg
End Function

Function PurpleYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellow = Msg
End Function

Function RedBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = Msg
End Function

Function RedBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = Msg
End Function

Function RedGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = Msg
End Function

Function RedPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurple = Msg
End Function

Function RedYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = Msg
End Function

Function YellowBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlack = Msg
End Function

Function YellowBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlue = Msg
End Function

Function YellowGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreen = Msg
End Function

Function YellowPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurple = Msg
End Function

Function YellowRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRed = Msg
End Function


'Pre-set 3 Color fade combinations begin here


Function BlackBlueBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlueBlack = Msg
End Function

Function BlackGreenBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreenBlack = Msg
End Function

Function BlackGreyBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreyBlack = Msg
End Function

Function BlackPurpleBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurpleBlack = Msg
End Function

Function BlackRedBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = Msg
End Function

Function BlackYellowBlack(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellowBlack = Msg
End Function

Function BlueBlackBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlackBlue = Msg
End Function

Function BlueGreenBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreenBlue = Msg
End Function

Function BluePurpleBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurpleBlue = Msg
End Function

Function BlueRedBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRedBlue = Msg
End Function

Function BlueYellowBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellowBlue = Msg
End Function

Function GreenBlackGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlackGreen = Msg
End Function

Function GreenBlueGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlueGreen = Msg
End Function

Function GreenPurpleGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurpleGreen = Msg
End Function

Function GreenRedGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRedGreen = Msg
End Function

Function GreenYellowGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellowGreen = Msg
End Function

Function GreyBlackGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlackGrey = Msg
End Function

Function GreyBlueGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlueGrey = Msg
End Function

Function GreyGreenGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreenGrey = Msg
End Function

Function GreyPurpleGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurpleGrey = Msg
End Function

Function GreyRedGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRedGrey = Msg
End Function

Function GreyYellowGrey(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellowGrey = Msg
End Function

Function PurpleBlackPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlackPurple = Msg
End Function

Function PurpleBluePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBluePurple = Msg
End Function

Function PurpleGreenPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreenPurple = Msg
End Function

Function PurpleRedPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRedPurple = Msg
End Function

Function PurpleYellowPurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellowPurple = Msg
End Function

Function RedBlackRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = Msg
End Function

Function RedBlueRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = Msg
End Function

Function RedGreenRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = Msg
End Function

Function RedPurpleRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = Msg
End Function

Function RedYellowRed(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = Msg
End Function

Function YellowBlackYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRedYellow = Msg
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    A = Hex(RGB)
    B = Len(A)
    If B = 5 Then A = "0" & A
    If B = 4 Then A = "00" & A
    If B = 3 Then A = "000" & A
    If B = 2 Then A = "0000" & A
    If B = 1 Then A = "00000" & A
    RGBtoHEX = A
End Function


'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
'Copyright ©1998 Programming World™
Sub killwin(windo)
'kills the window that you put in
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub
Function LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        Where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, Where + Len(look$))
        LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function

Sub AddRoomToListBox(ListBox As ListBox)
'Adds the room to a listbox
'Ex: Call AddRoomToListBox(List1)

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = UserSN Then GoTo Na
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub AntiIdle()
'This is used for idle bots.
Do
DoEvents
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Loop
End Sub
Sub Attention(thetext As String)
'used for attention bots
SendChat ("¤ ¤ ¤ ¤ ATTENTION ¤ ¤ ¤ ¤")
Call TimeOut(0.15)
SendChat (thetext)
Call TimeOut(0.15)
SendChat ("¤ ¤ ¤ ¤ ATTENTION ¤ ¤ ¤ ¤")
Call TimeOut(0.15)
End Sub
Sub CenterForm(F As Form)
'centers the form
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Function FindChatRoom()
'finds the chat room
'mostly used for room busters
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function
Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function
Function FreeProcess()
'keeps your prog from freezing
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function GetchatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function
Sub SendMail(Recipiants, subject, message)
'sends mail to the recipiants
'ex: SendMail "katt737", "Hi Pal!", "<font =999999999999999999999999999999999999999999999999999999999999999999999999999999>"

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Sub HideAOL()
'minimizes aol then hides it from the taskbar
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub
Sub IMBuddy(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMIgnore(thelist As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Sub IMKeyword(Recipiant, message)
'sends an im to the recipiant with the message
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call KeyWord("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "KoA owns You")
End Sub
Sub IMsOn()
Call IMKeyword("$im_on", "KoA owns You")
End Sub


Function IsUserOnline()
'checks to see if the user of the program is
'on-line
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Sub KeyWord(TheKeyWord As String)
'calls up the keyword box
'same as hitting ctrl + k
'Ex: KeyWord ("vb")
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below
'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon
Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub
Sub KillWait()
'makes the hourglass on AOL turn back to the normal
'mouse icon (the arrow)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function LastChatLine()
'gets the last chatline of the chatroom
'Ex: If LastChatLine = "KoA" Then
'SendChat ("KoA owns All!")
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Function LastChatLineWithSN()
'gets the last chat line from the chat room with
'the sn
chattext$ = GetchatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function
Sub Playwav(File)
'lets you play a .wav in your prog
'Ex: PlayWav ("C:\America Online 4.0\drop.wav")
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (E)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub SendChat(Chat)
'sends text to the chat room
'ex: SendChat ("KoA owns me")
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub ShowAOL()
'maximizes aol and shows it in the taskbar
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Function SNfromIM()
'returns the screen name from im
'ex: If SNfromIM = "katt737" then
'IMKeyword ("katt737", "<font =999999999999999999999999999999999999999999999999999999999999>")
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function
Function SNFromLastChatLine()
'returns the sn from last chat line
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        sn = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = sn
End Function
Sub StayOnTop(TheForm As Form)
'makes your form stay on top of aol
'ex: StayOnTop Me
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub TimeOut(Duration)
'makes your prog pause for the duration of seconds you put
'ex: TimeOut (60) 'makes it pause for 1 minute
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function
Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function
Sub UnUpchat()
'stops upchat
'dont do this if you have not started upchat
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub


Sub Upchat()
'makes you be able to do other things while
'uploading
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Function UserSN()
'returns the user sn
'ex: SendChat ("My SN is: " & UserSN)
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function
Sub waitforok()
'used in room busters
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
