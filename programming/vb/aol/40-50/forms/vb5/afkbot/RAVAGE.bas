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
Function BlackGreenBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)

End Function
'Pre-set 2 color fade combinations begin here
Sub BoldFadeBlack(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & U$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & C$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & J$ & "<FONT COLOR=#666666>" & K$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
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
Sub BoldFadeGreen(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & U$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & b$ & "<FONT COLOR=#009900>" & C$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & J$ & "<FONT COLOR=#007700>" & K$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub
Sub BoldFadeRed(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & U$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & L$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & b$ & "<FONT COLOR=#220000>" & C$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & J$ & "<FONT COLOR=#440000>" & K$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendChat (PC$)


End Sub
Sub BoldFadeBlue(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & U$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & L$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & b$ & "<FONT COLOR=#0000D7>" & C$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & J$ & "<FONT COLOR=#0000A5>" & K$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeYellow(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & U$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & b$ & "<FONT COLOR=#222200>" & C$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & J$ & "<FONT COLOR=#444400>" & K$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub


Function BoldBlackBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)

End Function

Function BlackGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)

End Function

Function BoldBlackGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlackPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
SendChat (Msg)
End Function

Function BoldBlackYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BlueBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldBlueGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBluePurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBlueYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldRedBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldRedGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldYellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack2(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><U><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function
Function BoldBlackBlueBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function
Function BoldBlackGreenBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBluePurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldBluePurpleBlue = Msg
    BoldSendChat (BoldBluePurpleBlue)
End Function
Function BoldBlackGreyBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function Bolditalic_BlackPurpleBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRedBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlueBlackBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueGreenBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function Bolditalic_BluePurpleBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next b
 SendChat (Msg)
End Function

Function BoldBlueRedBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlueYellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlackGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreenBlueGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurpleGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreenRedGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function


Function BoldGreenYellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlackGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyGreenGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyPurpleGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyRedGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellowGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldPurpleBlackPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBluePurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreenPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleYellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function RedBlackRed2(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & H & ">" & D
    Next b
  SendChat (Msg)
End Function
Function BoldRedBlackRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function
Function BoldRedBlueRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldRedGreenRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldRedPurpleRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedYellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlackYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowBlueYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowGreenYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowPurpleYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   BoldSendChat (Msg)
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
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


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    C = 0
    O = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(Text, X, 1)
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
            Msg = Msg + Mid$(Text, X, 1)
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

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    D = Len(Text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C)
    GoTo TheEnd
Odds:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C + 1)
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

Function RGB2HEX(R, G, b)
    Dim X&
    Dim XX&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For XX& = 1 To 2
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
        Next XX&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
    TrimSpaces = Text
    Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpace, 1)
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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = NextChr$ & newsent$
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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let newsent$ = newsent$ + NextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let newsent$ = newsent$ + NextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
Loop
BoldBlackBlueBlack (newsent$)


End Function
Function R_Hacker2(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
Loop
BoldBlackBlueBlack2 (newsent$)


End Function
Function R_Spaced2(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
 RedBlackRed2 (newsent$)

End Function

Function R_Spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
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
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
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
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub KillWait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
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
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub timeout(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub AOLChatPunter(SN1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo errhandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

Heh$ = AOLLastChatLine
Heh$ = UCase(Heh$)
Naw$ = Mid(Heh$, InStr(Heh$, ":") + 2)
timeout (0.3)
SN = Mid(Naw$, InStr(Naw$, ":") + 1)
SN = UCase(SN)
timeout (0.3)
pntstr = Mid$(Naw$, 1, (InStr(Naw$, ":") - 1))
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
Public Sub Macrothing(txt As TextBox)
'This scrolls a multilined textbox adding timeouts where needed
'This is basically for macro shops and things like that.
BoldPurpleRed "· ···•(\›• INCOMMING TEXT"
timeout 4
Dim onelinetxt$, X$, Start%, I%
Start% = 1
fa = 1
For I% = Start% To Len(txt.Text)
X$ = Mid(txt.Text, I%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
BoldPurpleRed ": " + onelinetxt$
timeout (0.5)
J% = J% + 1
I% = InStr(Start%, txt.Text, X$)
If I% >= Len(txt.Text) Then Exit For
Start% = I% + 1
onelinetxt$ = ""
End If
Next I%
BoldSendChat ":" + onelinetxt$
End Sub
Sub Anti45MinTimer()
'use this sub in a timer set at 100
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Public Sub AOLChatManipulator(who$, what$)
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "" & (who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
 



Sub SendCharNum(win, chars)
E = SendMessageByNum(win, WM_CHAR, chars, 0)

End Sub


Sub AntiIdle4o()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
 
Sub AntiIdle()
'use this sub in a timer set at 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub ClickIcon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, subject, message)

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
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub KeyWord(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
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
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
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
a = Len(G$)
For W = 1 To a Step 4
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

Call timeout(0.01)
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

Call timeout(0.01)
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
Sub IMKeyword(Recipiant, message)

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

Call timeout(0.01)
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

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetchatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetchatText = ChatText
End Function

Function LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
ChatText$ = GetchatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = LastLine
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
On Error Resume Next
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function


Sub AddRoomToListbox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
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
Next index
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

Public Sub AOLEightLine(txt As TextBox)
'a simple 8 line scroller
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
AOLFindRoom
timeout 2
ChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
AOLFindRoom
ChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
AOLFindRoom
ChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
AOLFindRoom
ChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
 

End Sub

Public Sub AOLFourLine(txt As TextBox)
'a simple 8 line scroller
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""

SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""


End Sub


Public Sub AOLFifteenLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$
timeout 1.5
End Sub
Public Sub AOLFiveLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$
timeout 0.3
End Sub




Public Sub AOLSixTeenLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.7
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.7
End Sub

Public Function AOLSupRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo Last
FindChatRoom
If FindChatRoom = 0 Then GoTo Last

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
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call SendChat("HeY! " & Person$ & " WaZ uP?")
timeout (0.5)
Next index
Call CloseHandle(AOLProcessThread)
End If
Last:
End Function

Public Sub AOLTenLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
End Sub

Public Sub AOLThirtyFiveLine(txt As TextBox)
a = String(116, Chr(4))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$
timeout 0.3
End Sub

Public Sub AOLTwentyFiveLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + ""
timeout 1.5

End Sub


Public Sub AOLTwentyLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 1.5
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""
timeout 0.3
End Sub

Function Scrambletext(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Scrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
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
Function DescrambleText(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Descrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo City
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
City:
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

Sub Window_Hide(hwnd)
'This will hide the window of your choice
X = ShowWindow(hwnd, SW_HIDE)
End Sub



Sub Window_Show(hwnd)
'This will show the window of your choice
X = ShowWindow(hwnd, SW_SHOW)
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Sub PhreakyAttention(Text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & Text)
SendChat ("<I>" & Text)
SendChat ("<U>" & Text)
SendChat ("<S>" & Text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Sub Punter(Text)
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
Call IMKeyword(Text, pu)
Call IMKeyword(Text, Punt)

End Sub


Sub AOL4_Invite(Person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo errhandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
timeout (1#)
Chat% = FindChildByTitle(MDI%, "Buddy Chat")
aoledit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(Chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
errhandler:
Exit Sub
End Sub

Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Public Sub FourLine(txt As TextBox)
'‹¬•· ReeFeRs Re-Enter · ReeFeR
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
ChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""

SendChat "" + txt.Text + "" & C$ & "" + txt.Text + ""







End Sub
Sub AOL4_KillWin(windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = SendMessageByNum(windo, WM_CLOSE, 0, 0)
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
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function



Function BoldYellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    WhitePurpleWhite (Msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Green_LBlue (Msg)
End Function

Function BoldLBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Yellow_LBlue (Msg)
End Function

Function BoldPurple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 450 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldDGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange (Msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange_LBlue (Msg)
End Function

Function BoldLGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen (Msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen_LGreen (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
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
Sub falling_form(Frm As Form, steps As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
BgColor = Frm.BackColor
Frm.BackColor = RGB(0, 0, 0)
For X = 0 To Frm.Count - 1
Frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
Frm.Show
X = ((Screen.Width - Frm.Width) - Frm.Left) / steps
Y = ((Screen.Height - Frm.Height) - Frm.Top) / steps
Do
    Frm.Move Frm.Left + X, Frm.Top + Y
Loop Until (Frm.Left >= (Screen.Width - Frm.Width)) Or (Frm.Top >= (Screen.Height - Frm.Height))
Frm.Left = Screen.Width - Frm.Width
Frm.Top = Screen.Height - Frm.Height
Frm.BackColor = BgColor
For X = 0 To Frm.Count - 1
Frm.Controls(X).Visible = True
Next X
End Sub
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Sub AOLSetText(win, txt)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AOLAntiPunter()
'this is not the best anti there is use this
'at your own risk it is pretty buggy
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
Sts% = FindChildByClass(ANT%, "_AOL_Static")
St% = GetWindow(Sts%, GW_HWNDNEXT)
St% = GetWindow(St%, GW_HWNDNEXT)
Call AOLSetText(St%, "Ritual2x¹ - This IM Window Should Remain OPEN.")
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
Public Sub AOLKillWindow(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
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
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
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
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

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
If X = 0 Then GoTo Last
AOLActivate
SendKeys "{TAB}"
AG:
timeout (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
Last:
End Function
Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Sub AOLHostManipulator(what$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLGuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call KeyWord("PC")
MsgBox "A Guide had entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub
Function AOLCountMail()
TheMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(TheMail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_ListMail(box As ListBox)
box.Clear
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
If Counter = AOLCountMail Then GoTo Last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    box.AddItem buffer$
 timeout (0.001)
Counter = Counter + 1
GoTo Start
Last:
End Function

Function Mail_Out_CloseMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function Mail_Out_ListMail(box As ListBox)
box.Clear
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
If Counter = AOLCountMail Then GoTo Last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    box.AddItem buffer$
 timeout (0.001)
Counter = Counter + 1
GoTo Start
Last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
TheMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(TheMail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function FindOpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

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
GoTo Last
End If
counterf = -1

Start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo Last
If lst.Selected(counterf) = True Then GoTo Last
If couterf = lst.ListCount Then GoTo Last
GoTo Start

Last:
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

SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub AddRoomToCombobox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListbox(ListBox)
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
ChatSend ("<b>" & BoldChat & "</b>")
End Sub
Sub BoldWavyChatBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
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
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next W
BoldSendChat (P$)
End Function
Sub BoldWavyColorbluegree(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblack(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (P$)
End Sub
Sub BoldWavyColorredandblue(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
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







Sub Attention(TheText As String)

BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
BoldSendChat (TheText)
Call timeout(0.15)
BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call timeout(0.15)
'BoldSendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
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





Function TrimTime()
b$ = Left$(Time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function
Function TrimTime2()
b$ = Time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
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


Sub IMIgnore(TheList As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For FindSN = 0 To TheList.ListCount
        If LCase$(TheList.List(FindSN)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next FindSN
End If
End Sub
Function SNfromIM()

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
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Black_LBlue = Msg
End Function



Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowPinkYellow = Msg
End Function

Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    WhitePurpleWhite = Msg
End Function

Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Green_LBlue = Msg
End Function

Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Yellow_LBlue = Msg
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Purple_LBlue_Purple = Msg
End Function

Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 450 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    DBlue_Black_DBlue = Msg
End Function

Function DGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    DGreen_Black = Msg
End Function



Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange = Msg
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange_LBlue = Msg
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen = Msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen_LGreen = Msg
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_DBlue = Msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_DBlue_LBlue = Msg
End Function

Function PinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    PinkOrange = Msg
End Function

Function PinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    PinkOrangePink = Msg
End Function

Function PurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleWhite = Msg
End Function

Function PurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleWhitePurple = Msg
End Function
Function YellowBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function



Function YellowRedYellow(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function
Function YellowRed(Text As String)
    a = Len(Text)
    For b = 1 To a
        C = Left(Text, b)
        D = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
   SendChat (Msg)
End Function
Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Yellow_LBlue_Yellow = Msg
End Function
Sub BoldWavY(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
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
Call timeout(0.8)
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



Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
