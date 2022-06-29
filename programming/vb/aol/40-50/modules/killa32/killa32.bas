Attribute VB_Name = "Killa32"
'Sup?
'This is the first AOL4.0 bas made by me
'E-Mail me things to add at   Ears98@hotmail.com
'Anything that has the names of  colors
'Ex.- YellowredYellow is a fading color
'this took me a long time to make to send it to a
'chat rooom do:  Sendtext ""& YellowredYellow ("l")
'And things like that
'
'
'Killa      '98'
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


Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
SendChat (P$)
End Sub


Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRB = P$
End Function
'Sub MyASCII(PPP$)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'LO$ = WavYChaT(PPP$ & "Loaded")
'B$ = WavYChaT("User: " & UserSN)
'TI$ = CoLoRChaT(TrimTime)
'V$ = CoLoRChaT("v¹·¹")
'FONTTT$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & V$ & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & LO$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •")
'Call timeout(0.15)
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & B$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •" & TI$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
'End Sub

Function WavyChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Function
'Sub MyASCII(PPP$)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'LO$ = WavYChaT(PPP$ & "Loaded")
'B$ = WavYChaT("User: " & UserSN)
'TI$ = CoLoRChaT(TrimTime)
'V$ = CoLoRChaT("v¹·¹")
'FONTTT$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & V$ & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & LO$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •")
'Call timeout(0.15)
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & B$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •" & TI$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
'End Sub

Function WavRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Function


Sub RedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Sub

Function WavY(thetext As String)

G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & S$ & "</sub>" & T$
Next w
WavY = P$

End Function


Function FadeBlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreenBlack = Msg
End Function


Function FadeBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreen = Msg
End Function


Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGrey = Msg
End Function
Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackGreyBlack = Msg
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackPurple = Msg
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackPurpleBlack = Msg
End Function


Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackRed = Msg
End Function

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackRedBlack = Msg
End Function


Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackYellow = Msg
End Function


Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackYellowBlack = Msg
End Function

Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueBlack = Msg
End Function
Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueBlackBlue = Msg
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueGreen = Msg
End Function


Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueGreenBlue = Msg
End Function


Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BluePurple = Msg
End Function


Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BluePurpleBlue = Msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueRed = Msg
End Function


Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueRedBlue = Msg
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueYellow = Msg
End Function


Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlueYellowBlue = Msg
End Function


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



Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenBlack = Msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenBlackGreen = Msg
End Function


Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenBlue = Msg
End Function


Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenBlueGreen = Msg
End Function


Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenPurple = Msg
End Function


Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenPurpleGreen = Msg
End Function


Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenRed = Msg
End Function


Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenRedGreen = Msg
End Function

Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenYellow = Msg
End Function

Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreenYellowGreen = Msg
End Function


Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlack = Msg
End Function


Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlackGrey = Msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlue = Msg
End Function


Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyBlueGrey = Msg
End Function


Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyGreen = Msg
End Function


Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyGreenGrey = Msg
End Function


Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyPurple = Msg
End Function


Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyPurpleGrey = Msg
End Function


Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyRed = Msg
End Function


Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyRedGrey = Msg
End Function


Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyYellow = Msg
End Function


Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    GreyYellowGrey = Msg
End Function


Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBlack = Msg
End Function


Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBlackPurple = Msg
End Function


Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBlue = Msg
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleBluePurple = Msg
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleGreen = Msg
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleGreenPurple = Msg
End Function
Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleRed = Msg
End Function


Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleRedPurple = Msg
End Function


Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleYellow = Msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleYellowPurple = Msg
End Function


Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlack = Msg
End Function


Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlackRed = Msg
End Function
Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlue = Msg
End Function


Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedBlueRed = Msg
End Function
Function fadeRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedGreen = Msg
End Function


Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedGreenRed = Msg
End Function
Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedPurple = Msg
End Function


Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedPurpleRed = Msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedYellow = Msg
End Function


Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    RedYellowRed = Msg
End Function


Function RGB2HEX(r, G, B)
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
        If X& = 3 Then Color& = r
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
Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function



Function ThreeColors(text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    d = Len(text)
        If d = 0 Then GoTo TheEnd
        If d = 1 Then Fade1 = text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    c = d \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c)
    GoTo TheEnd
Odds:
    c = d \ 2
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
    ThreeColors = Msg
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


Function TwoColors(text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = C1
    C2BAK = C2
    C3BAK = C3
    C4BAK = C4
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
        
        C1 = RGB2HEX(VAL1, VAL2, VAL3)
        C2 = RGB2HEX(VAL1, VAL2, VAL3)
        C3 = RGB2HEX(VAL1, VAL2, VAL3)
        C4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + C1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
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
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
            End If
        ElseIf WavY = False Then
            Msg = Msg + Mid$(text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
        End If
        End If
nc:     Next X
    C1 = C1BAK
    C2 = C2BAK
    C3 = C3BAK
    C4 = C4BAK
    TwoColors = Msg
End Function
Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlack = Msg
End Function


Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlackYellow = Msg
End Function

Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlue = Msg
End Function


Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowBlueYellow = Msg
End Function


Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowGreen = Msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowGreenYellow = Msg
End Function
Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowPurple = Msg
End Function


Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowPurpleYellow = Msg
End Function
Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRed = Msg
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    YellowRedYellow = Msg
End Function
Function FadeBlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackBlueBlack = Msg
End Function


Function FadeBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    BlackBlue = Msg
    Sendtext "" & Msg
End Function


Sub Oneim(Recipiant)
'This is the only AOL4 punt string i know it gives
'them an error messige and closes AOL
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
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999><Font = 99999999999999999999999999999999999999999999999999999999999>")

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

Sub Sendtext(Chat)
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

Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
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
Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
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
SendChat (Made$)
End Sub
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
Function FindChatRoom()
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
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
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
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function GetchatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function
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
Sub IMsOff()
Call IMKeyword("$IM_OFF", "Pussy")
End Sub
Sub IMsOn()
Call IMKeyword("$IM_ON", "TRU")
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

hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Sub KeyWord(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
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

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub KillGlyph()
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
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Function LastChatLineWithSN()
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
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub


Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub RespondIM(message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

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
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (e)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(IM%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
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

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
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
Function SNFromLastChatLine()
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
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub
Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
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
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

