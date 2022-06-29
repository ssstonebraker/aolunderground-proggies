Attribute VB_Name = "Hider"
'This is my first bas for AOL It is for
'AOL 40 it took me a while to make it
'So if you use it put me in the greetz
'of your prog this is the best bas out
'all the functions work any questions
'e-mail hider@hider.com or post on the board at my site
'at http://www.hider.com
'With the additions I added this is like a prog in a bas
'it has room bust,IManswer,MassIM,IdleBot,Attention and more
'Peace Hider
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
Declare Function ExitWindows Lib "user32" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellUse Lib "shell32.dll Alias (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long" ()
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppname As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

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

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Const EW_RESTARTWINDOWS = &H42
Public Const EW_REBOOTSYSTEM = &H43

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

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

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

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
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
Global giBeepBox As Integer
Global r&
Global entry$
Global iniPath$
'all you have to do to use this is in a button
'or a menu Call UpChat
Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
'all you do to use this is in a button
'or menu Call UnUpChat
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub
Sub KillWait()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")
For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 9)
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
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub
Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
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
Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Hider
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Hider
While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Hider
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Hider
Wend
FindChildByClass = 0
Hider:
Room% = firs%
FindChildByClass = Room%
End Function
Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo Hider
firs% = GetWindow(parentw, GW_CHILD)
While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo Hider
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo Hider
Wend
FindChildByTitle = 0
Hider:
Room% = firs%
FindChildByTitle = Room%
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Public Sub RunMenuByString(Application, StringSearch)
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
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub EightBall()
Dim Lst As String
Dim Text As String
Dim cht As Integer
Dim txt As String
Dim nws As String
Dim who As String
Dim wht As String
Dim r As Integer
Dim E As Integer
Dim M As Integer
Dim X
Dim Y
rider:
Y = SN()
AOL% = FindWindow("AOL Frame25", 0&)
cht = FindChildByClass(AOL%, "_AOL_View")
txt = WinCaption(cht)
If Lst = "" Then Lst = txt
If txt = Lst Then Exit Sub
Lst = txt
nws = LastChatLine(txt)
who = Mid(nws, 2, InStr(nws, ":") - 2)
wht = Mid(nws, Len(who) + 4, Len(nws) - Len(who))
If LCase(Trim(Trim(Y))) = LCase(Trim(Trim(who))) Then GoTo rider
r = GetParent(cht)
E = FindChildByClass(r, "_AOL_Edit")
tixt = RandomNumber(11)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then: tixt = "Definitely YES!"
ElseIf tixt = "3" Then: tixt = "Definitely No!"
ElseIf tixt = "4" Then: tixt = "Not a chance"
ElseIf tixt = "5" Then: tixt = "No way no"
ElseIf tixt = "6" Then: tixt = "Yesssss!"
ElseIf tixt = "7" Then: tixt = "Oh no try again."
ElseIf tixt = "8" Then: tixt = "Could Bee"
ElseIf tixt = "9" Then: tixt = "Yea baby"
ElseIf tixt = "10" Then: tixt = "I can't say fer sure"
ElseIf tixt = "11" Then: tixt = "Yes! Yes! Yes!"

End If
Text = wht$
w = InStr(LCase$(Text), LCase$("if"))
If w <> 0 Then
ChatRedYellowRed " & who & ", " The 8-ball say: " & tixt
TimeOut 0.5
GoTo rider
End If

End Sub
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
Chat (Made$)
End Sub
Function FadeForm1(ByVal frmIn As Form, iColor As Integer)
       Dim i As Integer
       Dim Y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 13
       .DrawWidth = 2
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With

For i = 0 To 255
       'To use this in the form load
       'FadeForm1 Formname,1 or
       'whatever # case you want to use
       Select Case iColor
       Case 1 'Black to Red
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, 0, 0), BF
       Case 2 'Black to Green
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, i, 0), BF
       Case 3 'Black to Blue
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, 0, i), BF
       Case 4 'Black To White
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, i, i), BF
       Case 5 'Black To Yellow
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, i, 0), BF
       Case 6 'Black To Agua
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, i, i), BF
       Case 7 'Black To Fusia
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, 0, i), BF
       Case 8 'Maroon To Blue
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(128, 0, i), BF
       Case 9 'Lime To Orange
       frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, 128, 0), BF
End Select

Y = Y + 2
Next i
End Function
Function FadeForm2(ByVal frmIn As Form, iColor As Integer)

       Dim i As Integer
       Dim Y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 4
       .DrawWidth = 2
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With
        'You call this FadeForm2 Formname,1
        'any case# you want to use
For i = 0 To 255
        Select Case iColor
        Case 1 'White To Agua
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, 0, 0), BF
        Case 2 'White To Purple
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, i, 0), BF
        Case 3 'White To Yellow
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, 0, i), BF
        Case 4 'White To Black
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, i, i), BF
        Case 5 'White To Blue
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(i, i, 0), BF
        Case 6 'White To Red
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(0, i, i), BF
        Case 7 'Agua to green
        frmIn.Line (0, Y)-(frmIn.Width, Y + 2), RGB(128, 0, i), BF
End Select

Y = Y + 2
Next i
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Sub Chat(Chat)
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
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub
Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
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
Function SN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
SN = User
End Function
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
Sub RespondIM(message)
'you need 2 listboxes 1 for th SNFromIM
'the other for MessageFromIm
'and a textbox for your message
'Call RespondIm Text1
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo ghost
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo ghost
Exit Sub
ghost:
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
e2 = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(e2, GW_HWNDNEXT)
Call SetText(e2, message)
ClickIcon (E)
Pause 0.8
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
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
End Sub
Function FindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
ghost% = FindChildByClass(Room%, "_AOL_Listbox")
rider% = FindChildByClass(Room%, "RICHCNTL")
If ghost% <> 0 And rider% <> 0 Then
   FindRoom = Room%
Else:
   FindRoom = 0
End If
End Function
Sub SetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear
Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
If person$ = SN Then GoTo Bs
ListBox.AddItem person$
Bs:
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
End Sub
Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Public Sub ClearChat()
'This will clear the chat screen
'I'm sure you seen progs do it now you
'know how it's done
getpar% = FindRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
End Sub
Function GetchatText()
Room% = FindRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function
Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub
Function LinkSender(txt As String, URL As String)
Hyperlink = ("<A HREF=" & Chr$(34) & Text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Function SNfromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo rider
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo rider
Exit Function
rider:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$
End Function
Function SNFromLastLine()
chattext$ = LastLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastLine = SN
End Function
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Function SupRoom()
'call this in a command button
Online
If Online = 0 Then GoTo ghost
FindRoom
If FindRoom = 0 Then GoTo ghost
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    
Room = FindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call Chat("Sup   " & person$)
TimeOut (1)
Next Index
Call CloseHandle(AOLProcessThread)
End If
ghost:
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
Sub IMSend(Recipiant, message)
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
Call IMSend("$IM_OFF", "HEHE ")
End Sub
Sub IMsOn()
Call IMSend("$IM_ON", "HOHO ")
End Sub
Function Online()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   Online = 1
Else:
   Online = 0
End If
End Function
Function LastLine()
'Gets the last line of chat
chattext = LastLineWithSN
ChatTrimNum = Len(SNFromLastLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastLine))
LastLine = ChatTrim$
End Function
Function LastLineWithSN()
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
LastLine = Mid(chattext$, lastlen, Len(thechars$))
LastLineWithSN = LastLine
End Function
Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo rider
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo rider
Exit Function
rider:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function
Sub MacroScroll(Text As String)
'put your macro in a text box then
'Call MacroScroll
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    Chat Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub
Sub Mail(Recipiants, subject, message)
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
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Sub MailHider()
    Mail t3terminat, "This bas", "This is the best bas for AOL 4.0"
End Sub
Sub IMHider()
    IMSend t3terminat, "sup"
End Sub
Sub Attention()
'just replace everything in between the " "
'with what you want to send also you call this in a timer
'set the interval to 60000
    Chat ("{S IM Attention:Hider.bas is the best")
    TimeOut (0.15)
    Chat ("<B>Get Your Copy At http://www.hider.com")
    TimeOut (0.15)
    Chat ("{S IM Attention")
End Sub
Public Sub LabelScroll()
'This will make a label scroll across
'a form below is the code you need
'to make it work
Dim Z As Integer
Dim q As Integer
Dim F As Integer
Dim CurrentColor As Integer
If CurrentColor = r Then
Z = Z + 5
End If
If CurrentColor = G Then
q = q + 5
End If
If CurrentColor = b Then
F = F + 5
End If
End Sub

'This is the code for the label scroller
'Put this in the Form_Load procedure
'TrueFalse = True
'Randomize Timer
'z = 0
'q = 0
'f = 0
'CurrentColor = R

'Put this in a timer
'If z = 255 And q = 255 And f = 255 Then
'z = 0
'q = 0
'f = 0
'End If
'If Label1.Left > Form1.Width Then Label1.Left = Form1.Left - Label1.Width
'Label1.Left = Label1.Left + 50
'ModifyColor
'If z = 255 Then
'CurrentColor = g
'UpDown = down
'End If
'If q = 255 Then
'CurrentColor = b
'b = 0
'UpDown = down
'End If
'If f = 255 Then
'CurrentColor = R
'R = 0
'UpDown = down
'End If
'Label1.ForeColor = RGB(z, q, f)
'End Sub

Public Sub Buster()
'Put 2 labels on your form and
'and a textbox for the roomname
'Call Buster in a command Button
'Put a stop button also in the stop
'button put Label2.Caption = "Stop Bust"
    Label1.Caption = "Busting"
    Label2.Caption = "0"
    Do
    If Label2.Caption = "Stop Bust" Then
    Exit Sub
    End If
    Call KeyWord("aol://2719:2-2-" & Text1.Text)
    Label2.Caption = Val(Label2.Caption) + 1
    waitforok
    Label2.Caption = Val(Label2.Caption) + 1
    Call FindChatRoom
    Loop Until FindChatRoom = Text1.Text
    Exit Sub
End Sub
Sub Advertize()
'I'd appreciate it if you used this
'sub to addvertize my bas
    ChatRedYellowRed ("This prog made with Hider.bas")
    ChatRedYellowRed ("Get your copy at")
    ChatRedYellowRed ("http://www.hider.com")
End Sub
Function ChatBlackBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 255 / A
        F = E * b
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatBlackGreenBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatBlackYellowBlack(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatGreenBlueGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatGreenYellowGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatRedYellowRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatRedGreenRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Function ChatRedPurpleRed(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 510 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    SendChat (Msg)
End Function
Function ChatGreyYellowGrey(Text1)
    A = Len(Text1)
    For b = 1 To A
        C = Left(Text1, b)
        D = Right(C, 1)
        E = 490 / A
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next b
    Chat (Msg)
End Function
Public Sub MassIM()
'To make this work put a comdialog control
'on your form and 7 command buttons
'a listbox and textbox
'in command 1 Call MassIm
'then copy the code below to the command buttons
    For i = 0 To List1.ListCount - 1
    IMSend (List1.List(i)), Text1.Text + Chr(13) + Chr(10) + Chr(13) + ("  <B><U>Retro*MassIM</B></U>  ")
    Next i
End Sub
'This is the MassIM code
'Private Sub Command2_Click()
    'List1.AddItem Text2
    'Text2 = "Add Name"
'End Sub

'Private Sub Command3_Click()
    'AddRoomToListBox List1
'End Sub

'Private Sub Command4_Click()
    'On Error GoTo openerror1
    'CommonDialog1.CancelError = True
    'CommonDialog1.Filter = "List Files(*.lst)1*.lst"
    'CommonDialog1.FilterIndex = 1
    'CommonDialog1.DialogTitle = "MassIM"
    'CommonDialog1.Action = 1
    'fname = CommonDialog1.filename
    'Open fname For Output As #1
    'For I = 0 To List1.ListCount - 1
    'Print #1, List1.List(I)
    'Next I
    'Close #1
    'Exit Sub
'openerror1:
    'On Error GoTo 0
    'Exit Sub
'End Sub

'Private Sub Command5_Click()
    'CommonDialog1.CancelError = True
    'CommonDialog1.Filter = "List Files(*.lst)1*.lst"
    'CommonDialog1.FilterIndex = 1
    'CommonDialog1.DialogTitle = "Text Files"
    'CommonDialog1.Action = 1
    'fname = CommonDialog1.filename
    'Open fname For Input As #1
    'Do While Not EOF(1)
    'Input #1, filedata
    'For l0072 = 0 To List1.ListCount - 1
    'DoEvents
    'l007C = List1.List(l0072)
    'l0080 = InStr(1, l007C, filedata, 1)
    'If l0080 Then
    'l0084 = Len(l007C)
    'l0088 = Len(filedata)
    'If l0084 = l0088 Then
    'GoTo 900
    'End If
    'End If
    'Next l0072
    'List1.AddItem filedata
'900:
   ' Loop
    'Close #1
'openerror:
    'On Error GoTo 0
    'Exit Sub
'End Sub

'Private Sub Command6_Click()
    'If List1.ListIndex < 0 Then Exit Sub
    'List1.RemoveItem List1.ListIndex
'End Sub

'Private Sub Command7_Click()
    'List1.Clear
    'Dim oldstr As String
    'Dim rmcnt As Variant
    'Dim mmcnt As Variant
    'Dim isim As Variant
    'Dim stoptext%
'End Sub
 Public Sub IMAnswer()
'In command1 Call IMAnswer
'To make this work put 1 text box,
'3 labels,a listbox and a stop button
'label1.caption = "What to Say"
'Put it above Text1
'label2.caption = " Messages"
'put this above the listbox
'label3 has no caption
'in command2 put Exit Sub
1:
IM% = FindChildByTitle(MDI(), ">Instant Message From:")
If IM% Then GoTo Z
IM% = FindChildByTitle(MDI(), "  Instant Message From:")
If IM% Then GoTo Z
Pause (1)
GoTo 1
Z:
Call MessageFromIM
Let X = MessageFromIM
Call SNfromIM
Let Y = SNfromIM
List1.AddItem Y & " : " & X
killwin (IM%)
IMSend Y, Text1
Pause (1)
GoTo 1
End Sub
Sub VisitHidersSite()
    Call KeyWord("http://www.hider.com")
End Sub
Sub killwin(Wind)
    X = SendMessageByNum(Wind, WM_CLOSE, 0, 0)
End Sub
 Public Sub IdleBot()
'In command button:Call IdleBot
'Put a textbox on your form for the reason your away
'you can also call this in a timer with interval at 60000
'put a stop button or it will loop forever
'if you call it in a timer comment out the
'Timeout (30)
Idle:
    Chat ("Retro*Active Idle Bot")
    TimeOut (0.15)
    Chat ("Reason: ") & Text1.Text
    TimeOut (30) 'This will make it send after 30 secs
    Chat ("Retro*Active Idle Bot")
    TimeOut (0.15)
    Chat ("Reason: ") & Text1.Text
    GoTo Idle
    
End Sub
Public Sub ErrorPunt1()
'Put 2 text boxes one for who text2
'one for howmany text1
    Dim A
    Dim howmany$
    Fuckee$ = Text2.Text
    howmany$ = Text1.Text
    For A = 1 To howmany$
    IMSend Fuckee$, ("<S><BR><B><I><U><h3></S>Retro*Active<FONT 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
    Text1.Text = Text1.Text - 1
    Next A
    Text1.Text = Text1.Text - 1
    If howmany$ = -1 Then
    ChatRedGreenRed Text2.Text & (" Is Gone ")
    TimeOut (0.25)
    ChatRedGreenRed ("Retro*Active v1.0 By GhostRider")
    'GhostRider was my old handle
    End If
End Sub
Public Sub ErrorPunt2()
   Dim A
    Dim howmany$
    Fuckee$ = Text2.Text
    howmany$ = Text1.Text
    For A = 1 To howmany$
    IMSend Fuckee$, ("<S><BR><B><I><U><h3></S>Retro*Active<FONT @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@>")
    Text1.Text = Text1.Text - 1
    Next A
    Text1.Text = Text1.Text - 1
    If howmany$ = -1 Then
    ChatRedGreenRed Text2.Text & (" Is Gone ")
    TimeOut (0.25)
    ChatRedGreenRed ("Retro*Active v1.0 By GhostRider")
    'Ghostrider is my old handle
    End If
End Sub
Function IMChecker()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "Person is online and IMs are OFF"
End If
If aolcl% = 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "Person is online and IMs are ON"
End If
End Function
Public Sub OnlineChecker(person)
Call IMSend(person, "Hello")
Pause 2
IMChecker
End Sub
Sub OpenEXE(FileName$)
OpenEXE = Shell(FileName$, 1): NoFreeze% = DoEvents()
End Sub
Public Sub LoadAol()
Dim X%
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
 Sub FormMove(Form As Form)
 'Call this in the Mouse_Move event
 'of a label or any other control
 'Call FormMove FormName
       Dim Ret&
       ReleaseCapture
       Ret& = SendMessage(Form.hwnd, &H112, &HF012, 0)
End Sub
Function RGBtoHEX(RGB)
    A = Hex(RGB)
    b = Len(A)
    If b = 5 Then A = "0" & A
    If b = 4 Then A = "00" & A
    If b = 3 Then A = "000" & A
    If b = 2 Then A = "0000" & A
    If b = 1 Then A = "00000" & A
    RGBtoHEX = A
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
Sub RoomKilla()
'call this in a timer with the interval
'set to 100 and put a start and stop button
'start button = Timer1.Enabled = True
'stop = Timer1.Enabled = False
'don't call this in a start button
'you won't be able to stop it
killa:
    Chat ("@@@@@@@@@@@@@@@@@@@@@@@@@@HEHE@@@@@@@@@@@@@@@@@@@@")
    Chat ("@@@@@@@@@@@@@@@@@@@@@@@@@@HEHE@@@@@@@@@@@@@@@@@@@@")
    Chat ("@@@@@@@@@@@@@@@@@@@@@@@@@@HEHE@@@@@@@@@@@@@@@@@@@@")
    Chat ("@@@@@@@@@@@@@@@@@@@@@@@@@@HEHE@@@@@@@@@@@@@@@@@@@@")
    TimeOut (1)
    GoTo killa
End Sub
Sub SpiralScroll(txt As TextBox)
X = txt.Text
rider:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
Chat "•[" + X + "]•"
If txt.Text = X Then
Exit Sub
End If
GoTo rider
End Sub
