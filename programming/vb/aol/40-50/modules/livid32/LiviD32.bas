Attribute VB_Name = "LiviD32"
'-= LiviD32.bas by LiviD =-

'email me: xLiviDx@juno.com
'scroll down for autophade, random,
'form moving and bot help

'Sup everyone..this is LiviD and
'this is the first module ive
'ever given out.
'this bas has all the basics, and
'even a few extra things i added
'that most bas' dont have...
'simple 3 and 4 color rgb fade
'a bunch of mail things
'AIM IMer
'how to autophade and make bots (scroll down)
'how to get random sayings(scroll down)
'****************************
'Directions for this bas:
'to use a Sub:
'a) if it has variables to use,
'   then u must define them
'   ex: Call IMSend("LiviD", "Hi")
'b)if it has no variables to define
'  then just call the sub
'  ex: Call AOLKillWait
'to use a function
'ex: LiviD% = FindChildByClass(im%, "RICHCNTL")

'********TUTORIALS*******

'Autophade (this is if u use spyworks,
'i never used anything else..)
'email xLiviDx@juno.com and ill send u
'spyworks 5.1
'
'put in a command button:
'subclass1.addhwnd = chatsendbox
'
'
'then in SubClass1_WndMessageX put:
'If wp = 13 Then
'    Richcntl% = chatsendbox
'    thetext$ = GetText(Richcntl%)
'    If thetext$ = "" Then Exit Sub
'    If InStr(LCase(thetext$), "<font color=") <> 0 Then Exit Sub
'    SubClass1.RemoveHwnd = Richcntl%
'    the2$ = ChatFade(thetext$,250,0,0,250,0,0,true)
'    Call SendMessageByString(Richcntl%, WM_SETTEXT, 0, "")
'    Call SendMessageByString(Richcntl%, WM_SETTEXT, 0, the2$)
'    Richcntl% = ChatSendBox
'    SubClass1.AddHwnd = Richcntl%
'End If
'
'
'Bots:
'ok bots are easy..this bot will say
'hello (the screen name) when u say hi
'put this in a timer with intervals of 1:
'on error goto DiviL
'lastline$ = chatlastline
'whoend = instr(lastline$, ":")
'who$ = left$(lastline$, whoend - 1)
'what$ = mid$(lastline$, whoend + 3)
'if lcase(what$) = "hi" then
'call chatsend("Hello " + who$)
'end if
'DiviL:
'
'randomizing..i have a sub in this bas
'that will give u a random number
'put in a button:
's1$ = "a"
's2$ = "b"
's3$ = "c"
's4$ = "d"
's5$ = "e"
'x = RandomNumber(5)
'select case(x)
'case 1:
'   s$ = s1$
'case 2:
'   s$ = s2$
'case 3:
'   s$ = s3$
'case 4:
'   s$ = s4$
'case 5:
'   s$ = s5$
'end select
'msgbox s$
'
'
'Moving a form without a title bar
'put in the mousedown procedure of
'any object like a label:
'MvForm Me
'
'
'
'thats it!..its that easy..
'ok now that i showed ur fat lazy
'lame ass how to autophade, randomize,
'move forms and make bots...
'now, whats the best aol 4 bas?..
'u damn right its this one!






'You shouldnt touch this shit

'Funcs and Subs:
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function getnextwindow Lib "user32" Alias "GetNextWindow" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpappname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
'Pretty Boring huh?
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Constants
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_gettext = &HD
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MDIDESTROY = &H221
Public Const LB_GETtext = &H189
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCURSEL = &H188
Public Const LB_INSERTSTRING = &H181
Public Const VK_END = &H23
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
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const FLAGz = SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
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
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_APPEND = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const CB_GETCOUNT = (WM_USER + 6)
Public Const CB_GETITEMDATA = (WM_USER + 16)
Public Const CB_GETLBTEXTLEN = (WM_USER + 9)
Public Const CB_INSERTSTRING = (WM_USER + 10)
Public Const CB_SETCURSEL = (WM_USER + 14)
Public Const CB_SETEDITSEL = (WM_USER + 2)
Public Const CB_SHOWDROPDOWN = (WM_USER + 15)
Public Const EM_GETLINE = WM_USER + 20
Public Const EM_GETLINECOUNT = WM_USER + 10
Public Const EM_GETSEL = WM_USER + 0
Public Const EM_REPLACESEL = WM_USER + 18
Public Const EM_SCROLL = WM_USER + 5
Public Const EM_SETFONT = WM_USER + 19
Public Const EM_SETREADONLY = (WM_USER + 31)
Public Const EW_REBOOTSYSTEM = &H43
Public Const KEY_DELETE = &H2E
Public Const LB_32GETCOUNT = &H18B
Public Const LB_32GETCURSEL = &H188
Public Const LB_32GETITEMDATA = &H199
Public Const LB_32GETTEXT = &H189
Public Const LB_32GETTEXTLEN = &H18A
Public Const LB_32SETCURSEL = &H186
Public Const LB_GETITEMRECT = (WM_USER + 25)
Public Const LBN_DBLCLK = 2
Public Const MB_TASKMODAL = &H2000
Public Const MF_BITMAP = &H4
Public Const SRCCOPY = &HCC0020
Public Const SW_NORMAL = 1
Public Const SW_SHOWNA = 8
Public Const WM_COPY = &H301
Public Const WM_GETFONT = &H31
Public Const WM_MOVE = &H3
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFONT = &H30
Public Const WS_BORDER = &H800000
Public Const WS_THICKFRAME = &H40000
Dim OldLast$
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

Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function
Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
AOLMDI = MDI%
End Function
Sub AOLRunMenuByString(StringSearch)
'this will only run the menus named "File" "Edit"
'etc..at this time, people would get pissed at
'me for showing how to run the pop up menus...
'but maybe ill show it in v2
Application = AOLWindow
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

Sub AOLChangeCaption(Change$)
Call SetText(AOLWindow, Change$)
End Sub
Sub AOLChangeSN(SN As String, aoldir As String, Replace As String)
'changes the screen names on your comp
' - LiviD
If Right$(aoldir, 1) = "\" Then aoldir = Left$(aoldir, Len(aoldir) - 1)
LiFile = FreeFile
Open aoldir$ + "\idb\main.idx" For Binary As #LiFile
If Len(SN$) > Len(Replace$) Then
    Replace$ = Replace$ + String(Len(SN$) - Len(Replace$), " ")
End If
For X = 1 To LOF(LiFile) Step 32000
    text$ = Space(32000)
    Get #LiFile, X, text$
Find:
    If InStr(1, text$, SN$, 1) Then
        Where = InStr(1, text$, SN$, 1)
        Put #LiFile, (X + Where) - 1, Replace$
        Mid$(text$, Where, 15) = String(15, " ")
        GoTo Find
    End If
    FreeProcess
Next X
Close #LiFile
End Sub
Sub SetText(Window, text)
ext% = SendMessageByString(Window, WM_SETTEXT, 0, "")
ext% = SendMessageByString(Window, WM_SETTEXT, 0, text)
End Sub
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function AOLUser()
On Error Resume Next
Welcome% = Findchildbytitle(AOLMDI, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUser = User
End Function

Sub AOLKeyword(TheKeyWord)
'by LiviD =]
'EX: AOLKeyword "aol://2719:2-2-LiviD"
'uses the toolbar combobox
aol% = FindWindow("AOL Frame25", vbNullString)
AOLTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOLTooL% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOLCombo% = FindChildByClass(AOLTooL%, "_AOL_ComboBox")
aoledit% = FindChildByClass(AOLCombo%, "Edit")
Call SendMessageByString(aoledit%, WM_SETTEXT, 0, TheKeyWord)
Call SendMessageByNum(aoledit%, WM_CHAR, VK_SPACE, 0)
Call SendMessageByNum(aoledit%, WM_CHAR, VK_RETURN, 0)
End Sub
Function Online() As Boolean
Welcome% = Findchildbytitle(AOLMDI, "Welcome, " + AOLUser + "!")
If Welcome% = 0 Then Online = False: Exit Function
Online = True
End Function
Function SearchFile(FileName As String, SearchString As String) As Long
Free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
    text$ = Space(32000)
    Get #Free, X, text$
    Debug.Print X
    If InStr(1, text$, SearchString$, 1) Then
        Where = InStr(1, text$, SearchString$, 1)
        SearchFile = (Where + X) - 1
        Close #Free
        Exit For
    End If
    Next X
Close #Free
End Function
Function FindChildByClass(parentw, childhand)

firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
Wend
FindChildByClass = 0

Flapjacks:
room% = firs%
FindChildByClass = room%

End Function
Function Findchildbytitle(parentw, childhand)

firs% = GetWindow(parentw, 5)
If UCase(GetText(firs%)) Like UCase(childhand) Then GoTo socks
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetText(firss%)) Like UCase(childhand) & "*" Then GoTo socks
firs% = GetWindow(firs%, 2)
If UCase(GetText(firs%)) Like UCase(childhand) & "*" Then GoTo socks
Wend
Findchildbytitle = 0

socks:
room% = firs%
Findchildbytitle = room%
End Function


Sub StayOnTop(the As Form)
Call SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Sub NotOnTop(the As Form)
Call SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Sub pause(amountoftime)
starttime = Timer
Do While Timer - starttime < amountoftime
FreeProcess
Loop
End Sub

Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
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
ChatSend (Made$)
End Sub

Function ChatFindRoom()
    aol% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(aol%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (l <> 100)
            FreeProcess
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        ChatFindRoom = firs%
        Exit Function
    End If
    ChatFindRoom = 0
End Function

Function ChatText() As String
rich% = FindChildByClass(ChatFindRoom, "RICHCNTL")
ChatText = GetText(rich%)
End Function
Sub ChatClearText()
rich% = FindChildByClass(ChatFindRoom, "RICHCNTL")
Call SetText(rich%, "")
ChatSend AOLUser & " Chat Cleared!"
End Sub
Function ChatLastLine()
ChatTexts$ = ChatText
For FindChar = 1 To Len(ChatTexts$)
thechar$ = Mid(ChatTexts$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(ChatTexts$, lastlen, Len(thechars$))
ChatLastLine = LastLine
End Function
Sub Click(icon%)
SendMessage icon%, WM_LBUTTONDOWN, 0, 0&
pause 0.0000001
SendMessage icon%, WM_LBUTTONUP, 0, 0&
End Sub
Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long
Dim i As Integer
If hWnd = 0 Then
GoTo Return_False
End If
hChild = GetWindow(hWnd, GW_CHILD)
While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend
GetChildCount = i
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Function GetLineCount(text)
theview$ = text
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
If thechar$ = Chr(13) Then
numline = numline + 1
End If
Next FindChar
If Mid(text, Len(text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function
Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer
Dim iIndex As Integer
With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With
GetListIndex = -2
End Function
Public Function GetComboIndex(oListBox As ComboBox, sText As String) As Integer
Dim iIndex As Integer
With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetComboIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With
GetComboIndex = -2
End Function
Sub IMSend(Person, Message)
Call AOLRunPopUpMenu3(10, 7177)
Do: FreeProcess
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
im% = Findchildbytitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSen% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSen% <> 0 Then Exit Do
Loop
Call SetText(aoledit%, Person)
Call SetText(aolrich%, Message)
IMSen% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMSen% = GetWindow(IMSen%, 2)
Next sends
Do: FreeProcess
Click (IMSen%)
pause 0.000000001
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
im% = Findchildbytitle(MDI%, "Send Instant Message")
msg% = FindWindow("#32770", "America Online")
If msg% <> 0 Then
text% = FindChildByClass(msg%, "Button")
closer = SendMessage(msg%, WM_CLOSE, 0, 0)
closer2 = SendMessage(im%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop Until im = 0 And msg = 0
who$ = Person
If TrimSpaces(LCase(who$)) = "$im_off" Or TrimSpaces(LCase(who$)) = "$im_on" Then Exit Sub
Do
pause 0.001
im% = Findchildbytitle(AOLMDI, "  Instant Message To: ")
Loop Until im% <> 0
pause 0.001
Do
im% = Findchildbytitle(AOLMDI, "  Instant Message To: ")
pause 0.0000000000001
WinKill im%
Loop Until im% = 0
End Sub

Sub MvFrm(frm As Form)
Dim Ret&
ReleaseCapture
Ret& = SendMessage(frm.hWnd, &H112, &HF012, 0)
End Sub


Sub Upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub ChatSend(text)
'this has a pause at the bottom, so u cant
'scroll off with the new tos thingy
If ChatFindRoom = 0 Then Exit Sub
R7% = ChatSendBox
FreeProcess
sBuffer = GetText(R7%)
Call SetText(R7%, "")
Call SetText(R7%, text)
Do
Call SendCharNum(R7%, 13)
pause 0.2
Loop Until GetText(ChatSendBox) <> text
Call SetText(R7%, sBuffer)
pause 0.6
End Sub
Sub PlayWav(File)
'This will play a WAV file
'*Doesnt wait for the wav to end*
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub PlayWavWait(File)
'This will play a WAV file
'*Waits for the wav to end*
SoundName$ = File

X = sndPlaySound(SoundName$, 0)
End Sub
Sub IMsOff()
Call IMSend("$IM_OFF", "¿LiviD?")
End Sub
Sub IMsOn()
Call IMSend("$IM_ON", "¿LiviD?")
End Sub

Sub ForAddRoomList(itm As String, Lst As ListBox)
'shouldnt fuck with this
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until XX = (Lst.ListCount)
Let diss_itm$ = Lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Sub ForAddRoomCombo(itm As String, Lst As ComboBox)
'shouldnt fuck with this
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until XX = (Lst.ListCount)
Let diss_itm$ = Lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub

Sub ChatAddRoomList(Lst As ListBox)

Dim Index As Long
Dim i As Integer
For Index = 0 To 25
    names$ = String$(256, " ")
    Ret = ForAddRoom(Index, names$)
    names$ = Left$(Trim$(names$), Len(Trim(names$)))
    ForAddRoomList names$, Lst
Next Index
endaddroom:
Lst.RemoveItem Lst.ListCount - 1
i = GetListIndex(Lst, AOLUser())
If i <> -2 Then Lst.RemoveItem i
End Sub
Sub ChatAddRoomCombo(Lst As ComboBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
    names$ = String$(256, " ")
    Ret = ForAddRoom(Index, names$)
    names$ = Left$(Trim$(names$), Len(Trim(names$)))
    ForAddRoomCombo names$, Lst
Next Index
endaddroom:
Lst.RemoveItem Lst.ListCount - 1
i = GetListIndex(Lst, AOLUser())
If i <> -2 Then Lst.RemoveItem i
End Sub
Sub SendCharNum(win, chars)
E = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub
Sub MailSend(SN, Subject, Message)
'This will send mail from AOL4.0

aol% = FindWindow("AOL Frame25", vbNullString)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
Toolbar% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
ico% = FindChildByClass(Toolbar%, "_AOL_Icon")
ico% = GetWindow(ico%, 2)
Click ico%
Do
mail% = Findchildbytitle(AOLMDI(), "Write Mail")
edit% = FindChildByClass(mail%, "_AOL_Edit")
rich% = FindChildByClass(mail%, "RICHCNTL")
icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And edit% <> 0 And rich% <> 0 And icon% <> 0
Call SendMessageByString(edit%, WM_SETTEXT, 0, SN)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
Call SendMessageByString(edit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(rich%, WM_SETTEXT, 0, Message)
For GetIcon = 1 To 18
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next GetIcon
Do
Call Click(icon%)
pause 0.6
mail% = Findchildbytitle(AOLMDI(), "Write Mail")
If mail% = 0 Then Exit Do
Loop
End Sub
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
Function ScrambleText(thetext)
findlastspace = Mid(thetext, Len(thetext), 1)
If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$
If thechar$ = " " Then
chars$ = Mid(Char$, 1, Len(Char$) - 1)
firstchar$ = Mid(chars$, 1, 1)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs
sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "
sniffs:
Char$ = ""
backchar$ = ""
End If
Next scrambling
ScrambleText = scrambled$
Exit Function
End Function

Public Sub SpiralScroll(what$)
wowlen = Len(what$)
wowsend$ = what$ + " "
ChatSend (wowsend$)
pause 1
For X = 1 To wowlen
    wowbck$ = Mid(wowsend$, 1, 1)
    wownew$ = Mid(wowsend$, 2, wowlen)
    wowsend$ = wownew$ + wowbck$
    ChatSend (wowsend$)
    pause 0.7
Next X
ChatSend (what$)

End Sub
Function ReplaceWords(text, charfind, charchange)
'Kicks replacecharacter's ass..finds
'multiple characters instead of one
hek = InStr(text, charfind)
If hek = 0 Then
ReplaceWords = text
Exit Function
End If
z = 0
phrig$ = text
Do
z = z + 1
newz = InStr(z, phrig$, charfind)
If newz = 0 Then GoTo imcool
F = newz - z
ar$ = Left$(phrig$, F)
ee$ = Mid$(phrig$, F + Len(charfind) + 1)
z = newz + 1
phrig$ = ar$ + charchange + ee$
Loop
imcool:
thechars$ = phrig$
ReplaceWords = thechars$
End Function
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
Call Click(AOIcon%)
End Sub
Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
Call Click(AOIcon%)
End Sub
Sub AOLKillModals()
Dim the As String
modal% = FindWindow("_AOL_Modal", vbNullString)
the = -1
Do:
If modal% = 0 Then Exit Do
modal% = FindWindow("_AOL_Modal", vbNullString)
WinKill (modal%)
the = the + 1
Loop
If the < 1 Then
MsgBox "You have no Modal Windows open", vbExclamation
Exit Sub
End If
If the = 1 Then
MsgBox "1 Modal Window has been destroyed!", vbInformation
Exit Sub
End If
If the > 1 Then
MsgBox the + " Modal Windows have been destroyed!", vbInformation
Exit Sub
End If
End Sub

Sub AOLKillWait()
Call AOLRunMenuByString("&About America Online")
Do
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
pause 0.00001
Loop Until modal% <> 0
Do: FreeProcess
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Call Click(AOIcon%)
pause 0.00001
Loop Until modal% = 0
End Sub


Function WavY(text$) As String
'makes chat wavy
lent = Len(text$)
er$ = ""
For w = 1 To a Step 4
    ar$ = Mid$(text$, w, 1)
    br$ = Mid$(text$, w + 1, 1)
    cr$ = Mid$(text$, w + 2, 1)
    dr$ = Mid$(text$, w + 3, 1)
    er$ = er$ & "<sup>" & ar$ & "</sup>" & br$ & "<sub>" & cr$ & "</sub>" & dr$
Next w
WavY = er$
End Function
Function RGB2HEX(RedGreenBlue) As String
'LiviD Rocks =]
HexVal$ = Hex(RedGreenBlue)
ZeroFact% = Len(HexVal$)
HexVal$ = String(6 - ZeroFact%, "0") + HexVal$
RGB2HEX = HexVal$
End Function
Sub KillListDupes(Lst As ListBox)
'Kill the duplicates in a listbox

For X = 0 To Lst.ListCount - 1
    current = Lst.List(X)
    For i = 0 To Lst.ListCount - 1
        Nower = Lst.List(i)
        If i = X Then GoTo dontkill
        If TrimSpaces(LCase(Nower)) = TrimSpaces(LCase(current)) Then Lst.RemoveItem (i)
dontkill:
    Next i
Next X
End Sub
Sub WinMin(hWnd%)
Call ShowWindow(hWnd%, SW_MINIMIZE)
End Sub
Sub WinRestore(hWnd%)
Call ShowWindow(hWnd%, SW_RESTORE)
End Sub
Sub WinMax(hWnd%)
Call ShowWindow(hWnd%, SW_MAXIMIZE)
End Sub
Sub WinKill(hWnd%)
Call SendMessage(hWnd%, WM_CLOSE, 0, 0)
End Sub
Sub WinHide(hWnd%)
Call ShowWindow(hWnd%, SW_HIDE)
End Sub
Sub WinShow(hWnd%)
Call ShowWindow(hWnd%, SW_SHOW)
End Sub
Public Function ForAddRoom(Index As Long, Buffer As String)
'DONT EDIT
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = ChatFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
Buffer$ = Person$
End Function
Function MSGTOP(what$, typer As VbMsgBoxStyle, Title$) As VbMsgBoxResult
'-LiviD =]
'Puts a msgbox on top of a form that
'uses api to be on top and has no
'title bar
X = MsgBox(what$, typer + vbSystemModal, Title$)
MSGTOP = X
End Function
Function InputTOP(what$, Title$) As String
'sets an inputbox on top...
'i had probs with it, so if
'u get it to werk right then tell me
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = InputBox(what$, Title$)
InputTOP = X
End Function
Function ChatSendBox()
room% = ChatFindRoom()
aR1% = FindChildByClass(room%, "RICHCNTL")
aR2% = GetWindow(aR1%, 2)
aR3% = GetWindow(aR2%, 2)
aR4% = GetWindow(aR3%, 2)
aR5% = GetWindow(aR4%, 2)
aR6% = GetWindow(aR5%, 2)
ar7% = GetWindow(aR6%, 2)
ChatSendBox = ar7%
End Function
Sub INIWrite(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub

Function INIRead(AppName, KeyName As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INIRead = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReplaceText(text, charfind, charchange)
'Replaces the text what your own charectors
If InStr(text, charfind) = 0 Then
ReplaceText = text
Exit Function
End If
For Replace = 1 To Len(text)
thechar$ = Mid(text, Replace, 1)
thechars$ = thechars$ & thechar$
If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace
ReplaceText = thechars$

End Function
Function ReverseText(text)
'Reverse's the text
'EX: ReverText txt%
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words
End Function
Sub IMSendAIM(who$, what$)
aim% = FindWindow("_Oscar_BuddyListwin", vbNullString)
If aim% = 0 Then Exit Sub
tab1% = FindChildByClass(aim%, "_Oscar_TabGroup")
tab2% = FindChildByClass(tab1%, "_Oscar_IconBtn")
tab2% = GetWindow(tab2%, 2)
Click tab2%
Do
im% = FindWindow("AIM_IMessage", vbNullString)
Combo% = FindChildByClass(im%, "_Oscar_PersistantCombo")
edit% = FindChildByClass(Combo%, "Edit")
aoeditt% = FindChildByClass(im%, "WndAte32class")
rich% = GetWindow(aoeditt%, 2)
but% = FindChildByClass(im%, "_Oscar_IconBtn")
pause 0.001
Loop Until im% <> 0 And edit% <> 0 And rich% <> 0 And but% <> 0
Call SetText(edit%, who$)
Call SetText(rich%, what$)
Do: FreeProcess
Call Click(but%)
pause 0.5
im% = FindWindow("AIM_IMessage", "Instant Message")
msg% = FindWindow("#32770", "IM Information")
If msg% <> 0 Then
text% = FindChildByClass(msg%, "Button")
closer = SendMessage(msg%, WM_CLOSE, 0, 0)
closer2 = SendMessage(im%, WM_CLOSE, 0, 0)
Exit Sub
End If
Loop Until im = 0 And msg = 0
Do
im% = FindWindow("AIM_IMessage", vbNullString)
pause 0.001
Call WinKill(im%)
Loop Until im% = 0


End Sub



Sub AOLUpChat()
moda% = FindWindow("_AOL_Modal", vbNullString)
If moda% = 0 Then Exit Sub
m% = FindWindow("_AOL_Modal", vbNullString)
WinMin moda%
X% = EnableWindow(moda%, 0)
X% = EnableWindow(AOLWindow, 1)
End Sub
Sub AOLUnUpchat()
moda% = FindWindow("_AOL_Modal", vbNullString)
If moda% = 0 Then Exit Sub
m% = FindWindow("_AOL_Modal", vbNullString)
WinRestore moda%
X% = EnableWindow(moda%, 1)
X% = EnableWindow(AOLWindow, 0)
End Sub
Sub MinToAOL(formname As Form)
'u can thank xarc for this...
'put this in code in the form_load sub
Dim a As Variant
Dim b As Variant
Dim C As Variant

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTooL%, "_AOL_Icon")

a = FindWindow("AOL Frame25", vbNullString)
b = FindChildByClass(a, "MDICLIENT")
C = SetParent(formname.hWnd, aol%)
l = Screen.Width
J = Screen.Height
formname.Top = 40
formname.Left = 100300
End Sub
Sub SignOnAsGuest()
modaa% = FindWindow("#32769", vbNullString)
Wel% = Findchildbytitle(AOLMDI, "Goodbye From America Online")
wel2% = Findchildbytitle(AOLMDI, "Sign On")
If Wel% <> 0 Then scr% = Wel%
If wel2% <> 0 Then scr% = wel2%
Com% = FindChildByClass(scr%, "_AOL_Combobox")
Click Com%
AppActivate GetText(AOLWindow)
SendKeys "{PGDN}"
End Sub
Function Find2ndChildByClass(parentw, childhand)
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    Wend
    Find2ndChildByClass = 0
found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function TrimDate()
'by LiviD =]
'this will return the date as:
'September 1, 1998 instead of 9/1/98
s$ = Date
slash1 = InStr(s$, "/")
slash2 = InStr(slash1 + 1, s$, "/")
mo$ = Left$(s$, slash1 - 1)
da$ = Mid$(s$, slash1 + 1, 2)
yr$ = Right$(s$, 2)
If Left$(da$, 1) = "0" Then
da$ = Right$(da$, 1)
End If
If Right$(da$, 1) = "/" Then da$ = Left$(da$, 1)
If mo$ = "1" Then mo$ = "January"
If mo$ = "2" Then mo$ = "February"
If mo$ = "3" Then mo$ = "March"
If mo$ = "4" Then mo$ = "April"
If mo$ = "5" Then mo$ = "May"
If mo$ = "6" Then mo$ = "June"
If mo$ = "7" Then mo$ = "July"
If mo$ = "8" Then mo$ = "August"
If mo$ = "9" Then mo$ = "September"
If mo$ = "10" Then mo$ = "October"
If mo$ = "11" Then mo$ = "November"
If mo$ = "12" Then mo$ = "December"
xyz$ = mo$ + " " + da$ + ", 19" + yr$
TrimDate = xyz$

End Function

Public Sub MailAddNewToListBox(ListBo As ListBox)
ListBo.MousePointer = 11
aol% = FindWindow("AOL Frame25", vbNullString)
mail% = Findchildbytitle(AOLMDI(), AOLUser() & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tree% = FindChildByClass(tabp%, "_AOL_Tree")
z = 0
For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1
Buff$ = String$(100, 0)
X = SendMessageByString(tree%, LB_GETtext, i, Buff$)
subj$ = Mid$(Buff$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga = Right(subj$, Len(subj$) - Layz)
ListBo.AddItem Str(z) + ")  " + Trim(nigga)
z = z + 1
Next i
ListBo.MousePointer = 0
End Sub

Public Sub MailAddOldToListBox(ListBo As ListBox)
ListBo.MousePointer = 11
aol% = FindWindow("AOL Frame25", vbNullString)
mail% = Findchildbytitle(AOLMDI(), AOLUser() & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tree% = FindChildByClass(tabp%, "_AOL_Tree")
z = 0
For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1
Buff$ = String$(100, 0)
X = SendMessageByString(tree%, LB_GETtext, i, Buff$)
subj$ = Mid$(Buff$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga = Right(subj$, Len(subj$) - Layz)
ListBo.AddItem Str(z) + ")  " + Trim(nigga)
z = z + 1
Next i
ListBo.MousePointer = 0
End Sub
Sub MailOpenNew()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call Click(icon%)
End Sub
Sub MailWaitNew()
mail% = Findchildbytitle(AOLMDI(), AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Do
pause 0.1
mail% = Findchildbytitle(AOLMDI(), AOLUser + "'s Online Mailbox")
Loop Until mail% <> 0
WinMin mail%
lis% = tree%
Do
FreeProcess
M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
M2% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
M3% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
Loop Until M1% = M2% And M2% = M3%
M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
End Sub
Public Function MailGetNewTitle(Index) As String
'returns the title of the specified index in new mail
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
aoltree% = FindChildByClass(tabp%, "_AOL_Tree")
txtlen% = SendMessageByNum(aoltree%, LB_GETTEXTLEN, Index, 0&)
txt$ = String(txtlen% + 1, 0&)
X = SendMessageByString(aoltree%, LB_GETtext, Index, txt$)
subj$ = Mid$(txt$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga$ = Right(subj$, Len(subj$) - Layz)
MailGetNewTitle = nigga$
End Function
Public Function MailGetOldTitle(Index) As String
'returns the title of the specified index in new mail

aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
aoltree% = FindChildByClass(tabp%, "_AOL_Tree")

'de = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
txtlen% = SendMessageByNum(aoltree%, LB_GETTEXTLEN, Index, 0&)
txt$ = String(txtlen% + 1, 0&)
X = SendMessageByString(aoltree%, LB_GETtext, Index, txt$)
subj$ = Mid$(txt$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga$ = Right(subj$, Len(subj$) - Layz)
MailGetOldTitle = nigga$
End Function
Sub MailWaitOld()
mail% = Findchildbytitle(AOLMDI(), AOLUser() & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Do
mail% = Findchildbytitle(AOLMDI(), AOLUser + "'s Online Mailbox")
Loop Until mail% <> 0
lis% = tree%
WinMin mail%
Do
FreeProcess
M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
M2% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
M3% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
pause 2
Loop Until M1% = M2% And M2% = M3%
M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)
End Sub
Sub MailClickRead()
mail% = Findchildbytitle(AOLMDI, AOLUser & "'s Online Mailbox")
Read% = FindChildByClass(mail%, "_AOL_Icon")
Call Click(Read%)
End Sub
Sub MailOpenOld()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabp% = FindChildByClass(mail%, "_AOL_TabControl")
Call SendCharNum(tabp%, vbKeyRight)
End Sub
Sub MailIgnoreNew(num)
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Call MailSelectNew(num)
Call SendMessage(tree%, WM_COMMAND, 515, 0)
End Sub
Sub MailIgnoreOld(num)
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Call MailSelectOld(num)
Call SendMessage(tree%, WM_COMMAND, 515, 0)
End Sub
Sub MailSelectOld(Number)
'selects a specified mail in you old mail

aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
aoltree% = FindChildByClass(tabp%, "_AOL_Tree")
If aoltree% = 0 Then Exit Sub
de = SendMessage(aoltree%, LB_SETCURSEL, Number, 0)
End Sub
Sub MailSelectNew(Number)
'select a specified mail in your new mail
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mail% = Findchildbytitle(MDI%, AOLUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
aoltree% = FindChildByClass(tabp%, "_AOL_Tree")
'If aoltree% = 0 Then Exit Sub
de = SendMessage(aoltree%, LB_SETCURSEL, Number, 0)
End Sub
Function FileOpenAsBinary(path$) As String
'By LiviD =]
Filenum = FreeFile
Open path$ For Binary As #Filenum
Anti$ = String$(LOF(Filenum), " ")
Get #Filenum, , Anti$
File_OpenAsBinary = Anti$
Close #Filenum
End Function
Sub FileSaveAsBinary(path$, what$, frm As Form)
'By LiviD =]
On Error GoTo Lover
Filenum = FreeFile
Open path$ For Binary As #Filenum
Put #Filenum, 1, what$
Close #Filenum
Exit Sub
Lover:
Close #Filenum
End Sub
Function FileSearch(FileName As String, SearchString As String) As Long
'searches through a file for a string
'if the string is found then it returns the
'place where its located
Free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
    text$ = Space(32000)
    Get #Free, X, text$
    Debug.Print X
    If InStr(1, text$, SearchString$, 1) Then
        Where = InStr(1, text$, SearchString$, 1)
        FileSearch = (Where + X) - 1
        Close #Free
        Exit For
    End If
    Next X
Close #Free
End Function

Function TrimTime() As String
s$ = Time$
colon = InStr(s$, ":")
our$ = Left$(s$, colon - 1)
colon2 = InStr(colon + 1, s$, ":")
minn$ = Mid$(s$, colon + 1, 2)
ou = Val(TrimSpaces(our$))
TrimTime = our$ + ":" + minn$
End Function
Sub BotEcho(persontoecho$)
'put in a timer with interval = 1:
'Call BotEcho("persons screen name to echo")
On Error GoTo hell
LastLine$ = ChatLastLine
If LastLine$ = OldLast$ Then Exit Sub
OldLast$ = LastLine$
whoend = InStr(LastLine$, ":")
who$ = Left$(LastLine$, whoend - 1)
what$ = Mid$(LastLine$, whoend + 3)
If LCase(TrimSpaces(who$)) = LCase(TrimSpaces(persontoecho$)) Then
Call ChatSend(what$)
End If
hell:
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
who$ = Left$(LastLine$, whoend - 1)
what$ = Mid$(LastLine$, whoend + 3)
If LCase(what$) = "/food" Then
num = RandomNumber(10)
Select Case num
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
osend$ = who$ + ", " + s$
ChatSend osend$
End If
hell:
End Sub

Sub ScrollMacro(Text1 As TextBox)

phrig$ = Text1.text
z = 0
ChatSend "Now Scrolling Macro..."
Do
z = z + 1
newz = InStr(z, phrig$, Chr(13))
If newz = 0 Then
ez$ = Mid$(phrig$, z)
Call ChatSend(ez$)
Call ChatSend("Macro Done!")
Exit Sub
End If
F = newz - z
r$ = Mid$(phrig$, z, F)
If newz <> 0 Then: ChatSend (r$)
z = newz + 1
Loop

End Sub
Function TrimTimer()
'this will take a portion of the timer
'so its a number like 5.2 instead of
'5.234423423
LaTime = Timer
scr$ = Str(LaTime)
Spot = InStr(scr$, ".")
If Spot = 0 Then GoTo klu
scr$ = Left$(scr$, Spot + 1)
klu:
TrimTimer = scr$
End Function
Sub IMMassIM(Lst As ListBox, whattosay$)
For X = 0 To Lst.ListCount - 1
Call IMSend(Lst.List(X), whattosay$)
Next X
End Sub

Sub IMIgnore(Lst As ListBox)
On Error GoTo vil
'put this in a timer
'call imignore(list1)
im% = Findchildbytitle(AOLMDI, ">Instant Message From: ")
If im% = 0 Then Exit Sub
wh$ = GetText(FindChildByClass(im%, "RICHCNTL"))
If wh$ = "" Then Exit Sub
whoend = InStr(wh$, ":")
If whoend = 0 Then Exit Sub
who$ = Left$(wh$, whoend - 1)
X = GetListIndex(Lst, who$)
If X > -2 Then
WinKill (im%)
End If
vil:
End Sub
Public Sub AOLClickToolBar(Number As Integer)
aol% = FindWindow("AOL Frame25", vbNullString)
tb% = FindChildByClass(aol%, "AOL Toolbar")
tc% = FindChildByClass(tb%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")
If Number = 1 Then
    Click (td%)
    Exit Sub
End If
For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T
Call Click(td%)
End Sub

Function ChatFade(what$, R1 As Integer, R2 As Integer, G1 As Integer, G2 As Integer, B1 As Integer, B2 As Integer, MakeWavy As Boolean) As String
'LiviD loves CoLoRs =P
textlen = Len(what$)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(what$, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy
        End If
    End If
AfterWavy:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
ChatFade = FadedText
End Function
Function ChatFade3(what$, R1, R2, R3, G1, G2, G3, B1, B2, B3, MakeWavy As Boolean) As String
'LiviD loves CoLoRs =P
textlen = Len(what$)
'Make 2 Different Strings:
Divide = Int(textlen / 2)
divide2 = textlen - Divide
fade1 = Left$(what$, Divide)
fade2 = Right$(what$, divide2)
'----------------------------------------
'Make first 2 colors..
textlen = Len(fade1)
WavyNumber = 0
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade1, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy1
        End If
    End If
AfterWavy1:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade1 = FadedText
'----------------------------------------
'Make Last 2 colors..
textlen = Len(fade2)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade2, MakeFade, 1)
    RGBVal = RGB((B3 - B2) / textlen * MakeFade + B2, (G3 - G2) / textlen * MakeFade + G2, (R3 - R2) / textlen * MakeFade + R2)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy2
        End If
    End If
AfterWavy2:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade2 = FadedText
ChatFade3 = fade1 + fade2
End Function
Function ChatFade4(what$, R1, R2, R3, R4, G1, G2, G3, G4, B1, B2, B3, B4, MakeWavy As Boolean) As String
'LiviD loves CoLoRs =P
thelen = Len(what$)
textlen = Len(what$)
Do
int1 = int1 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int2 = int2 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int3 = int3 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
Loop
fade1 = Left$(what$, int1)
fade2 = Mid$(what$, int1 + 1, int2)
Fade3 = Mid$(what$, int1 + int2 + 1, int3)
'----------------------------------------
'Make first 2 colors..
textlen = Len(fade1)
WavyNumber = 0
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade1, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy1
        End If
    End If
AfterWavy1:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade1 = FadedText
'----------------------------------------
'Make 2nd 2 colors..
textlen = Len(fade2)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade2, MakeFade, 1)
    RGBVal = RGB((B3 - B2) / textlen * MakeFade + B2, (G3 - G2) / textlen * MakeFade + G2, (R3 - R2) / textlen * MakeFade + R2)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy2
        End If
    End If
AfterWavy2:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade2 = FadedText
'----------------------------------------
'Make Last 2 colors..
textlen = Len(Fade3)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(Fade3, MakeFade, 1)
    RGBVal = RGB((B4 - B3) / textlen * MakeFade + B3, (G4 - G3) / textlen * MakeFade + G3, (R4 - R3) / textlen * MakeFade + R3)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavyq
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavyq
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavyq
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavyq
        End If
    End If
AfterWavyq:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
Fade3 = FadedText
ChatFade4 = fade1 + fade2 + Fade3
End Function
Function ChatFade5(what$, R1, R2, R3, R4, R5, G1, G2, G3, G4, G5, B1, B2, B3, B4, B5, MakeWavy As Boolean) As String
'LiviD loves CoLoRs =P
thelen = Len(what$)
textlen = Len(what$)
Do
int1 = int1 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int2 = int2 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int3 = int3 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int4 = int4 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
Loop
fade1 = Left$(what$, int1)
fade2 = Mid$(what$, int1 + 1, int2)
Fade3 = Mid$(what$, int1 + int2 + 1, int3)
fade4 = Mid$(what$, int1 + int2 + int3 + 1, int4)
'----------------------------------------
'Make first 2 colors..
textlen = Len(fade1)
WavyNumber = 0
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade1, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy1
        End If
    End If
AfterWavy1:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade1 = FadedText
'----------------------------------------
'Make 2nd 2 colors..
textlen = Len(fade2)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade2, MakeFade, 1)
    RGBVal = RGB((B3 - B2) / textlen * MakeFade + B2, (G3 - G2) / textlen * MakeFade + G2, (R3 - R2) / textlen * MakeFade + R2)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavys
        End If
    End If
AfterWavys:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade2 = FadedText
'----------------------------------------
'Make Last 2 colors..
textlen = Len(Fade3)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(Fade3, MakeFade, 1)
    RGBVal = RGB((B4 - B3) / textlen * MakeFade + B3, (G4 - G3) / textlen * MakeFade + G3, (R4 - R3) / textlen * MakeFade + R3)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavya
        End If
    End If
AfterWavya:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
Fade3 = FadedText
'----------------------------------------
'Make 2nd 2 colors..
textlen = Len(fade4)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade4, MakeFade, 1)
    RGBVal = RGB((B5 - B4) / textlen * MakeFade + B4, (G5 - G4) / textlen * MakeFade + G4, (R5 - R4) / textlen * MakeFade + R4)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy2
        End If
    End If
AfterWavy2:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade4 = FadedText
ChatFade5 = fade1 + fade2 + Fade3 + fade4
End Function

Function Fade6Colors(what$, R1, R2, R3, R4, R5, R6, G1, G2, G3, G4, G5, G6, B1, B2, B3, B4, B5, B6, MakeWavy As Boolean) As String
'LiviD loves CoLoRs =P
thelen = Len(what$)
textlen = Len(what$)
Do
int1 = int1 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int2 = int2 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int3 = int3 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int4 = int4 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
int5 = int5 + 1: thelen = thelen - 1
If thelen < 1 Then Exit Do
Loop
fade1 = Left$(what$, int1)
fade2 = Mid$(what$, int1 + 1, int2)
Fade3 = Mid$(what$, int1 + int2 + 1, int3)
fade4 = Mid$(what$, int1 + int2 + int3 + 1, int4)
fade5 = Mid$(what$, int1 + int2 + int3 + int4 + 1, int5)
'----------------------------------------
'Make first 2 colors..
textlen = Len(fade1)
WavyNumber = 0
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade1, MakeFade, 1)
    RGBVal = RGB((B2 - B1) / textlen * MakeFade + B1, (G2 - G1) / textlen * MakeFade + G1, (R2 - R1) / textlen * MakeFade + R1)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy1
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy1
        End If
    End If
AfterWavy1:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade1 = FadedText
'----------------------------------------
'Make 2nd 2 colors..
textlen = Len(fade2)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade2, MakeFade, 1)
    RGBVal = RGB((B3 - B2) / textlen * MakeFade + B2, (G3 - G2) / textlen * MakeFade + G2, (R3 - R2) / textlen * MakeFade + R2)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavys
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavys
        End If
    End If
AfterWavys:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade2 = FadedText
'----------------------------------------
'Make third 2 colors..
textlen = Len(Fade3)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(Fade3, MakeFade, 1)
    RGBVal = RGB((B4 - B3) / textlen * MakeFade + B3, (G4 - G3) / textlen * MakeFade + G3, (R4 - R3) / textlen * MakeFade + R3)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavya
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavya
        End If
    End If
AfterWavya:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
Fade3 = FadedText
'----------------------------------------
'Make 4th 2 colors..
textlen = Len(fade4)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade4, MakeFade, 1)
    RGBVal = RGB((B5 - B4) / textlen * MakeFade + B4, (G5 - G4) / textlen * MakeFade + G4, (R5 - R4) / textlen * MakeFade + R4)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy2
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy2
        End If
    End If
AfterWavy2:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade4 = FadedText
'----------------------------------------
'Make last 2 colors..
textlen = Len(fade5)
FadedText = ""
For MakeFade = 1 To textlen Step 1
    Char = Mid$(fade5, MakeFade, 1)
    RGBVal = RGB((B6 - B5) / textlen * MakeFade + B5, (G6 - G5) / textlen * MakeFade + G5, (R6 - R5) / textlen * MakeFade + R5)
    RGBColor = RGB2HEX(RGBVal)
    If MakeWavy = True Then
        WavyNumber = WavyNumber + 1
        If WavyNumber = 1 Then
            Char = "<sup>" + Char
            GoTo AfterWavy8
        ElseIf WavyNumber = 2 Then
            Char = "</sup>" + Char
            GoTo AfterWavy8
        ElseIf WavyNumber = 3 Then
            Char = "<sub>" + Char
            GoTo AfterWavy8
        ElseIf WavyNumber = 4 Then
            WavyNumber = 0
            Char = "</sub>" + Char
            GoTo AfterWavy8
        End If
    End If
AfterWavy8:
FadedText = FadedText + "<Font Color=" + Chr(34) + "#" + RGBColor + Chr(34) + ">" + Char
Next MakeFade
fade5 = FadedText
ChatFade6 = fade1 + fade2 + Fade3 + fade4 + fade5
End Function
