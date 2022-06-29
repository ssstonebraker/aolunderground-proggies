Attribute VB_Name = "LimP"
'yo dis is eklipse
'and this is LimP.BASv1
'v2 will be out sooner or later
'but for now im gonna work on my
'prog, so if you need any help
'email me at lil-limp@juno.com
'most emails get sent back to you
'within a week
'lates peepz!
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

Declare Function ReleaseCapture Lib "user32" () As Long

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
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

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
Private Const WM_NCLBUTTONDOWN = &HA1

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

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function FindChatRoom()
'finds the AOL chat room
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
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
'gets the users screen name
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub NOWait()
'hate waiting?..use this sub to end the wait
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AoIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function UserOnline()
'are they online?
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
End Function

Sub SendChat(Chat)
'sends text to the chat of course
Room% = FindChatRoom
AoRich% = FindChildByClass(Room%, "RICHCNTL")

AoRich% = GetWindow(AoRich%, 2)
AoRich% = GetWindow(AoRich%, 2)
AoRich% = GetWindow(AoRich%, 2)
AoRich% = GetWindow(AoRich%, 2)
AoRich% = GetWindow(AoRich%, 2)
AoRich% = GetWindow(AoRich%, 2)

Call SetFocusAPI(AoRich%)
Call SendMessageByString(AoRich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AoRich%, WM_CHAR, 13, 0)
End Sub

Sub TimeOut(Duration)
'stops for the entered duration
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Sub StayOnTop(TheForm As Form)
'makes form stay on top of everyting else
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Sub Anti45()
'keeps that 45 minute bullshit thing away
'i doint think they use the 45 minute timer anymore but oh well
'here it is anyways
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AoIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AoIcon%)
End Sub
Sub NOIdle()
'finds and closes that annoying little
'you have been idle ...blah blah blah.. window
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AoIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AoIcon%)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Mail(Recipiants, subject, Message)
'sends an email
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AoIcon% = GetWindow(AoIcon%, 2)

ClickIcon (AoIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AoRich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AoRich% <> 0 And AoIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AoRich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 18
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

ClickIcon (AoIcon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AoIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AoIcon%)
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

Sub Keyword(TheKeyword As String)
'calls up the keyword window on AOL
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AoIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyword)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")

If buddy% = 0 Then
    Keyword ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AoIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For l = 1 To 2
    AoIcon% = GetWindow(AoIcon%, 2)
Next l

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AoRich% = FindChildByClass(IMWin%, "RICHCNTL")
AoIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AoRich% <> 0 And AoIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AoRich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AoIcon% = GetWindow(AoIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IM(Recipiant, Message)
'sends an IM

AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")

Call Keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AoRich% = FindChildByClass(IMWin%, "RICHCNTL")
AoIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AoRich% <> 0 And AoIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AoRich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AoIcon% = GetWindow(AoIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetchatText()
Room% = FindChatRoom
AoRich% = FindChildByClass(Room%, "RICHCNTL")
ChatText = GetText(AoRich%)
GetchatText = ChatText
End Function

Function LastChatLineWithSN()
'gets last chat line with the screen name who typed it
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
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function SNFromLastChatLine()
' gets the screen name of the last line of chat
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function
'Last Chat Line, gets the last thing somone typed
Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToList(ListBox As ListBox)
'adds the room to a list
'use it like Call AddRoomToList (List#)
'where list# would be the name of your list
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

Sub AddRoomToCombo(ListBox As ListBox, ComboBox As ComboBox)
'adds the room to a combo box
'use it like Call AddRoomToCombo (combo#)
'where combo# would be the name of your combo box
Call AddRoomToListBox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
End Sub

Sub WavyChatBB(thetext)
'sends the text to the chat wavy
'in blue and black
'use it like WavyChatBB(text#)
'where text# would be your text box's name
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

Sub EliteTalker(word$)
'sends the text in the textbox to the xhat
'in elite form
'use it like Call EliteTalker(Text#)
'where text# is the name of your text box
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
SendChat (Made$)
End Sub

Sub imon()
'turns IM'z on
Call IMKeyword("$IM_ON", " ")
End Sub
Sub imoff()
'turns IM'z off
Call IMKeyword("$IM_OFF", " ")
End Sub



Sub KillGlyph()
'this sub gets rid of that stupid little
'spinning AOL icon in the corner

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

EliteText = Made$

End Function

Sub NO_IM(TheList As ListBox)
'closes IM'z as fast as you recieve them
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNFromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function SNFromIM()
'gets the screen name from an instant message
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNFromIM = TheSN$

End Function

Sub PlayWav(File)
'plays a wav file
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
'kills the modal
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
'waits for the message box or alert box and then
'clicks OK to close it
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

Function Wavy(thetext As String)
'sends the text to the chat all wavy like
'use it like Wavy(Text#) to use it with a text box
'where text# is the name of your text box
'or use it just like it is like:
'Wavy(this is what u want to be wavy)

G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & S$ & "</sub>" & T$
Next w
Wavy = P$

End Function

Sub CenterForm(F As Form)
'centers the form on the screen DuH!
'ex: CenterForm(Form#)
'form number is the name of your form
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub RespondIM(Message)
'finds the instant message you recieved,
'reply's with "message",
'and closes it
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

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
Call SendMessageByString(e2, WM_SETTEXT, 0, Message)
ClickIcon (e)
Call TimeOut(0.8)
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
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
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
End Sub

Function MessageFromIM()
'gets the message from an instant message
'that you get
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNFromIM()
snlen = Len(SNFromIM()) + 3
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

Sub UpchatON()
'upload files and chat at the same time
'this usb turns it on!
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UpchatOFF()
'this sub turns it OFF!
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub AOLHide()
'wanna guess what this does?
'it hides AOL
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub AOLshow()
'ooh another hard one..whats it do?:
'shows AOL
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub

Function RoomName()
'gets the name of the room you are in
'text1.text=roomname()
Call GetCaption(AOLFindChatRoom)
End Function

Function spaces(strin As String)
'spaces out the letters you type
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "•"
Let newsent$ = newsent$ + nextchr$
Loop
r_dots = newsent$

End Function

Function lagg(strin As String)
'this laggs the text you type
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "<html>"
Let newsent$ = newsent$ + nextchr$
Loop
r_html = newsent$

End Function

Sub AOLOff()
'does what it says..this sub
'signs your account off of AOL
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(AOL%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
ClickIcon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
ClickIcon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function EncryptType(Text, types)
'to encrypt:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt:
'decrypted$ = EncryptType("decryptedmessage", 1)
'first paramneter is the text to encrypt
'second is 0 for encrypt
'or 1 for decrypt

For God = 1 To Len(Text)
If types = 0 Then
current$ = Asc(Mid(Text, God, 1)) - 1
Else
current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(current$)
Next God

EncryptType = Process$
End Function
Function RanNum(finished)
'just gives you a random number.
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub ReadMail()
'reads your mail
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call ClickIcon(icon%)
End Sub

Public Sub AddNewMailToListBox(ListBo As ListBox)
'adds your new mail to a listbox.
ListBo.MousePointer = 11
AOL% = FindWindow("AOL Frame25", vbNullString)
Mail% = FindChildByTitle(AOLMDI(), AOLUser() & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Z = 0
For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1
Buff$ = String$(100, 0)
X = SendMessageByString(tree%, LB_GETTEXT, i, Buff$)
subj$ = Mid$(Buff$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga = Right(subj$, Len(subj$) - Layz)
ListBo.AddItem Str(Z) + ")  " + Trim(nigga)
Z = Z + 1
Next i
ListBo.MousePointer = 0

End Sub


Public Sub AddOldMailToListBox(ListBo As ListBox)
'adds your old mail to a list box
ListBo.MousePointer = 11
AOL% = FindWindow("AOL Frame25", vbNullString)
Mail% = FindChildByTitle(AOLMDI(), AOLUser() & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tree% = FindChildByClass(tabp%, "_AOL_Tree")
Z = 0
For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1
Buff$ = String$(100, 0)
X = SendMessageByString(tree%, LB_GETTEXT, i, Buff$)
subj$ = Mid$(Buff$, 14, 80)
Layz = InStr(subj$, Chr(9))
nigga = Right(subj$, Len(subj$) - Layz)
ListBo.AddItem Str(Z) + ")  " + Trim(nigga)
Z = Z + 1
Next i
ListBo.MousePointer = 0

End Sub

Sub MacroDraw(Text As String)
'this scrolls the text box in a chat room
'but it scrolls multi lines
'best for scrolling macros
'thats why its called MacrDraw
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub


Sub KillDupes(lst As ListBox)
'kills duplicate items
For X = 0 To lst.ListCount - 1
current = lst.List(X)
For i = 0 To lst.ListCount - 1
Nower = lst.List(i)
If i = X Then GoTo dontkill
If Nower = current Then lst.RemoveItem (i)
dontkill:
Next i
Next X
End Sub

Function CountMail()
'counts your mail...
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
TabControl% = FindChildByClass(Mailbox%, "_AOL_TabControl")
TabPage% = FindChildByClass(TabControl%, "_AOL_TabPage")
MailLB% = FindChildByClass(TabPage%, "_AOL_Tree")
CountMail = SendMessageByNum(MailLB%, LB_GETCOUNT, 0&, 0&)
End Function


Function StringInList(TheList As ListBox, FindMe As String)
If TheList.ListCount = 0 Then GoTo nope
For a = 0 To TheList.ListCount - 1
TheList.ListIndex = a
If UCase(TheList.Text) = UCase(FindMe) Then
StringInList = a
Exit Function
End If
Next a
nope:
StringInList = -1
End Function

Function ScrambleText(thetext)
'scrambles words man!!!!
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
On Error GoTo error
lastchar$ = Mid(chars$, Len(chars$), 1)
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo stuff
error:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo stuffs
stuff:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "
stuffs:
Char$ = ""
backchar$ = ""
End If
Next scrambling
ScrambleText = scrambled$
Exit Function

End Function

Function ReverseText(Text As String)

'reverses charachters in a string
'example: Text1 = ReverseText(Text1)
'This will take the string in Text1 and reverse it
'If Text1 was "Hello" it will now be "olleH"On Error GoTo error
Dim words As Integer
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words
Exit Function
error: MsgBox "Sorry Message Too Long: Caused Overflow Error!", vbOKOnly, "Overflow Error"
Err = 1
    End Function


'form fades begin here
'to use them correctly do it like this:
'On Error Resume Next
'Dim intLoop As Integer
'Form#.DrawStyle = vbInsideSolid
'Form#.DrawMode = vbCopyPen
'Form#.ScaleMode = vbPixels
'Form#.DrawWidth = 2
'Form#.ScaleHeight = 256
'For intLoop = 0 To 255
'Form#.Line (0, intLoop)-(Screen.Width, intLoop - 1), rgb(0, 0, 255 - intLoop), B
'Next intLoop
'form# is the the name of your form..

'work best in the form_paint()


Sub Blue(vForm As Form)
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

Sub Green(vForm As Form)
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

Sub Grey(vForm As Form)
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

Sub Purple(vForm As Form)
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

Sub Red(vForm As Form)
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

Sub Yellow(vForm As Form)
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

'----------------------------------------------
'start  yellow fades..all are preset
'----------------------------------------------
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    YellowRedYellow = Msg
End Function

'end yellow fades





'----------------------------------
'start the blue fades...all are preset
'-----------------------------------

Function BlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlueBlack = Msg
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlueGreen = Msg
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BluePurple = Msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlueRed = Msg
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlueYellow = Msg
End Function

'end blue fades




'--------------------------------
'start grey color phades
'--------------------------------

Function GreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 220 / a
        F = e * b
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyBlack = Msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyBlue = Msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyGreen = Msg
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyPurple = Msg
End Function

Function GreyRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyRed = Msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    GreyYellow = Msg
End Function

'end grey fades




'--------------------------
'start red fades
'--------------------------

Function RedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    RedBlack = Msg
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    RedBlue = Msg
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    RedGreen = Msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    RedPurple = Msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    RedYellow = Msg
End Function

'end red fades





'----------------------
'2 color black fades
'----------------------

Function BlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackBlue = Msg
End Function

Function BlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackGreen = Msg
End Function

Function BlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 220 / a
        F = e * b
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackGrey = Msg
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackPurple = Msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackRed = Msg
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    BlackYellow = Msg
End Function

'end black fades






