Attribute VB_Name = "AOL4MMer"
'WSUP everybody!  This is Mission.  I hope
'you enjoy using my bas file.  It is
'intended to be used with another bas
'file.  This is only for a Mass Mailer.
'There will be further updates to this
'module.  It can be freely distributed,
'and change it as much as you'd like.  If
'you have any questions, or additions to
'this file, please E-Mail me at
'IcyHotMission@juno.com with your comment.
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
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

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function
Function AddListToString2(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString2 = AddListToString2 & TheList.List(DoList) & "@aol.com, "
Next DoList
AddListToString2 = Mid(AddListToString2, 1, Len(AddListToString2) - 2)
End Function
Sub AddMailList(List As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
tol% = FindChildByClass(tool%, "_AOL_Toolbar")
Mail% = FindChildByClass(tol%, "_AOL_Icon")
ClickIcon (Mail%)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
chi% = FindChildByClass(mdi%, "AOL Child")
tabb% = FindChildByClass(chi%, "_AOL_TabControl")
pag% = FindChildByClass(tabb%, "_AOL_TabPage")
tree% = FindChildByClass(pag%, "_AOL_Tree")
If tree% Then Exit Do
Loop
Do
DoEvents
X = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Call TimeOut(2)
xg = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Loop Until X = xg
X = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Z = 0
For I = 0 To X - 1
mailstr$ = String$(255, " ")
    q% = SendMessageByString(tree%, LB_GETTEXT, I, mailstr$)
    nodate$ = Mid$(mailstr$, InStr(mailstr$, "/") + 8)
    nosn$ = Mid$(nodate$, InStr(nodate$, Chr(9)) + 1)
    List.AddItem Z & ") " & Trim(nosn$)
    Z = Z + 1
Next I
Call KillDupes(List)
End Sub
Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function
Function AOLWindow()
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Sub ClickForward()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AoIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 8
AoIcon% = GetWindow(AoIcon%, 2)
NoFreeze% = DoEvents()
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ClickKeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
AoIcon% = FindChildByClass(Mailbox%, "_AOL_Icon")
For l = 1 To 2
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickNext()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AoIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 5
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickRead()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
AoIcon% = FindChildByClass(Mailbox%, "_AOL_Icon")
For l = 1 To 0
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickSendAndForwardMail(Recipiants)

aol% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AoRich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AoRich% <> 0 And AoIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)
For GetIcon = 1 To 14
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon
ClickIcon (AoIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub CloseWindow(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub
Function CountMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mail% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
TabControl% = FindChildByClass(Mail%, "_AOL_TabControl")
TabPage% = FindChildByClass(TabControl%, "_AOL_TabPage")
MailLB% = FindChildByClass(TabPage%, "_AOL_Tree")
CountMail = SendMessageByNum(MailLB%, LB_GETCOUNT, 0&, 0&)
End Function
Sub DeleteItem(Lst As ListBox, Item$)
On Error Resume Next
Do
NoFreeze% = DoEvents()
If LCase$(Lst.List(A)) = LCase$(Item$) Then Lst.RemoveItem (A)
A = 1 + A
Loop Until A >= Lst.ListCount
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
Sub ForwardMail(Recipiants, Message)

aol% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AoRich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AoRich% <> 0 And AoIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AoRich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 14
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

ClickIcon (AoIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Sub IMBuddy(Recipiant, Message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")

If buddy% = 0 Then
Call IMKeyword(Recipiant, Message)
Exit Sub
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
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, Message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

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
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMsOff()
Call IMBuddy("$IM_OFF", " ")
End Sub
Sub IMsOn()
Call IMBuddy("$IM_ON", " ")
End Sub
Sub Keyword(TheKeyword As String)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call TimeOut(0.05)
ClickIcon (AoIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyword)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub KillDupes(Lst As ListBox)
For X = 0 To Lst.ListCount - 1
Current = Lst.List(X)
For I = 0 To Lst.ListCount - 1
Nower = Lst.List(I)
If I = X Then GoTo dontkill
If Nower = Current Then Lst.RemoveItem (I)
dontkill:
Next I
Next X
End Sub
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
'This is used like:
'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)
'where Label1 is how many mails have
'already been forwarded, and Label2 is
'how many total mails there are.
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print Percent(Done, Total, 100) & "%"
End Sub
Sub ReadMail()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
ToolBar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(ToolBar%, "_AOL_Icon")
Call ClickIcon(icon%)
End Sub
Sub SetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function
Sub WaitForMailToLoad()
ReadMail
Do
Box% = FindChildByTitle(AOLMDI(), UserSN & "'s Online Mailbox")
Loop Until Box% <> 0
List = FindChildByClass(Box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M2% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
ClickRead
End Sub
Function WaitForWin(Caption As String) As Integer
Do
DoEvents
win% = FindChildByTitle(AOLMDI, Caption$)
Loop Until win% <> 0
WaitForWin = win%
End Function
