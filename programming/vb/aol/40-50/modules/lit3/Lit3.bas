Attribute VB_Name = "Lit3Bas"
'Lit3's .Bas
'Made in Visual Basics 6.0
'Don't be coping my shit
'If you have any ?'s Then E=Mail me At DemonLit3@aol.com
'This .Bas Is for the use of making AOL4.0 Progs
'       ,,,,       ,,,              ,,;;;;;,                         ,,,,,,,,,         ,,
'      ´;;;;;;;;;;;;´´             ;;;;;;;;;    ,, ,,,,,;;´       ,;;;;;;;;;;;;;;;;;;;;;;
'        ´;;;;;;;;;;                ;;;;;;;;    ´;;;;;;;;;       ,;;;;;;;;;;;;;;;;;;;;;´
'         ;;;;;;;;;                 ´;;;;´´      ;;;;;;;;        ;;;;;;  ´´´´´;,;;;;´
'        ;;;;;;;;;               ;;;,,,,,;;´     ;;;;;;;,,,,;;  ;;;;;;;;   ,,;;;;´
'        ;;;;;;;;                 ;;;;;;;;    ,,;;;;;;;;;;;;;´   ;;;;´´  ,;;;;;,
'       ;;;;;;;;;  ,;;;;;;;;,    ;;;;;;;;   ´´´´;;;;;;;;;;;;´         ,;;;;;;;;;;;;;,,
'      ,;;;;;;;; ,;;;;;;;;;;;    ;;;;;;;       ;;;;;;;                 ´´´  ´;;;;;;;;;;
'      ;;;;;;;; ,;;;;;;;;;;;;   ,;;;;;;       ;;;;;;;´         ,,;;;;;;,     ;;;;;;;;;;
'     ;;;;;;;;;  ´´´;;;;;;;;;  ,;;;;;;;      ,;;;;;;;,;;;;;, ,;;;;;;;;       ;;;;;;;;;;
'    ;;;;;;;;;       ;;;;;;;   ;;;;;;;      ,;;;;;;;;;;;;;;; ;;;;;;;;       ;;;;;;;;;;
'  ,;;;;;;;;;;      ;;;;;;´   ;;;;;;;´    ,;;;;;;;;´´´;;;;;´´;;;;;;;;     ,;;;;;;;;;´
',;;;;;;;;;;;;;,,,;;;;;;´   ,;;;;;;;;,     ´;;;;;;;  ;;;;;´  ;;;;;;;;;,,,;;;;;;;;´´
';;´´    ´´´;;;;;;;;´´     ;;´   ´´´;;       ´;;;;;,;;´´´      ´´;;;;;;;;;;;´´´











Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
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

Public Enum MAILTEXT
        mtDATE
        mtSENDER
        mtSUBJECT
        mtALL
End Enum
Public Enum MAILTYPE
        mtFLASH
        mtNEW
        mtOLD
        mtSENT
End Enum





Public Const MOVE = &HA1
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
Sub TimePause(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Sub SendChatMsg(Chat)
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
Sub StayTopWin(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub MiniForm(frm As Form)
frm.WindowState = 1
End Sub
Sub MaxForm(frm As Form)
frm.WindowState = 2
End Sub
Sub NormalForm(frm As Form)
frm.WindowState = 0
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
Function UserScreenName()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function
Sub PauseForOk()
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
Sub AolShow()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Sub AOLHide()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub
Sub GlyphKill()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTOOL%, "_AOL_Toolbar")
glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(glyph%, WM_CLOSE, 0, 0)
End Sub
Sub KillToolBarIcon()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTOOL%, "_AOL_Toolbar")
Window% = FindChildByClass(AOTool2%, "_AOL_Icon")
Call SendMessage(Window%, WM_CLOSE, 0, 0)

End Sub
Sub KillBuddyList_ListBox()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "MDIClient")
AOTool2% = FindChildByClass(AOTOOL%, "AOLChild")
Window% = FindChildByClass(AOTool2%, "_AOL_Listbox")
Call SendMessage(Window%, WM_CLOSE, 0, 0)
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

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub
Sub KillWait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTOOL%, "_AOL_Toolbar")

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
Function FindChildByHandle(WHandle, childhand)
firs% = GetWindow(WHandle, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(WHandle, SW_SHOW = 5)

While firs%
firss% = GetWindow(WHandle, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByHandle = 5

bone:
Room% = firs%
FindChildByHandle = Room%
End Function
Sub Welcome_KillALLMessages()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "MDICLient")
AOTool2% = FindChildByClass(AOTOOL%, "AOL Child")
Window% = FindChildByClass(AOTool2%, "RICHCNTL")
Call SendMessage(Window%, WM_CLOSE, 0, 0)
End Sub
Sub Welcome_KillALLPictures()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTOOL% = FindChildByClass(AOL%, "MDICLient")
AOTool2% = FindChildByClass(AOTOOL%, "AOL Child")
Window% = FindChildByClass(AOTool2%, "_AOL_Icon")
Call SendMessage(Window%, WM_CLOSE, 0, 0)
End Sub
Sub IMs_On()
Call XxLit3_IMxX("$IM_On", "....,;`^~Lit3~^`;.... Turned Your IMsOn")
End Sub
Sub IMs_Off()
Call XxLit3_IMxX("$IM_OFF", "....,;`^~Lit3~^`;.... Turned Your IMsOff ")
End Sub
Sub XxLit3_IMxX(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call Keyword("aol://9293:")

Do: DoEvents
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(imwin%, "_AOL_Edit")
AORich% = FindChildByClass(imwin%, "RICHCNTL")
AOIcon% = FindChildByClass(imwin%, "_AOL_Icon")
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
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
okwin% = FindWindow("#32770", "America Online")
If okwin% <> 0 Then Call SendMessage(okwin%, WM_CLOSE, 0, 0): closer2 = SendMessage(imwin%, WM_CLOSE, 0, 0): Exit Do
If imwin% = 0 Then Exit Do
Loop

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
Sub Mini_Welcomescreen()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
Call SendMessage(welcome%, WM_CLOSE, 5, 5)
End Sub
Sub MessageBox()
'Code for messagebox put in your own message,vbokonly,andtitle
'MsgBox "message here",VBOKONLY , "title here"
End Sub
Sub ShowAol_Toolbar()
DeLTa& = FindWindow("AOL Frame25", vbNullString)
SocK& = FindChildByClass(DeLTa&, "AOL Toolbar")
PLoP& = ShowWindow(SocK&, 5)
End Sub
Sub HideAol_Toolbar()
DeLTa& = FindWindow("AOL Frame25", vbNullString)
SocK& = FindChildByClass(DeLTa&, "AOL Toolbar")
PLoP& = ShowWindow(SocK&, 0)
End Sub
Public Function FindMailWin() As Long
    Dim AOL As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", "America  Online")
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    FindMailWin& = FindWindowEx(MDI&, 0&, "AOL Child", GetUser() & "'s Online Mailbox")
End Function
Public Function FindOpenMailWin() As Long
    Dim AOL As Long, MDI As Long, child As Long, richcntl As Long
    Dim childtxt As String, richtxt As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    childtxt$ = GetCaption(child&)
    richtxt$ = GetText(richcntl&)
    If Left(richtxt$, Len(childtxt$) + 6) = "Subj:" & Chr(9) & childtxt$ Then
        FindOpenMailWin& = child&
        Exit Function
    Else
        Do: DoEvents
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            childtxt$ = GetCaption(child&)
            richtxt$ = GetText(richcntl&)
            If Left(richtxt$, Len(childtxt$) + 6) = "Subj:" & Chr(9) & childtxt$ Then
                FindOpenMailWin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
        FindOpenMailWin& = child&
    End If
    
End Function
Public Function FindWelcomeWin() As Long
    Dim AOL As Long, MDI As Long, child As Long, childtxt As String
    AOL& = FindWindow("AOL Frame25", "America  Online")
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    childtxt$ = GetCaption(child&)
    If InStr(childtxt$, "Welcome, ") <> 0& Then
        FindWelcomeWin& = child&
        Exit Function
    Else
        Do: DoEvents
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            childtxt$ = GetCaption(child&)
            If InStr(childtxt$, "Welcome, ") <> 0& Then
                FindWelcomeWin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
End Function

Public Function GetMailDate(TheText As String) As String
    Dim whereat As Long
    whereat& = InStr(TheText$, Chr(9))
    TheText$ = Left(TheText$, whereat& - 1)
    GetMailDate$ = TheText$
End Function
Public Function GetProfile(ScreenName As String) As String
    Dim AOL As Long, MDI As Long, toolbar1 As Long, toolbar2 As Long
    Dim Icon As Long, profilewin As Long, snbox As Long, OKButton As Long
    Dim prowin As Long, profile As Long, protxt1 As String, protxt2 As String
    Dim protxt3 As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    Icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(toolbar2&, Icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(Icon&, 11)
    Do: DoEvents
        profilewin& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        snbox& = FindWindowEx(profilewin&, 0&, "_AOL_Edit", vbNullString)
        OKButton& = FindWindowEx(profilewin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until profilewin& <> 0& And snbox& <> 0& And OKButton& <> 0&
    Call SendMessageByString(snbox&, WM_SETTEXT, 0&, ScreenName$)
    Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        prowin& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        profile& = FindWindowEx(prowin&, 0&, "_AOL_View", vbNullString)
        Do: DoEvents
            protxt1$ = GetText(profile&)
            Pause 0.2
            protxt2$ = GetText(profile&)
            Pause 0.2
            protxt3$ = GetText(profile&)
        Loop Until protxt1$ = protxt2$ And protxt2$ = protxt3$
    Loop Until prowin& <> 0& And profile& <> 0&
    GetProfile$ = GetText(profile&)
    Call SendMessage(profilewin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(prowin&, WM_CLOSE, 0&, 0&)
End Function
Public Sub Ghost(onoff As Boolean)
    Dim AOL As Long, MDI As Long, buddywin As Long
    Dim setupbutton As Long, setupwin As Long, ppbutton1 As Long
    Dim ppbutton As Long, ppwin As Long, blockall1 As Long
    Dim blockalloff As Long, blockallon As Long, blockiandb1 As Long
    Dim blockiandb As Long, savebutton1 As Long, savebutton As Long
    Dim okwin As Long, OKButton As Long, user As String
    user$ = GetUser()
    AOL& = FindWindow("AOL Frame25", "America  Online")
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0 Then
        Call Keyword("buddy list")
    ElseIf buddywin& <> 0 Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(MDI&, 0&, "AOL Child", user$ & "'s Buddy Lists")
        ppbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And ppbutton& <> 0&
    Call PostMessage(ppbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ppbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        ppwin& = FindWindowEx(MDI&, 0&, "AOL Child", "Privacy Preferences")
        blockall1& = FindWindowEx(ppwin&, 0&, "_AOL_Checkbox", vbNullString)
        blockall1& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockalloff& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockall1& = FindWindowEx(ppwin&, blockalloff&, "_AOL_Checkbox", vbNullString)
        blockallon& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockiandb1& = FindWindowEx(ppwin&, blockallon&, "_AOL_Checkbox", vbNullString)
        blockiandb& = FindWindowEx(ppwin&, blockiandb1&, "_AOL_Checkbox", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, 0&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
    Loop Until ppwin& <> 0& And blockallon& <> 0& And blockalloff& <> 0& And blockiandb& <> 0& And savebutton& <> 0&
    If onoff = True Then
        Call PostMessage(blockallon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(blockallon&, WM_LBUTTONUP, 0&, 0&)
    ElseIf onoff = False Then
        Call PostMessage(blockalloff&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(blockalloff&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(blockiandb&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(blockiandb&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
    Loop Until okwin& <> 0& And OKButton& <> 0&
    Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function LinkSender(url As String, Text As String) As String
    htmllink$ = "<a href=" & Chr(34) & url$ & Chr(34) & ">" & Text$ & "</a>"
End Function

Public Function locatemember(Person As String) As String
    Dim AOL As Long, MDI As Long, okwin As Long
    Dim childwin As Long, childwincaption As String
    Dim locatewin As Long, locatemsg1 As Long, locatemsg As String
    Dim OKButton As Long
    AOL& = FindWindow("AOL Frame25", "America  Online")
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://3548:" & Person$)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        If okwin& <> 0& Then Exit Do
        childwin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
        childwincaption$ = GetCaption(childwin&)
        If LCase(childwincaption$) = LCase("locate " & Person$) Then
            locatewin& = childwin&
        Else
            Do: DoEvents
                childwin& = FindWindowEx(MDI&, childwin&, "AOL Child", vbNullString)
                childwincaption$ = GetCaption(childwin&)
                If LCase(childwincaption$) = LCase("locate " & Person$) Then
                    locatewin& = childwin&
                    Exit Do
                End If
                okwin& = FindWindow("#32770", "America Online")
            Loop Until childwin& = 0& Or okwin& <> 0&
        End If
    Loop Until locatewin& <> 0& Or okwin& <> 0&
    If locatewin& <> 0& Then
        locatemsg1& = FindWindowEx(locatewin&, 0&, "_AOL_Static", vbNullString)
        locatemsg$ = GetText(locatemsg1&)
        If LCase(locatemsg$) = LCase(Person$ & " is online, but not in a chat area.") Then
            locatemember$ = "Not in a chat."
        ElseIf LCase(locatemsg$) = LCase(Person$ & " is online, but in a private room.") Then
            locatemember$ = "Private room."
        ElseIf LCase(locatemsg$) Like LCase(Person$ & " is in chat room *") Then
            locatemember$ = Right(locatemsg$, Len(locatemsg$) - Len(Person$ & " is in chat room "))
        End If
        Call SendMessage(locatewin&, WM_CLOSE, 0&, 0&)
    ElseIf okwin& <> 0& Then
        OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Do: DoEvents
            Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
            okwin& = FindWindow("#32770", "America Online")
            OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Loop Until okwin& = 0& And OKButton& = 0&
        locatemember$ = "Not signed on."
    End If
End Function
Public Sub RoomBuster(Room As String)
    Dim AOL As Long, MDI As Long, roomname As String, okwin As Long
    Dim OKButton As Long, chatwin As Long, toolbar1 As Long, toolbar2 As Long
    Dim Combo As Long, EditWin As Long, Modal As Long, Button As Long
    Dim chatwintxt As String, formwin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    roombustcount& = 0&
    roombuststop = False
    roomname$ = GetText(FindRoom())
    roomname$ = removechar(roomname$, " ")
    roomname$ = LCase(roomname$)
    Room$ = removechar(Room$, " ")
    Room$ = LCase(Room$)
    If roomname$ = Room$ Then Exit Sub
    Do: DoEvents
        If roombuststop = True Then Exit Do
        toolbar1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
        toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
        Combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
        EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
        Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, "aol://2719:2-2-" & Room$)
        Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
        okwin& = FindWindow("#32770", "America Online")
        If okwin& <> 0& Then
            roombustcount& = roombustcount& + 1
            OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
            Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
        End If
        chatwin& = FindRoom()
        chatwintxt$ = GetText(chatwin&)
        chatwintxt$ = LCase(chatwintxt$)
        chatwintxt$ = removechar(chatwintxt$, " ")
        formwin& = FindWindowEx(MDI&, 0&, "AOL Child", "Form")
    Loop Until chatwintxt$ = Room$ Or formwin& <> 0&
    Call RunMenu(4, 10)
    Do: DoEvents
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        Button& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until Modal& <> 0& And Button& <> 0&
    Do: DoEvents
        Call PostMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(Button&, WM_LBUTTONUP, 0&, 0&)
        Button& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until Button& = 0&
End Sub
Public Sub MailSend(Person As String, subject As String, message As String)
    Dim AOL As Long, MDI As Long, toolbar1 As Long, toolbar2 As Long
    Dim writeicon As Long, mailwin As Long, personbox As Long, ccbox As Long
    Dim subjectbox As Long, MessageBox As Long, SendButton As Long, sendbutton1 As Long
    Dim sendbutton2 As Long
    AOL& = FindWindow("AOL Frame25", "America  Online")
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, writeicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(writeicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(writeicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        mailwin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        personbox& = FindWindowEx(mailwin&, 0&, "_AOL_Edit", vbNullString)
        ccbox& = FindWindowEx(mailwin&, personbox&, "_AOL_Edit", vbNullString)
        subjectbox& = FindWindowEx(mailwin&, ccbox&, "_AOL_Edit", vbNullString)
        MessageBox& = FindWindowEx(mailwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton2& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(mailwin&, sendbutton2&, "_AOL_Icon", vbNullString)
    Loop Until mailwin& <> 0& And personbox& <> 0& And ccbox& <> 0& And subjectbox& <> 0& And MessageBox& <> 0& And SendButton& <> 0& And SendButton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
    Call SendMessageByString(personbox&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
    Call SendMessageByString(MessageBox&, WM_SETTEXT, 0&, message$)
    Do: DoEvents
        Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        mailwin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Loop Until mailwin& = 0&
End Sub
Public Sub SignOffAOL()
    Call RunMenu(3&, 1&)
End Sub
Public Sub WaitForListToLoad(list As Long)
   
    Dim getcount1 As Long, getcount2 As Long, getcount3 As Long
    Do
        getcount1& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        TimePause (0.8)
        getcount2& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        TimePause (0.8)
        getcount3& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
    Loop Until getcount1& = getcount2& And getcount2& = getcount3&
End Sub
Public Sub WindowHide(win As Long, onoff As Boolean)
    If onoff = True Then
        Call ShowWindow(win&, SW_HIDE)
    ElseIf onoff = False Then
        Call ShowWindow(win&, SW_SHOW)
    End If
End Sub
Sub Welcome_Kill()
' It took me two days to think of this code
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
Call SendMessage(welcome%, WM_CLOSE, 0, 0)
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(welcome%, 0)
End Sub
Sub Welcome_UnKill()
' It took me two days to think of this code
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
Call SendMessage(welcome%, WM_CLOSE, 0, 0)
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(welcome%, 5)
End Sub

Public Sub FormMoveable(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub WavPlayer(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Sub MidiPlayer(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Sub Lit3MacroKill1()
SendChatMsg ("<font color=#00ee00>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#00ff00>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
SendChatMsg ("<font color=#996600>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#00ff00>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#cccccc>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
SendChatMsg ("<font color=#ffffff>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#00cc00>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
SendChatMsg ("<font color=#eeffcc>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#111222>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
SendChatMsg ("<font color=#115533>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
SendChatMsg ("<font color=#11ee44>.<Pre=@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimePause (2)
End Sub
Sub Lit3MacroKill2()
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
SendChatMsg ("<font color=#00ee00>.<Pre=%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
TimePause (2)
End Sub
Sub Lit3Macro()
SendChatMsg ("       ,,,,       ,,,              ,,;;;;;,                         ,,,,,,,,,         ,,")
TimePause (1.5)
SendChatMsg ("      ´;;;;;;;;;;;;´´             ;;;;;;;;;    ,, ,,,,,;;´       ,;;;;;;;;;;;;;;;;;;;;;;")
TimePause (1.5)
SendChatMsg ("        ´;;;;;;;;;;                ;;;;;;;;    ´;;;;;;;;;       ,;;;;;;;;;;;;;;;;;;;;;´")
TimePause (1.5)
SendChatMsg ("          ;;;;;;;;;                 ´;;;;´´      ;;;;;;;;        ;;;;;;  ´´´´´;,;;;;´")
TimePause (1.5)
SendChatMsg ("         ;;;;;;;;;               ;;;,,,,,;;´     ;;;;;;;,,,,;;  ;;;;;;;;   ,,;;;;´")
TimePause (1.5)
SendChatMsg ("         ;;;;;;;;                 ;;;;;;;;    ,,;;;;;;;;;;;;;´   ;;;;´´  ,;;;;;,")
TimePause (1.5)
SendChatMsg ("        ;;;;;;;;;  ,;;;;;;;;,    ;;;;;;;;   ´´´´;;;;;;;;;;;;´         ,;;;;;;;;;;;;;,,")
TimePause (1.5)
SendChatMsg ("       ,;;;;;;;; ,;;;;;;;;;;;    ;;;;;;;       ;;;;;;;                 ´´´  ´;;;;;;;;;;")
TimePause (1.5)
SendChatMsg ("       ;;;;;;;; ,;;;;;;;;;;;;   ,;;;;;;       ;;;;;;;´         ,,;;;;;;,     ;;;;;;;;;;")
TimePause (1.5)
SendChatMsg ("      ;;;;;;;;;  ´´´;;;;;;;;;  ,;;;;;;;      ,;;;;;;;,;;;;;, ,;;;;;;;;       ;;;;;;;;;;")
TimePause (1.5)
SendChatMsg ("     ;;;;;;;;;       ;;;;;;;   ;;;;;;;      ,;;;;;;;;;;;;;;; ;;;;;;;;       ;;;;;;;;;;")
TimePause (1.5)
SendChatMsg ("   ,;;;;;;;;;;      ;;;;;;´   ;;;;;;;´    ,;;;;;;;;´´´;;;;;´´;;;;;;;;     ,;;;;;;;;;´")
TimePause (1.5)
SendChatMsg (" ,;;;;;;;;;;;;;,,,;;;;;;´   ,;;;;;;;;,     ´;;;;;;;  ;;;;;´  ;;;;;;;;;,,,;;;;;;;;´´")
TimePause (1.5)
SendChatMsg (" ;;´´    ´´´;;;;;;;;´´     ;;´   ´´´;;       ´;;;;;,;;´´´      ´´;;;;;;;;;;;´´´")
End Sub
Sub Credits()
'Programmer
'Lit3
'__________
' Doyoulay
' Best Visual Basic site there is KnK,.;`~^Http://WWW.KnK2000.Com/KnK/^~;.,
'KnK's Site has Helped Me Alot With Programming,so visit and it will most likly help you to
'Besure to visit my site as well ,.;`~^`';..;`..->Http://WWW.AngelFire.Com/ab/Necromancy/index.html
End Sub
