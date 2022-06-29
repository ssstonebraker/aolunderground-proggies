Attribute VB_Name = "unique32"
'unique32¹ by unique (aol_enterprise@hotmail.com)
'contribution: infinyti, sci, slik
'creditz: dos and sonic, PAA fer utility
'greetz: team express, mike, nbc, cbs
'note: this is my last bas i'll release. im outta
'aol since it ain't fittin ma style.copy watever
'you want,just gimme a credit(still gatta be known)
'im ganna make stuf out of yahoo massenger now
'(i b the first one ;x )
Option Explicit
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MCISendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const LB_ADDSTRING& = &H180
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const GW_MAX = 5
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Public Const PROCESS_READ = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
Public Const CB_RESETCONTENT = &H14B
Public Type POINTAPI
X As Long
Y As Long
End Type
Sub other_textcolor(Textbox As Textbox, Red As Long, Green As Long, Blue As Long)
Textbox.ForeColor = RGB(Red&, Green&, Blue&)
Call other_delay(0.2)
Textbox.ForeColor = RGB(Red& * 3 / 4, Green& * 3 / 4, Blue& * 3 / 4)
Call other_delay(0.2)
Textbox.ForeColor = RGB(Red& / 2, Green& / 2, Blue& / 2)
Call other_delay(0.2)
Textbox.ForeColor = RGB(Red& / 4, Green& / 4, Blue& / 4)
Call other_delay(0.2)
Textbox.ForeColor = RGB(0, 0, 0)
End Sub


Sub aol_addlisttobuddy(Lst As ListBox)
Dim EditBud As Long, AddIt As Long, AOEdit As Long
Dim Who As String, AOL As Long, AOMDI As Long, Bud As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOMDI& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
Call aol_keyword("buddylist")
Bud& = FindWindowEx(AOMDI&, 0, "AOL Child", aol_getuser + "'s Buddy Lists")
EditBud& = Bud&
If EditBud& = 0 Then Exit Sub
AOEdit& = FindWindowEx(EditBud&, 0&, "_AOL_Edit", vbNullString)
AOEdit& = FindWindowEx(EditBud&, AOEdit&, "_AOL_Edit", vbNullString)
For AddIt& = 0 To Lst.listcount - 1
Who$ = Lst.List(AddIt&)
Call SendMessageByString(AOEdit&, WM_SETTEXT, 0&, Who)
Call SendMessageLong(AOEdit&, WM_CHAR, ENTER_KEY, 0&)
Call other_delay(0.6)
Next AddIt&
End Sub
Sub aol_addbuddylist(Who As String)
Dim EditBuddyWin As Long, AOL As Long, AOMDI As Long
Dim Bud1 As Long, Bud2 As Long, EditBut As Long
Dim PutSnWin As Long, PutSnCombo As Long, PutSnStatic As Long
Dim PutSnEdit As Long, PutSnGlyph As Long, PutSnIcon As Long
Dim PutSnList As Long, DaSnWin As Long, SetSn As Long
Dim SaveBud As Long, HitDaOk As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOMDI& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
Call aol_keyword("buddylist")
Do
DoEvents
Bud1& = FindWindowEx(AOMDI&, 0, "AOL Child", aol_getuser() + "'s Buddy Lists")
Bud2& = FindWindowEx(AOMDI&, 0, "AOL Child", aol_getuser() + "'s Buddy List")
If Bud1& <> 0 Then
EditBuddyWin& = Bud1&
Exit Do
End If
If Bud2& <> 0 Then
EditBuddyWin& = Bud2&
Exit Do
End If
Loop
DoEvents
Do
DoEvents
EditBut& = FindWindowEx(EditBuddyWin&, 0, "_AOL_Icon", vbNullString)
EditBut& = FindWindowEx(EditBuddyWin&, EditBut&, "_AOL_Icon", vbNullString)
If EditBut& Then Exit Do
Loop
DoEvents
Call PostMessage(EditBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(EditBut&, WM_LBUTTONUP, 0, 0)
DoEvents
Do
DoEvents
PutSnWin& = FindWindowEx(AOMDI&, 0, "AOL Child", vbNullString)
PutSnList& = FindWindowEx(PutSnWin&, 0, "_AOL_Listbox", vbNullString)
PutSnCombo& = FindWindowEx(PutSnWin&, 0, "_AOL_Combobox", vbNullString)
PutSnGlyph& = FindWindowEx(PutSnWin&, 0, "_AOL_Glyph", vbNullString)
PutSnIcon& = FindWindowEx(PutSnWin&, 0, "_AOL_Icon", vbNullString)
PutSnStatic& = FindWindowEx(PutSnWin&, 0, "_AOL_Static", vbNullString)
PutSnEdit& = FindWindowEx(PutSnWin&, 0, "_AOL_Edit", vbNullString)
If PutSnCombo& And PutSnEdit& Then
DaSnWin& = PutSnWin&
Exit Do
End If
Loop
Call other_delay(0.5)
SetSn& = FindWindowEx(DaSnWin&, 0, "_Aol_Edit", vbNullString)
SetSn& = FindWindowEx(DaSnWin&, SetSn&, "_Aol_Edit", vbNullString)
Call SendMessageByString(SetSn&, WM_SETTEXT, 0&, Who)
DoEvents
Call SendMessageByString(SetSn&, WM_CHAR, ENTER_KEY, 0)
Call other_delay(0.5)
SaveBud& = FindWindowEx(DaSnWin&, 0, "_AOL_Icon", vbNullString)
SaveBud& = FindWindowEx(DaSnWin&, SaveBud&, "_AOL_Icon", vbNullString)
SaveBud& = FindWindowEx(DaSnWin&, SaveBud&, "_AOL_Icon", vbNullString)
Call PostMessage(SaveBud&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(SaveBud&, WM_LBUTTONUP, 0, 0)
For HitDaOk& = 1 To 20
Call other_delay(0.1)
Call aol_waitforok
Next HitDaOk
DoEvents
Call api_closewindow(EditBuddyWin&)
End Sub



Sub aol_addressadd(fName As String, lName As String, eMail As String, Note As String)
Dim MainParent As Long, Parent1 As Long, Parent2 As Long
Dim Parent3 As Long, Parent4 As Long, TheWindow As Long
Dim AOL As Long, MDI As Long, Tool As Long, WinEdit As Long
Dim ToolIcon As Long, WinVis As Long, sMod As Long
Dim Toolbar As Long, CurPos As POINTAPI, DoThis As Long
Dim MailIcon As Long, PrefWin As Long, MailWin As Long
Dim SetConf As Long, SetCloseMail As Long, PrefOk As Long
Dim TheWindow2 As Long, TheWindow3 As Long, TheWindow4 As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
DoEvents
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
DoEvents
Loop Until WinVis& = 1
DoEvents
For DoThis& = 1 To 6
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
DoEvents
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
DoEvents
Do
DoEvents
PrefWin& = FindWindowEx(MDI&, 0, "AOL Child", "Address Book")
Loop Until PrefWin&
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
TheWindow& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
WinEdit& = FindWindowEx(MDI&, 0, "AOL Child", "New Person")
Loop Until WinEdit&
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Parent3& = FindWindowEx(Parent2&, 0&, "_AOL_TabControl", vbNullString)
Parent4& = FindWindowEx(Parent3&, 0&, "_AOL_TabPage", vbNullString)
TheWindow& = FindWindowEx(Parent4&, 0&, "_AOL_Edit", vbNullString)
TheWindow2& = FindWindowEx(Parent4&, TheWindow&, "_AOL_Edit", vbNullString)
TheWindow3& = FindWindowEx(Parent4&, TheWindow2&, "_AOL_Edit", vbNullString)
TheWindow4& = FindWindowEx(Parent4&, TheWindow3&, "_AOL_Edit", vbNullString)
Call SendMessageByString(TheWindow&, &HC, 0&, fName$)
Call SendMessageByString(TheWindow2&, &HC, 0&, lName$)
Call SendMessageByString(TheWindow3&, &HC, 0&, eMail$)
Call SendMessageByString(TheWindow4&, &HC, 0&, Note$)
Call other_delay(0.5)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
TheWindow& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
Do
DoEvents
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
Loop Until TheWindow& <> 0
End Sub


Sub aol_addressremove(Index As Long)
Dim Toolbar As Long, CurPos As POINTAPI, AOL As Long, MDI As Long
Dim Tool As Long, ToolIcon As Long, sMod As Long, WinVis As Long
Dim DoThis As Long, PrefWin As Long, fList As Long, TheWindow As Long
Dim Child1 As Long, Child2 As Long, Child3 As Long, fCount As Long
Dim Parent1 As Long, Parent2 As Long, TheParent As Long, MainParent As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
DoEvents
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
DoEvents
Loop Until WinVis& = 1
DoEvents
For DoThis& = 1 To 6
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
DoEvents
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
DoEvents
Do
DoEvents
PrefWin& = FindWindowEx(MDI&, 0, "AOL Child", "Address Book")
Loop Until PrefWin&
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
Child2& = FindWindowEx(Parent2&, Child1&, "_AOL_Icon", vbNullString)
Child3& = FindWindowEx(Parent2&, Child2&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(Parent2&, Child3&, "_AOL_Icon", vbNullString)
fList& = FindWindowEx(Parent2&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < Index& Then Exit Sub
Call PostMessage(fList&, &H201, 0&, 0&)
Call PostMessage(fList&, &H202, 0&, 0&)
Call SendMessageLong(TheWindow&, &H186, Index&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
Call other_delay(0.6)
TheParent& = FindWindow("#32770", vbNullString)
TheWindow& = FindWindowEx(TheParent&, 0&, "Button", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub aol_changedownloaddir(eText As String)
Dim MainParent As Long, Child1 As Long, TheWindow As Long, TheParent As Long
Dim Parent1 As Long, Parent2 As Long, Child As Long
Dim AOL As Long, MDI As Long, Tool As Long, ToolIcon As Long, WinVis As Long, sMod As Long
Dim Toolbar As Long, CurPos As POINTAPI, DoThis As Long
Dim MailIcon As Long, PrefWin As Long, MailWin As Long
Dim SetConf As Long, SetCloseMail As Long, PrefOk As Long, X As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
DoEvents
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
DoEvents
Loop Until WinVis& = 1
DoEvents
For DoThis& = 1 To 3
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
DoEvents
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
DoEvents
Do
DoEvents
PrefWin& = FindWindowEx(MDI&, 0, "AOL Child", "Preferences")
Loop Until PrefWin&
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Child& = FindWindowEx(Parent2&, 0, "_AOL_Icon", vbNullString)
For X& = 1 To 5
Child& = FindWindowEx(Parent2&, Child&, "_AOL_Icon", vbNullString)
Next X&
TheWindow& = Child&
Do
DoEvents
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
Loop Until TheWindow& <> 0
Call other_delay(0.6)
TheParent& = FindWindow("_AOL_Modal", vbNullString)
TheWindow& = FindWindowEx(TheParent&, 0&, "_AOL_Edit", vbNullString)
SendMessageByString TheWindow&, &HC, 0&, eText$
Child1& = FindWindowEx(TheParent&, 0&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(TheParent&, Child1&, "_AOL_Icon", vbNullString)
Call PostMessage(TheWindow&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(TheWindow&, WM_KEYUP, VK_RETURN, 0&)
Call PostMessage(PrefWin&, WM_CLOSE, 0&, 0&)
End Sub


Sub aol_favoriteadd(Description As String, URL As String)
Dim AOL As Long, MDI As Long, Tool As Long, MainParent As Long
Dim ToolIcon As Long, WinVis As Long, sMod As Long
Dim Toolbar As Long, CurPos As POINTAPI, DoThis As Long
Dim MailIcon As Long, PrefWin As Long, MailWin As Long
Dim SetConf As Long, SetCloseMail As Long, PrefOk As Long
Dim Parent1 As Long, Parent2 As Long, Child1 As Long
Dim TheWindow As Long, TheWindow2 As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
DoEvents
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
DoEvents
Loop Until WinVis& = 1
DoEvents
For DoThis& = 1 To 1
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
DoEvents
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(Parent2&, Child1&, "_AOL_Icon", vbNullString)
Do
DoEvents
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
Loop Until TheWindow& <> 0
Call other_delay(0.6)
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(Parent2&, 0&, "_AOL_Edit", vbNullString)
TheWindow& = FindWindowEx(Parent2&, Child1&, "_AOL_Edit", vbNullString)
TheWindow2& = FindWindowEx(Parent2&, TheWindow&, "_AOL_Edit", vbNullString)
SendMessageByString TheWindow&, &HC, 0&, Description$
SendMessageByString TheWindow2&, &HC, 0&, URL$
Call other_delay(0.6)
Child1& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(Parent2&, Child1&, "_AOL_Icon", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
End Sub


Function api_cbcount(TheWindow As Long) As Long
api_cbcount& = SendMessageLong(TheWindow&, &H146, 0&, 0&)
End Function

Sub api_cbselect(TheWindow As Long, Index As Long)
Call SendMessageLong(TheWindow&, &H14E, Index&, 0&)
End Sub


Function api_lbgettext(TheWindow As Long, Index As Long) As String
Dim TheText As String
TheText$ = String(255, 0)
Call SendMessageByString(TheWindow&, &H189, Index&, TheText$)
api_lbgettext$ = TheText$
End Function

Sub api_lbselect(TheWindow As Long, Index As Long)
Call SendMessageLong(TheWindow&, &H186, Index&, 0&)
End Sub

Function form_lasershow(Frm As Form)
Frm.Line (Int(Rnd * Frm.Width), Int(Rnd * Frm.Height))-(Int(Rnd * Frm.Width), Int(Rnd * Frm.Height)), QBColor(Rnd * 15)
End Function

Sub aol_normalmailpref()
Dim AOL As Long, MDI As Long, Tool As Long
Dim ToolIcon As Long, WinVis As Long, sMod As Long
Dim Toolbar As Long, CurPos As POINTAPI, DoThis As Long
Dim MailIcon As Long, PrefWin As Long, MailWin As Long
Dim SetConf As Long, SetCloseMail As Long, PrefOk As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
DoEvents
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
DoEvents
Loop Until WinVis& = 1
DoEvents
For DoThis& = 1 To 3
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
DoEvents
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
DoEvents
Do
DoEvents
PrefWin& = FindWindowEx(MDI&, 0, "AOL Child", "Preferences")
Loop Until PrefWin&
Call other_delay(0.6)
Do
DoEvents
MailIcon& = FindWindowEx(PrefWin&, 0, "_AOL_Icon", vbNullString)
MailIcon& = FindWindowEx(PrefWin&, MailIcon&, "_AOL_Icon", vbNullString)
MailIcon& = FindWindowEx(PrefWin&, MailIcon&, "_AOL_Icon", vbNullString)
Loop Until MailIcon& <> 0
DoEvents
Call PostMessage(MailIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(MailIcon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
MailWin& = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until MailWin&
Call other_delay(0.6)
SetConf& = FindWindowEx(MailWin&, 0, "_AOL_Checkbox", vbNullString)
SetCloseMail& = FindWindowEx(MailWin&, SetConf&, "_AOL_Checkbox", vbNullString)
Call PostMessage(SetConf&, BM_SETCHECK, False, vbNullString)
Call PostMessage(SetCloseMail&, BM_SETCHECK, True, vbNullString)
PrefOk& = FindWindowEx(MailWin&, 0, "_AOL_Icon", vbNullString)
Do
DoEvents
Loop Until PrefOk&
DoEvents
Call PostMessage(PrefOk&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefOk&, WM_LBUTTONUP, 0&, 0&)
Call other_delay(1)
Call api_closewindow(PrefWin&)
End Sub


Sub aol_switchsn(Index As Long, password As String)
Dim AOL As Long, MDI As Long, switchwin As Long, snlist As Long
Dim switchbutton As Long, listcount As Long
Dim sThread As Long, mThread As Long, thesn As String, itmHold As Long
Dim psnHold As Long, rBytes As Long, swin2 As Long, sok2 As Long
Dim swin As Long, sok As Long, spw As Long, cProcess As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call aol_runmenu(3&, 0&)
Do: DoEvents
switchwin& = FindWindowEx(MDI&, 0&, "AOL Child", "Switch Screen Names")
snlist& = FindWindowEx(switchwin&, 0&, "_AOL_Listbox", vbNullString)
switchbutton& = FindWindowEx(switchwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until switchwin& <> 0& And snlist& <> 0& And switchbutton& <> 0&
listcount& = PostMessage(snlist&, LB_GETCOUNT, 0&, 0&)
Call PostMessage(snlist&, LB_SETCURSEL, CLng(Index&), 0&)
Call PostMessage(switchbutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(switchbutton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
swin2& = FindWindow("_AOL_Modal", "Switch Screen Name")
sok2& = FindWindowEx(swin2&, 0&, "_AOL_Icon", vbNullString)
Loop Until swin2& <> 0& And sok2& <> 0&
Call PostMessage(sok2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(sok2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
swin& = FindWindow("_AOL_Modal", "Switch Screen Name")
spw& = FindWindowEx(swin&, 0&, "_AOL_Edit", vbNullString)
sok& = FindWindowEx(swin&, 0&, "_AOL_Icon", vbNullString)
Loop Until swin& <> 0& And spw& <> 0& And sok& <> 0&
Call SendMessageByString(spw&, WM_SETTEXT, 0&, password$)
Call PostMessage(sok&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(sok&, WM_LBUTTONUP, 0&, 0&)
End Sub


Function form_nonaol(Frm As Form)
Call SetParent(Frm.hWnd, 0)
End Function



Sub list_additem(List As Control, Item As String)
Dim X As Long
On Error Resume Next
DoEvents
For X& = 0 To List.listcount - 1
If UCase$(List.List(X&)) = UCase$(Item$) Then Exit Sub
Next
If Len(Item$) <> 0 Then List.AddItem Item$
End Sub



Sub form_shrink(Frm As Form)
On Error Resume Next
Dim GotoVal As Long, GoInTo As Long
GotoVal = (Frm.Height / 60)
For GoInTo = 1 To GotoVal
DoEvents
Frm.Height = Frm.Height - 60
Frm.Top = (Screen.Height - Frm.Height) \ 2
DoEvents
Frm.Width = Frm.Width - 120
Frm.Left = (Screen.Width - Frm.Width) \ 2
If Frm.Width <= 100 Then Exit Sub
If Frm.Height <= 70 Then Exit Sub
DoEvents
Next GoInTo
End Sub

Function api_lbcount(List As Long) As Long
api_lbcount& = SendMessageByNum(List&, LB_GETCOUNT, 0, 0)
End Function



Function other_countchar(thestring As String, thechar As String) As Long
Dim Whereat As Long
Whereat& = 0&
Do
Whereat& = InStr(Whereat& + 1, thestring$, thechar$)
If Whereat& = 0& Then Exit Do
other_countchar& = other_countchar& + 1
Loop
End Function

Sub aol_mailattachment(Person As String, subject As String, message As String, filepath As String)
Dim SlashCount As Long, Indexx As Long, Whereat As Long, WhereAttemp As Long
Dim WherEat1 As Long, Folders As String, Folder As String, File As String
Dim AOL As Long, MDI As Long, ToolBar1 As Long, ToolBar2 As Long, WriteIcon As Long
Dim WriteWin As Long, SendToBox As Long, SubjectBox1 As Long, SubjectBox As Long
Dim MessageBox As Long, Attach1 As Long, SendButton1 As Long, SendButton2 As Long
Dim SendButton As Long, attachwin As Long, attachbutton As Long, okbutton1 As Long
Dim OKButton As Long, OpenWin As Long, FileBox As Long, Combo As Long
Dim OpenButton1 As Long, OpenButton As Long
SlashCount& = other_countchar(filepath$, "\")
For Indexx& = 0& To SlashCount& - 1
Whereat& = InStr(Whereat& + 1, filepath$, "\")
Next Indexx&
Do: DoEvents
WhereAttemp& = InStr(WhereAttemp& + 1, filepath$, "\")
If WhereAttemp& = Whereat& Then Exit Do
WherEat1& = WhereAttemp&
Loop
Folders$ = Left(filepath$, Whereat&)
Folder$ = Mid(filepath$, WherEat1& + 1, Whereat& - WherEat1& - 1)
File$ = Right(filepath$, Len(filepath$) - Whereat&)
If Len(Dir(filepath$)) Then
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
ToolBar1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolBar2& = FindWindowEx(ToolBar1&, 0&, "_AOL_Toolbar", vbNullString)
WriteIcon& = FindWindowEx(ToolBar2&, 0&, "_AOL_Icon", vbNullString)
WriteIcon& = FindWindowEx(ToolBar2&, WriteIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(WriteIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(WriteIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
WriteWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
SendToBox& = FindWindowEx(WriteWin&, 0&, "_AOL_Edit", vbNullString)
SubjectBox1& = FindWindowEx(WriteWin&, SendToBox&, "_AOL_Edit", vbNullString)
SubjectBox& = FindWindowEx(WriteWin&, SubjectBox1&, "_AOL_Edit", vbNullString)
MessageBox& = FindWindowEx(WriteWin&, 0&, "RICHCNTL", vbNullString)
Attach1& = FindWindowEx(WriteWin&, 0&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
Attach1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
SendButton1& = FindWindowEx(WriteWin&, Attach1&, "_AOL_Icon", vbNullString)
SendButton2& = FindWindowEx(WriteWin&, SendButton1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(WriteWin&, SendButton2&, "_AOL_Icon", vbNullString)
Loop Until WriteWin& <> 0& And SendToBox& <> 0& And SubjectBox1& <> 0& And SubjectBox& <> 0& And MessageBox& <> 0& And SendButton& <> 0& And SendButton& <> SendButton1& And SendButton1& <> SendButton2& And SendButton2& <> 0& And SendButton1 <> 0&
Call SendMessageByString(SendToBox&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(SubjectBox&, WM_SETTEXT, 0&, subject$)
Call SendMessageByString(MessageBox&, WM_SETTEXT, 0&, message$)
Call PostMessage(SendButton2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
attachwin& = FindWindow("_AOL_Modal", "Attachments")
attachbutton& = FindWindowEx(attachwin&, 0&, "_AOL_Icon", vbNullString)
okbutton1& = FindWindowEx(attachwin&, attachbutton&, "_AOL_Icon", vbNullString)
OKButton& = FindWindowEx(attachwin&, okbutton1&, "_AOL_Icon", vbNullString)
Loop Until attachwin& <> 0& And attachbutton& <> 0& And OKButton& <> 0&
Call PostMessage(attachbutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(attachbutton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
OpenWin& = FindWindow("#32770", "Attach")
FileBox& = FindWindowEx(OpenWin&, 0&, "Edit", vbNullString)
Combo& = FindWindowEx(OpenWin&, 0&, "ComboBox", vbNullString)
OpenButton1& = FindWindowEx(OpenWin&, 0&, "Button", vbNullString)
OpenButton& = FindWindowEx(OpenWin&, OpenButton1&, "Button", vbNullString)
Loop Until OpenWin& <> 0& And FileBox& <> 0& And Combo& <> 0& And OpenButton& <> 0&
Call SendMessageByString(FileBox&, WM_SETTEXT, 0&, Folders$)
Do: DoEvents
Call SendMessageByString(FileBox&, WM_SETTEXT, 0&, Folders$)
Call PostMessage(OpenButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OpenButton&, WM_LBUTTONUP, 0&, 0&)
Loop Until LCase(api_gettext(Combo&)) = LCase(Folder$)
Call SendMessageByString(FileBox&, WM_SETTEXT, 0&, File$)
Do: DoEvents
Call PostMessage(OpenButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OpenButton&, WM_LBUTTONUP, 0&, 0&)
OpenWin& = FindWindow("#32770", "Browse")
Loop Until OpenWin& = 0&
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Else
Exit Sub
End If
End Sub



Sub form_hideall()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = False
Next Frm
End Sub

Sub form_hidecontrols(Frm As Form)
On Error Resume Next
Dim Crl As Control
For Each Crl In Frm.Controls
Crl.Visible = False
Next Crl
End Sub



Sub form_introunload(Frm As Form, Ender As Boolean)
Dim Z As Long, X As Long, E As Long, D As Long
For Z& = 0 To (Screen.Width) / 2 Step 10
Frm.Left = Frm.Left - Z&
If Frm.Left <= 0 Then Exit For
Next
For X& = 0 To (Screen.Height) / 2 Step 10
Frm.Top = Frm.Top - X&
If Frm.Top <= 0 Then Exit For
Next
For E& = 0 To (Screen.Width) / 2 Step 10
Frm.Left = Frm.Left + E&
If Frm.Left >= (Screen.Width - Frm.Width) / 2 Then Exit For
Next
For D& = 0 To (Screen.Height) / 2 Step 10
Frm.Top = Frm.Top + D&
If Frm.Top >= (Screen.Height - Frm.Height) / 2 Then Exit For
Next
Do
DoEvents
Frm.Top = Trim(Str(Int(Frm.Top) + 300))
Loop Until Frm.Top > 7200
If Ender = True Then End Else
Exit Sub
End Sub

Sub form_showall()
Dim Frm As Form
For Each Frm In Forms
Frm.Visible = True
Next Frm
End Sub

Sub form_showcontrols(Frm As Form)
On Error Resume Next
Dim Crl As Control
For Each Crl In Frm.Controls
Crl.Visible = True
Next Crl
End Sub




Function file_freegdi() As String
file_freegdi$ = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) & "%"
End Function



Function file_freesys() As String
file_freesys$ = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) & "%"
End Function



Function file_freeuser() As String
file_freeuser$ = Format$(GetFreeSystemResources(GFSR_USERRESOURCES)) + "%"
End Function



Sub aol_buddyinvite(buddies As String, tosay As String, Room As String, gotochat As Boolean)
Dim AOL As Long, MDI As Long, BuddyWin As Long, InviteIcon As Long, InviteWin As Long
Dim Peoplebox As Long, ToSayBox As Long, RoomBox As Long, SendIcon As Long
Dim ItationWin As Long, GoIcon As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
BuddyWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Window")
If BuddyWin& = 0& Then
Call aol_keyword("buddy chat")
ElseIf BuddyWin& <> 0& Then
InviteIcon& = FindWindowEx(BuddyWin&, 0&, "_AOL_Icon", vbNullString)
InviteIcon& = FindWindowEx(BuddyWin&, InviteIcon&, "_AOL_Icon", vbNullString)
InviteIcon& = FindWindowEx(BuddyWin&, InviteIcon&, "_AOL_Icon", vbNullString)
InviteIcon& = FindWindowEx(BuddyWin&, InviteIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(InviteIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(InviteIcon&, WM_LBUTTONUP, 0&, 0&)
End If
Do: DoEvents
InviteWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy Chat")
Peoplebox& = FindWindowEx(InviteWin&, 0&, "_AOL_Edit", vbNullString)
ToSayBox& = FindWindowEx(InviteWin&, Peoplebox&, "_AOL_Edit", vbNullString)
RoomBox& = FindWindowEx(InviteWin&, ToSayBox&, "_AOL_Edit", vbNullString)
SendIcon& = FindWindowEx(InviteWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until InviteWin& <> 0& And Peoplebox& <> 0& And ToSayBox& <> 0& And RoomBox& <> 0& And SendIcon& <> 0&
Call SendMessageByString(Peoplebox&, WM_SETTEXT, 0&, buddies$)
Call SendMessageByString(ToSayBox&, WM_SETTEXT, 0&, tosay$)
Call SendMessageByString(RoomBox&, WM_SETTEXT, 0&, Room$)
Call PostMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
ItationWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Invitation from: " & aol_getuser())
GoIcon& = FindWindowEx(ItationWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until ItationWin& <> 0& And GoIcon& <> 0&
If gotochat = True Then
Call PostMessage(GoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(GoIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf gotochat = False Then
Call PostMessage(ItationWin&, WM_CLOSE, 0&, 0&)
End If
End Sub


Sub aol_chatsend(Text As String)
Dim Room As Long, AORich As Long, AORich2 As Long, TextBefore As String
Room& = aol_findroom&
AORich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
AORich2& = FindWindowEx(Room&, AORich, "RICHCNTL", vbNullString)
TextBefore$ = api_gettext(AORich2&)
Call SendMessageByString(AORich2&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AORich2&, WM_SETTEXT, 0&, Text$)
Do: DoEvents
Call SendMessageLong(AORich2&, WM_CHAR, ENTER_KEY, 0&)
If api_gettext(AORich2&) = "" Then Exit Do
Loop
Call SendMessageByString(AORich2&, WM_SETTEXT, 0&, TextBefore$)
End Sub


Function aol_chatcaption() As String
If aol_findroom& = False Then
Exit Function
aol_chatcaption$ = api_getcaption(aol_findroom&)
End If
End Function

Sub aol_chatclose()
If aol_findroom& = False Then
Exit Sub
Call PostMessage(aol_findroom&, WM_CLOSE, 0&, 0&)
End If
End Sub



Function aol_connectionlog() As Long
Dim AOLWindow As Long, MDIClient As Long, AOLChild As Long
Dim Child1 As Long
AOLWindow& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
If Child1& <> 0& Then
aol_connectionlog& = AOLChild&
Exit Function
Else
While AOLChild&
AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
Child1& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
If Child1& <> 0& Then
aol_connectionlog& = AOLChild&
Exit Function
End If
Wend
End If
End Function
Sub aol_cancelsignon()
Dim MainParent As Long, Child1 As Long, TheWindow As Long
MainParent& = FindWindow("_AOL_Modal", vbNullString)
Child1& = FindWindowEx(MainParent&, 0&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(MainParent&, Child1&, "_AOL_Icon", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
End Sub


Sub aol_clicksingon()
Dim MainParent As Long, Parent1 As Long, Parent2 As Long
Dim Child1 As Long, Child2 As Long, Child3 As Long, TheWindow As Long
Dim Checks As Long, PrefIcon As Long
Checks& = aol_signon&
If Checks& = 0 Then Exit Sub
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(Parent2&, 0&, "_AOL_Icon", vbNullString)
Child2& = FindWindowEx(Parent2&, Child1&, "_AOL_Icon", vbNullString)
Child3& = FindWindowEx(Parent2&, Child2&, "_AOL_Icon", vbNullString)
TheWindow& = FindWindowEx(Parent2&, Child3&, "_AOL_Icon", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
End Sub



Sub aol_killlog()
Dim Connection As Long
Connection& = aol_connectionlog&
If Connection& = 0 Then Exit Sub
Call PostMessage(Connection&, WM_CLOSE, 0&, 0&)
End Sub


Sub aol_setguest()
Dim MainParent As Long, Parent1 As Long, Parent2 As Long
Dim TheWindow As Long, Combo As Long
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "MDIClient", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "AOL Child", vbNullString)
TheWindow& = FindWindowEx(Parent2&, 0&, "_AOL_Combobox", vbNullString)
Dim TheCount As Long
TheCount& = SendMessageLong(TheWindow&, &H146, 0&, 0&)
Call SendMessageLong(TheWindow&, &H14E, 0, 0&)
Combo& = TheCount& - 1
Call SendMessageLong(TheWindow&, &H14E, Combo&, 0&)
End Sub

Function aol_signon() As Long
Dim AOLWindow As Long, MDIClient As Long, AOLChild As Long, Child1 As Long
Dim Child2 As Long, Child3 As Long, Child4 As Long
AOLWindow& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
Child1& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
Child2& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
Child3& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Child4& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
If Child1& <> 0& And Child2& <> 0& And Child3& <> 0& And Child4& <> 0& Then
aol_signon& = AOLChild&
Exit Function
Else
While AOLChild&
AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
Child1& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
Child2& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
Child3& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Child4& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
If Child1& <> 0& And Child2& <> 0& And Child3& <> 0& And Child4& <> 0& Then
aol_signon& = AOLChild&
Exit Function
End If
Wend
End If
End Function
Function aol_checkifmaster() As Boolean
Dim AOL As Long, MDI As Long, controlwin As Long, TryButton As Long
Dim Modal As Long, ModalStatic As Long, ModalText As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call aol_keyword("aol://4344:1580.prntcon.12263709.564517913")
Do: DoEvents
controlwin& = FindWindowEx(MDI&, 0&, "AOL Child", " Parental Controls")
TryButton& = FindWindowEx(controlwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until controlwin& <> 0& And TryButton& <> 0&
Call PostMessage(TryButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TryButton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
ModalStatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
Loop Until Modal& <> 0&
ModalText$ = api_gettext(ModalStatic&)
If Left(ModalText$, Len(ModalText$) - 1) <> "Set Parental Controls" Then
Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
Call PostMessage(controlwin&, WM_CLOSE, 0&, 0&)
aol_checkifmaster = False
ElseIf Left(ModalText$, Len(ModalText$) - 1) = "Set Parental Controls" Then
Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
Call PostMessage(controlwin&, WM_CLOSE, 0&, 0&)
aol_checkifmaster = True
End If
End Function

Sub aol_toolbarhide()
Dim AOL As Long, Toolbar As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Call ShowWindow(Toolbar&, SW_HIDE)
End Sub

Sub aol_toolbarshow()
Dim AOL As Long, Toolbar As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Call api_showwindow(Toolbar&)
End Sub


Sub api_click(Icon As String)
Call PostMessage(Icon$, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon$, WM_LBUTTONUP, 0&, 0&)
End Sub

Function api_getlistcount(List As Long) As Long
api_getlistcount& = PostMessage(List&, LB_GETCOUNT, 0&, 0&)
End Function



Sub aol_idle()
Dim Palette As Long, Modal As Long, PaletteButton As Long, ModalButton As Long
Do: DoEvents
Palette& = FindWindow("_AOL_Palette", vbNullString)
Modal& = FindWindow("_AOL_Modal", vbNullString)
PaletteButton& = FindWindowEx(Palette&, 0&, "_AOL_Icon", vbNullString)
ModalButton& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
If PaletteButton& <> 0& Or ModalButton& <> 0& Then
Call PostMessage(PaletteButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PaletteButton&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(ModalButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalButton&, WM_LBUTTONUP, 0&, 0&)
End If
Loop
End Sub
Sub aol_killwait()
Dim Modal As Long, OKButton As Long
Call aol_runmenu(4, 10)
Do: DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
OKButton& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Loop Until Modal& <> 0& And OKButton& <> 0&
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
End Sub






Sub aol_imignore(Who$)
aol_instantmessage "$IM_OFF " + Who$, Who$ + " is ignored"
End Sub
Sub aol_imoff()
Call aol_instantmessage("$IM_OFF", "ims disabled")
End Sub
Sub aol_imon()
Call aol_instantmessage("$IM_ON", "im enabled")
End Sub

Sub aol_imunignore(Who$)
Call aol_instantmessage("$IM_ON " & Who$, Who$ & " is unignored")
End Sub


Function aol_locatemember(Person As String) As String
Dim AOL As Long, MDI As Long, OKWin As Long
Dim ChildWin As Long, ChildWinCaption As String
Dim LocateWin As Long, LocateMsg1 As Long, LocateMsg As String
Dim OKButton As Long
AOL& = FindWindow("AOL Frame25", "America  Online")
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call aol_keyword("aol://3548:" & Person$)
Do: DoEvents
OKWin& = FindWindow("#32770", "America Online")
If OKWin& <> 0& Then Exit Do
ChildWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
ChildWinCaption$ = api_getcaption(ChildWin&)
If LCase(ChildWinCaption$) = LCase("locate " & Person$) Then
LocateWin& = ChildWin&
Else
Do: DoEvents
ChildWin& = FindWindowEx(MDI&, ChildWin&, "AOL Child", vbNullString)
ChildWinCaption$ = api_getcaption(ChildWin&)
If LCase(ChildWinCaption$) = LCase("locate " & Person$) Then
LocateWin& = ChildWin&
Exit Do
End If
OKWin& = FindWindow("#32770", "America Online")
Loop Until ChildWin& = 0& Or OKWin& <> 0&
End If
Loop Until LocateWin& <> 0& Or OKWin& <> 0&
If LocateWin& <> 0& Then
LocateMsg1& = FindWindowEx(LocateWin&, 0&, "_AOL_Static", vbNullString)
LocateMsg$ = api_gettext(LocateMsg1&)
If LCase(LocateMsg$) = LCase(Person$ & " is online, but not in a chat area.") Then
aol_locatemember$ = "Not in a chat."
ElseIf LCase(LocateMsg$) = LCase(Person$ & " is online, but in a private room.") Then
aol_locatemember$ = "Private room."
ElseIf LCase(LocateMsg$) Like LCase(Person$ & " is in chat room *") Then
aol_locatemember$ = Right(LocateMsg$, Len(LocateMsg$) - Len(Person$ & " is in chat room "))
End If
Call PostMessage(LocateWin&, WM_CLOSE, 0&, 0&)
ElseIf OKWin& <> 0& Then
OKButton& = FindWindowEx(OKWin&, 0&, "Button", "OK")
Do: DoEvents
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
OKWin& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(OKWin&, 0&, "Button", "OK")
Loop Until OKWin& = 0& And OKButton& = 0&
aol_locatemember$ = "Not signed on."
End If
End Function
Sub aol_ontop()
Dim TheWindow As Long
TheWindow& = FindWindow("AOL Frame25", vbNullString)
Call SetWindowPos(TheWindow&, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Sub aol_onbottom()
Dim TheWindow As Long
TheWindow& = FindWindow("AOL Frame25", vbNullString)
Call SetWindowPos(TheWindow&, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


Sub aol_rename(Name As String)
Dim TheWindow As Long
TheWindow& = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(TheWindow&, &HC, 0&, Name$)
End Sub


Sub aol_clickaoliconmenu(Icon As Long, itemnum As Long)
Dim sMod As Long, WinVis As Long, DoThis As Long
Dim CurPos As POINTAPI
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
For DoThis& = 1 To itemnum&
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub


Sub aol_addroom(TheList As Control, AddUser As Boolean)
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = aol_findroom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For Index& = 0 To PostMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = PostMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> aol_getuser$ Or AddUser = True Then
TheList.AddItem ScreenName$
End If
Next Index&
Call CloseHandle(mThread)
End If
End Sub
Sub aol_clearhistory()
Dim AOL As Long, ToolBar1 As Long, ToolBar2 As Long, Combo As Long, EditWin As Long
AOL& = FindWindow("AOL Frame25", "America  Online")
ToolBar1& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolBar2& = FindWindowEx(ToolBar1&, 0&, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(ToolBar2&, 0&, "_AOL_Combobox", vbNullString)
Call PostMessage(Combo&, CB_RESETCONTENT, 0&, 0&)
EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, "Type a Keyword Or Web Address and click Go")
End Sub



Function aol_checkim(Person As String) As Boolean
Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
Dim Available As Long, Available1 As Long, Available2 As Long
Dim Available3 As Long, oWindow As Long, oButton As Long
Dim oStatic As Long, oString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call aol_keyword("aol://9293:" & Person$)
Do
DoEvents
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
Available1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
Available2& = FindWindowEx(IM&, Available1&, "_AOL_Icon", vbNullString)
Available3& = FindWindowEx(IM&, Available2&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available3&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
DoEvents
Call PostMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Available&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
oWindow& = FindWindow("#32770", "America Online")
oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
Loop Until oWindow& <> 0& And oButton& <> 0&
Do
DoEvents
oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
oString$ = api_gettext(oStatic)
Loop Until oStatic& <> 0& And Len(oString$) > 15
If InStr(oString$, "is online and able to receive") <> 0 Then
aol_checkim = True
Else
aol_checkim = False
End If
Call PostMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function
Function aol_findim() As Long
Dim AOL As Long, MDI As Long, Child As Long, Caption As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Caption$ = api_gettext(Child&)
If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
aol_findim& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
Caption$ = api_gettext(Child&)
If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
aol_findim& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findim& = Child&
End Function
Function aol_findinfowindow() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
Dim AOLIcon2 As Long, AOLGlyph As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
aol_findinfowindow& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
aol_findinfowindow& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findinfowindow& = Child&
End Function

Function aol_findmailbox() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim TabControl As Long, TabPage As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabControl& <> 0& And TabPage& <> 0& Then
aol_findmailbox& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabControl& <> 0& And TabPage& <> 0& Then
aol_findmailbox& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findmailbox& = 0&
End Function

Function aol_findroom() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim Rich As Long, AOLList As Long
Dim AOLIcon As Long, AOLStatic As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
aol_findroom& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
aol_findroom& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findroom& = Child&
End Function


Function aol_findsendwindow() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim SendStatic As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
If SendStatic& <> 0& Then
aol_findsendwindow& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
If SendStatic& <> 0& Then
aol_findsendwindow& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findsendwindow& = 0&
End Function



Function api_getcaption(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
TextLength& = GetWindowTextLength(WindowHandle&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
api_getcaption$ = buffer$
End Function
Function api_getlisttext(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
TextLength& = PostMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
api_getlisttext$ = buffer$
End Function



Function aol_getuser() As String
Dim AOL As Long, MDI As Long, welcome As Long
Dim Child As Long, UserString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
UserString$ = api_getcaption(Child&)
If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
aol_getuser$ = UserString$
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
UserString$ = api_getcaption(Child&)
If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
aol_getuser$ = UserString$
Exit Function
End If
Loop Until Child& = 0&
End If
aol_getuser$ = ""
End Function


Function aol_imlastmsg() As String
Dim Rich As Long, MsgString As String, Spot As Long
Dim NewSpot As Long
Rich& = FindWindowEx(aol_findim&, 0&, "RICHCNTL", vbNullString)
MsgString$ = api_gettext(Rich&)
NewSpot& = InStr(MsgString$, Chr(9))
Do
Spot& = NewSpot&
NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
Loop Until NewSpot& <= 0&
MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
aol_imlastmsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function


Sub aol_imrespond(Msg As String)
Dim IM As Long, Rich As Long, Icon As Long
IM& = aol_findim&
If IM& = 0& Then Exit Sub
Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
Icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
DoEvents
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Function aol_imsender() As String
Dim IM As Long, Caption As String
Caption$ = api_getcaption(aol_findim&)
If InStr(Caption$, ":") = 0& Then
aol_imsender$ = ""
Exit Function
Else
aol_imsender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
End If
End Function
Function aol_imtext() As String
Dim Rich As Long
Rich& = FindWindowEx(aol_findim&, 0&, "RICHCNTL", vbNullString)
aol_imtext$ = api_gettext(Rich&)
End Function



Sub aol_instantmessage(Person As String, message As String)
Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
Dim SendButton As Long, OK As Long, Button As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call aol_keyword("aol://9293:" & Person$)
Do
DoEvents
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
OK& = FindWindow("#32770", "America Online")
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Loop Until OK& <> 0& Or IM& = 0&
If OK& <> 0& Then
Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End If
End Sub

Sub aol_keyword(KW As String)
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim Combo As Long, EditWin As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub


Sub api_hidewindow(hWnd As String)
Call ShowWindow(hWnd$, SW_HIDE)
End Sub


Function file_loadpicture(CTL As Control, Path As String) As Long
If Not file_exists(Path$) Then Exit Function Else
file_loadpicture& = CTL.Picture = (Path$)
End Function

Sub file_makedir(Dir As String)
If file_direxists(Dir$) Then Exit Sub
If Not file_direxists(Dir$) Then MkDir Dir$
End Sub

Function file_size(File As String) As String
Dim exists As Long
exists& = Len(Dir$(File$))
If Err Then Exit Function
file_size$ = FileLen(File$)
End Function

Sub form_controlstretch(Frm As Form, CTL As Control)
CTL.Height = Frm.Height
CTL.Width = Frm.Width
End Sub

Sub form_oval(Frm As Form)
Frm.Show
Call SetWindowRgn(Frm.hWnd, CreateEllipticRgn(0, 0, 195, 300), True)
End Sub

Sub form_transparent(Frm As Form)
On Error Resume Next
Call SetWindowLong(Frm.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
Call SetWindowPos(Frm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
End Sub

Sub other_creditz(Frm As Form, Label As Label, Name As String)
Dim X As Long, I As Long, Color As Long
Label.Alignment = 2
Label.Caption = Name$
Label.BackStyle = 0
Label.AutoSize = True
Frm.BackColor = vbBlack
For I& = 64 To 1 Step -5
X& = 255 - (I * 4 - 1)
Color& = RGB(X, X, X)
Label.ForeColor = Color&
other_delay (0.0000001)
Next I&
other_delay (0.8)
For I& = 1 To 64 Step 5
X = 255 - (I * 4 - 1)
Color& = RGB(X, X, X)
Label.ForeColor = Color&
other_delay (0.0000001)
Next I&
Label.ForeColor = &H0&
End Sub

Function other_linecount(MyString As String) As Long
Dim Spot As Long, Count As Long
If Len(MyString$) < 1 Then
other_linecount& = 0&
Exit Function
End If
Spot& = InStr(MyString$, Chr(13))
If Spot& <> 0& Then
other_linecount& = 1
Do
Spot& = InStr(Spot + 1, MyString$, Chr(13))
If Spot& <> 0& Then
other_linecount& = other_linecount& + 1
End If
Loop Until Spot& = 0&
End If
other_linecount& = other_linecount& + 1
End Function

Function aol_mailcountflash() As Long
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim Count As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
aol_mailcountflash& = Count&
End Function


Function aol_mailcountnew() As Long
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
aol_mailcountnew& = Count&
End Function

Function aol_mailcountold() As Long
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
aol_mailcountold& = Count&
End Function


Function aol_mailcountsent() As Long
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
aol_mailcountsent& = Count&
End Function

Sub aol_maildeleteflashbyindex(Index As Long)
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim fCount As Long, DeleteButton As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < Index& Then Exit Sub
DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
Call PostMessage(fList&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub aol_maildeleteflashduplicates(VBForm As Form, DisplayStatus As Boolean)
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim fCount As Long, DeleteButton As Long, SearchFor As Long
Dim SearchBox As Long, CurCaption As String
Dim sSender As String, sSubject As String
Dim cSender As String, cSubject As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < 2& Then Exit Sub
DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
CurCaption$ = VBForm.Caption
If fCount& = 0& Then Exit Sub
For SearchFor& = 0& To fCount& - 2
DoEvents
sSender$ = aol_mailsenderflash(SearchFor&)
sSubject$ = aol_mailsubjectflash(SearchFor&)
If sSender$ = "" Then
VBForm.Caption = CurCaption$
Exit Sub
End If
For SearchBox& = SearchFor& + 1 To fCount& - 1
If DisplayStatus = True Then
VBForm.Caption = "Checking #" & SearchFor& & " with #" & SearchBox&
End If
cSender$ = aol_mailsenderflash(SearchBox&)
cSubject$ = aol_mailsubjectflash(SearchBox&)
If cSender$ = sSender$ And cSubject$ = sSubject$ Then
Call PostMessage(fList&, LB_SETCURSEL, SearchBox&, 0&)
DoEvents
Call PostMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
DoEvents
SearchBox& = SearchBox& - 1
End If
Next SearchBox&
Next SearchFor&
VBForm.Caption = CurCaption$
End Sub

Function aol_mailsubjectflash(Index As Long) As String
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim fCount As Long, DeleteButton As Long, sLength As Long
Dim MyString As String, Spot As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < Index& Then Exit Function
DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
If fCount& = 0 Or Index& > fCount& - 1 Or Index& < 0& Then Exit Function
sLength& = PostMessage(fList&, LB_GETTEXTLEN, Index&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(fList&, LB_GETTEXT, Index&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
MyString$ = text_replacestring(MyString$, Chr(0), "")
aol_mailsubjectflash$ = MyString$
End Function


Sub pc_clickstart()
Dim TheParent As Long, TheWindow As Long
TheParent& = FindWindow("Shell_TrayWnd", vbNullString)
TheWindow& = FindWindowEx(TheParent&, 0&, "Button", vbNullString)
Call PostMessage(TheWindow&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(TheWindow&, WM_LBUTTONUP, 0&, 0&)
End Sub

Function text_replacestring(MyString As String, ToFind As String, ReplaceWith As String) As String
Dim Spot As Long, NewSpot As Long, LeftString As String
Dim RightString As String, NewString As String
Spot& = InStr(LCase(MyString$), LCase(ToFind))
NewSpot& = Spot&
Do
If NewSpot& > 0& Then
LeftString$ = Left(MyString$, NewSpot& - 1)
If Spot& + Len(ToFind$) <= Len(MyString$) Then
RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
Else
RightString = ""
End If
NewString$ = LeftString$ & ReplaceWith$ & RightString$
MyString$ = NewString$
Else
NewString$ = MyString$
End If
Spot& = NewSpot& + Len(ReplaceWith$)
If Spot& > 0 Then
NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
End If
Loop Until NewSpot& < 1
text_replacestring$ = NewString$
End Function


Function aol_mailsubjectnew(Index As Long) As String
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0 Or Index& > Count& - 1 Or Index& < 0& Then Exit Function
sLength& = PostMessage(mTree&, LB_GETTEXTLEN, Index&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(mTree&, LB_GETTEXT, Index&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
MyString$ = text_replacestring(MyString$, Chr(0), "")
aol_mailsubjectnew$ = MyString$
End Function


Function aol_mailsenderflash(Index As Long) As String
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim fCount As Long, DeleteButton As Long, sLength As Long
Dim MyString As String, Spot1 As Long, Spot2 As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < Index& Then Exit Function
DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
If fCount& = 0 Or Index& > fCount& - 1 Or Index& < 0& Then Exit Function
sLength& = PostMessage(fList&, LB_GETTEXTLEN, Index&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(fList&, LB_GETTEXT, Index&, MyString$)
Spot1& = InStr(MyString$, Chr(9))
Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
aol_mailsenderflash$ = MyString$
End Function


Function aol_mailsendernew(Index As Long) As String
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot1 As Long, Spot2 As Long, MyString As String
Dim Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0 Or Index& > Count& - 1 Or Index& < 0& Then Exit Function
sLength& = PostMessage(mTree&, LB_GETTEXTLEN, Index&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(mTree&, LB_GETTEXT, Index&, MyString$)
Spot1& = InStr(MyString$, Chr(9))
Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
aol_mailsendernew$ = MyString$
End Function


Sub aol_maildeletenewduplicates(VBForm As Form, DisplayStatus As Boolean)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long, dButton As Long
Dim SearchBox As Long, cSender As String, cSubject As String
Dim SearchFor As Long, sSender As String, sSubject As String
Dim CurCaption As String
MailBox& = aol_findmailbox&
CurCaption$ = VBForm.Caption
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0& Then Exit Sub
For SearchFor& = 0& To Count& - 2
DoEvents
sSender$ = aol_mailsendernew(SearchFor&)
sSubject$ = aol_mailsubjectnew(SearchFor&)
If sSender$ = "" Then
VBForm.Caption = CurCaption$
Exit Sub
End If
For SearchBox& = SearchFor& + 1 To Count& - 1
If DisplayStatus = True Then
VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
End If
cSender$ = aol_mailsendernew(SearchBox&)
cSubject$ = aol_mailsubjectnew(SearchBox&)
If cSender$ = sSender$ And cSubject$ = sSubject$ Then
Call PostMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
DoEvents
Call PostMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
DoEvents
SearchBox& = SearchBox& - 1
End If
Next SearchBox&
Next SearchFor&
VBForm.Caption = CurCaption$
End Sub

Sub aol_mailforward(SendTo As String, message As String, DeleteFwd As Boolean)
Dim AOL As Long, MDI As Long, Error As Long
Dim OpenForward As Long, OpenSend As Long, SendButton As Long
Dim DoIt As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, Rich As Long, fCombo As Long
Dim Combo As Long, Button1 As Long, Button2 As Long
Dim TempSubject As String
OpenForward& = aol_findforwardwindow
If OpenForward& = 0 Then Exit Sub
SendButton& = FindWindowEx(OpenForward&, 0&, "_AOL_Icon", vbNullString)
For DoIt& = 1 To 6
SendButton& = FindWindowEx(OpenForward&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
OpenSend& = aol_findsendwindow
EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
For DoIt& = 1 To 13
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&
Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
If DeleteFwd = True Then
TempSubject$ = api_gettext(EditSubject&)
TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
DoEvents
End If
Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SendTo$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents
Do Until OpenSend& = 0& Or Error& <> 0&
DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
OpenSend& = aol_findsendwindow
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
For DoIt& = 1 To 11
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
other_delay 1
Loop
If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
End Sub


Function aol_findforwardwindow() As Long
Dim AOL As Long, MDI As Long, Child As Long
Dim Rich1 As Long, Rich2 As Long, Combo As Long
Dim FontCombo As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
aol_findforwardwindow& = Child&
Exit Function
Else
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
aol_findforwardwindow& = Child&
Exit Function
End If
Loop Until Child& = 0&
End If
aol_findforwardwindow& = 0&
End Function


Sub aol_mailopenemailflash(Index As Long)
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim fCount As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
If fCount& < Index& Then Exit Sub
Call PostMessage(fList&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)
End Sub


Sub aol_mailopenemailnew(Index As Long)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& < Index& Then Exit Sub
Call PostMessage(mTree&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Sub aol_mailopenemailold(Index As Long)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& < Index& Then Exit Sub
Call PostMessage(mTree&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Sub aol_mailopenemailsent(Index As Long)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& < Index& Then Exit Sub
Call PostMessage(mTree&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Sub aol_mailopenflash()
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub


Sub aol_mailopennew()
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
Dim WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub

Sub aol_mailopenold()
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
For DoThis& = 1 To 4
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub

Sub aol_mailopensent()
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
For DoThis& = 1 To 5
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub


Sub aol_mailtolistflash(TheList As ListBox)
Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
Dim Count As Long, MyString As String, AddMails As Long
Dim sLength As Long, Spot As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
If fMail& = 0& Then Exit Sub
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(fList&, LB_GETCOUNT, 0&, 0&)
MyString$ = String(255, 0)
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = PostMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
MyString$ = text_replacestring(MyString$, Chr(0), "")
TheList.AddItem MyString$
Next AddMails&
End Sub
Sub aol_mailtolistnew(TheList As ListBox)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = PostMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
TheList.AddItem MyString$
Next AddMails&
End Sub


Sub aol_mailtolistold(TheList As ListBox)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = PostMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
TheList.AddItem MyString$
Next AddMails&
End Sub


Sub aol_mailtolistsent(TheList As ListBox)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = PostMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)
Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
TheList.AddItem MyString$
Next AddMails&
End Sub


Function aol_profileget(ScreenName As String) As String
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
Dim pWindow As Long, pTextWindow As Long, pString As String
Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
Dim WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
Do
DoEvents
pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
Call PostMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
pString$ = api_gettext(pTextWindow&)
NoWindow& = FindWindow("#32770", "America Online")
Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
DoEvents
If NoWindow& <> 0& Then
OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
aol_profileget$ = "< No Profile >"
Else
Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
aol_profileget$ = pString$
End If
End Function
Function aol_roomcount() As Long
Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
Dim Count As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
rMail& = aol_findroom
rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
Count& = PostMessage(rList&, LB_GETCOUNT, 0&, 0&)
aol_roomcount& = Count&
End Function

Sub aol_runmenu(TopMenu As Long, SubMenu As Long)
Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
Dim mVal As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
aMenu& = GetMenu(AOL&)
sMenu& = GetSubMenu(aMenu&, TopMenu&)
mnID& = GetMenuItemID(sMenu&, SubMenu&)
Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub


Sub aol_runmenubystring(SearchString As String)
Dim AOL As Long, aMenu As Long, mCount As Long
Dim LookFor As Long, sMenu As Long, sCount As Long
Dim LookSub As Long, sID As Long, sString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
aMenu& = GetMenu(AOL&)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
sMenu& = GetSubMenu(aMenu&, LookFor&)
sCount& = GetMenuItemCount(sMenu&)
For LookSub& = 0 To sCount& - 1
sID& = GetMenuItemID(sMenu&, LookSub&)
sString$ = String$(100, " ")
Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
If InStr(LCase(sString$), LCase(SearchString$)) Then
Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
Exit Sub
End If
Next LookSub&
Next LookFor&
End Sub
Sub aol_scroll(ScrollString As String)
Dim CurLine As String, Count As Long, ScrollIt As Long
Dim sProgress As Long
If aol_findroom& = 0 Then Exit Sub
If ScrollString$ = "" Then Exit Sub
Count& = other_linecount(ScrollString$)
sProgress& = 1
For ScrollIt& = 1 To Count&
CurLine$ = other_linefromstring(ScrollString$, ScrollIt&)
If Len(CurLine$) > 3 Then
If Len(CurLine$) > 92 Then
CurLine$ = Left(CurLine$, 92)
End If
Call aol_chatsend(CurLine$)
other_delay 0.7
End If
sProgress& = sProgress& + 1
If sProgress& > 4 Then
sProgress& = 1
other_delay 0.5
End If
Next ScrollIt&
End Sub

Function other_linefromstring(MyString As String, Line As Long) As String
Dim theline As String, Count As Long
Dim FSpot As Long, LSpot As Long, DoIt As Long
Count& = other_linecount(MyString$)
If Line& > Count& Then
Exit Function
End If
If Line& = 1 And Count& = 1 Then
other_linefromstring$ = MyString$
Exit Function
End If
If Line& = 1 Then
theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
theline$ = text_replacestring(theline$, Chr(13), "")
theline$ = text_replacestring(theline$, Chr(10), "")
other_linefromstring$ = theline$
Exit Function
Else
FSpot& = InStr(MyString$, Chr(13))
For DoIt& = 1 To Line& - 1
LSpot& = FSpot&
FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
Next DoIt
If FSpot = 0 Then
FSpot = Len(MyString$)
End If
theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
theline$ = text_replacestring(theline$, Chr(13), "")
theline$ = text_replacestring(theline$, Chr(10), "")
other_linefromstring$ = theline$
End If
End Function


Sub aol_sendmail(Person As String, subject As String, message As String)
Dim AOL As Long, MDI As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
Dim Rich As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, SendButton As Long
Dim Combo As Long, fCombo As Long, ErrorWindow As Long
Dim Button1 As Long, Button2 As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
For DoIt& = 1 To 13
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&
Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
DoEvents
Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents
other_delay 0.2
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub


Sub aol_setmailprefs()
Dim AOL As Long, Tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim MDI As Long, mPrefs As Long, mButton As Long
Dim gStatic As Long, mStatic As Long, fStatic As Long
Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
Dim CurPos As POINTAPI, WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
For DoThis& = 1 To 3
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
Do
DoEvents
mPrefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
gStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "General")
mStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Mail")
fStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Font")
maStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Marketing")
Loop Until mPrefs& <> 0& And gStatic& <> 0& And mStatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&
mButton& = FindWindowEx(mPrefs&, 0&, "_AOL_Icon", vbNullString)
mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
Do
DoEvents
Call PostMessage(mButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(mButton&, WM_LBUTTONUP, 0&, 0&)
dMod& = FindWindow("_AOL_Modal", "Mail Preferences")
other_delay 0.6
Loop Until dMod& <> 0&
ConfirmCheck& = FindWindowEx(dMod&, 0&, "_AOL_Checkbox", vbNullString)
CloseCheck& = FindWindowEx(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
OKButton& = FindWindowEx(dMod&, 0&, "_AOL_icon", vbNullString)
Call PostMessage(ConfirmCheck&, BM_SETCHECK, False, vbNullString)
Call PostMessage(CloseCheck&, BM_SETCHECK, True, vbNullString)
Call PostMessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Call PostMessage(mPrefs&, WM_CLOSE, 0&, 0&)
End Sub


Sub other_normalcursor()
Do
DoEvents
If Screen.MousePointer = 0 Then Exit Do
Loop
End Sub

Sub text_copy(Text As Textbox)
Text.SelStart = 0
Text.SelLength = Len(Text)
Clipboard.SetText Text.SelText
End Sub

Function text_count(Text As Textbox) As String
text_count$ = Len(Text)
End Function

Sub text_cut(Text As Textbox)
Text.SelStart = 0
Text.SelLength = Len(Text)
Clipboard.SetText Text.SelText
Text.SelText = ""
End Sub


Sub text_paste(Text As Textbox)
Text.SelText = Clipboard.GetText()
End Sub


Sub text_selall(Text As Textbox)
Text.SelStart = 0
Text.SelLength = Len(Text.Text)
End Sub

Function text_switchstrings(MyString As String, String1 As String, String2 As String) As String
Dim TempString As String, Spot1 As Long, Spot2 As Long
Dim Spot As Long, ToFind As String, ReplaceWith As String
Dim NewSpot As Long, LeftString As String, RightString As String
Dim NewString As String
If Len(String2) > Len(String1) Then
TempString$ = String1$
String1$ = String2$
String2$ = TempString$
End If
Spot1& = InStr(MyString$, String1$)
Spot2& = InStr(MyString$, String2$)
If Spot1& = 0& And Spot2& = 0& Then
text_switchstrings$ = MyString$
Exit Function
End If
If Spot1& < Spot2& Or Spot2& = 0 Or Len(String1$) = Len(String2$) Then
If Spot1& > 0 Then
Spot& = Spot1&
ToFind$ = String1$
ReplaceWith$ = String2$
End If
End If
If Spot2& < Spot1& Or Spot1& = 0& Then
If Spot2& > 0& Then
Spot& = Spot2&
ToFind$ = String2$
ReplaceWith$ = String1$
End If
End If
If Spot1& = 0& And Spot2& = 0& Then
text_switchstrings$ = MyString$
Exit Function
End If
NewSpot& = Spot&
Do
If NewSpot& > 0& Then
LeftString$ = Left(MyString$, NewSpot& - 1)
If Spot& + Len(ToFind$) <= Len(MyString$) Then
RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
Else
RightString$ = ""
End If
NewString$ = LeftString$ & ReplaceWith$ & RightString$
MyString$ = NewString$
Else
NewString$ = MyString$
End If
Spot& = NewSpot + Len(ReplaceWith$) - Len(ToFind$) + 1
If Spot& <> 0& Then
Spot1& = InStr(Spot&, MyString$, String1$)
Spot2& = InStr(Spot&, MyString$, String2$)
End If
If Spot1& = 0& And Spot2& = 0& Then
text_switchstrings$ = MyString$
Exit Function
End If
If Spot1& < Spot2& Or Spot2& = 0& Or Len(String1$) = Len(String2$) Then
If Spot1& > 0& Then
Spot& = Spot1&
ToFind$ = String1$
ReplaceWith$ = String2$
End If
End If
If Spot2& < Spot1& Or Spot1& = 0& Then
If Spot2& > 0& Then
Spot& = Spot2&
ToFind$ = String2$
ReplaceWith$ = String1$
End If
End If
If Spot1& = 0& And Spot2& = 0& Then
Spot& = 0&
End If
If Spot& > 0& Then
NewSpot& = InStr(Spot&, MyString$, ToFind$)
Else
NewSpot& = Spot&
End If
Loop Until NewSpot& < 1&
text_switchstrings$ = NewString$
End Function


Sub aol_waitforokorroom(Room As String)
Dim RoomTitle As String, FullWindow As Long, FullButton As Long
Room$ = LCase(text_replacestring(Room$, " ", ""))
Do
DoEvents
RoomTitle$ = api_getcaption(aol_findroom&)
RoomTitle$ = LCase(text_replacestring(Room$, " ", ""))
FullWindow& = FindWindow("#32770", "America Online")
FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
DoEvents
If FullWindow& <> 0& Then
Do
DoEvents
Call PostMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
FullWindow& = FindWindow("#32770", "America Online")
FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
Loop Until FullWindow& = 0& And FullButton& = 0&
End If
DoEvents
End Sub

Sub aol_hide()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL&, SW_HIDE)
End Sub


Sub aol_show()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL&, SW_SHOW)
End Sub

Sub aol_chatignorebyindex(Index As Long)
Dim Room As Long, sList As Long, iWindow As Long
Dim iCheck As Long, a As Long, Count As Long
Count& = aol_roomcount&
If Index& > Count& - 1 Then Exit Sub
Room& = aol_findroom&
sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Call PostMessage(sList&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
Do
DoEvents
iWindow& = aol_findinfowindow
Loop Until iWindow& <> 0&
DoEvents
iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
DoEvents
Do
DoEvents
a& = PostMessage(iCheck&, BM_GETCHECK, 0&, 0&)
Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
DoEvents
Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Loop Until a& <> 0&
DoEvents
Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub

Sub aol_chatignorebyname(Name As String)
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim lIndex As Long
Room& = aol_findroom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For Index& = 0 To PostMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = PostMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> aol_getuser$ And LCase(ScreenName$) = LCase(Name$) Then
lIndex& = Index&
Call aol_chatignorebyindex(lIndex&)
DoEvents
Exit Sub
End If
Next Index&
Call CloseHandle(mThread)
End If
End Sub


Sub aol_chatnotifyoff()
Dim Room As Long, ChatOptions As Long, PrefIcon As Long, NotifyCheck As Long, OkBut As Long
Room& = aol_findroom&
If Room& = 0 Then Exit Sub
PrefIcon& = FindWindowEx(Room&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(PrefIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
ChatOptions& = FindWindow("_AOL_Modal", "Chat Preferences")
NotifyCheck& = FindWindowEx(ChatOptions&, 0, "_AOL_Checkbox", vbNullString)
If ChatOptions& <> 0 And NotifyCheck& <> 0 Then Exit Do
Loop
Call other_delay(0.5)
Call PostMessage(NotifyCheck&, BM_SETCHECK, False, vbNullString)
Do
DoEvents
OkBut& = FindWindowEx(ChatOptions&, 0, "_AOL_Icon", vbNullString)
If OkBut& <> 0 Then Exit Do
Loop
DoEvents
Call PostMessage(OkBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(OkBut&, WM_LBUTTONUP, 0, 0)
End Sub

Sub aol_chatnotifyon()
Dim Room As Long, ChatOptions As Long, PrefIcon As Long, NotifyCheck As Long, OkBut As Long
Room& = aol_findroom&
If Room& = 0 Then Exit Sub
PrefIcon& = FindWindowEx(Room&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(PrefIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
ChatOptions& = FindWindow("_AOL_Modal", "Chat Preferences")
NotifyCheck& = FindWindowEx(ChatOptions&, 0, "_AOL_Checkbox", vbNullString)
If ChatOptions& <> 0 And NotifyCheck& <> 0 Then Exit Do
Loop
Call other_delay(0.5)
Call PostMessage(NotifyCheck&, BM_SETCHECK, True, vbNullString)
Do
DoEvents
OkBut& = FindWindowEx(ChatOptions&, 0, "_AOL_Icon", vbNullString)
If OkBut& <> 0 Then Exit Do
Loop
DoEvents
Call PostMessage(OkBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(OkBut&, WM_LBUTTONUP, 0, 0)
End Sub


Sub aol_chatsoundsoff()
Dim Room As Long, ChatOptions As Long, PrefIcon As Long, NotifyCheck As Long, OkBut As Long
Room& = aol_findroom&
If Room& = 0 Then Exit Sub
PrefIcon& = FindWindowEx(Room&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(PrefIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
ChatOptions& = FindWindow("_AOL_Modal", "Chat Preferences")
If ChatOptions& <> 0 Then Exit Do
Loop
Call other_delay(0.5)
NotifyCheck& = FindWindowEx(ChatOptions&, 0, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
DoEvents
Call PostMessage(NotifyCheck&, BM_SETCHECK, False, vbNullString)
Do
DoEvents
OkBut& = FindWindowEx(ChatOptions&, 0, "_AOL_Icon", vbNullString)
If OkBut& <> 0 Then Exit Do
Loop
DoEvents
Call PostMessage(OkBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(OkBut&, WM_LBUTTONUP, 0, 0)
End Sub


Sub aol_chatsoundson()
Dim Room As Long, ChatOptions As Long, PrefIcon As Long, NotifyCheck As Long, OkBut As Long
Room& = aol_findroom&
If Room& = 0 Then Exit Sub
PrefIcon& = FindWindowEx(Room&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(Room&, PrefIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(PrefIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon, WM_LBUTTONUP, 0&, 0&)
DoEvents
Do
DoEvents
ChatOptions& = FindWindow("_AOL_Modal", "Chat Preferences")
If ChatOptions& <> 0 Then Exit Do
Loop
Call other_delay(0.5)
NotifyCheck& = FindWindowEx(ChatOptions&, 0, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
NotifyCheck& = FindWindowEx(ChatOptions&, NotifyCheck&, "_AOL_Checkbox", vbNullString)
DoEvents
Call PostMessage(NotifyCheck&, BM_SETCHECK, True, vbNullString)
Do
DoEvents
OkBut& = FindWindowEx(ChatOptions&, 0, "_AOL_Icon", vbNullString)
If OkBut& <> 0 Then Exit Do
Loop
DoEvents
Call PostMessage(OkBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(OkBut&, WM_LBUTTONUP, 0, 0)
End Sub


Function aol_checkalive(ScreenName As String) As Boolean
Dim AOL As Long, MDI As Long, ErrorWindow As Long, ErrorTextWindow As Long, ErrorString As String
Dim MailWindow As Long, NoWindow As Long, NoButton As Long
Call aol_sendmail("*, " & ScreenName$, "check'n", "dobie dobie doo...")
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
Do
DoEvents
ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorString$ = api_gettext(ErrorTextWindow&)
Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
If InStr(LCase(text_replacestring(ErrorString$, " ", "")), LCase(text_replacestring(ScreenName$, " ", ""))) > 0 Then
aol_checkalive = False
Else
aol_checkalive = True
End If
MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
DoEvents
Do
DoEvents
NoWindow& = FindWindow("#32770", "America Online")
NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
Loop Until NoWindow& <> 0& And NoButton& <> 0
Call PostMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function



Sub aol_ghostoff()
Dim AOLWin As Long, AOLMDI As Long, PrefWin As Long
Dim BuddyPref As Long, UserBL As Long, UserBL2 As Long
Dim PrefIcon As Long, PrefIcon2 As Long, SaveIcon As Long
Call aol_keyword("Buddy")
Do: DoEvents
AOLWin& = FindWindow("AOL Frame25", "America  Online")
AOLMDI& = FindWindowEx(AOLWin&, 0, "MDIClient", vbNullString)
UserBL& = FindWindowEx(AOLMDI&, 0, "AOL Child", aol_getuser + "'s Buddy List")
UserBL2& = FindWindowEx(AOLMDI&, 0, "AOL Child", aol_getuser + "'s Buddy Lists")
If UserBL& Then GoTo BL1
If UserBL2& Then GoTo BL2
Loop
BL1:
Call other_delay(0.5)
Do: DoEvents
PrefIcon& = FindWindowEx(UserBL&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
If PrefIcon& Then
Call PostMessage(PrefIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon&, WM_LBUTTONUP, 0&, 0&)
GoTo SetPref
End If
Loop
BL2:
Call other_delay(0.5)
Do: DoEvents
PrefIcon2& = FindWindowEx(UserBL2&, 0, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
If PrefIcon2& Then
Call PostMessage(PrefIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon2&, WM_LBUTTONUP, 0&, 0&)
GoTo SetPref
End If
Loop
SetPref:
Do: DoEvents
PrefWin& = FindWindowEx(AOLMDI&, 0, "AOL Child", "Privacy Preferences")
If PrefWin& Then Exit Do
Loop
Call other_delay(0.5)
Do: DoEvents
BuddyPref& = FindWindowEx(PrefWin&, 0, "_AOL_Checkbox", vbNullString)
If BuddyPref& Then Exit Do
Loop
Call PostMessage(BuddyPref&, BM_SETCHECK, True, vbNullString)
Do: DoEvents
SaveIcon& = FindWindowEx(PrefWin&, 0, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
If SaveIcon& Then Exit Do
Loop
Call PostMessage(SaveIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SaveIcon&, WM_LBUTTONUP, 0&, 0&)
aol_waitforok
Call other_delay(0.5)
If UserBL& Then api_closewindow UserBL&
If UserBL2& Then api_closewindow UserBL2&
End Sub

Sub api_closewindow(Window As Long)
Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub


Sub aol_waitforok()
Dim FullWindow As Long, FullButton As Long
Do: DoEvents
FullWindow& = FindWindow("#32770", "America Online")
FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
If FullWindow& <> 0& And FullButton& <> 0& Then
Call PostMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
Exit Do
End If
Loop
End Sub

Sub aol_ghoston()
Dim AOLWin As Long, AOLMDI As Long, PrefWin As Long
Dim BuddyPref As Long, UserBL As Long, UserBL2 As Long
Dim PrefIcon As Long, PrefIcon2 As Long, SaveIcon As Long
Call aol_keyword("Buddy")
Do: DoEvents
AOLWin& = FindWindow("AOL Frame25", "America  Online")
AOLMDI& = FindWindowEx(AOLWin&, 0, "MDIClient", vbNullString)
UserBL& = FindWindowEx(AOLMDI&, 0, "AOL Child", aol_getuser + "'s Buddy List")
UserBL2& = FindWindowEx(AOLMDI&, 0, "AOL Child", aol_getuser + "'s Buddy Lists")
If UserBL& Then GoTo BL1
If UserBL2& Then GoTo BL2
Loop
BL1:
Call other_delay(0.5)
Do: DoEvents
PrefIcon& = FindWindowEx(UserBL&, 0, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
If PrefIcon& Then
Call PostMessage(PrefIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon&, WM_LBUTTONUP, 0&, 0&)
GoTo SetPref
End If
Loop
BL2:
Call other_delay(0.5)
Do: DoEvents
PrefIcon2& = FindWindowEx(UserBL2&, 0, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
If PrefIcon2& Then
Call PostMessage(PrefIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(PrefIcon2&, WM_LBUTTONUP, 0&, 0&)
GoTo SetPref
End If
Loop
SetPref:
Do: DoEvents
PrefWin& = FindWindowEx(AOLMDI&, 0, "AOL Child", "Privacy Preferences")
If PrefWin& Then Exit Do
Loop
Call other_delay(0.5)
Do: DoEvents
BuddyPref& = FindWindowEx(PrefWin&, 0, "_AOL_Checkbox", vbNullString)
BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
If BuddyPref& Then Exit Do
Loop
Call PostMessage(BuddyPref&, BM_SETCHECK, True, vbNullString)
Do: DoEvents
SaveIcon& = FindWindowEx(PrefWin&, 0, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
If SaveIcon& Then Exit Do
Loop
Call PostMessage(SaveIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SaveIcon&, WM_LBUTTONUP, 0&, 0&)
aol_waitforok
Call other_delay(0.5)
If UserBL& Then api_closewindow UserBL&
If UserBL2& Then api_closewindow UserBL2&
End Sub


Sub aol_kill46min()
Dim Win46Button As Long, AOLWin As Long
AOLWin& = FindWindow("_AOL_Palette", "America Online Timer")
Win46Button& = FindWindowEx(AOLWin&, 0, "_AOL_Icon", vbNullString)
If Win46Button& <> 0 Then
Call PostMessage(Win46Button&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Win46Button&, WM_LBUTTONUP, 0&, 0&)
End If
End Sub


Sub aol_maildeletenewbyindex(Index As Long)
Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long, dButton As Long
MailBox& = aol_findmailbox&
If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = PostMessage(mTree&, LB_GETCOUNT, 0&, 0&)
If Index& > Count& - 1 Or Index& < 0& Then Exit Sub
Call PostMessage(mTree&, LB_SETCURSEL, Index&, 0&)
dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
Call PostMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
End Sub


Sub file_copy(File As String, DestFile As String)
If File$ = "" Then Exit Sub
If DestFile$ = "" Then Exit Sub
If Not file_exists(File$) Then Exit Sub
If InStr(Right$(File$, 4), ".") = 0 Then Exit Sub
If InStr(Right$(DestFile$, 4), ".") = 0 Then Exit Sub
FileCopy File$, DestFile$
Exit Sub
End Sub
Function cd_changetracks(Track As String)
Call MCISendString("seek cd to " & Track$, 0, 0, 0)
End Function



Sub cd_closedoor()
Call MCISendString("set cd door closed", 0, 0, 0)
End Sub


Sub cd_opendoor()
Call MCISendString("set cd door open", 0, 0, 0)
End Sub


Sub cd_pause()
Call MCISendString("pause cd", 0, 0, 0)
End Sub


Sub cd_play()
Call MCISendString("play cd", 0, 0, 0)
End Sub


Sub cd_stop()
Call MCISendString("stop cd wait", 0, 0, 0)
End Sub







Function file_direxists(TheDirectory As String) As Boolean
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory$, 1) <> "/" Then TheDirectory$ = TheDirectory$ + "/"
Check = Len(Dir$(TheDirectory$))
If Err Or Check = 0 Then
file_direxists = False
Else
file_direxists = True
End If
End Function

Sub file_playwav(File As String)
Dim SafeFile As String
SafeFile$ = Dir(File$)
If SafeFile$ <> "" Then
Call sndPlaySound(File$, SND_FLAG)
End If
End Sub




Function ini_receive(appname$, Keyname$, FileName$) As String
Dim Retstr As String
Retstr = String(255, Chr(0))
ini_receive = Left(Retstr, GetPrivateProfileString(appname$, ByVal Keyname$, "", Retstr, Len(Retstr), FileName$))
End Function

Function ini_write(Section$, Colum$, Text$, Path$)
Dim r As Long
r& = WritePrivateProfileString(Section$, Colum$, Text$, Path$)
End Function
Function pc_date() As String
pc_date$ = Format$(Now, "mmmm/dd/yyyy")
End Function
Function pc_day() As String
pc_day$ = Format$(Now, "dddd")
End Function

Function pc_time() As String
pc_time$ = Format$(Now, "h:mm:ss AM/PM")
End Function



Function pc_restart()
pc_restart = ExitWindows(EWX_REBOOT, 0)
End Function



Function pc_shutdown()
pc_shutdown = ExitWindows(EWX_SHUTDOWN, 0)
End Function

Sub file_wavstop()
Call file_playwav(" ")
End Sub
Sub file_open(File As String)
If Not file_exists(File$) Then Exit Sub
Shell (File$)
End Sub


Function file_exists(ByVal sFileName As String) As Boolean
Dim I As Integer
On Error Resume Next
I = Len(Dir$(sFileName))
If Err Or I = 0 Then
file_exists = False
Else
file_exists = True
End If
End Function
Sub file_delete(File As String)
Dim NoFreeze As Integer
If Not file_exists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub

Function cd_isplaying() As Boolean
cd_isplaying = False
Dim s As String * 50
Call MCISendString("status cd mode", s, Len(s), 0)
If Mid$(s, 1, 7) = "playing" Then
cd_isplaying = True
Else
cd_isplaying = False
End If
End Function





Sub list_fonts(Lst As Control)
Dim X As Long
For X& = 1 To Screen.FontCount
Lst.AddItem Screen.Fonts(X&)
Next X&
End Sub



Sub form_center(Frm As Form)
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub


Sub form_drag(Frm As Form)
Call ReleaseCapture
Call PostMessage(Frm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub



Sub form_onbottom(Frm As Form)
Call SetWindowPos(Frm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub



Sub text_print(Text As Textbox)
Printer.Print Text
other_delay (0.4)
Printer.EndDoc
End Sub

Function text_upcase(Text As Textbox) As String
text_upcase$ = UCase(Text)
End Function
Function text_locase(Text As Textbox) As String
text_locase$ = LCase(Text)
End Function

Function text_trimspaces(Text)
Dim thechar As String
Dim TheChars As String
If InStr(Text, " ") = 0 Then
text_trimspaces = Text
Exit Function
End If
For text_trimspaces = 1 To Len(Text)
thechar$ = Mid(Text, text_trimspaces, 1)
TheChars$ = TheChars$ & thechar$
If thechar$ = " " Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
End If
Next text_trimspaces
text_trimspaces = TheChars$
End Function


Sub list_remove(Lst As Control)
Dim listcount As Long
listcount& = Lst.listcount
Do While listcount& > 0&
listcount& = listcount& - 1
If Lst.Selected(listcount&) = True Then
Lst.RemoveItem (listcount&)
End If
Loop
End Sub



Function spy_class() As String
spy_class$ = api_getclass(spy_handle&)
End Function


Sub list_killdupe(Lst As Control)
Dim X As Long, I As Long, Current As Long
Dim Nower As Long
For X& = 0 To Lst.listcount - 1
Current& = Lst.List(X&)
For I& = 0 To Lst.listcount - 1
Nower& = Lst.List(I&)
If I& = X& Then GoTo dontkill
If Nower& = Current& Then Lst.RemoveItem (I&)
dontkill:
Next I&
Next X&
End Sub

Sub aol_killglyph()
Dim MainParent As Long, Parent1 As Long, Parent2 As Long
Dim TheWindow As Long
MainParent& = FindWindow("AOL Frame25", vbNullString)
Parent1& = FindWindowEx(MainParent&, 0&, "AOL Toolbar", vbNullString)
Parent2& = FindWindowEx(Parent1&, 0&, "_AOL_Toolbar", vbNullString)
TheWindow& = FindWindowEx(Parent2&, 0&, "_AOL_Glyph", vbNullString)
Call PostMessage(TheWindow&, &H10, 0&, 0&)
End Sub


Sub aol_killmodal()
Dim Modal As Long
Modal& = FindWindow("_AOL_Modal", vbNullString)
Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
End Sub



Sub aol_unupchat()
Dim AOL As Long, AOModal As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOModal& = FindWindowEx(AOL&, 0&, "MDIClient", "_AOL_Modal")
Call EnableWindow(AOModal&, 1)
Call EnableWindow(AOL&, 0)
End Sub

Sub aol_upchat()
Dim AOL As Long, AOModal As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOModal& = FindWindowEx(AOL&, 0&, "MDIClient", "_AOL_Modal")
Call EnableWindow(AOModal&, 0)
Call EnableWindow(AOL&, 1)
End Sub

Function aol_useronline() As Boolean
Dim Online As String
Online$ = aol_getuser()
If Online$ = "" Then
aol_useronline = False
Else
aol_useronline = True
End If
End Function


Sub aol_kwfeatuerdent(Name As String)
aol_keyword ("aol://2719:22-2-" & Name$)
End Sub
Sub aol_kwfeaturedplaces(Name As String)
aol_keyword ("aol://2719:25-2-" & Name$)
End Sub


Sub aol_kwfeaturedromance(Name As String)
aol_keyword ("aol://2719:26-2-" & Name$)
End Sub


Sub aol_kwfeaturedspecial(Name As String)
aol_keyword ("aol://2719:27-2-" & Name$)
End Sub


Sub aol_kwfeaturedtown(Name As String)
aol_keyword ("aol://2719:21-2-" & Name$)
End Sub


Sub aol_kwmemberenter(Name As String)
aol_keyword ("aol://2719:62-2-" & Name$)
End Sub


Sub aol_kwmemberfreinds(Name As String)
aol_keyword ("aol://2719:74-2-" & Name$)
End Sub


Sub aol_kwmemberlife(Name As String)
aol_keyword ("aol://2719:63-2-" & Name$)
End Sub


Sub aol_kwmembernews(Name As String)
aol_keyword ("aol://2719:64-2-" & Name$)
End Sub
Sub form_aolparent(Frm As Form)
Dim AOL As Long, MDI As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call SetParent(Frm.hWnd, MDI&)
End Sub


Sub aol_kwmemberplaces(Name As String)
aol_keyword ("aol://2719:65-2-" & Name$)
End Sub


Sub aol_kwmemberromance(Name As String)
aol_keyword ("aol://2719:66-2-" & Name$)
End Sub


Sub aol_kwmemberspecial(Name As String)
aol_keyword ("aol://2719:67-2-" & Name$)
End Sub

Sub aol_kwmembertown(Name As String)
aol_keyword ("aol://2719:61-2-" & Name$)
End Sub


Sub aol_kwprivateroom(Name As String)
aol_keyword ("aol://2719:2-2-" & Name$)
End Sub


Sub aol_signoff()
If aol_getuser <> "" Then Exit Sub
aol_runmenubystring ("&Sign Off")
End Sub


Sub aol_kwfeaturedlife(Name As String)
aol_keyword ("aol://2719:23-2-" & Name$)
End Sub


Sub aol_kwfeaturednews(Name As String)
aol_keyword ("aol://2719:24-2-" & Name$)
End Sub


Sub aol_kwfeaturedfreind(Name As String)
aol_keyword ("aol://2719:34-2-" & Name$)
End Sub


Function other_percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
other_percent = Int(Complete / Total * TotalOutput)
End Function

Sub other_percentbar(Shape As Control, Done As Integer, Total As Variant)
Dim X As Long
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "Small Fonts"
Shape.FontSize = 7
Shape.FontBold = False
X& = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(X& - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print other_percent(Done, Total, 100) & "%"
End Sub



Function aol_lastchatline()
Dim ChatText As Long, ChatTrimNum As Long, ChatTrim As String
ChatText& = aol_lastchatlinewithsn
ChatTrimNum& = Len(aol_snfromlastchatline)
ChatTrim$ = Mid$(ChatText&, ChatTrimNum& + 4, Len(ChatText&) - Len(aol_snfromlastchatline))
aol_lastchatline = ChatTrim$
End Function

Function aol_lastchatlinewithsn()
Dim ChatText As String, FindChar As Long, thechar As String, TheChars As String
Dim thechattext As String
Dim lastlen%, LastLine%
ChatText$ = api_getchattext
For FindChar& = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar&, 1)
TheChars$ = TheChars$ & thechar$
If thechar$ = Chr(13) Then
thechattext$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
TheChars$ = ""
End If
Next FindChar&
lastlen% = Val(FindChar&) - Len(TheChars$)
LastLine% = Mid(ChatText$, lastlen, Len(TheChars$))
aol_lastchatlinewithsn = LastLine
End Function
Function api_getchattext()
Dim Room As Long, AORich As Long, ChatText As Long
Room& = aol_findroom()
AORich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
ChatText& = api_gettext(AORich&)
api_getchattext = ChatText&
End Function


Function aol_snfromlastchatline()
Dim ChatText As String, ChatTrim As String, Z As Long, sn As Long
ChatText$ = aol_lastchatlinewithsn
ChatTrim$ = Left$(ChatText$, 11)
For Z& = 1 To 11
If Mid$(ChatTrim$, Z&, 1) = ":" Then
sn& = Left$(ChatTrim$, Z& - 1)
End If
Next Z&
aol_snfromlastchatline = sn&
End Function





Sub api_showwindow(hWnd)
Call ShowWindow(hWnd, SW_SHOW)
End Sub



Sub file_sethidden(TheFile As String)
Dim File As String
File$ = Dir(TheFile$)
If File$ <> "" Then SetAttr TheFile$, vbHidden
End Sub

Function text_replace(Text As Textbox, Find As String, Changeto As String)
Dim X As Long, char As String, chars As String
If InStr(Text, Find$) = 0 Then
text_replace = Text
Exit Function
End If
For X = 1 To Len(Text)
char$ = Mid(Text, X, 1)
chars$ = chars$ & char$
If char$ = Find$ Then
chars$ = Mid(chars$, 1, Len(chars$) - 1) + Changeto$
End If
Next X
text_replace = chars$
End Function

Sub file_create(File As String)
Dim Free As Long
Free& = FreeFile
Open File$ For Random As Free&
Close Free&
End Sub

Function file_input(File As String) As Long
Dim Free As Long, I As Long, X As Long
Free& = FreeFile
Open File$ For Input As Free&
I& = FileLen(File$)
X& = Input(I&, Free&)
Close Free&
file_input& = X&
End Function

Sub file_setreadonly(TheFile As String)
Dim File As String
File$ = Dir(TheFile$)
If File$ <> "" Then SetAttr TheFile$, vbReadOnly
End Sub

Function file_preload() As Boolean
file_preload = False
If (App.PrevInstance = True) Then
file_preload = True
End If
End Function
Sub other_delay(Interval)
Dim Current
Current = Timer
Do While Timer - Current < Val(Interval)
DoEvents
Loop
End Sub

Function api_getclass(WinHandle As Long) As String
Dim FixedString As String
FixedString$ = String(250, 0)
Call GetClassName(WinHandle&, FixedString$, 250)
api_getclass$ = FixedString$
End Function


Function spy_handle() As Long
Dim CursorPos As POINTAPI
Call GetCursorPos(CursorPos)
spy_handle& = WindowFromPointXY(CursorPos.X, CursorPos.Y)
End Function



Function spy_id() As String
spy_id$ = GetWindowLong(spy_handle&, (-12))
End Function


Function spy_parent() As Long
spy_parent& = GetParent(spy_handle&)
End Function


Function spy_parentclass() As String
spy_parentclass$ = api_getclass(spy_parent&)
End Function


Function spy_parentid() As String
spy_parentid$ = GetWindowLong(spy_parent&, (-12))
End Function

Function spy_parentstyle() As String
spy_parentstyle$ = GetWindowLong(spy_parent&, (-16))
End Function


Function spy_parenttext() As String
spy_parenttext$ = api_gettext(spy_parent&)
End Function


Function api_gettext(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
TextLength& = PostMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
api_gettext$ = buffer$
End Function



Function spy_style() As String
spy_style$ = GetWindowLong(spy_handle&, (-16))
End Function


Function spy_text() As String
spy_text$ = api_gettext(spy_handle&)
End Function



Function cd_check()
Dim s As String * 30
Call MCISendString("status cd media present", s, Len(s), 0)
cd_check = s
End Function


Function cd_convsec()
Dim T As String * 50
Dim I, RSec, Seconds, FSeconds, MTS As Integer
Dim ms As String
T = cd_position
MTS = Mid(T, 1, 2)
FSeconds = MTS * 60
RSec = Mid(T, 4, 2)
Seconds = (FSeconds + RSec)
cd_convsec = Seconds
End Function

Function cd_currtrack()
Dim s As String * 50
Call MCISendString("status cd current track", s, Len(s), 0)
cd_currtrack = s
End Function

Function cd_fastforward(Spd)
Dim s As String * 40
cd_setformat_millisec
Call MCISendString("status cd position wait", s, Len(s), 0)
If cd_isplaying = True Then
Call MCISendString("play cd from " & CStr(CLng(s) + Spd), 0, 0, 0)
Else
Call MCISendString("seek cd to " & CStr(CLng(s) + Spd), 0, 0, 0)
End If
cd_setformat_tmsf
End Function
Function cd_setformat_tmsf()
Call MCISendString("set cd time format tmsf wait", 0, 0, 0)
End Function


Function cd_setformat_millisec()
Call MCISendString("set cd time format milliseconds", 0, 0, 0)
End Function


Function cd_getnumtracks()
Dim s As String * 30
Call MCISendString("status cd number of tracks wait", s, Len(s), 0)
cd_getnumtracks = CInt(Mid$(s, 1, 2))
End Function



Function cd_playing() As Boolean
cd_playing = False
Dim s As String * 50
Call MCISendString("status cd mode", s, Len(s), 0)
If Mid$(s, 1, 7) = "playing" Then
cd_playing = True
Else
cd_playing = False
End If
End Function

Function cd_length()
Dim s As String * 30
Call MCISendString("status cd length wait", s, Len(s), 0)
cd_length = s
End Function

Function cd_position()
Dim mm, Sec, Min, Track As Integer
Dim s As String * 30
Call MCISendString("status cd position", s, Len(s), 0)
Track = CInt(Mid$(s, 1, 2))
Min = CInt(Mid$(s, 4, 2))
Sec = CInt(Mid$(s, 7, 2))
cd_position = "Track[" & Track & "] Min[" & Min & "] Sec[" & Sec & "]"
End Function

Function cd_settrack(Track)
Call MCISendString("seek cd to " & Str(Track), 0, 0, 0)
End Function

Sub cd_unload()
Call MCISendString("close all", 0, 0, 0)
End Sub


Sub form_unload()
Dim OfTheseForms As Form
For Each OfTheseForms In Forms
Unload OfTheseForms
Set OfTheseForms = Nothing
Next OfTheseForms
End Sub


Sub form_ontop(Frm As Form)
Call SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub



Sub list_load(Path As String, Lst As Control)
Dim Strin As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, Strin$
DoEvents
Lst.AddItem Strin$
Wend
Close #1
Exit Sub
End Sub
Sub text_load(Path As String, Text As Textbox)
Dim Strin As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, Strin$
DoEvents
Text = Strin$
Wend
Close #1
Exit Sub
End Sub

Sub list_save(Path As String, Lst As Control)
Dim SaveList As Long
On Error Resume Next
Open Path$ For Output As #1
For SaveList& = 0 To Lst.listcount - 1
Print #1, Lst.List(SaveList&)
Next SaveList&
Close #1
End Sub

Sub text_save(Path As String, Text As Textbox)
On Error Resume Next
Open Path$ For Output As #1
Print #1, Text
Close 1
End Sub


Sub list_fontsizes(Lst As Control)
Dim X As Long
For X& = 6 To 32 Step 2
Lst.AddItem X&
Next X&
Lst.AddItem "48"
Lst.AddItem "36"
Lst.AddItem "72"
End Sub

Sub list_ascii(Lst As Control)
Dim X As Long
For X& = 33 To 223
Lst.AddItem Chr(X&)
Next X&
End Sub



Function text_reverse(Text As String)
On Error Resume Next
Dim RT$
Dim Words As Long
For Words& = Len(Text$) To 1 Step -1
RT$ = RT$ & Mid(Text$, Words&, 1)
Next Words&
text_reverse = RT$
End Function












