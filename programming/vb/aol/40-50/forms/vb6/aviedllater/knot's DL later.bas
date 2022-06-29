Attribute VB_Name = "knot2"
' knot2.bas by knot for auto download later
' this works for me... if it doesn't for you
' make the pauses longer i made them so they can work for me

' if you use this bas please put me in the greets or some shit
' all subs made by knot except 2 or 3
' which were made by chichis

' don't be gay and steal the codes, learn from this!
' Theirs nothing lamer, then stealing codes

' contact info for questions, comments, or shit
' Aim = itsknot, emails: knotfx@aol.com, knot9@aol.com
' private room: aviempire
' for warez/movies/porn and other shit

Option Explicit

Declare Function GetDriveType Lib "Kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetLogicalDriveStrings Lib "Kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTickCount Lib "Kernel32" () As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function getnextwindow Lib "User" Alias "GetNextWindow" (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

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
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112



Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Const EM_GETLINE = &HC4
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1


Public Const LB_GETTEXTLEN = &H18A
Public Const LB_FINDSTRINGEXACT = &H1A2


Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetDiskFreeSpaceEx2 Lib "Kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailable As Long, lpTotalNumberOfBytes As Long, lpTotalNumberOfFreeBytes As Long) As Long
Private Const DRW_CAPTIONMINIMIZE = &H3
Dim FB, BT, FBT As Currency
Dim drivesize As String
Const Gigabyte = 1073741824
Const Megabyte = 1048576
Dim retval As Long

Private Declare Function GetDiskFreeSpace_FAT32 _
    Lib "Kernel32" Alias "GetDiskFreeSpaceExA" _
    (ByVal lpRootPathName As String, _
    FreeBytesToCaller As Currency, BytesTotal _
    As Currency, FreeBytesTotal As Currency) _
    As Long



            
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1


Public Const LB_GETITEMDATA = &H199
Public Const LB_SETSEL = &H185

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

Public Const VK_UP = &H26

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type







Public Sub SendLater(People As String, Subject As String, Tag As String, FilePath As String)
Dim AOLFrame As Long
Dim aoltoolbar As Long
Dim AOLIcon As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim RICHCNTL As Long
Dim AOLModal As Long
Dim AOLTree As Long
Dim X As Long
Dim editx As Long
Dim Button As Long

AOLFrame = FindWindow("aol frame25", vbNullString)

If AOLFrame = 0 Then
   MsgBox "Error: Cannot find window"
   Exit Sub
End If

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   aoltoolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
   aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
   AOLIcon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(aoltoolbar, AOLIcon, "_aol_icon", vbNullString)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   aoltoolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
   aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
   AOLIcon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(aoltoolbar, AOLIcon, "_aol_icon", vbNullString)
   Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
Loop Until AOLChild <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
Loop Until AOLEdit <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
Loop Until AOLEdit <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Loop Until RICHCNTL <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
   Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, People)
Loop Until AOLEdit <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
   AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
   Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, Subject)
Loop Until AOLEdit <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
   Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Tag)
Loop Until RICHCNTL <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
Loop Until AOLModal <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
   AOLTree = FindWindowEx(AOLModal, 0&, "_aol_tree", vbNullString)
Loop Until AOLTree <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, 0&, "_aol_icon", vbNullString)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, AOLIcon, "_aol_icon", vbNullString)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, 0&, "_aol_icon", vbNullString)
   Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Loop Until AOLIcon <> 0

Pause (0.7)

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   X = FindWindow("#32770", vbNullString)
Loop Until X <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   X = FindWindow("#32770", vbNullString)
   editx = FindWindowEx(X, 0&, "edit", vbNullString)
Loop Until editx <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   X = FindWindow("#32770", vbNullString)
   Button = FindWindowEx(X, 0&, "button", vbNullString)
   Button = FindWindowEx(X, Button, "button", vbNullString)
Loop Until Button <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   X = FindWindow("#32770", vbNullString)
   editx = FindWindowEx(X, 0&, "edit", vbNullString)
   Call SendMessageByString(editx, WM_SETTEXT, 0&, FilePath)
Loop Until editx <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   X = FindWindow("#32770", vbNullString)
   Button = FindWindowEx(X, 0&, "button", vbNullString)
   Button = FindWindowEx(X, Button, "button", vbNullString)
   Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
Loop Until Button <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   AOLModal = FindWindow("_aol_modal", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLModal, AOLIcon, "_aol_icon", vbNullString)
   Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Loop Until AOLIcon <> 0

Do
   DoEvents
   AOLFrame = FindWindow("aol frame25", vbNullString)
   MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
   AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
   Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Loop Until AOLIcon <> 0
End Sub
Public Sub MailOpenNew()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
    Dim WinVis As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
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

Public Function MailGetSubject(Index As Long) As String
  ' made by chichis
    Dim Mailbox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String
    Dim Count As Long
    Mailbox = FindMailBox
    If Mailbox = 0& Then Exit Function
    TabControl = FindWindowEx(Mailbox, 0&, "_AOL_TabControl", vbNullString)
    TabPage = FindWindowEx(TabControl, 0&, "_AOL_TabPage", vbNullString)
    mTree = FindWindowEx(TabPage, 0&, "_AOL_Tree", vbNullString)
    Count = SendMessage(mTree, LB_GETCOUNT, 0&, 0&)
    If Count = 0 Or Index > Count - 1 Or Index < 0 Then Exit Function
    sLength = SendMessage(mTree, LB_GETTEXTLEN, Index, 0&)
    MyString = String(sLength + 1, 0)
    SendMessageByString mTree, LB_GETTEXT, Index, MyString
    Spot = InStr(MyString, Chr(9))
    Spot = InStr(Spot + 1, MyString, Chr(9))
    MyString = Mid(MyString, Spot + 1)
    MailGetSubject = Left(MyString, Len(MyString) - 1)
End Function

Public Sub MailOpenEmailNew(Index As Long)
    Dim Mailbox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    Mailbox& = FindMailBox&
    If Mailbox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(Mailbox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Public Sub SendChat(chat As String)
    Dim Room As Long, RICHCNTL As Long, X As Integer
    Dim OldChatText As String, Timer As Long, TempSend As String
    Room = FindRoom
    RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    RICHCNTL = FindWindowEx(Room, RICHCNTL, "RICHCNTL", vbNullString)
    TempSend = chat
    WindowHide RICHCNTL
    OldChatText = GetSendText
    If Not (OldChatText = Empty) Then
        Do Until GetSendText = ""
            DoEvents
            SetText RICHCNTL, ""
            If IsWindow(RICHCNTL) = False Then Exit Sub
        Loop
    End If
    SetText RICHCNTL, chat
    Timer = GetTickCount / 1000
    Do Until Left(GetSendText, Len(chat)) = chat
        DoEvents
        If GetTickCount / 1000 - Timer <= 0.1 Then
            SetText RICHCNTL, chat
            Timer = GetTickCount / 1000
        End If
        If IsWindow(RICHCNTL) = False Then Exit Sub
    Loop
    SendChar RICHCNTL, 13
    Timer = GetTickCount / 1000
    Do Until GetSendText = ""
        DoEvents
        If GetTickCount / 1000 - Timer <= 0.1 Then
            SendChar RICHCNTL, 13
            Timer = GetTickCount / 1000
        End If
        If IsWindow(RICHCNTL) = False Then Exit Sub
    Loop
    If Not OldChatText = Empty Then
        SetText RICHCNTL, OldChatText
    End If
     'RICHCNTL
End Sub
Public Sub SendChar(hwnd As Long, char As Byte)
    If IsWindow(hwnd) = False Then Exit Sub
    SendMessageByNum hwnd, WM_CHAR, char, 0
End Sub

Public Sub SetText(hwnd As Long, Text As String)
    SendMessageByString hwnd, WM_SETTEXT, 0&, Text
End Sub

Public Function GetSendText() As String
    Dim LineIndex As Long, LineLength As Long, X As String
    Dim result As Long, Chatroom As Long, RICHCNTL As Long
    Chatroom = FindRoom
    RICHCNTL = FindWindowEx(Chatroom, 0&, "RICHCNTL", vbNullString)
    RICHCNTL = FindWindowEx(Chatroom, RICHCNTL, "RICHCNTL", vbNullString)

        GetSendText = GetText(RICHCNTL)
   
End Function
Public Function GetText(child As Long) As String
    Dim TrimSpace As String, GetString As Long, gettrim
    gettrim = SendMessageByNum(child, 14, 0&, 0&)
    TrimSpace = Space(gettrim)
    GetString = SendMessageByString(child, 13, gettrim + 1, TrimSpace)
    GetText = TrimSpace
End Function

Public Function MKI(X As Integer) As String
    Dim Y As Long
    Y = CLng(X) And &HFFFF&
    MKI = Chr(Y And &HFF&) & Chr((Y And &HFF00&) \ &H100&)
End Function

Public Function GetAolVer() As Boolean
    Dim lngAOL As Long, lngTool As Long, lngToolbar As Long
    Dim lngCombo As Long, lngKW As Long, lngSearch As Long
    lngAOL = FindWindow("AOL Frame25", vbNullString)
    lngTool = FindWindowEx(lngAOL, 0, "AOL Toolbar", vbNullString)
    lngToolbar = FindWindowEx(lngTool, 0, "_AOL_Toolbar", vbNullString)
    lngSearch = FindWindowEx(lngToolbar, 0, "_AOL_Edit", vbNullString)
    lngSearch = FindWindowEx(lngToolbar, lngSearch, "_AOL_Edit", vbNullString)
    lngSearch = FindWindowEx(lngToolbar, lngSearch, "_AOL_Edit", vbNullString)
    lngSearch = FindWindowEx(lngToolbar, lngSearch, "_AOL_Edit", vbNullString)
    lngSearch = FindWindowEx(lngToolbar, lngSearch, "_AOL_Edit", vbNullString)
    lngCombo = FindWindowEx(lngToolbar, 0, "_AOL_Combobox", vbNullString)
    lngKW = FindWindowEx(lngCombo, 0, "Edit", vbNullString)
    If lngSearch > 0 Then
    GetAolVer = False
    ElseIf lngSearch = 0 And lngCombo > 0 And lngKW > 0 Then
    GetAolVer = True
    ElseIf lngCombo = 0 And lngKW = 0 And lngToolbar = 0 And lngTool > 0 Then
    GetAolVer = True
    Else
    GetAolVer = False
    End If
End Function
Public Function FindDlLater() As Boolean
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
If AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString) Then
FindDlLater = True
Else
FindDlLater = False
End If

End Function


Public Function FindMailBox() As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim TabControl As Long, TabPage As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindMailBox& = 0&
End Function


Public Sub DlLater()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, AOLModal As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Pause 0.55
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLModal = FindWindow("_aol_modal", vbNullString)
AOLIcon = FindWindowEx(AOLModal, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub Keyword(KW As String)
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim combo As Long, editwin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    editwin& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(editwin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub PrivateRoom(Room As String)
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub

Public Sub FormOnTop(FormName As Form, Ontop As Boolean)
    If Ontop = True Then
        SetWindowPos FormName.hwnd, -1, 0&, 0&, 0&, 0&, &H1 Or &H2
    Else
        SetWindowPos FormName.hwnd, -2, 0&, 0&, 0&, 0&, &H1 Or &H2
    End If
End Sub

Public Sub Click(icon As Long)
    SendMessage icon, WM_LBUTTONDOWN, 0&, 0&
    SendMessage icon, WM_KEYDOWN, VK_SPACE, 0&
    SendMessage icon, WM_LBUTTONUP, 0&, 0&
    SendMessage icon, WM_KEYUP, VK_SPACE, 0&
End Sub
Public Sub DoubleClick(icon As Long)
    SendMessage icon, WM_LBUTTONDBLCLK, 0&, 0&
End Sub
Public Function FindMail(Subject As String) As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim Caption As String
    aol = FindWindow("AOL Frame25", vbNullString)
    MDI = FindWindowEx(aol, 0, "MDIClient", vbNullString)
    child = FindWindowEx(MDI, 0, "AOL Child", vbNullString)
    Caption = Left(Replace(GetCaption(child), " ", ""), 40)
    If Left(Replace(Subject, " ", ""), 40) = Caption Then
        FindMail = child
        Exit Function
    Else
        Do
            child = FindWindowEx(MDI, child, "AOL Child", vbNullString)
            Caption = Left(Replace(GetCaption(child), " ", ""), 40)
            If Left(Replace(Subject, " ", ""), 40) = Caption Then
                FindMail = child
                Exit Function
            End If
        Loop Until child = 0
    End If
    FindMail = 0
End Function

Public Function FindRoom() As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    aol = FindWindow("AOL Frame25", vbNullString)
    MDI = FindWindowEx(aol, 0, "MDIClient", vbNullString)
    child = FindWindowEx(MDI, 0, "AOL Child", vbNullString)
    Rich = FindWindowEx(child, 0, "RICHCNTL", vbNullString)
    AOLList = FindWindowEx(child, 0, "_AOL_Listbox", vbNullString)
    AOLIcon = FindWindowEx(child, 0, "_AOL_Icon", vbNullString)
    AOLStatic = FindWindowEx(child, 0, "_AOL_Static", vbNullString)
    If Rich <> 0 And AOLList <> 0 And AOLIcon <> 0 And AOLStatic <> 0 Then
        FindRoom = child
        Exit Function
    Else
        Do
            child = FindWindowEx(MDI, child, "AOL Child", vbNullString)
            Rich = FindWindowEx(child, 0, "RICHCNTL", vbNullString)
            AOLList = FindWindowEx(child, 0, "_AOL_Listbox", vbNullString)
            AOLIcon = FindWindowEx(child, 0, "_AOL_Icon", vbNullString)
            AOLStatic = FindWindowEx(child, 0, "_AOL_Static", vbNullString)
            If Rich <> 0 And AOLList <> 0 And AOLIcon <> 0 And AOLStatic <> 0 Then
                FindRoom = child
                Exit Function
            End If
        Loop Until child = 0
    End If
    FindRoom = child
End Function
Public Function GetCaption(hwnd As Long) As String
    Dim length As Long
    Dim Title As String
    length = GetWindowTextLength(hwnd)
    Title = String(length, 0)
    GetWindowText hwnd, Title, (length + 1)
    GetCaption = Title
End Function
Public Function ListBoxCheckDup(List As ListBox, Query As String) As Boolean
' made by chichis :D
    If Query = "" Then Exit Function
    If Not TypeOf List Is ListBox Then Exit Function
    Dim X As Long
    
    X = SendMessageByString(List.hwnd, LB_FINDSTRINGEXACT, 0&, Query)
    ListBoxCheckDup = IIf(X > -1, True, False)
End Function
Public Function MailCountNew() As Long
    Dim Mailbox As Long, AOLChild As Long, Count As Long
    Dim AOLTabControl As Long, AOLTabPage As Long, AOLTree As Long
    Mailbox = FindMailBox
    AOLTabControl = FindWindowEx(Mailbox, 0&, "_AOL_TabControl", vbNullString)
    AOLTabPage = FindWindowEx(AOLTabControl, 0&, "_AOL_TabPage", vbNullString)
    AOLTree = FindWindowEx(AOLTabPage, 0&, "_AOL_Tree", vbNullString)
    Count = SendMessage(AOLTree, LB_GETCOUNT, 0, 0)
    MailCountNew = Count
End Function

Public Function MailCountSubject(Subject As String) As Long
    Dim Mailbox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, i As Long, sLength As Long
    Dim MyString As String, Spot As Long, Count As Long
    Mailbox = FindMailBox
    TabControl = FindWindowEx(Mailbox, 0&, "_AOL_TabControl", vbNullString)
    TabPage = FindWindowEx(TabControl, 0&, "_AOL_TabPage", vbNullString)
    mTree = FindWindowEx(TabPage, 0&, "_AOL_Tree", vbNullString)
    For i = 0 To MailCountNew - 1
        sLength = SendMessage(mTree, LB_GETTEXTLEN, i, 0&)
        MyString = String(sLength + 1, 0)
        SendMessageByString mTree, LB_GETTEXT, i, MyString
        Spot = InStr(MyString, Chr(9))
        Spot = InStr(Spot + 1, MyString, Chr(9))
        MyString = Mid(MyString, Spot + 1)
        If InStr(LCase(MyString), LCase(Subject)) > 0 Then Count = Count + 1
    Next
    MailCountSubject = Count
End Function
Public Function Percent(found&, Total&) As String
' makes it so its like 45.53%
Percent$ = (found& / Total&) * 100
Percent$ = Round(Percent$, 2)
Percent$ = Percent$ + "%"
End Function
Public Function percent2(found&, Total&) As String
' makes it so its like 45%
percent2$ = (found& / Total&) * 100
percent2$ = Round(percent2$, 0)
percent2$ = percent2$ + "%"
End Function


Public Function percent3(found&, Total&, rnd&) As String
' makes it so its like 45.rnd&
If IsNumeric(rnd&) = False Then Exit Function
percent3$ = (found& / Total&) * 100
percent3$ = Round(percent3$, rnd&)
percent3$ = percent3$ + "%"
End Function


Public Sub WindowHide(hwnd As Long)
    ShowWindow hwnd&, SW_HIDE
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

