Attribute VB_Name = "DaCrazyOne70"
'            ,,,                       ,,,                     ,,,         ,,
'    ,,;;;;;;;;;                      ;;;;;,                  ;;;;,      ,;;;
'  ,;;;;;;;;;;;;,                    ;;;;;;;,                  ;;;;,   ,;;;´
' ;;;;;´     ;;;´                   ;;;; ´;;;,                  ´;;;;,;;;;
',;;;´              ,, ,,,         ;;;;   ;;;;      ,,,,,,,,     ´;;;;;;´
';;;;              ;;;´´´;;,      ;;;;     ;;;    ´;´´´;;;;´      ´;;;;´
';;;;               ;;,  ;;;     ,;;;;;;;;;;;;;       ,;;´         ;;;;,
'´;;;;        ,,,   ;;;,;;´    ,,;;;;;;;´´;;;;´     ,;;´          ,;;;´
' ´;;;;,    ,,;;;  ,;;;;;,    ;;;;;;´     ;;;;     ,;;´           ;;;;
'  ´;;;;;;;;;;;;´   ;;;  ;;,   ´;;;´      ;;;;    ,;;,;;;;;,     ,;;;;
'    ´´;;;;;;´´     ;;;   ´´    ;;;       ´;;´    ;;´´    ´      ´;;;
'                        The .bas created by DaCrazyOne
'                          Email:ThatsMrPsP2U@aol.com
'
'Pause function and some decs taken from dos32.bas
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const ABS_ALWAYSONTOP = &H2
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const EM_GETLINE = &HC4
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINECOUNT = &HBA

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

Public Const SWP_NOMOVE = &H2
Public Const SW_SHOWNOACTIVATE = 4
Public Const SWP_HIDEWINDOW = &H80

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT


Global Const SWP_NOSIZE = 1
Public Const SW_RESTORE = 9

Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Const LB_SETITEMDATA = &H19A
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMDATA = &H150

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3

Public Const VK_TAB = &H9
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
Public Const WM_MOVE = &HF012 '
Public Const WM_SETTEXT = &HC '
Public Const WM_SYSCOMMAND = &H112 '
Public Const ENTER_KEY = 13 '

Private Const PROCESS_READ = &H10
Private Const RIGHTS_REQUIRED = &HF0000


Public Const SC_CLOSE = &HF060&
Public Const MF_BYCOMMAND = &H0&


Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETLBTEXT = &H148


Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const SW_NORMAL = 1

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Sub OpenMailbox()
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub OpenMyCalender()
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub SendChat(text As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, richcntl As Long, AOLEdit As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, text$)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub imoff()
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
Dim MDIClient As Long, AOLChild As Long, AOLEdit As Long, richcntl As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (0.6)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, "$IM_OFF")
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, "7.0")
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
Call SendMessageLong(AOLChild, WM_CLOSE, 0&, 0&)
Pause (0.6)
Call Close_msgbox
End Sub

Public Sub imon()
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
Dim MDIClient As Long, AOLChild As Long, AOLEdit As Long, richcntl As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (0.6)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, "$IM_ON")
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, "7.0")
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
Call SendMessageLong(AOLChild, WM_CLOSE, 0&, 0&)
Pause (0.6)
Call Close_msgbox
End Sub
'close the America Online message boxes
Public Sub Close_msgbox()
Dim x As Long
x = FindWindow("#32770", vbNullString)
If FindWindow("#32770", "America Online") <> 0& Then
Call SendMessage(FindWindow("#32770", "America Online"), WM_CLOSE, 0&, 0&)
End If

End Sub

Public Sub CloseWindow()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
Call SendMessageLong(AOLChild, WM_CLOSE, 0&, 0&)
End Sub
'clicks the setup button on the buddy list window
Public Sub Click_buddysetup()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)


End Sub
'clicks the buddy preferences button on the buddylist setup window
Public Sub Click_buddypref()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)

End Sub
'sets the buddlylist preferences to block all

Public Sub Ghost_on()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Call Pause(1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
aoltabcontrol = FindWindowEx(AOLChild, 0&, "_aol_tabcontrol", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, 0&, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, aoltabpage, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, aoltabpage, "_aol_tabpage", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, 0&, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, aolradiobox, "_aol_radiobox", vbNullString)
Call SendMessageLong(aolradiobox, BM_SETCHECK, True, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'Sets the buddylist preferences to allow all

Public Sub Ghost_off()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Call Pause(1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
aoltabcontrol = FindWindowEx(AOLChild, 0&, "_aol_tabcontrol", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, 0&, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, aoltabpage, "_aol_tabpage", vbNullString)
aoltabpage = FindWindowEx(aoltabcontrol, aoltabpage, "_aol_tabpage", vbNullString)
aolradiobox = FindWindowEx(aoltabpage, 0&, "_aol_radiobox", vbNullString)
Call SendMessageLong(aolradiobox, BM_SETCHECK, True, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (1)
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'click the Save button on the buddylist preferences window
Public Sub BuddyPrefSave()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'click the apply button on the buddylist preferences window
Public Sub BuddyPrefApply()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'clicks the return to buddylist button on the buddly list setup window
Public Sub ReturntoBuddy()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'Opens a new write mail window
Public Sub OpenWriteMail()
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)

End Sub
'sets text to the send to of a write mail window
Public Sub WriteMail_To(text As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLEdit As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, text$)
End Sub
'sets text to the body of the mail
Public Sub WriteMailBody(text As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim richcntl As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, text$)
End Sub

'sets the text of the subject line in a new mail
Public Sub MailWriteSubject(text As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLEdit As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, text$)
End Sub
'click the send now button
Public Sub ClickMailSendnow()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
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
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)


End Sub
'sends text to the aol search engine in the tool bar and clicks the button
Public Sub AOLSearch(Txt As String)
Dim AOLFrame As Long, AOLToolbar As Long, AOLEdit As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, 0&, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, Txt$)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub OpenMailNew() 'Opens mail to the new tab
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim aoltabcontrol As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLChild = FindWindowEx(MDIClient, AOLChild, "aol child", vbNullString)
aoltabcontrol = FindWindowEx(AOLChild, 0&, "_aol_tabcontrol", vbNullString)

End Sub
'used for sending Instant Messages
Public Sub IMSend(who As String, Msg As String)
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
Dim MDIClient As Long, AOLChild As Long, AOLEdit As Long, richcntl As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
Pause (0.6)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, who$)

MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Msg$)
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub BoldTextChat() 'sets chatroom to bold text
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub ItalicsTextChat() 'sets chat room to italics
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)


End Sub

Public Sub UnderlineTextChat() 'sets chat room to unerlinetext
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'use this code to send a message to a Instant Message that is already open
Public Sub SendImBody(Txt As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim richcntl As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Txt$)
Pause (0.3)
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'sets the im text to BOLD
Public Sub IMBold()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)


End Sub
'sets the im text to Italics
Public Sub IMItalics()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
'sets the im text to underline
Public Sub IMUnderline()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
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
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)


End Sub
'I got this code from Microsoft database
Public Sub disableX()
'Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    'Public Const SC_CLOSE = &HF060&
    'Public Const MF_BYCOMMAND = &H0&
'Place the Following in the Form Load Field of your APP
'the above Decs in a module
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

End Sub
Public Function OpenURL(ByVal URL As String) As Long
 OpenURL = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function


Public Sub OpenFilingCabinet()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Filing")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub OpenPictureGallery()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Picture")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub CapturePicture()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Capture")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub AddTopWindowToFavorites()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Favorites")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub Cascadewindows()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Cascade")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub SwitchScreenNames()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Switch")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub openaboutamericaonline()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("About")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub Signoff()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Sign")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub

'gets all the text from a chat room and adds it to a richtextbox
Public Function getroomtext(Txt As RichTextBox)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim richcntlreadonly As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntlreadonly = FindWindowEx(AOLChild, 0&, "richcntlreadonly", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(richcntlreadonly, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(richcntlreadonly, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
Txt = TheText
End Function

'opens keywords using the the AOL combobox in the tool bar
Public Sub Keyword(KWord As String)
    Dim AOLFrame As Long, AOLToolbar As Long, AOLCombobox As Long
    Dim Editx As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
    AOLCombobox = FindWindowEx(AOLToolbar, 0&, "_aol_combobox", vbNullString)
    Editx = FindWindowEx(AOLCombobox, 0&, "Edit", vbNullString)
    Call SendMessageByString(Editx, WM_SETTEXT, 0&, KWord$)
    Call SendMessageLong(Editx, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Editx, WM_CHAR, VK_RETURN, 0&)
End Sub
'use of this works as follows icon number, then the underlined letter of the item in that menu you want to open
'settings = icon number 9
'people = icon number 3
'services = icon number 6
'favorites = icon number 11
'mail = icon number 0
'example Call clickToolbar("0", "M")
Public Sub clickToolbar(IconNumber&, letter$)

Dim AOLFrame As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter, 0&)
End Sub

'use this one if the menu item you want has a second menu
'for example Read mail = New Mail,Old Mail,Sent Mail

'Call clickToolbar2("0", "R", "N")<~~ so you would use this to open the new mail
'Call clickToolbar2("0", "R", "O")<~~ this would open Old Mail
'Call clickToolbar2("0", "R", "S")<~~ this would open Sent Mail
Public Sub clickToolbar2(IconNumber&, letter$, letter2$)

Dim AOLFrame As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
letter2 = Asc(letter2)
Call PostMessage(menu, WM_CHAR, letter, 0&)
Call PostMessage(menu, WM_CHAR, letter2, 0&)
End Sub



Public Sub OpenDownLoadManager()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Dim TheWindow As Long
Dim aMenu As Long
Dim mCount As Long
Dim LookFor As Long
Dim sMenu As Long
Dim sCount As Long
Dim LookSub As Long
Dim sID As Long
Dim sString As String

TheWindow = AOLFrame
aMenu& = GetMenu(TheWindow)
mCount& = GetMenuItemCount(aMenu&)
For LookFor& = 0& To mCount& - 1
    sMenu& = GetSubMenu(aMenu&, LookFor&)
    sCount& = GetMenuItemCount(sMenu&)
    For LookSub& = 0 To sCount& - 1
        sID& = GetMenuItemID(sMenu&, LookSub&)
        sString$ = String$(100, " ")
        Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
        If InStr(LCase(sString$), LCase("Download")) Then
            Call SendMessageLong(TheWindow, WM_COMMAND, sID&, 0&)
            Exit Sub
        End If
    Next LookSub&
Next LookFor&
End Sub
Public Sub HideAOL()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Call ShowWindow(AOLFrame, SW_HIDE)
End Sub
Public Sub ShowAOL()
Dim AOLFrame As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
Call ShowWindow(AOLFrame, SW_SHOW)
End Sub
'gets all the text from a chat room and adds it to a richtextbox
Public Function getroomtext(Txt As RichTextBox)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim richcntlreadonly As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
richcntlreadonly = FindWindowEx(AOLChild, 0&, "richcntlreadonly", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(richcntlreadonly, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(richcntlreadonly, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
Txt = TheText
End Function
'Taken From dos32
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
