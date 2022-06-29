Attribute VB_Name = "PATorJKaimsysop"
Option Explicit

'PAT or JK's AIM Sysop Bas version 1.0

'Ok, this is a little bas that I wrote after
'making aim flash 1.0, and yes it's written
'100% by me, unless u count subs like timeout
'which everyone knows.

'Cool trick:
'For some reason aim sysop chat rooms remove
'certain html tags in the chat window when u try
'and set the text (tags like "<br>" and "<body>")
'anyways 2 get around this just place html tags
'inside html tags, example: "<<p>br>", that way
'when the text is set it removes "<p>" and leaves
'the tag "<br>"

'If u find any errors or want 2 contact me my
'e-mail is: "patorjk@aol.com" and my webpage is:

'http://members.xoom.com/thepatmaster/index.htm


Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
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

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Sub AddRoom2List(List1 As ListBox)
'Adds the room 2 a list
Dim ab As String, c As Long, d As Long, Person$
Dim x As Integer, pat$, e%
Dim Room&
pat$ = GetUserSN()

Room& = FindWindow("AIM_ChatWnd", vbNullString)
If Room& = 0& Then Exit Sub
c = FindWindowEx(Room&, 0&, "_Oscar_Tree", vbNullString)
d = SendMessage(c, LB_GETCOUNT, 0, 0)
For x = 0 To d - 1
Person$ = String(255, 0)
ab = SendMessageByString(c, LB_GETTEXT, x, Person$)
If Mid$(Trim$(Person$), 1, Len(pat$)) <> Trim$(pat$) Then
 e% = InStr(Person$, Chr$(0))
 If e% <> 0 Then
   List1.AddItem Mid$(Person$, 1, e% - 1)
 Else
   List1.AddItem Person$
 End If
End If
Next x
End Sub

Sub AIMFlash()
'Cool room flash
SendChat "<<br>body bgcolor=""#ff0000""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#0000ff""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#ffff00""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#00ff00""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#000000""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#008888""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.3
SendChat "<<br>body bgcolor=""#fffffe""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
End Sub

Sub AIMFlash2()
'Cool room flash
SendChat "<<br>body bgcolor=""#000000""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#888888""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#fffffe""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#888888""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#000000""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#888888""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
TimeOut 0.1
SendChat "<<br>body bgcolor=""#fffffe""><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
End Sub

Function ChangeCap(Wnd&, Caption$)
'This change the caption of a window
Dim cc&
cc& = SendMessageByString(Wnd&, WM_SETTEXT, 0, Caption$)
End Function

Sub ClearChat()
'Clears the chat, and it's real
SendChat "<<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br><<p>br>"
End Sub

Sub ClickIcon(Icon&)
Dim Click&
Click& = SendMessage(Icon&, WM_LBUTTONDOWN, 0, 0&)
Click& = SendMessage(Icon&, WM_LBUTTONUP, 0, 0&)
End Sub

Function CountLines(pat As String) As Integer
Dim x%
x% = 1
Do While InStr(pat, Chr$(13))
x% = x% + 1
pat = Mid$(pat, InStr(pat, Chr$(13)) + 2)
Loop
CountLines = x%
End Function

Public Function Text_Encrypt(Text As String) As String
Dim x$, X2$, i%, Boo%, Boo2$, PATorJK$
x$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡²³ÀÁÂÃÄÅÒÓÔÕÖÙÛÜàáâãäåØ¶§Ú¥"
X2$ = " ¿¡@#$%^&*()_+|01²³456789ÀbÁdÂÃghÄjklmÅÒÓqÔÕÖÙvwÛÜz.,-~AàáâãFGHäJKåMNØ¶QR§TÚVWX¥Z?!23acefinoprstuxyBCDEILOPSUY"
For i% = 1 To Len(Text)
Boo% = InStr(x$, Mid(Text, i%, 1))
If Not Boo% = 0 Then
Boo2$ = Mid(X2$, Boo%, 1)
PATorJK$ = PATorJK$ + Boo2$
Else
Boo2 = Mid(Text, i, 1)
PATorJK = PATorJK + Boo2
End If
Next
Encrypt = PATorJK$
'2 encrypt a word just put something like this:
'Text1.Text = Text_Encrypt(Text1)
'And then 2 unencrypt the word just call the function again:
'Text1.Text = Text_Encrypt(Text1)
End Function

Public Function GetCaption(hwnd As Long) As String
Dim Length%, Title$, a%
Length% = GetWindowTextLength(hwnd)
Title$ = String$(Length%, 0)
a% = GetWindowText(hwnd, Title$, (Length% + 1))
GetCaption = Title$
End Function

Public Function GetUserSN()
'This will get your SN
Dim oscarbuddylistwin&, sn$, Length%, Title$, a%

oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Length% = GetWindowTextLength(oscarbuddylistwin&)
sn$ = String$(Length%, 0)
a% = GetWindowText(oscarbuddylistwin&, sn$, (Length% + 1))
If InStr(sn$, "'s Buddy List") <> 0 Then
  GetUserSN = Mid$(sn$, 1, InStr(sn$, "'s Buddy List") - 1)
Else
  GetUserSN = "Dude"
End If
End Function

Public Sub GoToRoom(Room As String)
'This sub takes u 2 a room
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
Dim aimchatinvitesendwnd&
Dim edit&, x&, roomx&

roomx& = FindWindow("AIM_ChatWnd", vbNullString)
If roomx& <> 0 Then
KillWindow roomx&
TimeOut 0.2
End If

oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
If oscariconbtn& = 0 Then Exit Sub
ClickIcon oscariconbtn&

Do
  DoEvents
  TimeOut 0.1
  aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
Loop Until aimchatinvitesendwnd& <> 0

edit& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)
x& = SendMessageByString(edit&, WM_SETTEXT, 0, GetUserSN)

edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
x& = SendMessageByString(edit&, WM_SETTEXT, 0, Room$)

oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
ClickIcon oscariconbtn&
End Sub

Public Sub HideAIMwin()
'This hides the aim window
Dim AIM&, Hide&
AIM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Hide& = ShowWindow(AIM&, SW_HIDE)
End Sub
Public Sub HideAd()
'This hides the aim ad
Dim oscarbuddylistwin&, wndateclass&, x&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
x& = ShowWindow(wndateclass&, SW_HIDE)
End Sub

Public Sub KillWindow(Window&)
'Closes a window
Dim x&
x = SendMessageByNum(Window&, WM_CLOSE, 0, 0)
End Sub

Public Function IsOnAIM() As Boolean
'this will return true if your online and
'false if your not
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
If oscarbuddylistwin& = 0 Then
IsOnAIM = False
Else
IsOnAIM = True
End If
End Function

Public Sub IMSend(Person$, Message$)
'This sends an im 2 someone
Dim oscarbuddylistwin&, oscartabgroup&, oscariconbtn&
Dim aimimessage&
Dim oscarpersistantcombo&
Dim edit&, x&
Dim wndateclass&
Dim ateclass&, button&

oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)

If LCase$(Trim$(IMSN)) = LCase$(Trim$(Person$)) Then
aimimessage& = FindWindow("aim_imessage", vbNullString)
x& = FindChildByClass(aimimessage&, "WndAte32Class")
x& = GetWindow(x&, 2)
button& = FindChildByClass(aimimessage&, "_Oscar_IconBtn")
x& = SendMessageByString(x&, WM_SETTEXT, 0, Message$)
TimeOut 0.3
ClickIcon (button&)
Exit Sub
End If

oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
If oscariconbtn& = 0 Then Exit Sub
ClickIcon oscariconbtn&

Do
  DoEvents
  TimeOut 0.1
  aimimessage& = FindWindow("aim_imessage", vbNullString)
Loop Until aimimessage& <> 0

oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
x& = SendMessageByString(edit&, WM_SETTEXT, 0, Person$)

wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
x& = SendMessageByString(wndateclass&, WM_SETTEXT, 0, Message$)

oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
ClickIcon oscariconbtn&
End Sub

Public Function IMSN() As String
'this will get the screen name from an im
Dim aimimessage&, sn$, b%
aimimessage& = FindWindow("aim_imessage", vbNullString)
sn$ = GetCaption(aimimessage&)
b% = InStr(sn$, " - ")
If b% <> 0 Then
IMSN$ = Mid$(sn$, 1, b% - 1)
Else
IMSN$ = "(no one)"
End If
End Function

Public Function GetLastChatLine(txt$) As String
'this grabs up the last chat line, see below
Dim x As Integer, M As Integer
Do
DoEvents
x = x + 1
If Mid(txt, Len(txt) - x, 4) = "<BR>" Then
M = M + 1
End If
Loop Until M = 2
GetLastChatLine = Mid$(txt, Len(txt) - x, x - 10)
'Example on how 2 get last chat line info using
'these function: (put in button)
'
'Dim chat&, chatwin&, txt$, news$, who$, what$
'chat& = FindWindow("AIM_ChatWnd", vbNullString)
'If chat& = 0 Then Exit Sub
'chatwin& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
'txt$ = WinTxt(chatwin&)
'news$ = GetLastChatLine(txt$)
'who$ = GetLastSN(news$)
'what$ = GetLastSaid(news$)
'MsgBox who$ & "=" & what$
End Function

Public Function GetLastSN(txt$) As String
DoEvents '56
GetLastSN = Mid$(txt, 44, InStr(44, txt, "<") - 44)
End Function
Public Function GetLastSaid(txt$) As String
Dim x As String
x = Mid$(txt, InStr(txt, ":</FONT>") + 31)
GetLastSaid = Right$(x, Len(x))
End Function

Public Sub LCaseChat()
'This makes the chat lower case
Dim chat&, cht&, txt As String, x As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
cht& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
x = SendMessageByString(cht&, WM_SETTEXT, 0&, LCase$(WinTxt(cht&)))
End Sub

Public Function scrambleword(Text1 As String) As String
Dim PATorJK$, patorjk2$, i%, X1$, X2%, L%
PATorJK$ = Text1 + " "
Do
X1$ = Mid$(PATorJK$, 1, InStr(PATorJK$, " ") - 1)
L% = Len(X1$)
For i% = 1 To L%
Randomize
X2% = Int(Len(X1$) * Rnd) + 1
patorjk2$ = patorjk2$ + Mid$(X1$, X2%, 1)
X1$ = Mid$(X1$, 1, X2% - 1) + Mid$(X1$, X2% + 1)
Next
patorjk2$ = patorjk2$ + " "
PATorJK$ = Mid$(PATorJK$, InStr(PATorJK$, " ") + 1)
Loop Until InStr(PATorJK$, " ") = 0
scrambleword = RTrim(patorjk2$)
End Function

Public Sub SendImage(File$)
'This lets u send a pic into the chat (see below)
SendText "<<p>img src=" & Chr$(34) & File$ & Chr$(34) & ">"
'example: Call SendImage("c:\0034.jpg")
End Sub

Public Sub SendLink(URL$)
'This lets u send a link 2 the chat room
SendText "<A HREF=" & Chr$(34) & URL & Chr$(34) & ">" & URL$ & "</A>"
End Sub

Public Sub SendSize(txt$, size%)
'This lets u select the size of the text that
'goes 2 the chat room.
SendText "<<p>font size=" & size% & ">" & txt$
End Sub

Public Sub SendText(txt$)
'This sends some text 2 a chat room
Dim aimchatwnd&
Dim wndateclass&, x&
Dim ateclass&, oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
x& = SendMessageByString(wndateclass&, WM_SETTEXT, 0, txt$)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
ClickIcon oscariconbtn&
End Sub

Public Sub ShowAIMwin()
'This shows the aim window
Dim AIM&, Show&
AIM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Show& = ShowWindow(AIM&, SW_SHOW)
End Sub
Public Sub showad()
'Shows the aim ad
Dim oscarbuddylistwin&, wndateclass&, x&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
x& = ShowWindow(wndateclass&, SW_SHOW)
End Sub

Public Sub StayOnTop(Frm As Form)
'This makes it so the form stays ontop of all the other windows
Dim Top As Long
Top = SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub StayOffTop(Frm As Form)
'This makes it so the form is no longer ontop of all the other windows
Dim Top As Long
Top = SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub SuperMacros(M%)
'Some cool macros, once u see how there done
'u'll know how 2 do um, just don't make the
'macro 2 big of it wont work, example:
'Call SuperMacros(1)
Dim x%, Yeahman As String, i%, Isthere%
Dim Yeahman2 As String

Select Case M%
Case 1  ' AIM
SendText "<<p>br>                         o<<p>br>                      o88o<<p>br>                    o88888o<<p>br>                  o88888888o<<p>br>                o88º°´¨ ¸,.¸ ¨`°º<<p>br>             o8º¸ , o888 '88o,    ,<<p>br>           o8¹ , o880000088, ¹o<<p>br>         o8,oº¸8000000008   ,8o<<p>br>       o888 ' ,8000000000' ,o888o<<p>br>     o8888   '880000088,o0º¸8888o<<p>br>  o888888 °080008880°¸o8888888o<<p>br>o88888888o.      `° ººº°´  .o'8888888o"
Case 2  ' Cat
SendText "<<p>br>               /|<<p>br>         \`O.o'<<p>br>         =(_|_)=<<p>br>             U     "
Case 3  ' AIM SuX
SendText "<<p>br>                         o<<p>br>                      o88o<<p>br>                    o88888o<<p>br>                  o88888888o<<p>br>                o88º°´¨ ¸,.¸ ¨`°º<<p>br>             o8º¸ , o888 '88o,    ,<<p>br>           o8¹ , o880000088, ¹o<<p>br>         o8,oº¸8000000008   ,8o<<p>br>       o888 ' ,8000000000' ,o888o<<p>br>     o8888   '880000088,o0º¸8888o<<p>br>  o888888 °080008880°¸o8888888o<<p>br>o88888888o.      `° ººº°´  .o'8888888o<<p>br>                AOL/AIM SUX!"
Case 4  ' Finger
SendText "<<p>br>                       /´¯/)<<p>br>                     ,/¯  /<<p>br>                    /    /<<p>br>              /´¯`/'   '/´¯¯`·¸<<p>br>           /'/    /    /    /¨  /¯\<<p>br>          ('(    ´    ´     ¯ /'   ')<<p>br>           \                 '     /<<p>br>            '\'   \           _.·´<<p>br>              \              (<<p>br>                \             \"
Case 5  ' Poof
SendText "<<p>br>         /\              \    /<<p>br>        /  \            -- O --<<p>br>       / (* \             /  \<<p>br>      /____\          / ::Waving Da Magic Wand::<<p>br>      (       )         /<<p>br>      (  *o* )      /    --Abra-Ca-Dabra--<<p>br>     (((((U)))))  /<<p>br>       /      \   /\     <poof!> jus' like magic... ;o)<<p>br>                                      I'm gone!"
Case 6  ' LoL
SendText "<<p>br>   __ _                  __ _<<p>br>  l¯¯¯l\                 l¯¯¯l\<<p>br>  l     l ll         __    l     l  l<<p>br>  l     l_l _    /l¯¯l\\ l     l_l _<<p>br>  l______ l\ l/ ¯¯ \l l______ l\<<p>br>   \______\ll \ __ /  \______\ll<<p>br>    ¯¯¯ ¯ ¯     ¯ ¯     ¯¯¯ ¯ ¯"
Case 7  ' Pigman
SendText "<<p>br>          //\\____//\\<<p>br>         (   O)  ( º  )<<p>br>        /      ( oo )    \<<p>br>        \____ O  ___/     \IIIII<<p>br>  _____I        I______/  /<<p>br> /   ___                ____/<<p>br>/  /      /   º       º   \   Smoke<<p>br>IIII\   /      °   °       \   Your<<p>br>        I      ° o °       I  WeeD<<p>br>         \     °__°      /   Kids!!<<p>br>       \IIIIIII   IIIIIII/  Its good for ya!{S Vomit"
Case 8  ' Feet
SendText "<<p>br>                         Oooo<<p>br>                          (     )<<p>br>                            )  /<<p>br>                           (_/<<p>br>                oooO<<p>br>                (     )<<p>br>                 \  (<<p>br>                  \_)"
Case 9  ' Alien Head
SendText "<<p>br>     .·-·-·-·-·-·-·-··.<<p>br> .·´                      `·.<<p>br>:                            :<<p>br>:                            :<<p>br> `·.                      .·´<<p>br> :´¯`·.           .·´¯`:<<p>br>  '·.   0`·.   .·´0   .·'<<p>br>      `:--·´   `·--:´<<p>br>        `·.  ' ' .·´<<p>br>           `·-·´"
Case 10 ' Nyah Nyah
SendText "<<p>br>   \ | /     ______     \ | /<<p>br>    @   /     oo    \ ~  @<<p>br>    /__ (    \___/   )__\    nyah nyah!!<<p>br>           \___U __/"
Case 11 ' Evil Face
SendText "<<p>br> ·´¯`.     .´¯`·<<p>br>      '.  .'<<p>br>    (\ '`·´ /)<<p>br> .              .<<p>br>·´`,· . __ . ·,´`·<<p>br>   \::::::::::::::::/<<p>br>     `'·····'´"
Case 12 ' Crown
SendText "<<p>br>                 @<<p>br>        @      /\       @<<p>br>        '/\      /: :\      /\<<p>br>©    /: :\    /::  ::\    /: :\     ©<<p>br>|\   /::  ::\  /:::   :::\  /::  ::\     '/|<<p>br>|; \/::    ::\/:::     :::\/::     ::\/ ;|<<p>br> '\,-=·=·=·'````´´'·=·=·=-,/'<<p>br>  `0-''```'''''''`````'''''''```''-0´<<p>br>    Í¯```````****´´´´´´´¯Ì<<p>br>     \:::::::::::::::::::::::::::::::::::::/"
Case 13 ' Taco Bell Dog
SendText "<<p>br>',`*·.¸   Yo quiero  ¸.·*´ ,'<<p>br>  ',     `·. ¸ .  ·  . ¸ .·´     ,'<<p>br>   `·.,¸    ¸.  ¸   ¸ .¸    ¸,.·´<<p>br>        `.  `“´“O“`“´  .´ - taco bell<<p>br>          `.  · ­ ­ ·  .´<<p>br>              `·. . .·´"
Case 14 ' Peace
SendText "<<p>br>      (\ (\<<p>br>     (  (  \    _<<p>br>    ( (     \(  * )>/º<<p>br>\`\\`\           `)   \   PÊÃÇÊ<<p>br> \________/"
Case 15 ' LOL2
SendText "<<p>br>|¯¯|                  '|¯¯|<<p>br>|   '|__ '/¯¯/\¯¯\ |   '|__<<p>br>|____/||\__\/__/||____/|<<p>br>|¸___'|/'\|__|'__|/'|¸___'|/"
Case 16 ' Uncle Sam
SendText "<<p>br>   ___<<p>br>  I     I<<p>br>_I __I_    _ _ _<<p>br> q*^*p < SuP? |<<p>br>   )'''''(      ¯ ¯ ¯"
End Select
End Sub

Public Function Text_Dot(Text As String) As String
Dim i As Integer, pat$
For i = 1 To Len(Text)
If Mid$(Text, i, 1) <> Chr(13) And i <> Len(Text) Then
pat$ = pat$ + Mid$(Text, i, 1) & "•"
Else
pat$ = pat$ + Mid$(Text, i, 1)
End If
Next
Text_Dot = pat$
End Function

Public Function Text_Elite(Text As String) As String
Dim x$, X2$, i%, Boo%, Boo2$, PATorJK
x$ = "?!0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ¿¡01²³456789åb¢dèƒghîjklmñºÞq®$†µvw×ýzÁßÇÐÊFGH‡JK£MÑØ¶QR§TÚVWX¥Z"
X2$ = "¿¡01²³456789åb¢dèƒghîjklmñºÞq®$†µvw×ýzÁßÇÐÊFGH‡JK£MÑØ¶QR§TÚVWX¥Z?!0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
For i = 1 To Len(Text)
Boo = InStr(x, Mid(Text, i, 1))
If Not Boo = 0 Then
Boo2 = Mid(X2, Boo, 1)
PATorJK = PATorJK + Boo2
Else
Boo2 = Mid(Text, i, 1)
PATorJK = PATorJK + Boo2
End If
Next
Text_Elite = PATorJK
End Function

Public Function Text_Hacker(Text As String) As String
Dim i As Integer, pat As String, r As Integer
Randomize
For i = 1 To Len(Text)
r = Int(2 * Rnd) + 1
If r = 1 Then
pat$ = pat$ & UCase$(Mid$(Text, i, 1))
Else
pat$ = pat$ & LCase$(Mid$(Text, i, 1))
End If
Next
Text_Hacker = pat$
End Function

Public Function Text_Space(Text As String) As String
Dim i As Integer, pat$
For i = 1 To Len(Text)
If Mid$(Text, i, 1) <> Chr(13) Then
pat$ = pat$ + Mid$(Text, i, 1) & " "
Else
pat$ = pat$ + Mid$(Text, i, 1)
End If
Next
Text_Space = pat$
End Function

Public Sub TimeOut(Duration As Double)
'Make a timeout in a program
Dim starttime As Double
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub

Public Sub UCaseChat()
'This subs upper cases the chat
Dim chat&, cht&, txt As String, x As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
cht& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
x = SendMessageByString(cht&, WM_SETTEXT, 0&, UCase$(WinTxt(cht&)))
End Sub

Public Function WinTxt(ByVal hwnd As Long) As String
'This sub grabs up the text out of a window
Dim x As Integer, y As String, z As Integer
DoEvents
x = SendMessage(hwnd, &HE, 0&, 0&)
y = String(x + 1, " ")
z = SendMessageByString(hwnd, &HD, x + 1, y)
WinTxt = Left(y, x)
End Function
