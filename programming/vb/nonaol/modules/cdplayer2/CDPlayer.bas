Attribute VB_Name = "RushCD"
'RushOneRok CD Player BAS
'Version 2.0
'Made in VB5
'Email: rushonerok@htomail.com
'
'This module controls the windows 95/98 cd
'player and volume control.
'Aight Check it out.  If you get any
'problems or want me to add anything,
'email me.
'
'Later,
'Rush

Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Function Artist() As String
'Returns the title of the CD
    Dim win As Long, Title As Long, Text As Long, Art As String
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Title& = FindWindowEx(win&, 0&, "SJE_TextClass", vbNullString)
    Text& = GetWindowTextLength(Title&)
    Art$ = String(Text&, 0&)
    Call GetWindowText(Title&, Art$, Text& + 1)
    Artist = Art$
End Function

Function TimeBox() As String
'Returns the CD track and time.
    Dim win As Long, Title As Long, Text As Long, Art As String
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Title& = FindWindowEx(win&, 0&, "SJE_LEDClass", vbNullString)
    Text& = GetWindowTextLength(Title&)
    Art$ = String(Text&, 0&)
    Call GetWindowText(Title&, Art$, Text& + 1)
    TimeBox = Art$
End Function

Sub PLAY()
'Clicks the play button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub Pause()
'Clicks the pause button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub SStop()
'Clicks the stop button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub Previous()
'Clicks the previous (<<) button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub SkipBack()
'Clicks the back skip (<) button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub SkipForward()
'Clicks the forward skip (>) button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub NNext()
'Clicks the next (>>) button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub Eject()
'Clicks the eject button.
    Dim win As Long, Btn As Long
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Btn& = FindWindowEx(win&, 0&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Btn& = FindWindowEx(win&, Btn&, "Button", vbNullString)
    Call SendMessage(Btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub MasterMute()
'Mutes everything.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute all")
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute1()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute2()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute3()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute4()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute5()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute6()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDMute7()
'Mutes CD Player.
    Dim win As Long, Mute As Long
    win& = FindWindow("Volume Control", vbNullString)
    Mute& = FindWindowEx(win&, 0&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    Mute& = FindWindowEx(win&, Mute&, "Button", "&Mute")
    If Mute& = 0 Then: MsgBox ("Unable to mute CD Player:  Select Help from the options menu to fix problem."): Exit Sub
    Call SendMessage(Mute&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Mute&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub CDLeft2()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight2()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp2()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown2()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft3()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight3()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp3()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown3()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft4()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight4()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp4()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown4()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft5()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight5()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp5()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown5()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft6()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight6()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp6()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown6()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft7()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight7()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp7()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown7()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub CDLeft8()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub CDRight8()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub CDUp8()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub CDDown8()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Public Sub RunMenuByString(Class As String, SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow(Class$, vbNullString)
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

Sub Open_EPL()
'Opens the edit play list window (note: stalls program
'until edit play list is closed.
    Call RunMenuByString("SJE_CdPlayerClass", "Edit Play &List...")
End Sub

Sub Close_EPL()
'Closes edit play list
    Dim ParHand1 As Long, OurParent As Long, Hand1 As Long, Hand2 As Long, Hand3 As Long, Hand4 As Long, Hand5 As Long, OurHandle As Long
    ParHand1& = FindWindow("SJE_CdPlayerClass", "CD Player")
    OurParent& = FindWindowEx(ParHand1&, 0, "#32770", vbNullString)
    Hand1& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
    Hand2& = FindWindowEx(OurParent&, Hand1&, "Button", vbNullString)
    Hand3& = FindWindowEx(OurParent&, Hand2&, "Button", vbNullString)
    Hand4& = FindWindowEx(OurParent&, Hand3&, "Button", vbNullString)
    Hand5& = FindWindowEx(OurParent&, Hand4&, "Button", vbNullString)
    OurHandle& = FindWindowEx(OurParent&, Hand5&, "Button", vbNullString)
    Call SendMessage(OurHandle&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OurHandle&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Function ArtName() As String
'Returns the Artist Name
    Dim win As Long, Win2 As Long, Box As Long, Text As Long, Art As String
    Open_EPL
    win& = FindWindow("SJE_CdPlayerClass", "CD Player")
    Win2& = FindWindowEx(win&, 0, "#32770", vbNullString)
    Box& = FindWindowEx(Win2&, 0, "Edit", vbNullString)
    Close_EPL
    Text& = GetWindowTextLength(Box&)
    Art$ = String(Text&, 0&)
    Call GetWindowText(Box&, Art$, Text& + 1)
    ArtName = Art$
End Function

Sub OpenPros()
'Opens the windows CD Player program and the
'volume control and then hides them.
    Dim CD, VC, win As Long, Hide, Win2 As Long, Hide2
    CheckVol
    CD = Shell("C:\WINDOWS\CDPlayer.EXE", 1)
    'AppActivate CD
    VC = Shell("C:\WINDOWS\SNDVOL32.EXE", 1)
    'AppActivate VC
    win& = FindWindow("SJE_CdPlayerClass", vbNullString)
    Hide = ShowWindow(win&, SW_HIDE)
    Win2& = FindWindow("Volume Control", vbNullString)
    Hide2 = ShowWindow(Win2&, SW_HIDE)
End Sub

Sub CloseCD()
'closes cd player.
    Call RunMenuByString("SJE_CdPlayerClass", "E&xit")
End Sub

Sub CloseVol()
'closes volume control.
    Call RunMenuByString("Volume Control", "E&xit")
End Sub

Sub ClosePros()
'Closes the cd player and the volume control
    CloseCD
    CloseVol
End Sub

Sub CheckVol()
'checks if volume control is open and closes it.
    Dim win As Long, Win2 As Long
    win& = FindWindow("Volume Control", vbNullString)
    If win& = 0 Then
    Exit Sub
    Else
    CloseVol
    End If
End Sub

Sub TTE()
'sets the cd time to track time elapsed
    Call RunMenuByString("SJE_CdPlayerClass", "Track Time &Elapsed")
End Sub

Sub TTR()
'sets the cd time to track time remaining
    Call RunMenuByString("SJE_CdPlayerClass", "Track Time &Remaining")
End Sub

Sub DTR()
'sets the cd time to disc time remaining
    Call RunMenuByString("SJE_CdPlayerClass", "Dis&c Time Remaining")
End Sub

Sub RO()
'sets the play to random order
    Call RunMenuByString("SJE_CdPlayerClass", "&Random Order")
End Sub

Sub CP()
'sets the play to continuous play
    Call RunMenuByString("SJE_CdPlayerClass", "&Continuous Play")
End Sub

Sub IP()
'sets the play to intro play
    Call RunMenuByString("SJE_CdPlayerClass", "&Intro Play")
End Sub

Sub StayOnTop(TheForm As Form)
    Dim SetWinOnTop
    SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub MasterLeft()
'Sets the master volume 1 notch to the left speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_LEFT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_LEFT, 0&)
End Sub

Sub MasterRight()
'Sets the master volume 1 notch to the right speaker.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_RIGHT, 0&)
End Sub

Sub MasterUp()
'Turns the master volume up 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_UP, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_UP, 0&)
End Sub

Sub MasterDown()
'Turns the master volume down 1 notch.
    Dim win As Long, SB As Long
    win& = FindWindow("Volume Control", vbNullString)
    SB& = FindWindowEx(win&, 0&, "msctls_trackbar32", vbNullString)
    SB& = FindWindowEx(win&, SB&, "msctls_trackbar32", vbNullString)
    Call SendMessage(SB&, WM_KEYDOWN, VK_DOWN, 0&)
    Call SendMessage(SB&, WM_KEYUP, VK_DOWN, 0&)
End Sub

Sub MiniCap(Mini As Form, Cap As String)
'Fixes a problem I had when my program minimized/restored.
Dim Ht As Long, Wd As Long
If Mini.WindowState <> 1 Then
    Ht& = Mini.Height
    Wd& = Mini.Width
End If
If Mini.Caption = Cap Then
    Mini.Caption = ""
    Mini.Width = Wd&
    Mini.Height = Ht&
End If
End Sub

Sub Pause4(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Function SetName(Text As TextBox) As String
Dim MyString As String, FindString As String, Spot1 As Long
Dim Spot2 As Long
MyString$ = Text.Text
FindString$ = "\"
Spot1& = InStr(MyString$, FindString$)
Spot2& = InStr(Spot1& + 1, MyString$, FindString$)
MsgBox Spot1&
MsgBox Spot2&

End Function
