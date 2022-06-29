Attribute VB_Name = "InsaNe40"

' ________________________________________________ '
'|***   InsaNe40.bas [Version 1] By InsaNiTy   ***|'
'|***      For 32-Bit VB/API Programming       ***|'
'|***       This BAS Was Made For AOL4.0       ***|'
'|***      This May Be Freely Distributed      ***|'
'|***  Any Questions/Comments/Problems Email:  ***|'
'|***           InsaNiTy84@juno.com            ***|'
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯ '


'Copyright 1998. All Rights Reserved. Any modification
'with release is not permitted without consent of
'     InsaNiTy. And may posers rot in hell!


'** Windows 95 API Public Function Declarations **'
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

'** Windows 95 API Public Functions Substitutes **'
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal Dest As Long, ByVal nCount&)
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal Dest&, ByVal nCount&)
Declare Sub ReleaseCapture Lib "user32" ()

'Windows 95 API Private Function & Sub Declarations'
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellUse Lib "shell32.dll Alias (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long" ()
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)


'  ** Public Windows 95 API Constant Functions **  '

'WindowsMessage()
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

'ListBox()
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

'GetWindow()
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

'sndSoundPlay()
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

'ExitWindows()
Public Const EW_RESTARTWINDOWS = &H42
Public Const EW_REBOOTSYSTEM = &H43

'ShowWindow()
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

'ErrorHandling()
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

'OpenFile()
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

'SystemMetrics()
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


Global giBeepBox As Integer
Global r&
Global entry$
Global iniPath$


Sub AOL40_ClickForward()
'This will click the forward icon in AOL 4.0 but you
'must have the forwareded window open
AOL% = FindWindow("AOL Frame25", 0&)
icon% = FindChildByClass(AOL%, "_AOL_ICON")
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
AOLClickIcon (icon%)
End Sub


Function AOL40_FindFwdWindow()

End Function

Public Function GetListIndex(LB As ListBox, txt As String) As Integer
Dim Index As Integer
With LB
For Index = 0 To .ListCount - 1
If .List(iIndex) = txt Then
GetListIndex = Index
Exit Function
End If
Next Index
End With
GetListIndex = -2
End Function
Function GetWinText(hWnd As Integer) As String
'This gets the caption of any window
Dim LengthOfText, Buffer$, GetTheText
LengthOfText = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(LengthOfText)
GetTheText = SendMessageByString(hWnd, WM_GETTEXT, LengthOfText + 1, Buffer$)
GetWinText = Buffer$
End Function
Sub Form_Move(frm As Form)
'This will allow you to move the form to a different
'part of your screen
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hWnd, &HA1, 2, 0)
End Sub
Sub AOL40_ReadMail()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call AOLClickIcon(icon%)
End Sub

Sub AOLClose()
'This will close AOL very quickly
Call Window_Close(AOLWindow())
End Sub
Sub AOLChangeCaption(newcaption)
'This will change AOL's caption
Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub AOLSetText(win, txt)
'This features allows you to change the text from a
'window.
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Function AOLIsOnline() As Integer
'This will return if the user is online
Welcome% = FindChildByTitle(AOLMDI(), "Welcome, ")
If Welcome% = 0 Then
MsgBox "You Must Sign On Before Using This Feature.", 64, "Must Be Online"
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Function AOLRoomCount()
'This returns the number of people currently in the
'chat room you are in
Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Sub AOL40_AddRoomCombo(ListBox As ListBox, ComboBox As ComboBox)
Call AOL40_AddRoomList(ListBox)
For Q = 0 To ListBox.ListCount
ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub AOL40_AntiIdle()
Modal% = FindWindow("_AOL_Modal", vbNullString)
icon% = FindChildByClass(Modal%, "_AOL_Icon")
AOLClickIcon (AOIcon%)
End Sub

Sub AOL40_KillGlyph()
'This will close that little annoying AOL spinning
'thingy on the top corner of AOL 4.0
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolbar%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


Sub ADD_AOL_LB(itm As String, lst As ListBox)
'Add a list of names to a VB ListBox
If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub
Sub AOL40_AddRoomList(lst As ListBox)
'This will add AOL Chat's listbox to your listbox
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
    names$ = String$(256, " ")
    ret = AOLGetList(Index, names$)
    names$ = Left$(Trim$(names$), Len(Trim(names$)))
    ADD_AOL_LB names$, lst
Next Index
endaddroom:
lst.RemoveItem lst.ListCount - 1
i = GetListIndex(lst, AOLUserSN())
If i <> -2 Then lst.RemoveItem i
End Sub
Public Function AOLGetList(Index As Long, Buffer As String)
'This gets the list you request
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = AOL40_FindChatRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
Buffer$ = person$
End Function
Sub WAVStop()
'This will stop a WAV that is playing
Call WAVPlay(" ")
End Sub

Sub WAVLoop(File)
'This will play the WAV you want over and over
SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Function AOLMessageFromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo txt
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo txt
Exit Function
txt:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetWinText(imtext%)
sn = AOLSNFromIM()
snlen = Len(AOLSNFromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
AOLMessageFromIM = Left(blah, Len(blah) - 1)
End Function
Sub WAVPlay(File)
'This will play a WAV file
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub SendCharNum(win, chars)
'This sends any character number of your choice to
'your destinative window
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Sub AOLRespondIM(MESSAGE)
'This find the Instant Message window and responds
'it with the message you want, than closes
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
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
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call AOLSetText(e2, MESSAGE)
AOLClickIcon (e)
pause 0.8
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
e = GetWindow(e, GW_HWNDNEXT)
AOLClickIcon (e)
End Sub

Sub Enter(win)
'This will press enter
Call SendCharNum(win, 13)
End Sub
Sub File_Delete(File)
'This will delete a file straight from the users HD
Kill (File)
End Sub
Sub File_Open(File)
'This will open a file... whole dir and file name needed
Shell (File)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)
'This will immediately rename a file for you
Name sOldLoc As sNewLoc
End Sub
Sub Directory_Create(dir)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
MkDir dir
End Sub
Sub Directory_Delete(dir)
'This deletes a directory automatically from your HD
RmDir (dir)
End Sub

Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Sub SendIM(sn, MESSAGE)
'This will send the message to the SN you wish to on
'AOL 4.0
Call AOL40_Keyword("aol://9293:" & sn)
Do: DoEvents
IMWin% = FindChildByTitle(AOLMDI(), "Send Instant Message")
Rich% = FindChildByClass(IMWin%, "RICHCNTL")
icon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until Rich% <> 0 And icon% <> 0
Call SendMessageByString(Rich%, WM_SETTEXT, 0, MESSAGE)
For X = 1 To 9
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next X
Call pause(0.01)
AOLClickIcon (icon%)
Do: DoEvents
IMWin% = FindChildByTitle(AOLMDI(), "Send Instant Message")
oK% = FindWindow("#32770", "America Online")
If oK% <> 0 Then Call SendMessage(oK%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Function AOLUserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUserSN = User
End Function
Function AOLSNFromIM()
'this will return a Screen Name from an IM
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOLSNFromIM = Naw$
End Function
Public Sub Disable_Ctrl_Alt_Del()
'Disables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
'Enables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub SendMail(sn, subject, MESSAGE)
'This will send mail from AOL4.0
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
icon% = GetWindow(icon%, GW_HWNDNEXT)
Call AOLClickIcon(icon%)
Do: DoEvents
mail% = FindChildByTitle(AOLMDI(), "Write Mail")
Edit% = FindChildByClass(mail%, "_AOL_Edit")
Rich% = FindChildByClass(mail%, "RICHCNTL")
icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And Edit% <> 0 And Rich% <> 0 And icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SendMessageByString(Edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, MESSAGE)
For GetIcon = 1 To 18
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next GetIcon
Call AOLClickIcon(icon%)
End Sub

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Function AOLMDI()
'This function sets focus on AOL's parent window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function




Sub SendChat(txt)
'This will send text to AOL 4.0
Rich% = FindChildByClass(AOL40_FindChatRoom, "RICHCNTL")
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Call SetFocusAPI(Rich%)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, txt)
DoEvents
Call SendMessageByNum(Rich%, WM_CHAR, 13, 0)
End Sub
Sub AOL40_Keyword(KeyWord As String)
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call pause(0.05)
Call AOLClickIcon(icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, KeyWord)
Call pause(0.05)
Call AOLClickIcon(Icon2%)
Call AOLClickIcon(Icon2%)
End Sub
Function FreeProcess()
'This feature will allow you to be in your prog and
'not freeze or have too many errors
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function AOL40_FindChatRoom()
room% = FindChildByClass(AOLMDI(), "AOL Child")
roomlst% = FindChildByClass(room%, "_AOL_Listbox")
roomtxt% = FindChildByClass(room%, "RICHCNTL")
If roomlst% <> 0 And roomtxt% <> 0 Then
AOL40_FindChatRoom = room%
Else
AOL40_FindChatRoom = 0
End If
End Function
Function FindChildByClass(parentw, childhand)
'This will find an MDI Child by the childhand's class
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
room% = firs%
FindChildByClass = room%

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
room% = firs%
FindChildByTitle = room%
End Function
Sub AOLHide()
'This will hide, but no close, the AOL Window
X = FindWindow("AOL Frame25", 0&)
Window_Hide (X)
End Sub

Sub AOLShow()
'This will hide the AOL windows
X = FindWindow("AOL Frame25", 0&)
Window_Show (X)
End Sub

Function AOLClickList(List)
Click% = SendMessage(List, WM_LBUTTONDBLCLK, 0, 0)
End Function
Sub AOLClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Window_Close(win)
'This will close and window of your choice
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub
Sub StayOnTop(frm As Form)
'Allows your form to stay on top of all other windows
Dim ontop%
ontop% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Window_Minimize(win)
'This will minize any window of your choice
X = ShowWindow(win, SW_MINIMIZE)
End Sub

Sub Window_Maximize(win)
'This will maximize the window of your choice
X = ShowWindow(win, SW_MAXIMIZE)
End Sub


Sub Window_Hide(hWnd)
'This will hide the window of your choice
X = ShowWindow(hWnd, SW_HIDE)
End Sub



Sub Window_Show(hWnd)
'This will show the window of your choice
X = ShowWindow(hWnd, SW_SHOW)
End Sub
Sub pause(interval)
'This pauses all activity in your program for the
'amount of time you wish
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Sub Form_Center(frm As Form)
frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Sub Form_Maximize(frm As Form)
'This will maximize the form of your choice
frm.WindowState = 2
End Sub
Sub Form_Minimize(frm As Form)
'This will minimize the form of your choice
frm.WindowState = 1
End Sub
Sub WaitForOk()
'Waits for the AOL OK messages that pops up
Do
DoEvents
Okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until Okw <> 0
   OKB = FindChildByTitle(Okw, "OK")
   okd = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
   oku = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub
Sub AOLIMsOn()
'This will turn your IMs on
Call SendIM("$IM_ON", " ")
End Sub
Sub AOLIMsOff()
'This will turn your IMs on
Call SendIM("$IM_OFF", " ")
End Sub
Sub AOLRunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)
For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)
For GetString = 0 To MenuItemCount% - 1
Subcount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, Subcount%, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = Subcount%
GoTo MatchString
End If
Next GetString
Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub AOLRunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub


Function AOLWindow()
'This sets focus on the AOL window
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function

