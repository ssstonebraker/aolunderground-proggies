Attribute VB_Name = "Module1"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
'Declare Function extfn11A0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
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
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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

Global r&       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
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


Sub aolchangecaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
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

AppActivate "America  Online"
End Sub


Public Sub AddRoom(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)


Person$ = String$(10000, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Function ListToList(Source, destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
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
Public Function AOLGetList(Index)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
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





Function AddListToString(thelist As ListBox)
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function


Sub AddStringToList(theitems, thelist As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
thelist.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub


Function AOLClickList(hWnd)
clicklist% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function


Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLGetListString(parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
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
End If

If which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Sent")
End If

End Sub


Sub AOLRespondIM(MESSAGE)
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Z
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Z
Exit Sub
Z:
E = FindChildByClass(im%, "RICHCNTL")

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
Call AOLSetText(e2, MESSAGE)
AOLIcon (E)
Pause 4
killwin (im%)
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
Pause (10)
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
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
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
city:
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
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Sub HideWindow(hWnd)
hi = ShowWindow(hWnd, SW_HIDE)
End Sub

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(Text, theline)
theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
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



Sub MaxWindow(hWnd)
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hWnd)
mi = ShowWindow(hWnd, SW_MINIMIZE)
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
For Words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, Words, 1)
Next Words


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

Function ScrambleGame(thestring As String)
Dim bytestring As String

thestringcount = Len(thestring$)
If Not Mid(thestring$, thestringcount, 1) = " " Then thestring$ = thestring$ & " "
For Stringe = 1 To Len(thestring$)
characters$ = Mid(thestring$, Stringe, 1)
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
ScrambleGame = Mid(thestrings2$, 1, Len(thestring$) - 1)
End Function


Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
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

Pause (0.2)
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
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For trimspace = 1 To Len(Text)
thechar$ = Mid(Text, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function


Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
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


Function KTEncrypt(ByVal PassWord, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(PassWord) = 0 Then Error 31100
  
  'Is password too long
  If Len(PassWord) > 255 Then Error 31100

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
  PassMax = Len(PassWord)
  
  
  'Tack on leading characters to prevent repetative recognition
  PassWord = Chr$(Asc(Left$(PassWord, 1)) Xor PassMax) + PassWord
  PassWord = Chr$(Asc(Mid$(PassWord, 1, 1)) Xor Asc(Mid$(PassWord, 2, 1))) + PassWord
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 1)) Xor PassMax)
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 2)) Xor Asc(Right$(PassWord, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(PassWord, PassUp, 1))

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
    If Left$(strng, 9) <> Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") Then
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



Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

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
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLGetUser()
On Error Resume Next
AOL& = FindWindow("AOL Frame25", "America  Online")
MDI& = FindChildByClass(AOL&, "MDIClient")
Welcome% = FindChildByTitle(MDI&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength& + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User$
End Function

Sub AOLIMOff()
Call AOLInstantMessage("$IM_OFF", "United Nation Mass IMer")

End Sub

Sub AOLIMsOn()
Call AOLInstantMessage("$IM_ON", "United Nation Mass IMer")

End Sub


Sub AOLChatSend(Txt)
Room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(Room%, "_AOL_Edit"), Txt)
DoEvents
Call SendCharNum(FindChildByClass(Room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub


Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub

Sub AOLCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

theview$ = trimspace$
AOLGetChat = theview$
End Function

Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

AOLGetText = trimspace$
End Function

Sub AOLIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage(Person, MESSAGE)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, MESSAGE)
imsend% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends
AOLIcon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub

Function AOLIsOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
If Welcome% = 0 Then
MsgBox "Please sign on before using this feature.", 64, "TrU-Online Error"
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function

Sub AOLKeyword(Text)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending% = SendMessage(kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

theview$ = trimspace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
Lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = Lastline
End Function

Sub AOLMail(Person, subject, MESSAGE)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
MESS% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And MESS% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(MESS%, WM_SETTEXT, 0, MESSAGE)
AOLIcon (icone%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub

Function AOLRoomCount()
thechild% = AOLFindRoom()
Listerk% = FindChildByClass(thechild%, "_AOL_Listbox")

getcount = SendMessage(Listerk%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLSetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOLSignOff()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(AOL%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
clickicon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function AOLVersion()
AOL% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(AOL%)

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
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function



Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

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
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

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

Pause (0.2)
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

Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
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


Sub UnHideWindow(hWnd)
un = ShowWindow(hWnd, SW_SHOW)
End Sub



Sub WaitForOk()
Do: DoEvents
AOL% = FindWindow("#32770", "America Online")

If AOL% Then
closeaol% = SendMessage(AOL%, WM_CLOSE, 0, 0)
Exit Do
End If

aolw% = FindWindow("_AOL_Modal", vbNullString)

If aolw% Then
AOLButton (FindChildByTitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub

Sub WaitWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    Do
    im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    killwin (im%)
    Loop Until im% = 0



End Function

Sub killwin(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

