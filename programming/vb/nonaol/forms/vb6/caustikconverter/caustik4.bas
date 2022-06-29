Attribute VB_Name = "caustik4"
 'Aol4.0 Bas file created by caustik
 'Email: caustik@hotmail.com
 'Aim: ocaustik
 
Public Const WM_PASTE = &H302
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function SetFileAttributes Lib "Kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Declare Function GetFileAttributes Lib "Kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long


Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClipboardViewer Lib "user32" () As Long
Public Declare Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardOwner Lib "user32" () As Long
Public Declare Function GetClassWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long



Public Const LB_FINDSTRING = &H18F
Public Declare Function GetFocus Lib "user32" () As Long

Const MAXIMUM_ALLOWED = &H2000000
Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type WINDOWPLACEMENT
        Length As Long
        FLAGS As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Declare Function SetFocus2 Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Sub RtlMoveMemory Lib "Kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Public Declare Sub GetSystemTime Lib "Kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetCursor Lib "user32" () As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function OpenThreadToken Lib "advapi32.dll" (ByVal ThreadHandle As Long, ByVal DesiredAccess As Long, ByVal OpenAsSelf As Long, TokenHandle As Long) As Long

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Public Type COLORRGB
        red As Integer
        blue As Integer
        green As Integer
End Type

Public Const HTCAPTION = 2
Public Const EM_SETSEL = &HB1
Public Const WM_MOVE = &H3
Public Const WM_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_NCDESTROY = &H82
Public Const WM_COMMAND = &H111
Public Const WM_SIZE = &H5
Public Const WM_ENABLE = &HA
Public Const WM_DELETEITEM = &H2D
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SETCURSOR = &H20
Public Const WM_CLOSE = &H10
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_SETFOCUS = &H7
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CUT = &H300
Public Const VK_SPACE = &H20
Public Const VK_DELETE = &H2E
Public Const VK_RETURN = &HD
Public Const VK_CLEAR = &HC
Public Const VK_LCONTROL = &HA2
Public Const VK_CONTROL = &H11
Public Const VK_BACK = &H8
Public Const VK_UP = &H26
Public Const VK_TAB = &H9
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MINIMIZE = &H20000000
Public Const WM_DESTROY = &H2
Public Const WM_USER = &H400
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CLEAR = &H303
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_MOUSEMOVE = &H200
Public Const CB_GETCOUNT = &H146
Public Const CB_INSERTSTRING = &H14A
Public Const CB_RESETCONTENT = &H14B
Public Const LB_SETITEMDATA = &H19A
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETtext = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Global Const LB_ADDSTRING = (WM_USER + 1)
Public showing As Integer



Public Function AolMain() As Long
AOL% = FindWindow("AOL Frame25", vbNullString)
AolMain = AOL%
End Function
Public Function AolMainClone(Clone As Integer) As Long
AOL% = FindWindow("AOL Frame25", vbNullString)
If Clone = 1 Then AolMainClone = AOL%: Exit Function
v = 1
Do While v < Clone
AOL% = GetWindow(AOL%, 2)
If InStr(1, ClassName(AOL%), "AOL Frame25") <> 0 Then v = v + 1
DoEvents
Loop
AolMainClone = AOL%
End Function

Public Function FindChildByClass(Parent, tafind As String) As Long
child% = GetWindow(Parent, 5)
Dim buffer As String
buffer = String(250, " ")
blank = GetClassName(child, buffer, 250)
    If InStr(1, LCase(buffer), LCase(tafind)) <> 0 Then GoTo found
While child% <> 0
child% = GetWindow(child, 2)
buffer$ = String(250, " ")
blank = GetClassName(child, buffer, 250)
    If InStr(1, LCase(buffer), LCase(tafind)) <> 0 Then GoTo found
Wend
Exit Function
found:
FindChildByClass = child%
End Function
Public Function Wabber(txt As String)
thelen = Len(txt)
txt = Mid(txt, 2, thelen) + Mid(txt, 1, 1)
Wabber = txt
End Function
Public Function Wabber2(txt As String)
thelen = Len(txt)
txt = Mid(txt, thelen, 1) + Mid(txt, 1, thelen - 1)
Wabber2 = txt
End Function

Public Sub Pause(the)
start = Timer + the
Do While Timer < val(start)
DoEvents
Loop
End Sub
Public Function AOLChatSend(text As String) As Long
On Error Resume Next
aolchild% = GetChatRoom()
child% = FindChildByClass(aolchild%, "RICHCNTL")
For numbit = 1 To 10
child% = GetWindow(child%, 2)
buffer$ = String(250, " ")
blank = GetClassName(child%, buffer$, 250)
If InStr(1, buffer$, "RICHCNTL") <> 0 Then GoTo skipout
Next
skipout:
'OLDY = FLD_Main.WinHook1.HwndParam
'FLD_Main.WinHook1.HwndParam = 0
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, text)
'SendText = SendMessage(child%, WM_CHAR, 13, 0)
'SendText = SendMessage(child%, WM_CHAR, 13, 0)
'SendText = SendMessage(child%, WM_CHAR, 13, 0)
DoEvents
SendText = ClickIcon(GetWindow(child%, 2))
'FLD_Main.WinHook1.HwndParam = OLDY
AOLChatSend = 1
End Function
Public Function AOLChatSend4scroll(text As String) As Long
If text = "" Then Exit Function
aolchild% = GetChatRoom()
child% = FindChildByClass(aolchild%, "RICHCNTL")
For numbit = 1 To 10
child% = GetWindow(child%, 2)
buffer$ = String(250, " ")
blank = GetClassName(child%, buffer$, 250)
If InStr(1, buffer$, "RICHCNTL") <> 0 Then GoTo skipout
Next
skipout:

SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, text)
DoEvents
'SendText = SendMessage(child%, WM_CHAR, 13, 0)
'SendText = SendMessage(child%, WM_CHAR, 13, 0)
SendText = ClickIcon(GetWindow(child%, 2))
AOLChatSend4scroll = 1
End Function

Public Function AOLChatSend2(text As String) As Long
On Error Resume Next
If text = "" Then Exit Function
start = Timer + 3
While child% = 0 And start > Timer
aolchild% = GetChatRoom()
child% = FindChildByClass(aolchild%, "RICHCNTL")
DoEvents
For numbit = 1 To 10
child% = GetWindow(child%, 2)
buffer$ = String(250, " ")
blank = GetClassName(child%, buffer$, 250)
If InStr(1, buffer$, "RICHCNTL") <> 0 Then GoTo skipout
Next
Wend

skipout:
old = child%
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, text & Chr(13))
DoEvents
SendText = ClickIcon(GetWindow(child%, 2))
DoEvents
SendText = ClickIcon(GetWindow(child%, 2))
AOLChatSend2 = 1
End Function

Public Sub StayOnTop(the As Form)
stay = SetWindowPos(the.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1)
End Sub
Public Function HideWelcome()
child% = FindChildByClass(AolMain, "MDIClient")
killme = FindChildbyText(child%, "Welcome")
A = ShowWindow(killme, WS_VISIBLE)
On Error GoTo handler
playwav App.Path + "\sfx\rollup.dat"
handler:
End Function
Public Function ShowWelcome()
child% = FindChildByClass(AolMain, "MDIClient")
killme = FindChildbyText(child%, "Welcome")
A = ShowWindow(killme, WS_VISIBLE + 1)
On Error GoTo handler
playwav App.Path + "\sfx\rolldown.dat"
handler:
End Function
Public Function HideAol()
killme = AolMain
blank = ShowWindow(killme, WS_VISIBLE)
On Error GoTo handler
playwav App.Path + "\sfx\rollup.dat"
handler:
End Function
Public Function ShowAol()
killme = AolMain
blank = ShowWindow(killme, 3)
On Error GoTo handler
playwav App.Path + "\sfx\rolldown.dat"
handler:
End Function

Public Function FindChildbyText(Parent, tafind As String) As Long
child% = GetWindow(Parent, 5)
Dim buffer As String
buffer = String(250, " ")
blank = GetWindowText(child%, buffer, 250)
    If InStr(1, UCase(buffer), UCase(tafind)) <> 0 Then GoTo found
While child% <> 0
child% = GetWindow(child, 2)
buffer$ = String(250, " ")
blank = GetWindowText(child, buffer, 250)
    If InStr(1, UCase(buffer), UCase(tafind)) <> 0 Then GoTo found
Wend
Exit Function
found:
FindChildbyText = child%
End Function

Public Function GetFullChatText() As String
child% = GetWindow(GetChatRoom, 5)
thelen% = SendMessageByNum(child%, &HE, 0, 0)
theText$ = String(thelen + 1, " ")
geter% = SendMessageByString(child%, &HD, thelen% + 1, theText$)
GetFullChatText = theText$
End Function
Public Function GetChatRoom() As Long
child% = FindChildByClass(AolMain, "MDIClient")
'aolchild% = FindChildByClass(child%, "AOL Child")
aolchild% = GetWindow(child%, 5)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_Listbox") And FindChildByClass(aolchild%, "_AOL_Combobox") Then GoTo found
While aolchild% <> 0
aolchild% = GetWindow(aolchild%, 2)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_Listbox") And FindChildByClass(aolchild%, "_AOL_Combobox") Then GoTo found
Wend
child% = FindChildByClass(AolMain, "MDIClient")
'aolchild% = FindChildByClass(child%, "AOL Child")
aolchild% = GetWindow(child%, 5)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_Combobox") Then GoTo found
While aolchild% <> 0
aolchild% = GetWindow(aolchild%, 2)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_Combobox") Then GoTo found
Wend
child% = FindChildByClass(AolMain, "MDIClient")
'aolchild% = FindChildByClass(child%, "AOL Child")
aolchild% = GetWindow(child%, 5)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_ListBox") Then GoTo found
While aolchild% <> 0
aolchild% = GetWindow(aolchild%, 2)
If FindChildByClass(aolchild%, "RICHCNTL") And FindChildByClass(aolchild%, "_AOL_ListBox") Then GoTo found
Wend

GetChatRoom = 0
Exit Function

found:
GetChatRoom = aolchild%
End Function
Public Function GetLastChatLine() As String
word$ = GetFullChatText
thelen = Len(word$)
For finder = thelen + 1 To 1 Step -1
If InStr(finder, word$, Chr(13)) Then GoTo aight
Next
Exit Function
aight:
GetLastChatLine = Mid(word$, finder + 1, finder + 20)
End Function
Public Function GetLastChatGuy() As String
word$ = GetLastChatLine
For finder = 1 To 20
If Mid(word$, finder, 1) = ":" Then GoTo aight
Next
aight:
GetLastChatGuy = Mid(word$, 1, finder - 1)
End Function
Public Function GetUser() As String
child = FindChildByClass(AolMain, "MDIClient")
child = FindChildbyText(child, "Welcome,")
thelen = GetWindowTextLength(child)
If thelen = 0 Then GetUser = "unknown": Exit Function
buffer$ = String(thelen, " ")
geter = GetWindowText(child, buffer$, thelen)
buffer$ = Right(buffer$, thelen - 9)
GetUser = Left(buffer$, Len(buffer$) - 1)
End Function
Public Function Toolbar(numb)
child% = FindChildByClass(AolMain, "AOL Toolbar")
child% = FindChildByClass(child%, "_AOL_Toolbar")
child% = FindChildByClass(child%, "_AOL_Icon")
numb = numb - 1
If numb > 0 Then
    For v = 1 To numb
        child% = GetWindow(child%, 2)
    Next
End If
ClickIcon child%
End Function
Public Function ClickIcon(icon)
Click% = PostMessage(icon, WM_LBUTTONDOWN, VK_SPACE, 0&)
DoEvents
Click% = PostMessage(icon, WM_LBUTTONUP, VK_SPACE, 0&)
End Function

Public Function SendEmail(Who As String, subject As String, message As String) As Long
Toolbar (2)

mdich% = FindChildByClass(AolMain, "MDIClient")
child% = 0
start = Timer + 10
Do While childml% = 0
If start < Timer Then Exit Function
    childml% = FindChildbyText(mdich%, "Write Mail")
DoEvents
Loop

who2% = 0
start = Timer + 10
Do While who2% = 0
If start < Timer Then Exit Function
who2% = FindChildByClass(childml%, "_AOL_Edit")
who3% = who2%
DoEvents
Loop
who3o% = who3%
For lp = 1 To 4
start = Timer + 10
Do While who3% = who3o%
If start < Timer Then Exit Function
who3% = GetWindow(who3%, 2)
DoEvents
Loop
who3o% = who3%
Next

start = Timer + 10
Do While who4% = 0
If start < Timer Then Exit Function
who4% = FindChildByClass(childml%, "RICHCNTL")
DoEvents
Loop
start = Timer + 10
Do While sendh% = 0
If start < Timer Then Exit Function
sendh% = FindChildByClass(childml%, "_AOL_Icon")
For lp = 1 To 18
sendh% = GetWindow(sendh%, 2)
Next
DoEvents
Loop
find1% = 23
seter% = SendMessageByString(who2%, &HC, 0, Who)
DoEvents
seter% = SendMessageByString(who3%, &HC, 0, subject)
DoEvents
'For v = 1 To Len(message) Step 150
seter% = SendMessageByString(who4%, &HC, 0, message)
DoEvents
'ClickIcon (who2%)
'ClickIcon (who4%)
'DoEvents
'Next

Do While find1% <> 0
ClickIcon sendh%
start = Timer + 3
While start > Timer
find1% = FindChildbyText(mdich%, "Write Mail")
If find1% = 0 Then GoTo outtie
DoEvents
Wend
outtie:
find1% = FindChildbyText(mdich%, "Write Mail")
DoEvents
Loop

SendEmail = 1
End Function

Public Function OpenNewMail() As Long
Toolbar (1)
Do While Mailop% = 0
mdich% = FindChildByClass(AolMain, "MDIClient")
Mailop% = FindChildbyText(mdich%, "Online Mailbox")
DoEvents
Loop
OpenNewMail = 1
End Function
Public Function Clickmail(numb) As Long
mdich% = FindChildByClass(AolMain, "MDIClient")
Do While mailbx% = 0
mailbx% = FindChildbyText(mdich%, "Online Mailbox")
DoEvents
Loop
Do While tabr% = 0
tabr% = FindChildByClass(mailbx%, "_AOL_TabControl")
DoEvents
Loop
Do While tabr2% = 0
tabr2% = FindChildByClass(tabr%, "_AOL_TabPage")
DoEvents
Loop
Do While tabr3% = 0
tabr3% = FindChildByClass(tabr2%, "_AOL_Tree")
DoEvents
Loop

Dim theprocess As Long
Dim geting As Long

geting = GetWindowThreadProcessId(tabr3%, theprocess)
MsgBox theprocess

End Function
Public Function GetRoomCount() As Long
lister = FindChildByClass(GetChatRoom, "_AOL_Listbox")
If lister = 0 Then Exit Function
Counter = SendMessageByNum(lister, LB_GETCOUNT, 0, 0)
GetRoomCount = Counter
End Function
Public Function KeyWord(txt) As Long
MDI = FindChildByClass(AolMain, "MDIClient")
child = FindChildByClass(AolMain, "AOL Toolbar")
child = FindChildByClass(child, "_AOL_Toolbar")
menucheck = FindChildByClass(child, "_AOL_Icon")
child = FindChildByClass(child, "_AOL_Combobox")
child3 = FindChildByClass(child, "Edit")
ClickIcon (child)
ClickIcon (child3)
blank = SendMessageByString(child3, WM_SETTEXT, 0, "")
blank = SendMessageByString(child3, WM_SETTEXT, 0, txt)
Pause 0.025
blank = PostMessage(child3, WM_CHAR, 32, 0)
blank = PostMessage(child3, WM_CHAR, 8, 0)
Pause 0.025
blank = PostMessage(child3, WM_CHAR, 13, 0)
'blank = sendmessagebystring(child3, WM_SETTEXT, 0, "~~~~~~~~~~~~~~~~~~~~~~Fluid by caustik~~~~~~~~~~~~~~~~~~~~~~")
End Function
Public Function Fadehoptext(FadeMe As String, redstart, redend, greenstart, greenend, bluestart, blueend) As String
If redstart < 20 Then redstart = 20
If redend < 20 Then redend = 20
If greenstart < 20 Then greenstart = 20
If greenend < 20 Then greenend = 20
If bluestart < 20 Then bluestart = 20
If blueend < 20 Then blueend = 20
If Len(FadeMe) < 1 Then Exit Function
thelen = Len(FadeMe)
word$ = ""
stepred = (redend - redstart) \ thelen
stepgreen = (greenend - greenstart) \ thelen
stepblue = (blueend - bluestart) \ thelen
ared = redstart
agreen = greenstart
ablue = bluestart
For v = 1 To thelen + 1
A = A + 1
If A > 4 Then A = 1
If A = 1 Then word$ = word$ + "<sub><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1) + "</sub>"
If A = 2 Then word$ = word$ + "<FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1)
If A = 3 Then word$ = word$ + "<sup><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1) + "</sup>"
If A = 4 Then word$ = word$ + "<FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1)
ared = ared + stepred
agreen = agreen + stepgreen
ablue = ablue + stepblue
Next
Fadehoptext = word$
End Function

Public Function Listlines(ScreenNames As ListBox, request As ListBox, Trigger As String, Badlist As ListBox)
theline = GetLastChatLine
word$ = theline
For finder = 1 To 20
If Mid(word$, finder, 1) = ":" Then GoTo aight
Next

aight:
finder2 = InStr(finder, word$, Chr(32))
v = finder2
Do While Char = 32
v = v + 1
Char = Asc(Mid(word$, v, 1))
Loop
finder2 = v
aight2:
theguy = Mid(word$, 1, finder - 1)
theline = Mid(word$, finder + 3)
If UCase(theguy) = UCase("Caustik") And InStr(1, UCase(theline), UCase("-SCARE-")) <> 0 And FLD_Sleeping.Visible = False Then Call FLD_Buffer.FKO_Click
If theline = publastline Then Exit Function
'If theguy = TheUser Then Exit Function
If InStr(1, LCase(theline), LCase(Trigger)) > 0 And InStr(1, LCase(theline), LCase(Trigger)) < 5 Then
        If Badlist.ListCount > 0 Then
        For check = 0 To Badlist.ListCount
            If Badlist.List(check) = theguy Then Exit Function
        Next
        End If
    ScreenNames.AddItem theguy
    request.AddItem theline
End If
publastline = theline
End Function
Public Function ListlinesVote(ScreenNames As ListBox, request As ListBox, screennames2 As ListBox, request2 As ListBox, Trigger As String, Trigger2 As String, Badlist As ListBox)
theline = GetLastChatLine
word$ = theline
For finder = 1 To 20
If Mid(word$, finder, 1) = ":" Then GoTo aight
Next

aight:
finder2 = InStr(finder, word$, Chr(32))
v = finder2
Do While Char = 32
v = v + 1
Char = Asc(Mid(word$, v, 1))
Loop
finder2 = v
aight2:
theguy = Mid(word$, 1, finder - 1)
theline = Mid(word$, finder + 3)
If theline = publastline Then Exit Function
'If theguy = TheUser Then Exit Function
If InStr(1, LCase(theline), LCase(Trigger)) > 0 And InStr(1, LCase(theline), LCase(Trigger)) < 5 Then
        If Badlist.ListCount > 0 Then
        For check = 0 To Badlist.ListCount
            If Badlist.List(check) = theguy Then Exit Function
        Next
        End If
    ScreenNames.AddItem theguy
    request.AddItem theline
End If
If InStr(1, LCase(theline), LCase(Trigger2)) > 0 And InStr(1, LCase(theline), LCase(Trigger2)) < 5 Then
        If Badlist.ListCount > 0 Then
        For check = 0 To Badlist.ListCount
            If Badlist.List(check) = theguy Then Exit Function
        Next
        End If
    screennames2.AddItem theguy
    request2.AddItem theline
End If

publastline = theline
End Function
Public Function tAFKBot(Peephold As ListBox, Reqhold As ListBox, logmsgs1 As ListBox, Trigger As String, Badlist As ListBox, Reasony As String, logmsgs2 As ListBox, Timerlabel As Label, Allowed)
On Error Resume Next
Dim starttime
Dim mintime
Allowed = Allowed - 1
starttime = Timer
mintime = Timer
Do While afkon = 1
If Timer > mintime + 60 Then
    Timerlabel.Caption = Timerlabel.Caption + 1
    AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(header + " Fluid AFK bot", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
    AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(header + " " + LCase(TheUser) + " has been AFK for " + Str(((Timer - starttime) \ 60)) + " minute(s)", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(header + " Reason-" + Reasony, MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
mintime = Timer
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(header + " Type " + Chr(34) + Trigger + Chr(34) + " to leave a message", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
mintime = Timer
End If
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
If Peephold.ListCount > 0 Then
           If InStr(1, UCase(Reqhold.List(0)), "FUCK") <> 0 Or InStr(1, UCase(Reqhold.List(0)), "YOU SUCK") <> 0 Or InStr(1, UCase(Reqhold.List(0)), " FAG") <> 0 Or InStr(1, UCase(Reqhold.List(0)), " GAY") <> 0 Or InStr(1, UCase(Reqhold.List(0)), " ASS ") <> 0 Or InStr(1, UCase(Reqhold.List(0)), "BITCH") <> 0 Or InStr(1, UCase(Reqhold.List(0)), "ASS!") <> 0 Or InStr(1, UCase(Reqhold.List(0)), " SHIT") <> 0 Or InStr(1, UCase(Reqhold.List(0)), "DICK") <> 0 Then
                AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(Peephold.List(0) + ", Your messages are being ignored : x", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
            Call AddIfNotOn(Badlist, Peephold.List(0))
                Peephold.RemoveItem (0)
                Reqhold.RemoveItem (0)
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
DoEvents
    GoTo skiper
    End If
    thespot = InStr(1, Reqhold.List(0), Trigger) + Len(Trigger)
    reqer = Mid(Reqhold.List(0), thespot)
    checkallows = 0
    For check = 0 To logmsgs1.ListCount
        If logmsgs1.List(check) = Peephold.List(0) Then checkallows = checkallows + 1
        If checkallows > Allowed And Allowed <> -1 Then Call AddIfNotOn(Badlist, (Peephold.List(0))): GoTo skipy:
    Next
    Call logmsgs1.AddItem(Peephold.List(0)): logmsgs2.AddItem (Mid(Reqhold.List(0), Len(Trigger) + 1)) ' new
    If checkallows = Allowed And Allowed <> -1 Then
    AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(Peephold.List(0) + ", Fluid has saved your message [" + Str(logmsgs1.ListCount) + "] That was your last message!", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
    Else
    If FLD_Buffer.MenNotifyMessages.Checked = True Then playwav App.Path + "\sfx\message.dat"
    If Allowed = -1 Then
    AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(Peephold.List(0) + ", Fluid has saved your message [ " + Str(logmsgs1.ListCount) + "]", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
    Else
    AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(Peephold.List(0) + ", Fluid has saved your message [ " + Str(logmsgs1.ListCount) + "] You have [" + Str(Allowed - checkallows) + "] Messages Left", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
    End If
    End If
skipy:
start = Timer
Do While Timer < start + 1
DoEvents
keepupdate = Listlines(Peephold, Reqhold, Trigger, Badlist)
Loop
Peephold.RemoveItem (0)
    Reqhold.RemoveItem (0)
End If
DoEvents
skiper:
Loop
AOLChatSend2 MakeFont("Comic Sans MS") + Fadetext(header + " Fluid AFK bot OFF", MainFade.Red1, MainFade.Red2, MainFade.Green1, MainFade.Green2, MainFade.Blue1, MainFade.Blue2)
End Function
Public Sub MoveForm(frm As Form)

ReleaseCapture
X = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

End Sub
Public Sub Resize(frm As Form)

ReleaseCapture
X = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, 17, 0&)

End Sub
Function ReverseText(text)
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words
End Function
Public Function Fadetext(FadeMe As String, redstart, redend, greenstart, greenend, bluestart, blueend) As String
If redstart < 20 Then redstart = 20
If redend < 20 Then redend = 20
If greenstart < 20 Then greenstart = 20
If greenend < 20 Then greenend = 20
If bluestart < 20 Then bluestart = 20
If blueend < 20 Then blueend = 20
If Len(FadeMe) < 1 Then Exit Function
thelen = Len(FadeMe)
word$ = ""
stepred = (redend - redstart) \ thelen
stepgreen = (greenend - greenstart) \ thelen
stepblue = (blueend - bluestart) \ thelen
ared = redstart
agreen = greenstart
ablue = bluestart
For v = 1 To thelen + 1
word$ = word$ + "<FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid(FadeMe, v, 1)
ared = ared + stepred
agreen = agreen + stepgreen
ablue = ablue + stepblue
Next
Fadetext = word$
End Function
Public Function MakeLink(Title As String, Linkto As String) As String
MakeLink = "<FONT SIZE=3 PTSIZE=10> < a HREF=" + Chr(34) + Linkto + Chr(34) + "></u>" + Title + "</A>"
End Function
Public Function MakeFont(Fontname) As String
MakeFont = "<FONT FACE=" + Chr(34) + Fontname + Chr(34) + ">"
End Function
Public Function Fadehoplaggtext(FadeMe As String, redstart, redend, greenstart, greenend, bluestart, blueend, Bold) As String
If redstart < 20 Then redstart = 20
If redend < 20 Then redend = 20
If greenstart < 20 Then greenstart = 20
If greenend < 20 Then greenend = 20
If bluestart < 20 Then bluestart = 20
If blueend < 20 Then blueend = 20
If Len(FadeMe) < 1 Then Exit Function
thelen = Len(FadeMe)
word$ = ""
stepred = (redend - redstart) \ thelen
stepgreen = (greenend - greenstart) \ thelen
stepblue = (blueend - bluestart) \ thelen
ared = redstart
agreen = greenstart
ablue = bluestart
For v = 1 To thelen + 1
A = A + 1
If A > 4 Then A = 1
If Bold = 1 Then word$ = word$ + "<B>"
If A = 1 Then word$ = word$ + "<HTML></HTML><HTML></HTML><sub><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1) + "</sub>"
If A = 2 Then word$ = word$ + "<HTML></HTML><HTML></HTML><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1)
If A = 3 Then word$ = word$ + "<HTML></HTML><HTML></HTML><sup><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1) + "</sup>"
If A = 4 Then word$ = word$ + "<HTML></HTML><HTML></HTML><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid$(FadeMe, v, 1)
ared = ared + stepred
agreen = agreen + stepgreen
ablue = ablue + stepblue
Next
Fadehoplaggtext = word$
End Function
Public Function Fadelaggtext(FadeMe As String, redstart, redend, greenstart, greenend, bluestart, blueend, Bold) As String
If redstart < 20 Then redstart = 20
If redend < 20 Then redend = 20
If greenstart < 20 Then greenstart = 20
If greenend < 20 Then greenend = 20
If bluestart < 20 Then bluestart = 20
If blueend < 20 Then blueend = 20
If Len(FadeMe) < 1 Then Exit Function
thelen = Len(FadeMe)
word$ = ""
stepred = (redend - redstart) \ thelen
stepgreen = (greenend - greenstart) \ thelen
stepblue = (blueend - bluestart) \ thelen
ared = redstart
agreen = greenstart
ablue = bluestart
For v = 1 To thelen + 1
If Bold = 1 Then word$ = word$ + "<B>"
word$ = word$ + "<HTML></HTML><HTML></HTML><HTML></HTML><FONT COLOR=#" + Hex(Int(ared)) + Hex(Int(agreen)) + Hex(Int(ablue)) + ">" + Mid(FadeMe, v, 1)
ared = ared + stepred
agreen = agreen + stepgreen
ablue = ablue + stepblue
Next
Fadelaggtext = word$
End Function
Public Function Showactive(thelabel As Label)

GoTo skipy2
If showing = 1 Then Exit Function
showing = 1
original = thelabel.Caption
For v = 1 To Len(thelabel.Caption)
thelabel.Caption = Wabber(thelabel.Caption)
start = Timer + 0.02
Do While Timer < val(start)
If showing = 0 Then thelabel.Caption = original: Exit Function
DoEvents
Loop
DoEvents
Next
For v = 1 To Len(thelabel.Caption)
thelabel.Caption = Wabber2(thelabel.Caption)
start = Timer + 0.02
Do While Timer < val(start)
If showing = 0 Then thelabel.Caption = original: Exit Function
DoEvents
Loop
DoEvents
Next
showing = 0
Exit Function

skipy:
If showing = 1 Then Exit Function
showing = 1
original = thelabel.Caption
thesize = thelabel.FontSize
For v = thesize To thesize - 3 Step -1
thelabel.FontSize = v
start = Timer + 0.02
Do While Timer < val(start)
DoEvents
Loop
DoEvents
Next
For v = thesize - 2 To thesize
thelabel.FontSize = v
start = Timer + 0.02
Do While Timer < val(start)
DoEvents
Loop
DoEvents
Next
thelabel.FontSize = thesize
showing = 0
Exit Function

skipy2:
If showing = 1 Then Exit Function
showing = 1
original = thelabel.Caption
thesize = thelabel.FontSize
For v = thesize To thesize + 2
thelabel.FontSize = v
start = Timer + 0.02
Do While Timer < val(start)
DoEvents
Loop
DoEvents
Next
For v = thesize + 2 To thesize Step -1
thelabel.FontSize = v
start = Timer + 0.02
Do While Timer < val(start)
DoEvents
Loop
DoEvents
Next
thelabel.FontSize = thesize
showing = 0
Exit Function


End Function
Public Function EmailBody(text As String)
MDI = FindChildByClass(AolMain, "MDIClient")
child = FindChildByClass(MDI, "AOL Child")
If FindChildByClass(child, "_AOL_Static") <> 0 And FindChildByClass(child, "_AOL_Edit") <> 0 And FindChildByClass(child, "_AOL_FontCombo") <> 0 And FindChildByClass(child, "Combobox") <> 0 And FindChildByClass(child, "RICHCNTL") <> 0 Then GoTo found:
Do While child <> 0
child = GetWindow(child, 2)
If FindChildByClass(child, "_AOL_Static") <> 0 And FindChildByClass(child, "_AOL_Edit") <> 0 And FindChildByClass(child, "_AOL_FontCombo") <> 0 And FindChildByClass(child, "Combobox") <> 0 And FindChildByClass(child, "RICHCNTL") <> 0 Then GoTo found:
Loop
Exit Function
found:
buffer$ = String(250, " ")
child = FindChildByClass(child, "RICHCNTL")
blank = SendMessageByString(child, WM_SETTEXT, 0, text)
End Function
Public Function InstantMessageBody(text As String) As Long
MDI = FindChildByClass(AolMain, "MDIClient")
If child = 0 Then child = FindChildbyText(MDI, "Instant Message"): ver = 2
If child = 0 Then child = FindChildbyText(MDI, "Send Instant Message")
If child = 0 Then InstantMessageBody = 0
body = FindChildByClass(child, "RICHCNTL")
If ver = 2 Then
For v = 1 To 9
    body = GetWindow(body, 2)
Next
End If

If body = 0 Then Exit Function
blank = SendMessageByString(body, WM_SETTEXT, 0, "")
blank = SendMessageByString(body, WM_SETTEXT, 0, "")
blank = SendMessageByString(body, WM_SETTEXT, 0, text)
InstantMessageBody = child
End Function
Public Function GetInstantMessageBody() As String
MDI = FindChildByClass(AolMain, "MDIClient")
child = FindChildbyText(FLD_ImOrganizer.Frame1.hWnd, "Instant Message"): ver = 2
If child = 0 Then child = FindChildbyText(MDI, "Instant Message"): ver = 2
If child = 0 Then child = FindChildbyText(MDI, "Send Instant Message")
If child = 0 Then GetInstantMessageBody = ""
body = FindChildByClass(child, "RICHCNTL")
If ver = 2 Then
For v = 1 To 9
    body = GetWindow(body, 2)
Next
End If

If body = 0 Then Exit Function
thelen = SendMessageByNum(body, WM_GETTEXTLENGTH, 0, 0) + 1
text$ = String(thelen, " ")
blank = SendMessageByString(body, WM_gettext, thelen, text$)
GetInstantMessageBody = text$
End Function

Public Function ClickInstantMessageSend()
MDI = FindChildByClass(AolMain, "MDIClient")
If child = 0 Then child = FindChildbyText(MDI, "Instant Message"): ver = 2
If child = 0 Then child = FindChildbyText(MDI, "Send Instant Message"): Screenname = FindChildByClass(child, "_AOL_Edit")
If child = 0 Then Exit Function
sendbutton = FindChildByClass(child, "RICHCNTL")
If ver = 2 Then
For v = 1 To 9
    sendbutton = GetWindow(sendbutton, 2)
Next
End If
sendbutton = GetWindow(sendbutton, 2)
If sendbutton = 0 Then Exit Function
DoEvents
ClickIcon (sendbutton)
End Function

Public Function TurnImsOFF()
SendInstantMessage "$IM_OFF", "Fluid Says Turn My Ims OFF !"
MDI = FindChildByClass(AolMain, "MDIClient")
check = FindChildbyText(MDI, "Send Instant Message")
blank = SendMessageByNum(check, WM_CLOSE, 0, 0)
start = Timer + 5
Do While (Okey1 = 0 Or Okey2 = 0) And start > Timer
Okey1 = FindWindow("#32770", vbNullString)
Okey2 = FindChildbyText(Okey1, "You are")
DoEvents
Loop
Okey = FindChildByClass(Okey1, "Button")
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
End Function
Public Function TurnImsON()
SendInstantMessage "$IM_ON", "Fluid Says Turn My Ims ON !"
MDI = FindChildByClass(AolMain, "MDIClient")
check = FindChildbyText(MDI, "Send Instant Message")
blank = SendMessageByNum(check, WM_CLOSE, 0, 0)
start = Timer + 5
Do While (Okey1 = 0 Or Okey2 = 0) And start > Timer
Okey1 = FindWindow("#32770", vbNullString)
Okey2 = FindChildbyText(Okey1, "You are")
DoEvents
Loop
Okey = FindChildByClass(Okey1, "Button")
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(Okey, WM_LBUTTONUP, 0, 0)
End Function
Public Function SendInstantMessage(Screenname As String, message As String)
MDI = FindChildByClass(AolMain, "MDIClient")
KeyWord "Instant message"
start = Timer + 10
Do While InstantMessageWindow = 0 And start > Timer
InstantMessageWindow = FindChildbyText(MDI, "Send Instant Message")
FreeProcess
Loop
start = Timer + 20
Do While editbox = 0
editbox = FindChildByClass(InstantMessageWindow, "_AOL_Edit")
If start < Timer Then Exit Function
FreeProcess
Loop
Do While richcntl = 0
richcntl = FindChildByClass(InstantMessageWindow, "RICHCNTL")
FreeProcess
Loop
Do While clickme = 0
clickme = GetWindow(richcntl, 2)
FreeProcess
Loop
blank = SendMessageByString(editbox, WM_SETTEXT, 0, Screenname)
blank = SendMessageByString(richcntl, WM_SETTEXT, 0, message)
ClickIcon (clickme)
End Function
Sub playwav(File)
On Error Resume Next
SoundName$ = File
SoundFlags& = &H20000 Or &H1
snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub
Public Function GotoPrivateRoom(roomname)
KeyWord ("aol://2719:2-2-" + roomname)
End Function
Public Function GotoPrivateRoomfast(roomname, n)
For v = 1 To n
KeyWord ("aol://2719:2-2-" + roomname)
Next
End Function

Public Function AIMInstantMessageBody(txt As String) As Long
For trys = 1 To 10
aimim = FindWindow("AIM_IMessage", vbNullString)
If aimim = 0 Then AIMInstantMessageBody = 0
setme = FindChildByClass(aimim, "_Oscar_PersistantCombo")
For v = 1 To 2
setme = GetWindow(setme, 2)
Next
setme = GetWindow(setme, 5)
If setme > 0 Then GoTo skipout:
Next

Exit Function

skipout:
If Len(txt) < 1 Then Exit Function
For v = 1 To Len(txt)
blank = SendMessageByNum(setme, WM_CHAR, Asc(Mid(txt, v, 1)), 0)
Next
sendme = FindChildByClass(aimim, "_Oscar_IconBtn")
blank = SendMessageByNum(sendme, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(sendme, WM_LBUTTONUP, 0, 0)
AIMInstantMessageBody = 1
End Function
Public Function FreeProcess()
For v = 1 To 10
DoEvents
Next
End Function
Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, WM_gettext, GetTrim, TrimSpace$)

AOLGetText = TrimSpace$
End Function
Public Function KillWait()
menu1 = 4
menu2 = 10
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
Do While caustik = 0
caustik = FindWindow("_AOL_Modal", vbNullString)
killa = SendMessage(caustik, WM_CLOSE, 0, 0)
killa = SendMessage(caustik, WM_CLOSE, 0, 0)
killa = SendMessage(caustik, WM_CLOSE, 0, 0)
DoEvents
Loop
End Function
Public Function AddIfNotOn(listbox1 As ListBox, Item As String)
If listbox1.ListCount = 0 Then listbox1.AddItem Item
For check = 0 To (listbox1.ListCount) - 1
    If listbox1.List(check) = Item Then Exit Function
Next
listbox1.AddItem Item
End Function
Public Function GetInstantMessageText(IMhwnd) As String
child = FindChildByClass(IMhwnd, "RICHCNTL")
thelen = SendMessage(child, WM_GETTEXTLENGTH, 0, 0)
buffer$ = String(thelen, " ")
blank = SendMessageByString(child, WM_gettext, thelen, buffer$)
GetInstantMessageText = buffer$
End Function
Public Function GetInstantMessageLastLine(IMhwnd) As String
word$ = GetInstantMessageText(IMhwnd)
thelen = Len(word$)
If thelen < 1 Then GetInstantMessageLastLine = "": Exit Function
    thelen2 = GetWindowTextLength(IMhwnd)
    buffer$ = String(thelen2 + 1, " ")
    blank = GetWindowText(IMhwnd, buffer$, thelen2 + 1)
    thespot = InStr(1, buffer$, ":")
    theguy$ = Mid(buffer$, thespot + 3)
For finder = (thelen + 1) To 1 Step -1
If InStr(finder, word$, ":") <> 0 Then GoTo aight
DoEvents
Next
Exit Function

aight:
GetInstantMessageLastLine = Trim(Mid(word$, finder + 2))
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
Public Function Scrambleword(text As String) As String
thelen = Len(text)
If thelen < 1 Then Exit Function
For v = 1 To 100
spot1 = Int(Rnd * thelen - 1) + 2
spot2 = Int(Rnd * thelen - 1) + 2
word1$ = Mid(text, spot1, 1)
word2$ = Mid(text, spot2, 1)
If word1$ = " " Or word2$ = " " Then GoTo skipy:
Mid(text, spot1, 1) = word2$
Mid(text, spot2, 1) = word1$
skipy:
Next
Scrambleword = text
End Function
Public Function Scramble(text As String) As String
thelen = Len(text)
If thelen < 1 Then Exit Function
spot = 1
If Mid(text, 1, 1) <> " " Then
text = " " + text
thelen = Len(text)
End If
If Mid(text, thelen, 1) <> " " Then
text = text + " "
End If

Loopy:
If Mid(text, spot, 1) = " " Then
        spot2 = InStr(spot + 1, text, " ")
        If spot2 = 0 Then GoTo skipy
        word$ = Mid(text, spot, spot2 - spot)
        Mid(text, spot, spot2 - spot) = Scrambleword(word$)
End If
skipy:
spot = spot + 1
If spot > thelen Then GoTo outty:
GoTo Loopy
outty:
Scramble = Trim(text)
End Function
Public Function NoSpace(txt) As String
For v = 1 To Len(txt)
    If Mid(txt, v, 1) = " " Then txt = Mid(txt, 1, v - 1) & Mid(txt, v + 1)
Next
NoSpace = txt
End Function
Public Function CaustikRypte(word As String, Pw As String) As String
B = 1
For v = 1 To Len(word)
Mid(word, v, 1) = Chr(Asc(Mid(word, v, 1)) + Asc(Mid(Pw, B, 1)))
B = B + 1
If B > Len(Pw) Then B = 1
Next
CaustikRypte = word
End Function
Public Function UNCaustikRypte(word As String, Pw As String) As String
B = 1
For v = 1 To Len(word)
Mid(word, v, 1) = Chr(Asc(Mid(word, v, 1)) - Asc(Mid(Pw, B, 1)))
B = B + 1
If B > Len(Pw) Then B = 1
Next
UNCaustikRypte = word
End Function
Public Function SignOffAol()
menu1 = 3
menu2 = 1
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Function
Public Sub BotIsRunning()
On Error Resume Next
If EchoOn = 1 Then FLD_Error.Label1.Caption = "You must stop Echo Bot before using this bot"
If IdleOn = 1 Then FLD_Error.Label1.Caption = "You must stop Idle Bot before using this bot"
If afkon = 1 Then FLD_Error.Label1.Caption = "You must stop AFK Bot before using this bot"
If Ball8On = 1 Then FLD_Error.Label1.Caption = "You must stop 8Ball Bot before using this bot"
If GuessNumberOn = 1 Then FLD_Error.Label1.Caption = "You must stop Guess the Number Bot before using this bot"
If BustOn = 1 Then FLD_Error.Label1.Caption = "You must stop Room Bust / Room Gagg before using this bot"
If RequestOn = 1 Then FLD_Error.Label1.Caption = "You must stop Request Bot before using this bot"

FLD_Error.Image1.Picture = FLD_Main.skin.Picture: FLD_Error.Show 1
playwav (App.Path + "\SFX\err.dat")

End Sub

Public Function AOLClearRICHSend()
aolchild% = GetChatRoom()
child% = FindChildByClass(aolchild%, "RICHCNTL")
For numbit = 1 To 10
child% = GetWindow(child%, 2)
buffer$ = String(250, " ")
blank = GetClassName(child%, buffer$, 250)
If InStr(1, buffer$, "RICHCNTL") <> 0 Then GoTo skipout
Next
skipout:
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
DoEvents
SetText% = SendMessageByString(child%, WM_SETTEXT, 0, "")
End Function
Public Function AIMChatSend(txt As String) As Long
aimim = FindWindow("AIM_ChatWnd", vbNullString)
If aimim = 0 Then AIMChatSend = 0
setme = FindChildByClass(aimim, "_Oscar_RateMeter")
setme = GetWindow(setme, 3)
setme = GetWindow(setme, 3)
setme = GetWindow(setme, 5)

If Len(txt) < 1 Then Exit Function
For v = 1 To Len(txt)
blank = SendMessageByNum(setme, WM_CHAR, Asc(Mid(txt, v, 1)), 0)
Next
sendme = FindChildByClass(aimim, "_Oscar_RateMeter")
sendme = GetWindow(sendme, 3)
blank = SendMessageByNum(sendme, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(sendme, WM_LBUTTONUP, 0, 0)

Do While aimim <> 0
aimim = GetWindow(aimim, 2)
buffer$ = String(30, " ")
blank = GetClassName(aimim, buffer$, 30)

If InStr(1, buffer$, "AIM_ChatWnd") <> 0 Then
If aimim = 0 Then AIMChatSend = 0
setme = FindChildByClass(aimim, "_Oscar_RateMeter")
setme = GetWindow(setme, 3)
setme = GetWindow(setme, 3)
setme = GetWindow(setme, 5)

If Len(txt) < 1 Then Exit Function
For v = 1 To Len(txt)
blank = SendMessageByNum(setme, WM_CHAR, Asc(Mid(txt, v, 1)), 0)
Next
sendme = FindChildByClass(aimim, "_Oscar_RateMeter")
sendme = GetWindow(sendme, 3)
blank = SendMessageByNum(sendme, WM_LBUTTONDOWN, 0, 0)
blank = SendMessageByNum(sendme, WM_LBUTTONUP, 0, 0)
End If

DoEvents
Loop

AIMChatSend = 1
End Function
Public Function Upchat()
Modal = FindWindow("_AOL_Modal", vbNullString)
blank = EnableWindow(Modal, 0)
blank = EnableWindow(AolMain, 1)
menupchat.Caption = "UpChat OFF"
End Function
Public Function UnUpchat()
Modal = FindWindow("_AOL_Modal", vbNullString)
blank = EnableWindow(Modal, 1)
menupchat.Caption = "UpChat ON"
End Function
Public Function ClassName(hWnd) As String
If hWnd = 0 Then Exit Function
Dim buffer As String
buffer = String(250, " ")
thelen = GetClassName(hWnd, buffer, 25)
buffer = Trim(buffer)
ClassName = buffer
End Function
Public Function AOLChatSendClone(text As String, Clone As Integer) As Long
On Error Resume Next
aolchild% = GetChatRoomClone(Clone)
child% = FindChildByClass(aolchild%, "_AOL_Edit")
SetText% = SendMessageByString(child%, WM_SETTEXT, &H20, text)
blank = SendMessageByNum(child%, WM_CHAR, 13, &H20)
blank = SendMessageByNum(child%, WM_CHAR, 13, &H20)
AOLChatSendClone = 1
End Function

Public Function GetChatRoomClone(Clone As Integer) As Long
child% = FindChildByClass(AolMainClone(Clone), "MDIClient")
'aolchild% = FindChildByClass(child%, "AOL Child")
aolchild% = GetWindow(child%, 5)
If FindChildByClass(aolchild%, "_AOL_Edit") And FindChildByClass(aolchild%, "_AOL_ListBox") Then GoTo found
While aolchild% <> 0
aolchild% = GetWindow(aolchild%, 2)
If FindChildByClass(aolchild%, "_AOL_Edit") And FindChildByClass(aolchild%, "_AOL_ListBox") Then GoTo found
Wend

GetChatRoomClone = 0
Exit Function

found:
GetChatRoomClone = aolchild%
End Function
Public Function FadeAwayForm(frm As Form)
frm.ScaleMode = 3
WidthStep = frm.ScaleWidth / 10
HeightStep = frm.ScaleHeight / 10
Widthy = frm.ScaleWidth
Heighty = frm.ScaleHeight
Do While Heighty > frm.ScaleHeight / 2
Call SetWindowRgn(frm.hWnd, CreateEllipticRgn(frm.ScaleWidth - Widthy, frm.ScaleHeight - Heighty, Widthy, Heighty), True)
Widthy = Widthy - WidthStep
Heighty = Heighty - HeightStep
DoEvents
DoEvents
DoEvents
DoEvents
Loop
End Function
Public Function FadeInForm(frm As Form)
frm.ScaleMode = 3
WidthStep = frm.ScaleWidth / 10
HeightStep = frm.ScaleHeight / 10
Widthy = frm.ScaleWidth / 2
Heighty = frm.ScaleHeight / 2
Do While Heighty < frm.ScaleHeight
Call SetWindowRgn(frm.hWnd, CreateEllipticRgn(frm.ScaleWidth - Widthy, frm.ScaleHeight - Heighty, Widthy, Heighty), True)
If frm.Visible = False Then frm.Visible = True
Widthy = Widthy + WidthStep
Heighty = Heighty + HeightStep
DoEvents
DoEvents
DoEvents
DoEvents
Loop

End Function

Public Function AddRoomByNum(ListBox, Index) As String
Dim buffer As String
Data = SendMessageByNum(ListBox, LB_GETITEMDATA, Index, 0)
buffer = String(30, " ")


AddRoomByNum = buffer
End Function
Public Function AddRoom(ListBox As ListBox)
child = GetChatRoom
child = FindChildByClass(child, "_AOL_Listbox")
Count = SendMessageByNum(child, LB_GETCOUNT, 0, 0)
For v = 0 To Count - 1
ListBox.AddItem AddRoomByNum(child, v)
Next
End Function

Public Function VbtoAol(val As Long) As String
temp = Hex(val)
redo:
If InStr(1, temp, "FF") <> 0 Then
spot = InStr(1, temp, "FF")
    Mid(temp, spot, 2) = "FE"
    GoTo redo
End If


If Len(temp) = 1 Then temp = "00000" & temp
If Len(temp) = 2 Then temp = "0000" & temp
If Len(temp) = 3 Then temp = "000" & temp
If Len(temp) = 4 Then temp = "00" & temp
If Len(temp) = 5 Then temp = "0" & temp
temp2 = Mid(temp, 5, 2) + Mid(temp, 3, 2) + Mid(temp, 1, 2)
temp = temp2
If temp = "FFFFFF" Then temp = "FEFEFE"
VbtoAol = temp
End Function
Public Function caustikly(txt) As String
For v = 1 To Len(txt)
A$ = A$ & Mid(txt, v, 1)
If Mid(txt, v, 1) = " " Then A$ = A$ & "<font color=#004000><b><u>)" Else A$ = A$ & "</b></u><font color=#004000>"
Next
caustikly = "<font color=#004000><b><u>)" & A$
End Function
