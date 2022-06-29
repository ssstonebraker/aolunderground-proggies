Attribute VB_Name = "ravage3"
Public RoomHandle%
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cX&, ByVal cY&, ByVal wFlags&)

Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd&)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Public Const WM_CHAR = &H102
Public Const HWND_TOPMOST = -1
Public Const VK_SPACE = &H20
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const VK_RETURN = &HD
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function ReleaseDC Lib "User" (ByVal hwnd%, ByVal hdc%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hWndParent%, ByVal lpEnumFunc&, ByVal lParam&)
Declare Function FillRect Lib "User" (ByVal hdc As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function GetDC Lib "User" (ByVal hwnd%) As Integer
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


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

Public Const WM_SYSCOMMAND = &H112
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
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

Public Const SC_MOVE = &HF012

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
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0

Global Const SND_SYNC = &H0

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
Sub AcidTrip(frm As Form)
' for the best effect put this in a timer and watch the colors
' Call AcidTrip(Form1)
Dim cX, cY, Radius, Limit
    frm.ScaleMode = 3
    cX = frm.ScaleWidth / 2
    cY = frm.ScaleHeight / 2
    If cX > cY Then Limit = cY Else Limit = cX
    For Radius = 0 To Limit
frm.Circle (cX, cY), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next Radius
End Sub
Sub AddToList(List As ListBox, txt$)
' this will add something to a listbox if it is not all ready in there
' Call AddList(List1, "c")
On Error Resume Next
DoEvents
For X = 0 To List.ListCount - 1
    If UCase$(List.List(X)) = UCase$(txt$) Then Exit Sub
Next
If Len(txt$) <> 0 Then List.AddItem txt$
End Sub
Sub AOL40_AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
' this will add everyone in the chat room to a listbox, then to a combobox
' Call AOL40_AddRoomToComboBox(List1, Combo1)
Call AOL40_AddRoomToListBox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
End Sub
Sub AOL40_AddRoomToListBox(ListBox As ListBox)
' this will add all the names from the chat room to a listbox
' Call AOL40_AddRoomToListBox(List1)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
ListBox.Clear
room = AOL40_FindChatRoom()
If room = 0 Then MsgBox "There must be a chat room open to use this sub", vbInformation, ""
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Sub AOL40_Anti45MinTimer()
' this will click on the window that pops up saying how long u have been online
' Call AOL40_Anti45MinTimer
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AOL40_AntiIdle()
' this will click on an icon so that you will not get logged off do to inactivity
' Call AOL40_AntiIdle
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AOL40_BustInPrivate(URL As String)
' This will bust into a private room that you specify
' it will not run the room when u get in, it will stop
' the secong the room is found so u only enter the chat room
' once. The Keyword for all private rooms is "aol://2719:2-2-ROOMNAME"

busttin:
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
StuFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If StuFF% <> 0 And MoreStuff% <> 0 Then
Call SendMessage(room%, WM_CLOSE, 0, 0)
End If
tryagain:
Call keyword("aol://2719:2-2-" & URL$)
Call AOL40_FindChatRoom
If AOL40_FindChatRoom = 0 Then GoTo noroom:
GoTo FoundRoom:
noroom:
noroomagaiN:
full% = FindWindow("#32770", "America Online")
If AOL40_FindChatRoom <> 0 Then GoTo FoundRoom:
If full% = 0 Then GoTo noroomagaiN:
CloseWin (full%)
GoTo tryagain:
FoundRoom:
 Exit Sub
End Sub

Sub keyword(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = findchildbytitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call Timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub AOL40_ChatHandles()
' this will return the handles of the chat room
Dim e2%, class%
ChatRoomName$ = ""
RoomHandle% = 0
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
child% = FindChildByClass(mdi%, "AOL Child")
e2 = 1
GetNextWindow:
If e2 = 0 Then child% = GetWindow(child%, 2)
If child% = 0 Then GoTo ending
e2 = FindChildByClass(child%, "_AOL_Listbox")
AOLList% = e2
If e2 = 0 Then GoTo GetNextWindow
e2 = FindChildByClass(child%, "_AOL_Combobox")
If e2 = 0 Then GoTo GetNextWindow
e2 = FindChildByClass(child%, "RICHCNTL")
If e2 = 0 Then GoTo GetNextWindow
RoomHandle% = child%
ChatRoomName$ = GetCaption(RoomHandle%)
texthandle% = e2
chatbox% = texthandle%
For X = 1 To 6: texthandle% = GetWindow(texthandle%, 2): Next X
SendButton% = FindChildByClass(RoomHandle%, "_AOL_Icon")
For X = 1 To 5: SendButton% = GetWindow(SendButton%, 2): Next
ending:
End Sub
Sub AOL40_BustInPublic(URL As String)
' This will bust into a public room that you specify
' it will not run the room when u get in, it will stop
' the secong the room is found so u only enter the chat room
' once. The Keyword for all public rooms you can find buy holding
' down the heart and putting it in where you type text, then sending
' it to the chat room
' Call AOL40_BustInPublic("aol://2719:21-2-Lobby%20177")
On Error Resume Next
busttin:
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
StuFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If StuFF% <> 0 And MoreStuff% <> 0 Then
Call SendMessage(room%, WM_CLOSE, 0, 0)
End If
tryagain:
Call AOL40_Keyword(URL$)
Call AOL40_FindChatRoom
If AOL40_FindChatRoom = 0 Then GoTo noroom:
GoTo FoundRoom:
noroom:
noroomagaiN:
AOL40_MDI
a = findchildbytitle(AOL40_MDI, "Sorry...")
b = FindChildByClass(a, "_AOL_Icon")
b = GetWindow(b, 2)
If AOL40_FindChatRoom <> 0 Then GoTo FoundRoom:
If b = 0 Then GoTo noroomagaiN:
ClickIcon (b)
GoTo tryagain:
FoundRoom:
MsgBox "room was found"
End Sub
Function AOL40_ChatLink(link, txt)
' this will send a link to the chat room so all the people in the chat
' room can click on it
' Call AOL40_SendChat("mailto:bonginokc", "")
AOL40_SendChat ("< a href=" + link + ">" + txt + "</a>")
End Function
Function AOL40_ChatLink2(link, txt)
' this is just for my prog ;)
' Call AOL40_ChatLink2("mailto:bonginokc", "")
AOL40_SendChat ("<b><--^v^ < a href=" + link + ">" + txt + "</a>")
End Function
Function AOL40_ChatToList(List As ListBox)
' this will but the whole chat into a listbox. this is one way you can scan chatlines for text
' Call AOL40_ChatToList(List1)
again:
ChatText$ = AOL40_GetChatText
If AOL40_GetChatText = "" Then GoTo again:
For FindChar = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
List.AddItem (TheChatText$)
thechars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(ChatText$, lastlen, Len(thechars$))
List.AddItem (LastLine)
End Function
Function AOL40_ClearChat()
' this is not one of those lamer tools because i will not let anything in
' bas deal with that shit, but this will only clear your chat box.
' Call AOL40_ClearChat
Dim dat%
Call AOL40_ChatHandles
RoomHandle = FindChildByClass(RoomHandle, "RICHCNTL")
X = SendMessageByString(RoomHandle, 12, 0, "")
End Function
Sub AOL40_CloseChat()
' this will close anychat room you are in
' Call AOL40_CloseChat
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
StuFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If StuFF% <> 0 And MoreStuff% <> 0 Then
Call SendMessage(room%, WM_CLOSE, 0, 0)
End If
End Sub
Sub AOL40_decripter(word$)
' i need to just work on this a little more it is almost done
' Call AOL40_decripter("" + Text1.Text + "")
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "¿" Then Leet$ = " "
    If letter$ = "†" Then Leet$ = "a"
    If letter$ = "õ" Then Leet$ = "b"
    If letter$ = "ð" Then Leet$ = "c"
    If letter$ = "ô" Then Leet$ = "d"
    If letter$ = "î" Then Leet$ = "e"
    If letter$ = "ï" Then Leet$ = "f"
    If letter$ = "ì" Then Leet$ = "g"
    If letter$ = "é" Then Leet$ = "h"
    If letter$ = "ê" Then Leet$ = "i"
    If letter$ = "ë" Then Leet$ = "j"
    If letter$ = "ä" Then Leet$ = "k"
    If letter$ = "å" Then Leet$ = "l"
    If letter$ = "â" Then Leet$ = "m"
    If letter$ = "ù" Then Leet$ = "n"
    If letter$ = "û" Then Leet$ = "o"
    If letter$ = "ü" Then Leet$ = "p"
    If letter$ = "ç" Then Leet$ = "q"
    If letter$ = "ñ" Then Leet$ = "r"
    If letter$ = "š" Then Leet$ = "s"
    If letter$ = "v" Then Leet$ = "t"
    If letter$ = "ÿ" Then Leet$ = "u"
    If letter$ = "Ø" Then Leet$ = "v"
    If letter$ = "£" Then Leet$ = "w"
    If letter$ = "¹" Then Leet$ = "x"
    If letter$ = "©" Then Leet$ = "y"
    If letter$ = "#" Then Leet$ = "z"
    ' upercase letters
    If letter$ = "Ï" Then Leet$ = "A"
    If letter$ = "Î" Then Leet$ = "B"
    If letter$ = "Í" Then Leet$ = "C"
    If letter$ = "ß" Then Leet$ = "D"
    If letter$ = "Ç" Then Leet$ = "E"
    If letter$ = "Å" Then Leet$ = "F"
    If letter$ = "Ä" Then Leet$ = "G"
    If letter$ = "Ã" Then Leet$ = "H"
    If letter$ = "Ð" Then Leet$ = "I"
    If letter$ = "Ë" Then Leet$ = "J"
    If letter$ = "S" Then Leet$ = "K"
    If letter$ = "&" Then Leet$ = "L"
    If letter$ = "Y" Then Leet$ = "M"
    If letter$ = "W" Then Leet$ = "N"
    If letter$ = ">" Then Leet$ = "O"
    If letter$ = "<" Then Leet$ = "P"
    If letter$ = "Š" Then Leet$ = "Q"
    If letter$ = "Û" Then Leet$ = "R"
    If letter$ = "+" Then Leet$ = "S"
    If letter$ = "=" Then Leet$ = "T"
    If letter$ = "@" Then Leet$ = "U"
    If letter$ = "Ñ" Then Leet$ = "V"
    If letter$ = "%" Then Leet$ = "W"
    If letter$ = "*" Then Leet$ = "X"
    If letter$ = "Õ" Then Leet$ = "Y"
    If letter$ = "~" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next q
End Sub
Sub AOL40_elitetalker(word$)
' this will take the word$ that you enter and turn them into
' kool letters then send it to the chat room
' Call AOL40_elitetalker("This will be in elite letters")
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next q
AOL40_SendChat (Made$)
End Sub
Sub AOL40_encripter(word$)
' i need to just work on this a little more it is almost done
' Call AOL40_encripter("" + Text1.Text + "")
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    ' lower case letters
    If letter$ = "a" Then Leet$ = "†"
    If letter$ = "b" Then Leet$ = "õ"
    If letter$ = "c" Then Leet$ = "ð"
    If letter$ = "d" Then Leet$ = "ô"
    If letter$ = "e" Then Leet$ = "î"
    If letter$ = "f" Then Leet$ = "ï"
    If letter$ = "g" Then Leet$ = "ì"
    If letter$ = "h" Then Leet$ = "é"
    If letter$ = "i" Then Leet$ = "ê"
    If letter$ = "j" Then Leet$ = "ë"
    If letter$ = "k" Then Leet$ = "ä"
    If letter$ = "l" Then Leet$ = "å"
    If letter$ = "m" Then Leet$ = "â"
    If letter$ = "n" Then Leet$ = "ù"
    If letter$ = "o" Then Leet$ = "û"
    If letter$ = "p" Then Leet$ = "ü"
    If letter$ = "q" Then Leet$ = "ç"
    If letter$ = "r" Then Leet$ = "ñ"
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "v"
    If letter$ = "u" Then Leet$ = "ÿ"
    If letter$ = "v" Then Leet$ = "Ø"
    If letter$ = "w" Then Leet$ = "£"
    If letter$ = "x" Then Leet$ = "¹"
    If letter$ = "y" Then Leet$ = "©"
    If letter$ = "z" Then Leet$ = "#"
    ' upercase letters
    If letter$ = "A" Then Leet$ = "Ï"
    If letter$ = "B" Then Leet$ = "Î"
    If letter$ = "C" Then Leet$ = "Í"
    If letter$ = "D" Then Leet$ = "ß"
    If letter$ = "E" Then Leet$ = "Ç"
    If letter$ = "F" Then Leet$ = "Å"
    If letter$ = "G" Then Leet$ = "Ä"
    If letter$ = "H" Then Leet$ = "Ã"
    If letter$ = "I" Then Leet$ = "Ð"
    If letter$ = "J" Then Leet$ = "Ë"
    If letter$ = "K" Then Leet$ = "S"
    If letter$ = "L" Then Leet$ = "&"
    If letter$ = "M" Then Leet$ = "Y"
    If letter$ = "N" Then Leet$ = "W"
    If letter$ = "O" Then Leet$ = ">"
    If letter$ = "P" Then Leet$ = "<"
    If letter$ = "Q" Then Leet$ = "Š"
    If letter$ = "R" Then Leet$ = "Û"
    If letter$ = "S" Then Leet$ = "+"
    If letter$ = "T" Then Leet$ = "="
    If letter$ = "U" Then Leet$ = "@"
    If letter$ = "V" Then Leet$ = "Ñ"
    If letter$ = "W" Then Leet$ = "%"
    If letter$ = "X" Then Leet$ = "*"
    If letter$ = "Y" Then Leet$ = "Õ"
    If letter$ = "Z" Then Leet$ = "~"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next q
'Form1.Label1.Caption
End Sub
Function AOL40_FindChatRoom()
' this will find the chat room and set the focus on the chat room
' Call AOL40_FindChatRoom
' or
' If AOL40_FindChatRoom = 0 then MsgBox "Please Enter A Chat Room Before using this function"
    AOL% = FindWindow("AOL Frame25", vbNullString)
    mdi% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(mdi%, 5)
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
        AOL40_FindChatRoom = firs%
        Exit Function
    End If
    AOL40_FindChatRoom = 0
End Function
Function AOL40_GetChatText()
' this will get all the chat text and make it into one string
' this is usually just used by other subs
' tex$ = AOL40_GetChatText
' MsgBox tex$
room% = AOL40_FindChatRoom
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText$ = AOL40_GetText(AORich%)
AOL40_GetChatText = ChatText$
End Function
Function AOL40_GetRoomName() As String
' this will get the Room Name
' tex$ = AOL40_GetRoomName
' MsgBox tex$
On Error Resume Next
AOL40_GetRoomName = GetApiText(AOL40_FindChatRoom())
End Function
Function AOL40_GetText(child)
' this will get the text of the "child" window
' AOL40_GetText (im$)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
AOL40_GetText = TrimSpace$
End Function
Sub AOL40_HideAOL()
' this will hide your AOL screen
' Call AOL40_HideAOL
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub
Public Sub AOL40_HideWelcome()
' this will get rid of your welcome window so it is not minimized in aol
' Call AOL40_HideWelcome
X = findchildbytitle(AOL40_MDI(), "Welcome, " & AOL40_UserSN & "!")
Call ShowWindow(X, SW_HIDE)
End Sub
Sub AOL40_IM(Recipiant, message)
' This will send a IM.
' Call AOL40_IM("", "This is the message")
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Call AOL40_Keyword("aol://9293:")
Do: DoEvents
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMwin%, "_AOL_Edit")
AORich% = FindChildByClass(IMwin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call Timeout(0.01)
ClickIcon (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
IMwin% = findchildbytitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMwin%, WM_CLOSE, 0, 0): Exit Do
If IMwin% = 0 Then Exit Do
Loop
End Sub
Sub AOL40_IMsOff()
' this will turn your ims off
' Call AOL40_IMsOff
Call AOL40_IM("$IM_OFF", "")
End Sub
Sub AOL40_IMsOn()
' this will turn your ims on
' Call AOL40_IMsOn
Call AOL40_IM("$IM_ON", "")
End Sub
Function AOL40_IsOnline()
' this will check to see if they are online
' If AOL40_IsOnline = 0 then MsgBox "Please Sign Onto AOL"
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome1% = findchildbytitle(mdi%, "Welcome,")
If welcome1% <> 0 Then
   AOL40_IsOnline = 1 ' They are online
Else:
   AOL40_IsOnline = 0 ' They are notonline
End If
End Function
Sub AOL40_Keyword(TheKeyWord As String)
' This will go to the keyword specified buy insterting it into
' the toolbar of aol
' Call AOL40_Keyword("Help Me")
Dim tool%
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass2(AOL%, "AOL Toolbar")
tool% = FindChildByClass2(tool%, "_AOL_Toolbar")
tool% = FindChildByClass2(tool%, "_AOL_Combobox")
tool% = FindChildByClass2(tool%, "Edit")
Call SendMessageByString(tool%, 12, 0, txt)
Call SendMessageByNum(tool%, WM_CHAR, VK_SPACE, 0)
Call SendMessageByNum(tool%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub AOL40_KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
' Call AOL40_KillGlyph
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Function AOL40_LastChatLine()
' This will grab the lastchat line from the room you are in
' Call AOL40_LastChatLine
chatline$ = AOL40_LastChatLineWithSN
If chatline$ = "" Then Exit Function
ChatTrim$ = Left$(chatline$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
screenname$ = SN
ChatTrimNum = Len(screenname$)
ChatTrim$ = Mid$(chatline$, ChatTrimNum + 4, Len(chatline$) - Len(screenname$))
AOL40_LastChatLine = ChatTrim$
End Function
Function AOL40_LastChatLineWithSN()
' this will get the last chat line with sn
' Call AOL40_LastChatLineWithSN
ChatText$ = AOL40_GetChatText
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
AOL40_LastChatLineWithSN = LastLine
End Function
Sub AOL40_LocateMember(SN)
' this will open a screen telling you where a member is if they are online
' Call AOL40_locateMember("")
Call AOL40_Keyword("aol://3548:" & SN)
End Sub

Function AOL40_MDI()
' This is just soemthin for all my subs so i didnt have to keep writing it out
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL40_MDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function AOL40_RoomCount()
' this will count the number of people in the chatroom
' num$ = AOL40_RoomCount
' msgbox num$
Dim chat%
chat% = AOL40_FindChatRoom()
List% = FindChildByClass(chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOL40_RoomCount = Count%
End Function
Public Function AOL40_RoomFull()
' this will click that icon that pops up saying that the room
' you requested is full.
' Call AOL40_RoomFull
Do
Timeout 0.00002
msg1% = FindWindow("#32770", "America Online")
Button% = FindChildByClass(msg1%, "Button")
stat% = FindChildByClass(msg1%, "Static")
statcap% = findchildbytitle(msg1%, "The room you requested is full.")
If stat% <> 0 And Button% <> 0 And statcap% <> 0 Then Call ClickIcon(Button%)
Loop Until msg1% = 0
End Function
Function AOL40_SecondToLastChatLineWithSN()
' this will get the secong to last chat line with sn
' Call AOL40_SecondToLastChatLineWithSN
ChatText$ = AOL40_GetChatText
For FindChar = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
AOL40_SecondToLastChatLineWithSN = TheChatText$
End Function
Sub AOL40_SendChat(chat)
' this will send chat to the chat room
' Call AOL40_SendChat("")
Do
room% = AOL40_FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
Z = GetWinText(AORich%)
Timeout 0.05
Loop Until Z = chat
DoEvents
Do
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Z = GetWinText(AORich%)
Timeout 0.00001
Loop Until Z = ""
End Sub
Sub AOL40_ShowAOL()
' this will show your AOL screen after it has been hidden
' Call AOL40_ShowAOL
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Public Sub AOL40_ShowWelcome()
' this will show the welcome screen if it is hidden
' Call AOL40_ShowWelcome
X = findchildbytitle(AOL40_MDI(), "Welcome, " & AOL40_UserSN & "!")
Call ShowWindow(X, SW_SHOW)
End Sub
Function AOL40_SNFromLastChatLine()
' this will grab the screen name from the lastchatline
' Call AOL40_SNFromLastChatLine
ChatText$ = AOL40_LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
AOL40_SNFromLastChatLine = SN
End Function
Sub AOL40_strangeim(StuFF)
'This sends someone blank IMs with different colors
'in each one. It sends 5 IMs but then it loops so
'you better add a stop button!
Do:
DoEvents
Call AOL40_IM(StuFF, "<body bgcolor=#000000>")
Call AOL40_IM(StuFF, "<body bgcolor=#0000FF>")
Call AOL40_IM(StuFF, "<body bgcolor=#FF0000>")
Call AOL40_IM(StuFF, "<body bgcolor=#00FF00>")
Call AOL40_IM(StuFF, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub
Sub AOL40_TextSound(Wav$, text$)
' this will send text to the chat with invisible sound
' Call AOL40_TextSound("{S GotMail", "Got Mail")
Call AOL40_SendChat(text$ & " <font color=#fffffe> " & Wav$)
End Sub
Sub AOL40_UnUpChat()
' this will undo the upchat window
' Call AOL40_UnUpChat
aom% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(aom%, SW_RESTORE)
X = ShowWindow(aom%, SW_SHOW)
X = SetFocusAPI(aom%)
End Sub
Function AOL40_UpChat()
'this is an upchat that minimizes the
'upload window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
aom% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(aom%, SW_MINIMIZE)
X = SetFocusAPI(AOL%)
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Function
Function AOL40_UserSN()
' this will get the usersn from the welcome screen.
' Call AOL40_UserSN
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome1% = findchildbytitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome1%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome1%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL40_UserSN = User
End Function
Function AOL40_Win()
' This is used by a couple other of my subs
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL40_Win = AOL%
End Function
Function AOLVersion()
' this will find out what version of AOL they are on.
' If AOLVersion = 4 Then MsgBox "You are on AOL 4.0"
' If AOLVersion = 3 Then MsgBox "You are on AOL 3.0"
hMenu% = GetMenu(AOL40_Win())
submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3 ' They are on AOL 3.0
Else
AOLVersion = 4 ' They are on AOL 4.0
End If
End Function
Function BlackBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackBlueBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackGreenBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackGreyBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackPurpleBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackRedBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlackYellowBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueBlackBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueGreenBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BluePurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BluePurpleBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueRedBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BlueYellowBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function

Sub ClickIcon(icon%)
' this will click on the icon you specify
' Call ClickIcon (icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub CloseWin(wind)
' this will close the window you want
' Call closewin(im%)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
End Sub
Function EnterKey()
EnterKey = CStr(Chr(13) & Chr(10))
End Function

Function FindChildByClass2(parenthwnd, childhand)
' this is just used by my functions
Dim returnstring$, handles%, Parent, copy
copy = parenthwnd
Parent = GetWindow(parenthwnd, 5)
Top:
returnstring$ = String$(250, 0): handles% = GetClassName(Parent, returnstring$, 250)
If Left$(returnstring$, Len(childhand)) Like childhand Then GoTo ending:
Parent = GetWindow(Parent, 2)
If Parent > 0 Then GoTo Top
ending:
FindChildByClass2 = Parent
parenthwnd = copy
End Function

Public Function FindChildByTitle2(parenthwnd, childhand$)
' this is just used by my functions
Dim copy
copy = parenthwnd
GetCaption (parenthwnd)
Caption$ = Left$(Caption$, Len(childhand$))
If Caption$ = childhand$ Then GoTo foundit
Do While parenthwnd <> 0 And Caption$ <> childhand$
parenthwnd = GetWindow(parenthwnd, 2)
GetCaption (parenthwnd)
Caption$ = Left$(Caption$, Len(childhand$))
Loop
foundit:
FindChildByTitle2 = parenthwnd
parenthwnd = copy
End Function
Sub First_Time_Load()
' The vesru first time this program loads, a msgbox appears saying
' that this is there first time to load your prog
' Call First_Time_Load
If Len(dir(App.Path + "\" + "first.txt")) = 0 Then
 ' This is what it says when the program is first loaded
MsgBox "This is the first time your running this program"
Open (App.Path + "\" + "first.txt") For Append As #1
Print #1, "Hey!"
Close #1
Else
Open (App.Path + "\" + "first.txt") For Append As #1
Print #1, "Hey!"
Close #1
End If
End Sub
Function FixAPIString(sText As String) As String
' this is just used by most of my subs
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function


Sub Form_ScrollDown(frm As Form, finished)
' This will make the form slowly scroll down
' you can add a timeout to make it go slower
' or faster
' Call Call Form_ScrollDown(Form1, 1000)
If frm.Height > finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) + 1
Loop Until frm.Height = finished
End Sub
Sub Form_ScrollUp(frm As Form, finished)
' This will make the form slowly scroll up
' you can add a timeout to make it go slower
' or faster
' Call Form_ScrollUp(Form1, 1000)
If frm.Height < finished Then Exit Sub
If frm.Height = finished Then Exit Sub
Do
frm.Height = Val(frm.Height) - 1
Loop Until frm.Height = finished
End Sub
Sub FormDance(M As Form)
' this will make the form move back and forth
' Call FormDance(Form1)
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 5000
Timeout (0.1)
M.Left = 4000
Timeout (0.1)
M.Left = 3000
Timeout (0.1)
M.Left = 2000
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 5
Timeout (0.1)
M.Left = 400
Timeout (0.1)
M.Left = 700
Timeout (0.1)
M.Left = 1000
Timeout (0.1)
M.Left = 2000
End Sub
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Private Function GetAOLProcessHandle(ByVal hwnd As Long) As Long
' this is used by most of my other subs
On Error Resume Next
Dim m_AOLThreadId As Long
Dim m_AOLProcessID As Long
m_AOLThreadId = GetWindowThreadProcessId(hwnd, m_AOLProcessID)
GetAOLProcessHandle = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, m_AOLProcessID)
End Function
Function GetApiText(hwnd As Integer) As String
X = SendMessageByNum(hwnd%, WM_GETTEXTLENGTH, 0, 0)
    text$ = Space(X + 1)
    X = SendMessageByString(hwnd%, WM_GETTEXT, X + 1, text$)
    GetApiText = FixAPIString(text$)
End Function
Function GetCaption(hwnd)
' this is just a sub that is used in most of my other subs.
' GetCaption(win%)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Function GetClass(child)
' this is used by almost every sub. it will tell you the class of a window
' Text$ = GetClass(win%)     'this will get the class of win% and make it = text$
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Function GetWinText(hwnd As Integer) As String
' this is used by almost every sub.
lentos = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
X = SendMessageByString(hwnd, WM_GETTEXT, lentos + 1, Buffer$)
GetWinText = Buffer$
End Function
Function GreenBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenBlackGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenBlueGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenPurpleGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenRedGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreenYellowGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & H & "></b>" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 220 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyBlackGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyBlueGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyGreenGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b> " + Msg + "")
End Function
Function GreyPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyPurpleGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyRedGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function GreyYellowGrey(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 490 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Sub KillComboDupes(Cmb As Control)
For i = 0 To Cmb.ListCount - 1
For E = 0 To Cmb.ListCount - 1
If LCase(Cmb.List(i)) Like LCase(Cmb.List(E)) And i <> E Then
Cmb.RemoveItem (E)
End If
Next E
Next i
End Sub
Sub KillListDupes(lst As Control)
On Error Resume Next
For i = 0 To lst.ListCount - 1
For E = 0 To lst.ListCount - 1
If LCase(lst.List(i)) Like LCase(lst.List(E)) And i <> E Then
lst.RemoveItem (E)
End If
Next E
Next i
End Sub
Function List_Count(lst As ListBox)
X = lst.ListCount
List_Count = X
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Function PlayAvi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = MciSendString("play c:\windows\help\scroll.avi", 0&, 0, 0)
End Function
Function PlayMidi()
'Plays a Midi File Change the path Below to your
'Midi Path
lRet = MciSendString("play C:\INAGODA.mid", 0&, 0, 0) ' or whatever the File Name is
End Function
Sub PlayWav(File)
Dim X%
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub Prevent()
' Only Allows one version of you progg to run at a time like AOL
' Call Prevent
If App.PrevInstance Then End
End Sub
Function PurpleBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function BoldPurpleBlackPurple(text As String)
    a = Len(text)
    For b = 1 To a
        C = Left(text, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat (Msg)
End Function
Function PurpleBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function PurpleBluePurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function PurpleGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function PurpleGreenPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    PurpleGreenPurple = ("<b>" + Msg + "")
End Function
Function PurpleRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function PurpleRedPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    PurpleRedPurple = ("<b>" + Msg + "")
End Function
Function PurpleYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function PurpleYellowPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RandomNumber(finished As Integer)
' this will get a random number from the number FINISHED to 0
' Call RandomNumber (11)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Function Range(Lower As Integer, Upper As Integer) As Integer
' this will return the number between 2 and 5 and make it num$
' num$ = Range(2, 5)
Randomize
Range% = Int((Upper% - Lower% + 1) * Rnd + Lower%)
End Function
Function RedBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedBlackRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedBlueRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & H & "></b>" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedGreenRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedPurpleRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RedYellowRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
   AOL40_SendChat ("<b>" + Msg + "")
End Function
Function RemoveSpace(txt$) As String
' this will make the remove any spaces in a string
' num$ = RemoveSpace("all spaces will be removed")
' MsgBox num$
NoSpace$ = txt$
While InStr(NoSpace$, " ") <> 0
Where = InStr(NoSpace$, " ")
NoSpace$ = Mid(NoSpace$, 1, Where - 1) + Mid(NoSpace$, Where + 1)
Wend
RemoveSpace = NoSpace$
End Function
Function ReverseText(text)
' This will return the words backwards
' tex$ = ReverseText("brad")
' MsgBox tex$
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words
End Function
Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function
Function SpaceCase(text As String) As String
' this will make the remove any spaces in a string but also will make every letter uppercase
' num$ = SpaceCase("all spaces will be removed and all letters will be uppercase")
' MsgBox num$
txt$ = text$
txt$ = Trim(UCase(RemoveSpace(txt$)))
SpaceCase = txt$
End Function

Function TrimTime()
' this will take the seconds off of the time
' txt$ = TrimTime
' MsgBox txt$
bb$ = Left$(Time$, 5)
HourH$ = Left$(bb$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(bb$, 3) & " " & Ap$
End Function
Sub WAVLoop(File)
' This will play a wav file and keep looping it till WayStop is called
' Call WAVPlay("C:\WINDOWS\MEDIA\Tada.wav")
SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub WAVPlay(File)
' This will play a wav file
' Call WAVPlay("C:\WINDOWS\MEDIA\Tada.wav")
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub WAVStop()
' This is the only way that you can stop the wavloop
' Call WAVStop
Call WAVPlay(" ")
End Sub
Function Wavy(TheText As String)
' this will make the chat wavy
' Call WavY("This Chat Will be  wavy")
G$ = TheText
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next W
Wavy = p$
End Function
Sub WavyChatBlueBlack(TheText)
' this will faded and make the chat wavy
' Call WavyChatBlueBlack("This Chat Will be faded, wavy")
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
AOL40_SendChat (p$)
End Sub
Function WavYChaTRedBlue(TheText As String)
' this will faded and make the chat wavy
' Call WavYChaTRedBlue("This Chat Will be faded, wavy")
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
WavYChaTRedBlue = p$
End Function
Function WavYChaTRedGreen(TheText As String)
' this will faded and make the chat wavy
' Call WavYChaTRedGreen("This Chat Will be faded, wavy")G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
WavYChaTRedGreen = p$
End Function
Function WinCaption(win)
Dim GetWinText%
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Function Window_Close(win)
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Function
Sub Window_Hide(win)
' this will hide the window you specify
' Call Window_Hide(im%)
X = ShowWindow(win, SW_HIDE)
End Sub
Sub Window_Maximize(win)
' this will Maximize the win that you specify, you will need to find out
' the api of the window you want to Maximize
' Call Window_Maximize(im%)
X = ShowWindow(win, SW_MAXIMIZE)
End Sub
Sub Window_Minimize(win)
' this will minimize the win that you specify, you will need to find out
' the api of the window you want to minimize
' call Window_Minimize(im%)
X = ShowWindow(win, SW_MINIMIZE)
End Sub
Sub Window_Restore(win)
' this will Restore the window that is minimized
' Call Window_Restore(im%)
X = ShowWindow(win, SW_RESTORE)
End Sub
Sub Window_Show(win)
' this will show the window that is hidden
' Call Window_Show(im%)
X = ShowWindow(win, SW_SHOW)
End Sub
Function YellowBlack(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowBlackYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowBlue(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowBlueYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
  AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowGreen(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
    AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowGreenYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowPurple(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowPurpleYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowRed(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 255 / a
        F = E * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
 AOL40_SendChat ("<b>" + Msg + "")
End Function
Function YellowRedYellow(Text1)
a = Len(Text1)
    For b = 1 To a
        C = Left(Text1, b)
        d = Right(C, 1)
        E = 510 / a
        F = E * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next b
AOL40_SendChat ("<b>" + Msg + "")
End Function
