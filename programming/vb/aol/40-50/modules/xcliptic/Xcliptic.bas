Attribute VB_Name = "Module1"
' ------------------------------------
'| Xcliptic Bas By: Xcliptic for AOL 4.0 |
' ------------------------------------
'Version 1
'32 Bit Version of AOL4.0 Only!
'E-mail me at Xcliptic99@aol.com
'This is my very first bas file i made
'its pretty big file and took me quite a
'long time to make,well i hope i see
'it on alot of diffrent sites and people
'enjoy the subs.Most of them work but
'there is always some subs on bases that
'doent work so its not just mine,but 97%
'of the subs work,so enjoy and have fun
'I am not responsible for anything
'That you do with it and any
'Damage that it may cause to you,
'Your computer, or your aol account.

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const MF_ENABLED = &H0&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
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
Public Const WM_USER = &H400
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
Const WM_PSD_ENVSTAMPRECT = WM_USER + 5 'Subclass The Chat Window
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUared = &HF0000
Public Const ENTER_KEY = 13
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesaredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Type POINTAPI
        x As Long
        Y As Long
End Type
Type TheRgb
   Red As Long
   Green As Long
   Blue As Long
End Type

Function AOLWindow() As Long
'By: Xcliptic
AOLWindow& = FindWindow("AOL Frame25", vbNullString)
End Function
Function AOLMDI() As Long
'BY: Xcliptic
AOLMDI& = FindWindowEx(AOLWindow, 0&, "MDIClient", vbNullString)
End Function
Sub Key(Keywd As String)
'By: Xcliptic
'keyword
Tool1& = FindWindowEx(AOLWindow, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(Tool1&, 0&, "_AOL_Toolbar", vbNullString)
Box& = FindWindowEx(tool2&, 0&, "_AOL_Combobox", vbNullString)
Box& = FindWindowEx(Box&, 0&, "Edit", vbNullString)
SendMessageByString Box&, WM_SETTEXT, 0&, Keywd
SendMessageLong Box&, WM_CHAR, VK_SPACE, 0&
SendMessageLong Box&, WM_CHAR, VK_RETURN, 0&
End Sub
Sub SendIM(person As String, Message As String, Optional Bold As Boolean, Optional Italics As Boolean, Optional Underline As Boolean, Optional Strikeout As Boolean)
'bold, italics, strikeout, and
'underlined are optional statements
'You could have IM(ScreenName,Message)
'or IM(Screenname,message,true,true,true,true)
'By: Xcliptic
If Bold = True Then Text = "<b>" & Text & "</b>"
If Italics = True Then Text = "<i>" & Text & "</i>"
If Strikeout = True Then Text = "<s>" & Text & "</s>"
If Underline = True Then Text = "<u>" & Text & "</u>"
Call Key("aol://9293:" & person)
Do: DoEvents
IMWin& = FindWindowEx(AOLMDI, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
SendButton& = FindWindowEx(IMWin&, 0&, "_AOL_Icon", vbNullString)
For x = 1 To 8
SendButton& = FindWindowEx(IMWin&, SendButton&, "_AOL_Icon", vbNullString)
Next
Loop Until IMWin& <> 0 And Rich& <> 0 And SendButton& <> 0
WindowHide IMWin&
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message)
Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
IMWin& = FindWindowEx(AOLMDI, 0&, "AOL Child", "Send Instant Message")
Loop Until okwin& <> 0& Or IMWin& = 0&
If okwin& <> 0& Then
Call SendMessage(okwin&, WM_CLOSE, 0&, 0&)
Call SendMessage(IMWin&, WM_CLOSE, 0&, 0&)
Exit Sub
End If
Do Until U& <> 0
U& = FindChildByTitle(AOLMDI, "  Instant Message To:")
Loop
Pause 0.5
SendMessageLong U&, WM_CLOSE, 0&, 0&
End Sub
Sub IMsOn()
'By: Xcliptic
'you can change the wording in between the" "
'but i would like if you keep it the same and
'not steal my code
Call SendIM("$IM_ON", "Xcliptic is the best ")
End Sub
Sub IMsOff()
'By: Xcliptic
'you can change the wording in between the" "
'but i would like if you keep it the same and
'not steal my code
Call SendIM("$IM_OFF", "Xcliptic is the best ")
End Sub
Sub AddRoomToListBox(TheList As ListBox, AddUser As Boolean)
'Borrowed this sub From Dos32.bas
'I would like to say thanx to Dos32.
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChatRoom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> GetUserSn Or AddUser = True Then
If ListSearch(TheList, ScreenName$) = -1 Then TheList.AddItem ScreenName$
End If
Next index&
Call CloseHandle(mThread)
End If
End Sub

Sub AddRoomToComboBox(TheCombo As ComboBox, AddUser As Boolean)
'This is another sub borrowed from Dos32.
'thanx again.
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChatRoom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> GetUserSn Or AddUser = True Then
TheCombo.AddItem ScreenName$
End If
Next index&
Call CloseHandle(mThread)
End If
If TheCombo.ListCount > 0 Then
TheCombo.Text = TheCombo.List(0)
End If
End Sub
Function FindChatRoom() As Long
'By: Xcliptic
child& = FindWindowEx(AOLMDI, 0&, "AOL Child", vbNullString)
Do: DoEvents
TheList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AOLCombo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
If AOLIcon& <> 0 And Rich& <> 0 And AOLStatic& <> 0 And TheList& <> 0 And AOLCombo& <> 0 Then Exit Do
child& = FindWindowEx(AOLMDI, child&, "AOL Child", vbNullString)
Loop Until child& = 0
FindChatRoom& = child&
End Function
Function GetRoomName() As String
'name of the room
'By: Xcliptic
GetRoomName = GetCaption(FindChatRoom)
End Function
Function GetCaption(WindowHandle As Long) As String
'Well here it is again another sub from Dos32
'thanx again Dos
Dim Buffer As String, TextLength As Long
TextLength& = GetWindowTextLength(WindowHandle&)
Buffer$ = String(TextLength&, 0&)
Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
GetCaption$ = Buffer$
End Function
Function GetClass(child)
'Well here it is again another sub from Dos32
'thanx again Dos
Dim Buffer$
Dim getclas%

Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Sub SendChat(Text As String, Optional Bold As Boolean, Optional Italics As Boolean, Optional Underline As Boolean, Optional Strikeout As Boolean, Optional Wavy As Boolean)
'bold, italics, strikeout, and
'underlined are optional statements
'You could have sendchat(text)
'or sendchat(text,true,true,true,true)
'By: Xcliptic
If Wavy = True Then Text = WavyChat(Text)
If Bold = True Then Text = "<b>" & Text & "</b>"
If Italics = True Then Text = "<i>" & Text & "</i>"
If Strikeout = True Then Text = "<s>" & Text & "</s>"
If Underline = True Then Text = "<u>" & Text & "</u>"
thechat& = FindChatRoom
If thechat& = 0 Then Exit Sub
Box = FindWindowEx(thechat&, 0&, "RICHCNTL", vbNullString)
Box = FindWindowEx(thechat&, Box, "RICHCNTL", vbNullString)
SendMessageByString Box, WM_SETTEXT, 0&, ""
SendMessageByString Box, WM_SETTEXT, 0&, Text
SendMessageLong Box, WM_CHAR, 13, 0&
Pause 0.33
End Sub
Sub PrivateRoom(Room As String)
'enter a private room
'By: Xcliptic
Call Key("aol://2719:2-2-" & Room)
End Sub
Sub WindowHide(hWnd As Long)
'hides a window
'By: Xcliptic
Call ShowWindow(hWnd&, SW_HIDE)
End Sub

Sub WindowShow(hWnd As Long)
'shows a window
'By: Xcliptic
Call ShowWindow(hWnd&, SW_SHOW)
End Sub
Function GetText(WindowHandle As Long) As String
'By: Xcliptic
TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
GetText$ = Buffer$
End Function
Function GetchatText() As String
'gets chat text
'By: Xcliptic
Box& = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
GetchatText = GetText(Box)
End Function
Function GetChatBox() As Long
'handle of chat box
'By:Xcliptic
Box = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
GetChatBox = FindWindowEx(FindChatRoom, Box, "RICHCNTL", vbNullString)
End Function
Function LastChatLineWithSN() As String
'By: Xcliptic
TheStr = GetchatText
For T = Len(TheStr) To 1 Step -1
If Mid(TheStr, T, 1) = Chr(13) Then Exit For Else totalstr = Mid(TheStr, T, 1) & totalstr
Next
LastChatLineWithSN = totalstr
End Function
Function LastChatLine() As String
'last chat line
'By:Xcliptic
TheStr = LastChatLineWithSN
For T = Len(TheStr) To 1 Step -1
If Mid(TheStr, T, 1) = ":" Then Exit For Else totalstr = Mid(TheStr, T, 1) & totalstr
Next
LastChatLine = Mid(totalstr, 3, Len(totalstr))
End Function
Function LastChatSn() As String
'screen name from last chat line
'By:Xcliptic
TheStr = LastChatLineWithSN
For T = 1 To Len(TheStr)
If Mid(TheStr, T, 1) = ":" Then Exit For Else totalstr = totalstr & Mid(TheStr, T, 1)
Next
LastChatSn = totalstr
End Function
Function ListSearch(Lst As ListBox, Txt As String) As Integer
'exact search for a list
'By:Xcliptic
Dim x%
For x% = 0 To Lst.ListCount - 1
If UCase(Txt) = UCase(Lst.List(x%)) Then
ListSearch = x%
Exit Function
End If
Next x%
ListSearch = -1
End Function
Sub IMAnswer(SN As String, Message As String)
'By:Xcliptic
If SN = "*" Then GoTo doall
thewin = FindWindowEx(AOLMDI, 0&, "AOL Child", ">Instant Message from: " & SN)
If thewin = 0 Then Exit Sub Else
Text = FindWindowEx(thewin, 0&, "RICHCNTL", vbNullString)
typein = FindWindowEx(thewin, Text, "RICHCNTL", vbNullString)
If typein = 0 Then GoTo openit
typeinit:
typein = FindWindowEx(thewin, Text, "RICHCNTL", vbNullString)
SendMessageByString typein, WM_SETTEXT, 0&, Message
butt = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
For Butto = 1 To 8
butt = FindWindowEx(thewin, butt, "_AOL_Icon", vbNullString)
Next
ClickIcon (butt)
Exit Sub
openit:
g = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
g = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
ClickIcon (g)
GoTo typeinit
doall:
Do
thewin = FindChildByTitle(AOLMDI, ">Instant Message from: ")
If thewin = 0 Then Exit Sub
Text = FindWindowEx(thewin, 0&, "RICHCNTL", vbNullString)
typein = FindWindowEx(thewin, Text, "RICHCNTL", vbNullString)
If typein = 0 Then GoTo openit2
typeinit2:
typein = FindWindowEx(thewin, Text, "RICHCNTL", vbNullString)
SendMessageByString typein, WM_SETTEXT, 0&, Message
butt = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
For Butto = 1 To 8
butt = FindWindowEx(thewin, butt, "_AOL_Icon", vbNullString)
Next
ClickIcon (butt)
GoTo loopIT
openit2:
g = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
g = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
ClickIcon (g)
GoTo typeinit2
loopIT:
Loop
End Sub
Sub ClickIcon(hWnd As Long)
'By:Xcliptic
Call SendMessage(hWnd, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(hWnd, WM_LBUTTONUP, 0&, 0&)
End Sub
Function FindChildByTitle(ParenthWnd As Long, childtitle As String, Optional starthwnd As Long)
'I made this so I could find child
'windows without knowing the exact
'title.
'If you know the exact title, you can
'use this or FindWindowEx
'By:Xcliptic
g& = FindWindowEx(ParenthWnd, 0&, vbNullString, vbNullString)
If InStr(UCase(GetCaption(g&)), UCase(childtitle)) Then FindChildByTitle = g&: Exit Function
g& = starthwnd
Do
DoEvents
g& = FindWindowEx(ParenthWnd, g&, vbNullString, vbNullString)
If InStr(GetCaption(g&), childtitle) Then Exit Do
Loop Until g = 0
FindChildByTitle = g&
End Function
Function GetUserSn() As String
'gets user's screen name
'By:Xcliptic
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
If WelcomeWin& = 0& Then GetUserSn$ = "Not Online": Exit Function
UserSN$ = GetCaption(WelcomeWin&)
GetUserSn$ = Mid(UserSN$, 10, Len(UserSN$) - 10)
End Function
Sub AddBuddyListToList(TheList As ListBox)
'By:Xcliptic
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChildByTitle(AOLMDI, "Buddy List Window")
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> GetUserSn$ Or AddUser = True Then
If ListSearch(TheList, ScreenName$) = -1 Then TheList.AddItem ScreenName$
End If
Next index&
Call CloseHandle(mThread)
End If
End Sub
Function ReverseString(TheStr As String) As String
'By:Xcliptic
'AOL would return LOA
For g = 1 To Len(TheStr)
newstr = Mid(TheStr, g, 1) & newstr
Next
ReverseString = newstr
End Function
Function TrimSpaces(TheStr As String) As String
'By:Xcliptic
'gets rid of all spaces in a string.
For g = 1 To Len(TheStr)
If Mid(TheStr, g, 1) = " " Then Else newstr = newstr & Mid(TheStr, g, 1)
Next
TrimSpaces = newstr
End Function
Function TrimChar(TheStr As String, TheChar As String) As String
'trims a char from a string
'By:Xcliptic
If Len(TheChar) <> 1 Then Exit Function
For g = 1 To Len(TheStr)
If Mid(TheStr, g, 1) = TheChar Then Else newstr = newstr & Mid(TheStr, g, 1)
Next
TrimChar = newstr
End Function
Function TrimString(Text As String, TheStr As String) As String
'takes one string out of a string.
'ex
'trimstring("Xcliptic is the best!","the")
'returns "Xcliptic is best!"
'By:Xcliptic
For g = 1 To Len(Text)
If Mid(Text, g, Len(TheStr)) = TheStr Then g = g + Len(TheStr) - 1 Else newstr = newstr & Mid(Text, g, 1)
Next
TrimString = newstr
End Function
Function ReplaceString(Text As String, Replace As String, ReplaceWith As String)
'By:Xcliptic
For g = 1 To Len(Text)
If Mid(Text, g, Len(Replace)) = Replace Then newstr = newstr & ReplaceWith: g = g + Len(Replace) - 1 Else newstr = newstr & Mid(Text, g, 1)
Next
ReplaceString = newstr
End Function
Function SpacedString(Text As String) As String
'if you had "Xcliptic" as the text
'it would return "X c l i p t i c"
'By:Xcliptic
For x = 1 To Len(Text)
newstr = newstr & Mid(Text, x, 1) & " "
Next
SpacedString = Mid(newstr, 1, Len(newstr) - 1)
End Function
Sub Mail(person As String, subject As String, Message As String)
'sends mail
'By:Xcliptic
Tool& = FindWindowEx(AOLWindow, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(Tool&, 0&, "_AOL_Toolbar", vbNullString)
mailbutton& = FindWindowEx(tool2, 0&, "_AOL_Icon", vbNullString)
mailbutton& = FindWindowEx(tool2, mailbutton&, "_AOL_Icon", vbNullString)
Call SendMessage(mailbutton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(mailbutton&, WM_LBUTTONUP, 0&, 0&)
DoEvents
win& = 0
Do
win& = FindWindowEx(AOLMDI, win&, vbNullString, "Write Mail")
edit1& = FindWindowEx(win&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(win&, 0&, "RICHCNTL", vbNullString)
Edit& = FindWindowEx(win&, edit1&, "_AOL_Edit", vbNullString)
Edit& = FindWindowEx(win&, Edit&, "_AOL_Edit", vbNullString)
fCombo& = FindWindowEx(win&, 0&, "_AOL_FontCombo", vbNullString)
Combo& = FindWindowEx(win&, 0&, "_AOL_Combobox", vbNullString)
Check& = FindWindowEx(win&, 0&, "_AOL_Checkbox", vbNullString)
Sleep 10
Loop Until Check& <> 0 And Combo& <> 0 And fCombo& <> 0 And win& <> 0 And Edit& <> 0 And edit1& <> 0 And Rich& <> 0
butt& = 0
Sleep 30
For findsend = 1 To 14
butt& = FindWindowEx(win&, butt&, "_AOL_Icon", vbNullString)
Next
Call SendMessageByString(edit1&, WM_SETTEXT, 0, person$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
DoEvents
Call SendMessageByString(Edit&, WM_SETTEXT, 0, subject$)
ClickIcon (butt&)
End Sub
Sub LinkToChat(URL As String, urlname As String)
'sends a link to chat
'By:Xcliptic
g$ = "< a href=""" & URL & """>" & urlname & "</a>"
SendChat g$, False, False, False, False
End Sub
Sub MailToLink(mailto As String, Linktext As String)
'puts a mailto link in the chatroom
'By:Xcliptic
g$ = "< a href=""mailto:" & mailto & """>" & Linktext & "</a>"
SendChat g$, False, False, False, False
End Sub
Sub SpiralChat(Text As String)
'example AOL
'aol
'ola
'lao
'aol
'By:Xcliptic
For b = 1 To Len(Text)
H$ = Mid(Text, b, Len(Text)) & Mid(Text, 1, b - 1)
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & H$ & "</b><font color=#000000> }–•"
Pause 0.5
Next
Pause 0.5
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & Text & "</b><font color=#000000> }–•"
End Sub
Sub SpiralIM(person As String, Text As String)
'same as spiral chat only works with Imz
'By:Xcliptic
If Len(Text) * (Len(Text) + 1) > 592 Then MsgBox "Message to long!", vbOKOnly, App.Title
For b = 1 To Len(Text)
H$ = H$ & Chr(13) & Mid(Text, b, Len(Text)) & Mid(Text, 1, b - 1)
Next
H$ = H$ & Chr(13) & Text
SendIM person, H$, False, False, False, False
End Sub

Function WavyChat(Text As String)
'makes wavy text for the chatroom,
'im or mail.
'By:Xcliptic
Dim e As String
For x = 1 To Len(Text)
wave = wave + 1
If wave > 4 Then wave = 1
If wave = 4 Then wavetext = "</sub>"
If wave = 3 Then wavetext = "<sub>"
If wave = 2 Then wavetext = "</sup>"
If wave = 1 Then wavetext = "<sup>"
g$ = g$ & Mid(Text, x, 1) & wavetext
Next
WavyChat = g$
End Function
Sub WebSearch(search As String)
'searches the internet for a string
'By:Xcliptic
Key "http://search.yahoo.com/bin/search?p=" & search
End Sub
Sub ChatExtender(Text As String)
'lets you type more than the 92 char
'limit
'By:Xcliptic
If Len(Text) < 92 Then SendChat Text, False, False, False, False: Exit Sub
F = Int(Len(Text) / 92) + 1
For x = 0 To F
SendChat Mid(Text, x * 92 + 1, 92), False, False, False, False
Pause 0.5
Next
End Sub
Function BoldRotate(Text As String)
'makes text bold, not bold...
'By:Xcliptic
For x = 1 To Len(Text)
reg = reg + 1
If reg > 2 Then reg = 1
If reg = 1 Then html = "<b>"
If reg = 2 Then html = "</b>"
If Mid(Text, x, 1) = " " Then BoldRotate = BoldRotate & " ": reg = reg + 1: GoTo loopIT
BoldRotate = BoldRotate & Mid(Text, x, 1) & html
loopIT:
Next
End Function
Function BoldItalicRotate(Text As String)
'makes text bold then italic then nothing
'By:Xcliptic
For x = 1 To Len(Text)
reg = reg + 1
If reg > 4 Then reg = 1
If reg = 1 Then html = "<b>"
If reg = 2 Then html = "</b>"
If reg = 3 Then html = "<i>"
If reg = 4 Then html = "</i>"
If Mid(Text, x, 1) = " " Then BoldItalicRotate = BoldItalicRotate & " ": reg = reg + 1: GoTo loopIT
BoldItalicRotate = BoldItalicRotate & Mid(Text, x, 1) & html
loopIT:
Next
End Function
Sub Macro(Text As textbox)
'scrolls a multilined textbox to
'the chatroom one line at a time
'By:Xcliptic
For ctri% = 1 To CountTextBoxLines(Text)
g$ = GetTextBoxLine(Text, ctri%)
SendChat g$, False, False, False, False
Pause 0.5
Next
End Sub
Public Sub ChatIgnoreByIndex(index As Long)
'By:Xcliptic
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, Count As Long
    Count& = RoomCount&
    If index& > Count& - 1 Then Exit Sub
    Room& = FindChatRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        a& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Loop Until a& <> 0&
    DoEvents
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Function ChatIgnoreByName(name As String) As Boolean
'By:Xcliptic
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindChatRoom&
    If Room& = 0& Then ChatIgnoreByName = False: Exit Function
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If GetUserSn <> ScreenName$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                ChatIgnoreByName = True
                Exit Function
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
ChatIgnoreByName = False
End Function
Function TimeUntilDate(date1 As Date, date2 As Date)
'By:Xcliptic
S = DateDiff("s", date1, date2)
M = S / 60
H = M / 60
d = H / 24
dd = Int(d)
H = (d - dd) * 24
hh = Int(H)
M = (H - hh) * 60
mm = Int(M)
S = (M - mm) * 60
ss = Int(S)
TimeUntilDate = dd & " days, " & hh & " hours, " & mm & " minutes, " & ss & " seconds to " & date2
End Function
Function LastMessageFromIM()
'last message from an im
'By:Xcliptic
IMWin& = FindChildByTitle(AOLMDI, ">Instant Message From")
If IMWin& = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message From")
If IMWin = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message To")
If IMWin& = 0& Then Exit Function
richtext& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
Text = GetText(richtext&)
Do
Find& = b&
b& = InStr(b& + 1, Text, ":")
Loop Until b& <= 0&
LastMessageFromIM = Mid(Text, Find& + 3, Len(Text))
End Function
Function SNfromIM()
'gets sn from first im
'By:Xcliptic
IMWin& = FindChildByTitle(AOLMDI, ">Instant Message From")
If IMWin& = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message From")
If IMWin = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message To")
If IMWin& = 0& Then Exit Function
imtitle$ = GetText(IMWin&)
If InStr(imtitle$, "To") <> 0 Then GoTo too
If InStr(imtitle$, ">") <> 0 Then GoTo n
SNfromIM = Mid(imtitle$, 25, Len(imtitle$))
Exit Function
too:
SNfromIM = Mid(imtitle$, 22, Len(imtitle$))
Exit Function
n:
SNfromIM = Mid(imtitle$, 24, Len(imtitle$))
End Function
Function UserOnline() As Boolean
'returns true if the user is online
'false if not
'By:Xcliptic
If FindChildByTitle(AOLMDI, "Welcome,") <> 0 Then UserOnline = True Else UserOnline = False
End Function
Sub IMIgnore(person As String)
'By:Xcliptic
SendIM "$IM_OFF, " & person, ".", False, False, False, False
End Sub
Sub IMUnIgnore(person As String)
'By:Xcliptic
SendIM "$IM_ON, " & person, ".", False, False, False, False
End Sub
Sub IMIgnoreList(Lst As ListBox)
'ignores an im if from some on the list
'By:Xcliptic
Do: DoEvents
x& = FindChildByTitle(AOLMDI, ">Instant Message From", x&)
l = GetText(x&)
b$ = Mid(l, 24, Len(l))
If ListSearch(Lst, b$) = -1 Then Exit Do
SendMessageLong x&, WM_CLOSE, 0, 0
Loop Until x& = 0
End Sub
Function FindNewIM(SN As String)
'finds new im
'By:Xcliptic
thewin = FindWindowEx(AOLMDI, 0&, "AOL Child", ">Instant Message from: " & SN)
FindIM = thewin
End Function
Function TheTime()
'By:Xcliptic
TheTime = Format(Now, "h:mm:ss AM/PM")
End Function
Sub CloseAllIMs()
'closes all open ims
'By:Xcliptic
Do
DoEvents
v& = FindChildByTitle(AOLMDI, "Instant Message", v&)
SendMessageLong v&, WM_CLOSE, 0, 0
Loop Until v& <> 0&
End Sub
Function GetChatName() As String
'name of chatroom
GetChatName = GetCaption(FindChatRoom)
End Function
Sub Pause(interval)
'pauses
'By:Xcliptic
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function CoolChat(Text As String)
'takes a random font and sends it to
'chat
'By:Xcliptic
b = Mid(Time, Len(Format(Time, "h:mm:ss")) - 2, 2)
For M = 1 To Len(Text)
If b > Printer.FontCount Then b = 1 Else b = b + 1
ob$ = Printer.Fonts(b)
aSt$ = aSt$ & "<font face=""" & ob$ & """>" & Mid(Text, M, 1)
Next
CoolChat = aSt$
End Function
Function ListSearch2(Lst As ListBox, Txt As String) As Integer
'search for a list
'By:Xcliptic
Dim x%
For x% = 0 To Lst.ListCount - 1
If InStr(UCase(Lst.List(x%)), UCase((Txt))) Then
ListSearch2 = x%
Exit Function
End If
Next x%
ListSearch2 = -1
End Function

Sub ClearChat()
'Only the user can see this clearchat(that sucks but try it its
'fun)
'By:Xcliptic
H = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
SendMessageByString H, WM_SETTEXT, 0&, ""
End Sub
Function RandomColor() As String
'give random color
'By:Xcliptic
RandomColor = ""
For M = 1 To 3
Randomize
rand = Int((255 * Rnd) + 1)
RandomColor = RandomColor & Hex(rand)
Next
End Function

Public Function RoomCount() As Long
'how many people in room
'By:Xcliptic
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindChatRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function
Sub CircleGradient(frm As Form, rs%, gs%, Bs%, re%, ge%, be%, smooth As Boolean)
'By:Xcliptic

If frm.WindowState = vbMinimized Then Exit Sub
frm.BackColor = RGB(re, ge, be)
If smooth = True Then
frm.DrawStyle = 6
Else
frm.DrawStyle = 0
End If
If frm.ScaleWidth <> 255 Then
frm.ScaleWidth = 255
End If
If frm.ScaleHeight <> 255 Then
frm.ScaleHeight = 255
End If
frm.DrawWidth = 5
frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (Bs - be) / 255
rc = rs: bc = Bs: gc = gs
For x = 0 To 255
frm.Circle (frm.ScaleWidth / 2, frm.ScaleHeight / 2), x, RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
End Sub
Sub RectGradient(frm As Form, rs%, gs%, Bs%, re%, ge%, be%, smooth As Boolean)
'By:Xcliptic

If frm.WindowState = vbMinimized Then Exit Sub
frm.BackColor = RGB(re, ge, be)
If smooth = True Then
frm.DrawStyle = 6
Else
frm.DrawStyle = 0
End If
If frm.ScaleWidth <> 255 Then
frm.ScaleWidth = 255
End If
If frm.ScaleHeight <> 255 Then
frm.ScaleHeight = 255
End If
frm.DrawWidth = 5
frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (Bs - be) / 255
rc = rs: bc = Bs: gc = gs
For x = 255 To 0 Step -1
DoEvents
frm.Line ((x / 2), (x / 2))-(frm.ScaleWidth - (x / 2), frm.ScaleHeight - (x / 2)), RGB(rc, gc, bc), B
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
End Sub

Sub DiagGradient(frm As Form, rs%, gs%, Bs%, re%, ge%, be%, smooth As Boolean)
'By:Xcliptic
If frm.WindowState = vbMinimized Then Exit Sub
frm.BackColor = RGB(re, ge, be)
If smooth = True Then
frm.DrawStyle = 6
Else
frm.DrawStyle = 0
End If
If frm.ScaleWidth <> 255 Then
frm.ScaleWidth = 255
End If
If frm.ScaleHeight <> 255 Then
frm.ScaleHeight = 255
End If
frm.DrawWidth = 5
frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (Bs - be) / 255
rc = rs: bc = Bs: gc = gs
For x = 0 To 255
DoEvents
frm.Line (0, x)-(x, 0), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
ri = (re - rs) / 255
gi = (ge - gs) / 255
bi = (be - Bs) / 255
rc = re: bc = be: gc = ge
For x = 255 To 0 Step -1
DoEvents
frm.Line (255 - x, 255)-(255, 255 - x), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
End Sub

Sub SpinGradient(frm As Form, rs%, gs%, Bs%, re%, ge%, be%, smooth As Boolean)
'By:Xcliptic
If frm.WindowState = vbMinimized Then Exit Sub
frm.BackColor = RGB(rs, gs, Bs)
If smooth = True Then
frm.DrawStyle = 6
Else
frm.DrawStyle = 0
End If
If frm.ScaleWidth <> 255 Then
frm.ScaleWidth = 255
End If
If frm.ScaleHeight <> 255 Then
frm.ScaleHeight = 255
End If
frm.DrawWidth = 5
frm.Refresh
ri = (rs - re) / 255 / 2
gi = (gs - ge) / 255 / 2
bi = (Bs - be) / 255 / 2
rc = rs: bc = Bs: gc = gs
For x = 0 To 255
DoEvents
frm.Line (x, 0)-(255 - x, 255), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
For x = 0 To 255
DoEvents
frm.Line (255, x)-(0, 255 - x), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next x
End Sub
Function FileExists(sFileName As String) As Integer
'finds if a file exists
'By:Xcliptic
On Error Resume Next
Dim FileNum As Integer
FileNum = FreeFile
Open sFileName For Input As FileNum
If Err Then
    FileExists = False
Else
    FileExists = True
End If
End Function


Sub PlayWav(wav As String)
'By:Xcliptic
w% = sndplaysound(wav, 0)
End Sub
Sub PlayMIDI(midilocation As String)
'midilocation ex:
'c:\windows\midi.mid
'By:Xcliptic
Dim RET As Long
RET = mciSendString("play " & midilocation, 0&, 0, 0)
End Sub
Sub StopMidi(midlocation As String)
Dim RET As Long
RET = mciSendString("stop " & midilocation, 0&, 0, 0)
End Sub
Function LoadtxtToText(TxtFilename As String, TextboxToload As textbox)
'this is a tricky sub, but you can handle it
'By:Xcliptic
Dim nFile As Integer
Dim sFile As String

'Get a suitable file number
nFile = FreeFile

'Specify the file you want opened.
sFile = TxtFilename

'Open the file
Open sFile For Input As nFile

'Read the file into the text box
TextboxToload = Input(LOF(nFile), nFile)

Close nFile


End Function

Function Savetext(TxtFilename As String, TxtTextbox As textbox)
'By:Xcliptic
Dim sFile As String
Dim nFile As Integer

nFile = FreeFile
sFile = TxtFilename
Open sFile For Output As nFile
Print #nFile, TxtTextbox
Close nFile


End Function

Function GetTextBoxLine(Text As textbox, theline As Integer) As String
'gets a specific line from textbox
'By:Xcliptic
U = CountTextBoxLines(Text)
If theline > U Then Exit Function
For g = 1 To theline - 1
e = InStr(e + 1, Text.Text, Chr(13))
Next
e = e + 1
If e = 1 Then e = 0
For F = 1 To theline
j = InStr(j + 1, Text.Text, Chr(13))
Next
If j = 0 Then j = Len(Text.Text) + 1
GetTextBoxLine = Mid(Text.Text, e + 1, j - 1 - e)
End Function
Function CountTextBoxLines(Text As textbox) As Long
'counts lines in a textbox
'By:Xcliptic
i = 0
Do
a = InStr(a + 1, Text.Text, Chr(13))
If a = 0 Then Exit Do
i = i + 1
Loop
CountTextBoxLines = i + 1
End Function

Public Function FindInfoWindow() As Long
'By:Xcliptic
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function

Function TypeOut(Text As String)
'makes a string type out one letter
'at a time in chat, mail or im.
'By:Xcliptic
lagtext$ = "</html>"
If Text = "" Then Exit Function
For j = 1 To Len(Text)
lagtext$ = lagtext$ & Mid(Text, j, 1) & "<html></html>"
Next
TypeOut = lagtext$
End Function
Sub OnTop(frm As Form)
'put your form on top of all others
'By:Xcliptic
Dim RET As Long
On Error Resume Next
RET& = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

' Fade Subs

Function HexRGB(thecolor)
'By:Xcliptic

       Dim acol As Long
       Dim ared, agreen, ablue As Integer
       Dim rhex, ghex, bhex As Variant
       acol = Val(thecolor)
       ared = acol Mod &H100
       acol = acol \ &H100
       agreen = acol Mod &H100
       acol = acol \ &H100
       ablue = acol Mod &H100
    
       rhex = Hex(ared)

              If Len(rhex) < 2 Then
                      rhex = "0" & rhex
              End If

       ghex = Hex(agreen)

              If Len(ghex) < 2 Then
                      ghex = "0" & agreen
              End If

       bhex = Hex(ablue)

              If Len(bhex) < 2 Then
                      bhex = "0" & bhex
              End If

       HexRGB = "#" & rhex & ghex & bhex
End Function
Function GetRGB(colr As Long) As TheRgb
'By:Xcliptic
       acol = colr
       GetRGB.Red = acol Mod &H100
       acol = acol \ &H100
       GetRGB.Green = acol Mod &H100
       acol = acol \ &H100
       GetRGB.Blue = acol Mod &H100
End Function
Function Fade2Color(col1 As Long, col2 As Long, Text As String, Wavy As Boolean)
'two color fade
'By:Xcliptic

If col1 = FFFFFF Then col1 = FFFFFE
If col2 = FFFFFF Then col2 = FFFFFE
R1 = GetRGB(col1).Red
G1 = GetRGB(col1).Green
B1 = GetRGB(col1).Blue
R2 = GetRGB(col2).Red
G2 = GetRGB(col2).Green
B2 = GetRGB(col2).Blue

rdiff = Int((R1 - R2) / Len(TrimSpaces(Text)))
gdiff = Int((G1 - G2) / Len(TrimSpaces(Text)))
bdiff = Int((B1 - B2) / Len(TrimSpaces(Text)))

For H = 1 To Len(Text)
If Mid(Text, H, 1) = " " Then thes$ = thes$ & " ": H = H + 1 Else
If Wavy = True Then
wave = wave + 1
If wave > 4 Then wave = 1
If wave = 1 Then html = "<sub>"
If wave = 2 Then html = "</sub>"
If wave = 3 Then html = "<sup>"
If wave = 4 Then html = "</sup>"
Else
html = vbNullString
End If
If R1 > 255 Then R1 = R1 - (R1 - 255)
If R1 < 0 Then R1 = R1 - R1
If G1 > 255 Then G1 = G1 - (G1 - 255)
If G1 < 0 Then G1 = G1 - G1
If B1 > 255 Then B1 = B1 - (B1 - 255)
If B1 < 0 Then B1 = B1 - B1

newrgb& = RGB(Int(R1), Int(G1), Int(B1))
thes$ = thes$ & "<font color=" & HexRGB(newrgb&) & ">" & Mid(Text, H, 1) & html
R1 = R1 - rdiff
G1 = G1 - gdiff
B1 = B1 - bdiff
Next
Fade2Color = thes$
End Function
Function Fade3Color(col1 As Long, col2 As Long, col3 As Long, Text As String, Wavy As Boolean)
'By:Xcliptic
If Len(Text) <= 7 Then Fade3Color = Fade2Color(col1, col2, Text, Wavy): Exit Function
lens = Int(Len(Text) / 2)
Fade3Color = Fade2Color(col1, col2, Mid(Text, 1, lens), Wavy)
Fade3Color = Fade3Color & Fade2Color(col2, col3, Mid(Text, lens + 1, Len(Text) - lens), Wavy)
End Function
Function Fade4Color(col1 As Long, col2 As Long, col3 As Long, col4 As Long, Text As String, Wavy As Boolean)
'By:Xcliptic
If Len(Text) <= 7 Then Fade4Color = Fade2Color(col1, col2, Text, Wavy): Exit Function
lens = Int(Len(Text) / 3)
Fade4Color = Fade2Color(col1, col2, Mid(Text, 1, lens), Wavy)
Fade4Color = Fade4Color & Fade2Color(col2, col3, Mid(Text, lens + 1, lens), Wavy)
Fade4Color = Fade4Color & Fade2Color(col3, col4, Mid(Text, lens + lens + 1, Len(Text) - lens), Wavy)
End Function
Function Fade5color(col1 As Long, col2 As Long, col3 As Long, col4 As Long, col5 As Long, Text As String, Wavy As Boolean)
'By:Xcliptic
If Len(Text) <= 7 Then Fade5color = Fade2Color(col1, col2, Text, Wavy): Exit Function
lens = Int(Len(Text) / 4)
Fade5color = Fade2Color(col1, col2, Mid(Text, 1, lens), Wavy)
Fade5color = Fade5color & Fade2Color(col2, col3, Mid(Text, lens + 1, lens), Wavy)
Fade5color = Fade5color & Fade2Color(col3, col4, Mid(Text, lens + lens + 1, lens), Wavy)
Fade5color = Fade5color & Fade2Color(col4, col5, Mid(Text, lens + lens + lens + 1, Len(Text) - lens), Wavy)
End Function

Function BGFade(col1 As Long, col2 As Long, Speed As Integer)
'Fades the background of an im or mail.
'Speed = 1 - 25
'1 is the slowest but looks the best
'25 is the fastest but looks the worst
'If you want to use it on an IM then
'You have to have a speed of at least
'12 anything lower may give you an
'error saying that the message is too
'complex!
'By:Xcliptic
Dim x As Integer
Dim RedChange As Integer
Dim GreenChange As Integer
Dim BlueChange As Integer
If Speed > 25 Then Speed = 25
StartRed = GetRGB(col1).Red
StartGreen = GetRGB(col1).Green
StartBlue = GetRGB(col1).Blue
EndRed = GetRGB(col2).Red
EndGreen = GetRGB(col2).Green
EndBlue = GetRGB(col2).Blue
For x = 0 To 255 Step Speed
col$ = RGB(StartRed + RedChange, StartGreen + GreenChange, StartBlue + BlueChange) 'Draws Line With correct color
BGFade = BGFade & "<BODY BGCOLOR=" & HexRGB(Val(col$)) & ">"
For F = 1 To Speed
RedChange = RedChange + (EndRed - StartRed) / 255 '
GreenChange = GreenChange + (EndGreen - StartGreen) / 255
BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
Next F
Next x
End Function


Function BGFadeWithMessage(col1 As Long, col2 As Long, newstr As String)
'By:Xcliptic
Dim x As Integer
Dim RedChange As Integer
Dim GreenChange As Integer
Dim BlueChange As Integer
'same as bgfade just puts a message
'that types out one char at a time
'ex.
'mail ScreenName, subject, BGFadeWithMessage(vbRed, vbBlue, "Hey, whats up?")
StartRed = GetRGB(col1).Red
StartGreen = GetRGB(col1).Green
StartBlue = GetRGB(col1).Blue
EndRed = GetRGB(col2).Red
EndGreen = GetRGB(col2).Green
EndBlue = GetRGB(col2).Blue
U = 1
If letters > 256 Then Exit Function
letters = Int(256 / Len(newstr))
For x = 0 To 255
col$ = RGB(StartRed + RedChange, StartGreen + GreenChange, StartBlue + BlueChange) 'Draws Line With correct color
BGFadeWithMessage = BGFadeWithMessage & "<BODY BGCOLOR=" & HexRGB(Val(col$)) & ">"
If x Mod letters = 0 Then BGFadeWithMessage = BGFadeWithMessage & Mid(newstr, U, 1): U = U + 1
RedChange = RedChange + (EndRed - StartRed) / 255 '
GreenChange = GreenChange + (EndGreen - StartGreen) / 255
BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
Next x
End Function
Function RoomOrOK() As String
'By:Xcliptic
Dim Room&
Dim okwin&
Do
DoEvents
Room& = FindChatRoom
okwin& = FindWindow("#32770", "America Online")
Loop Until okwin& <> 0 Or Room& <> 0
If Room& <> 0 Then
RoomOrOK = "Room"
ElseIf okwin& <> 0 Then
RoomOrOK = "OK"
End If
End Function
Sub HideWelcome()
'By:Xcliptic
Dim WelcomeWin&
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
WindowHide (WelcomeWin&)
End Sub
Sub ShowWelcome()
'By:Xcliptic
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
WindowShow (WelcomeWin&)
End Sub
Sub RemoveFromListbyString(List As ListBox, remove As String)
'By:Xcliptic
g = ListSearch(List, remove)
If g = -1 Then Exit Sub
List.RemoveItem g
End Sub
Sub HideAOL()
'By:Xcliptic
WindowHide (AOLWindow)
End Sub
Sub ShowAOL()
'By:Xcliptic
WindowShow (AOLWindow)
End Sub
Function Uppercase(Text As String) As String
'By:Xcliptic
Uppercase = UCase(Text)
End Function
Function Lowercase(Text As String) As String
'By:Xcliptic
Lowercase = LCase(Text)
End Function
Function HackerText(Text As String) As String
'By:Xcliptic
For H = 1 To Len(Text)
If H Mod 2 <> 0 Then Mid(Text, H, 1) = UCase(Mid(Text, H, 1)) Else Mid(Text, H, 1) = LCase(Mid(Text, H, 1))
Next
HackerText = Text
End Function
Function EncryptText(Text As String) As String
'By:Xcliptic
For H = 1 To Len(Text)
theasc = Asc(Mid(Text, H, 1)) + 4
If theasc > 255 Then theasc = (theasc - 255)
d$ = d$ & Chr(theasc)
Next
EncryptText = d$
End Function
Function DecryptText(Text As String) As String
'By:Xcliptic
For H = 1 To Len(Text)
theasc = Asc(Mid(Text, H, 1)) - 4
If theasc <= 0 Then theasc = (255 + theasc)
d$ = d$ & Chr(theasc)
Next
DecryptText = d$
End Function
Sub SetClipboardText(Text As String)
'By:Xcliptic
Clipboard.Clear
Clipboard.SetText Text
End Sub
Function GetClipboardText() As String
'By:Xcliptic
GetClipboardText = Clipboard.GetText
End Function

Sub MadBot()
'need:
'2 textboxes, 2 command buttons, and 1 timer
'in Command1 put:
'Timer1.Enabled = True
'in Command2 put:
'Do
'DoEvents
'Loop
'Timer1.Enabled = False

'Text1 is the name of the person
'Text2 is something that would tick some1 off
'without any of these it wont work so doent even try
'By:Xcliptic
SendChat ("Man, " & Text1.Text & "you " & Text2.Text)
TimeOut (0.2)
SendChat ("You heard me, I said you " & Text2.Text)
Do
SendChat (Text1.Text & "suckz")
TimeOut (5)
Loop Until Timer1.Enabled = False
End Sub

'Public Sub SendChat(Text As String)
'here is where u put yor code for sending text to a chat room
'there are many bas's out there and if u ask properly, u kan
'use theirs.  There are other things u kan use besides SendChat
'i just used SendChat cuz i just thought of it.  It kan be
'ChatSend,Chat, AOL4Chat, etc...

'this sub doesn't work but the bots do if u use a sub that
'can send text to a chat room and put it here
End Sub

Sub AFKBot()
'need:
'2 textboxes, 2 command buttons, and a 1 timer
'Put Timer's interval to 5 and in it put:
'all need to run this AFK bot
'By:Xcliptic
SendChat ("AFK Bot Activated")
SendChat ("AFK for:" & Text1.Text & " min(z)") 'text1 is how long u r afk
TimeOut (0.1)
SendChat ("Reason: " & Text2.Text) 'text2 is the reason
TimeOut (Text1.Text)
'in Command1 put:
'Timer1.Enabled = True
'in Command2 put:
'Timer1.Enabled = False
'Do
'DoEvents
'Loop
End Sub

Sub IdleBot()
'need:
'1 textboxes, 2 commandbuttons, and 2 labels
'make Label1's caption "Reason:" and put it above the textbox
'make Command1's caption "Start" and Command2's "Stop
'in Command1.Click put: Call IdleBot
'in Command2.Click put:
'Do
'DoEvents
'Loop
'make the text1's text = ""
'make Label2's caption "How Long:" and put it above text2
'make Text2.Text's = ""
'******START OF SUB CODE******'
'By:Xcliptic
SendChat ("Idle Bot Activated")
TimeOut (1)
SendChat ("Reason:" & Text1.Text)
TimeOut (0.1)
SendChat ("How Long:" & Text2.Text)
TimeOut (0.1)
SendChat ("I'll be back...")
TimeOut (Text2.Text)
End Sub

Sub RequestBot()
'need:
'2 Textboxes and 2 Command buttons
'Text1 is what you want
'Text2 is yor name
'By:Xcliptic

SendChat ("Request Bot Activated")
SendChat ("Request: " & Text1.Text)
TimeOut (0.1)
SendChat ("Better give " & Text2.Text & " what he wants")

'in Command1 put:
'Call RequestBot

'in Command2 put:
'SendChat ("Everything's cool now cuz" & Text2.Text & "got what he wanted")
'Sendchat ("and that is:" & Text1.Text)
'TimeOut (0.3)
'Sendchat ("Request Bot Deactivated")
End Sub

Sub AttentionBot()
'to use it put: Call AttentionBot
'By:Xcliptic
SendChat ("{S IM")
TimeOut (1.2)
SendChat ("Attention Bot")
SendChat ("Gimme Attention!")
TimeOut (0.1)
SendChat ("{S IM")
TimeOut (1.1)
SendChat ("{S BUDDYIN")
TimeOut (1.1)
SendChat ("{S GOTMAIL")
TimeOut (1.1)
End Sub

Sub AdvertizeBot()
'to use it put: Call AdvertizeBot
'By:Xcliptic
SendChat ("[write the name of your prog really fancy here]")
TimeOut (0.1) ' the timesouts are so u dont get logged off for scrolling
SendChat ("Made By [your handle here]")
TimeOut (0.1)
SendChat ("[Some cocky or attention-getter for yor prog]")
End Sub

Sub EchoBot()
'By:Xcliptic
'need:
'1 timer and 2 Command buttons
'With Timer1's interval at 5 and enabled = false put:
Dim LastChatLine As String
Dim SNLastChatLine As String
LastChatLine$ = LastChatLine 'this would be a sub that would get
                          'the last text line in a chat room
SNLastChatLine$ = SNFromLastChatLine 'this would also be a sub that would
                            'get the SN from the last text line
                            'in a chat room
SendChat ("Echo Bot Active")
SendChat ("Echoing: " & SNFromLastChatLine$)
TimeOut (0.3)
Do
SendChat (LastChatLine$)
TimeOut (4)
Loop Until Timer1.Enabled = False
'in Command1 put: Timer1.Enabled = True
'in Command2 put: Timer1.Enabled = False
End Sub

Sub FightBot()
'By:Xcliptic
'need:
'2 textboxes and 1 label
'work:
'make Label1 NOT visible, so they kant see it.
'the first textbox is 1 person the other textbox is the other
Label1.Caption = Int(Rnd * 3)
If Label1.Caption = 1 Then
SendChat (Text1.Text & "punches" & Text2.Text)
TimeOut (1)
SendChat (Text1.Text & "kicks" & Text2.Text)
TimeOut (1)
SendChat (Text1.Text & "kills" & Text2.Text)
TimeOut (1)
SendChat ("The Winner is: " & Text1.Text)
Else
SendChat (Text2.Text & "punches" & Text1.Text)
TimeOut (1)
SendChat (Text2.Text & "kicks" & Text1.Text)
TimeOut (1)
SendChat (Text2.Text & "kills" & Text1.Text)
TimeOut (1)
SendChat ("The Winner is: " & Text2.Text)
End If
End Sub


Public Sub TimeOut(Duration)
'By:Xcliptic
Dim Starttime As Long
  Starttime = Timer
Do While Timer - Starttime > Duration
DoEvents
Loop
End Sub

Sub FakeBot()
'By:Xcliptic
'need:
'_ Command buttons, it depends on how many progs u wanna do it to
'dont call this sub, just use the code given and put in each
'command button this:
'Call SendChat("[name of prog which i will provide a lot of em]")

'these are the opening things that some progs say when they
'are activated.

'********************START OF CODE***************************
'º¯`v´¯¯) PhrostByte By: Progee (¯¯`v´¯º

'.·´¯`·-  gøthíc nightmâres by másta  ­·´¯`·
'·._.--   aøl 4.o punt tools · loaded ---._.·

'· úpr mácro stùdio · másta ·

'-=·Sting Anti Punta 2.o Loaded·=-
'-=·MaDe By SaBrE·=-

'(¯\_ GøDZîLLa³·º _/¯)
'(¯\_ ßy ÇoLd _/¯)

'¢º°¤÷®ÍP§ 2øøø÷¤°º¢
' ¢º°¤÷£øÃdÊD÷¤°º¢

'-•(`(`·•Fate Zero v¹ Loaded•·´)´)•-

'•·.·´).·÷•[ Outlaw Mass Mailer by Twiztid

'(¯`·.····÷• ärméñïå¹ · kðkô
'(¯`·.····÷• îøâdèd

'•·._.·´¯`·>AoL 4.0 TooLz By: X GeNuS X
'•·._.·´¯`·>Status: LoaDeD
'•·._.·´¯`·>Ya'll BeTTa NoT MeSS WiT ThiS NiG!

'^····÷• James Bond Toolz Ver .007
'^····÷• By: Saßan

'(¯`•Prophecy²·° Loaded

'···÷••(¯`·._ CoRn Fader _.·´¯)••÷···
'···÷••(¯`·._Created by :::PooP:::_.·´¯)••÷···

'Blue Ice Punter¹ For AOL 4.0
'By STaNK

'¤-----==America Onfire Platinum
'¤-----==Loaded
'¤-----==Created (²›y Fatal Error

'.­”ˆ”­.•Fí/\/ä£ Få/\/†ä§y \/ïïï•.­”ˆ”­.
'·­„¸„­•·ßy RšZz•
'.­”ˆ”­.•£õàÐëÐ

'¤¤†³¹¹º†¤¤ SANNMEN †oºLz ¤¤†³¹¹º†¤¤
'¤¤†³¹¹º†¤¤    By:Má§†é®MinÐ    ¤¤†³¹¹º†¤¤
'¤¤†³¹¹º†¤¤ LOADED ¤¤†³¹¹º†¤¤

'··¤÷×(Rapier Bronze)×÷¤··
'··¤÷×(By Excalibur)×÷¤··
'··¤÷×(Works for 3.0 and 4.0!!!)×÷¤··

'<-==(`(` Icy Hot 2.0 For AOL 4.0 ')')==->
'<-==(`(` Loaded ')')==->

'(\›•‹ Im Backfire KiLLer ›•‹/)
'(\›•‹ By:phire Status:Loaded ›•‹/)
'(\›•‹ Im Backfire KiLLer ›•‹/)

'[_.·´¯° Indian Invasion Punter Loaded °¯`·._]


'··¤·´¯° X ChAt CoMmAnDoR °¯`¤··
'··¤·´¯° By: Xcliptic °¯`¤··
'··¤·´¯° Get your compy today!°¯`¤··
'***********************END OF CODE**************************
End Sub

Sub HiBot()
'By: Xcliptic
SendChat ("Hi Bot Loaded.  Hi Everybody!")
Do
If LastChatLine = "Hi" Then
SendChat ("Hi " & SNFromLastChatLine & "!")
Loop Until LastChatLine = "Hi"
End If
End Sub


Sub CustomBot()
'By: Xcliptic
'make sure u make a stop button and in it put:
'Do
'DoEvents
'Loop

'u need 2 textboxes
'Text1 is what they say and Text2 is what u want to say
Do
If Text1.Text = LastChatLine Then
SendChat (Text2.Text)
Loop Until Text1.Text = LastChatLine
End If
End Sub

Sub QuizBot()
'By: Xcliptic
'need:
'1 label
Label1.Caption = Int(Rnd * 11)
If Label1.Caption = "1" Then
SendChat ("What state is Harrisburg in?")
Do
If LastChatLine = "Pennsylvania" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "Pennsylvania"
End If
    If Label1.Caption = "2" Then
    SendChat ("How many inches r in a foot?")
    Do
    If LastChatLine = "12" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "12"
    End If
If Label1.Caption = "3" Then
SendChat ("How many hours r in a day?")
Do
If LastChatLine = "24" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "24"
End If
    If Label1.Caption = "4" Then
    SendChat ("How many days r in a week?")
    Do
    If LastChatLine = "7" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "7"
    End If
If Label1.Caption = "5" Then
SendChat ("What country has the most population?")
Do
If LastChatLine = "China" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "China"
    If Label1.Caption = "6" Then
    SendChat ("Does money suck?")
    Do
    If LastChatLine = "no" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "no"
    End If
If Label1.Caption = "7" Then
SendChat ("Which is bigger, USA or Japan?")
Do
If LastChatLine = "USA" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "USA"
End If
    If Label1.Caption = "8" Then
    SendChat ("How many points is a touchdown?")
    Do
    If LastChatLine = "6" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "6"
    End If
If Label1.Caption = "9" Then
SendChat ("Which is bigger, USA or Japan?")
Do
If LastChatLine = "USA" Then SendChat (SNFromLastChatLine & ", you're right!")
Loop Until LastChatLine = "USA"
End If
    If Label1.Caption = "10" Then
    SendChat ("What kind of verb is ran?")
    Do
    If LastChatLine = "action" Then SendChat (SNFromLastChatLine & ", you're right!")
    Loop Until LastChatLine = "action"
    End If
End Sub

Sub ShhBot()
'By: Xcliptic
'need:
'_ Textboxes, it depends on how many people
Do
If SNFromLastChatLine = Text1.Text Then 'u kan add more textboxes for more people
SendChat ("STFU " & SNFromLastChatLine & "!")
Loop Until SNFromLastChatLine = Text1.Text
End If
End Sub

Sub EightBallBot()
'By: Xcliptic
'need:
'1 textbox, 1 label
'Text1 is what u r asking, and the label where u randomize
Label1.Caption = Int(Rnd * 9)
SendChat ("8Ball Bot Loaded")
TimeOut (0.4)
SendChat (Text1.Text & " = Question")
    If Label1.Caption = "1" Then
    SendChat ("Excellent Chance!")
    End If
If Label1.Caption = "2" Then
SendChat ("Great Chance!")
End If
    If Label1.Caption = "3" Then
    SendChat ("Good Chance!")
    End If
If Label1.Caption = "4" Then
SendChat ("OK Chance!")
End If
    If Label1.Caption = "5" Then
    SendChat ("Bad Chance!")
    End If
If Label1.Caption = "6" Then
SendChat ("Very Bad Chance!")
End If
    If Label1.Caption = "7" Then
    SendChat ("0% Chance!")
    End If
If Label1.Caption = "8" Then
SendChat ("Ur having a horrible day!")
End If
End Sub

Sub ScrambleBot()
'By: Xcliptic
'need:
'1 label
Dim aString As String, eString As String, iString As String
SendChat ("ScrambleBot Loaded")
TimeOut (0.2)
SendChat ("Try to unscramble the words:")
Label1.Caption = Int(Rnd * 6)

    If Label1.Caption = "1" Then
    aString$ = "eggs"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 2)
    TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "eggs" Then SendChat ("Correct, the word is eggs")
    Loop Until LastChatLine = "eggs"
    End If
 TimeOut (0.2)
If Label1.Caption = "2" Then
aString$ = "poop"
eString$ = Left(aString$, 3)
iString$ = Right(aString$, 1)
TimeOut (0.3)
SendChat (iString$ & eString$)
Do
If LastChatLine = "poop" Then SendChat ("Correct, the word is poop")
Loop Until LastChatLine = "poop"
End If
 TimeOut (0.2)
    If Label1.Caption = "3" Then
    aString$ = "bacon"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 3)
 TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "bacon" Then SendChat ("Correct, the word is bacon")
    Loop Until LastChatLine = "bacon"
    End If
 TimeOut (0.2)
If Label1.Caption = "4" Then
aString$ = "Sting"
eString$ = Left(aString$, 3)
iString$ = Right(aString$, 2)
TimeOut (0.3)
SendChat (iString$ & eString$)
Do
If LastChatLine = "Sting" Then SendChat ("Correct, the word is Sting")
Loop Until LastChatLine = "Sting"
End If
 TimeOut (0.2)
    If Label1.Caption = "5" Then
    aString$ = "SaBrE"
    eString$ = Left(aString$, 2)
    iString$ = Right(aString$, 3)
 TimeOut (0.3)
    SendChat (iString$ & eString$)
    Do
    If LastChatLine = "SaBrE" Then SendChat ("Correct, the word is SaBrE")
    Loop Until LastChatLine = "SaBrE"
End If
 TimeOut (0.2)
End Sub

Sub LuckyNumberBot()
'By: Xcliptic
'need:
'1 textbox and 1 label
Label1.Caption = Int(Rnd * 1000)
SendChat ("Lucky # Bot Loaded")
TimeOut (1)
SendChat ("type /luckynumber to see your lucky #")
TimeOut (1)
Do
If LastChatLine = "/luckynumber" Then SendChat (SNFromLastChatLine & " : " & Label1.Caption)
Loop Until LastChatLine = "/luckynumber"
End Sub

Sub GuessNumberBot()
'By: Xcliptic
'need:
'1 label
Label1.Caption = Int(Rnd * 6)
    If Label1.Caption = "1" Then
    SendChat ("I'm thinking of a number between 1-3")
    Do
    If LastChatLine = "2" Then SendChat ("2 is right!")
    Loop Until LastChatLine = "2"
    End If
If Label1.Caption = "2" Then
SendChat ("I'm thinking of a number between 10-13")
Do
If LastChatLine = "12" Then SendChat ("12 is right!")
Loop Until LastChatLine = "12"
End If
    If Label1.Caption = "3" Then
    SendChat ("I'm thinking of a number between 30-33")
    Do
    If LastChatLine = "31" Then SendChat ("31 is right!")
    Loop Until LastChatLine = "31"
    End If
If Label1.Caption = "4" Then
SendChat ("I'm thinking of a number between 40-43")
Do
If LastChatLine = "42" Then SendChat ("42 is right!")
Loop Until LastChatLine = "42"
End If
    If Label1.Caption = "5" Then
    SendChat ("I'm thinking of a number between 50-53")
    Do
    If LastChatLine = "51" Then SendChat ("51 is right!")
    Loop Until LastChatLine = "51"
    End If

End Sub

'This bas was made by:Xcliptic'                                                                                                                                                                                                                                                                                                                                                                                                                                                                         'All code written and made by SaBrE, Copyright 1999 : SaBrE.  If this is found on yor bas and u dont have SaBrE's permission, ur in deep crap

'                                               ______
'      ______________________________________  /
'     /--------------------------------------\/|||||\)\
'    /________________________________Xcliptic__/\|||||

Sub AddFilesToList(list1 As ListBox)

       Dim R As Long
       Dim pathSpec As String
       '     'fill the listbox yeee haah lol -doom
       pathSpec = "c:\windows\system\*.*"
       R = SendMessageStr(list1.hWnd, LB_DIR, DDL_FLAGS, pathSpec)
End Sub
Sub File_Copy(File$, DestFile$)
'By: Xcliptic
If Not File_IfileExists(File$) Then Exit Sub
FileCopy File$, DestFile$
End Sub
Sub File_Delete(File$)
'By: Xcliptic
Dim NoFreeze%
If Not File_IfileExists(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub
Function File_IfileExists(ByVal sFileName As String) As Integer
'By: Xcliptic
'Example: If Not File_ifileexists("win.com") then...
Dim i As Integer
On Error Resume Next
i = Len(dir$(sFileName))
    If Err Or i = 0 Then
        File_IfileExists = False
        Else
            File_IfileExists = True
    End If

End Function
Function Scan_Deltree(File As String)
'By: Xcliptic
'example : Call Scan_Deltree(text1.text)

Dim FileLenn As Variant
Dim FileLennn As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As String
Dim l0044 As Single
Dim l0046 As Single
Dim l0048 As Single
Dim l004a As Single
Dim l004c As Single
Dim l004e As Single
Dim l0050 As Single
Dim l0052 As Single
Dim l0054 As Single
Dim l0056 As Single
Dim l0058 As Single
Dim l005A As Variant
Dim l0045!
Open File For Binary As #2
DoEvents
FileLenn = LOF(2)
FileLennn = FileLenn
l003A = 1
While FileLennn >= 0
    If FileLennn > 32000 Then
        l003E = 32000
    ElseIf FileLennn = 0 Then
        l003E = 1
    Else
        l003E = FileLennn
    End If
    l0042$ = String$(l003E, " ")
    Get #2, l003A, l0042$
    l0044! = InStr(1, l0042$, "deltree \y", 1)
    l0045! = InStr(1, l0042$, "MZÿ C:\*.*", 1)
If l0044! Then DeltreeScan = True
Close: Exit Function

If Not l0044! Then DeltreeScan = False
Close: Exit Function
Wend
End Function


Function IfDirExists(TheDirectory)
'By: Xcliptic

Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(dir$(TheDirectory))
If Err Or Check = 0 Then
    IfDirExists = False
Else
    IfDirExists = True
End If
End Function
Function FreeProcess()
'By: Xcliptic
Dim DooM

Do: DoEvents
DooM = DooM + 1
If DooM = 50 Then Exit Do
Loop
End Function
Sub Directory_Delete(dir)
'By: Xcliptic
'This deletes a directory automatically from your HardDrive
RmDir (dir)
End Sub
Sub Directory_Create(dir)
'By: Xcliptic
'Call Directory_Create("C:\NewDir")
MkDir dir
End Sub
Private Function Scan_For(sFile$, ByVal sWhat$)
'By: Xcliptic
Dim VariantA As Variant
Dim VariantB As Variant
Dim VariantC As Variant
Dim VariantD As Variant
Dim SingleA As Single
Dim StringA As String
Dim EnterKey As String

On Error Resume Next
Open sFile$ For Binary As #1
    EnterKey$ = Chr$(13) + Chr$(10)
    msg$ = ""
    VariantA = LOF(1)
    VariantB = VariantA
    VariantC = 1

    If VariantB > 32000 Then
        VariantD = 32000
    ElseIf VariantB = 0 Then
        VariantD = 1
    Else
        VariantD = VariantB
    End If

    StringA$ = String$(VariantD, " ")
    Get #1, VariantC, StringA$

    SingleA! = InStr(1, StringA$, sWhat$, 1)

    If SingleA! Then
        Scan_For = 0
    Else
        Scan_For = 1
    End If
Close #1
End Function
Function ScanFile1(Filename$, Searchstring As String, Label As Label) As Long
'By: Xcliptic
'ok, Filename is the File to scan
'SearchSring is the string to search for.
'and label like tells if it is a virus or not
'this is a example:
'call scanfile (text1.text,"Deltree y c:",label1)
'thats all
'i have also listed a list of virus searchstring below
'Main.idx, Deltree, Kill C:, Win.ini
'@Juno.com, @Hotmail.com, @FreeMail.com
'Deltree y c:, .Com, Deltree.com
'that is only a FEW

'-ÐøøM
'
'
Dim free
Dim x
Dim Text$


free = FreeFile
Dim Where As Long
Open Filename$ For Binary Access Read As #free
For x = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, x, Text$
    Debug.Print x
    If InStr(1, Text$, Searchstring$, 1) Then
    
MsgBox "Virus Found!"

        Where = InStr(1, Text$, Searchstring$, 1)
        ScanFile1 = (Where + x) - 1
        Close #free
        Exit For
    End If
    Next x
    
    
    
    If Not InStr(1, Text$, Searchstring$, 1) Then
MsgBox "No virus Found"

  End If
  
    
Close #free
End Function
Sub Directory_Delete2(dirnAmes$)
'By: Xcliptic
If Not IfDirExists(dirnAmes$) Then MsgBox dirnAmes$ & Chr(13) & "Bad Dir File Name!", 16, "Error": Exit Sub
On Error GoTo ErrorInDeletion
Kill dirnAmes$
Exit Sub
ErrorInDeletion:
MsgBox Error$
Resume Exitinga
Exitinga:
Exit Sub
End Sub
Private Sub Glassify()
'By: Xcliptic
GlassifyForm Me
End Sub
Private Sub Unglassify()
'By: Xcliptic
UnglassifyForm Me
End Sub
Private Sub Minimize()
'By: Xcliptic
WindowState = vbMinimized
End Sub

Private Sub Unload()
'By: Xcliptic
Unload Me
End Sub

Public Sub XFader_TextColorChange(textbox As textbox, R, g, b)
'Contributed by Xcliptic
    textbox.ForeColor = RGB(R, g, b)
    Pause 0.2
    textbox.ForeColor = RGB(R * 3 / 4, g * 3 / 4, b * 3 / 4)
    Pause 0.1
    textbox.ForeColor = RGB(R / 2, g / 2, b / 2)
    Pause 0.1
    textbox.ForeColor = RGB(R / 4, g / 4, b / 4)
    Pause 0.1
    textbox.ForeColor = RGB(0, 0, 0)
End Sub
Function XFader_Rich2HTML(RichTXT As Control, StartPos%, EndPos%)
Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTXT.SelStart = posi%
RichTXT.SelLength = 1

If Bolded <> RichTXT.SelBold Or posi% = StartPos Then
If RichTXT.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTXT.SelUnderline Or posi% = StartPos Then
If RichTXT.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTXT.SelStrikeThru Or posi% = StartPos Then
If RichTXT.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTXT.SelItalic Or posi% = StartPos Then
If RichTXT.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTXT.SelColor Or posi% = StartPos Then
ColorX = RGB(XFader_GetRGB(RichTXT.SelColor).Blue, XFader_GetRGB(RichTXT.SelColor).Green, XFader_GetRGB(RichTXT.SelColor).Red)
colorhex = XFader_RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTXT.SelColor
End If

If LastFont <> RichTXT.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTXT.SelFontName + Chr(34) + ">"
LastFont = RichTXT.SelFontName
End If

HTMLString = HTMLString + RichTXT.SelText
Next posi%

Rich2HTML = HTMLString

End Function


Public Function PWSD_BVScan2(Filename$, Searchstring$) As Long
'By: Xcliptic
free = FreeFile
Dim Where As Long
Open Filename$ For Binary Access Read As #free
For x = 1 To LOF(free) Step 32000
Text$ = Space(32000)
Get #free, x, Text$
 Debug.Print x
 If InStr(1, Text$, Searchstring$, 1) Then
 Where = InStr(1, Text$, Searchstring$, 1)
FileSearch = (Where + x) - 1
 Close #free
  Exit For
 End If
  Next x
Close #free
End Function
Public Function PWSD_BVScan(Filename$, ByVal Searchstring$)
'By: Xcliptic
Dim Variant1 As Variant
Dim Variant2 As Variant
Dim Variant3 As Variant
Dim Variant4 As Variant
Dim Single1 As Single
Dim String1 As String
Dim EnterKey As String
On Error Resume Next
Open Filename$ For Binary As #1
EnterKey$ = Chr$(13) + Chr$(10)
msg$ = ""
Variant1 = LOF(1)
Variant2 = Variant1
Variant3 = 1
If Variant2 > 32000 Then
Variant4 = 32000
ElseIf Variant2 = 0 Then
Variant4 = 1
Else
Variant4 = Variant2
End If
StringA$ = String$(Variant4, " ")
Get #1, Variant3, String1$
Single1! = InStr(1, String1$, Searchstring$, 1)
If Single1! Then
PWSD_BVScan = 0
Else
PWSD_BVScan = 1
End If
Close #1
End Function

Function XFader_CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'By: Xcliptic
'This gets a color from 3 scroll bars
XFader_CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = XFader_CLRBars(RedScroll1, GreenScroll1, BlueScroll1)
End Function
Function XFader_FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, Wavy As Boolean)
'By: Xcliptic
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
dacolor5$ = XFader_RGBtoHEX(Colr5)
dacolor6$ = XFader_RGBtoHEX(Colr6)
dacolor7$ = XFader_RGBtoHEX(Colr7)
dacolor8$ = XFader_RGBtoHEX(Colr8)
dacolor9$ = XFader_RGBtoHEX(Colr9)
dacolor10$ = XFader_RGBtoHEX(Colr10)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))
XFader_FadeByColor10 = XFader_FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, Wavy)
End Function

Function XFader_FadeByColor2(Colr1, Colr2, thetext$, Wavy As Boolean)
'By: Xcliptic
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
XFader_FadeByColor2 = XFader_FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, Wavy)
End Function

Function XFader_FadeByColor3(Colr1, Colr2, Colr3, thetext$, Wavy As Boolean)
'By: Xcliptic
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
XFader_FadeByColor3 = XFader_FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, Wavy)
End Function
Function XFader_FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, Wavy As Boolean)
'By: Xcliptic
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
XFader_FadeByColor4 = XFader_FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, Wavy)
End Function
Function XFader_FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, Wavy As Boolean)
'By: Xcliptic
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
dacolor5$ = XFader_RGBtoHEX(Colr5)
rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
XFader_FadeByColor5 = XFader_FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, Wavy)
End Function
Function XFader_FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, Wavy As Boolean)
'By: Xcliptic
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    XFader_FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function


Sub XFader_FadeForm(FormX As Form, Colr1, Colr2)
'By: Xcliptic

    B1 = XFader_GetRGB(Colr1).Blue
    G1 = XFader_GetRGB(Colr1).Green
    R1 = XFader_GetRGB(Colr1).Red
    B2 = XFader_GetRGB(Colr2).Blue
    G2 = XFader_GetRGB(Colr2).Green
    R2 = XFader_GetRGB(Colr2).Red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub



Function XFader_RGBtoHEX(RGB)
'By: Xcliptic

    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function

Sub Xroom_Buster()
'By: Xcliptic
'Needed:
'1 timer,and 2 command buttons
'in the timer put in this code...
AppActivate "Vplaces"
SendKeys "%OE"
'Then in the command you want to start the room bust with input
'this code...

Timer1.Enabled = True

'IN the stop Button put this code

Timer1.Enabled = False
'set the timers' interval to something high
'like 75 or 100 and there's your room bust.
'Yet another note you may wanna "disable" the timer
'or else whenever the form loads it will start busting
'on it's own
End Sub
