Attribute VB_Name = "Plå†ïñüm32"
' ------------------------------------
'| Plå†ïñüm32 by Plå†ïñüm for AOL 4.0 |
' ------------------------------------
'Version 1
'32 Bit Version of AOL Only!
'E-mail me at dome@www2000.net
'Codes for this can be found at
'www.nwozone.com/knk4o/index.htm
'I made this bas file so that
'You guys could try to learn VB by
'Playing around with the functions
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CreateMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
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
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesaredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
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
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Type TheRgb
   Red As Long
   Green As Long
   Blue As Long
End Type

Function AOLWindow() As Long
AOLWindow& = FindWindow("AOL Frame25", vbNullString)
End Function
Function AOLMDI() As Long
AOLMDI& = FindWindowEx(AOLWindow, 0&, "MDIClient", vbNullString)
End Function
Sub Key(Keywd As String)
'keyword
Tool1& = FindWindowEx(AOLWindow, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(Tool1&, 0&, "_AOL_Toolbar", vbNullString)
Box& = FindWindowEx(tool2&, 0&, "_AOL_Combobox", vbNullString)
Box& = FindWindowEx(Box&, 0&, "Edit", vbNullString)
SendMessageByString Box&, WM_SETTEXT, 0&, Keywd
SendMessageLong Box&, WM_CHAR, VK_SPACE, 0&
SendMessageLong Box&, WM_CHAR, VK_RETURN, 0&
End Sub
Sub SendIM(PERSON As String, message As String, Optional Bold As Boolean, Optional Italics As Boolean, Optional Underline As Boolean, Optional Strikeout As Boolean)
'bold, italics, strikeout, and
'underlined are optional statements
'You could have IM(ScreenName,Message)
'or IM(Screenname,message,true,true,true,true)
If Bold = True Then text = "<b>" & text & "</b>"
If Italics = True Then text = "<i>" & text & "</i>"
If Strikeout = True Then text = "<s>" & text & "</s>"
If Underline = True Then text = "<u>" & text & "</u>"
Call Key("aol://9293:" & PERSON)
Do: DoEvents
IMWin& = FindWindowEx(AOLMDI, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
SendButton& = FindWindowEx(IMWin&, 0&, "_AOL_Icon", vbNullString)
For X = 1 To 8
SendButton& = FindWindowEx(IMWin&, SendButton&, "_AOL_Icon", vbNullString)
Next
Loop Until IMWin& <> 0 And Rich& <> 0 And SendButton& <> 0
WindowHide IMWin&
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message)
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
Call SendIM("$IM_ON", " ")
End Sub
Sub IMsOff()
Call SendIM("$IM_OFF", " ")
End Sub
Sub AddRoomToListBox(thelist As ListBox, AddUser As Boolean)
'Taken From Dos32.bas
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChatRoom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> GetUserSn Or AddUser = True Then
If ListSearch(thelist, ScreenName$) = -1 Then thelist.AddItem ScreenName$
End If
Next Index&
Call CloseHandle(mThread)
End If
End Sub

Sub AddRoomToComboBox(TheCombo As ComboBox, AddUser As Boolean)
'Taken From Dos32.bas
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChatRoom&
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
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
Next Index&
Call CloseHandle(mThread)
End If
If TheCombo.ListCount > 0 Then
TheCombo.text = TheCombo.List(0)
End If
End Sub
Function FindChatRoom() As Long
child& = FindWindowEx(AOLMDI, 0&, "AOL Child", vbNullString)
Do: DoEvents
thelist& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AOLCombo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
If AOLIcon& <> 0 And Rich& <> 0 And AOLStatic& <> 0 And thelist& <> 0 And AOLCombo& <> 0 Then Exit Do
child& = FindWindowEx(AOLMDI, child&, "AOL Child", vbNullString)
Loop Until child& = 0
FindChatRoom& = child&
End Function
Function GetRoomName() As String
'name of the room
GetRoomName = GetCaption(FindChatRoom)
End Function
Function GetCaption(WindowHandle As Long) As String
'Taken From Dos32.bas
Dim Buffer As String, TextLength As Long
TextLength& = GetWindowTextLength(WindowHandle&)
Buffer$ = String(TextLength&, 0&)
Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
GetCaption$ = Buffer$
End Function
Function GetClass(child)
'Taken From Dos32.bas
Dim Buffer$
Dim getclas%

Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Sub SendChat(text As String, Optional Bold As Boolean, Optional Italics As Boolean, Optional Underline As Boolean, Optional Strikeout As Boolean, Optional WavY As Boolean)
'bold, italics, strikeout, and
'underlined are optional statements
'You could have sendchat(text)
'or sendchat(text,true,true,true,true)
If WavY = True Then text = WavyChat(text)
If Bold = True Then text = "<b>" & text & "</b>"
If Italics = True Then text = "<i>" & text & "</i>"
If Strikeout = True Then text = "<s>" & text & "</s>"
If Underline = True Then text = "<u>" & text & "</u>"
thechat& = FindChatRoom
If thechat& = 0 Then Exit Sub
Box = FindWindowEx(thechat&, 0&, "RICHCNTL", vbNullString)
Box = FindWindowEx(thechat&, Box, "RICHCNTL", vbNullString)
SendMessageByString Box, WM_SETTEXT, 0&, ""
SendMessageByString Box, WM_SETTEXT, 0&, text
SendMessageLong Box, WM_CHAR, 13, 0&
Pause 0.33
End Sub
Sub PrivateRoom(Room As String)
'enter a private room
Call Key("aol://2719:2-2-" & Room)
End Sub
Sub WindowHide(hwnd As Long)
'hides a window
Call ShowWindow(hwnd&, SW_HIDE)
End Sub

Sub WindowShow(hwnd As Long)
'shows a window
Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Function GetText(WindowHandle As Long) As String
TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
GetText$ = Buffer$
End Function
Function GetChatText() As String
'gets chat text
Box& = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
GetChatText = GetText(Box)
End Function
Function GetChatBox() As Long
'handle of chat box
Box = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
GetChatBox = FindWindowEx(FindChatRoom, Box, "RICHCNTL", vbNullString)
End Function
Function LastChatLineWithSN() As String
'last chat line with the screen name
TheStr = GetChatText
For T = Len(TheStr) To 1 Step -1
If Mid(TheStr, T, 1) = Chr(13) Then Exit For Else totalstr = Mid(TheStr, T, 1) & totalstr
Next
LastChatLineWithSN = totalstr
End Function
Function LastChatLine() As String
'last chat line
TheStr = LastChatLineWithSN
For T = Len(TheStr) To 1 Step -1
If Mid(TheStr, T, 1) = ":" Then Exit For Else totalstr = Mid(TheStr, T, 1) & totalstr
Next
LastChatLine = Mid(totalstr, 3, Len(totalstr))
End Function
Function LastChatSn() As String
'screen name from last chat line
TheStr = LastChatLineWithSN
For T = 1 To Len(TheStr)
If Mid(TheStr, T, 1) = ":" Then Exit For Else totalstr = totalstr & Mid(TheStr, T, 1)
Next
LastChatSn = totalstr
End Function
Function ListSearch(lst As ListBox, txt As String) As Integer
'exact search for a list
Dim X%
For X% = 0 To lst.ListCount - 1
If UCase(txt) = UCase(lst.List(X%)) Then
ListSearch = X%
Exit Function
End If
Next X%
ListSearch = -1
End Function
Sub IMAnswer(SN As String, message As String)
'Use a * in sn if you want to answer all new im's
'answers an im from a specific person
If SN = "*" Then GoTo doall
thewin = FindWindowEx(AOLMDI, 0&, "AOL Child", ">Instant Message from: " & SN)
If thewin = 0 Then Exit Sub Else
text = FindWindowEx(thewin, 0&, "RICHCNTL", vbNullString)
typein = FindWindowEx(thewin, text, "RICHCNTL", vbNullString)
If typein = 0 Then GoTo openit
typeinit:
typein = FindWindowEx(thewin, text, "RICHCNTL", vbNullString)
SendMessageByString typein, WM_SETTEXT, 0&, message
butt = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
For Butto = 1 To 8
butt = FindWindowEx(thewin, butt, "_AOL_Icon", vbNullString)
Next
ClickIcon (butt)
Exit Sub
openit:
G = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
G = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
ClickIcon (G)
GoTo typeinit
doall:
Do
thewin = FindChildByTitle(AOLMDI, ">Instant Message from: ")
If thewin = 0 Then Exit Sub
text = FindWindowEx(thewin, 0&, "RICHCNTL", vbNullString)
typein = FindWindowEx(thewin, text, "RICHCNTL", vbNullString)
If typein = 0 Then GoTo openit2
typeinit2:
typein = FindWindowEx(thewin, text, "RICHCNTL", vbNullString)
SendMessageByString typein, WM_SETTEXT, 0&, message
butt = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
For Butto = 1 To 8
butt = FindWindowEx(thewin, butt, "_AOL_Icon", vbNullString)
Next
ClickIcon (butt)
GoTo loopit
openit2:
G = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
G = FindWindowEx(thewin, 0&, "_AOL_Icon", vbNullString)
ClickIcon (G)
GoTo typeinit2
loopit:
Loop
End Sub
Sub ClickIcon(hwnd As Long)
'clicks an icon
Call SendMessage(hwnd, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(hwnd, WM_LBUTTONUP, 0&, 0&)
End Sub
Function FindChildByTitle(parenthwnd As Long, childtitle As String, Optional starthwnd As Long)
'I made this so I could find child
'windows without knowing the exact
'title.
'If you know the exact title, you can
'use this or FindWindowEx
G& = FindWindowEx(parenthwnd, 0&, vbNullString, vbNullString)
If InStr(UCase(GetCaption(G&)), UCase(childtitle)) Then FindChildByTitle = G&: Exit Function
G& = starthwnd
Do
DoEvents
G& = FindWindowEx(parenthwnd, G&, vbNullString, vbNullString)
If InStr(GetCaption(G&), childtitle) Then Exit Do
Loop Until G = 0
FindChildByTitle = G&
End Function
Function GetUserSn() As String
'gets user's screen name
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
If WelcomeWin& = 0& Then GetUserSn$ = "Not Online": Exit Function
UserSN$ = GetCaption(WelcomeWin&)
GetUserSn$ = Mid(UserSN$, 10, Len(UserSN$) - 10)
End Function
Sub AddBuddyListToList(thelist As ListBox)
On Error Resume Next
Dim cProcess As Long, itmHold As Long, ScreenName As String
Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
Dim rList As Long, sThread As Long, mThread As Long
Room& = FindChildByTitle(AOLMDI, "Buddy List Window")
If Room& = 0& Then Exit Sub
rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
If mThread& Then
For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
ScreenName$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
psnHold& = psnHold& + 6
ScreenName$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
If ScreenName$ <> GetUserSn$ Or AddUser = True Then
If ListSearch(thelist, ScreenName$) = -1 Then thelist.AddItem ScreenName$
End If
Next Index&
Call CloseHandle(mThread)
End If
End Sub
Function ReverseString(TheStr As String) As String
'AOL would return LOA
For G = 1 To Len(TheStr)
newstr = Mid(TheStr, G, 1) & newstr
Next
ReverseString = newstr
End Function
Function TrimSpaces(TheStr As String) As String
'gets rid of all spaces in a string.
For G = 1 To Len(TheStr)
If Mid(TheStr, G, 1) = " " Then Else newstr = newstr & Mid(TheStr, G, 1)
Next
TrimSpaces = newstr
End Function
Function TrimChar(TheStr As String, TheChar As String) As String
'trims a char from a string
If Len(TheChar) <> 1 Then Exit Function
For G = 1 To Len(TheStr)
If Mid(TheStr, G, 1) = TheChar Then Else newstr = newstr & Mid(TheStr, G, 1)
Next
TrimChar = newstr
End Function
Function TrimString(text As String, TheStr As String) As String
'takes one string out of a string.
'ex
'trimstring("Dome32 is the best!","the")
'returns "Dome32 is best!"
For G = 1 To Len(text)
If Mid(text, G, Len(TheStr)) = TheStr Then G = G + Len(TheStr) - 1 Else newstr = newstr & Mid(text, G, 1)
Next
TrimString = newstr
End Function
Function ReplaceString(text As String, Replace As String, ReplaceWith As String)

For G = 1 To Len(text)
If Mid(text, G, Len(Replace)) = Replace Then newstr = newstr & ReplaceWith: G = G + Len(Replace) - 1 Else newstr = newstr & Mid(text, G, 1)
Next
ReplaceString = newstr
End Function
Function SpacedString(text As String) As String
'if you had "Visual Basic" as the text
'it would return "V i s u a l  B a s i c"
For X = 1 To Len(text)
newstr = newstr & Mid(text, X, 1) & " "
Next
SpacedString = Mid(newstr, 1, Len(newstr) - 1)
End Function
Sub Mail(PERSON As String, SUBJECT As String, message As String)
'sends mail
tool& = FindWindowEx(AOLWindow, 0&, "AOL Toolbar", vbNullString)
tool2& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
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
Call SendMessageByString(edit1&, WM_SETTEXT, 0, PERSON$)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents
Call SendMessageByString(Edit&, WM_SETTEXT, 0, SUBJECT$)
ClickIcon (butt&)
End Sub
Sub LinkToChat(URL As String, urlname As String)
'sends a link to chat
G$ = "< a href=""" & URL & """>" & urlname & "</a>"
SendChat G$, False, False, False, False
End Sub
Sub MailToLink(mailto As String, Linktext As String)
'puts a mailto link in the chatroom
G$ = "< a href=""mailto:" & mailto & """>" & Linktext & "</a>"
SendChat G$, False, False, False, False
End Sub
Sub SpiralChat(text As String)
'example AOL
'aol
'ola
'lao
'aol
For B = 1 To Len(text)
h$ = Mid(text, B, Len(text)) & Mid(text, 1, B - 1)
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & h$ & "</b><font color=#000000> }–•"
Pause 0.5
Next
Pause 0.5
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & text & "</b><font color=#000000> }–•"
End Sub
Sub SpiralIM(PERSON As String, text As String)
'same as spiral chat only in im
If Len(text) * (Len(text) + 1) > 592 Then MsgBox "Message to long!", vbOKOnly, App.Title
For B = 1 To Len(text)
h$ = h$ & Chr(13) & Mid(text, B, Len(text)) & Mid(text, 1, B - 1)
Next
h$ = h$ & Chr(13) & text
SendIM PERSON, h$, False, False, False, False
End Sub

Function WavyChat(text As String)
'makes wavy text for the chatroom,
'im or mail.
Dim e As String
For X = 1 To Len(text)
wave = wave + 1
If wave > 4 Then wave = 1
If wave = 4 Then wavetext = "</sub>"
If wave = 3 Then wavetext = "<sub>"
If wave = 2 Then wavetext = "</sup>"
If wave = 1 Then wavetext = "<sup>"
G$ = G$ & Mid(text, X, 1) & wavetext
Next
WavyChat = G$
End Function
Sub WebSearch(search As String)
'searches the internet for a string
Key "http://search.yahoo.com/bin/search?p=" & search
End Sub
Sub ChatExtender(text As String)
'lets you type more than the 92 char
'limit
If Len(text) < 92 Then SendChat text, False, False, False, False: Exit Sub
F = Int(Len(text) / 92) + 1
For X = 0 To F
SendChat Mid(text, X * 92 + 1, 92), False, False, False, False
Pause 0.5
Next
End Sub
Function BoldRotate(text As String)
'makes text bold, not bold...
For X = 1 To Len(text)
reg = reg + 1
If reg > 2 Then reg = 1
If reg = 1 Then html = "<b>"
If reg = 2 Then html = "</b>"
If Mid(text, X, 1) = " " Then BoldRotate = BoldRotate & " ": reg = reg + 1: GoTo loopit
BoldRotate = BoldRotate & Mid(text, X, 1) & html
loopit:
Next
End Function
Function BoldItalicRotate(text As String)
'makes text bold then italic then nothing

For X = 1 To Len(text)
reg = reg + 1
If reg > 4 Then reg = 1
If reg = 1 Then html = "<b>"
If reg = 2 Then html = "</b>"
If reg = 3 Then html = "<i>"
If reg = 4 Then html = "</i>"
If Mid(text, X, 1) = " " Then BoldItalicRotate = BoldItalicRotate & " ": reg = reg + 1: GoTo loopit
BoldItalicRotate = BoldItalicRotate & Mid(text, X, 1) & html
loopit:
Next
End Function
Sub Macro(text As TextBox)
'scrolls a multilined textbox to
'the chatroom one line at a time

For ctri% = 1 To CountTextBoxLines(text)
G$ = GetTextBoxLine(text, ctri%)
SendChat G$, False, False, False, False
Pause 0.5
Next
End Sub
Public Sub ChatIgnoreByIndex(Index As Long)
'From Dos32
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, Count As Long
    Count& = RoomCount&
    If Index& > Count& - 1 Then Exit Sub
    Room& = FindChatRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, Index&, 0&)
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
'From Dos32
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindChatRoom&
    If Room& = 0& Then ChatIgnoreByName = False: Exit Function
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUared, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If GetUserSn <> ScreenName$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = Index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                ChatIgnoreByName = True
                Exit Function
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
ChatIgnoreByName = False
End Function
Function TimeUntilDate(date1 As Date, date2 As Date)
S = DateDiff("s", date1, date2)
m = S / 60
h = m / 60
d = h / 24
dd = Int(d)
h = (d - dd) * 24
hh = Int(h)
m = (h - hh) * 60
mm = Int(m)
S = (m - mm) * 60
ss = Int(S)
TimeUntilDate = dd & " days, " & hh & " hours, " & mm & " minutes, " & ss & " seconds to " & date2
End Function
Function LastMessageFromIM()
'last message from an im
IMWin& = FindChildByTitle(AOLMDI, ">Instant Message From")
If IMWin& = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message From")
If IMWin = 0& Then IMWin& = FindChildByTitle(AOLMDI, "  Instant Message To")
If IMWin& = 0& Then Exit Function
richtext& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
text = GetText(richtext&)
Do
Find& = B&
B& = InStr(B& + 1, text, ":")
Loop Until B& <= 0&
LastMessageFromIM = Mid(text, Find& + 3, Len(text))
End Function
Function SNfromIM()
'gets sn from first im
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
If FindChildByTitle(AOLMDI, "Welcome,") <> 0 Then UserOnline = True Else UserOnline = False
End Function
Sub IMIgnore(PERSON As String)
SendIM "$IM_OFF, " & PERSON, ".", False, False, False, False
End Sub
Sub IMUnIgnore(PERSON As String)
SendIM "$IM_ON, " & PERSON, ".", False, False, False, False
End Sub
Sub IMIgnoreList(lst As ListBox)
'ignores an im if from some on the list
Do: DoEvents
X& = FindChildByTitle(AOLMDI, ">Instant Message From", X&)
L = GetText(X&)
B$ = Mid(L, 24, Len(L))
If ListSearch(lst, B$) = -1 Then Exit Do
SendMessageLong X&, WM_CLOSE, 0, 0
Loop Until X& = 0
End Sub
Function FindNewIM(SN As String)
'finds new im
thewin = FindWindowEx(AOLMDI, 0&, "AOL Child", ">Instant Message from: " & SN)
FindIM = thewin
End Function
Function TheTime()
TheTime = Format(Now, "h:mm:ss AM/PM")
End Function
Sub CloseAllIMs()
'closes all open ims
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
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function CoolChat(text As String)
'takes a random font and sends it to
'chat
B = Mid(Time, Len(Format(Time, "h:mm:ss")) - 2, 2)
For m = 1 To Len(text)
If B > Printer.FontCount Then B = 1 Else B = B + 1
ob$ = Printer.Fonts(B)
aSt$ = aSt$ & "<font face=""" & ob$ & """>" & Mid(text, m, 1)
Next
CoolChat = aSt$
End Function
Function ListSearch2(lst As ListBox, txt As String) As Integer
'list search
Dim X%
For X% = 0 To lst.ListCount - 1
If InStr(UCase(lst.List(X%)), UCase((txt))) Then
ListSearch2 = X%
Exit Function
End If
Next X%
ListSearch2 = -1
End Function

Sub ClearChat()
'Only the user can see this clearchat
h = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
SendMessageByString h, WM_SETTEXT, 0&, ""
End Sub
Function RandomColor() As String
'give random color
RandomColor = ""
For m = 1 To 3
Randomize
rand = Int((255 * Rnd) + 1)
RandomColor = RandomColor & Hex(rand)
Next
End Function

Public Function RoomCount() As Long
'dos32
'how many people in room
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindChatRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function
Sub CircleGradient(Frm As Form, rs%, gs%, bs%, re%, ge%, be%, smooth As Boolean)
'Visit DiP's VB World At:
'http://come.to/dipsvbworld
'from dip's vbworld

If Frm.WindowState = vbMinimized Then Exit Sub
Frm.BackColor = RGB(re, ge, be)
If smooth = True Then
Frm.DrawStyle = 6
Else
Frm.DrawStyle = 0
End If
If Frm.ScaleWidth <> 255 Then
Frm.ScaleWidth = 255
End If
If Frm.ScaleHeight <> 255 Then
Frm.ScaleHeight = 255
End If
Frm.DrawWidth = 5
Frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (bs - be) / 255
rc = rs: bc = bs: gc = gs
For X = 0 To 255
Frm.Circle (Frm.ScaleWidth / 2, Frm.ScaleHeight / 2), X, RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
End Sub
Sub RectGradient(Frm As Form, rs%, gs%, bs%, re%, ge%, be%, smooth As Boolean)
'Visit DiP's VB World At:
'http://come.to/dipsvbworld
'more form gradients

If Frm.WindowState = vbMinimized Then Exit Sub
Frm.BackColor = RGB(re, ge, be)
If smooth = True Then
Frm.DrawStyle = 6
Else
Frm.DrawStyle = 0
End If
If Frm.ScaleWidth <> 255 Then
Frm.ScaleWidth = 255
End If
If Frm.ScaleHeight <> 255 Then
Frm.ScaleHeight = 255
End If
Frm.DrawWidth = 5
Frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (bs - be) / 255
rc = rs: bc = bs: gc = gs
For X = 255 To 0 Step -1
DoEvents
Frm.Line ((X / 2), (X / 2))-(Frm.ScaleWidth - (X / 2), Frm.ScaleHeight - (X / 2)), RGB(rc, gc, bc), B
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
End Sub

Sub DiagGradient(Frm As Form, rs%, gs%, bs%, re%, ge%, be%, smooth As Boolean)
'Visit DiP's VB World At:
'http://come.to/dipsvbworld
'more gradients
If Frm.WindowState = vbMinimized Then Exit Sub
Frm.BackColor = RGB(re, ge, be)
If smooth = True Then
Frm.DrawStyle = 6
Else
Frm.DrawStyle = 0
End If
If Frm.ScaleWidth <> 255 Then
Frm.ScaleWidth = 255
End If
If Frm.ScaleHeight <> 255 Then
Frm.ScaleHeight = 255
End If
Frm.DrawWidth = 5
Frm.Refresh
ri = (rs - re) / 255
gi = (gs - ge) / 255
bi = (bs - be) / 255
rc = rs: bc = bs: gc = gs
For X = 0 To 255
DoEvents
Frm.Line (0, X)-(X, 0), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
ri = (re - rs) / 255
gi = (ge - gs) / 255
bi = (be - bs) / 255
rc = re: bc = be: gc = ge
For X = 255 To 0 Step -1
DoEvents
Frm.Line (255 - X, 255)-(255, 255 - X), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
End Sub

Sub SpinGradient(Frm As Form, rs%, gs%, bs%, re%, ge%, be%, smooth As Boolean)
'Visit DiP's VB World At:
'http://come.to/dipsvbworld
'more gradients
If Frm.WindowState = vbMinimized Then Exit Sub
Frm.BackColor = RGB(rs, gs, bs)
If smooth = True Then
Frm.DrawStyle = 6
Else
Frm.DrawStyle = 0
End If
If Frm.ScaleWidth <> 255 Then
Frm.ScaleWidth = 255
End If
If Frm.ScaleHeight <> 255 Then
Frm.ScaleHeight = 255
End If
Frm.DrawWidth = 5
Frm.Refresh
ri = (rs - re) / 255 / 2
gi = (gs - ge) / 255 / 2
bi = (bs - be) / 255 / 2
rc = rs: bc = bs: gc = gs
For X = 0 To 255
DoEvents
Frm.Line (X, 0)-(255 - X, 255), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
For X = 0 To 255
DoEvents
Frm.Line (255, X)-(0, 255 - X), RGB(rc, gc, bc)
rc = rc - ri
gc = gc - gi
bc = bc - bi
Next X
End Sub
Function FileExists(sFileName As String) As Integer
'finds if a file exists
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


Sub playwav(wav As String)
'program waits till wave has ended
'to keep going
'if you want to change this use
'this instead
'w% = sndPlaySound(wav, 1)
W% = sndPlaySound(wav, 0)
End Sub
Sub PlayMidi(midilocation As String)
'midilocation ex:
'c:\windows\midi.mid
Dim ret As Long
ret = mciSendString("play " & midilocation, 0&, 0, 0)
End Sub
Sub StopMidi(midlocation As String)
Dim ret As Long
ret = mciSendString("stop " & midilocation, 0&, 0, 0)
End Sub
Function LoadtxtToText(TxtFilename As String, TextboxToload As TextBox)
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

Function Savetext(TxtFilename As String, TxtTextbox As TextBox)
Dim sFile As String
Dim nFile As Integer

nFile = FreeFile
sFile = TxtFilename
Open sFile For Output As nFile
Print #nFile, TxtTextbox
Close nFile


End Function

Function GetTextBoxLine(text As TextBox, theline As Integer) As String
'gets a specific line from textbox
U = CountTextBoxLines(text)
If theline > U Then Exit Function
For G = 1 To theline - 1
e = InStr(e + 1, text.text, Chr(13))
Next
e = e + 1
If e = 1 Then e = 0
For F = 1 To theline
j = InStr(j + 1, text.text, Chr(13))
Next
If j = 0 Then j = Len(text.text) + 1
GetTextBoxLine = Mid(text.text, e + 1, j - 1 - e)
End Function
Function CountTextBoxLines(text As TextBox) As Long
'counts lines in a textbox
i = 0
Do
a = InStr(a + 1, text.text, Chr(13))
If a = 0 Then Exit Do
i = i + 1
Loop
CountTextBoxLines = i + 1
End Function

Public Function FindInfoWindow() As Long
'dos32
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

Function TypeOut(text As String)
'makes a string type out one letter
'at a time in chat, mail or im.
lagtext$ = "</html>"
If text = "" Then Exit Function
For j = 1 To Len(text)
lagtext$ = lagtext$ & Mid(text, j, 1) & "<html></html>"
Next
TypeOut = lagtext$
End Function
Sub OnTop(Frm As Form)
'put your form on top of all others
Dim ret As Long
On Error Resume Next
ret& = SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

' Fade Subs

Function HexRGB(thecolor)

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
       acol = colr
       GetRGB.Red = acol Mod &H100
       acol = acol \ &H100
       GetRGB.Green = acol Mod &H100
       acol = acol \ &H100
       GetRGB.Blue = acol Mod &H100
End Function
Function Fade2Color(col1 As Long, col2 As Long, text As String, WavY As Boolean)
'two color fade

If col1 = FFFFFF Then col1 = FFFFFE
If col2 = FFFFFF Then col2 = FFFFFE
R1 = GetRGB(col1).Red
G1 = GetRGB(col1).Green
B1 = GetRGB(col1).Blue
R2 = GetRGB(col2).Red
G2 = GetRGB(col2).Green
B2 = GetRGB(col2).Blue

rdiff = Int((R1 - R2) / Len(TrimSpaces(text)))
gdiff = Int((G1 - G2) / Len(TrimSpaces(text)))
bdiff = Int((B1 - B2) / Len(TrimSpaces(text)))

For h = 1 To Len(text)
If Mid(text, h, 1) = " " Then thes$ = thes$ & " ": h = h + 1 Else
If WavY = True Then
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
thes$ = thes$ & "<font color=" & HexRGB(newrgb&) & ">" & Mid(text, h, 1) & html
R1 = R1 - rdiff
G1 = G1 - gdiff
B1 = B1 - bdiff
Next
Fade2Color = thes$
End Function
Function Fade3Color(col1 As Long, col2 As Long, col3 As Long, text As String, WavY As Boolean)

If Len(text) <= 7 Then Fade3Color = Fade2Color(col1, col2, text, WavY): Exit Function
lens = Int(Len(text) / 2)
Fade3Color = Fade2Color(col1, col2, Mid(text, 1, lens), WavY)
Fade3Color = Fade3Color & Fade2Color(col2, col3, Mid(text, lens + 1, Len(text) - lens), WavY)
End Function
Function Fade4Color(col1 As Long, col2 As Long, col3 As Long, col4 As Long, text As String, WavY As Boolean)
If Len(text) <= 7 Then Fade4Color = Fade2Color(col1, col2, text, WavY): Exit Function
lens = Int(Len(text) / 3)
Fade4Color = Fade2Color(col1, col2, Mid(text, 1, lens), WavY)
Fade4Color = Fade4Color & Fade2Color(col2, col3, Mid(text, lens + 1, lens), WavY)
Fade4Color = Fade4Color & Fade2Color(col3, col4, Mid(text, lens + lens + 1, Len(text) - lens), WavY)
End Function
Function Fade5color(col1 As Long, col2 As Long, col3 As Long, col4 As Long, col5 As Long, text As String, WavY As Boolean)
If Len(text) <= 7 Then Fade5color = Fade2Color(col1, col2, text, WavY): Exit Function
lens = Int(Len(text) / 4)
Fade5color = Fade2Color(col1, col2, Mid(text, 1, lens), WavY)
Fade5color = Fade5color & Fade2Color(col2, col3, Mid(text, lens + 1, lens), WavY)
Fade5color = Fade5color & Fade2Color(col3, col4, Mid(text, lens + lens + 1, lens), WavY)
Fade5color = Fade5color & Fade2Color(col4, col5, Mid(text, lens + lens + lens + 1, Len(text) - lens), WavY)
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
Dim X As Integer
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
For X = 0 To 255 Step Speed
col$ = RGB(StartRed + RedChange, StartGreen + GreenChange, StartBlue + BlueChange) 'Draws Line With correct color
BGFade = BGFade & "<BODY BGCOLOR=" & HexRGB(Val(col$)) & ">"
For F = 1 To Speed
RedChange = RedChange + (EndRed - StartRed) / 255 '
GreenChange = GreenChange + (EndGreen - StartGreen) / 255
BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
Next F
Next X
End Function


Function BGFadeWithMessage(col1 As Long, col2 As Long, newstr As String)
Dim X As Integer
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
For X = 0 To 255
col$ = RGB(StartRed + RedChange, StartGreen + GreenChange, StartBlue + BlueChange) 'Draws Line With correct color
BGFadeWithMessage = BGFadeWithMessage & "<BODY BGCOLOR=" & HexRGB(Val(col$)) & ">"
If X Mod letters = 0 Then BGFadeWithMessage = BGFadeWithMessage & Mid(newstr, U, 1): U = U + 1
RedChange = RedChange + (EndRed - StartRed) / 255 '
GreenChange = GreenChange + (EndGreen - StartGreen) / 255
BlueChange = BlueChange + (EndBlue - StartBlue) / 255 '
Next X
End Function
Function RoomOrOK() As String
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
Dim WelcomeWin&
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
WindowHide (WelcomeWin&)
End Sub
Sub ShowWelcome()
WelcomeWin& = FindChildByTitle(AOLMDI, "Welcome, ")
WindowShow (WelcomeWin&)
End Sub
Sub RemoveFromListbyString(List As ListBox, remove As String)
G = ListSearch(List, remove)
If G = -1 Then Exit Sub
List.RemoveItem G
End Sub
Sub HideAOL()
WindowHide (AOLWindow)
End Sub
Sub ShowAOL()
WindowShow (AOLWindow)
End Sub
Function Uppercase(text As String) As String
Uppercase = UCase(text)
End Function
Function Lowercase(text As String) As String
Lowercase = LCase(text)
End Function
Function HackerText(text As String) As String
For h = 1 To Len(text)
If h Mod 2 <> 0 Then Mid(text, h, 1) = UCase(Mid(text, h, 1)) Else Mid(text, h, 1) = LCase(Mid(text, h, 1))
Next
HackerText = text
End Function
Function EncryptText(text As String) As String
For h = 1 To Len(text)
theasc = Asc(Mid(text, h, 1)) + 4
If theasc > 255 Then theasc = (theasc - 255)
d$ = d$ & Chr(theasc)
Next
EncryptText = d$
End Function
Function DecryptText(text As String) As String
For h = 1 To Len(text)
theasc = Asc(Mid(text, h, 1)) - 4
If theasc <= 0 Then theasc = (255 + theasc)
d$ = d$ & Chr(theasc)
Next
DecryptText = d$
End Function
Sub SetClipboardText(text As String)
Clipboard.Clear
Clipboard.SetText text
End Sub
Function GetClipboardText() As String
GetClipboardText = Clipboard.GetText
End Function
