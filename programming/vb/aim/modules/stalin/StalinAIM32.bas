Attribute VB_Name = "StalinAIM32"
'***********************************************
'************ -= StalinAIM32.bas =- ************
'************ -= Made by Stalin! =- ************
'****************** -= 99' =- ******************
'*********** -= Stalin000@aol.com =- ***********
'*********http://uprise.virtualave.net/*********
'***********************************************
'**** -= This bas file was made for VB5! =- ****
'******* -= Thanks for using my first =- *******
'****** -= ever AIM bas file. I'm gonna =- *****
'***** -= start making the second pretty =- ****
'*********** -= soon. Until Then! =- ***********
'***********************************************


Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

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
        x As Long
        y As Long
End Type







Function Attention(Message$)
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
SendChat Message$
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
End Function
Sub playwav(Wav$)
x = sndPlaySound("wav$", 1):
     NoFreeze% = DoEvents()

End Sub
Function GetClass(Child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(Child, Buffer$, 250)

GetClass = Buffer$
End Function
Function GoToWebPage(page$)
'Goes to a webpage using AIM
Dim aim As Long, Box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Go&)
End Function

Function HideAIM()
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = ShowWindow(aim&, SW_HIDE)
End Function

Function ShowAIM()
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = ShowWindow(aim&, SW_SHOW)
End Function


Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function ChangeAIMCaption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function

Function SNfromIM()
'Won't work if you've changed the IM caption!
Dim IM As Long
On Error Resume Next
IM& = FindWindow("AIM_IMessage", vbNullString)
name$ = GetCaption(IM&)
If InStr(name$, "- Instant Message") <> 0 Then
Text% = GetWindowTextLength(IM&)
Text% = (Text%) - 19
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + Text%)
SNfromIM = SN$
Else
SNfromIM = "(Unknown)"
End If
End Function

Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Function ChangeIMCaption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Function ChangeChatCaption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_ChatWnd", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function


Function exitaim()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "E&xit")
End Function

Function signoffAIM()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Sign O&ff")
End Function


Function KillWindow(Window%)
KillWindow = SendMessageByNum(Window%, WM_CLOSE, 0, 0)
End Function

Function MacroKill()
SendChat "~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~ ~@%@%@%Stalin%@%@%@~"
End Function

Function SendInvite(Who$, Message$, Chat$)
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, Edit1 As Long, Edit2 As Long, Edit3 As Long, send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit1&, WM_SETTEXT, 0, Who$)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, Message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, Chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function

Function SendLink(Link$, Message$)
SendChat "<a href=""" + Link$ + """>" + Message$ + ""
End Function

Sub TimeOut(Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop

End Sub

Function clearChat()
Dim Chat As Long, Box As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, "")

End Function

Function GetChatName()
'This won't work if you've changed the chat caption!
Dim Chat As Long
On Error Resume Next
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = GetCaption(Chat&)
Room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
GetChatName = Room$
End Function

Function IsAIMOnline()
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
IsAIMOnline = aim&
End Function

Function FindChatRoom()
Dim Chat As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
FindChatRoom = Chat&
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function Scroller(Message$)
'Scrolls 10 lines of your message
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
TimeOut (0.25)
SendChat Message$
End Function

Sub HideAIMAd()
'Hides that annoying advertisement
Dim part1 As Long, part2 As Long, HideIt As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
HideIt& = ShowWindow(part2&, SW_HIDE)
End Sub

Sub ClickIcon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

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
Room% = firs%
FindChildByTitle = Room%
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
Room% = firs%
FindChildByClass = Room%

End Function

Function GetText(Child)
GetTrim = SendMessageByNum(Child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(Child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function MassIM(List As ListBox, Message As String)
'Needs a list box and a text box
If List.ListCount = 0 Then
Do: DoEvents: Loop
End If
List.Enabled = False
i = List.ListCount - 1
List.ListIndex = 0
For x = 0 To i
List.ListIndex = x
Call SendIM(List.Text, Message)
TimeOut (0.8)
Next x
List.Enabled = True
End Function

Function openIM()
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
End Function

Function SayRoomName()
'Won't work if you've changed the chat caption!
SendChat "You have just entered room """ + GetChatName + "."""
End Function

Function ChatSend(Message$)
Dim Chat As Long, Box As Long, Box2 As Long, Box3 As Long, send As Long, send2 As Long, send3 As Long, send4 As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(Chat&, Box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, Message$)
send& = FindWindowEx(Chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Chat&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(Chat&, send3&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send4&)
End Function

Function SendIM(Who$, Message$)
Dim aim As Long, Group As Long, Button As Long, IM As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
IM& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(IM&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, Who$)
IMmessage& = FindWindowEx(IM&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, Message$)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send&)
End Function

Function ShowAIMAd()
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = ShowWindow(part2&, SW_SHOW)
End Function

Function UserSN()
'Won't work if you've changed the AIM caption!!!
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
name$ = GetCaption(aim&)
If InStr(name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 14
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + Text%)
UserSN = SN$
Else
UserSN = "(Unknown)"
End If
End Function

Sub GetInfo(Who As String)

Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Get Member Inf&o")
Do
CIO1% = FindWindow("_Oscar_Locate", vbNullString)
Loop Until CIO1% <> 0
NF1% = FindChildByClass(CIO1%, "_Oscar_PersistantComb")
NF2% = FindChildByClass(NF1%, "Edit")
NF3% = SendMessageByString(NF2%, WM_SETTEXT, 0, Who)
NF4% = FindChildByClass(CIO1%, "Button")
ClickIcon (NF4%)
ClickIcon (NF4%)
NF5% = FindChildByClass(CIO1%, "WndAte32Class")
NF6% = FindChildByClass(NF5%, "Ate32Class")
End Sub
