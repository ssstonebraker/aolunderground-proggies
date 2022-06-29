Attribute VB_Name = "CaLoReZ AIM40"
'THE FIRST EVER! AIM4.0 .BAS!
'by: CaLô ReZ
'OK maby you think im lie-ing but im not
'my uncle works for aol :-)
'so i got to gat aim4.0 :-)
'http://www.calo2k.cjb.net
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
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
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETtext = &H189
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

Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100

Public Const WM_LBUTTONDBLCLK = &H203


Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Function ATTENTION(text$)
SendChat "_-=\_-=\[A.T.T.E.N.T.I.O.N]/=-_/=-_"
SendChat Message$
SendChat "_-=\_-=\[A.T.T.E.N.T.I.O.N]/=-_/=-_"
End Function
Sub PlayWav(Sound As String)
X = sndPlaySound("" + Sound + "", 1):
     NoFreeze% = DoEvents()
End Sub
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function
Function GoToWebPage(page$)
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
Dim im As Long
On Error Resume Next
im& = FindWindow("AIM_IMessage", vbNullString)
name$ = GetCaption(im&)
If InStr(name$, "- Instant Message") <> 0 Then
text% = GetWindowTextLength(im&)
text% = (text%) - 19
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + text%)
SNfromIM = SN$
Else
SNfromIM = "-=[ Unknown ]=-"
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
Function Close_Aim()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "E&xit")
End Function
Function SignOffAIM()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Sign O&ff")
End Function
Function KillWindow(Window%)
KillWindow = SendMessageByNum(Window%, WM_CLOSE, 0, 0)
End Function
Function SendInvite(who$, Message$, chat$)
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, Edit1 As Long, Edit2 As Long, Edit3 As Long, send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit1&, WM_SETTEXT, 0, who$)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, Message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function
Function SendLinkBLUE(Link$, Message$)
SendChat "<a href=""" + Link$ + """><font color=#0000ff>" + Message$ + ""
End Function
Function SendLinkRED(Link$, Message$)
SendChat "<a href=""" + Link$ + """><font color=#ff0000>" + Message$ + ""
End Function
Function ClearChat()
Dim chat As Long, Box As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, "")
End Function
Function GetChatName()
Dim chat As Long
On Error Resume Next
chat& = FindWindow("AIM_ChatWnd", vbNullString)
roomname$ = GetCaption(chat&)
room$ = Mid(roomname$, InStr(roomname$, ":") + 2)
GetChatName = room$
End Function
Function IsAIMOnline()
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
IsAIMOnline = aim&
End Function
Function FindChatRoom()
Dim chat As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
FindChatRoom = chat&
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Sub HideAIMAdvertisement()
Dim part1 As Long, part2 As Long, HideIt As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
HideIt& = ShowWindow(part2&, SW_HIDE)

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
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function MassIM(List As ListBox, Message As String)
If List.ListCount = 0 Then
Do: DoEvents: Loop
End If
List.Enabled = False
I = List.ListCount - 1
List.ListIndex = 0
For X = 0 To I
List.ListIndex = X
Call SendIM(List.text, Message)
TimeOut (0.8)
Next X
List.Enabled = True
End Function

Function OpenIM()
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
End Function

Function SayRoomName()
SendChat "You have just entered room """ + GetChatName + "."""
End Function
Sub ClickIcon(ICON%)
Click% = SendMessage(ICON%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(ICON%, WM_LBUTTONUP, 0, 0&)
End Sub

Function SendChat(Message$)
Dim chat As Long, Box As Long, Box2 As Long, Box3 As Long, send As Long, send2 As Long, send3 As Long, send4 As Long, send5 As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(chat&, Box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, Message$)
send& = FindWindowEx(chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(chat&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(chat&, send3&, "_Oscar_IconBtn", vbNullString)
send5& = FindWindowEx(chat&, send4&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send5&)
End Function
Function SendIM(who$, Message$)
Dim aim As Long, Group As Long, Button As Long, im As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
im& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(im&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, who$)
IMmessage& = FindWindowEx(im&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, Message$)
send& = FindWindowEx(im&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send&)
End Function
Function ShowAIMadvertisement()
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = ShowWindow(part2&, SW_SHOW)
End Function
Function UserSN()
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
name$ = GetCaption(aim&)
If InStr(name$, "Buddy List") <> 0 Then
text% = GetWindowTextLength(aim&)
text% = (text%) - 14
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + text%)
UserSN = SN$
Else
UserSN = "-=[ WHO THE HELL KNOWZ ]=-"
End If
End Function

