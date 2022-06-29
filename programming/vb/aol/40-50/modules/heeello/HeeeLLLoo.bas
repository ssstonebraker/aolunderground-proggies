Attribute VB_Name = "HeeeLLLoo"
'                HeeeLLLoo Version 1.0
'This bas includes functions for AIM and AOL 4.0 combined.
'I made this bas for my program, it will not be released
'for a while because I am still trying to get some of
'the bugs out.  It is going to be call "AIM Toolz."
'If you are having trouble with this bas or you have
'request for up coming programs, email me at
'TopCard23@hotmail.com or TopCardBak@aol.com

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
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
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

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
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        Y As Long
End Type







Function Attention(Message$)
ChatSend "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
ChatSend Message$
ChatSend "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
End Function
Sub Playwav(Wav$)
x = sndPlaySound("wav$", 1):
     NoFreeze% = DoEvents()

End Sub
Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Function GoToWebPage(page$)
'Goes to a webpage using AIM
Dim aim As Long, Box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMEssageByString(Box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Go&)
End Function

Function HideAIM()
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = showwindow(aim&, SW_HIDE)
End Function

Function ShowAIM()
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = showwindow(aim&, SW_SHOW)
End Function


Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Function Change_AIM_Caption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
captionchange& = SendMEssageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function

Function SNfromIM()
'Won't work if you've changed the IM caption!
Dim IM As Long
On Error Resume Next
IM& = FindWindow("AIM_IMessage", vbNullString)
Name$ = GetCaption(IM&)
If InStr(Name$, "- Instant Message") <> 0 Then
Text% = GetWindowTextLength(IM&)
Text% = (Text%) - 19
sn$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
SNfromIM = sn$
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
Function Change_AIM_IM_Caption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
captionchange& = SendMEssageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Function Change_AIM_Chat_Caption(Caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_ChatWnd", vbNullString)
captionchange& = SendMEssageByString(caption1&, WM_SETTEXT, 0, Caption$)
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
Call SendMEssageByString(Edit1&, WM_SETTEXT, 0, Who$)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMEssageByString(Edit2&, WM_SETTEXT, 0, Message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMEssageByString(Edit3&, WM_SETTEXT, 0, Chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function

Function SendLink(Link$, Message$)
SendChat "<a href=""" + Link$ + """>" + Message$ + ""
End Function

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Function clearChat()
Dim Chat As Long, Box As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Call SendMEssageByString(Box&, WM_SETTEXT, 0, "")

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
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
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
HideIt& = showwindow(part2&, SW_HIDE)
End Sub

Sub ClickIcon(icon%)
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

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMEssageByString(child, 13, GetTrim + 1, TrimSpace$)
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
Box3& = SendMEssageByString(Box2&, WM_SETTEXT, 0, Message$)
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
Call SendMEssageByString(IMto&, WM_SETTEXT, 0, Who$)
IMmessage& = FindWindowEx(IM&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMEssageByString(IMmessage&, WM_SETTEXT, 0, Message$)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send&)
End Function

Function ShowAIMAd()
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = showwindow(part2&, SW_SHOW)
End Function

Function UserSN()
'Won't work if you've changed the AIM caption!!!
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Name$ = GetCaption(aim&)
If InStr(Name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 14
sn$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
UserSN = sn$
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
NF3% = SendMEssageByString(NF2%, WM_SETTEXT, 0, Who)
NF4% = FindChildByClass(CIO1%, "Button")
ClickIcon (NF4%)
ClickIcon (NF4%)
NF5% = FindChildByClass(CIO1%, "WndAte32Class")
NF6% = FindChildByClass(NF5%, "Ate32Class")
End Sub

Public Sub AOL_SignOff()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'Ignore Since Of New AOL.
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
ClickIcon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
ClickIcon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Public Sub EliteTalker()
On Error Resume Next
Do
If InStr(Message$, "A") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "A") - 1) + "Ã" + Right$(Message$, Len(Message$) - InStr(Message$, "A"))
Message$ = macstringz
Loop Until InStr(Message$, "A") = 0
Do
If InStr(Message$, "a") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "a") - 1) + "å" + Right$(Message$, Len(Message$) - InStr(Message$, "a"))
Message$ = macstringz
Loop Until InStr(Message$, "a") = 0
Do
If InStr(Message$, "B") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "B") - 1) + "ß" + Right$(Message$, Len(Message$) - InStr(Message$, "B"))
Message$ = macstringz
Loop Until InStr(Message$, "B") = 0
Do
If InStr(Message$, "b") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "b") - 1) + "ß" + Right$(Message$, Len(Message$) - InStr(Message$, "b"))
Message$ = macstringz
Loop Until InStr(Message$, "b") = 0
Do
If InStr(Message$, "C") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "C") - 1) + "Ç" + Right$(Message$, Len(Message$) - InStr(Message$, "C"))
Message$ = macstringz
Loop Until InStr(Message$, "C") = 0
Do
If InStr(Message$, "c") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "c") - 1) + "ç" + Right$(Message$, Len(Message$) - InStr(Message$, "c"))
Message$ = macstringz
Loop Until InStr(Message$, "c") = 0
Do
If InStr(Message$, "D") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "D") - 1) + "Ð" + Right$(Message$, Len(Message$) - InStr(Message$, "D"))
Message$ = macstringz
Loop Until InStr(Message$, "D") = 0
Do
If InStr(Message$, "d") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "d") - 1) + "Ð" + Right$(Message$, Len(Message$) - InStr(Message$, "d"))
Message$ = macstringz
Loop Until InStr(Message$, "d") = 0
Do
If InStr(Message$, "E") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "E") - 1) + "Ê" + Right$(Message$, Len(Message$) - InStr(Message$, "E"))
Message$ = macstringz
Loop Until InStr(Message$, "E") = 0
Do
If InStr(Message$, "e") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "e") - 1) + "ê" + Right$(Message$, Len(Message$) - InStr(Message$, "e"))
Message$ = macstringz
Loop Until InStr(Message$, "e") = 0
Do
If InStr(Message$, "F") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "F") - 1) + "ƒ" + Right$(Message$, Len(Message$) - InStr(Message$, "F"))
Message$ = macstringz
Loop Until InStr(Message$, "F") = 0
Do
If InStr(Message$, "f") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "f") - 1) + "ƒ" + Right$(Message$, Len(Message$) - InStr(Message$, "f"))
Message$ = macstringz
Loop Until InStr(Message$, "f") = 0
Do
If InStr(Message$, "I") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "I") - 1) + "Ï" + Right$(Message$, Len(Message$) - InStr(Message$, "I"))
Message$ = macstringz
Loop Until InStr(Message$, "I") = 0
Do
If InStr(Message$, "i") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "i") - 1) + "ï" + Right$(Message$, Len(Message$) - InStr(Message$, "i"))
Message$ = macstringz
Loop Until InStr(Message$, "i") = 0
Do
If InStr(Message$, "L") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "L") - 1) + "£" + Right$(Message$, Len(Message$) - InStr(Message$, "L"))
Message$ = macstringz
Loop Until InStr(Message$, "L") = 0
Do
If InStr(Message$, "N") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "N") - 1) + "Ñ" + Right$(Message$, Len(Message$) - InStr(Message$, "N"))
Message$ = macstringz
Loop Until InStr(Message$, "N") = 0
Do
If InStr(Message$, "n") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "n") - 1) + "ñ" + Right$(Message$, Len(Message$) - InStr(Message$, "n"))
Message$ = macstringz
Loop Until InStr(Message$, "n") = 0
Do
If InStr(Message$, "O") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "O") - 1) + "Ø" + Right$(Message$, Len(Message$) - InStr(Message$, "O"))
Message$ = macstringz
Loop Until InStr(Message$, "O") = 0
Do
If InStr(Message$, "o") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "o") - 1) + "Ø" + Right$(Message$, Len(Message$) - InStr(Message$, "o"))
Message$ = macstringz
Loop Until InStr(Message$, "o") = 0
Do
If InStr(Message$, "R") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "R") - 1) + "®" + Right$(Message$, Len(Message$) - InStr(Message$, "R"))
Message$ = macstringz
Loop Until InStr(Message$, "R") = 0
Do
If InStr(Message$, "r") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "r") - 1) + "®" + Right$(Message$, Len(Message$) - InStr(Message$, "r"))
Message$ = macstringz
Loop Until InStr(Message$, "r") = 0
Do
If InStr(Message$, "S") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "S") - 1) + "§" + Right$(Message$, Len(Message$) - InStr(Message$, "S"))
Message$ = macstringz
Loop Until InStr(Message$, "S") = 0
Do
If InStr(Message$, "s") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "s") - 1) + "š" + Right$(Message$, Len(Message$) - InStr(Message$, "s"))
Message$ = macstringz
Loop Until InStr(Message$, "s") = 0
Do
If InStr(Message$, "T") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "T") - 1) + "†" + Right$(Message$, Len(Message$) - InStr(Message$, "T"))
Message$ = macstringz
Loop Until InStr(Message$, "T") = 0
Do
If InStr(Message$, "t") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "t") - 1) + "†" + Right$(Message$, Len(Message$) - InStr(Message$, "t"))
Message$ = macstringz
Loop Until InStr(Message$, "t") = 0
Do
If InStr(Message$, "U") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "U") - 1) + "Ü" + Right$(Message$, Len(Message$) - InStr(Message$, "U"))
Message$ = macstringz
Loop Until InStr(Message$, "U") = 0
Do
If InStr(Message$, "u") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "u") - 1) + "ú" + Right$(Message$, Len(Message$) - InStr(Message$, "u"))
Message$ = macstringz
Loop Until InStr(Message$, "u") = 0
Do
If InStr(Message$, "X") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "X") - 1) + "×" + Right$(Message$, Len(Message$) - InStr(Message$, "X"))
Message$ = macstringz
Loop Until InStr(Message$, "X") = 0
Do
If InStr(Message$, "x") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "x") - 1) + "×" + Right$(Message$, Len(Message$) - InStr(Message$, "x"))
Message$ = macstringz
Loop Until InStr(Message$, "x") = 0
Do
If InStr(Message$, "Y") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "Y") - 1) + "¥" + Right$(Message$, Len(Message$) - InStr(Message$, "Y"))
Message$ = macstringz
Loop Until InStr(Message$, "Y") = 0
Do
If InStr(Message$, "y") = 0 Then Exit Do
macstringz = Left$(Message$, InStr(Message$, "y") - 1) + "ÿ" + Right$(Message$, Len(Message$) - InStr(Message$, "y"))
Message$ = macstringz
Loop Until InStr(Message$, "y") = 0
SendChat Message$
End Sub

Public Sub Stop_Button()
Do
DoEvents:
Loop
End Sub

Public Sub Change_AIM_Buddy_Caption()
BuddyCaption1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
BuddyCaption2% = SendMEssageByString(BuddyCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub

Public Sub Link_Chat()
ChatLink = "<A HREF= " + Chr(34) + Link$ + Chr(34) + ">" + Name$ = "</A>"
End Sub

Public Sub StayOnTop2()
' Example: Call StayOnTop(Form1.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub

Public Sub IdleBot()
'You Will Need:
'1 Textboxes, 2 Command Buttons, And 2 Labels
'Make Label1's Caption "Reason:" And Put It Above The Textbox
'Make Command1's Caption "Start" And Command2's "Stop
'In Command1.Click Put: Call IdleBot
'In Command2.Click Put:
'Do
'DoEvents
'Loop
'Make The Text1's Text = ""
'Make Label2's Caption "How Long:" And Put It Above Text2
'Make Text2.Text's = ""
'******Start Of SUB Code******'
ChatSend ("Idle Bot Activated")
TimeOut (1)
ChatSend ("Reason:")
TimeOut (0.1)
ChatSend ("How Long:")
TimeOut (0.1)
ChatSend ("I'll be back...")
End Sub

Public Sub AdvertiseBot()
'To Use It Put: Call AdvertizeBot
SendChat ("[Program Name]")
TimeOut (0.1) ' The Timesouts Are So You Don't Get Logged Off For Scrolling
SendChat ("Made By [Your Handle]")
TimeOut (0.1)
SendChat ("[Some Advertisement About Your Prog]")
End Sub

Public Sub AFKBot()
'You Will Need:
'2 Textboxes, 2 Command Buttons, And A 1 Timer
'Put Timer's Interval To 5 And In It Put:
ChatSend ("AFK Bot Activated")
ChatSend ("AFK for:" & Text1.Text & " min(z)") 'text1 is how long u r afk
TimeOut (0.1)
ChatSend ("Reason: " & Text2.Text) 'Text2 Is The Reason
TimeOut (Text1.Text)
'In The Command1 Put:
'Timer1.Enabled = True
'In The Command2 Put:
'Timer1.Enabled = False
'Do
'DoEvents
'Loop
End Sub

Public Sub EchoBot()
'You Will Need:
'1 Timer And 2 Command Buttons
'With Timer1's Interval At 5 And Enabled = False Put:
Dim LastChatLine As String
Dim SNLastChatLine As String
LastChatLine$ = LastChatLine 'This Would Be A Sub That Would Get
                          'The Last Text Line In A Chat Room
SNLastChatLine$ = SNFromLastChatLine 'This would Also Be A Sub That Would
                            'Get The SN From The Last Text Line
                            'In A Chat Room
SendChat ("Echo Bot Active")
SendChat ("Echoing: " & SNFromLastChatLine$)
TimeOut (0.3)
Do
SendChat (LastChatLine$)
TimeOut (4)
Loop Until Timer1.Enabled = False
'In Command1 Put: Timer1.Enabled = True
'In Command2 Put: Timer1.Enabled = False
End Sub

Public Sub AttentionBot()
'To Use This Put: Call AttentionBot
ChatSend ("{S IM")
TimeOut (1.2)
ChatSend ("Attention Bot")
ChatSend ("Gimme Attention!")
TimeOut (0.1)
ChatSend ("{S IM")
TimeOut (1.1)
ChatSend ("{S BUDDYIN")
TimeOut (1.1)
ChatSend ("{S GOTMAIL")
TimeOut (1.1)
End Sub

Public Sub EightBallBot()
'You Will Need:
'1 Textbox, 1 Label
'Text1 Is what You Are Asking, And The Label Is Where You Randomize
Label1.Caption = Int(Rnd * 9)
ChatSend ("8Ball Bot Loaded")
TimeOut (0.4)
ChatSend (Text1.Text & " = Question")
    If Label1.Caption = "1" Then
    ChatSend ("Excellent Chance!")
    End If
If Label1.Caption = "2" Then
ChatSend ("Great Chance!")
End If
    If Label1.Caption = "3" Then
    ChatSend ("Good Chance!")
    End If
If Label1.Caption = "4" Then
ChatSend ("Okay Chance!")
End If
    If Label1.Caption = "5" Then
    ChatSend ("Bad Chance!")
    End If
If Label1.Caption = "6" Then
ChatSend ("Very Bad Chance!")
End If
    If Label1.Caption = "7" Then
    SendChat ("0% Chance!")
    End If
If Label1.Caption = "8" Then
ChatSend ("Your Having A Horrible Day!")
End If
End Sub

Public Sub PlayMIDI()
Dim Safefile As String
    Safefile$ = Dir(MIDIFile$)
    If Safefile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub CheckIfAlive()
Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call SendMail("" & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub CheckIfMaster()
Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mStatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    Pause 0.3
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 0.8
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        mStatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = GetText(mStatic&)
    Loop Until Modal& <> 0 And mStatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub PrivateRoom()
Call Keyword("aol://2719:2-2-" & Room$)
End Sub

Public Sub RoomCount()
Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Sub

Public Sub PublicRoom()
Call Keyword("aol://2719:21-2-" & Room$)
End Sub
