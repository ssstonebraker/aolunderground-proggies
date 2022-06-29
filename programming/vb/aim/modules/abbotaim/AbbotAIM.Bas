Attribute VB_Name = "AbbotAIM"

'        __       ____   ____     ____  ______
'     / /  \    /|    \/|    \ / /    \|__  __|
'    / /    \  | |  o / |  o / | |     |'|| |_/
'   / /  O   \ | |    | |    | | |  O  |'|| |
'  / /   __   \| |  o  \|  o  \| |     |'|| |
' / /___/ \ \__\ |_____/|_____/| \_____/'||_|
'/____ /   \___//____ //_____/ \_____ / '|__/
'AbbotAIM v1 for use with AIM v 2.0.741
'
'Sup everybody?????
'So I guess you're gonna make an AIM program
'have fun!!!!
'Programming for AIM is very easy!
'
'If you use this all I want is that you give
'me some credit in your prog
'If you have any questions you can mail me:
'          xThaAbbotx@hotmail.com
'GREETZ & SHOUT OUTZ:
'prophecy
'Soulsnyper
'CldPhuSiON
'MindSt0rmz
'MenaceNYC
'IIITRUlll
'NoLiSoJa1
'AXION884
'PunkinMHS
'KNK's phat ass programming page
'and anyone else I may have forgot
'
'Enjoy the .bas


Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type







Function Attention(message$)
'Gets peoples attention
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
SendChat message$
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
End Function
Sub playwav(Wav$)
X = sndPlaySound("wav$", 1):
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
Call SendMessageByString(Box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Go&)
End Function

Function HideAIM()
'Hides the AIM window
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = ShowWindow(aim&, SW_HIDE)
End Function

Function ShowAIM()
'Shows the AIM window
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = ShowWindow(aim&, SW_SHOW)
End Function


Sub StayOnTop(the As Form)
'keeps the form on top of everything
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function ChangeAIMCaption(Caption$)
'changes AIM's caption
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Function EliteTalker(message$)
'sends your message using ascii characters
On Error Resume Next
Do
If InStr(message$, "A") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "A") - 1) + "Ã" + Right$(message$, Len(message$) - InStr(message$, "A"))
message$ = macstringz
Loop Until InStr(message$, "A") = 0
Do
If InStr(message$, "a") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "a") - 1) + "å" + Right$(message$, Len(message$) - InStr(message$, "a"))
message$ = macstringz
Loop Until InStr(message$, "a") = 0
Do
If InStr(message$, "B") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "B") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "B"))
message$ = macstringz
Loop Until InStr(message$, "B") = 0
Do
If InStr(message$, "b") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "b") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "b"))
message$ = macstringz
Loop Until InStr(message$, "b") = 0
Do
If InStr(message$, "C") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "C") - 1) + "Ç" + Right$(message$, Len(message$) - InStr(message$, "C"))
message$ = macstringz
Loop Until InStr(message$, "C") = 0
Do
If InStr(message$, "c") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "c") - 1) + "ç" + Right$(message$, Len(message$) - InStr(message$, "c"))
message$ = macstringz
Loop Until InStr(message$, "c") = 0
Do
If InStr(message$, "D") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "D") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "D"))
message$ = macstringz
Loop Until InStr(message$, "D") = 0
Do
If InStr(message$, "d") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "d") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "d"))
message$ = macstringz
Loop Until InStr(message$, "d") = 0
Do
If InStr(message$, "E") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "E") - 1) + "Ê" + Right$(message$, Len(message$) - InStr(message$, "E"))
message$ = macstringz
Loop Until InStr(message$, "E") = 0
Do
If InStr(message$, "e") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "e") - 1) + "ê" + Right$(message$, Len(message$) - InStr(message$, "e"))
message$ = macstringz
Loop Until InStr(message$, "e") = 0
Do
If InStr(message$, "F") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "F") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "F"))
message$ = macstringz
Loop Until InStr(message$, "F") = 0
Do
If InStr(message$, "f") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "f") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "f"))
message$ = macstringz
Loop Until InStr(message$, "f") = 0
Do
If InStr(message$, "I") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "I") - 1) + "Ï" + Right$(message$, Len(message$) - InStr(message$, "I"))
message$ = macstringz
Loop Until InStr(message$, "I") = 0
Do
If InStr(message$, "i") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "i") - 1) + "ï" + Right$(message$, Len(message$) - InStr(message$, "i"))
message$ = macstringz
Loop Until InStr(message$, "i") = 0
Do
If InStr(message$, "L") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "L") - 1) + "£" + Right$(message$, Len(message$) - InStr(message$, "L"))
message$ = macstringz
Loop Until InStr(message$, "L") = 0
Do
If InStr(message$, "N") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "N") - 1) + "Ñ" + Right$(message$, Len(message$) - InStr(message$, "N"))
message$ = macstringz
Loop Until InStr(message$, "N") = 0
Do
If InStr(message$, "n") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "n") - 1) + "ñ" + Right$(message$, Len(message$) - InStr(message$, "n"))
message$ = macstringz
Loop Until InStr(message$, "n") = 0
Do
If InStr(message$, "O") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "O") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "O"))
message$ = macstringz
Loop Until InStr(message$, "O") = 0
Do
If InStr(message$, "o") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "o") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "o"))
message$ = macstringz
Loop Until InStr(message$, "o") = 0
Do
If InStr(message$, "R") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "R") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "R"))
message$ = macstringz
Loop Until InStr(message$, "R") = 0
Do
If InStr(message$, "r") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "r") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "r"))
message$ = macstringz
Loop Until InStr(message$, "r") = 0
Do
If InStr(message$, "S") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "S") - 1) + "§" + Right$(message$, Len(message$) - InStr(message$, "S"))
message$ = macstringz
Loop Until InStr(message$, "S") = 0
Do
If InStr(message$, "s") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "s") - 1) + "š" + Right$(message$, Len(message$) - InStr(message$, "s"))
message$ = macstringz
Loop Until InStr(message$, "s") = 0
Do
If InStr(message$, "T") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "T") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "T"))
message$ = macstringz
Loop Until InStr(message$, "T") = 0
Do
If InStr(message$, "t") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "t") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "t"))
message$ = macstringz
Loop Until InStr(message$, "t") = 0
Do
If InStr(message$, "U") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "U") - 1) + "Ü" + Right$(message$, Len(message$) - InStr(message$, "U"))
message$ = macstringz
Loop Until InStr(message$, "U") = 0
Do
If InStr(message$, "u") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "u") - 1) + "ú" + Right$(message$, Len(message$) - InStr(message$, "u"))
message$ = macstringz
Loop Until InStr(message$, "u") = 0
Do
If InStr(message$, "X") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "X") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "X"))
message$ = macstringz
Loop Until InStr(message$, "X") = 0
Do
If InStr(message$, "x") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "x") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "x"))
message$ = macstringz
Loop Until InStr(message$, "x") = 0
Do
If InStr(message$, "Y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "Y") - 1) + "¥" + Right$(message$, Len(message$) - InStr(message$, "Y"))
message$ = macstringz
Loop Until InStr(message$, "Y") = 0
Do
If InStr(message$, "y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "y") - 1) + "ÿ" + Right$(message$, Len(message$) - InStr(message$, "y"))
message$ = macstringz
Loop Until InStr(message$, "y") = 0
SendChat message$
End Function




Function SNfromIM()
'gets the SN of the sender of an IM
'won't work if you've changed the IM caption!
Dim IM As Long
On Error Resume Next
IM& = FindWindow("AIM_IMessage", vbNullString)
Name$ = GetCaption(IM&)
If InStr(Name$, "- Instant Message") <> 0 Then
Text% = GetWindowTextLength(IM&)
Text% = (Text%) - 19
SN$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
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
'changes the caption of an IM
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Function ChangeChatCaption(Caption$)
'changes the chatroom's caption
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_ChatWnd", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function


Function exitaim()
'exits AIM
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "E&xit")
End Function

Function signoffAIM()
'signs off AIM
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Sign O&ff")
End Function


Function KillWindow(Window%)
KillWindow = SendMessageByNum(Window%, WM_CLOSE, 0, 0)
End Function

Function MacroKill()
'Messes up someones macro
SendChat "~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~"
TimeOut (0.25)
SendChat "~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~ ~@%@%@%Abbot%@%@%@~"
End Function

Function SendInvite(Who$, message$, Chat$)
'Sends an chat invitation to someone
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
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, Chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function

Function SendLink(Link$, message$)
'Sends a link in a chat
SendChat "<a href=""" + Link$ + """>" + message$ + ""
End Function

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Function clearChat()
'clears the chat text
Dim Chat As Long, Box As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, "")

End Function
Function GetChatName()
'gets the name of the chat room
'This won't work if you've changed the chat caption!
Dim Chat As Long
On Error Resume Next
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = GetCaption(Chat&)
Room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
GetChatName = Room$
End Function
Function IsAIMOnline()
'sees if AIM is online
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
IsAIMOnline = aim&
End Function

Function FindChatRoom()
'looks for a chatroom
Dim Chat As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
FindChatRoom = Chat&
End Function
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function Scroller(message$)
'Scrolls 10 lines of your message
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
TimeOut (0.25)
SendChat message$
End Function

Sub HideAIMAd()
'Hides that annoying advertisement
Dim part1 As Long, part2 As Long, HideIt As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
HideIt& = ShowWindow(part2&, SW_HIDE)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
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
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function GetchatText()
'gets the text from a chat
'sorry, but it gets the text with full HTML
'so it's kinda useless, but I put in here anyway
Dim Chat As Long, ChatWin As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
ChatWin& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
it = GetText(ChatWin&)
GetchatText = it
End Function


Function MassIM(List As ListBox, message As String)
'Mass IM the people on your list
'You must have a listbox with screen names
'and a message to send!
If List.ListCount = 0 Then
Do: DoEvents: Loop
End If
List.Enabled = False
i = List.ListCount - 1
List.ListIndex = 0
For X = 0 To i
List.ListIndex = X
Call SendIM(List.Text, message)
TimeOut (0.8)
Next X
List.Enabled = True
End Function

Function openIM()
'Opens an IM
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
End Function

Function SayRoomName()
'Says the room name
'Won't work if you've changed the chat caption!
SendChat "You have just entered room """ + GetChatName + "."""
End Function

Function SendChat(message$)
'Sends you're message to a chat
Dim Chat As Long, Box As Long, Box2 As Long, Box3 As Long, send As Long, send2 As Long, send3 As Long, send4 As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(Chat&, Box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(Chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Chat&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(Chat&, send3&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send4&)
End Function
Function SendIM(Who$, message$)
'Sends an IM to someone
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
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send&)
End Function
Function ShowAIMAd()
'Shows that gay advertisement
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = ShowWindow(part2&, SW_SHOW)
End Function





Function UserSN()
'Get's the User's SN
'Won't work if you've changed the AIM caption!!!
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Name$ = GetCaption(aim&)
If InStr(Name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 14
SN$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
UserSN = SN$
Else
UserSN = "(Unknown)"
End If
End Function

