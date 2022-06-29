Attribute VB_Name = "SimpleAOL6Beta"
'***************************************'
'*                                     *'
'* File: simple-aol6.bas               *'
'* Coded by: Skew & Jaze               *'
'* URL: http://go.to/GSoftware/        *'
'*                                     *'
'* Notes:                              *'
'* -If anything is taken from this bas *'
'* please give us credit because we    *'
'* worked pretty hard it.              *'
'* -Besides 2 subs taken from          *'
'* sloveaol6.bas, everything else was  *'
'* coded by Skew & Jaze                *'
'*                                     *'
'***************************************'

Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
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
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_DELETEITEM = &H2D
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

Public Const COLON_KEY = 58
Public Const TAB_KEY = 9
Public Const ENTER_KEY = 13 'Taken from Dos's bas
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const New_Mail = 1
Public Const Old_Mail = 2
Public Const Sent_Mail = 3

'Menu Commands
'Figured out by progee & slove
'Taken from sloveaol6.bas

'Mail
Public Const Menu_Mail = 1
Public Const Menu_NewMail = 7170
Public Const Menu_OldMail = 7171
Public Const Menu_SentMail = 7172
Public Const Menu_WriteMail = 7173
Public Const Menu_MailPref = 7185
'People
Public Const Menu_People = 2
Public Const Menu_IM = 7169
Public Const Menu_BuddyList = 7177
Public Const Menu_GetProfile = 7178
Public Const Menu_Locate = 7179
'Settings
Public Const Menu_Settings = 4
Public Const Menu_PControls = 7172
Public Const Menu_Profile = 7173
Public Const Menu_ScreenName = 7174
Public Const Menu_PassWord = 7175

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Function FindAOL() As Long

FindAOL = FindWindow("AOL Frame25", vbNullString)
End Function

Public Function FindMDI() As Long

FindMDI = FindWindowEx(FindAOL, 0&, "MDIClient", vbNullString)
End Function

Public Function FindToolBar() As Long

Dim tool As Long
tool = FindWindowEx(FindAOL, 0&, "AOL Toolbar", vbNullString)
FindToolBar = FindWindowEx(tool, 0&, "_AOL_Toolbar", vbNullString)
End Function

Public Sub KillPlus(SW_Command As Long)

Dim mdi, win As Long
mdi = FindMDI
win = FindWindowEx(mdi, 0&, "AOL Child", " AOL Plus")
Call ShowWindow(win, SW_Command)
End Sub

Public Sub KillWelcome(SW_Command As Long)

Dim mdi, win As Long
mdi = FindMDI
win = FindWindowEx(mdi, 0&, "AOL Child", "Welcome, " & UserSN & "!")
Call ShowWindow(win, SW_Command)
End Sub

Public Sub HideAOL(SW_Command As Long)

Call ShowWindow(FindAOL, SW_Command)
End Sub

Public Function UserSN() As String

On Error Resume Next
Dim mdi, win As Long
Dim cap, sn As String
mdi = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
cap = GetCaption(win)
If InStr(cap, "Welcome, ") Then
   UserSN = Mid(cap, InStr(cap, ",") + 2, InStr(cap, "!") - (InStr(cap, ",") + 2))
   Exit Function
End If
Loop Until win = 0&
UserSN = ""
End Function

Private Function GetCaption(ByVal hWnd As Long) As String
    
Dim buffer As String
Dim txtlen As Long
txtlen = GetWindowTextLength(hWnd)
buffer = String(txtlen, 0&)
Call GetWindowText(hWnd, buffer, txtlen + 1)
GetCaption = buffer
End Function

Private Function GetText(ByVal hWnd As Long) As String

Dim txtlen As Long
Dim buffer As String
txtlen = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
buffer = String(txtlen, 0&)
Call SendMessageByString(hWnd, WM_GETTEXT, txtlen + 1, buffer)
GetText = buffer
End Function

Private Sub ClickButton(ByVal Button As Long)

Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub ClickIcon(ByVal hWnd As Long)
'Written by progee & slove
'Taken from sloveaol6.bas

Dim Size As RECT
Dim CurPos As POINTAPI
Call GetWindowRect(hWnd, Size)
Call GetCursorPos(CurPos)
Call SetCursorPos(Size.Left + (Size.Right - Size.Left) / 2, Size.Top + (Size.Bottom - Size.Top) / 2)
Call SendMessageByNum(hWnd, WM_LBUTTONDOWN, 0, 0)
Call SendMessageByNum(hWnd, WM_LBUTTONUP, 0, 0)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub

Public Sub Keyword(keywrd As String)

Dim tool, Combo, txtbx, icon As Long
tool = FindToolBar
Combo = FindWindowEx(tool, 0&, "_AOL_Combobox", vbNullString)
txtbx = FindWindowEx(Combo, 0&, "Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, keywrd)
Call SendMessageLong(txtbx, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(txtbx, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Function FindChatRoom() As Long

On Error Resume Next
Dim mdi, win, stat, i As Long
mdi = FindMDI
win = 0&
Do
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
stat = 0&
For i = 1 To 4
stat = FindWindowEx(win, stat, "_AOL_Static", vbNullString)
Next i
If GetText(stat) = "people here" Then
   FindChatRoom = win
   Exit Function
End If
Loop Until win = 0&
FindChatRoom = 0&
End Function

Public Sub ClearChat()

Dim txtbx As Long
txtbx = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, "")
End Sub

Public Sub CloseChatRoom()

Dim room As String
room = FindChatRoom
If room <> 0& Then Call PostMessage(room, WM_CLOSE, 0&, 0&)
End Sub

Public Function ChatRoomName() As String

ChatRoomName = GetCaption(FindChatRoom())
End Function

Public Sub SendChat(What_To_Say As String)

Dim chat, txtbx As Long
chat = FindChatRoom
txtbx = FindWindowEx(chat, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(chat, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, What_To_Say)
Call SendMessageLong(txtbx, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function FindIM() As Long

Dim mdi, win, icon As Long
Dim Caption As String
mdi = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
Caption = GetCaption(win)
If InStr(Caption, "IM To") Or InStr(Caption, "IM From") Or InStr(Caption, "Send Instant Message") Then
   If InStr(Caption, "IM From") Then Call ClickRespond(win)
   FindIM = win
   Exit Function
End If
Loop Until win = 0&
End Function

Public Sub InstantMessage(sn As String, msg As String)
'Waits to see if user is online and
'closes msgbox if user is not online

Dim IMwin, txtbx, icon As Long
Call Keyword("im")
Pause (1)
IMwin = FindIM
txtbx = FindWindowEx(IMwin, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, sn)
txtbx = FindWindowEx(IMwin, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(IMwin, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
If WaitForError <> 0& Then Call PostMessage(IMwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IM(sn As String, msg As String)
'Does not wait to see if user is online
'Also waits to make sure the im window closes

Dim IMwin, txtbx, icon As Long
Call Keyword("im")
Pause (1)
IMwin = FindIM
txtbx = FindWindowEx(IMwin, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, sn)
txtbx = FindWindowEx(IMwin, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(IMwin, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Sub IMsOn()

Dim IMwin As Long
Call IM("$IM_ON", "Skew & Jaze")
Pause (1)
Call WaitForError
IMwin = 0&
Do
DoEvents
IMwin = FindIM
Loop Until IMwin <> 0&
Call PostMessage(IMwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IMsOff()

Dim IMwin As Long
Call IM("$IM_OFF", "Skew & Jaze")
Pause (1)
Call WaitForError
IMwin = 0&
Do
DoEvents
IMwin = FindIM
Loop Until IMwin <> 0&
Call PostMessage(IMwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IMRespond(msg As String, Close_Win As Boolean)

Dim IM, txtbx, icon As Long
IM = FindIM
txtbx = FindWindowEx(IM, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(IM, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(IM, txtbx, "_AOL_Icon", "")
Call SendMessageLong(icon, WM_CHAR, ENTER_KEY, 0&)
If Close_Win = True Then Call PostMessage(IM, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ClickRespond(ByVal imHandle As Long)

On Error Resume Next
Dim win, txtbx, icon As Long
txtbx = FindWindowEx(imHandle, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(imHandle, txtbx, "RICHCNTL", vbNullString)
icon = FindWindowEx(imHandle, txtbx, "_AOL_Icon", vbNullString)
icon = FindWindowEx(imHandle, icon, "_AOL_Icon", vbNullString)
Call SendMessageLong(icon, WM_CHAR, ENTER_KEY, 0&)
WaitForOK
End Sub

Public Sub WaitForOK()

On Error Resume Next
Dim msgbx, icon, start As Long
start = Timer
Do
DoEvents
msgbx = FindWindow("_AOL_Modal", "America Online")
If Timer > start + 2 Then Exit Sub
Loop Until msgbx <> 0&
icon = FindWindowEx(msgbx, 0&, "_AOL_Icon", vbNullString)
Call ClickButton(icon)
End Sub

Public Function WaitForError() As Long

On Error Resume Next
Dim start, err, icon As Long
start = Timer
Do
DoEvents
err = FindWindow("#32770", "America Online Error")
If err = 0& Then err = FindWindow("#32770", "America Online")
If Timer > start + 2 Then
WaitForError = 0&
Exit Function
End If
Loop Until err <> 0&
icon = FindWindowEx(err, 0&, "Button", "OK")
Call ClickButton(icon)
WaitForError = icon
End Function

Public Function FindMailBox() As Long

FindMailBox = FindWindowEx(FindMDI, 0&, "AOL Child", UserSN & "'s Online Mailbox")
End Function

Public Function ListBoxToMailString(lstbox As ListBox) As String

Dim lst As String, i As Integer
If lstbox.ListCount = 0 Then Exit Function
For i = 0 To lstbox.ListCount - 2
lst = lst & lstbox.List(i) & ","
Next i
ListBoxToMailString = lst & lstbox.List(lstbox.ListCount - 1)
End Function

Public Sub SendMail(Person As String, cc As String, subject As String, msg As String)

Dim mdi, mail, i, txtbx, icon As Long
Call Keyword("mailto: " & Person)
Pause (1)
mdi = FindMDI
mail = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
txtbx = FindWindowEx(mail, 0&, "_AOL_Edit", vbNullString)
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, cc)
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, subject)
txtbx = FindWindowEx(mail, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(mail, txtbx, "_AOL_Icon", vbNullString)
For i = 1 To 3
DoEvents
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Sub OpenMailBox(which As Integer)
Select Case which
Case 1
Call RunIconMenu(Menu_Mail, Menu_NewMail)
Case 2
Call RunIconMenu(Menu_Mail, Menu_OldMail)
Case 3
Call RunIconMenu(Menu_Mail, Menu_SentMail)
End Select
Pause (0.5)
End Sub

Public Function WaitForMailBox(which As Integer) As Integer
Dim mail, aolTab, aolTab2, aoltree, lstbx, num, i, num2 As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
num = num2 = 0
Do
num = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
Pause (1.5)
num2 = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
Loop Until num = num2
WaitForMailBox = num2
End Function

Public Sub RunIconMenu(mnu As Integer, cmd As Long)
'PostMessage() idea originally by progee & slove in sloveaol6.bas
'Optimized Slightly

Dim tool, icon, i As Long
tool = FindToolBar
icon = FindWindowEx(tool, 0&, "_AOL_Icon", vbNullString)
Select Case mnu
Case 1
DoEvents
Case 2
For i = 1 To 3
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
Case 3
For i = 1 To 6
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
Case 4
For i = 1 To 9
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
Case 5
For i = 1 To 11
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
End Select
Call PostMessage(icon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(icon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
tool = FindWindow("#32768", vbNullString)
Loop Until tool <> 0&
Call PostMessage(FindAOL, 273, cmd, 0) 'Taken from sloveaol6.bas
End Sub

Public Sub AwayMsgOff()
'Assumes away message is on

Dim mdi, win, i, icon, BuddyLst As Long
mdi = FindMDI
BuddyLst = FindWindowEx(mdi, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Pause (0.75)
win = FindWindowEx(mdi, 0&, "AOL Child", "Away Message Off")
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Sub AwayMsgOn()
'Assumes away messsage is "not" already on
'Turns on the default message

Dim mdi, win, i, icon, BuddyLst As Long
mdi = FindMDI
BuddyLst = FindWindowEx(mdi, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Pause (0.75)
win = FindWindowEx(mdi, 0&, "AOL Child", "Away Message")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Sub AddAwayMsg(title As String, msg As String)
'Assumes no away message is curently on

Dim mdi, win, win2, txtbx, i, icon, BuddyLst As Long
mdi = FindMDI
BuddyLst = FindWindowEx(mdi, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Pause (0.75)
win = FindWindowEx(mdi, 0&, "AOL Child", "Away Message")
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Pause (0.75)
win2 = FindWindowEx(mdi, 0&, "AOL Child", "New Away Message")
txtbx = FindWindowEx(win2, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, title)
txtbx = FindWindowEx(win2, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(win2, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Call PostMessage(win, WM_CLOSE, 0&, 0&)
End Sub

Public Function MsgFromIM() As String

Dim txtbx, tabspot, finalspot As Long
Dim msgtxt As String
txtbx = FindWindowEx(FindIM, 0&, "RICHCNTL", vbNullString)
msgtxt = GetText(txtbx)
tabspot = InStr(msgtxt, Chr(COLON_KEY))
Do
finalspot = tabspot
tabspot = InStr(tabspot + 1, msgtxt, Chr(COLON_KEY))
Loop Until tabspot <= 0
msgtxt = Right(msgtxt, Len(msgtxt) - finalspot - 1)
MsgFromIM = Left(msgtxt, Len(msgtxt) - 1)
End Function

Public Function SNFromIM() As String

Dim txtbx As Long
Dim cap As String
cap = GetCaption(FindIM)
If InStr(cap, Chr(COLON_KEY)) <= 0 Then
SNFromIM = ""
Exit Function
Else
SNFromIM = Right(cap, Len(cap) - InStr(cap, Chr(COLON_KEY)) - 1)
End If
End Function

Public Function SNFromChatline(Line As String) As String

If InStr(Line, ":") = 0 Then
SNFromChatline = ""
Exit Function
End If
SNFromChatline = Left(Line, InStr(Line, ":") - 1)
End Function

Public Function MsgFromChatLine(Line As String) As String

If InStr(Line, Chr(TAB_KEY)) = 0 Then
MsgFromChatLine = ""
Exit Function
End If
MsgFromChatLine = Right(Line, Len(Line) - InStr(Line, Chr(TAB_KEY)))
End Function

Public Function LastChatLine() As String

Dim chattext, thechar, thechars, TheChatText, lastline As String
Dim findchat, lastlen, findchar As Long

DoEvents
chattext = GetChatText
For findchar = 1 To Len(chattext)
   DoEvents
   thechar = Mid(chattext, findchar, 1)
   thechars = thechars & thechar
   If thechar = Chr(ENTER_KEY) Then
      DoEvents
      TheChatText = Mid(thechars, 1, Len(thechars) - 1)
      thechars = ""
   End If
   DoEvents
Next findchar
lastlen = Val(findchar) - Len(thechars)
lastline = Mid(chattext, lastlen, Len(thechars))
DoEvents
LastChatLine = lastline
End Function

Public Function GetChatText() As String

Dim chat, txtbx As Long
chat = FindChatRoom
txtbx = FindWindowEx(chat, 0&, "RICHCNTL", vbNullString)
GetChatText = GetText(txtbx)
End Function

Public Function IsUserOnline() As Boolean

If UserSN = "" Then
   IsUserOnline = False
Else
   IsUserOnline = True
End If
End Function

Public Function CountMail(which As Integer) As Integer

Call OpenMailBox(which)
CountMail = WaitForMailBox(which)
End Function

Public Sub DeleteMailByIndex(which As Integer, idx As Long)

Dim count As Integer
Dim icon, mail, aolTab, aolTab2, i, aoltree As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If count = 0 Or idx > count - 1 Then Exit Sub
Call SendMessage(aoltree, LB_SETCURSEL, idx, 0&)
icon = 0&
For i = 1 To 7
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Sub MailToList(which As Integer, lstbx As ListBox)

Dim mail, aolTab, aolTab2, i, aoltree, count, maillist, length, pos, pos2
Dim mailer As String
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If count = 0 Then Exit Sub
For maillist = 0 To count - 1
DoEvents
length = SendMessage(aoltree, LB_GETTEXTLEN, maillist, 0&)
mailer = String(length + 1, 0)
Call SendMessageByString(aoltree, LB_GETTEXT, maillist, mailer)
pos = InStr(mailer, Chr(TAB_KEY))
pos = InStr(pos + 1, mailer, Chr(TAB_KEY))
mailer = Right(mailer, Len(mailer) - pos)
lstbx.AddItem mailer
Next maillist
End Sub

Public Function SenderFromMail(which As Integer, idx As Long) As String

Dim mail, aolTab, aolTab2, i, aoltree, count, maillist, length, pos, pos2
Dim mailer As String
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If count = 0 Or idx > count - 1 Then Exit Function
length = SendMessage(aoltree, LB_GETTEXTLEN, idx, 0&)
mailer = String(length + 1, 0)
Call SendMessageByString(aoltree, LB_GETTEXT, idx, mailer)
pos = InStr(mailer, Chr(TAB_KEY))
pos2 = InStr(pos + 1, mailer, Chr(TAB_KEY))
mailer = Mid(mailer, pos + 1, pos2 - pos - 1)
SenderFromMail = mailer
End Function

Public Function FindForward() As Long

Dim mdi, win As Long
Dim cap As String
mdi = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
cap = GetCaption(win)
If Left(cap, 4) = "Fwd:" Then
FindForward = win
Exit Function
End If
Loop Until win = 0&
End Function

Public Sub OpenEmailByIndex(which As Integer, idx As Long)

Dim count As Integer
Dim icon, mail, aolTab, aolTab2, i, aoltree As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If count = 0 Or idx > count - 1 Then Exit Sub
Call SendMessage(aoltree, LB_SETCURSEL, idx, 0&)
icon = FindWindowEx(mail, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Function FindEmail() As Long

Dim mdi, win, stat, i As Long
mdi = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
stat = 0&
For i = 1 To 3
DoEvents
stat = FindWindowEx(win, stat, "_AOL_Static", vbNullString)
Next i
If GetCaption(stat) = "Reply" Then
FindEmail = win
Exit Function
End If
Loop Until win = 0&
FindEmail = 0&
End Function

Public Sub ForwardEmail(Person As String, cc As String, msg As String, removefwd As Boolean, closemail As Boolean)

On Error Resume Next
Dim icon, mail, txtbx, win, i As Long
Dim subj, cap As String
mail = FindEmail
icon = 0&
For i = 1 To 8
DoEvents
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Do
DoEvents
win = FindForward
Loop Until win <> 0&
txtbx = FindWindowEx(win, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Person)
txtbx = FindWindowEx(win, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, cc)
txtbx = FindWindowEx(win, txtbx, "_AOL_Edit", vbNullString)
If removefwd Then
cap = GetText(txtbx)
subj = Right(cap, Len(cap) - 5)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, subj)
End If
txtbx = FindWindowEx(win, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg)
icon = FindWindowEx(win, txtbx, "_AOL_Icon", vbNullString)
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Call ClickModal
If closemail Then
Call PostMessage(mail, WM_CLOSE, 0&, 0&)
End If
endit:
End Sub

Public Sub ClickModal()
Dim icon, win As Long
Do
DoEvents
win = FindWindow("_AOL_Modal", vbNullString)
Loop Until win <> 0&
Do
DoEvents
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Loop Until icon <> 0&
Call ClickIcon(icon)
End Sub

Public Function GetValidChatLine() As String

Static lastchat As String
Dim length As String
Do
DoEvents
length = GetChatText
Loop Until length <> lastchat
lastchat = length
GetValidChatLine = LastChatLine
End Function
