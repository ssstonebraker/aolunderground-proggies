Attribute VB_Name = "SimpleAOL6"
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
'* -when dealing with any of the mail  *'
'* functions, it will usually require  *'
'* "which" as one of the parameters.   *'
'* this refers to which mail box you   *'
'* want to use.                        *'
'* 1 for new, 2 for old, 3 for sent    *'
'*                                     *'
'***************************************'

Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
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
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

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
Public Const Menu_WaitingToBeSent = 7180
Public Const Menu_MailPref = 7185
'People
Public Const Menu_People = 2
Public Const Menu_IM = 7169
Public Const Menu_BuddyList = 7177
Public Const Menu_GetProfile = 7178
Public Const Menu_Locate = 7179
'Settings
Public Const Menu_Settings = 4
Public Const Menu_Pref = 7171
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

'Closes and shows the AOL PLus window which appears on broadband connections
Dim MDI, win As Long
MDI = FindMDI
win = FindWindowEx(MDI, 0&, "AOL Child", " AOL Plus")
Call ShowWindow(win, SW_Command)
End Sub

Public Sub KillWelcome(SW_Command As Long)

Dim MDI, win As Long
MDI = FindMDI
win = FindWindowEx(MDI, 0&, "AOL Child", "Welcome, " & UserSN & "!")
Call ShowWindow(win, SW_Command)
End Sub

Public Sub HideAOL(SW_Command As Long)

Call ShowWindow(FindAOL, SW_Command)
End Sub

Public Function UserSN() As String

On Error Resume Next
Dim MDI, win As Long
Dim cap, sn As String
MDI = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(MDI, win, "AOL Child", vbNullString)
cap = GetCaption(win)
If InStr(cap, "Welcome, ") Then
   UserSN = Mid(cap, InStr(cap, ",") + 2, InStr(cap, "!") - (InStr(cap, ",") + 2))
   Exit Function
End If
Loop Until win = 0&
UserSN = ""
End Function

Private Function GetCaption(ByVal hwnd As Long) As String
    
Dim buffer As String
Dim txtlen As Long
txtlen = GetWindowTextLength(hwnd)
buffer = String(txtlen, 0&)
Call GetWindowText(hwnd, buffer, txtlen + 1)
GetCaption = buffer
End Function

Public Function GetText(ByVal hwnd As Long) As String

Dim txtlen As Long
Dim buffer As String
txtlen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
buffer = String(txtlen, 0&)
Call SendMessageByString(hwnd, WM_GETTEXT, txtlen + 1, buffer)
GetText = buffer
End Function

Public Sub ClickButton(ByVal Button As Long)

Call SendMessage(Button, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(Button, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub ClickIcon(ByVal hwnd As Long)
'Written by progee & slove
'Taken from sloveaol6.bas

Dim Size As RECT
Dim CurPos As POINTAPI
Call GetWindowRect(hwnd, Size)
Call GetCursorPos(CurPos)
Call SetCursorPos(Size.Left + (Size.Right - Size.Left) / 2, Size.Top + (Size.Bottom - Size.Top) / 2)
Call SendMessageByNum(hwnd, WM_LBUTTONDOWN, 0, 0)
Call SendMessageByNum(hwnd, WM_LBUTTONUP, 0, 0)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub

Public Sub KeyWord(keywrd As String)

Dim tool, Combo, txtbx, icon As Long
tool = FindToolBar
Combo = FindWindowEx(tool, 0&, "_AOL_Combobox", vbNullString)
txtbx = FindWindowEx(Combo, 0&, "Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, keywrd)
Call SendMessageLong(txtbx, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(txtbx, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub SendChat(What_To_Say As String)
Dim Chat, txtbx As Long
Dim Msg As String
Chat = FindChatRoom
txtbx = FindWindowEx(Chat, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(Chat, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, What_To_Say)
Call SendMessageLong(txtbx, WM_CHAR, ENTER_KEY, 0&)
Do
DoEvents
Msg = GetText(txtbx)
Loop Until Msg = ""
End Sub

Public Function FindChatRoom() As Long

On Error Resume Next
Dim MDI, win, stat, i As Long
MDI = FindMDI
win = 0&
Do
win = FindWindowEx(MDI, win, "AOL Child", vbNullString)
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

Public Function FindIM() As Long

Dim MDI, win, icon As Long
Dim Caption As String
MDI = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(MDI, win, "AOL Child", vbNullString)
Caption = GetCaption(win)
If InStr(Caption, "IM To") Or InStr(Caption, "IM From") Or InStr(Caption, "Send Instant Message") Then
   If InStr(Caption, "IM From") Then Call ClickRespond(win)
   FindIM = win
   Exit Function
End If
Loop Until win = 0&
End Function

Public Sub IMRespond(Msg As String, Close_Win As Boolean)

Dim IM, txtbx, icon As Long
IM = FindIM
txtbx = FindWindowEx(IM, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(IM, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(IM, txtbx, "_AOL_Icon", "")
Call SendMessageLong(icon, WM_CHAR, ENTER_KEY, 0&)
If Close_Win Then Call PostMessage(IM, WM_CLOSE, 0&, 0&)
End Sub

Public Sub InstantMessage(sn As String, Msg As String)
'Waits to see if user is online and
'closes msgbox if user is not online

Dim imwin, txtbx, icon As Long
Call KeyWord("im")
imwin = 0&
Do
imwin = FindIM
Loop Until imwin <> 0&
Pause (0.5)
txtbx = FindWindowEx(imwin, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, sn)
txtbx = FindWindowEx(imwin, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(imwin, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
If WaitForError <> 0& Then Call PostMessage(imwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IM(sn As String, Msg As String)
'Does not wait to see if user is online

Dim imwin, txtbx, icon As Long
Call KeyWord("im")
imwin = 0&
Do
imwin = FindIM
Loop Until imwin <> 0&
Pause (0.5)
txtbx = FindWindowEx(imwin, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, sn)
txtbx = FindWindowEx(imwin, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(imwin, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Sub IMsOn()

Dim imwin As Long
Call IM("$IM_ON", "Skew & Jaze")
Pause (1)
Call WaitForError
imwin = 0&
Do
DoEvents
imwin = FindIM
Loop Until imwin <> 0&
Call PostMessage(imwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub IMsOff()

Dim imwin As Long
Call IM("$IM_OFF", "Skew & Jaze")
Pause (1)
Call WaitForError
imwin = 0&
Do
DoEvents
imwin = FindIM
Loop Until imwin <> 0&
Call PostMessage(imwin, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ClickRespond(ByVal IMHandle As Long)
'Clicks on the respond button when somebody IM's you

On Error Resume Next
Dim win, txtbx, icon As Long
txtbx = FindWindowEx(IMHandle, 0&, "RICHCNTL", vbNullString)
txtbx = FindWindowEx(IMHandle, txtbx, "RICHCNTL", vbNullString)
icon = FindWindowEx(IMHandle, txtbx, "_AOL_Icon", vbNullString)
icon = FindWindowEx(IMHandle, icon, "_AOL_Icon", vbNullString)
Call SendMessageLong(icon, WM_CHAR, ENTER_KEY, 0&)
Call WaitForOK
End Sub

Public Sub WaitForOK()

On Error Resume Next
Dim msgbx, icon, start As Long
start = Timer
Do
DoEvents
msgbx = FindWindow("_AOL_Modal", "America Online")
If Timer > start + 3 Then Exit Sub
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

Public Function ListToMailString(lstbox As ListBox) As String

Dim lst As String, i As Integer
If lstbox.ListCount = 0 Then Exit Function
For i = 0 To lstbox.ListCount - 2
lst = lst & lstbox.List(i) & ","
Next i
ListToMailString = lst & lstbox.List(lstbox.ListCount - 1)
End Function

Public Sub SendEmail(Person As String, cc As String, subject As String, Msg As String)

Dim MDI, mail, i, txtbx, icon, tool As Long
tool = FindToolBar
icon = 0&
For i = 1 To 3
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Do
mail = FindNewMail
Loop Until mail <> 0&
Pause (0.5)
txtbx = FindWindowEx(mail, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Person)
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, cc)
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, subject)
txtbx = FindWindowEx(mail, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(mail, txtbx, "_AOL_Icon", vbNullString)
For i = 1 To 3
DoEvents
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Function FindCommonDialog() As Long

Dim win As Long
win = 0&
Do
DoEvents
win = FindWindow("#32770", "Attach")
Loop Until win <> 0&
FindCommonDialog = win
End Function

Public Function FindAttach() As Long

Dim win As Long
win = 0&
Do
win = FindWindow("_AOL_Modal", "Attachments")
Loop Until win <> 0&
FindAttach = win
End Function

Public Sub SendWaiting()

'Assumes there is already mail waiting to be sent
Dim win, icon, i As Long
Call RunIconMenu(Menu_Mail, Menu_WaitingToBeSent)
win = 0&
Do
DoEvents
win = FindWindowEx(FindMDI, 0&, "AOL Child", "Mail Waiting to be Sent")
Loop Until win <> 0&
Pause (0.5)
icon = 0&
For i = 1 To 4
DoEvents
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Pause (0.5)
Call PostMessage(win, WM_CLOSE, 0&, 0&)
End Sub

Public Sub EmailAttachSendLater(ByVal Person As String, subject As String, Msg As String, File As String)
Dim mail, win, win2, i, txtbx, icon, staticbox, tool As Long
tool = FindToolBar
icon = 0&
For i = 1 To 3
icon = FindWindowEx(tool, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
mail = FindNewMail
Pause (0.75)
txtbx = FindWindowEx(mail, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Person)
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, "")
txtbx = FindWindowEx(mail, txtbx, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, subject)
txtbx = FindWindowEx(mail, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(mail, txtbx, "_AOL_Icon", vbNullString)
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
win = FindAttach
Pause (0.75)
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
win2 = FindCommonDialog
Pause (0.75)
txtbx = FindWindowEx(win2, 0&, "Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, File)
icon = FindWindowEx(win2, txtbx, "Button", "&Open")
Call ClickButton(icon)
Do
DoEvents
Loop Until FindWindow("#32770", "Attach") = 0&
icon = 0&
For i = 1 To 3
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Do
DoEvents
Loop Until FindWindow("_AOL_Modal", "Attachments") = 0&
staticbox = FindWindowEx(mail, 0&, "_AOL_Static", "Send Now")
icon = FindWindowEx(mail, staticbox, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Do
DoEvents
Loop Until FindWindowEx(FindMDI, 0&, "AOL Child", "Write Mail") = 0&
Pause (0.25)
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
mail = 0&
Do
DoEvents
mail = FindMailBox
Loop Until mail <> 0&
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
num = num2 = 0
Do
num = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
Pause (2.5)
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
DoEvents
Call PostMessage(icon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon, WM_LBUTTONUP, 0&, 0&)
Do
tool = FindWindow("#32768", vbNullString)
Loop Until tool <> 0&
Call PostMessage(FindAOL, 273, cmd, 0) 'Taken from sloveaol6.bas
End Sub

Public Sub AwayMsgOff()
'Assumes away message is on

Dim MDI, win, i, icon, BuddyLst As Long
MDI = FindMDI
BuddyLst = FindWindowEx(MDI, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Do
win = FindWindowEx(MDI, 0&, "AOL Child", "Away Message Off")
Loop Until win <> 0&
Pause (0.75)
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Sub AwayMsgOn()
'Assumes away messsage is "not" already on
'Turns on the default message

Dim MDI, win, i, icon, BuddyLst As Long
MDI = FindMDI
BuddyLst = FindWindowEx(MDI, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
Do
win = FindWindowEx(MDI, 0&, "AOL Child", "Away Message")
Loop Until win <> 0&
Pause (0.5)
icon = 0&
For i = 1 To 4
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Sub AddAwayMsg(Title As String, Msg As String)
'Assumes no away message is curently on

Dim MDI, win, win2, txtbx, i, icon, BuddyLst As Long
MDI = FindMDI
BuddyLst = FindWindowEx(MDI, 0&, "AOL Child", "Buddy List")
icon = 0&
For i = 1 To 4
icon = FindWindowEx(BuddyLst, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
win = 0&
Do
win = FindWindowEx(MDI, 0&, "AOL Child", "Away Message")
Loop Until win <> 0&
Pause (0.75)
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
win2 = 0&
Do
DoEvents
win2 = FindWindowEx(MDI, 0&, "AOL Child", "New Away Message")
Loop Until win2 <> 0&
Pause (0.75)
txtbx = FindWindowEx(win2, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Title)
txtbx = FindWindowEx(win2, txtbx, "RICHCNTL", vbNullString)
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, Msg)
icon = FindWindowEx(win2, txtbx, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Call PostMessage(win, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetIMText()

Dim txtbx, imwin As Long
imwin = FindIM
txtbx = FindWindowEx(imwin, 0&, "RICHCNTL", vbNullString)
GetIMText = GetText(txtbx)
End Function

Public Function GetIMText2(IMHandle As Long)

Dim txtbx As Long
txtbx = FindWindowEx(IMHandle, 0&, "RICHCNTL", vbNullString)
GetIMText2 = GetText(txtbx)
End Function


Public Function MsgFromIM2(IMHandle As Long) As String

Dim txtbx, tabspot, finalspot As Long
Dim msgtxt As String
txtbx = FindWindowEx(IMHandle, 0&, "RICHCNTL", vbNullString)
msgtxt = GetIMText2(IMHandle)
tabspot = InStr(msgtxt, Chr(COLON_KEY))
Do
finalspot = tabspot
tabspot = InStr(tabspot + 1, msgtxt, Chr(COLON_KEY))
Loop Until tabspot <= 0
msgtxt = Right(msgtxt, Len(msgtxt) - finalspot - 1)
MsgFromIM2 = Left(msgtxt, Len(msgtxt) - 1)
End Function

Public Function SNFromIM2(IMHandle) As String

Dim txtbx As Long
Dim cap As String
cap = GetCaption(IMHandle)
If InStr(cap, Chr(COLON_KEY)) <= 0 Then
SNFromIM2 = ""
Exit Function
Else
SNFromIM2 = Right(cap, Len(cap) - InStr(cap, Chr(COLON_KEY)) - 1)
End If
End Function

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

If InStr(Line, Chr(COLON_KEY)) = 0 Then
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

Dim chattext, TheChar, TheChars, TheChatText, lastline As String
Dim findchat, lastlen, findchar As Long

DoEvents
chattext = GetChatText
For findchar = 1 To Len(chattext)
   DoEvents
   TheChar = Mid(chattext, findchar, 1)
   TheChars = TheChars & TheChar
   If TheChar = Chr(ENTER_KEY) Then
      DoEvents
      TheChatText = Mid(TheChars, 1, Len(TheChars) - 1)
      TheChars = ""
   End If
   DoEvents
Next findchar
lastlen = Val(findchar) - Len(TheChars)
lastline = Mid(chattext, lastlen, Len(TheChars))
DoEvents
LastChatLine = lastline
End Function

Public Function GetChatText() As String

Dim Chat, txtbx As Long
Chat = FindChatRoom
txtbx = FindWindowEx(Chat, 0&, "RICHCNTL", vbNullString)
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
'which refers to which mailbox to delete from

Dim Count As Integer
Dim icon, mail, aolTab, aolTab2, i, aoltree As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
Count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If Count = 0 Or idx > Count - 1 Then Exit Sub
Call SendMessage(aoltree, LB_SETCURSEL, idx, 0&)
icon = 0&
For i = 1 To 7
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Sub MailToListBox(which As Integer, lstbx As ListBox)

Dim mail, aolTab, aolTab2, i, aoltree, Count, maillist, length, pos, pos2
Dim mailer As String
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
Count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If Count = 0 Then Exit Sub
For maillist = 0 To Count - 1
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

Public Function SenderFromMail(which As Integer, idx As Integer) As String

Dim mail, aolTab, aolTab2, i, aoltree, Count, maillist, length, pos, pos2
Dim mailer As String
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
Count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If Count = 0 Or idx > Count - 1 Then Exit Function
length = SendMessage(aoltree, LB_GETTEXTLEN, idx, 0&)
mailer = String(length + 1, 0)
Call SendMessageByString(aoltree, LB_GETTEXT, idx, mailer)
pos = InStr(mailer, Chr(TAB_KEY))
pos2 = InStr(pos + 1, mailer, Chr(TAB_KEY))
mailer = Mid(mailer, pos + 1, pos2 - pos - 1)
SenderFromMail = mailer
End Function

Public Function FindForward() As Long

Dim MDI, win As Long
Dim cap As String
MDI = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(MDI, win, "AOL Child", vbNullString)
cap = GetCaption(win)
If Left(cap, 4) = "Fwd:" Then
FindForward = win
Exit Function
End If
Loop Until win = 0&
End Function

Public Sub OpenEmailByIndex(which As Integer, idx As Integer)

Dim Count As Integer
Dim icon, mail, aolTab, aolTab2, i, aoltree As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
Count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If Count = 0 Or idx > Count - 1 Then Exit Sub
Call SendMessage(aoltree, LB_SETCURSEL, idx, 0&)
icon = FindWindowEx(mail, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
End Sub

Public Sub KeepMailAsNew(which As Integer, idx As Integer)

Dim Count As Integer
Dim icon, mail, aolTab, aolTab2, i, aoltree As Long
mail = FindMailBox
aolTab = FindWindowEx(mail, 0&, "_AOL_TabControl", vbNullString)
aolTab2 = 0&
For i = 1 To which
aolTab2 = FindWindowEx(aolTab, aolTab2, "_AOL_TabPage", vbNullString)
Next i
aoltree = FindWindowEx(aolTab2, 0&, "_AOL_Tree", vbNullString)
Count = SendMessage(aoltree, LB_GETCOUNT, 0&, 0&)
If Count = 0 Or idx > Count - 1 Then Exit Sub
Call SendMessage(aoltree, LB_SETCURSEL, idx, 0&)
icon = 0&
For i = 1 To 3
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
End Sub

Public Function FindNewMail() As Long
Dim mail As Long
mail = 0&
Do
DoEvents
mail = FindWindowEx(FindMDI, 0&, "AOL Child", "Write Mail")
Loop Until mail <> 0&
FindNewMail = mail
End Function

Public Function FindEmail() As Long

Dim MDI, win, stat, i As Long
MDI = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(MDI, win, "AOL Child", vbNullString)
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

Public Sub ForwardEmail(Person As String, cc As String, Msg As String, removefwd As Boolean, closemail As Boolean)

On Error Resume Next
Dim icon, mail, txtbx, win, i As Long
Dim subj, cap, msg2 As String
Do
mail = FindEmail
Loop Until mail <> 0&
icon = 0&
For i = 1 To 8
DoEvents
icon = FindWindowEx(mail, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
win = 0&
Call WaitForFwdToOpen
win = FindForward
Pause (0.5)
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
msg2 = GetText(txtbx)
msg2 = Msg & vbCrLf & msg2
Call SendMessageByString(txtbx, WM_SETTEXT, 0&, msg2)
icon = FindWindowEx(win, txtbx, "_AOL_Icon", vbNullString)
icon = FindWindowEx(win, icon, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Call WaitForFwdToClose
If closemail Then
Call PostMessage(FindEmail, WM_CLOSE, 0&, 0&)
End If
End Sub

Public Sub ClickModal()
Dim icon, win As Long
Do
win = FindWindow("_AOL_Modal", vbNullString)
Loop Until win <> 0&
Do
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Loop Until icon <> 0&
Call ClickIcon(icon)
End Sub

Public Sub SetMailPrefs()

Dim MDI, win, win2, icon, check, i As Long
If FindAOL = 0& Then Exit Sub
If IsUserOnline Then Call msgbx("SimpleAOL6.bas", "You must be offline to set the mail prefernces")
Call RunIconMenu(Menu_Settings, Menu_Pref)
MDI = FindMDI
win2 = 0&
Do
win2 = FindWindowEx(MDI, win2, "AOL Child", "Preferences")
Loop Until win2 <> 0&
Pause (0.25)
icon = 0&
For i = 1 To 8
icon = FindWindowEx(win2, icon, "_AOL_Icon", vbNullString)
Next i
Call ClickIcon(icon)
win = 0&
Do
win = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until win <> 0&
Pause (0.75)
check = 0&
For i = 1 To 5
DoEvents
check = FindWindowEx(win, check, "_AOL_Checkbox", vbNullString)
Next i
Call SendMessage(check, BM_SETCHECK, False, vbNullString)
check = FindWindowEx(win, check, "_AOL_Checkbox", vbNullString)
Call SendMessage(check, BM_SETCHECK, True, vbNullString)
check = FindWindowEx(win, check, "_AOL_Checkbox", vbNullString)
Call SendMessage(check, BM_SETCHECK, False, vbNullString)
check = FindWindowEx(win, check, "_AOL_Checkbox", vbNullString)
Call SendMessage(check, BM_SETCHECK, False, vbNullString)
check = FindWindowEx(win, check, "_AOL_Checkbox", vbNullString)
Call SendMessage(check, BM_SETCHECK, False, vbNullString)
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Call PostMessage(win2, WM_CLOSE, 0&, 0&)
End Sub

Public Sub WaitForEmailToOpen()

Do
DoEvents
Loop Until IsWindowVisible(FindEmail) <> 0&
End Sub

Public Sub WaitForEmailToClose()

Do
DoEvents
Loop Until IsWindowVisible(FindEmail) = 0&
End Sub

Public Sub WaitForFwdToClose()
Do
DoEvents
Loop Until IsWindowVisible(FindForward) = 0&
End Sub

Public Sub WaitForFwdToOpen()
Do
DoEvents
Loop Until IsWindowVisible(FindForward) <> 0&
End Sub

Public Sub CloseStatus()

Dim win As Long
win = FindWindowEx(FindMDI, 0&, "AOL Child", "Status")
Call PostMessage(win, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetAOLVersion() As Integer

If FindMenuByString(FindAOL, "&What's New in AOL 5.0") Then GetAOLVersion = 5
If FindMenuByString(FindAOL, "&What's New in AOL 6.0") Then GetAOLVersion = 6
End Function

