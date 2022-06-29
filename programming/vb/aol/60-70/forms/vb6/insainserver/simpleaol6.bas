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
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
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
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Const VK_CONTROL = &H11
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const VK_ESCAPE = &H1B
Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_DELETEITEM = &H2D
Public Const WM_DESTROY = &H2
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
Public Const Esc_Key = 96
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
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
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
cap = Getcaption2(win)
If InStr(cap, "Welcome, ") Then
   UserSN = Mid(cap, InStr(cap, ",") + 2, InStr(cap, "!") - (InStr(cap, ",") + 2))
   Exit Function
End If
Loop Until win = 0&
UserSN = ""
End Function

Private Function Getcaption2(ByVal hwnd As Long) As String
    
Dim Buffer As String
Dim txtlen As Long
txtlen& = GetWindowTextLength(hwnd)
Buffer$ = String(txtlen&, 0&)
Call GetWindowText(hwnd&, Buffer$, txtlen + 1)
Getcaption2 = Buffer$
End Function

Private Function GetText(ByVal hwnd As Long) As String

Dim txtlen As Long
Dim Buffer As String
txtlen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer = String(txtlen, 0&)
Call SendMessageByString(hwnd, WM_GETTEXT, txtlen + 1, Buffer)
GetText = Buffer
End Function

Private Sub ClickButton(ByVal Button As Long)

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

Dim Room As String
Room = FindChatRoom
If Room <> 0& Then Call PostMessage(Room, WM_CLOSE, 0&, 0&)
End Sub

Public Function ChatRoomName() As String

ChatRoomName = Getcaption2(FindChatRoom())
End Function

Public Sub SendChat(Chat As String)
Dim dctime As Variant, lroom As Long
lroom = AOLRoomEdit 'uses the sub to find the text box to type in
SetText lroom, Chat 'uses the sub to set the edit box to the text trying to be sent
dctime = Timer
DoEvents
man:
SendChar lroom, 13  'see above
If InStr(1, GetText(lroom), Chat) = 0 Then
    Exit Sub
Else
    GoTo man
End If
checkdeadmsg
End Sub
Public Sub SendChar(hwnd As Long, Char As Byte)
    Dim bResult As Long
    bResult = IsWindow(hwnd)    'see above
    If bResult = 0 Then Exit Sub    'see above
    Call SendMessageByNum(hwnd, WM_CHAR, Char, 0)  'sends the specified character to the specified hwnd
End Sub
Public Function FindIM() As Long

Dim mdi, win, icon As Long
Dim Caption As String
mdi = FindMDI
win = 0&
Do
DoEvents
win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
Caption = Getcaption2(win)
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
Pause2 (1)
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
Pause2 (1)
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
Pause2 (1)
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
Pause2 (1)
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
lst = lst & lstbox.list(i) & ","
Next i
ListBoxToMailString = lst & lstbox.list(lstbox.ListCount - 1)
End Function

Public Sub SendMail(Person As String, cc As String, subject As String, msg As String)
Dim mdi, mail, i, txtbx, icon As Long, aoerror As Long, aomodal As Long, email As Long
Dim mike As String, X As Long, emailedit As Long, emailrich As Long, subjectbox As Long
Dim current As Variant
openit:
Call Keyword("mailto: " & Person)
current = Timer
Do:
    If Timer - current > 1 Then GoTo openit
    checkaol
    mdi = FindMDI
    mail = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
    email& = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
    emailedit& = FindWindowEx(email&, 0&, "_AOL_Edit", vbNullString)
    emailrich& = FindWindowEx(email&, 0&, "RICHCNTL", vbNullString)
Loop Until mail <> 0& And emailedit <> 0& And emailrich <> 0
subjectbox& = FindWindowEx(email&, emailedit&, "_AOL_EDIT", vbNullString)
subjectbox& = FindWindowEx(email&, subjectbox&, "_AOL_EDIT", vbNullString)
Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
Call SendMessageByString(emailrich&, WM_SETTEXT, 0&, msg$)
mike = 0
current = Timer
Do:
    If Timer - current > 10 Then
        Restartaol2
        Exit Sub
    End If
    icon = Findsendnow(FindEmail2)
    Call PostMessage(icon&, WM_KEYDOWN, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_RETURN, 0&)
mike = mike + 1
Loop Until mike = 3
current = Timer
Do: DoEvents
    If Timer - current > 10 Then
        Restartaol2
        Exit Sub
    End If
    aoerror& = FindWindowEx(mdi, 0&, "AOL Child", "Error")
    aomodal& = FindWindow("_AOL_Modal", vbNullString)
    If aomodal& <> 0 Then
        killwin FindEmail2
        Exit Sub
    End If
    If aoerror& <> 0 Then
        killwin aoerror&
fucker:
        killwin FindEmail2
        Exit Do
    End If
Loop Until FindEmail2 = 0&
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
Pause2 (0.5)
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
Pause2 (1.5)
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
Pause2 (0.75)
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
Pause2 (0.75)
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
Pause2 (0.75)
win = FindWindowEx(mdi, 0&, "AOL Child", "Away Message")
icon = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(icon)
Pause2 (0.75)
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
cap = Getcaption2(FindIM)
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

Public Sub MailToList(which As Integer, lstbx As ListBox)

Dim mail, aolTab, aolTab2, i, aoltree, Count, maillist, Length, pos, pos2
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
Length = SendMessage(aoltree, LB_GETTEXTLEN, maillist, 0&)
mailer = String(Length + 1, 0)
Call SendMessageByString(aoltree, LB_GETTEXT, maillist, mailer)
pos = InStr(mailer, Chr(TAB_KEY))
pos = InStr(pos + 1, mailer, Chr(TAB_KEY))
mailer = Right(mailer, Len(mailer) - pos)
lstbx.AddItem mailer
Next maillist
End Sub

Public Function SenderFromMail(which As Integer, idx As Long) As String

Dim mail, aolTab, aolTab2, i, aoltree, Count, maillist, Length, pos, pos2
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
Length = SendMessage(aoltree, LB_GETTEXTLEN, idx, 0&)
mailer = String(Length + 1, 0)
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
cap = Getcaption2(win)
If Left(cap, 4) = "Fwd:" Then
FindForward = win
Exit Function
End If
Loop Until win = 0&
End Function

Public Sub OpenEmailByIndex(which As Integer, idx As Long)

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

Public Function FindEmail() As Long

Dim mdi As Long, win As Long, stat As Long, i As Long
mdi = FindMDI
win = 0&
Do
    DoEvents
    win = FindWindowEx(mdi, win, "AOL Child", vbNullString)
    stat = 0&
    For i = 0 To 1
        DoEvents
        stat = FindWindowEx(win, stat, "_AOL_Static", vbNullString)
    Next i
    If Getcaption2(stat) Like "* of *" Then
        FindEmail = win
        Exit Function
    End If
Loop Until win = 0&
FindEmail = 0&
End Function
Public Function FindEmail2() As Long
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
If Getcaption2(stat) = "Subject:" Then
FindEmail2 = win
Exit Function
End If
Loop Until win = 0&
FindEmail2 = 0&
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
Dim Length As String
Do
DoEvents
Length = GetChatText
Loop Until Length <> lastchat
lastchat = Length
GetValidChatLine = LastChatLine
End Function
Public Function Gettext2(Window As Long) As String

'gets the text of a window.
Dim Buffer As String, TextLength As Long
TextLength& = SendMessage(Window, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength&, 0&)
Call SendMessageByString(Window, WM_GETTEXT, TextLength& + 1, Buffer$)
Gettext2$ = Buffer$
End Function
Public Sub KeepFormOnTop(frm As Form)

' This sub will keep your form ontop of all other's and other applications.
    
' This example will keep the current from on top.
    
' Call KeepFromOnTop(Me)

Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub SetText(hwnd As Long, Text As String)
    Dim bResult As Long, aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)  'locates the aol window
    mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString) 'locates the mdi clinet

    bResult = IsWindow(hwnd)    'see above
    If bResult = 0 Then Exit Sub    'see above
    Call SendMessageByString(hwnd, WM_SETTEXT, 0, Text)   'sets the windows text to the text specified
End Sub
Sub killexplorer()
Dim AOLFrame25 As Long, MDIClient As Long, AOLChild As Long, ATLX60512150 As Long, InternetExplorerServer As Long
AOLFrame25& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame25&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
ATLX60512150& = FindWindowEx(AOLChild&, 0&, "ATL:60512150", vbNullString)
InternetExplorerServer& = FindWindowEx(ATLX60512150&, 0&, "Internet Explorer_Server", vbNullString)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Public Function FindWelcomeScreen() As Long
'gets the aol4 welcome screen.
Dim AOL2 As Long, mdi As Long, child As Long
AOL2& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
mdi& = FindWindowEx(AOL2&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)

If InStr(GetCaption(child&), "Welcome, ") <> 0& Then
    FindWelcomeScreen = child&
    Exit Function
Else
    Do
        child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
        If InStr(GetCaption(child&), "Welcome, ") <> 0& Then
            FindWelcomeScreen = child&
            Exit Function
        End If
    Loop Until child& = 0&
End If
FindWelcomeScreen = 0&
End Function
Public Function FindSignOnScreen() As Long
'gets the aol4 welcome screen.
Dim AOL2 As Long, MDI2 As Long, child As Long
AOL2& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
MDI2& = FindWindowEx(AOL2&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI2&, 0&, "AOL Child", vbNullString)

If InStr(GetCaption(child&), "Sign, ") <> 0& Then
    FindSignOnScreen& = child&
    Exit Function
Else
    Do
        child& = FindWindowEx(MDI2&, child&, "AOL Child", vbNullString)
        If InStr(GetCaption(child&), "Sign, ") <> 0& Then
            FindSignOnScreen& = child&
            Exit Function
        End If
    Loop Until child& = 0&
End If
FindSignOnScreen& = 0&
End Function
Public Function Findclicksignon(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 3
        icon& = FindWindowEx(hwnd, icon&, "_AOL_Icon", vbNullString)
    Next X
Findclicksignon = icon&
End Function
Sub checkdeadmsg()
Dim FullWindow As Long, FullButton As Long
FullWindow& = FindWindow("#32770", "America Online")
FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
If FullWindow& <> 0& Then
    Do
        DoEvents
        Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until FullWindow& = 0& And FullButton& = 0&
End If
End Sub
Sub checkaol()
Dim aolframe As Long, current As Variant, aol As Long, mdi As Long, fmail As Long
aolframe& = FindWindow("AOL Frame25", vbNullString)
If aolframe <> 0& Then
    aolframe& = FindWindow("AOL Frame25", vbNullString)
Else
    'Form1.Line2.Visible = False
    'Form1.Label7.Caption = "There was an error in AOL please wait"
    Do:
        aolframe& = FindWindow("AOL Frame25", vbNullString)
        Pause2 0.1
    Loop Until aolframe& = 0
    Pause2 3
    Restartaol
    ServerForm.roombust1.Room = ServerForm.Text4.Text
    ServerForm.roombust1.PrivateBust
    Do: DoEvents
    If ServerForm.roombust1.Busted = True Then Exit Do
    Loop
openplease:
    Call RunMenuByString("Filing &Cabinet")
    current = Timer
    Do:
        If Timer - current > 3 Then GoTo openplease
        aol& = FindWindow("AOL Frame25", vbNullString)
        mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
        fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
    Loop Until fmail& <> 0
    Call ShowWindow(fmail&, SW_MINIMIZE)
    ServerForm.Timer1 = True
    ServerForm.Timer2 = True
    ServerForm.Timer3 = True
    ServerForm.Timer4 = True
    ServerForm.Timer5 = True
    ServerForm.Timer6 = True
    ServerForm.Timer8 = True
    SendChat " Recoving from WAOL!"
    ServerForm.Restartaol.Caption = "0"
    Pause2 1
    MailOpenEmailFlash (7)
End If
End Sub
Public Sub SendMail2(Person As String, cc As String, subject As String, msg As String)
Dim mdi, mail, i, txtbx, icon As Long, aoerror As Long, aomodal As Long, email As Long
Dim mike As String, X As Long, emailedit As Long, emailrich As Long, subjectbox As Long
Dim current As Variant
openit:
Call Keyword("mailto: " & Person)
current = Timer
Do:
    If Timer - current > 1 Then
        If FindEmail2 = 0 Then
            GoTo openit:
        End If
    End If
    checkaol
    mdi = FindMDI
    mail = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
    email& = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
    emailedit& = FindWindowEx(email&, 0&, "_AOL_Edit", vbNullString)
    emailrich& = FindWindowEx(email&, 0&, "RICHCNTL", vbNullString)
Loop Until mail <> 0& And emailedit <> 0& And emailrich <> 0
subjectbox& = FindWindowEx(email&, emailedit&, "_AOL_EDIT", vbNullString)
subjectbox& = FindWindowEx(email&, subjectbox&, "_AOL_EDIT", vbNullString)
Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
Call SendMessageByString(emailrich&, WM_SETTEXT, 0&, msg$)
mike = 0
current = Timer
Do:
    If Timer - current > 10 Then
        Restartaol2
        Exit Sub
    End If
    icon = Findsendnow(FindEmail2)
    Call PostMessage(icon&, WM_KEYDOWN, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_RETURN, 0&)
mike = mike + 1
Loop Until mike = 3
current = Timer
Do:
    If Timer - current > 10 Then
        StopIncomingText
        killwin FindEmail2
        Exit Sub
    End If
    aoerror& = FindWindowEx(mdi, 0&, "AOL Child", "Error")
    aomodal& = FindWindow("_AOL_Modal", vbNullString)
    If aomodal& <> 0 Then
        killwin FindEmail2
        Exit Sub
    End If
    If aoerror& <> 0 Then
        killwin aoerror&
fucker:
        killwin FindEmail2
        Exit Do
    End If
Loop Until FindEmail2 = 0&
End Sub

