Attribute VB_Name = "digitek"
' digitek AIM module 1.0 |
' by grn eminem          |
' greetz to digiwizard   |
' for use with AIM 3.5   |
' made in vb5            |
'-------------------------
'there are no ad hiders in this bas
'simply because if you do remove them
'they are distorted and more annoying
'then what they were originally
'also, getting the buddy list users was
'excluded since they have the "Offline"
'section now

'all subs/functions written by me, unless
'stated otherwise in the sub/function

'coming attractions...
'in digitek 1.1 there will be:
'- distort for chat/im
'- colored links for chat/im
'- im popup

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_SETCURSEL = &H186

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20

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
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

'---------------------------------------'
' Sub's/Function's that aid other Sub's '
' in this Module                        '
'---------------------------------------'

Sub Pause(interval)
' No use to you...
Dim Current
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Sub RunMenuByString(Application, StringSearch)
' No use to you...
Dim ToSearch As Integer, MenuCount As Integer, FindString
Dim ToSearchSub As Integer, MenuItemCount As Integer, GetString
Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
Dim MenuItem As Integer, RunTheMenu As Integer
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
Function FilterHtml(Amount%) As String
' From NewPsyche.bas
' No use to you...
Dim intLoop As Integer, stringChat As String
stringChat$ = ChatText(Amount%)
stringChat$ = Win_Replace(LCase(stringChat$), "<br>", Chr(13))
For intLoop% = 1 To CountTags(stringChat$)
lLeft2& = InStr(1, stringChat$, "<")
lRight2& = InStr(1, stringChat$, ">")
sLeft$ = Left(stringChat$, lLeft2& - 1)
sFull& = lRight2& - lLeft2& + 1
sRight$ = Right(stringChat$, Len(stringChat$) - (Len(sLeft$) + sFull&))
stringChat$ = sLeft$ + sRight$
Next intLoop%
End2:
FilterHtml$ = stringChat$
End Function
Function LastEnter(strText As String) As Integer
' From NewPsyche.bas
' No use to you...
Dim intLoop As Integer
For intLoop% = 1 To Len(strText$)
lEnter& = InStr(intLoop%, strText$, Chr(13))
If lEnter& > 0 Then
lEnters% = lEnter&
intLoop% = lEnter& + 1
Else:
GoTo End1
End If
Next intLoop%
End1:
LastEnter% = lEnters%
End Function
Function Win_Replace(strText01 As String, stToReplace As String, stReplaceWith As String) As String
' No use to you...
For intLoop% = 1 To Len(strText01)
lFindChar& = InStr(1, LCase(strText01$), LCase(stToReplace$))
If lFindChar& = 0 Then
Win_Replace$ = strText01$
Exit Function
End If
lFindChar& = InStr(1, LCase(strText01$), LCase(stToReplace$))
strCharL$ = Left(strText01$, lFindChar& - 1)
strCharR$ = Right(strText01$, Len(strText01$) - (lFindChar& + Len(stToReplace$) - 1))
strCharF$ = strCharL$ + stReplaceWith$ + strCharR$
strText01$ = strCharF$
intLoop% = lFindChat& + 1
Next intLoop%
Win_Replace$ = strCharF$
End Function
Function CountTags(strText As String) As Integer
' From NewPsyche.bas
' No use to you...
Dim lLeft As Integer, lRight As Integer, intLoop As Integer
For intLoop% = 1 To Len(strText$)
If InStr(intLoop%, strText$, "<") > 0 Then
lLeft% = lLeft% + 1
Else:
GoTo End1
End If
If InStr(intLoop%, strText$, ">") > 0 Then
intLoop% = InStr(intLoop%, strText$, ">") '+ 1
lRight% = lRight% + 1
Else: GoTo End1
End If
Next intLoop%
End1: If lLeft% <> lRight% Then
CountTags% = 0
Else: CountTags% = lRight%
End If
End Function
Function ChatText(Amount As Integer) As String
' From NewPsyche.bas
' No use to you...
If Amount% = 0 Then Exit Function
FindChat1& = FindWindow("AIM_ChatWnd", vbNullString)
lChText& = FindWindowEx(FindChat1&, 0, "WndAte32Class", vbNullString)
String1$ = Win_GetTxt(lChText&)
If Len(String1$) = 0 Then Exit Function
If Len(String1$) > Amount% Then
String2$ = Right(String1$, Amount%)
lLeft& = InStr(1, String2$, "<")
String1$ = Right(String2$, Len(String2$) - lLeft& + 1)
End If
ChatText$ = String1$
End Function
Function Win_GetTxt(ByVal lWinHandle As Long) As String
' No use to you...
Dim tWinLength As Long, stTxt As String
tWinLength = SendMessage(lWinHandle, WM_GETTEXTLENGTH, 0&, 0&)
stTxt$ = String(tWinLength&, 0&)
Call SendMessageByString(lWinHandle&, WM_GETTEXT, tWinLength& + 1, stTxt$)
Win_GetTxt$ = stTxt$
End Function
Function Win_GetCap(ByVal lWinHandle As Long) As String
' No use to you...
Dim lWinLength As Long, stCap As String
lWinLength = GetWindowTextLength(lWinHandle)
stCap$ = String(lWinLength&, 0&)
Call GetWindowText(lWinHandle&, stCap$, lWinLength& + 1)
Win_GetCap$ = stCap$
End Function

'-----------------------------------'
' Sub's/Function's for use with AIM '
'-----------------------------------'

Sub AIMLoad()
Dim AIM As Long, NoFreeze As Integer
AIM& = Shell("C:\Program Files\AIM95\aim.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Sub AIMExit()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call SendMessageLong(oscarbuddylistwin&, WM_CLOSE, 0&, 0&)
End Sub
Sub AIMMaximize()
Dim BuddyList As Long, Mini As Long
BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Mini& = ShowWindow(BuddyList&, SW_MAXIMIZE)
End Sub
Sub AIMMinimize()
Dim BuddyList As Long, Mini As Long
BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Mini& = ShowWindow(BuddyList&, SW_MINIMIZE)
End Sub
Sub AIMHide()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call ShowWindow(oscarbuddylistwin&, SW_HIDE)
End Sub
Sub AIMShow()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call ShowWindow(oscarbuddylistwin&, SW_SHOW)
End Sub
Sub AIMHideButtons()
'IM Button
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_HIDE)

'Chat Button
Dim oscarbuddylistwin2&
Dim oscartabgroup2&
Dim oscariconbtn2&
oscarbuddylistwin2& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup2& = FindWindowEx(oscarbuddylistwin2&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn2& = FindWindowEx(oscartabgroup2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(oscartabgroup2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_HIDE)

'Talk Button
Dim oscarbuddylistwin3&
Dim oscartabgroup3&
Dim oscariconbtn3&
oscarbuddylistwin3& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup3& = FindWindowEx(oscarbuddylistwin3&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_HIDE)
End Sub
Sub AIMShowButtons()
'IM Button
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_SHOW)

'Chat Button
Dim oscarbuddylistwin2&
Dim oscartabgroup2&
Dim oscariconbtn2&
oscarbuddylistwin2& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup2& = FindWindowEx(oscarbuddylistwin2&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn2& = FindWindowEx(oscartabgroup2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(oscartabgroup2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_SHOW)

'Talk Button
Dim oscarbuddylistwin3&
Dim oscartabgroup3&
Dim oscariconbtn3&
oscarbuddylistwin3& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup3& = FindWindowEx(oscarbuddylistwin3&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(oscartabgroup3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_SHOW)
End Sub
Sub AIMHideBuddyList()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartree&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)
Call ShowWindow(oscartree&, SW_HIDE)
End Sub
Sub AIMShowBuddyList()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartree&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)
Call ShowWindow(oscartree&, SW_SHOW)
End Sub
Sub AIMHideTabs()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartabctrl&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartabctrl& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tabctrl", vbNullString)
Call ShowWindow(oscartabctrl&, SW_HIDE)
End Sub
Sub AIMShowTabs()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartabctrl&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartabctrl& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tabctrl", vbNullString)
Call ShowWindow(oscartabctrl&, SW_SHOW)
End Sub
Sub AIMHideAll()
' Hides Tabs/Buttons/BuddyList
Dim oscarbuddylistwin&
Dim oscartabgroup&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
Call ShowWindow(oscartabgroup&, SW_HIDE)
End Sub
Sub AIMShowAll()
' Show Tabs/Buttons/BuddyList
Dim oscarbuddylistwin&
Dim oscartabgroup&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
Call ShowWindow(oscartabgroup&, SW_SHOW)
End Sub
Sub AIMHideSearchBar()
Dim oscarbuddylistwin&
Dim editx&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call ShowWindow(editx&, SW_HIDE)

Dim oscarbuddylistwin2&
Dim oscariconbtn2&
oscarbuddylistwin2& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn2& = FindWindowEx(oscarbuddylistwin2&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_HIDE)
End Sub
Sub AIMShowSearchBar()
Dim oscarbuddylistwin&
Dim editx&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call ShowWindow(editx&, SW_SHOW)

Dim oscarbuddylistwin2&
Dim oscariconbtn2&
oscarbuddylistwin2& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn2& = FindWindowEx(oscarbuddylistwin2&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_SHOW)
End Sub
Sub IMOpen(ScreenName As String, Message As String, SendIM As Boolean, CloseIM As Boolean)
' Ex:  Call IMOpen("grn eminem", "", False, False)
'      That opens a blank IM for grn eminem
' Ex2: call IMOpen("grn eminem", "hello", True, True)
'      That opens an IM for grn eminem with
'      the message hello, then sends and
'      closes the IM.
x = ShellExecute(0, "open", "aim:goim?screenname=" & ScreenName & "&message=" & Message, vbNullString, vbNullString, 3)
If SendIM = True Then
Pause 0.2
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
End If
If SendIM = False Then
GoTo 10
End If
If CloseIM = True Then
Dim aimimessage2&
aimimessage2& = FindWindow("aim_imessage", vbNullString)
Call SendMessageLong(aimimessage&, WM_CLOSE, 0&, 0&)
End If
If CloseIM = False Then
GoTo 10
End If
10 End Sub
Sub IMClose()
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call SendMessageLong(aimimessage&, WM_CLOSE, 0&, 0&)
End Sub
Sub IMClear()
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "")
End Sub
Sub IMHide()
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call ShowWindow(aimimessage&, SW_HIDE)
End Sub
Sub IMShow()
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call ShowWindow(aimimessage&, SW_SHOW)
End Sub
Sub IMMaximize()
Dim IMWin As Long, Mini As Long
IMWin& = FindWindow("AIM_IMessage", vbNullString)
Mini& = ShowWindow(IMWin&, SW_MAXIMIZE)
End Sub
Sub IMMinimize()
Dim IMWin As Long, Mini As Long
IMWin& = FindWindow("AIM_IMessage", vbNullString)
Mini& = ShowWindow(IMWin&, SW_MINIMIZE)
End Sub
Function IMGetText()
' Ex: Text1.Text = IMGetText
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(ateclass&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(ateclass&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
IMGetText = TheText$
End Function
Sub IMHideButtons()
'Warn Button
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_HIDE)

'Block Button
Dim aimimessage2&
Dim oscariconbtn2&
aimimessage2& = FindWindow("aim_imessage", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_HIDE)

'Add Buddy Button
Dim aimimessage3&
Dim oscariconbtn3&
aimimessage3& = FindWindow("aim_imessage", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_HIDE)

'Talk Button
Dim aimimessage4&
Dim oscariconbtn4&
aimimessage4& = FindWindow("aim_imessage", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn4&, SW_HIDE)

'Info Button
Dim aimimessage5&
Dim oscariconbtn5&
aimimessage5& = FindWindow("aim_imessage", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn5&, SW_HIDE)

'Send Button
Dim aimimessage6&
Dim oscariconbtn6&
aimimessage6& = FindWindow("aim_imessage", vbNullString)
oscariconbtn6& = FindWindowEx(aimimessage6&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn6&, SW_HIDE)

'Seperators
Dim aimimessage7&
Dim oscarseparator7&
aimimessage7& = FindWindow("aim_imessage", vbNullString)
oscarseparator7& = FindWindowEx(aimimessage7&, 0&, "_oscar_separator", vbNullString)
Call ShowWindow(oscarseparator7&, SW_HIDE)

Dim aimimessage8&
Dim oscarseparator8&
aimimessage8& = FindWindow("aim_imessage", vbNullString)
oscarseparator8& = FindWindowEx(aimimessage8&, 0&, "_oscar_separator", vbNullString)
oscarseparator8& = FindWindowEx(aimimessage8&, oscarseparator8&, "_oscar_separator", vbNullString)
Call ShowWindow(oscarseparator8&, SW_HIDE)
End Sub
Sub IMShowButtons()
'Warn Button
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_SHOW)

'Block Button
Dim aimimessage2&
Dim oscariconbtn2&
aimimessage2& = FindWindow("aim_imessage", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_SHOW)

'Add Buddy Button
Dim aimimessage3&
Dim oscariconbtn3&
aimimessage3& = FindWindow("aim_imessage", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimimessage3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_SHOW)

'Talk Button
Dim aimimessage4&
Dim oscariconbtn4&
aimimessage4& = FindWindow("aim_imessage", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimimessage4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn4&, SW_SHOW)

'Info Button
Dim aimimessage5&
Dim oscariconbtn5&
aimimessage5& = FindWindow("aim_imessage", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
oscariconbtn5& = FindWindowEx(aimimessage5&, oscariconbtn5&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn5&, SW_SHOW)

'Send Button
Dim aimimessage6&
Dim oscariconbtn6&
aimimessage6& = FindWindow("aim_imessage", vbNullString)
oscariconbtn6& = FindWindowEx(aimimessage6&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn6&, SW_SHOW)

'Seperators
Dim aimimessage7&
Dim oscarseparator7&
aimimessage7& = FindWindow("aim_imessage", vbNullString)
oscarseparator7& = FindWindowEx(aimimessage7&, 0&, "_oscar_separator", vbNullString)
Call ShowWindow(oscarseparator7&, SW_SHOW)
Dim aimimessage8&
Dim oscarseparator8&
aimimessage8& = FindWindow("aim_imessage", vbNullString)
oscarseparator8& = FindWindowEx(aimimessage8&, 0&, "_oscar_separator", vbNullString)
oscarseparator8& = FindWindowEx(aimimessage8&, oscarseparator8&, "_oscar_separator", vbNullString)
Call ShowWindow(oscarseparator8&, SW_SHOW)
End Sub
Sub IMHideRateMeter()
Dim aimimessage&
Dim oscarratemeter&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarratemeter& = FindWindowEx(aimimessage&, 0&, "_oscar_ratemeter", vbNullString)
Call ShowWindow(oscarratemeter&, SW_HIDE)
End Sub
Sub IMShowRateMeter()
Dim aimimessage&
Dim oscarratemeter&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarratemeter& = FindWindowEx(aimimessage&, 0&, "_oscar_ratemeter", vbNullString)
Call ShowWindow(oscarratemeter&, SW_SHOW)
End Sub
Sub IMHiddenMessage(ScreenName As String, Message As String)
'Ex: Call IMHiddenMessage("grn eminem", "hello")
'    This hides the message in the IM with a
'    code used for TimeStamps, for your friend
'    to view it he/she must press F2
Call IMOpen(ScreenName, "<!-- " & Message & " -->" & " ", True, False)
End Sub
Sub IMFontFreak(ScreenName As String, FontName As String, Text As String)
' Ex: Call IMFontFreak("grn eminem", "I own you", "hi")
'     This Instant Message's grn eminem, and in
'     his font list adds I own you
'     and also sends the text to the IM hi.
'     Font list located anywhere you type:
'     Right Click/Text/Font Name
Call IMOpen(ScreenName, "</font><font face=" & Chr(34) & FontName & Chr(34) & ">" & Text & "</font>", True, False)
End Sub
Sub IMFocus(ScreenName As String)
' This simply focus' the IM.
' An IM must be open before this will work
' Or else it will open a new one.
x = ShellExecute(0, "open", "aim:goim?screenname=" & ScreenName, vbNullString, vbNullString, 3)
End Sub
Sub IMCloseAll()
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Do
DoEvents:
If aimimessage& > 0 Then
Call IMClose
End If
Loop Until aimimessage& = 0
End Sub
Sub IMGetFile(ScreenName As String)
x = ShellExecute(0, "open", "aim:getfile?screenname=" & ScreenName, vbNullString, vbNullString, 3)
End Sub
Sub IMSendFile(ScreenName As String, File As String, Send As Boolean)
x = ShellExecute(0, "open", "aim:goim?screenname=" & ScreenName, vbNullString, vbNullString, 3)
aimimessage& = FindWindow("AIM_IMessage", vbNullString)
x0& = FindWindowEx(aimimessage&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(aimimessage, "Send &File...")
Dim aimimessage1&
Dim x1&
Dim editx1&
x1& = FindWindow("#32770", vbNullString)
editx1& = FindWindowEx(x1&, 0&, "edit", vbNullString)
Call SendMessageByString(editx1&, WM_SETTEXT, 0&, File)
If Send = True Then
Dim aimimessage3&
Dim x3&
Dim button3&
x3& = FindWindow("#32770", vbNullString)
button3& = FindWindowEx(x3&, 0&, "button", vbNullString)
button3& = FindWindowEx(x3&, button3&, "button", vbNullString)
Call SendMessageLong(button3&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessageLong(button3&, WM_KEYUP, VK_SPACE, 0&)
End If
If Send = False Then
GoTo 10
End If
10 End Sub
Sub IMTalk(ScreenName As String)
Call IMOpen(ScreenName, "", False, False)
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
Call IMClose
End Sub
Function IMGetName() As String
' From NewPsyche.bas
On Error Resume Next
String1$ = Win_GetCap(FindIM&)
IMGetName$ = Left(String1$, Len(String1$) - 18)
End Function
Function IMCount() As Integer
' From NewPsyche.bas
Dim IMWin As Long, lngInt As Long
lngInt& = -1
IMWin& = 0
Do: DoEvents
IMWin& = FindWindowEx(0, IMWin&, "AIM_IMessage", vbNullString)
lngInt& = lngInt& + 1
Loop Until IMWin& = 0
IMCount% = lngInt&
End Function
Sub AIMGetInfo(ScreenName As String)
Call IMOpen(ScreenName, "", False, False)
Pause 0.1
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
Call IMClose
End Sub
Function AIMGetUserName() As String
' From NewPsyche.bas
On Error Resume Next
String1$ = Win_GetCap(FindMain&)
String2$ = Left(String1$, Len(String1$) - 20)
string01$ = Win_GetCap(FindMain&)
lApost& = InStr(1, string01$, "'")
string02$ = Left(string01$, lApost& - 1)
If String2$ = string02$ Then
GetUserName$ = String2$
Else
GetUserName$ = "unknown"
End If
End Function
Sub AIMBlockUser(ScreenName As String)
Call IMOpen(ScreenName, "", False, False)
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
Call IMClose
End Sub
Sub AIMCloseErrors()
Dim aimchatwnd&
Dim x&
x& = FindWindow("#32770", vbNullString)
Call SendMessageLong(x&, WM_CLOSE, 0&, 0&)
End Sub
Function IMLastLine() As String
' From NewPsyche.bas
String1$ = FilterHtml2(400)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml2(700)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml2(900)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml2(1200)
IMLastLine$ = Right(String1$, Len(String1$) - LastEnter(String1$))
End Function
Function IMLastName() As String
' From NewPsyche.bas
On Error Resume Next
String1$ = IMLastLine$
lColon& = InStr(1, String1$, ":")
IMLastName$ = Left(String1$, lColon& - 1)
End Function
Function IMLastText() As String
' From NewPsyche.bas
String1$ = IMLastLine$
lColon& = InStr(1, String1$, ":")
IMLastText$ = Right(String1$, Len(String1$) - (lColon + 1))
End Function
Sub IMPuntUser(ScreenName As String, FakeText As String)
Call IMOpen(ScreenName, "<a href=" & Chr(34) & "aim:kill?user" & Chr(34) & ">" & FakeText & "</a>", True, False)
End Sub
Sub ChatOpen(RoomName As String, Exchange As String)
' Ex:  Call ChatOpen("vb", "")
'      That opens to the normal chat room vb.
' Ex2: Call ChatOpen("vb", "5")
'      This opens to the chat room vb on
'      exchange 5, there are 3 exchanges
'      4, 5 and 6. 4 is default.
x = ShellExecute(0, "open", "aim:gochat?roomname=" & RoomName & "&exchange=" & Exchange, vbNullString, vbNullString, 3)
If Exchange = "" Then
Exchange = 4
End If
End Sub
Sub ChatSend(Message As String)
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, Message)

Dim aimchatwnd2&
Dim oscariconbtn&
aimchatwnd2& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub ChatClose()
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call SendMessageLong(aimchatwnd&, WM_CLOSE, 0&, 0&)
End Sub
Sub ChatClear()
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "")
End Sub
Sub ChatHide()
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call ShowWindow(aimchatwnd&, SW_HIDE)
End Sub
Sub ChatShow()
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call ShowWindow(aimchatwnd&, SW_SHOW)
End Sub
Function ChatGetText()
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(ateclass&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(ateclass&, WM_GETTEXT, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
ChatGetText = TheText$
End Function
Sub ChatInvite(ScreenNames As String, Message As String, ChatRoom As String)
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)

Dim aimchatinvitesendwnd&
Dim editx&
aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, ScreenNames)

Dim aimchatinvitesendwnd2&
Dim editx2&
aimchatinvitesendwnd2& = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx2& = FindWindowEx(aimchatinvitesendwnd2&, 0&, "edit", vbNullString)
editx2& = FindWindowEx(aimchatinvitesendwnd2&, editx2&, "edit", vbNullString)
Call SendMessageByString(editx2&, WM_SETTEXT, 0&, Message)

Dim aimchatinvitesendwnd3&
Dim editx3&
aimchatinvitesendwnd3& = FindWindow("aim_chatinvitesendwnd", vbNullString)
editx3& = FindWindowEx(aimchatinvitesendwnd3&, 0&, "edit", vbNullString)
editx3& = FindWindowEx(aimchatinvitesendwnd3&, editx3&, "edit", vbNullString)
editx3& = FindWindowEx(aimchatinvitesendwnd3&, editx3&, "edit", vbNullString)
Call SendMessageByString(editx3&, WM_SETTEXT, 0&, ChatRoom)

Dim aimchatinvitesendwnd4&
Dim oscariconbtn4&
aimchatinvitesendwnd4& = FindWindow("aim_chatinvitesendwnd", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatinvitesendwnd4&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatinvitesendwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatinvitesendwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
Call SendMessageLong(oscariconbtn4&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(oscariconbtn4&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub ChatHiddenMessage(Message As String)
'Ex: Call ChatHiddenMessage("hello")
'    This hides the message in the chat room
'    with a code used for TimeStamps,
'    for your friend to view it he/she
'    must press F2
Call ChatSend("<!-- " & Message & " -->" & " ")
End Sub
Sub ChatFontFreak(FontName As String, Text As String)
' Ex: Call ChatFontFreak("I own you", "hi")
'     This makes everyone's font list in the
'     chat room I own you, and sends the text
'     hi to the whole room normally.
'     Font list located anywhere you type:
'     Right Click/Text/Font Name
Call ChatSend("</font><font face=" & Chr(34) & FontName & Chr(34) & ">" & Text & "</font>")
End Sub
Sub ChatHideButtons()
'IM Button
Dim aimchatwnd&
Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_HIDE)

'Ignore Button
Dim aimchatwnd2&
Dim oscariconbtn2&
aimchatwnd2& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn2& = FindWindowEx(aimchatwnd2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimchatwnd2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_HIDE)

'Talk Button
Dim aimchatwnd3&
Dim oscariconbtn3&
aimchatwnd3& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_HIDE)

'Info Button
Dim aimchatwnd4&
Dim oscariconbtn4&
aimchatwnd4& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn4&, SW_HIDE)

'More Button
Dim aimchatwnd5&
Dim button5&
aimchatwnd5& = FindWindow("aim_chatwnd", vbNullString)
button5& = FindWindowEx(aimchatwnd5&, 0&, "button", vbNullString)
Call ShowWindow(button5&, SW_HIDE)

'Send Button
Dim aimchatwnd6&
Dim oscariconbtn6&
aimchatwnd6& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn6&, SW_HIDE)
End Sub
Sub ChatShowButtons()
'IM Button
Dim aimchatwnd&
Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn&, SW_SHOW)

'Ignore Button
Dim aimchatwnd2&
Dim oscariconbtn2&
aimchatwnd2& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn2& = FindWindowEx(aimchatwnd2&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn2& = FindWindowEx(aimchatwnd2&, oscariconbtn2&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn2&, SW_SHOW)

'Talk Button
Dim aimchatwnd3&
Dim oscariconbtn3&
aimchatwnd3& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
oscariconbtn3& = FindWindowEx(aimchatwnd3&, oscariconbtn3&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn3&, SW_SHOW)

'Info Button
Dim aimchatwnd4&
Dim oscariconbtn4&
aimchatwnd4& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
oscariconbtn4& = FindWindowEx(aimchatwnd4&, oscariconbtn4&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn4&, SW_SHOW)

'More Button
Dim aimchatwnd5&
Dim button5&
aimchatwnd5& = FindWindow("aim_chatwnd", vbNullString)
button5& = FindWindowEx(aimchatwnd5&, 0&, "button", vbNullString)
Call ShowWindow(button5&, SW_SHOW)

'Send Button
Dim aimchatwnd6&
Dim oscariconbtn6&
aimchatwnd6& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
oscariconbtn6& = FindWindowEx(aimchatwnd6&, oscariconbtn6&, "_oscar_iconbtn", vbNullString)
Call ShowWindow(oscariconbtn6&, SW_SHOW)
End Sub
Sub ChatHideRateMeter()
Dim aimchatwnd&
Dim oscarratemeter&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscarratemeter& = FindWindowEx(aimchatwnd&, 0&, "_oscar_ratemeter", vbNullString)
Call ShowWindow(oscarratemeter&, SW_HIDE)
End Sub
Sub ChatShowRateMeter()
Dim aimchatwnd&
Dim oscarratemeter&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscarratemeter& = FindWindowEx(aimchatwnd&, 0&, "_oscar_ratemeter", vbNullString)
Call ShowWindow(oscarratemeter&, SW_SHOW)
End Sub
Sub ChatFocus(RoomName As String, Exchange As String)
' This simply focus' the Chat Room.
' Chat must be already open to use or it will
' create a new one.
x = ShellExecute(0, "open", "aim:gochat?roomname=" & RoomName & "&exchange=" & Exchange, vbNullString, vbNullString, 3)
End Sub
Sub ChatBlankLine()
' This sends a blank line to the chat room,
' with just your screen name, it doesn't
' use just a space and sends, as that won't
' show up, it uses a special code.
Call ChatSend(" ")
End Sub
Function ChatLastLine() As String
' From NewPsyche.bas
String1$ = FilterHtml(400)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml(700)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml(900)
If InStr(1, String1$, ":") = 0 Then _
String1$ = FilterHtml(1200)
ChatLastLine$ = Right(String1$, Len(String1$) - LastEnter(String1$))
End Function
Function ChatLastName() As String
' From NewPsyche.bas
On Error Resume Next
stringChat$ = ChatLastLine
If InStr(1, stringChat$, ":") = 0 Then
ChatLastName$ = "None"
End If
lColon& = InStr(1, stringChat$, ":")
ChatLastName$ = Left(stringChat$, lColon& - 1)
End Function
Function ChatLastText() As String
' From NewPsyche.bas
On Error Resume Next
stringChat$ = ChatLastLine
If InStr(1, stringChat$, ":") = 0 Then
ChatLastText$ = "None"
End If
lColon& = InStr(1, stringChat$, ": ")
ChatLastText$ = Right(stringChat$, Len(stringChat$) - lColon& - 1)
End Function
Function ChatName() As String
' From NewPsyche.bas
String1$ = Win_GetCap(FindChat&)
ChatName$ = Right(String1$, Len(String1$) - 11)
End Function
Sub ChatCloseAll()
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Do
DoEvents:
If aimchatwnd& > 0 Then
Call ChatClose
End If
Loop Until aimchatwnd& = 0
End Sub
Function ChatCountUsers()
Dim aimchatwnd&
Dim oscartree&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscartree& = FindWindowEx(aimchatwnd&, 0&, "_oscar_tree", vbNullString)
Dim LCount&
LCount& = SendMessageLong(oscartree&, LB_GETCOUNT, 0&, 0&)
ChatCountUsers = LCount&
End Function
Sub ChatGetList(List As ListBox)
' From Digital AIM.bas
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
Dim name As String, NameLen, buffer As String
Dim TabPos, NameText As String, Text As String
Dim mooz, Well As Integer, BuddyTree As Long
ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)
If ChatRoom& <> 0 Then
Do
BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
buffer$ = String$(NameLen, 0)
Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
TabPos = InStr(buffer$, Chr$(9))
NameText$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
TabPos = InStr(NameText$, Chr$(9))
Text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
name$ = Text$
For mooz = 0 To List.ListCount - 1
If name$ = List.List(mooz) Then
Well% = 123
GoTo Endz
End If
Next mooz
If Well% <> 123 Then
List.AddItem name$
Else
End If
Endz:
Next MooLoo
End If
End Sub
Sub ChatGetCombo(Combo As ComboBox)
' From Digital AIM.bas
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
Dim name As String, NameLen, buffer As String
Dim TabPos, NameText As String, Text As String
Dim mooz, Well As Integer, BuddyTree As Long
ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)
If ChatRoom& <> 0 Then
Do
BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
buffer$ = String$(NameLen, 0)
Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
TabPos = InStr(buffer$, Chr$(9))
NameText$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
TabPos = InStr(NameText$, Chr$(9))
Text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
name$ = Text$
For mooz = 0 To Combo.ListCount - 1
If name$ = Combo.List(mooz) Then
Well% = 123
GoTo Endz
End If
Next mooz
If Well% <> 123 Then
Combo.AddItem name$
Else
End If
Endz:
Next MooLoo
End If
End Sub
Sub ChatPuntUser(FakeText As String)
Call ChatSend("<a href=" & Chr(34) & "aim:kill?user" & Chr(34) & ">" & FakeText & "</a>")
End Sub
Sub SetBuddyCaption(Caption As String)
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call SendMessageByString(oscarbuddylistwin&, WM_SETTEXT, 0&, Caption)
End Sub
Sub SetIMCaption(Caption As String)
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call SendMessageByString(aimimessage&, WM_SETTEXT, 0&, Caption)
End Sub
Sub SetChatCaption(Caption As String)
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call SendMessageByString(aimchatwnd&, WM_SETTEXT, 0&, Caption)
End Sub
Sub SetAwayCaption(Caption As String)
Dim oscarbuddylistwin&
Dim AwayWindow&
AwayWindow& = FindWindow("#32770", vbNullString)
Call SendMessageByString(AwayWindow&, WM_SETTEXT, 0&, Caption)
End Sub
Sub SetInfoCaption(Caption As String)
Dim oscarlocate&
oscarlocate& = FindWindow("_oscar_locate", vbNullString)
Call SendMessageByString(oscarlocate&, WM_SETTEXT, 0&, Caption)
End Sub
Sub SetUserProfile(Text As String)
' This Sub changes the text in a
' profile you are viewing of a user.
' It only appears to you to be changed.
Dim oscarlocate&
Dim wndateclass&
Dim ateclass&
oscarlocate& = FindWindow("_oscar_locate", vbNullString)
wndateclass& = FindWindowEx(oscarlocate&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, Text)
End Sub


