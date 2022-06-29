Attribute VB_Name = "HixAIMBas"
' Show/Hide Menu: Deleted: AOL Deleted that and added it at the top
' Bold_Underlined_Italic SendChat: Deleted: The AIM Chat won't recognize that it was stopped with </B>, </I>, </U>
' FuckBuddyList: Deleted: No point in it
' StayOnTop: Fixed
' SendChatInvite Code: Fixed: Sends to Chat Now, correctly
' The New IM/Chat Invite: Fixed: Doesn't click ChatBtn, and vice versa
' SendIM_2 and SendChatInvite_2: Added.
' OpenNewIM2 and OpenNewChatInvite2: Added.

' Give me suggestions on any other subs I should add.  I
' just might listen to your ideas ;P.  Any bugs should also be
' reported here at:
' Hix-Prog-Media@n64rocks.com, AIM: Hix Media
'
' This bas is for AOL Instant Messenger version Beta 2.o
' NOTE: Works for and tested with: 2.0.741(newest)
'
' This is the ONLY AOL Instant Messenger 2.0 Beta Bas Out there!!!
' As Aol keeps coming out with new ver., I will keep updating.
' Be proud to have on your comp a bas file by Hix! lol
'
' For now, I am just going to release this bas, as a Beta.  I
' just got vb6, Don't worry, next I will add a LOT more subs
' so, wait for a week, or so.
'
' If you use this, I ask you: Please, I put hella lot of time
' into this bas, so please give credit where it is due, and I
' ask you NOT to copy, steal, and edit any of my codes.
' If you want to uses these, why copy the codes?  Its easier
' just to use this bas!  Thanx, in advance, I and I hope you
' have phun!!
'
' Hix progging it in da 9d8!
' bow down to da masta!
' 1998 Hix Inc.

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal giFlags As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Const WM_CLOSE = &H10
Public Const WM_SETTEXT = &HC
Public Const WM_COMMAND = &H111
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const GW_HWNDNEXT = 2

Sub ChatSend(Text As String)

ChatSend1% = FindWindow("AIM_ChatWnd", vbNullString)
    If ChatSend1% = 0 Then Info% = MsgBox("There is no chat room open, so please open one", vbInformation + vbOKOnly, "Error!")
ChatSend2% = FindChildByClass(ChatSend1%, "_Oscar_Separator")
ChatSend3% = GetWindow(ChatSend2%, GW_HWNDNEXT)
ChatSend4% = GetWindow(ChatSend3%, GW_HWNDNEXT)
ChatSend5% = SendMessageByString(ChatSend3%, WM_SETTEXT, 0, Text$)
ClickIcon (ChatSend4%)
End Sub

Sub SendIM(Who As String, What As String)

Call OpenNewIM
SendIM1% = FindWindow("AIM_IMessage", vbNullString)
SendIM2% = FindChildByClass(SendIM1%, "_Oscar_PersistantComb")
SendIM3% = FindChildByClass(SendIM2%, "Edit")
SendIM4% = SendMessageByString(SendIM3%, WM_SETTEXT, 0, Who$)
SendIM5% = FindChildByClass(SendIM1%, "Ate32class")
SendIM6% = GetWindow(SendIM5%, GW_HWNDNEXT)
SendIM7% = SendMessageByString(SendIM6%, WM_SETTEXT, 0, What$)
SendIM8% = FindChildByClass(SendIM1%, "_Oscar_IconBtn")
ClickIcon (SendIM8%)
End Sub

Sub SendIM_2(Who As String, What As String)

Call OpenNewIM_2
SendIM1% = FindWindow("AIM_IMessage", vbNullString)
SendIM2% = FindChildByClass(SendIM1%, "_Oscar_PersistantComb")
SendIM3% = FindChildByClass(SendIM2%, "Edit")
SendIM4% = SendMessageByString(SendIM3%, WM_SETTEXT, 0, Who$)
SendIM5% = FindChildByClass(SendIM1%, "Ate32class")
SendIM6% = GetWindow(SendIM5%, GW_HWNDNEXT)
SendIM7% = SendMessageByString(SendIM6%, WM_SETTEXT, 0, What$)
SendIM8% = FindChildByClass(SendIM1%, "_Oscar_IconBtn")
ClickIcon (SendIM8%)
End Sub

Sub ChangeBuddyCaption(NewCaption As String)

BuddyCaption1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
BuddyCaption2% = SendMessageByString(BuddyCaption1%, WM_SETTEXT, 0, NewCaption$)
End Sub

Sub ChangeIMCaption(NewCaption As String)

IMCaption1% = FindWindow("AIM_IMessage", vbNullString)
IMCaption2% = SendMessageByString(IMCaption1%, WM_SETTEXT, 0, NewCaption$)
End Sub

Sub ChangeChatCaption(NewCaption As String)

ChatCaption1% = FindWindow("AIM_IMessage", vbNullString)
ChatCaption2% = SendMessageByString(ChatCaption1%, WM_SETTEXT, 0, NewCaption$)
End Sub

Sub ChatMacroKill()

ChatSend ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ Genius")
TimeOut 0.75
ChatSend ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@        Genius")
TimeOut 0.75
ChatSend ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ Genius")
End Sub

Sub StopButton()

Do
DoEvents:
Loop
End Sub

Sub HideAd()

Hide1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Hide2% = FindChildByClass(Hide1%, "Ate32Class")
Hide3% = ShowWindow(Hide2%, SW_HIDE)
End Sub

Sub ShowAd()

Show1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Show2% = FindChildByClass(Show1%, "Ate32Class")
Show3% = ShowWindow(Show2%, SW_SHOW)
End Sub

Function FindChildByTitle(Parent, Child As String) As Integer

childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, buffer$, (hwndLength% + 1))

If InStr(UCase(buffer$), UCase(Child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByClass(Parent, Child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, buffer$, 250)

If InStr(UCase(buffer$), UCase(Child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Sub OpenNewIM()

OpenNewIM1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewIM2% = FindChildByClass(OpenNewIM1%, "_Oscar_TabGroup")
OpenNewIM3% = FindChildByClass(OpenNewIM2%, "_Oscar_IconBtn")
ClickIcon (OpenNewIM3%)
End Sub
Sub OpenNewIM_2()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Send &Instant Message" + Chr(9) + "Alt-I")
End Sub

Sub OpenNewChatInvite()

OpenNewChatInvite1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewChatInvite2% = FindChildByClass(OpenNewChatInvite1%, "_Oscar_TabGroup")
OpenNewChatInvite3% = FindChildByClass(OpenNewChatInvite2%, "_Oscar_IconBtn")
OpenNewChatInvite4% = GetWindow(OpenNewChatInvite3%, GW_HWNDNEXT)
ClickIcon (OpenNewChatInvite4%)
End Sub

Sub OpenNewChatInvite_2()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Send &Buddy Chat Invitation" + Chr(9) + "Alt-C")
End Sub

Sub ClickIcon(Icon%)

Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub TimeOut(Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop
End Sub

Sub SendChatInvite(Who$, Message$, ChatName$)
' If the Message and ChatName are "", i.e. blank, it won't
' change the default rooms and such.
Call OpenNewChatInvite
ChatInvite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
ChatInvite2% = FindChildByClass(ChatInvite1%, "Edit")
ChatInvite3% = SendMessageByString(ChatInvite2%, WM_SETTEXT, 0, Who$)
ChatInvite4% = FindChildByTitle(ChatInvite1%, "Join me in this Buddy Chat.")
If Not Message$ = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, Message$)
For ChatInvite5% = 1 To 2
ChatInvite4% = GetWindow(ChatInvite4%, GW_HWNDNEXT)
Next ChatInvite5%
If Not ChatName = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, ChatName$)
ChatInvite6% = FindChildByClass(ChatInvite1%, "_Oscar_IconBtn")
For ChatInvite7% = 1 To 2
ChatInvite6% = GetWindow(ChatInvite6%, GW_HWNDNEXT)
Next ChatInvite7%
ClickIcon (ChatInvite6%)
End Sub


Sub SendChatInvite_2(Who$, Message$, ChatName$)
' If the Message and ChatName are "", i.e. blank, it won't
' change the default rooms and such.
Call OpenNewChatInvite_2
ChatInvite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
ChatInvite2% = FindChildByClass(ChatInvite1%, "Edit")
ChatInvite3% = SendMessageByString(ChatInvite2%, WM_SETTEXT, 0, Who$)
ChatInvite4% = FindChildByTitle(ChatInvite1%, "Join me in this Buddy Chat.")
If Not Message$ = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, Message$)
For ChatInvite5% = 1 To 2
ChatInvite4% = GetWindow(ChatInvite4%, GW_HWNDNEXT)
Next ChatInvite5%
If Not ChatName = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, ChatName$)
ChatInvite6% = FindChildByClass(ChatInvite1%, "_Oscar_IconBtn")
For ChatInvite7% = 1 To 2
ChatInvite6% = GetWindow(ChatInvite6%, GW_HWNDNEXT)
Next ChatInvite7%
ClickIcon (ChatInvite6%)
End Sub

Sub IMPunt(Who As String)
    Punt1$ = MsgBox("When you do this, your Rate Meter Has to be *FULLY* Charged!! Is it?  If its not, and you say Ok, you will cause an error!" + Chr(13) + Chr(10) + "Also, this only punts *3.0*, so, make sure the guys is 3.0!", vbOKCancel + vbQuestion, "Before Punt:")
    If Punt1$ = 2 Then Exit Sub
    Punt2$ = "<H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3>"
Call SendIM(Who$, Punt2$)
TimeOut 1.25
Punt3% = FindWindow("AIM_IMessage", vbNullString)
Punt4% = FindChildByClass(Punt3%, "_Oscar_IconBtn")
Punt5% = FindWindow("#32770", vbNullString)
Punt6% = ShowWindow(Punt3%, SW_MINIMIZE)
If Punt5% Then
    Call SendMessage(Punt5%, WM_CLOSE, 0, 0)
    Call SendMessage(Punt3%, WM_CLOSE, 0, 0)
    MsgBox "The guy you are trying to punt is not online, has his IMs off, or has blocked you off.  Sorry, you cannot punt him.  ;i", vbInformation + vbOKOnly, "Error!"
    Exit Sub
End If
Punt7% = FindChildByClass(Punt3%, "Ate32Class")
Punt8% = GetWindow(Punt7%, GW_HWNDNEXT)
For PuntLoop = 1 To 8
    TimeOut 0.05
    Punt9% = SendMessageByString(Punt8%, WM_SETTEXT, 0, Punt2$)
    TimeOut 0.05
    ClickIcon (Punt4%)
    Punt5% = FindWindow("#32770", vbNullString)
    If Punt5% Then
        Call SendMessage(Punt5%, WM_CLOSE, 0, 0)
        Call SendMessage(Punt3%, WM_CLOSE, 0, 0)
        MsgBox "Great job 0wning him!  ;i", vbInformation + vbOKOnly, "Hell yeah!"
        Exit Sub
        End If
Next PuntLoop
Do
    TimeOut 4
    Punt10% = SendMessageByString(Punt8%, WM_SETTEXT, 0, Punt2$)
    TimeOut 0.1
    ClickIcon (Punt4%)
    Punt5% = FindWindow("#32770", vbNullString)
    If Punt5% Then
        Call SendMessage(Punt5%, WM_CLOSE, 0, 0)
        Call SendMessage(Punt3%, WM_CLOSE, 0, 0)
        MsgBox "Great job 0wning him!  ;i", vbInformation + vbOKOnly, "Hell yeah!"
    Exit Sub
    End If
Loop
End Sub
Sub ClearIMText()
ClearIMText1% = FindWindow("AIM_IMessage", vbNullString)
ClearIMText2% = FindChildByClass(ClearIMText1%, "Ate32Class")
ClearIMText3% = SendMessageByString(ClearIMText2%, WM_SETTEXT, 0, "")
End Sub

Sub ClearChatText()
ClearChatText1% = FindWindow("AIM_ChatWnd", vbNullString)
ClearChatText2% = FindChildByClass(ClearChatText1%, "Ate32Class")
ClearChatText3% = SendMessageByString(ClearChatText2%, WM_SETTEXT, 0, "")
End Sub

Function AIMGetUserSn()
If IsUserOnline = False Then Exit Sub
On Error Resume Next
UserSn1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
UserSn2% = GetWindowTextLength(UserSn1%)
UserSn3$ = String$(UserSn2%, 0)
UserSn4% = GetWindowText(UserSn1%, UserSn3$, (UserSn2% + 1))
If Not Right(UserSn3$, 13) = "'s Buddy List" Then Exit Function
UserSn5$ = Mid$(UserSn3$, 1, (UserSn2% - 13))
AIMGetUserSn = UserSn5$

End Function

Function AIMGetSnFromIM()

On Error Resume Next
SnFromIM1% = FindWindow("AIM_IMessage", vbNullString)
SnFromIM2% = GetWindowTextLength(SnFromIM1%)
SnFromIM3$ = String$(SnFromIM2%, 0)
SnFromIM4% = GetWindowText(SnFromIM1%, SnFromIM3$, (SnFromIM2% + 1))
If Not Right(SnFromIM3$, 18) = " - Instant Message" Then Exit Function
SnFromIM5$ = Mid$(SnFromIM3$, 1, (SnFromIM2% - 18))
AIMGetSnFromIM = SnFromIM5$

End Function
Sub AIMMassIM(List As ListBox, Text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call SendIM(List.Text, Text)
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub

Function ChatLink(Link$, Name$)
ChatLink = "<A HREF= " + Chr(34) + Link$ + Chr(34) + ">" + Name$ = "</A>"
End Function

Sub Attention(Text As String)
ChatSend ("<-_-_-_-| Attention! |-_-_-_->")
TimeOut 0.5
ChatSend (Text$)
TimeOut 0.5
ChatSend ("<-_-_-_-| Attention! |-_-_-_->")
End Sub

Sub RunMenuByString(AppHwnd As Long, MenuName As String)
AppToSearch% = GetMenu(AppHwnd)
MenuAmnt% = GetMenuItemCount(AppToSearch%)

For StringFind = 0 To MenuAmnt% - 1
ToSearchSub% = GetSubMenu(AppToSearch%, StringFind)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(MenuName)) Then
MenuItem% = SubCount%
GoTo StringMatch:
End If

Next GetString

Next StringFind
StringMatch:
RunTheMenu% = SendMessage(AppHwnd, WM_COMMAND, MenuItem%, 0)
End Sub

Public Sub StayOnTop(hWindow As Long, bTopMost As Boolean)
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
 
Function IsUserOnline()
' Returns a True if Online, False if Not Online.
IsUserOnline1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
If IsUserOnline1% <> 0 Then
IsUserOnline = True
Else
IsUserOnline = False
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
