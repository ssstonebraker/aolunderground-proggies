Attribute VB_Name = "NewPsyche"
'Psyche's last bas.

'All you biznitches takin from Raid need to stop!
'Write your own code, stop trailing!  Its not wo-
'rth copying if everyone can tell anyways.


'This was fully written by me, no stealing.
'First Subs taken from my Frubal~4:20 module.

'Greetz to all these leeto`s:
'FuBu, Sketch, e01j, kloned, mushy, nero
'________________________________________
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'Public Const for use.
Public Const WM_CLOSE = &H10
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETTEXT = &HD
Public Const WM_LBUTTONDBLCLK = &H203

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6

Public Const LB_GETCOUNT = &H18B
Public Const LB_SETCURSEL = &H186

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Dim bError As Boolean
Dim bInvite As Boolean
Dim bInfo As Boolean
Dim bProf As Boolean
Dim bOnline As Boolean
Dim bMessage As Boolean
Dim bChat As Boolean
Dim sMessage As String
Dim lWindow As Long


Public Sub Win_Pause(Duration As Long)
    Dim strCurrent As String
    strCurrent = Timer
    Do While Timer - strCurrent < Val(Duration)
    DoEvents
    Loop 'StrCurrent
End Sub

Function Win_GetCap(ByVal lWinHandle As Long) As String
'This gets the title of any window.
    Dim lWinLength As Long, stCap As String
    lWinLength = GetWindowTextLength(lWinHandle)
    stCap$ = String(lWinLength&, 0&)
    Call GetWindowText(lWinHandle&, stCap$, lWinLength& + 1)
    Win_GetCap$ = stCap$
End Function

Function Win_GetTxt(ByVal lWinHandle As Long) As String
'This gets the text of any object
    Dim tWinLength As Long, stTxt As String
    tWinLength = SendMessage(lWinHandle, WM_GETTEXTLENGTH, 0&, 0&)
    stTxt$ = String(tWinLength&, 0&)
    Call SendMessageByString(lWinHandle&, WM_GETTEXT, tWinLength& + 1, stTxt$)
    Win_GetTxt$ = stTxt$
End Function

Public Sub Win_SetTxt(ByVal lWinHandle As Long, ByVal stText As String)
'This writes text to any object
    If Len(stText) > 0 Then
    Call SendMessageByString(lWinHandle, WM_SETTEXT, 0, stText)
    End If
End Sub

Public Sub Win_Click(ByVal lWinHandle As Long)
'This clicks any command button
    DoEvents
    Call SendMessage(lWinHandle, WM_LBUTTONDOWN, 0, vbNullString)
    Call SendMessage(lWinHandle, WM_LBUTTONUP, 0, vbNullString)
    DoEvents
End Sub

Public Sub Win_Close(ByVal lWinHandle As Long)
'This will kill a window
    Call PostMessage(lWinHandle, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Win_Hide(ByVal lWinHandle As Long)
'This will hide any instance of a window
    DoEvents:
    Call ShowWindow(lWinHandle, SW_HIDE)
End Sub

Public Sub Win_Show(ByVal lWinHandle As Long)
'This will show any instance of a window
    DoEvents:
    Call ShowWindow(lWinHandle, SW_SHOW)
End Sub






'Now to the AIM subs and functions!
'I'm going to start off Finding All
'The different windows on Aim.

Function FindMain() As Long
'Main is your buddy list window.
    DoEvents:
    lMain& = FindWindow("_Oscar_BuddyListWin", vbNullString)

    If lMain& = 0 Then bOnline = False Else: bOnline = True
     FindMain& = lMain&
End Function

Function FindIM() As Long
'Will find first Focused IM.
    DoEvents:
    lMess& = FindWindow("AIM_IMessage", vbNullString)

    If lMess& = 0 Then bChat = False Else: bChat = True
     FindIM& = lMess&

End Function

Function FindInfo() As Long
'Will find Info on a person.
    DoEvents:
    lInfo& = FindWindow("_Oscar_Locate", vbNullString)
    
    If lInfo& = 0 Then bInfo = False Else: bInfo = True
     FindInfo& = lInfo&
End Function

Function FindProfile() As Long
'Will find profile
    DoEvents:
    lProfile& = FindWindow("#32770", "Create a Profile - More Info")
    
    If lProfile& = 0 Then bProf = False Else: bProf = True
     FindProfile& = lProfile&
End Function

Function FindInvite() As Long
'Will find Invite  window.
    DoEvents:
    lInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")

    If lInvite& = 0 Then bInvite = False Else: bInvite = True
     FindInvite& = lInvite&
End Function

Function FindChat() As Long
'Will find focused Chat
    DoEvents:
    lChat& = FindWindow("AIM_ChatWnd", vbNullString)
    
    If lChat& = 0 Then bChat = False Else: bChat = True
     FindChat& = lChat&
End Function


Function FindError() As Long
'Will find focused error
    DoEvents:
    lError& = FindWindow("#32770", vbNullString)
    
    If lError& = 0 Then bError = False Else: bError = True
    AimFindError& = lError&
End Function




'These are all for changing
'AIM caps.    Please!  This
'Will cause  errors in many
'Other subs, so use w/ caution!

Public Sub CapMain(sText As String)
    If Len(sText$) > 0 Then
    Call Win_SetTxt(FindMain&, sText$)
    Else 'Set to nothing.
    Call Win_SetTxt(FindMain&, " ")
    End If 'sText$
End Sub

Public Sub CapIm(sText As String)
    If Len(sText$) > 0 Then
    Call Win_SetTxt(FindIM&, sText$)
    Else 'Set to nothing
    Call Win_SetTxt(FindIM&, " ")
    End If 'sText$
End Sub

Public Sub CapChat(sText As String)
    If Len(sText$) > 0 Then
    Call Win_SetTxt(FindChat&, sText$)
    Else 'Set to nothing
    Call Win_SetTxt(FindChat&, " ")
    End If
End Sub




'These are all for closing
'Aim windows.  If the word
'Is plural, it closes them
'All, if its singular it -
'Closes the focused one.

Public Sub CloseMain()
'Close main if open.
    DoEvents:
    lChat& = FindMain&
    If lChat& > 0 Then 'Close
    Call Win_Close(lChat&)
    End If 'lChat&
End Sub

Public Sub CloseIM()
'Close focused Im if open.
    DoEvents:
    lIm& = FindIM&
    If lIm& > 0 Then 'Close
    Call Win_Close(lIm&)
    End If 'lIm&
End Sub

Public Sub CloseIMs()
'Close ALL Ims open.
    DoEvents:
    If FindIM& > 0 Then
    Do Until FindIM& = 0
    Call Win_Close(FindIM&)
    Loop ' FindIm&
    End If
End Sub

Public Sub CloseChat()
'Close Focused Chat.
    DoEvents:
    lChat& = FindChat&
    If lChat& > 0 Then
    Call Win_Close(lChat&)
    End If 'lChat
End Sub

Public Sub CloseChats()
'Close all Chats.
    DoEvents:
    If FindChat& > 0 Then
    Do Until FindChat& = 0
    Call Win_Close(FindChat&)
    Loop ' FindChat&
    End If
End Sub

Public Sub CloseInvite()
'Close Focused Invite
    DoEvents:
    lInvite& = FindInvite&
    If lInvite& > 0 Then
    Call Win_Close(lInvite&)
    End If 'lInvite&
End Sub




'Now I'm going to work with the
'Im's a bit more.  Find out how
'many are open, get text from -
'them, etc.

Public Sub ImOpen()
'Opens New Im.
    DoEvents:
    lParIm& = FindWindowEx(FindMain&, 0, "_Oscar_TabGroup", vbNullString)
    lHanIm& = FindWindowEx(lParIm&, 0, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lHanIm&)
End Sub

Function ImCount() As Integer
'This counts the amount of ims open.
    Dim imWin As Long, lngInt As Long
    lngInt& = -1
    imWin& = 0
    
    Do: DoEvents
    imWin& = FindWindowEx(0, imWin&, "AIM_IMessage", vbNullString)
    lngInt& = lngInt& + 1
    Loop Until imWin& = 0
    
     ImCount% = lngInt&
End Function

Public Sub ImSend(sWho As String, sMessage As String)
'Send An Im.
    Call ImOpen
    
    'Set sWho$
    lMessage& = FindIM&
    lPareIm& = FindWindowEx(lMessage&, 0, "_Oscar_PersistantCombo", vbNullString)
    lHandIm& = FindWindowEx(lPareIm&, 0, "Edit", vbNullString)
    Call Win_SetTxt(lHandIm&, sWho$)
    
    'Set sMessage$
    lPareIm2& = FindWindowEx(lMessage&, 0, "WndAte32Class", vbNullString)
    lHandIm2& = GetWindow(lPareIm2&, 2)
    Call Win_SetTxt(lHandIm2&, sMessage$)
    
    'Click Send.
    lHandIm3& = FindWindowEx(lMessage&, 0, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lHandIm3&)
    
End Sub

Public Sub ImSend2(sWho As String, sMessage As String, bBold As Boolean, bItalic As Boolean, bUnderlined As Boolean, bClose As Boolean)
'This is for those special Ims.
'It`s just a little Fancier.
   If bBold = True Then sMessage$ = "<B>" + sMessage$
   If bItalic = True Then sMessage$ = "<I>" + sMessage$
   If bUnderlined = True Then sMessage$ = "<U>" + sMessage$

   If Len(sWho$) > 0 And Len(sMessage$) > 0 Then
   If Len(sMessage$) > 946 Then
   sMessage$ = Right(sMessage$, 946)
   End If
   
   'send message
   Call ImSend(sWho, sMessage$)
   End If

   If bClose = True Then _
   Call Win_Close(FindIM&)
End Sub

Function Win_Filter(stText As String) As String
'Everybody and their grandma has been copying
'Off Raid's HTML filter!  Well, I did my own.
'It's much better for 32 bit.

'If you copy, please give me credit:
'Written by: Psyche [NewPsyche.bas]
Dim intLoop As Integer
On Error Resume Next

    For intLoop% = 1 To Len(stText)
    stL1& = InStr(1, stText$, ">")
    stR1& = InStr(1, stText$, "<")
    stLeft$ = Left(stText$, stR1& - 1)

    lFull% = (stL1& - (stR1& + 1)) + 2
    stRight$ = Right(stText$, Len(stText$) - (Len(stLeft$) + lFull%))
        
    Filter$ = stLeft$ + stRight$
    stText$ = Filter$
    Next intLoop%


Win_Filter$ = stText$
End Function

Function Win_Replace(stText As String, stToReplace As String, stReplaceWith As String) As String
'This finds any char/word you want in a text and replaces it
'With what you want. I'm going to use if for my second HTML
'Filter

    On Error Resume Next
    For intLoop% = 1 To Len(stText)
        
    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    If lFindChar& = 0 Then
    Win_Replace$ = stText$
    Exit Function
    End If
    

    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    strCharL$ = Left(stText$, lFindChar& - 1)
    strCharR$ = Right(stText$, Len(stText$) - (lFindChar& + Len(stToReplace$) - 1))
    strCharF$ = strCharL$ + stReplaceWith$ + strCharR$
    stText$ = strCharF$
        
    Next intLoop%
    Win_Replace$ = strCharF$
End Function

Function Win_Filter2(stText As String) As String
    stText$ = Win_Replace(stText$, "<BR>", Chr(13))
    stText$ = Win_Replace(stText$, "&nbsp;", " ")
    
    stText$ = Win_Filter(stText$)
    Win_Filter2$ = stText$
End Function

Function ImText() As String
    lImTxt& = FindWindowEx(FindIM&, 0, "WndAte32Class", vbNullString)
    stText1$ = Win_GetTxt(lImTxt&)
    
     ImText$ = Win_Filter2(stText1$)
End Function

Function CountWrds(stText As String) As Integer
'This counts words by spaces.
On Error Resume Next
    If Len(stText) = 0 Then
    CoundWrds% = 0
    Exit Function
    Else
    intWordCount% = 1
    End If

    Dim intLoop As Integer
    For intLoop% = 1 To Len(stText$) Step 1
    stWrd1$ = Mid(stText, intLoop, 1)
    stWrd2$ = Mid(stText, intLoop + 1, 1)
    
    If stWrd1$ = " " Then
        If stWrd2$ = " " Then
        Else    'count another word
        intWordCount% = intWordCount% + 1
        End If
    End If
    Next intLoop%
    
    If Left(stText, 1) = " " Then
    intWordCount% = intWordCount% - 1
    End If
    
    If Right(stText, 1) = " " Then
    intWordCount% = intWordCount% - 1
    End If
    
    If stWordCount% < 0 Then stWordCount% = 0
    CountWrds% = intWordCount%
End Function

Function CountEnters(stText As String) As Integer
'This finds enter by character (chr)   Enter is 13
'chr([number])
    On Error Resume Next
    Dim intLoop As Integer
    For intLoop% = 1 To Len(stText$) Step 1
    stWrd1$ = Mid(stText, intLoop, 1)
    If stWrd1$ = Chr(13) Then intEnters% = intEnters% + 1
    Next intLoop%

    CountEnters% = intEnters%
End Function

Function Win_GetLine(stText As String, iLine As Integer) As String
'This will get one line of a textbox.
'You can use this for every line except the first!

    On Error Resume Next
    If InStr(1, stText, Chr(13)) = 0 Then
    If iLine = 1 Then
    Win_GetLine$ = stText$
    Exit Function
    End If 'iLine
    End If 'InStr
    
    iStop% = 0
    For intLoop3% = 1 To Len(stText$) Step 1
    stLine1$ = Mid(stText$, intLoop3%, 1)
    
    If stLine1$ = Chr(13) Then
    iStop% = iStop% + 1
        If iStop% >= iLine% Then
            Win_GetLine$ = stLine2$
            Exit Function
        End If 'iStop
        stLine1$ = ""
        stLine2$ = ""
        End If
    stLine2$ = stLine2$ + stLine1$
    Next intLoop3%
    

    Win_GetLine$ = stLine2$
End Function

Function CountLines(stText As String) As Integer
'Find out how many lines your textbox has
    On Error Resume Next
    CountLines% = CountEnters(stText$) + 1
    
End Function

Function ImLine(iLine As Integer) As String
'I'm using all those Function
'Above that I wrote.
    stText$ = ImText$
     ImLine$ = Win_GetLine(stText$, iLine%)

End Function

Function ImLastLine() As String
'Use  ImLine to get any line
    'Gets last line of focused IM.
     ImLastLine$ = ImLine(CountLines(ImText$))
    
End Function


Function ImLastText2() As String
'Using  IMline to get Text
'This uses ":" to find it.
'I recommend you use the first one.
    string1$ = ImLastLine$
    For iLoop% = 1 To Len(string1$)
    string2$ = Mid(string1$, iLoop%, 1)
    If string2$ = ":" Then
    string3$ = ""
    Else 'add letter
    string3$ = string3$ + string2$
    End If
    Next iLoop%

string3$ = Right(string3$, Len(string3$) - 1)
 ImLastText2$ = string3$
End Function

Function ImLastName() As String
'Using  IM line to get Name
'This also uses ":" to find it.
    string1$ = ImLastLine$
    lColon& = InStr(1, string1$, ":")
    
     ImLastName$ = Left(string1$, lColon& - 1)
End Function

Function ImLimit() As Boolean
'If  ImLimit is true then you can't send ims
'Else,  ImLimit is false, and you can.

    lAimError1& = FindWindow("#32770", "AOL Instant Messenger (TM)")
    If lAimError1& = 0 Then
     ImLimit = True
    Else 'False
     ImLimit = False
    End If 'lAimError1

End Function

Function ImFromWho() As String
'Gets who sent you the IM.
    On Error Resume Next
    string1$ = Win_GetCap(FindIM&)
    ImFromWho$ = Left(string1$, Len(string1$) - 18)
    
End Function

Public Sub ImReply(stMessage As String)
'Responds to One Focused Im, then closes it.
    On Error Resume Next

    'Find An Open IM!
    If FindIM& <> 0 Then
    strName$ = ImFromWho$
    Call CloseIM

    'Send the IM!
    Call ImSend(strName$, stMessage$)
    strName$ = ""
    End If

End Sub

Public Sub MiniIMs()
'Minimize all IMs.
    
    Dim intLoop As Integer
    For intLoop = 0 To ImCount() Step 1
    im1win& = FindWindowEx(0, 0, "AIM_IMessage", vbNullString)
    IntMin% = ShowWindow(im1win&, SW_MINIMIZE)
    Next intLoop

End Sub

Public Sub ImClear()
'This clears text from an IM.
'strNew is "" to clear fully.
    If Len(strNew$) = 0 Then strNew$ = " "
    lIm001& = FindWindow("AIM_IMessage", vbNullString)
    lIm002& = FindWindowEx(lIm001&, 0, "WndAte32Class", vbNullString)
    Call Win_SetTxt(lIm002&, " ")
    
End Sub




'Thats All the Im subs.
'Now I got A lot of Chat
'Ones to do, so Here we
'Go.


Public Sub ChatSend(stMessage As String)
'Sends string to chat.
    
    If Len(stMessage) > 400 Then
    stMessage2$ = Left(stMessage$, 400)
    Else 'keep same
    stMessage2$ = stMessage$
    End If 'changing len

    'Set Text
    ChatWndow1& = FindWindowEx(FindChat&, 0&, "WndAte32Class", vbNullString)
    ChatWndow2& = FindWindowEx(FindChat&, ChatWndow1&, "WndAte32Class", vbNullString)
    Call Win_SetTxt(ChatWndow2&, stMessage2$)
    
    'Click Send
    lChtHand1& = FindWindowEx(FindChat&, 0, "_Oscar_IconBtn", vbNullString)
    lChtHand2& = FindWindowEx(FindChat&, lChtHand1&, "_Oscar_IconBtn", vbNullString)
    lChtHand3& = FindWindowEx(FindChat&, lChtHand2&, "_Oscar_IconBtn", vbNullString)
    lChatSend& = FindWindowEx(FindChat&, lChtHand3&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lChatSend&)

End Sub

Public Sub ChatSend2(stMessage As String, bBold As Boolean, bUnderlined As Boolean, bItalic As Boolean)
'Just a fancy chat send.
    If bBold = True Then stMessage$ = "<B>" + strName$
    If bItalic = True Then stMessage$ = "<I>" + strName$
    If bUnderlined = True Then stMessage$ = "<U>" + strName$
    Call ChatSend(stMessage$)

End Sub

Public Sub ChatSend3(tbMessage As TextBox)
'This will send according to a textbox.
    Call ChatSend2(tbMessage.Text, tbMessage.FontBold, _
    tbMessage.FontUnderline, tbMessage.FontItalic)

End Sub

Public Sub ChatScroll(iTimes As Integer, lPause As Long, stMessage As String)
'Scrolls To Chat
    On Error Resume Next
    For iLoop% = 1 To iTimes%
    Call ChatSend(stMessage)
    Call Win_Pause(lPause&)
    Next iLoop%
End Sub

Function UserName() As String
'Gets name of user

    string1$ = Win_GetCap(FindMain&)
     UserName$ = Left(string1$, Len(string1$) - 20)
    
End Function

Public Sub Greeting()
'This is pretty cool.  It says Good Morning or Afternoon depending on the time
'Plus the user's name.
    string1$ = Right(Time, 2)
    If string1$ = "PM" Then
    string2$ = "afternoon"
    Else
    string2$ = "morning"
    End If 'string1$
    
    Call ChatSend("Good " + string2$ + ", " + UserName$)
End Sub

Public Sub CloseErrors(lLimit As Long)
'lLimit is 1000 to allow long check.
    
    If FindWindow("#32770", vbNullString) > 0 Then
    Do Until lAimError& = 0 Or lLim& >= lLimit
    lLimit = lLim& + 1
    lAimError& = FindWindow("#32770", vbNullString)
    Call Win_Close(lAimError&)
    Loop 'lAimError&
    End If
End Sub

Public Sub OpenInvite()
'Opens Invite for Chat
    lInvPar& = FindWindowEx(FindMain&, 0, "_Oscar_TabGroup", vbNullString)
    lInvHan& = FindWindowEx(lInvPar&, 0, "_Oscar_IconBtn", vbNullString)
    lInvite& = FindWindowEx(lInvPar&, lInvHan&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lInvite&)
    
End Sub


Function ChatName() As String
'Get name of Chat
    string1$ = Win_GetCap(FindChat&)
    ChatName$ = Right(string1$, Len(string1$) - 11)
    
End Function



Function ChatLink(sLink As String, sCaption As String) As String
'Notice this is a function! Not a sub.
    ahref$ = "<A HREF=" + Chr(34) + sLink$ + Chr(34) + ">"
    ahref$ = ahref$ + sCaption$ + "</A>"
    ChatLink$ = ahref$
End Function

Function ChatText() As String
'Get Full Chat
    lChText& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
     ChatText$ = Win_Filter2(Win_GetTxt(lChText&))
    
End Function

Function ChatLine(iLine As Integer) As String
'Get any line out of the chat.
    string1$ = ChatText$
     ChatLine$ = Win_GetLine(string1$, iLine%)

End Function

Function ChatLastLine() As String
'Gets Lastline
    On Error Resume Next
     ChatLastLine$ = ChatLine(CountLines(ChatText$))
    
End Function

Function ChatLastText() As String
'Gets Text from last line
    string1$ = ChatLastLine$
    lColon& = InStr(1, string1$, ":")
     ChatLastText$ = Right(string1$, Len(string1$) - (lColon& + 1))

End Function

Function ChatLastName() As String
'Gets Name from last line
    string1$ = ChatLastLine$
    lColon& = InStr(1, string1$, ":")
     ChatLastName$ = Left(string1$, lColon& - 1)

End Function

Function ChatLetter(stLetter As String) As Boolean
'This finds the first letter in the last line
'Used for afks   Ex:  If  chatletter(".") = true then msgbox( lastchatline)
    string1$ = Left(ChatLastText$, 1)
    If stLetter$ = string1$ Then
     ChatLetter = True
    Else
     ChatLetter = False
    End If
End Function

Public Sub ChatSendMult(strMessage As String, iPause As Long)
'Send Multiple lines to Chat
    On Error Resume Next
    If InStr(1, strMessage$, Chr(13)) = 0 Then
    Call ChatSend(strMessage$)
    Exit Sub
    End If
    
    For intLoop% = 1 To CountLines(strMessage$)
    strL$ = Win_GetLine(strMessage, intLoop%)
    Call ChatSend(strL$)
    Call Win_Pause(iPause&)
    Next intLoop%
End Sub

Public Sub ChatClear()
'Clear the Chat.
    lChaPar& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    lChPar2& = FindWindowEx(lChaPar&, 0, "Ate32Class", vbNullString)
    Call Win_SetTxt(lChPar2&, " ")
    
End Sub

Public Sub AddRoom(objBox As Object, blnAddYourself As Boolean)
'I always have had trouble with Addrooms, this one works!
'My example:                call  addroom(List1, False)
Dim intLoop As Integer

    lngChatWnd& = FindWindow("Aim_ChatWnd", vbNullString)
    lngAllPeop& = FindWindowEx(lngChatWnd&, 0, "_Oscar_Tree", vbNullString)
    lngGetCoun& = SendMessageByNum(lngAllPeop&, LB_GETCOUNT, 0, 0)

    For intLoop = 0 To Str(lngGetCoun& - 1) Step 1
    DoEvents:
        lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, intLoop, 0)
        lngDblClic& = SendMessageByNum(lngAllPeop&, WM_LBUTTONDBLCLK, 0, 0)
    
        lngGetIm& = FindWindow("AIM_IMessage", vbNullString)
        lngIMPar& = FindWindowEx(lngGetIm&, 0, "_Oscar_PersistantCombo", vbNullString)
        lngIMhan& = FindWindowEx(lngIMPar&, 0, "Edit", vbNullString)
        strName$ = Win_GetTxt(lngIMhan&)
        If blnAddYourself = False Then
        If UserName = strName$ Then
        'Do nothin` at all
        Else
        'Add the ol' feller!
        objBox.AddItem strName$
        End If
        Else
        'If you are adding yourself
        objBox.AddItem strName$
        End If
        
    Call Win_Close(lngGetIm&)
    Next intLoop%
    
End Sub

Function Win_TrimText(strText As String) As String
'This converts any string to no spaces and
'No Caps.
stringM$ = ""

    If InStr(strText, " ") = 0 Then
    stringM$ = strText
    Else
    For intLoop = 1 To Len(strText)
    stringNew$ = Mid(strText, intLoop, 1)
    If stringNew$ = " " Then Else: _
    stringM$ = stringM$ + stringNew$
    Next intLoop
    End If
        
    stringN$ = LCase(stringM$)
    Win_TrimText = stringN$
    
End Function

Function FindChatter(strSName As String) As Boolean
'This Checks to see if a certain Name is in a room.
'Example: If  NameInRoom("um psyche") = true then msgbox("PSYCHE is here!")
Dim intLoop As Integer

    lngChatWnd& = FindWindow("Aim_ChatWnd", vbNullString)
    lngAllPeop& = FindWindowEx(lngChatWnd&, 0, "_Oscar_Tree", vbNullString)
    lngGetCoun& = SendMessageByNum(lngAllPeop&, LB_GETCOUNT, 0, 0)

    For intLoop = 0 To Str(lngGetCoun& - 1) Step 1
    DoEvents:
        lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, intLoop, 0)
        lngDblClic& = SendMessageByNum(lngAllPeop&, WM_LBUTTONDBLCLK, 0, 0)
    
        lngGetIm& = FindWindow("AIM_IMessage", vbNullString)
        lngIMPar& = FindWindowEx(lngGetIm&, 0, "_Oscar_PersistantCombo", vbNullString)
        lngIMhan& = FindWindowEx(lngIMPar&, 0, "Edit", vbNullString)
        strName$ = Win_GetTxt(lngIMhan&)
        If Win_TrimText(strName$) = Win_TrimText(strSName$) Then
         FindChatter = True
        Call Win_Close(lngGetIm&)
        Exit Function
        Else
        'Keep Looking
         FindChatter = False
        End If
        
    Call Win_Close(lngGetIm&)
    Next intLoop%
End Function

Public Sub ChatFilter()
    lChText& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    string1$ = Win_GetTxt(lChText&)
    
    string1$ = Win_Replace(LCase(string1$), "<br>", "!br!")
    string1$ = Win_Filter2(string1$)
    string1$ = Win_Replace(LCase(string1$), "!br!", String$(116, " "))

    lChaPar& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    lChPar2& = FindWindowEx(lChaPar&, 0, "Ate32Class", vbNullString)
    Call Win_SetTxt(lChPar2&, string1$)
End Sub

Public Sub ChatIgnore(iIndex As Integer)
'Ignore by index
    lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, iIndex%, 0)
    lLbHan& = FindWindowEx(FindChat&, 0, "_Oscar_IconBtn", vbNullString)
    lLbHan2& = FindWindowEx(FindChat&, lLbHan&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lLbHan2&)
    

End Sub


Public Sub ImFilter()
'This is good for anti punt on im strings.
'Find out if there is someone trying to punt, then
'call ImFilter and strip out the html.
    lImTxt& = FindWindowEx(FindIM&, 0, "WndAte32Class", vbNullString)
    stText1$ = Win_GetTxt(lImTxt&)
    
    stText1$ = Win_Replace(LCase(stText1$), "<br>", "!br!")
    stText1$ = Win_Filter2(stText1$)
    stText1$ = Win_Replace(LCase(stText1$), "!br!", "<br>")

    lIm002& = FindWindowEx(FindIM&, 0, "WndAte32Class", vbNullString)
    Call Win_SetTxt(lIm002&, stText1$)
End Sub


'This is the Excess Codes for
'everything Else you may need.
'Copy at will!  I dont care
'On these.

Public Sub MainClose()
    DoEvents:
    If FindMain& <> 0 Then _
    Call Win_Close(FindMain&)
End Sub

Public Sub MainOpen(sDirectory As String)
    On Error Resume Next
    DoEvents:
    If aimfindmain& = 0 Then _
    Call Shell(sDirectory$, vbNormalFocus)
End Sub

Public Sub MiniMain()
    DoEvents:
    Call ShowWindow(FindMain&, SW_MINIMIZE)
End Sub

Public Sub Win_OnTop(ByVal lHandle As Long, ByVal bOnTop As Boolean)
'Have any windows handle and set in on top, or
'Take it off top. bOnTop is true if on top.

    If bOnTop = True Then
    Call SetWindowPos(lHandle&, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else 'Not On Top
    Call SetWindowPos(lHandle&, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
    End If
End Sub

Public Sub TopMain(bOnTop As Boolean)
    'Set Main on top.
    Call Win_OnTop(FindMain, bOnTop)
End Sub

Public Sub TopChat(bOnTop As Boolean)
    'Set Chat to top
    Call Win_OnTop(FindChat, bOnTop)
End Sub

Public Sub TopIm(bOnTop As Boolean)
    'Set IM to top.
    Call Win_OnTop(FindIM, bOnTop)
End Sub



Public Sub MiniChat()
    'Minimize Chat
    Call ShowWindow(FindChat&, SW_MINIMIZE)
End Sub


      
Public Sub WriteSN(ByVal SNtxt As String)
'This one's for you.  This allows you to set the the Sign On screen name
    lParSN& = FindWindow("#32770", "Sign On")
    lParent& = FindWindowEx(lParSN&, 0, "ComboBox", vbNullString)
    lHandle& = FindWindowEx(lParent&, 0, "Edit", vbNullString)
    Call Win_SetTxt(lHandle&, SNtxt)

End Sub

Public Sub WritePW(PWtxt As String)
'Writes your password
    On Error Resume Next
    lParent& = FindWindow("#32770", "Sign On")
    lHandle& = FindWindowEx(lParent&, 0, "Edit", vbNullString)
    Call Win_SetTxt(lHandle&, PWtxt)
        
End Sub

Public Sub HideMain(bHide As Boolean)
'Hidden is bHide. Set True to Hide, False to Show.
    If bHide = True Then
    Win_Hide (FindMain&)
    Else 'Show it.
    Win_Show (FindMain&)
    End If
End Sub

Public Sub HideIM(bHide As Boolean)
'Hiddens is bHide. Set True to Hide, False to Show
    If bHide = True Then
    Win_Hide (FindIM&)
    Else 'Show it
    Win_Show (FindIM&)
    End If
End Sub

Public Sub HideChat(bHide As Boolean)
'Hidden is bHide. Set True to Hide, False to Show
    If bHide = True Then
    Win_Hide (FindChat&)
    Else 'show it
    Win_Show (FindChat&)
    End If
End Sub

Function IsOnline() As Boolean
' IsOnline is true if Online
    If FindMain& <> 0 Then IsOnline = True
    If FindMain& = 0 Then IsOnline = False

End Function

Public Sub SayHello(stSayWhat As String)
'Will send [name], + [stSayWhat]
'Example: stSayWhat is "Hows it goin?", so it
'would send to chat "name, hows it goin?"
Dim intLoop As Integer

    lngChatWnd& = FindWindow("Aim_ChatWnd", vbNullString)
    lngAllPeop& = FindWindowEx(lngChatWnd&, 0, "_Oscar_Tree", vbNullString)
    lngGetCoun& = SendMessageByNum(lngAllPeop&, LB_GETCOUNT, 0, 0)

    For intLoop = 0 To Str(lngGetCoun& - 1) Step 1
    DoEvents:
        lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, intLoop, 0)
        lngDblClic& = SendMessageByNum(lngAllPeop&, WM_LBUTTONDBLCLK, 0, 0)
    
        lngGetIm& = FindWindow("AIM_IMessage", vbNullString)
        lngIMPar& = FindWindowEx(lngGetIm&, 0, "_Oscar_PersistantCombo", vbNullString)
        lngIMhan& = FindWindowEx(lngIMPar&, 0, "Edit", vbNullString)
        strName$ = Win_GetTxt(lngIMhan&)
        If UserName = strName$ Then
        'Do nothin` at all
        Else
        'Add the ol' feller!
        ChatSend (strName$ + ", " + stSayWhat)
        End If

        
    Call Win_Close(lngGetIm&)
    Next intLoop%
End Sub

Public Sub SayRoom()
'Scrolls something like: Hey it's 3:30 P.M. and
'We're in Room, Psyche.
    string1$ = Time
    string2$ = ChatName
    Call ChatSend("Hey, it's " + string1$ + " ,and I am in room " + string2$)

End Sub


