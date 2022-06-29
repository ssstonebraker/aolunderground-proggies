Attribute VB_Name = "VK_AOL"

'this version  9-1-98

'coded by KRhyME and SkaFia
'|¯|    |¯|\¯\ |¯|    |_|/¯//¯/|¯|\¯\ /¯/|¯|
'| |/¯/ | |/ / | |\¯\   / / | ||_|| ||  |/_/_
'|_|\_\ |_|\_\ |_| |_| /_/  |_|   |_| \______|
'HEAD of the [Voltron Kru]
'Voltron Kru '98
'www.voltronkru.com
'voltronkru@juno.com

'This Bas file requires Voltron.bas to work
'Voltron.Bas is the core bas file for the
'Voltron Kru. You can get all our bas files at
'www.voltronkru.com

'many ideas for this bas came from other Voltron Kru
'members. I would Like to thank SkaFia for all the
'things he did for the series of Bas files.
'Please do not steal our codes without giving us
'credit. I would like to say thank you to KnK for
'making so many files avaible to the public, The makers
'of DiVe32.bas (the first bas i used), Toast, Magus,
'and all the other great programmers out there who
'have infuinced us

'Please join our VB mailing list
'www.voltronkru.com

Function AOLHyperLink(ByVal text As String, ByVal link As String)
'Text: STRING - What you want the hyperlink to say
'Link: STRING - The keyword/link for the hyperlink to be linked to
'RETURN VALUE: STRING - A string ready to be put in an IM or Mail window
AOLHyperLink = "<HTML><A HREF=""" & link & """>" & text & "</A></HTML>"
End Function




Sub AOL30_Set_Guest_SN_PW(SN, PW)
'This finds the guest signon window
'and will set the SN and PW
'and press enter

Do
DoEvents
modal1 = FindWindow("_AOL_Modal", vbNullString)
thetext1 = FindChildByClass(modal1, "_AOL_Edit")
If modal1 <> 0 And thetext1 <> 0 Then Exit Do
Loop

Call SetText(thetext1, SN)
thetext2 = GetWindow(thetext1, 2)
thetext3 = GetWindow(thetext2, 2)
Call SetText(thetext3, PW)
cancelbut = FindChildByClass(modal1, "_AOL_button")
okbut% = GetWindow(cancelbut, 2)
Call Click(okbut%)
End Sub

Public Sub AOL30_AddRoom(the_AOL_list As Long, listbox_to_add_to As ListBox)
'This will add the room to a list
'use the name of the listbox only
'EXAMPLE
'thelist& = FindChildByClass(AOL30_FindChatRoom, "_AOL_Listbox")
'Call AddRoom(thelist, List1)

chatcount = SendMessage(the_AOL_list, LB_GETCOUNT, 0, 0)
For i = 0 To chatcount - 1
DoEvents
sname = GetListSpecific(the_AOL_list, i)
If sname <> GetUser Then
For X = 0 To listbox_to_add_to.ListCount - 1
DoEvents
If listbox_to_add_to.List(X) = sname Then GoTo gohere
Next
If sname <> "" Then listbox_to_add_to.AddItem (sname)
gohere:
End If
Next

End Sub

Function AOL30_ChatRoom()
GetMDI
aoc% = GetWindow(MDi, GW_CHILD)
aoc% = GetWindow(aoc%, GW_hWndFIRST)
aon% = aoc%
Do
    ce% = FindChildByClass(aon%, "_AOL_Edit")
    cv% = FindChildByClass(aon%, "_AOL_View")
    cl% = FindChildByClass(aon%, "_AOL_Listbox")
    If ce% <> 0 And cv% <> 0 And cl% <> 0 Then GoTo 1
    aon% = GetWindow(aon%, GW_hWndNEXT)
Loop Until aon% = aoc%
AOL30_ChatRoom = 0
Exit Function
1:
AOL30_ChatRoom = aon%
End Function

Function AOL30_CheckOnline()
'This will see if AOL is signed ON

AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
welcome = findchildbytitle(MDi, "Welcome, ")
If welcome <> 0 Then
AOL30_CheckOnline = 0
Exit Function
End If
AOL30_CheckOnline = 1
End Function

Public Function AOL30_FindChatRoom()
GetAOL
GetMDI
windws = 0
Chat = FindChildByClass(MDi, "AOL CHILD")
LB = FindChildByClass(Chat, "_AOL_LISTBOX")
EB = FindChildByClass(Chat, "_AOL_EDIT")
If EB <> 0 And LB <> 0 Then
AOL30_FindChatRoom = Chat
Exit Function
End If
Do
Chat = GetWindow(Chat, GW_hWndNEXT)
LB = FindChildByClass(Chat, "_AOL_LISTBOX")
EB = FindChildByClass(Chat, "_AOL_EDIT")
If EB <> 0 And LB <> 0 Then
AOL30_FindChatRoom = Chat
Exit Function
End If
windws = windws + 1
Loop While windws < 100
Chat = 0

End Function

Function AOL30_GetChatText()
'This will get the entire chat window
'see AOL30_GetLastChatLine for getting just the
'last line
childs% = AOL30_FindChatRoom()
child = FindChildByClass(childs%, "_AOL_View")
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)
theview$ = trimspace$
AOL30_GetChatText = theview$
End Function

Public Function AOL30_GetIMSender() As String
'This will get the person sending the IM
GetMDI
im% = FindChildByTitlePartial(MDi, ">Instant Message From:")
If im% Then
imlength = GetWindowTextLength(im%)
imtitle$ = Space(imlength)
Call GetWindowText(im%, imtitle$, (imlength + 1))
finper = Mid(imtitle$, 23)
AOL30_GetIMSender = finper
End If
GetMDI
im% = findchildbytitle(MDi, "  Instant Message From:")
If im% Then
End If
End Function

Function AOL30_GetLastChatLine()
'This will get the last line in the chat
'room

getpar = AOL30_FindChatRoom()
child = FindChildByClass(getpar, "_AOL_View")
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)

theview$ = trimspace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
If thechars$ = "" Then GoTo bad
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
If lastline <> "" Then
AOL30_GetLastChatLine = lastline
Else
bad:
AOL30_GetLastChatLine = " "
End If
End Function

Public Function AOL30_GetPercent(hwndwin As Long) As String
'Hehe this was used with AOL_Modal to
'get the percent from a upload or download

welcomelength = GetWindowTextLength(hwndwin)
welcometitle$ = Space(welcomelength)
wintext = GetWindowText(hwndwin, welcometitle$, (welcomelength + 1))
finper = Mid(welcometitle$, 17)
AOL30_GetPercent = finper
End Function

Function AOL30_GetRoomCount()
'This will get the number of users in a
'room

thechild% = AOL30_FindChatRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOL30_GetRoomCount = getcount
End Function

Function AOL30_GetUser()
'This Will get the USER
On Error Resume Next
AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
welcome = FindChildByTitlePartial(MDi, "Welcome, ")
welcomelength = GetWindowTextLength(welcome)
welcometitle$ = String$(200, 0)
a = GetWindowText(welcome, welcometitle$, (welcomelength + 1))
User = Mid$(welcometitle$, 10, (InStr(welcometitle$, "!") - 10))
AOL30_GetUser = User
End Function

Sub AOL30_IMsOff()
'This Turns IM's off
Call AOL30_SendInstantMessage("$IM_OFF", "Voltron Kru")
End Sub

Sub AOL30_IMsOn()
'This Turns IM's on
Call AOL30_SendInstantMessage("$IM_ON", "Voltron Kru")
End Sub

Sub AOL30_OpenMailSpecific(which)
'If you pass in the number 1, it will open
'your newmail, number2 your old mail, any
'other number your mail you've read

If which = 1 Then
Call RunMenuByString("Read &New Mail")
End If

If which = 2 Then
Call RunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call RunMenuByString("Check Mail You've &Sent")
End If

End Sub

Sub AOL30_RespondIM(message)
'this will respond to a open IM
GetMDI
im% = findchildbytitle(MDi, ">Instant Message From:")
If im% Then GoTo Z
im% = findchildbytitle(MDi, "  Instant Message From:")
If im% Then GoTo Z
Exit Sub
Z:
E = FindChildByClass(im%, "RICHCNTL")
For i = 1 To 9
DoEvents
E = GetWindow(E, 2)
Next
E2 = GetWindow(E, 2) 'Send Text
E = GetWindow(E2, 2) 'Send Button
Call SetText(E2, message)
ClickIcon (E)
End Sub

Public Sub Click_Button(but)
'This will click on a "_AOL_Button"

Call sendmessagebynum(but, WM_KEYDOWN, VK_SPACE, 0)
Call sendmessagebynum(but, WM_KEYUP, VK_SPACE, 0)
End Sub

Public Sub Clicker(Handle)
a% = sendmessagebynum(Handle, WM_LBUTTONDOWN, 0, 0)
B% = sendmessagebynum(Handle, WM_LBUTTONUP, 0, 0)
End Sub

Sub ClickIcon(icon%)
'This will click on a "_AOL_Icon"

Clicka% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clicka% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub GetAOL()
'This Will return AOLs handle and AOL
'is a public var it can be used every
'where.
'Example
'GetAOL
AOL = FindWindow("AOL Frame25", vbNullString)
End Sub

Public Sub GetMDI()
'This Will return the MDI handle and MDI
'is a public var it can be used every
'where.
GetAOL
MDi = FindChildByClass(AOL, "MDIClient")
End Sub

Sub AOL30_SendChat(txt)
'This will send text to the chat room
room% = AOL30_FindChatRoom()
Call SetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
End Sub

Function AOL30_SendInstantMessage(Person As String, message As String)
GetMDI
RunMenuByString ("Send an Instant Message")
Do: DoEvents
im = findchildbytitle(MDi, "Send Instant Message")
imrich = FindChildByClass(im, "RICHCNTL")
imtext = FindChildByClass(im, "_AOL_Static")
imicon = FindChildByClass(im, "_AOL_Icon")
If im <> 0 And imrich <> 0 And imtext <> 0 And imicon <> 0 Then Exit Do
Loop
imedit = GetWindow(imtext, 2)
For i = 1 To 8
DoEvents
imicon = GetWindow(imicon, 2)
Next
Call SetText(imedit, Person)
Call SetText(imrich, message)
imicon = FindChildByClass(im, "_AOL_Icon")
For i = 1 To 9
DoEvents
imicon = GetWindow(imicon, 2)
Next
ClickIcon (imicon)
Do: DoEvents
im = findchildbytitle(MDi, "Send Instant Message")
aolcl = FindWindow("#32770", "America Online")
If aolcl <> 0 Then closer = SendMessage(aolcl, WM_CLOSE, 0, 0): closer2 = SendMessage(im, WM_CLOSE, 0, 0): Exit Do
If im = 0 Then Exit Do
Loop
End Function

Sub RunMenuByString(stringer As String)
'This will run the string from AOLs menu
'that u enter

'GetAOL
Dim AOL
AOL = FindWindow("AOL Frame25", vbNullString)
Call RMBS(AOL, stringer)
End Sub

Sub AOL30_SendKeyword(text)
'This will goto to a keyword

Call RunMenuByString("keyword...")
Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
keyw = findchildbytitle(MDi, "Keyword")
kedit = FindChildByClass(keyw, "_AOL_Edit")
If kedit Then Exit Do
Loop

editsend = sendmessagebystring(kedit, WM_SETTEXT, 0, text)
pausing = DoEvents()
Sending = SendMessage(kedit, 258, 13, 0)
pausing = DoEvents()
End Sub

Sub AOL30_SendMail(Person, SUBJECT, message)
'This will send a EMail

Call RunMenuByString("Compose Mail")

Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDi, "Compose Mail")
icone = FindChildByClass(MailWin, "_AOL_Icon")
peepz = FindChildByClass(MailWin, "_AOL_Edit")
subjt = findchildbytitle(MailWin, "Subject:")
subjec = GetWindow(subjt, 2)
mess = FindChildByClass(MailWin, "RICHCNTL")
If icone <> 0 And peepz <> 0 And subjec <> 0 And mess <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz, WM_SETTEXT, 0, Person)
a = sendmessagebystring(subjec, WM_SETTEXT, 0, SUBJECT)
a = sendmessagebystring(mess, WM_SETTEXT, 0, message)

ClickIcon (icone)


Do: DoEvents
ClickIcon (icone)
AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDi, "Compose Mail")
erro = findchildbytitle(MDi, "Error")
aolw = FindWindow("#32770", "America Online")
If MailWin = 0 Then Exit Do
If aolw <> 0 Then
a = SendMessage(aolw, WM_CLOSE, 0, 0)
a = SendMessage(MailWin, WM_CLOSE, 0, 0)
Exit Do
End If
If erro <> 0 Then
a = SendMessage(erro, WM_CLOSE, 0, 0)
a = SendMessage(MailWin, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub AOL30_SendMail2(Person, SUBJECT, message)
Call RunMenuByString("Compose Mail")

Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDi = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDi, "Compose Mail")
icone = FindChildByClass(MailWin, "_AOL_Icon")
peepz = FindChildByClass(MailWin, "_AOL_Edit")
subjt = findchildbytitle(MailWin, "Subject:")
subjec = GetWindow(subjt, 2)
mess = FindChildByClass(MailWin, "RICHCNTL")
If icone <> 0 And peepz <> 0 And subjec <> 0 And mess <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz, WM_SETTEXT, 0, Person)
a = sendmessagebystring(subjec, WM_SETTEXT, 0, SUBJECT)
a = sendmessagebystring(mess, WM_SETTEXT, 0, message)

ClickIcon (icone)
ClickIcon (icone)

End Sub

Sub AOL30_SetBackPre()
GetMDI
Call RunMenuByString("Preferences")
Do: DoEvents
prefer% = findchildbytitle(MDi, "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_hWndNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
ClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Closewindows% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And Closewindows% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(Closewindows%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

Click_Button (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Sub AOL30_SetPreference()
GetMDI
Call RunMenuByString("Preferences")

Do: DoEvents
prefer% = findchildbytitle(MDi, "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_hWndNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
ClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Closewindows% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And Closewindows% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(Closewindows%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

Click_Button (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Sub AOL30_GetMemberProfile(name As String)
'This gets the profile of member "name"
RunMenuByString ("Get a Member's Profile")
timeout 0.3
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
prof% = findchildbytitle(MDi%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call SetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
Click_Button okbutton%
End Sub



Sub AOL30_SignOff()
'This will sign u off
RunMenuByString ("Sign Off")
End Sub

Public Sub AOL30_WaitForOk()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
DoEvents
Loop Until okb <> 0
Do
okw = FindWindow("#32770", "America Online")
    okb = findchildbytitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
DoEvents
Loop Until okw = 0

End Sub


Function AOL30_LastChatLineWithSN()
chattext$ = AOL30_GetChatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(chattext$, lastlen, Len(thechars$))

AOL30_LastChatLineWithSN = lastline
End Function

Function AOL30_SNFromLastChatLine()
chattext$ = AOL30_LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
AOL30_SNFromLastChatLine = SN
End Function



Sub AOL4_ChatSend(txt)
'Sendz txt to a chat room
    room% = AOL4_FindRoom()
    If room% Then
        hChatEdit% = FindChildByClass(room%, "RICHCNTL")
        ret = sendmessagebystring(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = sendmessagebynum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub


Sub AOL4_SendChat(txt)
'Sendz txt to a chat room
    room% = AOL4_FindRoom()
    If room% Then
        hChatEdit% = FindChildByClass(room%, "RICHCNTL")
        ret = sendmessagebystring(hChatEdit%, WM_SETTEXT, 0, txt)
        ret = sendmessagebynum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub

Function AOL4_FindRoom()
'Findz the chat room/setz focus on it
    AOL% = FindWindow("AOL Frame25", 0&)
    MDi% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDi%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
        AOL4_FindRoom = firs%
        Exit Function
    End If
    
    AOLFindRoom = 0
End Function

Function AOL4_GetText(child)
'This getz text from "Child" Window
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)
AOL4_GetText = trimspace$
End Function

Sub AOL4_Invite(Person)
'This will send an Invite to a person
On Error GoTo errhandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
bud% = findchildbytitle(MDi%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
Call Clicker(E)
timeout (1)
Chaty% = findchildbytitle(MDi%, "Buddy Chat")
aoledit% = FindChildByClass(Chaty%, "_AOL_Edit")
If Chaty% Then GoTo FILL
FILL:
Call SetText(aoledit%, Person)
de = FindChildByClass(Chaty%, "_AOL_Icon")
Call Clicker(de)
Killit% = findchildbytitle(MDi%, "Invitation From:")
Call CloseWindow(Killit%)
errhandler:
Exit Sub
End Sub

Sub AOL4_Keyword(txt)
'This goes to an AOL Keyword
    AOL% = FindWindow("AOL Frame25", 0&)
    Temp% = FindChildByClass(AOL%, "AOL Toolbar")
    Temp% = FindChildByClass(Temp%, "_AOL_Toolbar")
    Temp% = FindChildByClass(Temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(Temp%, "Edit")
    d = sendmessagebystring(KWBox%, WM_SETTEXT, 0, txt)
    E = sendmessagebynum(KWBox%, WM_CHAR, VK_SPACE, 0)
    F = sendmessagebynum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Function AOL4_MessageFromIM()
'This gets the Message from an IM
GetMDI
im% = findchildbytitle(MDi, ">Instant Message From:")
If im% Then GoTo DAMN
im% = findchildbytitle(MDi, "  Instant Message From:")
If im% Then GoTo DAMN
Exit Function
DAMN:
imtext = FindChildByClass(im%, "RICHCNTL")
Imessage = GetText(imtext)
FUK$ = IMmessage
Naw$ = Mid(FUK$, InStr(FUK$, ":") + 2)
AOL4_MessageFromIM = Naw$
End Function

Function AOL4_SNfromIM()
'This will return the Screen Name from an IM
GetMDI
im% = findchildbytitle(MDi, ">Instant Message From:")
If im% Then GoTo Greed
im% = findchildbytitle(MDi, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
heh$ = GetText(im%)
Naw2$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL4_SNfromIM = Naw2$
End Function

Function AOLVersion()
'this wil find version of aol is being used
'returns 25 if version 2.5
'        4  if version 4
'        3  if version 3
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL ToolBar")
car% = FindChildByClass(tool%, "_AOL_ComboBox")
Welc% = findchildbytitle(AOL%, "Welcome,")
rch% = FindChildByClass(Welc%, "RICHCNTL")
If rch% = 0 Then AOLVersion = 25: Exit Function
If car% <> 0 Then AOLVersion = 4: Exit Function
If rch% <> 0 Then AOLVersion = 3: Exit Function

End Function

Sub AOL4_KillGlyph()
'This will close that little annoying AOL spinning
'thingy on the top corner of AOL 4.0
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolbar%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Sub AOL4_locateMember(SN)
'This will locate where a member is
Call AOL4_Keyword("aol://3548:" & SN)
End Sub
Sub AOL30_SignON()
'make sure the signon window is set to guest
'for this code to work properly
'it will click the signon button, use
'AOL30_Guest_enter after you call this code
'to enter the SN, an PW
'this code was written for use in a cracker
'or a quick signon....will not set signon window
'to guest

Do: DoEvents
aol1 = FindWindow("AOL Frame25", vbNullString)
mdi1 = FindChildByClass(aol1, "MDICLIENT")
modal1 = FindWindow("_AOL_Modal", vbNullString)

goodbye = FindChildByTitlePartial(mdi1, "Goodbye ")
goodbut1 = FindChildByClass(goodbye, "_AOL_Icon")
goodbut2 = GetWindow(goodbut1, 2)
goodbut3 = GetWindow(goodbut2, 2)

welcomewindo = findchildbytitle(mdi1, "WELCOME")
button1% = FindChildByClass(welcomewindo, "_AOL_Icon")
BUTTON2% = GetWindow(button1%, 2)
SignONr% = GetWindow(BUTTON2%, 2)
   
   If goodbye <> 0 And goodbut3 <> 0 Then
      Call timeout(2)
      Call Clicker(goodbut3)
      Exit Do
   End If

   If aol1 <> 0 And welcomewindo <> 0 And button1 <> 0 Then
      Call Clicker(SignONr%)
      Exit Do
   End If

Loop

End Sub
Sub AOL30_Guest_enter(name, PW)
'this will enter then name and PW into
'the guest signon window
modal1 = FindWindow("_AOL_Modal", vbNullString)

thetext1 = FindChildByClass(modal1, "_AOL_Edit")
Call SetText(thetext1, name)
thetext2 = GetWindow(thetext1, 2)
thetext3 = GetWindow(thetext2, 2)
Call SetText(thetext3, PW)
End Sub

Sub AOL4_Kill_DL_advertise()
'kill download advertisement
home% = findchildbytitle(AOLMDI, "File Transfer")
DL% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(DL%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_Kill_Mail_Advertise()
'kill mail advertisement
mail% = findchildbytitle(AOLMDI, AOLUserSN & "'s Online Mailbox")
Add% = FindChildByClass(mail%, "_AOL_Image")
Call SendMessage(Add%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_AntiIdle()
'keep form getting signed off
modal% = FindWindow("_AOL_Modal", vbNullString)
icon% = FindChildByClass(modal%, "_AOL_Icon")
ClickIcon (icon%)
End Sub
Sub AOL4_Kill_Chat_Advertise()
'Kills the annoying advertisemenat in member chats.
Chati% = AOL4_FindRoom()

pict% = FindChildByClass(Chati%, "_AOL_Image")
Call SendMessage(pict%, WM_CLOSE, 0, 0)
End Sub
Sub AOL4_SignOff()
'This will sign-off AOL very quickly
Call RunMenuByString("&Sign Off")
End Sub

Sub AOL4_ShowToolBar()
'show the tool bar
DeLTa& = FindWindow("AOL Frame25", vbNullString)
SocK& = FindChildByClass(DeLTa&, "AOL Toolbar")
PLoP& = ShowWindow(SocK&, 5)
End Sub
Function AOL30_IMScan()
'call this after sending an im...
'checks their status
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
im% = findchildbytitle(MDi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs OFF and can't be punted."
End If
If aolcl% = 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
im% = findchildbytitle(MDi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs ON and can be punted."
End If
End Function

Sub AOL4_KillWait()
'makes the hourglass on AOL turn back to the normal
'mouse icon (the arrow)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDi% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = findchildbytitle(MDi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_UnUpChat()
'stop upchatting
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub
Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function
Sub AOL4_SetFocus()
'set focus to aol
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Function AOLWindow()
'gets the aol window
AOL = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL
End Function

Function AOL30_SNfromIM()
'This will return the Screen Name from an IM
GetMDI
im% = findchildbytitle(MDi, ">Instant Message From:")
If im% Then GoTo Greed
im% = findchildbytitle(MDi, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
heh$ = GetText(im%)
Naw2$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL30_SNfromIM = Naw2$
End Function

Sub AOL4_IMOff()
'ims off
Call AOL4_InstantMessage("$IM_OFF", "Voltron Kru Owns Me")
End Sub
Sub AOL4_IMOn()
'ims on
Call AOL4_InstantMessage("$IM_ON", "Voltron Kru Owns Me")
End Sub
Sub AOL4_InstantMessage(Person, message)
'sends an im
Call AOL4_Keyword("aol://9293:" & Person)
timeout (2)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
im% = findchildbytitle(MDi%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SetText(aolrich%, message)
For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_hWndNEXT)
Next sends
Call Click(imsend%)
If im% Then Call CloseWindow(im%)
End Sub


Sub AOL4_Mail(Person, SUBJECT, message)
'sends mail
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(Tool2%, "_AOL_Icon")
Icon2% = GetWindow(ico3n%, 2)
X = sendmessagebynum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
X = sendmessagebynum(Icon2%, WM_LBUTTONUP, 0&, 0&)
Call timeout(4)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDi% = FindChildByClass(AOL%, "MDIClient")
    mail% = findchildbytitle(MDi%, "Write Mail")
    aoledit% = FindChildByClass(mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(mail%, "RICHCNTL")
    subjt% = findchildbytitle(mail%, "Subject:")
    subjec% = GetWindow(subjt%, 2)
        Call SetText(aoledit%, Person)
        Call SetText(subjec%, SUBJECT)
        Call SetText(aolrich%, message)
E = FindChildByClass(mail%, "_AOL_Icon")
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
E = GetWindow(E, GW_hWndNEXT)
Call Clicker(E)
End Sub
Public Sub AOL4_MassIM(lst As ListBox, txt As TextBox)
'easy mass im
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For X = 0 To i
lst.ListIndex = X
Call AOL4_InstantMessage(lst.text, txt.text)
timeout (1)
Next X
lst.Enabled = True
End Sub
Sub AOL4_OpenChat()
'go to chat
AOL4_Keyword ("PC")
End Sub
Sub AOL4_OpenPR(PRrm As TextBox)
'go to a private room
Call AOL4_Keyword("aol://2719:2-2-" & PRrm)
End Sub

Sub IMIgnore(thelist As ListBox)
'ignores im's from people on the list
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDi% = FindChildByClass(AOL%, "MDIClient")
im% = findchildbytitle(MDi%, ">Instant Message From:")
If im% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(AOL30_SNfromIM) Then
            BadIM% = im%
            imrich% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(imrich%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If

Loop
End Sub



Sub AOL30_BuddyBLOCK(SN As String)
'blocks some one from adding you to
'their buddylist
'
GetMDI
BUDLIST% = findchildbytitle(MDi, "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_hWndNEXT)
setup% = GetWindow(IM1%, GW_hWndNEXT)
Call Click(setup%)
Call timeout(2)
STUPSCRN% = findchildbytitle(MDi, AOL30_GetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_hWndNEXT)
Delete% = GetWindow(Edit%, GW_hWndNEXT)
View% = GetWindow(Delete%, GW_hWndNEXT)
PRCYPREF% = GetWindow(View%, GW_hWndNEXT)
Call Click(PRCYPREF%)
Call timeout(1.8)
Call CloseWindow(STUPSCRN%)
Call timeout(2)
PRYVCY% = findchildbytitle(MDi, "Privacy Preferences")
DABUT% = findchildbytitle(PRYVCY%, "Block only those people whose screen names I list")
Call Click(DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call SetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Edit% = GetWindow(Edit%, GW_hWndNEXT)
Call Click(Edit%)
Call timeout(1)
Save% = GetWindow(Edit%, GW_hWndNEXT)
Save% = GetWindow(Save%, GW_hWndNEXT)
Save% = GetWindow(Save%, GW_hWndNEXT)
Call Click(Save%)
End Sub
Function AOLClickList(hWnd)
'clicks a list
ClickList% = sendmessagebynum(hWnd, &H203, 0, 0&)
End Function

Function AOL30_CountMail()
'counts your mail
GetMDI
themail% = FindChildByClass(MDi, "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOL30_CountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function


Sub AOL30_SetFocus()
'aol gets focus
X = GetCaption(AOLWindow)
AppActivate X
End Sub

Sub AOL30_SetToGuest()
'this moves the aol signon window to guest
X = GetCaption(AOLWindow)
AppActivate X
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.2)
SendKeys "{down}"
Call timeout(0.3)
SendKeys "{up}"
Call timeout(0.2)
End Sub


Sub AOL30_SignOnASGuest(name As String, password As String)
'written 100% by KRhyME and SkaFia
'this will let you sign on to aol
'as a guest....
'(it even moves the name to guest for you)
'
Call RunMenuByString("set up && sign on")

Do: DoEvents
aol1 = FindWindow("AOL Frame25", vbNullString)
mdi1 = FindChildByClass(aol1, "MDICLIENT")
modal1 = FindWindow("_AOL_Modal", vbNullString)

goodbye = FindChildByTitlePartial(mdi1, "Goodbye ")
goodbut1 = FindChildByClass(goodbye, "_AOL_Icon")
goodbut2 = GetWindow(goodbut1, 2)
goodbut3 = GetWindow(goodbut2, 2)

welcomewindo = findchildbytitle(mdi1, "WELCOME")


Call RunMenuByString("Set Up && Sign On")


button1% = FindChildByClass(welcomewindo, "_AOL_Icon")
BUTTON2% = GetWindow(button1%, 2)
SignONr% = GetWindow(BUTTON2%, 2)
Call AOL30_SetToGuest
   If goodbye <> 0 And goodbut3 <> 0 Then
      Call timeout(2)
      Call Clicker(goodbut3)
      Exit Do
   End If

   If aol1 <> 0 And welcomewindo <> 0 And button1 <> 0 Then
      Call Clicker(SignONr%)
      Exit Do
   End If

Loop
Call AOL30_Set_Guest_SN_PW(name, password)
End Sub

Function AOL30_MessageFromIM()
GetMDI
'This gets the Message from an IM
im = findchildbytitle(MDi, ">Instant Message From:")
If im Then GoTo DAMN
im = findchildbytitle(MDi, "  Instant Message From:")
If im Then GoTo DAMN
Exit Function
DAMN:
imtext = FindChildByClass(im, "RICHCNTL")
IMmessage = GetText(imtext)
FUK$ = IMmessage
Naw$ = Mid(FUK$, InStr(FUK$, ":") + 2)
AOL30_MessageFromIM = Naw$
End Function

Function AOL30_TestForMASTER()
'this code written 100% by KRhyME and SkaFia
'checks an account to see it its a
'master or a sub
'returns 0 for master
'returns 1 for sub

'If AOL30_TestForMASTER = 0 Then
'     MsgBox "master account"
'  Else
'     MsgBox "sub account"
'  End If
  
Call AOL30_SendKeyword("billing")
Do: DoEvents
AOLe = FindWindow("AOL Frame25", vbNullString)
MDIe = FindChildByClass(AOLe, "MDICLIENT")

pwnoa = FindWindow("_AOL_Modal", vbNullString)
pwnoa2% = FindChildByClass(pwnoa, "_AOL_Button")

billing = findchildbytitle(MDIe, "Accounts & Billing")
Advertise = FindChildByClass(billing, "_AOL_Image")
but1% = FindChildByClass(billing, "_AOL_Icon")
but2% = GetWindow(but1%, 2)
but3% = GetWindow(but2%, 2)
but4% = GetWindow(but3%, 2)
but5% = GetWindow(but4%, 2)
but6% = GetWindow(but5%, 2)
but7% = GetWindow(but6%, 2)

If pwnoa <> 0 And pwnoa2% <> 0 Then GoTo blocked:
If billing <> 0 And but1% <> 0 And but7% <> 0 Then GoTo Goody:
Loop


Goody:
    Call CloseWindow(Advertise)
    Call timeout(1)
    Call Click(but7%)
    Do: DoEvents
    
    pwver = findchildbytitle(MDIe, "Password Verification")
    pwno = FindWindow("#32770", "America Online")
    If pwver <> 0 Or pwno <> 0 Then Exit Do
    Loop

         If pwver <> 0 Then
         AOL30_TestForMASTER = 0
         Call CloseWindow(pwver)
         Call CloseWindow(billing)
         Exit Function
         End If

         If pwno <> 0 Then
         AOL30_TestForMASTER = 1
         pwno2% = findchildbytitle(pwno, "OK")
         Call Click(pwno2%)
         Call CloseWindow(billing)
         Exit Function
         End If

blocked:
    AOL30_TestForMASTER = 1
    pwnoa2% = FindChildByClass(pwnoa, "_AOL_Button")
    Call Click(pwnoa2%)
    Call CloseWindow(billing)
    Exit Function
    
End Function

Public Function GetListSpecific(ByVal hChatList As Long, ByVal nListIdx As Integer) As String
'this will get names 1 at a time from a
'from a _Aol_listbox

    '/Setup error handling
    On Error GoTo Err_GetListSpecific
    
    Dim hAOLProcess As Long   ' A handle of AOL's process.
    
    Dim lAddrOfItemData As Long ' Memory location of a list items item data.
    Dim sBuffer As String       ' Buffer to place data read by ReadProcessMemory. NOTE: This will only work if a string
                                ' buffer is used...I don't know why though, but would like to so if you have an idea I
                                ' be be happy to hear it.
    Dim lAddrOfName As Long     ' Memory location of the screen name.
    Dim lBytesRead As Long      ' Number of bytes read by ReadProcessMemory.
    
    
    ' Make sure 0 wasn't passed to this function.
    If hChatList Then
        ' Get a valid process handle for AOL that enables you to read(PROCESS_VM_READ) memory in it's process space.
        hAOLProcess = GetAOLProcessHandle(hChatList)
            
        ' Make sure a handle was retrieved
        If hAOLProcess Then
        
            ' Setup the buffer
            sBuffer = String$(4, vbNullChar)
            
            ' Get the item data for the list item
            lAddrOfItemData = SendMessage(hChatList, LB_GETITEMDATA, ByVal CLng(nListIdx%), ByVal 0&)
                            
            ' lAddrOfItemData is actually a pointer to a 7 element array of 4byte values...reading begins from
            ' the address of the 7th element.  Since lAddrOfItemData already equals the address of the first element
            ' you add 4 for each element, so the 7th element's memory address is...
            lAddrOfItemData = lAddrOfItemData + (4 * 6)
            
            ' lAddrOfItemData is now the address of the 7th element of the array, this element contains a 4byte pointer
            ' to a string(the screen name)
            Call ReadProcessMemory(hAOLProcess, lAddrOfItemData, sBuffer, 4, lBytesRead)
                        
            ' The 4 bytes in sBuffer are actually a pointer, this pointer needs to be incremented by 6 so sBuffer needs
            ' to be convertd to long value
            RtlMoveMemory lAddrOfName, ByVal sBuffer, 4
            
            ' Increment the address
            lAddrOfName = lAddrOfName + 6
            
            ' Setup buffer
            sBuffer = String$(16, vbNullChar)
            
            ' lAddrOfName now holds a pointer to a string(screen name), so retrieve the string
            Call ReadProcessMemory(hAOLProcess, lAddrOfName, sBuffer, Len(sBuffer), lBytesRead)
            
            ' That's it, so add the screen name to the array, but be sure to trim off any extra characters.
            GetListSpecific = Left$(sBuffer$, InStr(sBuffer$, vbNullChar) - 1)
            
            ' Close the handle to AOL's process
            Call CloseHandle(hAOLProcess)
        End If
    End If
    
    Exit Function

'/Error handler
Err_GetListSpecific:
    ' Make sure the handle to AOL's process is closed
    Call CloseHandle(hAOLProcess)
    
    Exit Function
    

End Function

Private Function GetAOLProcessHandle(ByVal hWnd As Long) As Long
    
    '/Setup error handling
    On Error Resume Next
    
    Dim m_AOLThreadID As Long   ' A value that uniquely identifies the thread throughout the system.
    Dim m_AOLProcessID As Long  ' A value that uniquely identifies the process throughout the system.
    
    ' Get the process ID for AOL's main thread. Since AOL is not a multithreaded application each window use the same
    ' thread.
    m_AOLThreadID = GetWindowThreadProcessId(hWnd, m_AOLProcessID)
    
    ' Get a valid process handle for AOL that enables you to read(PROCESS_VM_READ) memory in it's process space.
    GetAOLProcessHandle = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, m_AOLProcessID)
                
End Function

Public Sub AOL30_AskToBeAdded(trigger As String, box As ListBox)
'steps...
'1. in a Timer Call aol30_AskToBeAdded
'2. make 2 command buttons
'3. in command button 1 type-
'Timer.enbled = True
'AOL30_ChatSend "Type / (what ever your trigger is) to be added to (what ever)"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOL30_ChatSend "Bot Off!"

'set the timers interval to 1

FreeProcess
On Error Resume Next
Dim last As String
Dim name As String
Dim a As String
Dim n As Integer
Dim X As Integer
DoEvents
a = AOL30_GetLastChatLine
last = Len(a)
For X = 1 To last
name = Mid(a, X, 1)
final = final & name
If name = ":" Then Exit For
Next X
final = Left(final, Len(final) - 1)
'If final = AOL30_GetUser Then
'Exit Sub
'Else
If InStr(a, "/" + trigger) Then
box.AddItem (name)
Call AOL30_SendChat(Chr(62) + Chr(126) + Chr(126) + Chr(58) + Chr(32) + name + " was added" + Chr(32) + Chr(32) + Chr(58) + Chr(126) + Chr(126) + Chr(60))
Call timeout(0.6)
'End If
End If
End Sub

