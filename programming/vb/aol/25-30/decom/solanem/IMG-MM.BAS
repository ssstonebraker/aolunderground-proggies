'''''''''''

'''''''''

'''''''
'''''''
'''''''

Global FirstMail As Integer
Global LastMail As Integer

Sub AnswerIMs ()
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Do While IM% <> 0
IM% = findchildbytitle(MDI%, ">Instant Message From:")
If IM% = 0 Then IM% = findchildbytitle(MDI%, "  Instant Message From:")
IMCap$ = GrabText(IM%): DoEvents
SN$ = Trim$(Mid$(IMCap$, InStr(IMCap$, ":")))
Call SendIM(Trim$(SN$), "Sorry, but I'm currently massmailing with SoLuTioN massmailer.  Please E-Mail me or talk to me later. Thanks!"): DoEvents
closewin IM%: DoEvents
Loop
End Sub

Sub ChatSay (who$, wut$)
AOL% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(AOL%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = sendmessagebystring(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (who$) & ":" & Chr(9) & (wut$))
SendNow% = sendmessagebynum(ChatVew%, WM_CHAR, &HD, 0)
End Sub

Sub MailPref ()
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call runmenubystring(AOL%, "Preferences")

Do: DoEvents
prefer% = findchildbytitle(MDI%, "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = getwindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Call Click(mailbut%, 1)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, True, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, False, 0)

Call Click(aolOK%, 1)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function Percent (Complete As Integer, total As Integer, TotalOutput As Integer) As Integer
R = Int(Complete / total * TotalOutput)
Percent = R
End Function

Function Replace (object$, RepWhat$, RepWit$) As String
Yay:
If InStr(UCase$(object$), UCase$(RepWhat$)) Then
        LeftSide$ = Left$(object$, InStr(UCase$(object$), UCase$(RepWhat$)) - 1)
        RightSide$ = Mid$(object$, InStr(UCase$(object$), UCase$(RepWhat$)) + Len(RepWhat$), Len(object$))
        Finish$ = LeftSide$ & RepWit$ & RightSide$
    Else :
        Finish$ = object$
End If
Replace = Finish$
End Function

Sub SendAll ()

Dim EndMM As Integer

AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Tree% = FindChildByClass(MDI%, "_AOL_TREE")
mailbox% = getparent(Tree%): DoEvents
AOLVersion = AOLv()
    
    If InStr(GrabText(mailbox%), "Mail") = False Then MsgBox "Unable to locate your mailbox, please try again", 32: Exit Sub

For f = 0 To 5
frmMMer.Hide
Next f
frmStatus.Show

nCount = SendMessage(Tree%, LB_GETCOUNT, 0, 0&): DoEvents

Call MailPref: DoEvents

If frmOptions.chkStartatBottom.Value = False Then x% = SendMessage(Tree%, LB_SETCURSEL, 0, 0): DoEvents
If frmOptions.chkStartatBottom.Value = True Then x% = SendMessage(Tree%, LB_SETCURSEL, nCount - 1, 0&): DoEvents

startTime = Int(Timer)

nxtMail:

Call AnswerIMs: DoEvents

If frmStatus.cmdPause.Caption = "Resume" Then
    Do: DoEvents
    Loop Until frmStatus.cmdPause.Caption = "Pause"
End If

If frmStatus.cmdStop.Caption = "Ending..." Then
    EndMM = True
End If

mailStart = Int(Timer)


Do: DoEvents
    fwdlbl% = findchildbytitle(MDI%, "Forward")
        tolbl% = findchildbytitle(MDI%, "To: ")

Call closewin(getparent(fwdlbl%)): DoEvents
Call closewin(getparent(tolbl%)): DoEvents
Loop Until fwdlbl% = 0 And tolbl% = 0


If GrabText(mailbox%) = "Incoming FlashMail" Then
    readbtn% = findchildbytitle(mailbox%, "Open")
        Else
    readbtn% = findchildbytitle(mailbox%, "Read")
End If
Click readbtn%, 1: DoEvents
    For j% = 0 To 15
        Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
    Next j%
Do: DoEvents
    msg% = FindWindow("#32770", 0&)
    fwdlbl% = findchildbytitle(MDI%, "Forward")
    Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
Loop Until fwdlbl% <> 0 Or msg% <> 0
    
    If msg% <> 0 Then
        Call waitforok: DoEvents
            GoTo nxtMail
            Exit Sub
    End If

fwdlbl% = findchildbytitle(MDI%, "Forward")
fwdparent% = getparent(fwdlbl%): DoEvents
fwdbtn% = getnextwindow(fwdlbl%, 2): DoEvents
Click fwdbtn%, 1: DoEvents
    Do: DoEvents
        tolbl% = findchildbytitle(MDI%, "To: ")
    Loop Until tolbl% <> 0
tolbl% = findchildbytitle(MDI%, "To: ")
mailwnd% = getparent(tolbl%): DoEvents
tolbl% = findchildbytitle(mailwnd%, "To: ")
totxt% = getnextwindow(tolbl%, 2): DoEvents
    '/* RESETING SCREEN NAMES HERE */
resetsnall:
SN$ = ""
    '/* END OF SNz */
    For j% = 0 To frmMMer.lstPeople.ListCount - 1
        SN$ = SN$ & "(" & frmMMer.lstPeople.List(j%) & "), ": DoEvents
    Next j%
sendtext totxt%, SN$: DoEvents
subjlbl% = findchildbytitle(mailwnd%, "Subject: ")
subjtxt% = getnextwindow(subjlbl%, 2): DoEvents
fwdtxt$ = GrabText(subjtxt%): DoEvents
    If UCase(Left$(fwdtxt$, 5)) = UCase("FWD: ") Then
        fwdtxt$ = Mid$(fwdtxt$, 6): DoEvents
        sendtext subjtxt%, fwdtxt$: DoEvents
    End If

If AOLVersion = 25 Then
    msgtxt% = getnextwindow(subjtxt%, 2)
        Else
    msgtxt% = FindChildByClass(mailwnd%, "RICHCNTL")
End If

mailFinish = Int(Timer)

MailSecond = Int(((mailFinish - startTime)) / 60 * 100)
MailMinute = 0
MailHour = 0

If MailSecond >= 60 Then
    MailMinute = Int(MailSecond / 60)
    MailSecond = Int(((mailFinish - startTime) / 60 * 100) - MailMinute * 60)
End If

sendtext msgtxt%, "···• SoLuTioN Mailing System By Glen" & Chr(13) & Chr(10) & "···• Sending Mail: " & Trim$(Str(nSel + 1)) & " of " & Trim$(Str(nCount)) & Chr(13) & Chr(10) & "···• This Mail Took; " & Trim$(Str(mailFinish - mailStart)) & " Second[s]" & Chr(13) & Chr(10) & "···• This Mailing has lasted for; " & Trim$(Str(MailMinute)) & " Minute[s] and " & Trim$(Str(MailSecond)) & " Second[s]": DoEvents

sendlbl% = findchildbytitle(mailwnd%, "Send Now")
sendbtn% = getnextwindow(sendlbl%, 2): DoEvents
Click sendbtn%, 1: DoEvents
    Do: DoEvents
        dead% = FindWindow("#32770", 0&)
        sendlbl% = findchildbytitle(mailwnd%, "Send Now")
        errmsg% = findchildbytitle(MDI%, "Error")
    Loop Until sendlbl% = 0 Or errmsg% <> 0 Or dead% <> 0

If dead% <> 0 Then
    closewin mailwnd%: DoEvents
    GoTo nolbl
    Exit Sub
End If


If errmsg% <> 0 Then
            errmsg% = findchildbytitle(MDI%, "Error")
            errview% = FindChildByClass(errmsg%, "_AOL_VIEW")
    errtext$ = GrabText(errview%)
    errtext$ = Trim$(errtext$)
        hyphen = InStr(errtext$, ":")
    blah = Mid$(errtext$, hyphen + 5)
    dash = InStr(blah, "-")
    errtexta$ = Mid$(blah, 1, dash - 2)
    okbtn% = FindChildByClass(errmsg%, "_AOL_BUTTON")
        Do: DoEvents
        Call Click(okbtn%, 1): DoEvents
            errmsg% = findchildbytitle(MDI%, "Error")
        Loop Until errmsg% = 0
    For f = 0 To 5
        Call removeitemlst(frmMMer.lstPeople, Trim$(errtexta$)): DoEvents
        Call addlist(frmStatus.lstRemoved, Trim$(errtexta$)): DoEvents
    Next f
        
        If frmOptions.chkIMFull.Value = True Then Call SendIM(errtexta$, "I'm sorry but you have been removed off of the SoLuTioN massmail because of a mailbox error."): DoEvents
    GoTo resetsnall
End If
nolbl:
If sendlbl% = 0 Then
    closewin fwdparent%: DoEvents
    
    If EndMM = True Then
    Unload frmStatus
    frmMMer.Show
        MsgBox "Mass Mail ended early", 32
    Exit Sub
    End If

    If nSel + 1 = nCount Then
    Unload frmStatus
    frmMMer.Show
        MsgBox "Mass Mail is complete", 32
    Exit Sub
    Else
        mCount% = nCount
        frmStatus.pnlPercent.FloodPercent = Percent(nSel + 1, mCount%, 100)

            nSel = SendMessage(Tree%, LB_GETCURSEL, 0&, 0&): DoEvents

            If frmOptions.chkStartatBottom.Value = False Then x% = SendMessage(Tree%, LB_SETCURSEL, nSel + 1, 0&): DoEvents
            If frmOptions.chkStartatBottom.Value = True Then x% = SendMessage(Tree%, LB_SETCURSEL, nSel - 1, 0&): DoEvents

        GoTo nxtMail
    End If
End If
End Sub

Sub SendCurrent ()
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Tree% = FindChildByClass(MDI%, "_AOL_TREE")
mailbox% = getparent(Tree%): DoEvents
AOLVersion = AOLv()
    
    If InStr(GrabText(mailbox%), "Mail") = False Then MsgBox "Unable to locate your mailbox, please try again", 32: Exit Sub

Call MailPref: DoEvents
Do: DoEvents
    fwdlbl% = findchildbytitle(MDI%, "Forward")
        tolbl% = findchildbytitle(MDI%, "To: ")

Call closewin(getparent(fwdlbl%)): DoEvents
Call closewin(getparent(tolbl%)): DoEvents
Loop Until fwdlbl% = 0 And tolbl% = 0


If GrabText(mailbox%) = "Incoming FlashMail" Then
    readbtn% = findchildbytitle(mailbox%, "Open")
        Else
    readbtn% = findchildbytitle(mailbox%, "Read")
End If
Click readbtn%, 1: DoEvents
    For x = 0 To 15
        Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
    Next x
Do: DoEvents
    msg% = FindWindow("#32770", 0&)
    fwdlbl% = findchildbytitle(MDI%, "Forward")
    Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
Loop Until fwdlbl% <> 0 Or msg% <> 0
    
    If msg% <> 0 Then
        Call waitforok: DoEvents
            MsgBox "Sorry, the current mail is unavailable, and it could not be sent", 32
            Exit Sub
    End If

    

fwdlbl% = findchildbytitle(MDI%, "Forward")
fwdparent% = getparent(fwdlbl%): DoEvents
fwdbtn% = getnextwindow(fwdlbl%, 2): DoEvents
Click fwdbtn%, 1: DoEvents
    Do: DoEvents
        tolbl% = findchildbytitle(MDI%, "To: ")
    Loop Until tolbl% <> 0
tolbl% = findchildbytitle(MDI%, "To: ")
mailwnd% = getparent(tolbl%): DoEvents
tolbl% = findchildbytitle(mailwnd%, "To: ")
totxt% = getnextwindow(tolbl%, 2): DoEvents
    '/* RESETING SCREEN NAMES HERE */
resetsncurrent:
SN$ = ""
    '/* END OF SNz */
    For j% = 0 To frmMMer.lstPeople.ListCount - 1
        SN$ = SN$ & "(" & frmMMer.lstPeople.List(j%) & "), ": DoEvents
    Next j%
sendtext totxt%, SN$: DoEvents
subjlbl% = findchildbytitle(mailwnd%, "Subject: ")
subjtxt% = getnextwindow(subjlbl%, 2): DoEvents
fwdtxt$ = GrabText(subjtxt%): DoEvents
    If UCase(Left$(fwdtxt$, 5)) = UCase("FWD: ") Then
        fwdtxt$ = Mid$(fwdtxt$, 6): DoEvents
        sendtext subjtxt%, fwdtxt$: DoEvents
    End If

If AOLVersion = 25 Then
    msgtxt% = getnextwindow(subjtxt%, 2)
        Else
    msgtxt% = FindChildByClass(mailwnd%, "RICHCNTL")
End If

sendtext msgtxt%, "···• SoLuTioN Mailing System By Glen": DoEvents

sendlbl% = findchildbytitle(mailwnd%, "Send Now")
sendbtn% = getnextwindow(sendlbl%, 2): DoEvents
Click sendbtn%, 1: DoEvents
    Do: DoEvents
        sendlbl% = findchildbytitle(mailwnd%, "Send Now")
        errmsg% = findchildbytitle(MDI%, "Error")
    Loop Until sendlbl% = 0 Or errmsg% <> 0

If errmsg% <> 0 Then
            errmsg% = findchildbytitle(MDI%, "Error")
            errview% = FindChildByClass(errmsg%, "_AOL_VIEW")
    errtext$ = GrabText(errview%)
    errtext$ = Trim$(errtext$)
        hyphen = InStr(errtext$, ":")
    blah = Mid$(errtext$, hyphen + 5)
    dash = InStr(blah, "-")
    errtexta$ = Mid$(blah, 1, dash - 2)
    okbtn% = FindChildByClass(errmsg%, "_AOL_BUTTON")
        Do: DoEvents
        Call Click(okbtn%, 2): DoEvents
            errmsg% = findchildbytitle(MDI%, "Error")
        Loop Until errmsg% = 0
    For f = 0 To 5
        Call removeitemlst(frmMMer.lstPeople, Trim$(errtexta$)): DoEvents
            Call addlist(frmStatus.lstRemoved, Trim$(errtexta$)): DoEvents
    Next f

        If frmOptions.chkIMFull.Value = True Then Call SendIM(errtexta$, "I'm sorry but you have been removed off of the SoLuTioN massmail because of a mailbox error."): DoEvents
    GoTo resetsncurrent
End If

If sendlbl% = 0 Then
    closewin fwdparent%: DoEvents
    Rem UNLOAD STATUS FORM HERE
    MsgBox "Mass Mail of One Mail is Complete", 32
    Exit Sub
End If
'End If
End Sub

Sub SendRange ()

Dim EndMM As Integer

AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Tree% = FindChildByClass(MDI%, "_AOL_TREE")
mailbox% = getparent(Tree%): DoEvents
AOLVersion = AOLv()
    
    If InStr(GrabText(mailbox%), "Mail") = False Then MsgBox "Unable to locate your mailbox, please try again", 32: Exit Sub

For f = 0 To 5
frmMMer.Hide
Next f
frmStatus.Show

Rem nCount = SendMessage(Tree%, LB_GETCOUNT, 0, 0&): DoEvents
nCount = LastMail% - FirstMail% + 1

Call MailPref: DoEvents

x% = SendMessage(Tree%, LB_SETCURSEL, FirstMail% - 1, 0&): DoEvents

startTime = Int(Timer)

nxtMailr:

Call AnswerIMs: DoEvents

If frmStatus.cmdPause.Caption = "Resume" Then
    Do: DoEvents
    Loop Until frmStatus.cmdPause.Caption = "Pause"
End If

If frmStatus.cmdStop.Caption = "Ending..." Then
    EndMM = True
End If

mailStart = Int(Timer)

nSel = SendMessage(Tree%, LB_GETCURSEL, 0&, 0&): DoEvents
x% = SendMessage(Tree%, LB_SETCURSEL, nSel + 1, 0&): DoEvents

Do: DoEvents
    fwdlbl% = findchildbytitle(MDI%, "Forward")
        tolbl% = findchildbytitle(MDI%, "To: ")

Call closewin(getparent(fwdlbl%)): DoEvents
Call closewin(getparent(tolbl%)): DoEvents
Loop Until fwdlbl% = 0 And tolbl% = 0


If GrabText(mailbox%) = "Incoming FlashMail" Then
    readbtn% = findchildbytitle(mailbox%, "Open")
        Else
    readbtn% = findchildbytitle(mailbox%, "Read")
End If
Click readbtn%, 1: DoEvents
    For j% = 0 To 15
        Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
    Next j%
Do: DoEvents
    msg% = FindWindow("#32770", 0&)
    fwdlbl% = findchildbytitle(MDI%, "Forward")
    Call runmenubystring(AOL%, "Stop Incoming Text"): DoEvents
Loop Until fwdlbl% <> 0 Or msg% <> 0
    
    If msg% <> 0 Then
        Call waitforok: DoEvents
            GoTo nxtMailr
            Exit Sub
    End If

fwdlbl% = findchildbytitle(MDI%, "Forward")
fwdparent% = getparent(fwdlbl%): DoEvents
fwdbtn% = getnextwindow(fwdlbl%, 2): DoEvents
Click fwdbtn%, 1: DoEvents
    Do: DoEvents
        tolbl% = findchildbytitle(MDI%, "To: ")
    Loop Until tolbl% <> 0
tolbl% = findchildbytitle(MDI%, "To: ")
mailwnd% = getparent(tolbl%): DoEvents
tolbl% = findchildbytitle(mailwnd%, "To: ")
totxt% = getnextwindow(tolbl%, 2): DoEvents
    '/* RESETING SCREEN NAMES HERE */
resetsnrange:
SN$ = ""
    '/* END OF SNz */
    For j% = 0 To frmMMer.lstPeople.ListCount - 1
        SN$ = SN$ & "(" & frmMMer.lstPeople.List(j%) & "), ": DoEvents
    Next j%
sendtext totxt%, SN$: DoEvents
subjlbl% = findchildbytitle(mailwnd%, "Subject: ")
subjtxt% = getnextwindow(subjlbl%, 2): DoEvents
fwdtxt$ = GrabText(subjtxt%): DoEvents
    If UCase(Left$(fwdtxt$, 5)) = UCase("FWD: ") Then
        fwdtxt$ = Mid$(fwdtxt$, 6): DoEvents
        sendtext subjtxt%, fwdtxt$: DoEvents
    End If

If AOLVersion = 25 Then
    msgtxt% = getnextwindow(subjtxt%, 2)
        Else
    msgtxt% = FindChildByClass(mailwnd%, "RICHCNTL")
End If

mailFinish = Int(Timer)

MailSecond = Int(((mailFinish - startTime)) / 60 * 100)
MailMinute = 0
MailHour = 0

If MailSecond >= 60 Then
    MailMinute = Int(MailSecond / 60)
    MailSecond = Int(((mailFinish - startTime) / 60 * 100) - MailMinute * 60)
End If

sendtext msgtxt%, "···• SoLuTioN Mailing System By Glen" & Chr(13) & Chr(10) & "···• Sending Mail: " & Trim$(Str(nSel + 2 - FirstMail%)) & " of " & Trim$(Str(nCount)) & Chr(13) & Chr(10) & "···• This Mail Took; " & Trim$(Str(mailFinish - mailStart)) & " Second[s]" & Chr(13) & Chr(10) & "···• This Mailing has lasted for; " & Trim$(Str(MailMinute)) & " Minute[s] and " & Trim$(Str(MailSecond)) & " Second[s]": DoEvents

sendlbl% = findchildbytitle(mailwnd%, "Send Now")
sendbtn% = getnextwindow(sendlbl%, 2): DoEvents
Click sendbtn%, 1: DoEvents
    Do: DoEvents
        dead% = FindWindow("#32770", 0&)
        sendlbl% = findchildbytitle(mailwnd%, "Send Now")
        errmsg% = findchildbytitle(MDI%, "Error")
    Loop Until sendlbl% = 0 Or errmsg% <> 0

If dead% <> 0 Then
    closewin mailwnd%: DoEvents
    GoTo nolblr
    Exit Sub
End If

If errmsg% <> 0 Then
            errmsg% = findchildbytitle(MDI%, "Error")
            errview% = FindChildByClass(errmsg%, "_AOL_VIEW")
    errtext$ = GrabText(errview%)
    errtext$ = Trim$(errtext$)
        hyphen = InStr(errtext$, ":")
    blah = Mid$(errtext$, hyphen + 5)
    dash = InStr(blah, "-")
    errtexta$ = Mid$(blah, 1, dash - 2)
    okbtn% = FindChildByClass(errmsg%, "_AOL_BUTTON")
        Do: DoEvents
        Call Click(okbtn%, 1): DoEvents
            errmsg% = findchildbytitle(MDI%, "Error")
        Loop Until errmsg% = 0
    
    For f = 0 To 5
        Call removeitemlst(frmMMer.lstPeople, Trim$(errtexta$)): DoEvents
        Call addlist(frmStatus.lstRemoved, Trim$(errtexta$)): DoEvents
    Next f

        If frmOptions.chkIMFull.Value = True Then Call SendIM(errtexta$, "I'm sorry but you have been removed off of the SoLuTioN massmail because of a mailbox error."): DoEvents
    GoTo resetsnrange
End If

nolblr:
If sendlbl% = 0 Then
    closewin fwdparent%: DoEvents
    
    If EndMM = True Then
    Unload frmStatus
    frmMMer.Show
        MsgBox "Mass Mail ended early", 32
    Exit Sub
    End If

    If nSel + 1 = LastMail% Then
    Unload frmStatus
    frmMMer.Show
        MsgBox "Mass Mail is complete", 32
    Exit Sub
    Else
        GoTo nxtMailr
    End If
End If
End Sub

