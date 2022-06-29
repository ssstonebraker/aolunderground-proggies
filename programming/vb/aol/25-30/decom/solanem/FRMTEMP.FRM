VERSION 2.00
Begin Form frmTemp 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   3000
   Height          =   2160
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   3000
   Top             =   1140
   Width           =   3120
   Begin CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1020
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
End

Sub Command1_Click ()
aol% = findwindow("AOL Frame25", 0&)
mdi% = findchildbyclass(aol%, "MDIClient")
tree% = findchildbyclass(mdi%, "_AOL_TREE")
mailbox% = getparent(tree%): DoEvents
AOLVersion = AOLv()
    
    If InStr(grabtext(mailbox%), "Mail") = False Then MsgBox "Unable to locate your mailbox, please try again", 32: Exit Sub

Call MailPref: DoEvents
Do: DoEvents
    fwdlbl% = findchildbytitle(mdi%, "Forward")
        tolbl% = findchildbytitle(mdi%, "To: ")

Call closewin(getparent(fwdlbl%)): DoEvents
Call closewin(getparent(tolbl%)): DoEvents
Loop Until fwdlbl% = 0 And tolbl% = 0


If grabtext(mailbox%) = "Incoming FlashMail" Then
    readbtn% = findchildbytitle(mailbox%, "Open")
        Else
    readbtn% = findchildbytitle(mailbox%, "Read")
End If
Click readbtn%, 1: DoEvents
    For X = 0 To 15
        Call runmenubystring(aol%, "Stop Incoming Text"): DoEvents
    Next X
Do: DoEvents
    msg% = findwindow("#32770", 0&)
    fwdlbl% = findchildbytitle(mdi%, "Forward")
    Call runmenubystring(aol%, "Stop Incoming Text"): DoEvents
Loop Until fwdlbl% <> 0 Or msg% <> 0
    
    If msg% <> 0 Then
        Call waitforok: DoEvents
            MsgBox "Sorry, the current mail is unavailable, and it could not be sent", 32
            Exit Sub
    End If

fwdlbl% = findchildbytitle(mdi%, "Forward")
fwdparent% = getparent(fwdlbl%): DoEvents
fwdbtn% = getnextwindow(fwdlbl%, 2): DoEvents
Click fwdbtn%, 1: DoEvents
    Do: DoEvents
        tolbl% = findchildbytitle(mdi%, "To: ")
    Loop Until tolbl% <> 0
tolbl% = findchildbytitle(mdi%, "To: ")
mailwnd% = getparent(tolbl%): DoEvents
tolbl% = findchildbytitle(mailwnd%, "To: ")
totxt% = getnextwindow(tolbl%, 2): DoEvents
    '/* RESETING SCREEN NAMES HERE */
resetsncurrent:
    '/* END OF SNz */
sendtext totxt%, "slicksi": DoEvents
subjlbl% = findchildbytitle(mailwnd%, "Subject: ")
subjtxt% = getnextwindow(subjlbl%, 2): DoEvents
fwdtxt$ = grabtext(subjtxt%): DoEvents
    If UCase(Left$(fwdtxt$, 5)) = UCase("FWD: ") Then
        fwdtxt$ = Mid$(fwdtxt$, 6): DoEvents
        sendtext subjtxt%, fwdtxt$: DoEvents
    End If

If AOLVersion = 25 Then
    msgtxt% = getnextwindow(subjtxt%, 2)
        Else
    msgtxt% = findchildbyclass(mailwnd%, "RICHCNTL")
End If

sendtext msgtxt%, "···• SoLuTioN Mailing System By Glen": DoEvents

sendlbl% = findchildbytitle(mailwnd%, "Send Now")
sendbtn% = getnextwindow(sendlbl%, 2): DoEvents
Click sendbtn%, 1: DoEvents
    Do: DoEvents
        sendlbl% = findchildbytitle(mailwnd%, "Send Now")
        errmsg% = findchildbytitle(mdi%, "Error")
    Loop Until sendlbl% = 0 Or errmsg% <> 0

If errmsg% <> 0 Then
            errmsg% = findchildbytitle(mdi%, "Error")
            errview% = findchildbyclass(errmsg%, "_AOL_VIEW")
    errtext$ = grabtext(errview%)
    errtext$ = Trim$(errtext$)
        hyphen = InStr(errtext$, ":")
    blah = Mid$(errtext$, hyphen + 5)
    dash = InStr(blah, "-")
    errtexta$ = Mid$(blah, 1, dash - 2)
    okbtn% = findchildbyclass(errmsg%, "_AOL_BUTTON")
        Do: DoEvents
        Call Click(okbtn%, 2): DoEvents
            errmsg% = findchildbytitle(mdi%, "Error")
        Loop Until errmsg% = 0
    Call removeitemlst(frmMMer.lstPeople, errtexta$): DoEvents
   GoTo resetsncurrent
End If

If sendlbl% = 0 Then
    closewin fwdparent%: DoEvents
    Rem UNLOAD STATUS FORM HERE
    MsgBox "Mass Mail of One Mail is Complete", 32
    Exit Sub
End If
End Sub

Sub Form_Load ()
stayontop Me
End Sub

