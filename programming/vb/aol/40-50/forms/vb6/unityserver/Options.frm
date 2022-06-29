VERSION 5.00
Begin VB.Form Options 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu OPTIONS 
      Caption         =   "OPTIONS"
      Begin VB.Menu mnuserve 
         Caption         =   "&Serving Name"
      End
      Begin VB.Menu mnuviewmsg 
         Caption         =   "&View Messages"
      End
      Begin VB.Menu mnubanning 
         Caption         =   "&View Banning"
      End
      Begin VB.Menu mnuchangeascii 
         Caption         =   "&Change Ascii"
      End
      Begin VB.Menu mnugreetz 
         Caption         =   "&View Greetz"
      End
      Begin VB.Menu mnuscrolls 
         Caption         =   "&Scrolls"
         Begin VB.Menu mnucommandsi 
            Caption         =   "&Commands Now"
         End
         Begin VB.Menu mnucommands 
            Caption         =   "&Commands"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnustatus 
            Caption         =   "&Status"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnupause 
         Caption         =   "&Pause"
         Begin VB.Menu mnufinds 
            Caption         =   "&Finds"
         End
         Begin VB.Menu mnumailz 
            Caption         =   "&Mailz"
         End
         Begin VB.Menu mnulist 
            Caption         =   "&List"
         End
      End
      Begin VB.Menu mnustyle 
         Caption         =   "&Style"
         Begin VB.Menu mnufastr 
            Caption         =   "&Fast(Risky)"
         End
         Begin VB.Menu mnuslows 
            Caption         =   "&Slow(Safer)"
         End
      End
      Begin VB.Menu lists 
         Caption         =   "&Lists"
         Begin VB.Menu mnutime 
            Caption         =   "&Time (60)"
         End
         Begin VB.Menu mnulistsize 
            Caption         =   "&Size"
         End
         Begin VB.Menu listsend 
            Caption         =   "S&end"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnumisc 
         Caption         =   "&Misc."
         Begin VB.Menu mnurestart 
            Caption         =   "&Restarter"
            Begin VB.Menu mnurestartnum 
               Caption         =   "&Restart ()"
            End
            Begin VB.Menu mnusetup 
               Caption         =   "&Set up"
            End
         End
         Begin VB.Menu mnudelsentmailz 
            Caption         =   "&Delete Sent Mailz"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnusaveit 
            Caption         =   "&Save Pending"
         End
         Begin VB.Menu mnumaxpend 
            Caption         =   "&Max Pending"
         End
         Begin VB.Menu mnureport 
            Caption         =   "&Send Report"
         End
         Begin VB.Menu mnunotify 
            Caption         =   "&Notification"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusendfinds 
            Caption         =   "&Send Finds"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnusavelog 
            Caption         =   "&Save Log"
         End
         Begin VB.Menu mnublocks 
            Caption         =   "&Blocks"
         End
      End
      Begin VB.Menu mnukills 
         Caption         =   "&Kills"
         Begin VB.Menu mnuflashmail 
            Caption         =   "&Flash Mail Dupes"
         End
         Begin VB.Menu mnumodal 
            Caption         =   "M&odal"
         End
         Begin VB.Menu mnuchat 
            Caption         =   "&Chat"
         End
         Begin VB.Menu mnuwait 
            Caption         =   "&Wait"
         End
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'I BELIEVE THIS WHOLE SECTION IS SELF EXPLANATORY, IF YOU DO NOT UNDERSTAND THE CODE HERE
'YOU SHOULD NOT BE ATEMPTING TO MAKE A SERVER

'ALL THE OPTIONS ARE FOR THE USER TO CUSTOMIZE THEIR SERVER
'THIS IS WHERE A LOT OF THE VARIABLES THAT YOU DIDN'T UNDERSTAND WHERE THEY CAME FROM
'COME FROM, LOOK IT OVER!
End Sub
Private Sub listsend_Click()
If listsend.Checked = True Then
    listsend.Checked = False
    FrmMain.Timer3.Enabled = False
Else
    listsend.Checked = True
    FrmMain.Timer3.Enabled = True
End If
End Sub
Private Sub mnubanning_Click()
Ban.Show
End Sub
Private Sub mnublocks_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many mails per block? 0 is for no blocks", "AMOUNT OF BLOCKS")
If IsNumeric(Answer) = False Then Exit Sub
If Answer > 99 Or Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number 0 - 99", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
If PendAmt = 0 Then GoTo nahman
If Answer > PendAmt Then
    check = MsgBox("Your block amount can not excede your max pending amount would you like use to adjust both?", vbYesNo, "CAN'T EXCEDE MAX PENDING")
    If check = vbYes Then
    PendAmt = Answer
    mnumaxpend.Caption = "&Max Pending (" & PendAmt& & ")"
    Else
    Answer = PendAmt
    End If
End If
nahman:
BlockAmt& = Answer
mnublocks.Caption = "&Blocks (" & BlockAmt& & ")"
FormOnTop FrmMain
End Sub
Private Sub mnuchangeascii_Click()
FormNotOnTop FrmMain
LAscii = InputBox("Input Left Ascii", "LEFT ASCII")
RAscii = InputBox("Input Right Ascii", "RIGHT ASCII")
FormOnTop FrmMain
End Sub
Private Sub mnuchat_Click()
room& = FindRoom&
Call SendMessage(room&, WM_CLOSE, 0&, 0&)
End Sub
Private Sub mnucommands_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many minutes to wait before sending commands?", "AMOUNT IN MINUTES")
If IsNumeric(Answer) = False Then Exit Sub
CmdTime& = Answer
mnucommands.Caption = "&Commands (" & CmdTime & ")"
FormOnTop FrmMain
End Sub
Private Sub mnucommandsi_Click()
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity Server(AOL4) -By- KiD" & RAscii
Pause 0.2
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "''/" & User$ & " Send '' & X, X-Y(" & BlockAmt & ")" & RAscii
Pause 0.2
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "''/" & User$ & " Send '' & List(" & Count2 & ")(" & ListTime & " secs)" & RAscii
Pause 0.7
End Sub
Private Sub mnudelsentmailz_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many mailz to send before killing sent mailz", "DELETE SENT MAILZ")
If IsNumeric(Answer) = False Then Exit Sub
If Answer > 1000 Or Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number 0 - 1000", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
KillSentMail& = Answer
mnudelsentmailz.Caption = "&Delete Sent Mail (" & KillSentMail & ")"
FormOnTop FrmMain
End Sub
Private Sub mnufastr_Click()
If mnufastr.Checked = True Then Exit Sub
    mnufastr.Checked = True
    mnuslows.Checked = False
End Sub
Private Sub mnufinds_Click()
FormNotOnTop FrmMain
findpausechecker = InputBox("Time to pause for lists(in tenthes of a second:", "LIST PAUSE")
If IsNumeric(findpausechecker) = False Then Exit Sub
If findpausechecker > 10 Then
    MsgBox "That pause time is to long", , "TOO LONG"
    FormOnTop FrmMain
    Exit Sub
End If
FindPause = findpausechecker / 10
mnufinds.Caption = "&Finds (" & FindPause & ")"
FormOnTop FrmMain
End Sub
Private Sub mnuflash_Click()
mnuflash.Checked = True
mnunew.Checked = False
End Sub
Private Sub mnuflashmail_Click()
MailOpenFlash
Pause 1
Call MailDeleteFlashDuplicates(Me, False)
End Sub
Private Sub mnugreetz_Click()
Load greetz
End Sub
Private Sub mnulist_Click()
FormNotOnTop FrmMain
listpausechecker = InputBox("Time to pause for lists(in tenthes of a second:", "LIST PAUSE")
If IsNumeric(listpausechecker) = False Then Exit Sub
If listpausechecker > 10 Then
    MsgBox "That pause time is to long", , "TOO LONG"
    FormOnTop FrmMain
    Exit Sub
End If
ListPause = listpausechecker / 10
mnulist.Caption = "&List (" & ListPause & ")"
FormOnTop FrmMain
End Sub
Private Sub mnulistsize_Click()
FormNotOnTop FrmMain
Answer = InputBox("Many items per list do you want", "ITEMS PER LIST")
If IsNumeric(Answer) = False Then Exit Sub
If Answer > 500 Or Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number 0 - 500", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
ListSize& = Answer
mnulistsize.Caption = "&Size (" & ListSize & ")"
FormOnTop FrmMain
End Sub
Private Sub mnumailz_Click()
FormNotOnTop FrmMain
listpausechecker = InputBox("Time to pause for lists(in tenthes of a second:", "LIST PAUSE")
If IsNumeric(listpausechecker) = False Then Exit Sub
If listpausechecker > 10 Then
    MsgBox "That pause time is to long", , "TOO LONG"
    FormOnTop FrmMain
    Exit Sub
End If
MailzPause = listpausechecker / 10
mnumailz.Caption = "&Mail (" & MailzPause & ")"
FormOnTop FrmMain
End Sub
Private Sub mnumaxpend_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many mails can a single person have pending at a time, 0 is for unlimited", "AMOUNT PENDING")
If IsNumeric(Answer) = False Then Exit Sub
If Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number greater then or equal to 0", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
PendAmt& = Answer
If PendAmt& = 0 Then GoTo skippndaif
If BlockAmt > PendAmt Then
    BlockAmt = PendAmt
    MsgBox "Your Block Amount Was Also Adjusted", , "BLOCK AMOUNT ADJUSTED"
    mnublocks.Caption = "&Blocks(" & PendAmt& & ")"
End If
skippndaif:
mnumaxpend.Caption = "&Max Pending (" & PendAmt& & ")"
FormOnTop FrmMain
End Sub
Private Sub mnumodal_Click()
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
Call PostMessage(aolmodal&, WM_CLOSE, 0, 0&)
End Sub
Private Sub mnunew_Click()
mnuflash.Checked = False
mnunew.Checked = True
End Sub
Private Sub mnunewdupe_Click()
MailOpenNew
Call MailDeleteNewDuplicates(Me, False)
End Sub
Private Sub mnunotify_Click()
If mnunotify.Checked = True Then
    mnunotify.Checked = False
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "<font face=" & Chr(34) & "tahoma" & Chr(34) & "> Notifaction Has Been Turned Off <font face=" & Chr(34) & "verdana" & Chr(34) & ">" & RAscii
Else
    mnunotify.Checked = True
End If
End Sub
Private Sub mnureport_Click()
FormNotOnTop FrmMain
usercheck$ = InputBox("Send To Who:", "TO WHOM")
If usercheck$ = "" Then Exit Sub
If STOPIT = True Then
    STOPIT = False
    turnit = True
End If
aol4_mail_send usercheck$, "¹~·-.¸  Unity Server(AOL4) -By- KiD Server Report", "Date: " & Date & Chr(13) & "Total Served: " & FrmMain.Finished.ListCount, True
If turnit = True Then STOPIT = True
FormOnTop FrmMain
End Sub
Private Sub mnurestartnum_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many mailz to send before restarting", "MAILZ TO SEND BEFORE RESTART")
If IsNumeric(Answer) = False Then Exit Sub
If Answer > 2000 Or Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number 0 - 2000", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
AolRestartNum& = Answer
mnurestartnum.Caption = "&Restart (" & AolRestartNum & ")"
FormOnTop FrmMain
End Sub
Private Sub mnusaveit_Click()
If mnusaveit.Checked = False Then
    sure = MsgBox("Are you sure you want to add this feature?, although it's helpful incase of crashes, depending on the amount you have pending this could add a noticable lag", vbYesNo, "ARE YOU SURE")
    If sure = vbYes Then
        mnusaveit.Checked = True
        Exit Sub
    End If
End If
mnusaveit.Checked = False
End Sub
Private Sub mnusavelog_Click()
FormNotOnTop FrmMain
usercheck$ = InputBox("Pick a name for the log:", "LOG NAME")
If usercheck$ = "" Then Exit Sub
If InStr(usercheck, ".") > 0 Then
    On Error Resume Next
    puini = FreeFile
    Open App.Path & "/" & usercheck$ For Output As #puini
    For x = 0 To FrmMain.Finished.ListCount - 1
        Write #puini, FrmMain.Finished.List(x)
    Next x
    Close #puini
    MsgBox "Log Saved", , "LOG SAVED"
Else
    On Error Resume Next
    puini = FreeFile
    Open App.Path & "/" & usercheck$ & ".log" For Output As #puini
    For x = 0 To FrmMain.Finished.ListCount - 1
        Write #puini, FrmMain.Finished.List(x)
    Next x
    Close #puini
    MsgBox "Log Saved", , "LOG SAVED"
End If
FormOnTop FrmMain
End Sub
Private Sub mnusendfinds_Click()
If mnusendfinds.Checked = True Then
    mnusendfinds.Checked = False
Else
    mnusendfinds.Checked = True
End If
End Sub
Private Sub mnuserve_Click()
FormNotOnTop FrmMain
usercheck$ = InputBox("Pick a new handle:", "NEW HANDLE")
If Len(usercheck$) > 10 Or Len(usercheck$) < 1 Then
    MsgBox "Please pick a handle 1-10 characters", , "INVALID HANDLE"
    FormOnTop FrmMain
    Exit Sub
End If
User$ = usercheck$
mnuserve.Caption = "&Serving Name (" & User$ & ")"
FormOnTop FrmMain
FrmMain.totalLbl = Total& + FrmMain.Finished.ListCount
FrmMain.userlbl = User$
FrmMain.pendinglbl = FrmMain.Pending.ListCount
FrmMain.finishedlbl = FrmMain.Finished.ListCount
End Sub
Private Sub mnusetup_Click()
Restarter.Show
End Sub
Private Sub mnuslows_Click()
If mnuslows.Checked = True Then Exit Sub
    mnuslows.Checked = True
    mnufastr.Checked = False
End Sub
Private Sub mnustatus_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many minutes to wait before sending status?", "AMOUNT IN MINUTES")
If IsNumeric(Answer) = False Then Exit Sub
StatusTime& = Answer
mnustatus.Caption = "&Status (" & StatusTime & ")"
FormOnTop FrmMain
End Sub
Private Sub mnutime_Click()
FormNotOnTop FrmMain
Answer = InputBox("How many seconds to wait before sending lists?", "AMOUNT IN SECONDS")
If IsNumeric(Answer) = False Then Exit Sub
If Answer > 99 Or Answer < 0 Then
    FormNotOnTop FrmMain
    MsgBox "Please pick a number 0 - 99", , "INVALID NUMBER"
    Exit Sub
    FormOnTop FrmMain
End If
ListTime& = Answer
mnutime.Caption = "&Time (" & ListTime & ")"
FormOnTop FrmMain
End Sub
Private Sub mnuviewmsg_Click()
Mesages.Show
End Sub
Private Sub mnuwait_Click()
aol4_killwait
End Sub
