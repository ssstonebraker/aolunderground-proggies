VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Planet Unity(AOL4) Server -By- KiD"
   ClientHeight    =   2685
   ClientLeft      =   -15615
   ClientTop       =   -3135
   ClientWidth     =   5145
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer getrequests 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4320
      Top             =   405
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   4725
      Top             =   0
   End
   Begin VB.ListBox Lists 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2430
      Width           =   4290
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   0
   End
   Begin VB.ListBox Mailz 
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   2160
      Width           =   4290
   End
   Begin VB.ListBox Finished 
      Height          =   450
      Left            =   2610
      TabIndex        =   17
      Top             =   810
      Width           =   1365
   End
   Begin VB.ListBox Pending 
      Height          =   450
      Left            =   360
      TabIndex        =   16
      Top             =   810
      Width           =   1320
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   3960
      MouseIcon       =   "Form1.frx":164A
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   3735
      MouseIcon       =   "Form1.frx":179C
      MousePointer    =   99  'Custom
      Top             =   45
      Width           =   195
   End
   Begin VB.Label statuslbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stopped"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   2970
      TabIndex        =   15
      Top             =   1845
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   2430
      TabIndex        =   14
      Top             =   1845
      Width           =   495
   End
   Begin VB.Label userlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1620
      TabIndex        =   13
      Top             =   1845
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1170
      TabIndex        =   12
      Top             =   1845
      Width           =   375
   End
   Begin VB.Label totalLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000018&
      Height          =   150
      Left            =   585
      TabIndex        =   11
      Top             =   1845
      Width           =   90
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   1845
      Width           =   405
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000018&
      X1              =   2340
      X2              =   2340
      Y1              =   1755
      Y2              =   2025
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000018&
      X1              =   1080
      X2              =   1080
      Y1              =   1755
      Y2              =   2025
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000018&
      X1              =   135
      X2              =   4185
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000018&
      X1              =   135
      X2              =   4185
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label pausebut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   900
      MouseIcon       =   "Form1.frx":18EE
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label optionsbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   3375
      MouseIcon       =   "Form1.frx":1A40
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1530
      Width           =   660
   End
   Begin VB.Label loadbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   2340
      MouseIcon       =   "Form1.frx":1B92
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label stopbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1665
      MouseIcon       =   "Form1.frx":1CE4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1530
      Width           =   405
   End
   Begin VB.Label startbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   270
      MouseIcon       =   "Form1.frx":1E36
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1530
      Width           =   420
   End
   Begin VB.Label Percentlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-------->"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1935
      TabIndex        =   4
      Top             =   945
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Finished List:"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   2655
      TabIndex        =   3
      Top             =   540
      Width           =   915
   End
   Begin VB.Label finishedlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000018&
      Height          =   150
      Left            =   3690
      TabIndex        =   2
      Top             =   540
      Width           =   90
   End
   Begin VB.Label pendinglbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H80000018&
      Height          =   150
      Left            =   1530
      TabIndex        =   1
      Top             =   540
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pending List:"
      ForeColor       =   &H80000018&
      Height          =   150
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   -45
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   0
      Picture         =   "Form1.frx":1F88
      Top             =   0
      Width           =   4320
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next    'Incase of error it resumes the next command without sending an error message
SplashScreen.statuslbl = "Setting Mail Preferences" 'Sets the label caption on the Splash Screen to let the user know what's going on
Call SetMailPrefs   'Calls a sub in the bas, check the sub for documentation
Me.Height = Image1.Height   'Resizes the form to fit the image
Me.Width = Image1.Width 'Resizes the form to fit the image
welcometxt$ = GetText(aol4_welcomescreen)   'Calls subs from the bas to get the text from the welcome screen(this is how the aol users screen name is found)
If aol4_welcomescreen = 0 Then Exit Sub 'Checks to see if there is a welcome screen, if not then it doesn't coninue loading, if there is no welcome screen that means the user is not signed on
User$ = Right(welcometxt$, Len(welcometxt$) - Len("welcome, ")) 'Gets the text from the right of welcome, (this is where the aol users sn will be found)
User$ = Left(User$, Len(User$) - 1) 'Trims the screen name of the added character on the welcome screens title
SN$ = User$ 'Sets a variable easy for me to remember to the users screen name
OPTIONS.mnuserve.Caption = "&Serving Name(" & User$ & ")"   'Changes the menu name to fit the aol users screen name
FormOnTop Me    'Calls formontop from the bas to make this form stay on top of all the others
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString) 'Searches for the window AOL Modal, this is a common aol error and aol form
SplashScreen.statuslbl = "Killing Modal"    'Notify's the user what it's doing again
Call PostMessage(aolmodal&, WM_CLOSE, 0, 0&)    'Closes the modal if it's found
On Error Resume Next    'See Above
puini = FreeFile    'Sets the variable puini as freefile so that vb can open it and work with it as a freefile
SplashScreen.statuslbl = "Loading Options...."  'Notifys user of wut's going on
Open App.Path & "\planet unity.ini" For Input As #puini 'Opens the specified file for input, which means that it is inputing variables into vb not inputing into the file
Input #puini, checker$  'Sets the variable checker to the first line in the file
Input #puini, ListTime& 'Sets the variable ListTime& to the Second line in the file
Input #puini, BlockAmt& 'Sets the variable BlockAmt& to the Third line in the file
Input #puini, Total     'Etc......
Input #puini, ListPause 'Etc......
Input #puini, MailzPause 'Etc......
Input #puini, FindPause 'Etc......
Input #puini, PendAmt   'Etc......
Input #puini, saved 'Etc......
Input #puini, MailMsg   'Etc......
Input #puini, SentMsg   'Etc......
Input #puini, LAscii    'Etc......
Input #puini, RAscii    'Etc......
Input #puini, StatusTime    'Etc......
StatusTime = 0 'Was an option i had in but removed due to Waol errors
Input #puini, CmdTime
CmdTime = 0 'Was an option i had in but removed due to Waol errors
Input #puini, findsman  'Etc......
Input #puini, styledood 'Etc......
Input #puini, KillSentMail  'Etc......
Input #puini, ListSize  'Etc......
Input #puini, AolRestartNum 'Etc......
Input #puini, AolDir    'Etc......
Input #puini, AolRestartPw  'Etc......
'The following just sets option menu's names to make the server more user friendly
OPTIONS.mnurestartnum.Caption = "&Restart (" & AolRestartNum & ")"
OPTIONS.mnulistsize.Caption = "&Size (" & ListSize& & ")"
OPTIONS.mnudelsentmailz.Caption = "&Delete Sent Mail (" & KillSentMail& & ")"
OPTIONS.mnutime.Caption = "&Time (" & ListTime& & ")"
OPTIONS.mnublocks.Caption = "&Block (" & BlockAmt& & ")"
OPTIONS.mnulist.Caption = "&List (" & ListPause & ")"
OPTIONS.mnumailz.Caption = "&Mail (" & MailzPause & ")"
OPTIONS.mnufinds.Caption = "&Find (" & FindPause & ")"
OPTIONS.mnucommands.Caption = "&Commands (" & CmdTime& & ")"
OPTIONS.mnustatus.Caption = "&Status (" & StatusTime& & ")"
OPTIONS.mnumaxpend.Caption = "&Max Pending (" & PendAmt& & ")"
If saved = 1 Then OPTIONS.mnusaveit.Checked = True
If findsman = 1 Then OPTIONS.mnusendfinds.Checked = True
If styledood = 1 Then
    OPTIONS.mnufastr.Checked = True
    OPTIONS.mnuslows.Checked = False
ElseIf styledood = 0 Then
    OPTIONS.mnuslows.Checked = True
    OPTIONS.mnufastr.Checked = False
End If
If checker$ = "<sn>" Then   'Checks the checker to see if the server has been used before based on the file
OPTIONS.mnuserve.Caption = "&Serving Name(" & checker$ & ")"
    Close #puini
    SplashScreen.Hide
    Exit Sub
End If
User$ = checker$
OPTIONS.mnuserve.Caption = "&Serving Name(" & User$ & ")"
Close #puini
SplashScreen.statuslbl = "Loading Pending List"
Open App.Path & "\pending save.ini" For Input As #puini 'opens the specified file for input
    While Not EOF(1)    'a while loop that goes by one line until the file is not at its end
        Input #puini, requesttoadd$ 'adds the request
        Pending.AddItem requesttoadd$   'adds the request
    Wend    'end of the while loop
Close #puini    'closes the file
SplashScreen.statuslbl = "Sending Ascii"
'ChatSend "</a>< a href=" & Chr(34) & "{s Application Data\microsoft\WELCOME\Welcom98" & Chr(34) & "></a>"
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity Server(AOL4)ฒทณ -By- KiD" & RAscii   'sends chat
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Total Since Install " & Total & "" & RAscii    'sends chat
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "</a>< a href=" & Chr(34) & "http://parplex.com/chip/Unity.ZIP" & Chr(34) & ">Get It Here</a>" & RAscii 'sends chat
SplashScreen.Hide   'Hides the splash screen
bill = checkdeadmsg 'checks for a dead message
totalLbl = Total& + FrmMain.Finished.ListCount  'Sets the lable to the total sent count
userlbl = User$
pendinglbl = FrmMain.Pending.ListCount
finishedlbl = FrmMain.Finished.ListCount
End Sub
Private Sub getrequests_Timer()
Static StatusLeft As Long
GetReqs = "Ready"   'tells the server that it is ready to collect the requests
StatusLeft = StatusLeft + 5
If StatusLeft >= 120 Then
    StatusLeft = 0
    StatusReady = True
End If
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me 'makes the form moveable by dragging it
FormOnTop MChat 'makes the mchat stay on top
MChat.Left = FrmMain.Left - ((MChat.Width - FrmMain.Width) / 2) 'centers the mchat on the bottom of the form
MChat.Top = FrmMain.Top + FrmMain.Height
End Sub
Private Sub Image3_Click()
Me.WindowState = 1  'minimizes the main form
End Sub
Private Sub Image4_Click()
Call ShowWindow(FindRoom, SW_SHOW)  'Makes the aol chatrooom visable
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity Server(AOL4)ฒทณ -By- KiD" & RAscii
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Total Since Install " & Total & "" & RAscii
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "</a>< a href=" & Chr(34) & "http://parplex.com/chip/Unity.ZIP" & Chr(34) & ">Get It Here</a>" & RAscii
'saving preferences
On Error Resume Next    'on error resume next command
puini = FreeFile    'designate puini as a freefile
Open App.Path & "\Planet Unity.ini" For Output As #puini    'open the specified file for output(out put into a file not into vb)
Write #puini, User$ 'write the variable user in the first line
Write #puini, ListTime& 'write the variable listtime in the second line
Write #puini, BlockAmt& 'write the variable BlockAmt& in the thrid line
Write #puini, Total + FrmMain.Finished.ListCount    'etc...
Write #puini, ListPause    'etc...
Write #puini, MailzPause    'etc...
Write #puini, FindPause    'etc...
Write #puini, PendAmt    'etc...
If OPTIONS.mnusaveit.Checked = True Then saved = 1
If OPTIONS.mnusaveit.Checked = False Then saved = 0
Write #puini, saved    'etc...
Write #puini, MailMsg    'etc...
Write #puini, SentMsg    'etc...
Write #puini, LAscii    'etc...
Write #puini, RAscii    'etc...
Write #puini, StatusTime    'etc...
Write #puini, CmdTime    'etc...
If OPTIONS.mnusendfinds.Checked = True Then saved = 1
If OPTIONS.mnusendfinds.Checked = False Then saved = 0
Write #puini, saved    'etc...
If OPTIONS.mnufastr.Checked = True Then styledood = 1
If OPTIONS.mnuslows.Checked = True Then styledood = 0
Write #puini, styledood    'etc...
Write #puini, KillSentMail    'etc...
Write #puini, ListSize    'etc...
Write #puini, AolRestartNum    'etc...
Write #puini, AolDir    'etc...
Write #puini, AolRestartPw    'etc...
Close #puini
On Error Resume Next
'saves your pending list incase of server or aol crash
puini = FreeFile
If OPTIONS.mnusaveit.Checked = True Then
Open App.Path & "\pending save.ini" For Output As #puini
    For x = 0 To FrmMain.Pending.ListCount - 1
        Write #puini, Pending.List(x)
    Next x
Close #puini
End If
End
End Sub
Private Sub loadbut_Click()
On Error Resume Next
Dim count, counter&
    Mailz.Clear 'clears your mailz list
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Loading FlashMail" & RAscii
    statuslbl = "Loading"
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'locates the flashmail box
    If fMail& <> 0 Then GoTo nextstep   'checks to see if it is alreayd open, if so it skips the next part
    MailOpenFlash   'calls a sub in the bas to open the flash mail box
    aol& = findwindow("AOL Frame25", vbNullString)  'locates the aol window
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdi window
    Do  'begins a do loop
        If STOPIT = True Then Exit Sub  'if stopit = true it exits, stopit is set to true when the user clicks stop.... i did this incase the server freezes, the only way this is possible is in a loop so if the user clicks stop it exits all loops and then if he clicks start it'll start as normal
        DoEvents    'a set command to let vb and other applications do needed events so that the server doesn't freeze the computer while in a loop
        fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'locates the flashmail box
    Loop Until fMail& <> 0  'loops until the box is found
nextstep:   'label, there is a goto above that is used when the mail box is already open
    MailToListFlash Mailz   'calls a sub to add the mailz to a list box
    Call Killwin(fMail&)    'calls a sub to close the window with the handle fmail&, which is the flashmail box as we said above
Count2& = 0 'sets a variable to be used later to 0
    If InStr(Mailz.List(0), "น~ท-.ธ  Unity Server(AOL4) -By- KiD List ") > 0 Then   'checks the list of mailz for old lists by it's title
        Count2& = Count2& + 1   'for every old list it finds it sets count2 to count2 + 1 so count is keeping track of the number of lists to be used later in loops
        For x = 1 To Mailz.ListCount - 1    'a for loop to continue finding lists
             If InStr(Mailz.List(x), "น~ท-.ธ  Unity Server(AOL4) -By- KiD List ") > 0 Then Count2& = Count2& + 1 'adds one if the next index in the mailz listbox is a list
             If InStr(Mailz.List(x), "น~ท-.ธ  Unity Server(AOL4) -By- KiD List ") = 0 Then Exit For 'exits the loop if the next mail in the list is not a list
        Next x  'ends the for
        ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Old Lists Found" & RAscii
        FormNotOnTop Me 'negates the forontop, to prepare for a message box
        checkhim = MsgBox("Unity Server Has Noticed You Have Valid Old Lists, Would You Like To Reuse Them?(IF PROBLEMS OCCUR YOU MIGHT WANT TO JUST MAKE NEW ONES", vbYesNo, "REUSE OLD LISTS")    'asks the user if he would like to use old lists
        FormOnTop Me    'sets the form back ontop
        If checkhim = vbYes Then    'if the user clicks yes then it uses old lists
        ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Using Old Lists(" & Count2 & ")" & RAscii
            For x = 0 To Count2& - 1    'a loop to remove the lists from the list box, some you are serving doesn't want to see the lists in the list
                Mailz.RemoveItem 0  'removes the index 0, after removing one item the next item's index becomes 0
            Next x  'ends the for loop
            GoTo skipdasendpart 'goes to a label to skip the part in which it sends the ilsts
        End If
        If checkhim = vbNo Then Count2& = 0 'if the user clicked no, the he wants to make new lists and there for we have no lists so we set count2 back down to 0
    End If  'ends the if that began if the first item in the list was a list
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Creating New Lists" & RAscii
    For x = 0 To Mailz.ListCount - 1    'begins a loop with the demensions based on the amount of mailz you have
        statuslbl = "Making " & Count2& + 1
        If x <= 9 Then dastring$ = dastring$ & x & ".)" & Chr(9) & Chr(9) & LTrim(Mailz.List(x)) & Chr(13)  'just organizes the lists in an ordily fashion with the proper indentation and index numbers for the people you are serving to request from, based on the loop
        If x > 9 Then dastring$ = dastring$ & x & ".)" & Chr(9) & LTrim(Mailz.List(x)) & Chr(13)    'same as above
        count = count + 1   'increases count by one, if the lists are to big then aol can't send them so we must limit the size of each list
        If count >= ListSize Then   'checks to see if count is = to the amount of items per list designated by the user
            Count2& = Count2& + 1   'if the above is true then we have a list, so 1 is added to count2
            statuslbl = "Sending " & Count2&
            aol4_mail_send SN$, "จน~ท-.ธ  Unity Server(AOL4) -By- KiD List " & Count2&, dastring$, False    'calls a sub to send a mail to the user containing the list and index numbers
            ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "List " & Count2 & " Created/Sent" & RAscii
            count = 0   'sets count back to 0 so that it can count the amount for the next list
            dastring$ = ""  'resets the dastring, this was where the string for list was stored previously
        End If  'ends the if
    Next x  'ends the loop
    If dastring$ = "" Then Exit Sub 'checks dastring for any remaining items that weren't sent because the last list didn't reach the max size
    Count2& = Count2& + 1   'adds one more to send the last list
    statuslbl = "Sending " & Count2&
    aol4_mail_send SN$, "จน~ท-.ธ  Unity Server(AOL4) -By- KiD List " & Count2&, dastring$, False    'sends the last list
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "List " & Count2 & " Created/Sent" & RAscii
    statuslbl = "Counting Mail"
    MailOpenNew 'calls a sub to open your new mailbox
    MailzRdy = MailCountNew& - Count2&  'calls a sub and puts it in the variable mailzrdy, it's a sub to count your new mail in your new maibox
    Call Killwin(FindMailBox)   'closes your flashmail box, you are about to flash the lists to your flashmail, it's best that you do close the mailbox one last time
    statuslbl = "Running Flash"
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Running Flasmail" & RAscii
    MailRunFlash    'calls a sub to begin a flashmail session and wait until the end of it
     ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Finished And Ready To Serve" & RAscii
skipdasendpart:
    MailOpenFlash   'reopens your mailbox one last time
    fMail& = 0& 'sets fmail to 0 so that it can check to see when it opens
    Do  'begins a loop
    DoEvents    'see above
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'checks for the flashmail box
    Loop Until fMail& <> 0  'loops until the mailbox is found
    Call ShowWindow(fMail&, SW_MINIMIZE)    'minimizes the mailbox
    statuslbl = "Stopped"
    startbut.Enabled = True 'enables the start button
End Sub
Private Sub optionsbut_Click()
On Error Resume Next
PopupMenu OPTIONS!OPTIONS, , optionsbut.Left, optionsbut.Top + optionsbut.Height    'opens a menu from the form options with the proper coordinates
End Sub
Private Sub pausebut_Click()
On Error Resume Next
ServerPaused = True
Call ShowWindow(FindRoom, SW_SHOW)  'shows the chat room
Timer3.Enabled = False  'turns off timer3
startbut.Enabled = True 'enables the start button
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Paused Still Taking Requests" & RAscii
End Sub
Private Sub Pending_DblClick()
If FrmMain.Pending.ListCount = 0 Then Exit Sub  'checks to see if there are any pending requests, if not there is no reason to clear them
FormNotOnTop Me 'negates the formontop from before to get ready for a message box or inputbox
sure = MsgBox("Are you sure you want to clear the pending list?", vbYesNo, "CLEAR PENDING LIST")    'asks the user if they really want to clear the pending list box
If sure = vbYes Then Pending.Clear  'if the user clicked yes then it clears the pending list box
FormOnTop Me    'sets the form back on top
End Sub
Private Sub startbut_Click()
On Error Resume Next
'If FindRoom = 0 Then Exit Sub
Call ShowWindow(FindRoom, SW_HIDE)  'minimizes the chat room
ServerPaused = False
MChat.Show  'shows the form m-chat
FormOnTop MChat 'makes mchat the top most form
STOPIT = False
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity Server(AOL4) -By- KiD" & RAscii
Pause 0.2
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "''/" & User$ & " Send '' & X, X-Y(" & BlockAmt & ")" & RAscii
Pause 0.2
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "''/" & User$ & " Send '' & List(" & Count2 & ")(" & ListTime & " secs)" & RAscii
Pause 0.7
loadbut.Enabled = False 'makes the load mail button false so that the user doesn't load mail while serving
startbut.Enabled = False    'disables the start button, if the user is serving there's no reason to click start again
pausebut.Enabled = True 'enables the pause button
stopbut.Enabled = True  'enables the stop button
Timer2.Enabled = True   'starts timer 2
If OPTIONS.listsend.Checked = False Then GoTo skipit    'if the user has the option to not send lists checked, then it doesn't start timer3, the timer in which is used to tell the server it's ready to send lists
Timer3.Enabled = True   'starts timer3
skipit: 'the label to skip enables the send list timer
getrequests.Enabled = True  'starts the getrequests timer
statuslbl = "Waiting"
Timer3.Interval = ListTime& * 1000  'sets timer3 to the appropriot interval designated by the user, each interval is 1/1000 of a second, the user is asked how many seconds to wait before sending lists, so you take that variable and multiply by 1000 to get the correct amount of seconds
End Sub
Private Sub stopbut_Click()
On Error Resume Next
Call ShowWindow(FindRoom, SW_SHOW)  'shows the chat room
MChat.Hide  'hides the mchat form
STOPIT = True   'sets stopit to true, this is a variable used to exit all other loops
Timer2.Enabled = False  'stops timer 2
Timer3.Enabled = False  'stops timer 3
getrequests.Enabled = False 'stops the timer to get requests
startbut.Enabled = True 'enables the start button for the user to restart the server
pausebut.Enabled = False    'enables the pause button for the user to pause the server
loadbut.Enabled = True  'enables the load button incase the user wants to reload the mail
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity Server(AOL4) -By- KiD" & RAscii
ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Stopped Not Taking Requests" & RAscii
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Dim screenname As String, Place As Integer, num As Long, Num2 As Long
Dim toadd As Long, towho As String, x As Long, roomname As String

If GetReqs = "Ready" Then   'checks to see if it's time to collect requests
    GetReqs = "Not Ready"   'makes sure it doesn't continually collect requests
    Call Get_Reqs   'calls a sub to get requests
End If  'ends the if

If ServerPaused = True Then Exit Sub

If TillAolRestart >= AolRestartNum And AolRestartNum <> 0 And AolRestartPw <> "" And AolDir <> "" Then  'checks to see if the user filled out all the correct information to restart aol and that the variable used to keep track of the amount sent is equal to the amount needed to restart aol
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Restarting AOL BRB!" & RAscii
    room = FindRoom 'makes room equal to the sub findroom, which finds the aol chat room
    roomname = GetCaption(FindRoom) 'gets the caption of the aol chat room
    aol& = findwindow("AOL Frame25", vbNullString)  'finds aol
    Killwin aol&    'closes aol
    FrmMain.statuslbl = "Restarting AOL"
    Do: DoEvents    'begins a loop
    If STOPIT = True Then Exit Sub  'see above
    aol& = findwindow("AOL Frame25", vbNullString)  'find aol
    Loop Until aol& = 0 'loops until aol is gone
    Pause 3 'waits 3 seconds to make sure the computer is ready
    FrmMain.statuslbl = "Opening AOL"
    Shell AolDir, vbNormal  'opens the aol directory set by the user
    Do: DoEvents    'begins a loop
        If STOPIT = True Then Exit Sub  'see above
        aol& = findwindow("AOL Frame25", vbNullString)  'locates the aol window
        MDI& = FindWindowEx(aol&, 0, "MDIClient", vbNullString) 'locates the mdi clinet
        SignOnScreen& = FindWindowEx(MDI&, 0, "AOL Child", "Sign On")   'locates the sign on screen
    Loop Until SignOnScreen& <> 0&  'loops until the sign on screen is found
    Current = Timer 'sets current = to timer, timer is a set command by vb to give an integer for timing reasons
    FrmMain.statuslbl = "Signing On AOL"
    Do: DoEvents    'begins a new loop
        If STOPIT = True Then Exit Sub  'see above
        If Timer - Current >= 3 Then GoTo clicksignon   'if 3 seconds have gone by it quits waiting and decides to move on
        PwBox& = FindWindowEx(SignOnScreen&, 0, "_AOL_Edit", vbNullString)  'locates the password box
    Loop Until PwBox& <> 0& 'loops until the password box is found
    Call textset(PwBox&, AolRestartPw)  'sets the passwordbox to the users password using the seub settext
clicksignon:    'the label used to goto from above
    icon1& = FindWindowEx(SignOnScreen&, 0, "_AOL_Icon", vbNullString)  'finds one icon
    icon2& = FindWindowEx(SignOnScreen&, icon1&, "_AOL_Icon", vbNullString) 'finds another icon
    icon3& = FindWindowEx(SignOnScreen&, icon2&, "_AOL_Icon", vbNullString) 'and yet another one
    SignOnBut& = FindWindowEx(SignOnScreen&, icon3&, "_AOL_Icon", vbNullString) 'ahhh finally, the forth icon is the icon you click to sign on
    clickicon SignOnBut&    'clicks the sign on icon
    TillAolRestart = 0  'resets the number of mailz kept track to check when it's time to restart aol
    Do: DoEvents    'begins a loop
        If STOPIT = True Then Exit Sub
    Windo = aol4_welcomescreen  'lfinds the welcome screen is found using the sub aol4_welcome screen
    Loop Until Windo <> 0   'loops until the welcome screen is found using the sub aol4_welcome screen
    Pause 4 'waits 4 seconds to make sure aol is done loading
    FrmMain.statuslbl = "Killing Windows"
    Child& = FindWindowEx(MDI, 0, "AOL Child", vbNullString)    'finds an aol child
    Call ShowWindow(Child&, SW_MINIMIZE)    'minizes the found aol child
    Killwin Child&  'attempts to close it
    For x = 1 To 10 'begins a for loop of 10
    Child& = FindWindowEx(MDI, Child&, "AOL  Child", vbNullString)  'finds another child
    Call ShowWindow(Child&, SW_MINIMIZE)    'minimizes it
    Killwin Child&  'atempts to close it
    Next x  'repeats
    FrmMain.statuslbl = "Entering Room"
    Do: DoEvents    'begins a do looop
        If STOPIT = True Then Exit Sub  'see above
        PrivateRoom roomname    'trys re entering the room with a sub called private room
    Loop Until WaitForOKOrRoom(roomname) <> "OK"    'waits for room full message or room using a sub
    Current = Timer 'see above
    MailOpenFlash   'reopens the flashmail
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient
    Do  'begins a do loop
        If STOPIT = True Then Exit Sub  'see above
        DoEvents    'see above
        fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'locates the flashmail box
    Loop Until fMail& <> 0  'loops until the flashmail box is found
    FrmMain.statuslbl = "Waiting"
    Pause 1 'waits one second
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Restart Complete!" & RAscii
End If  'ends if

'If TillSentKill >= KillSentMail And KillSentMail <> 0 Then  'checks to see if the variable used to keep track of mailz sent is equal to the variable designated by the user at which it's time to delete all the sent mail
'ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Killing Sent Mail" & RAscii
'    TillSentKill = 0    'resets that variable
'openthesetnmail:    'label used to goto when the sent mail box is not found after a period of time
'    Call aol4_runpopup2(3, "S") 'calls a sub to open the sent mail box
'    Do: DoEvents    'begins a loop
'        If STOPIT = True Then Exit Sub
'        counter& = MailCountSent&   'counts the sent mail once
'            Pause 0.65
'        counter2& = MailCountSent&  'then again
'            Pause 0.65
'        counter3& = MailCountSent&  'and once more
'            Pause 0.65
'    Loop Until counter& = counter2& And counter2& = counter3&   'makes sure they're all equal before continuing
'    Do: DoEvents    'begins another loop
'        MailDeleteNewByIndex 0  'begins deleting the sent mail
'    Loop Until MailCountSent = 0    'loops until it's all gone
'    Killwin FindMailBox 'closes the sentmail box
'End If  'ends the if
If StatusReady = True Then
    StatusReady = False
    Pause 0.7
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Total Pending: " & Pending.ListCount & " Lists Pending: " & Lists.ListCount & RAscii
    Pause 0.7
    ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "/" & User & " Send (list,status,x,x-y(" & BlockAmt & "))" & RAscii
    Exit Sub
End If
'used to keep the user online
Timer45& = findwindow("_AOL_PALETTE", "America Online Timer")   'checks for the 45 minutes timere
If Timer45& > 0 Then    'if it's there then.......
   Timer45Button& = FindWindowEx(Timer45&, 0&, "_aol_icon", vbNullString)   'locates the button
   clickicon Timer45Button& 'clicks it
End If  'end if

bill = checkdeadmsg 'checks for a dead mail message
totalLbl = Total& + FrmMain.Finished.ListCount  'resets the total
userlbl = User$ 'resets the user label
pendinglbl = FrmMain.Pending.ListCount  'resets the pending count label
finishedlbl = FrmMain.Finished.ListCount    'resets the finished count label

If ListsReady = True Then   'checks to see if the lists are ready to be sent
    If Lists.ListCount = 0 Then 'checks if any one is waiting for a list
        ListsReady = False  'makes them not ready
        Exit Sub    'exits the sub
    End If  'ends the if
    If OPTIONS.listsend.Checked = False Then    'checks to make sure the user wants them sent
        ListsReady = False  'makes the lists not ready
        Exit Sub    'exits the sub
    End If  'ends the if
    statuslbl = "Sending List(s)"
    For x = 0 To Lists.ListCount - 1    'begins a for loop
        Place% = InStr(1, Lists.List(x), "-")   'sets a place in the listbox to extract the screen name from
        screenname$ = Left(Lists.List(x), Place% - 1)   'extracts the screen name
        Finished.AddItem screenname$ & "-" & "List", 0  'adds them to the finished box
        TillAolRestart = TillAolRestart + 1 'adds one to the variables to killsent mail and restart aol
        TillSentKill = TillSentKill + 1     'adds one to the variables to killsent mail and restart aol
        towho$ = towho$ & screenname$ & ", "    'adds the screen name and a comma for the names to send lists to, so that you send them all at once instead of one at a time
    Next x  'ends the for loop

    Lists.Clear 'clears the list box with pending lists in it so you don't send the list to the person again and again
    For x = 0 To Count2& - 1    'begins a loop based on the variable count2 which if you remember was sent when the user loaded mailz
opendaholist:   'a label to be sent to when the list isn't found after a period of time
    Call MailOpenEmailFlash(x)  'calls a sub to open the list with the index based on the loop which is based on the amount of lists you have
    If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'checks for the status window
        Call RunMenuByString("S&top Incoming Text") 'if it's found, call a sub to run the menu stop incoming text
        Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'closes the status window
    End If  'ends the if
    Current = Timer 'see above
    Do: DoEvents    'begins a do loop
        If STOPIT = True Then Exit Sub  'see above
        Call RunMenuByString("S&top Incoming Text") 'see above
        Handle& = FindSendWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List")   'calls a sub to find the list that you have just opened and sets it the handle
        If Timer - Current >= 2 Then GoTo opendaholist  'if 2 seconds has elapsed it goes to the label where you tried to open the list before
    Loop Until Handle& <> 0 And findfowardbutton(Handle&) <> 0 'loops until the mail is open
        aol& = findwindow("AOL Frame25", vbNullString)  'locates the aol window
        MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient window
        If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'checks for that annoying atatus window
            Call RunMenuByString("S&top Incoming Text") 'see above
            Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'closes that son of a bitch(ha ha got him!)
        End If  'ends the if
opendafuckinlist:   'label to come to when the window used to forward the mail isn't found(note the label name, i was annoyed at this section)
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        clickfowardbutton Handle&   'calls a sub to click the forward button
        Current = Timer 'see above
        Do: DoEvents    'yup you got it another do loop
            If Timer - Current >= 1 Then GoTo opendafuckinlist    'if 1/2 of a second has elapsed it goes to the section to click the forward button
            If STOPIT = True Then Exit Sub  'see above
            If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'fucking status window
                Call RunMenuByString("S&top Incoming Text") 'see above
                Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'kilt that bitch again!!!
                GoTo opendafuckinlist   'goes to the section to click the forward button again, after killing the status window you can instatly click the forward button instead of waiting 1/2 of a second, this is only done to maximize the speed
            End If  'ends the if
        Loop Until FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List") <> 0& 'loops until the window is found
        MailForward "จน~ท-.ธ  Unity Server(AOL4) -By- KiD List", towho$, " </u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity(AOL4) Server -By- KiD" & RAscii & Chr(13) & " " & LAscii & "Finished: " & finishedlbl + 1 & " Pending: " & pendinglbl & RAscii & Chr(13) & " " & LAscii & "< a href=" & Chr(34) & "http://parplex.com/chip/Unity.ZIP" & Chr(34) & "></u>Click Here To Download</a>" & RAscii & Chr(13) & MailMsg, True
        If OPTIONS.mnufastr.Checked = True Then 'checks to see the setting the user has selected
            Do Until FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List") = 0 'loops until there is no lists open
                Killwin FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List")  'locates the list
            Loop    'loops
            aol4_killwait   'calls a sub to kill wait on aol
        Else    'if slow mode is selected then.....
            Do: DoEvents    'hmmm wut's this? look familiar?, yup you got it another do loop
                If STOPIT = True Then Exit Sub  'see above
            Loop Until FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List") = 0&  'waits until the list is gone without killing it
        End If  'ends the if
        Killwin Handle& 'kills the other mail
        Do Until FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List") = 0 'loops until the mail is gonme
        Killwin FindForwardWindow("จน~ท-.ธ  Unity Server(AOL4) -By- KiD List")  'kills the mail
        Loop    'loops
        ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & Left(towho$, Len(towho$) - 2) & " List " & x + 1 & " Sent" & RAscii
        Pause ListPause 'pauses the designated tile for a list set by the user
        If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'status sone of a bitch again
            Call RunMenuByString("S&top Incoming Text") 'see above
            Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'uh oh, time to pull out a six shooter.... Bang got dat mofo
        End If  'ends if
    Next x  'ends loop to send the next list(wasn't that fun?!)

    ListsReady = False  'sets the lists to not be ready
    statuslbl = "Waiting"
End If  'ends the if

'checking for finds, thanks, status requests
    If FrmMain.Pending.ListCount = 0 Then Exit Sub  'checks for pending requests
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient
    If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'STATUS WINDOW DIE MUTCHA FUCKAH!
        Call RunMenuByString("S&top Incoming Text") 'see above
        Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'kilt dat mofo
    End If  'ends the if
    For x = 0 To FrmMain.Pending.ListCount - 1  'starts a loop to scan the pending listbox to check for the type of requests
        If InStr(Pending.List(x), "=") > 0 Then 'for different types of commands i just used different indecaters such as: -, >, =, etc... the = is used for find commands, I've always believed in sending finds as soon as they come in since i came out with premium gold, this line is used to find the point in which you will extract the screen name
            Place% = InStr(1, Pending.List(x), "=") 'this is used to actually find the point
            screenname$ = Left(Pending.List(x), Place% - 1) 'this gets all the text to the left of the point you found before(everything left of the =)
            find$ = Right(Pending.List(x), Len(Pending.List(x)) - Place%)   'this is to extract their command, this gets everything to the right of the point
            If InStr(LCase(find), "my ") > 0 Then   'most of the time you get assholes they search for: my dick, my cock, etc.... so if the server sees the word my in it it will ban the user, i was planning on making a place for the user to designate which words to ignore but i released this source prior to that
                ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " I Don't Take Bullshit Commands" & RAscii
                ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " You're Banned" & RAscii
                Pause 0.7   'waits after sending text to avoid being scrolled off
                Pending.RemoveItem x    'removes the bullshit request
                Ban.List2.AddItem screenname$   'adds the user to the ban list
                Exit Sub    'exits the sub without continuing
            End If  'ends the if
            For b = 0 To Mailz.ListCount - 1    'begins the loop to search for the string they requested
                If InStr(LCase(Mailz.List(b)), LCase(find$)) Then   'uses the instring command to check if they string they requested is found in the string of a mail in our mailz list, when we loaded the list we loaded them into the listbox with they subject of the mail, so basically it's searching the subject of the mail
                    dastring$ = dastring$ & b & ".)" & Chr(9) & Chr(9) & LTrim(Mailz.List(b)) & Chr(13) 'if the string is found in the subject it adds it to the string with the enter character which is character 13
                    found = found + 1   'makes your find +1, if you get an asshole to search for a space they'll get your whole list, if your whole list is over 500 items then the user would crash the server so we have to keep track of the number and limit it
                    If found >= 50 Then Exit For    'i've chosen to limit it to 50, this is a small amount i agree but 50 finds should be sufficient, i was planning on having the user select the amount of items per find but once again i released this source before this was implemented, it's a rather simple code though
                End If  ' ends the if
            Next b  'goes to the begining to check the next mail
            If dastring$ = "" Then  'if dastring = "" then nothing was found
                Pending.RemoveItem x    'removes the item
                ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " Find Was Not Found" & RAscii
                Pause 0.7   'waits to not be scrolled off
                Exit Sub    'exits without continuing
            Else    'self explanatory
                aol4_mail_send screenname$, "จน~ท-.ธ  Unity Server(AOL4) -By- KiD Find " & Request, dastring$, False    'calls the sub to send mail to the user with their finds
                Pause FindPause 'pauses for the duration designated by the user
                If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'checks for that damn status window
                    Call RunMenuByString("S&top Incoming Text") 'see above
                    Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'Yee Haw, one more down!
                End If  'ends the if
                ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " Find Was Found" & RAscii
            End If  'ends the if
            Pending.RemoveItem x    'removes the item
            Exit Sub    'exits the sub
        ElseIf InStr(Pending.List(x), ">") > 0 Then 'i've used the > character to designate some one saying thanks to the server, this checks for that character in the pending box
            Place% = InStr(1, Pending.List(x), ">") 'sets the point
            screenname$ = Left(Pending.List(x), Place% - 1) 'extracts the screen name
            ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " Your're Welcome" & RAscii
            Pause 0.7   'waits to not be scrolled off
            Pending.RemoveItem x    'removes the item
            Exit Sub    'exits the sub
        ElseIf InStr(Pending.List(x), "+") > 0 Then 'i've used the + character to designate a status request, this checks the pending list with the item equal to the loop variable for that character
            Place% = InStr(1, Pending.List(x), "+") 'sets the point in which to extract the screen name
            screenname$ = Left(Pending.List(x), Place% - 1) 'extracts the screen name
            For b = 0 To FrmMain.Pending.ListCount - 1  'begins a loop to find the total requests from the user
                If InStr(Pending.List(b), screenname$) > 0 Then pendnumforhim = pendnumforhim + 1   'checks each command to see if the users screen name is in the item with the index equal to the loop variable, if so then that's one more command
            Next b  'loops
            ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " You Have  " & pendnumforhim - 1 & "  Pending" & RAscii
            ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " There Is  " & pendinglbl - 1 & "  Total Pending" & RAscii
            Pause 0.7   'see above
            Pending.RemoveItem x    'removes the item
            Exit Sub    'see above
        End If  'exits the sub
    Next x  'checks the next item
    
'sends actual mail
    Place% = InStr(1, Pending.List(0), "-") 'i've used the - character for sending mail this sets the point
    screenname$ = Left(Pending.List(0), Place% - 1) 'extracts the screen name
    num& = Right(Pending.List(0), Len(Pending.List(0)) - Place%)    'extracts the number they've requested
    For x = 1 To FrmMain.Pending.ListCount - 1  'begins a loop to check for multiple requests
    If InStr(1, Pending.List(x), "-") <> 0 Then 'checks for the - character to see if the item with the index equal to the loop is also a mail
           Place% = InStr(1, Pending.List(x), "-")  'sets the point
           screenname2$ = Left(Pending.List(x), Place% - 1) 'extracts the screen name
           Num2& = Right(Pending.List(x), Len(Pending.List(x)) - Place%)    'extracts the command
           If Num2& = num& Then 'checks the commands to see if they are the same number
               toadd& = toadd& + 1  'this is used to see how many you will be adding to the total finished
               screenname$ = screenname$ & ", " & screenname2$  'adds the screen name to the screen name variable
               TillAolRestart = TillAolRestart + 1  'see above
               TillSentKill = TillSentKill + 1  'see above
               Finished.AddItem Pending.List(x), 0  'adds this command to the finished list
               Pending.RemoveItem x 'removes the command from the pending list
           End If   'ends the if
        End If  'ends the if
    Next x  'checks the next item
    statuslbl = "Sending " & num&
opendaho:
    checknumtoopen = checknumtoopen + 1 'variable used to see if the mail the server is trying to open can be opened, after atempting 10 times the server reads it as un openable and moves on
    If checknumtoopen >= 10 Then    ' checks to see if it has tried to open the mail 10 times
        Pending.RemoveItem 0    'removes their request
        ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & screenname$ & " Can't Open Mail " & num & " " & RAscii
        Exit Sub    'exits the sub
    End If  'ends the if
    MailOpenEmailFlash num& + (Count2 + MailzRdy&)  'callz a sub to open a mail from the flash mail box with an index = to the number they've requested plus their lists plus the mailz that you counted in their mailbox before running the flash session
    If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'yup you guessed it another status sob
        Call RunMenuByString("S&top Incoming Text") 'see above
        Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'killed it again, we're winning the battle agains the status window!
    End If  'ends the if
    Current = Timer 'see above
    Do: DoEvents    'begins a loop that will wait for the mail to be opened
        If STOPIT = True Then Exit Sub  'see above
        Call RunMenuByString("S&top Incoming Text") 'continues running the stop incoming text menu with a sub from the bas to minize the text loaded in the mail and to maximize the speed
        Handle& = FindSendWindow(Mailz.List(num))   'sets handle to = the window found with the caption of the mail being requested
        If Timer - Current >= 2 Then GoTo opendaho  'if 2 seconds have elapsed then it goes to the label to try to open it again
    Loop Until Handle& <> 0 And findfowardbutton(Handle&) <> 0 'continues checking for the mail until it's found
    Call RunMenuByString("S&top Incoming Text") 'see above
    Call RunMenuByString("S&top Incoming Text") 'see above, i've done it twice to try and stop the mail text from loading
    If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'the status window again
        Call RunMenuByString("S&top Incoming Text") 'yup you guessed it see above
        Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'another point for us in the way against the status window
    End If  'ends the if
opendabitch:    'once again i was annoyed when i made this label
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    clickfowardbutton Handle&   'calls a sub to click the forward button on the window with the handle in the variable handle&
    trackdabitch = 0    'I wrote all this code first then came and documented it later so i can't remember while this was put here, maybe it'll come up later, my guess is i used it before and now i removed it
    Current = Timer 'see above
    Do: DoEvents    'begins a do loop to wait for the next window to come up after clicking the forward window
    checkingit = checkdeadmsg   'see above
    If checkingit = "True" Then 'if the dead mail was found then...
        Pending.RemoveItem 0    'remove the item
        GoTo skipnextok 'skips the rest of the sending point
    End If  'ends the if
        If STOPIT = True Then Exit Sub  'see above
        If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'man is it just me or this status window undead?
            checkdeadmsg    'checks for a dead message again
            Call RunMenuByString("S&top Incoming Text") 'see above
            Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'aol keeps racking em up and i keep shooting them down
            GoTo opendabitch:   'goes and clicks fwd again
        End If  'ends the if
        If Timer - Current >= 1.5 Then GoTo opendabitch 'if 3 tenths of a second has elapsed then it trys to click forward again
    Loop Until FindForwardWindow(Mailz.List(num&)) <> 0 'calls a sub to find a window based on the caption
    MailForward Mailz.List(num&), screenname$, " </u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "Unity(AOL4) Server -By- KiD" & RAscii & Chr(13) & " " & LAscii & "Finished: " & finishedlbl + 1 & " Pending: " & pendinglbl & RAscii & Chr(13) & " " & LAscii & "< a href=" & Chr(34) & "http://parplex.com/chip/Unity.ZIP" & Chr(34) & "></u>Click Here To Download</a>" & RAscii & Chr(13) & MailMsg, True
    Call SendMessage(FindSendWindow(Mailz.List(num&)), WM_CLOSE, 0&, 0&)    'close the first mail with the forward button
    aol4_killwait   'uses a sub to kill the wait
    Finished.AddItem Pending.List(0), 0 'adds the item to the finished list
    TillAolRestart = TillAolRestart + 1 'see above
    TillSentKill = TillSentKill + 1 'see above
    Pending.RemoveItem 0    'removes the item
    If InStr(SentMsg, "^") > 0 Then 'checks the message to be sent to the chat room set by the user for a character i used to let the user notify the room of the total sent
        Place% = InStr(1, SentMsg, "^") 'sets the point
        lMsg$ = Left(SentMsg, Place% - 1)   'gets the left of the message
        rmsg$ = Right(SentMsg, Len(SentMsg) - Place%)   'gets the right of the message
        If OPTIONS.mnunotify.Checked = True Then ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & screenname$ & " ท " & num& & " " & lMsg & " " & finishedlbl + 1 & " " & rmsg & "<font face=" & Chr(34) & "verdana" & Chr(34) & ">" & RAscii
        GoTo skipnextok 'skips the next send chat
    End If
    If OPTIONS.mnunotify.Checked = True Then ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & screenname$ & " ท " & num& & " " & SentMsg & "<font face=" & Chr(34) & "verdana" & Chr(34) & ">" & RAscii
skipnextok: 'looky here we are
    statuslbl = "Waiting"
    If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then 'last mofo!
        Call RunMenuByString("S&top Incoming Text") 'see above
        Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)   'kills it for the last time whoo hoo!!!!!!!!!!!!!!!!!
    End If  'ends the if
    If MailzPause <> 0 Then Pause MailzPause    'checks to see if there is a pause designated by the user and pause
    checkdeadmsg    'see above
End Sub
Private Sub Timer3_Timer()
    ListsReady = True   'sets the lists ready = to true
End Sub
