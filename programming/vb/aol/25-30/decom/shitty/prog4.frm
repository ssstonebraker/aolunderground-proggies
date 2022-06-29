VERSION 4.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deicide's Prog"
   ClientHeight    =   2160
   ClientLeft      =   1815
   ClientTop       =   2115
   ClientWidth     =   4095
   ForeColor       =   &H00000000&
   Height          =   2565
   Icon            =   "PROG4.frx":0000
   KeyPreview      =   -1  'True
   Left            =   1755
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4095
   Top             =   1770
   Width           =   4215
   Begin VB.CommandButton Command20 
      Caption         =   "Im stuff"
      Height          =   195
      Left            =   3180
      TabIndex        =   33
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Charactr"
      Height          =   195
      Left            =   2430
      TabIndex        =   32
      Top             =   660
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Left            =   390
      Top             =   2700
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Punter"
      Height          =   195
      Left            =   2430
      TabIndex        =   26
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Kill Wait"
      Height          =   195
      Left            =   2430
      TabIndex        =   25
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Im's On"
      Height          =   195
      Left            =   1680
      TabIndex        =   22
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Im's Off"
      Height          =   195
      Left            =   930
      TabIndex        =   21
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hide"
      Height          =   195
      Left            =   2430
      TabIndex        =   2
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load it"
      Height          =   195
      Left            =   930
      TabIndex        =   1
      Top             =   30
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Left            =   2430
      Top             =   2280
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AOL Dir"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Change "
      Height          =   195
      Left            =   930
      TabIndex        =   15
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "UpChat"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Exit"
      Height          =   195
      Left            =   3180
      TabIndex        =   13
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dialer"
      Height          =   195
      Left            =   3180
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Advrtse"
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Chat"
      Height          =   195
      Left            =   3180
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Count"
      Height          =   195
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Bust In"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Metal"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   450
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Text"
      Height          =   195
      Left            =   930
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Minimize"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   315
      Left            =   2220
      TabIndex        =   31
      Top             =   3270
      Width           =   1365
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   30
      Top             =   1650
      Width           =   3855
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   2250
      TabIndex        =   29
      Top             =   2670
      Width           =   1635
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   270
      TabIndex        =   28
      Top             =   2310
      Width           =   1485
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   0
         weight          =   400
         size            =   6.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   1830
      Width           =   3915
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   90
      TabIndex        =   24
      Top             =   1620
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   345
      Left            =   2250
      TabIndex        =   23
      Top             =   2310
      Width           =   1725
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1170
      TabIndex        =   20
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   345
      Left            =   360
      TabIndex        =   19
      Top             =   2310
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   18
      Top             =   1320
      Width           =   705
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Show
Form4.Hide
End Sub


Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Do this FIRST! Or else stuff will mess up"
Label9.Caption = "bob"
End Sub


Private Sub Command10_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Bust's you into a Private room"
Label9.Caption = "bob"
End Sub


Private Sub Command11_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
NER% = FindChildByTitle(MDI%, "New Mail")
If NER% <> 0 Then GoTo J
TOO% = FindChildByClass(AOL%, "AOL Toolbar")
NEE% = FindChildByClass(TOO%, "_AOL_Icon")
D = SendMessageByNum(NEE%, WM_LBUTTONDOWN, 0, 0)
D = SendMessageByNum(NEE%, WM_LBUTTONUP, 0, 0)
Do
c% = DoEvents()
NER% = FindChildByTitle(MDI%, "New Mail")
Loop Until NER% <> 0
TRE% = FindChildByClass(NER%, "_AOL_Tree")
Do
u = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
Pause (2)
f = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
Loop Until u = f
J:
TRE% = FindChildByClass(NER%, "_AOL_Tree")
f = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
MsgBox "You have " & f & " Poopy'z of mail!", 64, "Eraser Own'z Me cause he's kewlest leet0 guy in the world!"


End Sub

Private Sub Command11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Doesn't work...yet"
Label9.Caption = "bob"
End Sub


Private Sub Command12_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then
MsgBox "Aol could not be found"
Exit Sub
End If
Call RunMenuByString(AOL%, "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Private Sub Command12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Get's rid of the hourglass crap"
Label9.Caption = "bob"
End Sub


Private Sub Command13_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then
MsgBox "Aol could not be found"
Exit Sub
End If
Call RunMenuByString(AOL%, "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub


Private Sub Command13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Makes it so you can do stuff and Upload"
Label9.Caption = "bob"
End Sub


Private Sub Command14_Click()
Form4.Hide
Form9.Show
End Sub

Private Sub Command14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Changes AOL's title to something else"
Label9.Caption = "bob"
End Sub


Private Sub Command15_Click()
End
End Sub

Private Sub Command15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "If you can't figure this out, shoot yourself"
Label9.Caption = "bob"
End Sub


Private Sub Command16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Punt's people....sometimes :)"
Label9.Caption = "bob"
End Sub

Private Sub Command17_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then
MsgBox "Aol could not be found"
Exit Sub
End If
Call RunMenuByString(AOL%, "Send An Instant Message")
MDI% = FindChildByClass(AOL%, "MDIClient")
Do: DoEvents
IMt% = FindChildByTitle(MDI%, "Send Instant Message")
Loop Until IMt% <> 0

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Edit"
Call AOLSetText(hChild, "$im_off")

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "RICHCNTL"
Call AOLSetText(hChild, "-=Deicide Rulez Me=-")

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Call Sendclick(hChild)
SendMessage IMt%, WM_CLOSE, 0, 0
End Sub

Private Sub Command17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Turn Im's off in a hurry"
Label9.Caption = "bob"
End Sub


Private Sub Command18_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then
MsgBox "Aol could not be found"
Exit Sub
End If
Call RunMenuByString(AOL%, "Send An Instant Message")
MDI% = FindChildByClass(AOL%, "MDIClient")
Do: DoEvents
IMt% = FindChildByTitle(MDI%, "Send Instant Message")
Loop Until IMt% <> 0

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Edit"
Call AOLSetText(hChild, "$im_on")

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "RICHCNTL"
Call AOLSetText(hChild, "-=Deicide Rulez Me=-")

hChild = GetWindow(IMt%, GW_CHILD)
rooms = GetClass(hChild)
Label6.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label6.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label6.Caption = "_AOL_Icon"
Call Sendclick(hChild)
SendMessage IMt%, WM_CLOSE, 0, 0
End Sub

Private Sub Command19_Click()
Form10.Show
Form4.Hide
End Sub

Private Sub Command16_Click()
Form4.Hide
Form15.Show
End Sub

Private Sub Command18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Turns your IM's on in a hurry"
Label9.Caption = "bob"
End Sub

Private Sub Command19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Easy to use extended characters"
Label9.Caption = "bob"
End Sub

Private Sub Command2_Click()
        mypath = CurDir
        Open mypath & "\deicide.ini" For Random As #1
        Get #1, 1, Path$
        If Path$ <> "" Then
        AmericaStart = Shell(Path$, 1)
        Close #1
        Else
        MsgBox "Go to AOL Dir and choose your default Aol Path"
        Exit Sub
        End If
End Sub


Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Load's up AOL for you..."
Label9.Caption = "bob"
End Sub


Private Sub Command20_Click()
Form4.Hide
Form3.Show
End Sub

Private Sub Command20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Scrolling IM's and other things"
Label9.Caption = "bob"
End Sub


Private Sub Command21_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMs% = FindChildByTitle(MDI%, "Send Instant Message")

hChild = GetWindow(IMs%, GW_CHILD)
rooms = GetClass(hChild)
Label4.Caption = rooms
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label4.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label4.Caption = "RICHCNTL"
ca = AOLGetText(hChild)
MsgBox ca

End Sub

Private Sub Command3_Click()
a% = FindWindow("AOL Frame25", vbNullString)  'Find AOL
    

If Command3.Caption = "Hide" Then
  
    X = ShowWindow(a%, SW_HIDE)
    Command3.Caption = "Show"
    Exit Sub
End If
If Command3.Caption = "Show" Then
    
    X = ShowWindow(a%, SW_SHOW)
    Command3.Caption = "Hide"
    Exit Sub
End If

End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Hides AOL, quite useful!"
Label9.Caption = "bob"
End Sub


Private Sub Command4_Click()
Form5.Show
Form4.Hide
End Sub


Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "A nice auto Dialer program"
Label9.Caption = "bob"
End Sub


Private Sub Command5_Click()
'String(950, 13)
'SW_minimze
End Sub


Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Makes this form smaller...much smaller"
Label9.Caption = "bob"
End Sub


Private Sub Command6_Click()
Form4.Hide
Form7.Show
End Sub


Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Text Pad w/ Encryption stuff"
Label9.Caption = "bob"
End Sub


Private Sub Command7_Click()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
If room% = 0 Then
MsgBox "Room could not be found"
Exit Sub
End If
View% = FindChildByClass(room%, "_AOL_View")
If View% = 0 Then
MsgBox "Room could not be verified"
Exit Sub
End If
Z = aolhwnd = room%
hChild = GetWindow(room%, GW_CHILD)
hchild2 = hChild
rooms = GetClass(hChild)
Label4.Caption = rooms

Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
Bob = GetClass(hChild)
Label4.Caption = Bob
Y = aolhwnd = hChild
Loop Until Label4.Caption = "_AOL_Icon"

Do: DoEvents
hchild2 = GetWindow(hchild2, GW_HWNDNEXT)
rooms = GetClass(hchild2)
Label4.Caption = rooms
Y = aolhwnd = hchild2
Loop Until Label4.Caption = "_AOL_Edit"
Call sendtext(hchild2, "___________________________________________")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hchild2)
Loop Until a = ""
Call sendtext(hchild2, "»»»   I am using GÕÐ§ßÃÑE by Deicide402™   «««")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hchild2)
Loop Until a = ""
Call sendtext(hchild2, "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯")
Do: DoEvents
Call SendCharNum(hChild, 13)
a = AOLGetText(hchild2)
Loop Until a = ""
End Sub

Private Sub m_Click()
AppActivate "America"
America% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(America%, "MDIClient")
roomma% = FindChildByTitle(MDI%, "Welcome")
Z = aolhwnd = roomma%
hChild = GetWindow(roomma%, GW_CHILD)
a = GetClass(hChild)
p = "_AOL_Glyph"
Text1.text = a
If Text1.text = "_AOL_Edit" Then
Z = aolhwnd = hChild
Call sendtext(hChild, Text1.text)
Else
Do: DoEvents
hChild = GetWindow(hChild, GW_HWNDNEXT)
a = GetClass(hChild)
Text1.text = a
Loop Until Text1.text = "_AOL_Edit"
Z = aolhwnd = hChild
Call sendtext(hChild, Text1.text)
End If

End Sub


Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Tell everybody you are using this"
Label9.Caption = "bob"
End Sub

Private Sub Command8_Click()
Form4.Hide
Form3.Show
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Makes a chat window that StaysOnTop"
Label9.Caption = "bob"
End Sub


Private Sub Command9_Click()
Form4.Hide
Form12.Show
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Really lame scrolling stuff..."
Label9.Caption = "bob"
End Sub


Private Sub Form_Load()
Call StayOnTop(Form4)
Timer1.interval = 5
Timer2.Enabled = False
mypath = CurDir
Bob = Dir(mypath & "\Deicide.ini")
If Bob = "" Then
    Msg = "K now here it is...to get started go to AOL Dir in the upper left corner, you should really do this or else alot of things won't work and you will get annoying error messages..that's about it.                                                                                                 -= Deicide =-                                     -=ShdwODeath=-"
    Style = vbOKOnly
    Title = "Hell doth welcome thee..."
    Response = MsgBox(Msg, Style, Title)
End If
Label7.Caption = "GÕÐ§ßÃÑE                        ßý                        ÐËIÇIÐË"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "bob"
Label5.Caption = "Umm...this does nothing."
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Label10_Change()
Timer2.Enabled = True
Timer2.interval = 5000
ass = "Munky"
asss = "Curve463"
Bob$ = "Contributions from " & ass & " and " & asss & "."
label8.Caption = Bob$
MyValue = 8
Do: DoEvents
MyValue = MyValue + 1
Pause (2.5)
label8.ForeColor = QBColor(MyValue)
If MyValue = 15 Then
MyValue = 8
End If
Loop
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "Yep, it's the time...I know"
Label9.Caption = "bob"
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.Caption = "bob"
Label5.Caption = "Ohh look you broke it!!!! :)"
End Sub


Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "-=Don't mess with these People=-"
Label9.Caption = "bob"
End Sub


Private Sub label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.Caption = "-=Don't mess with these People=-"
End Sub


Private Sub Label9_Change()
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                       ßý                       ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                      ßý                      ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                     ßý                     ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                    ßý                    ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                   ßý                   ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                  ßý                  ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                 ßý                 ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE                ßý                ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE               ßý               ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE              ßý              ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE             ßý             ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE            ßý            ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE           ßý           ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE          ßý          ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE         ßý         ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE        ßý        ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE       ßý       ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE      ßý      ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE     ßý     ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE    ßý    ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE   ßý   ÐËIÇIÐË"
Pause (0.2)
Label7.Caption = "GÕÐ§ßÃÑE  ßý  ÐËIÇIÐË"
Pause (0.5)
Label7.ForeColor = QBColor(0)
Pause (0.4)
Label7.ForeColor = QBColor(15)
Pause (0.3)
Label7.ForeColor = QBColor(12)
Pause (1.5)
Label10.Caption = "hi"
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format(Now, "Long Time")
End Sub


Private Sub Timer2_Timer()
'1st time
X = Label1.Caption
If X = "" Then
X = 0
End If
X = X + 1
If X = 1 Then
Label7.ForeColor = QBColor(4)
Do: DoEvents
MyValue = Int((7 - 1 + 1) * Rnd + 1)
Loop Until MyValue <> 4
Label12.ForeColor = QBColor(MyValue)
Label12.Caption = "GÕÐ§ßÃÑE  ßý  ÐËIÇIÐË"
Label11.Caption = MyValue
End If

If X = 2 Then
Label7.ForeColor = QBColor(0)
MyValue = Label11.Caption
MyValue = MyValue + 8
Label12.ForeColor = QBColor(MyValue)
End If

If X = 3 Then
MyValue = Label11.Caption
Label12.ForeColor = QBColor(MyValue)
Label7.ForeColor = QBColor(4)
End If

If X = 4 Then
Label12.Caption = ""
Label7.ForeColor = QBColor(12)
End If
If X = 5 Then
X = 0
End If
Label1.Caption = X
End Sub
