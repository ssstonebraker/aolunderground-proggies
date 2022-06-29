VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H80000007&
   Caption         =   "Room Buster"
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form16"
   ScaleHeight     =   1740
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   1080
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Text            =   "Fate"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bust"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Tea$ = "-æ-[ HyPO ™ ]-æ-Busted Into " + Text1 + "In " + Text2 + " Secs."
 fnt$ = "10"
A = Len(Tea$)
For w = 1 To A Step 4
    R$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    S$ = Mid$(Tea$, w + 2, 1)
    T$ = Mid$(Tea$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & R$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRBb = P$
  
Do
Timer1.Enabled = True
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   Timer1.Enabled = False
  MsgBox "Please Close the Chat room before trying to bust into another room", vbCritical, "HyPO"
   Exit Sub
   End If
 
KeyWord ("aol://2719:2-2-" + Text1)

'
'
waitforok
'
'
 
 
DoEvents

Loop

End Sub

Private Sub Command2_Click()
Form16.Hide
Unload Form16
Timer1.Enabled = False


End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form16
End Sub

Private Sub Timer1_Timer()
 
 TimeOut 0.000001
Text2.Text = Val(Text2.Text) + 0.000001
 
 
End Sub
