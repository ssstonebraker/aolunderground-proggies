VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voter"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   2520
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Text            =   "Question"
      Top             =   120
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Answered No"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Answered YES"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Tea$ = "-æ-[ HyPO ™ ]-æ-Voter Bot.. Type Yes or No"

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
SendChat WavYChaTRBb
 

TimeOut 0.5
'

'
'
ToaST$ = "And The Question IS.. """ + Text2 + """"

A = Len(ToaST$)
For w = 1 To A Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup><b>" & ToaSTR$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
SendChat WavYChaTRBb

Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer1.Enabled = False
Form19.Hide
Unload Form19




End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
counterz$ = List2.ListCount
Counter$ = List1.ListCount
SendChat "-æ-[ HyPO ™ ]-æ-" + counterz$ + " Votes for no"
SendChat "-æ-[ HyPO ™ ]-æ-" + Counter + " Votes for yes"
Timer1.Enabled = False
Timer1.Enabled = False

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Timer1_Timer()
 
Text1 = LastChatLine
Text1.Text = LCase(Text1)
If Text1 = ("no") Then
SendChat SNFromLastChatLine + "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "> Vote for No Counted"
List2.AddItem SNFromLastChatLine
TimeOut 0.9
End If

If Text1 = ("yes") Then
SendChat SNFromLastChatLine + "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "> Vote for Yes Counted"
TimeOut 0.9
List1.AddItem SNFromLastChatLine
End If
DoEvents

 
End Sub
