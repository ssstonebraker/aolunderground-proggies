VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guess"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1755
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   1755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guess Bot"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text2 = Int((20 * Rnd) + 1)
Tea$ = "-æ-[ HyPO ™ ]-æ-Guess Bot.. Guess a number 1-20"

fnt$ = "10"
a = Len(Tea$)
For w = 1 To a Step 4
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
ToaST$ = "···÷••(¯`·._By ToaST _.·´¯)••÷ "

a = Len(ToaST$)
For w = 1 To a Step 4
    ToaSTR$ = Mid$(ToaST$, w, 1)
    ToaSTu$ = Mid$(ToaST$, w + 1, 1)
    ToaSTs$ = Mid$(ToaST$, w + 2, 1)
    ToaSTT$ = Mid$(ToaST$, w + 3, 1)
    ToaSTP$ = ToaSTP$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sup><b>" & ToaSTR$ & "</sup></b>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & ToaSTu$ & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & "><sub>" & ToaSTs$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#EE2C2C" & Chr$(34) & ">" & ToaSTT$
Next w
WavYChaTRBb = ToaSTP$
'···÷••(¯`·._   _.·´¯)••÷
SendChat WavYChaTRBb
 
 On Error Resume Next
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False

Form18.Hide
Unload Form18


End Sub

Private Sub Timer1_Timer()
  
Text1 = LastChatLine
Text1.Text = LCase(Text1)
If Text1 = (Text2) Then
SendChat SNFromLastChatLine + "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "> Good Job you Got it!!"
SendChat "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">Game Over"
SendChat WavYChaTRBb
SendChat "The number was " + Text1
Timer1.Enabled = False
Exit Sub
End If
If IsNumeric(Text1) = (Text2) Then
SendChat SNFromLastChatLine + " Wrong Number But Close"
TimeOut 0.5
End If

 
End Sub
