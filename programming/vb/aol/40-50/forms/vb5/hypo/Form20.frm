VERSION 5.00
Begin VB.Form Form20 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warez Requester"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Text            =   "HyPO 2.0"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Request"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3015
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 BlueFade ("HyPO request Bot by ToaST")
timeout 0.5
BlueFade ("Requesting """ + Text2 + """")
timeout 0.5
BlueFade ("If you have it type""" + Text2 + """")

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Timer1_Timer()
Do
Text1 = LastChatLine
Text2.Text = LCase(Text2)
Text1.Text = LCase(Text1)
If Text1 = Text2 Then
   

SendChat SNFromLastChatLine + " Can you send " + Text2
timeout 1
'
'
Tea$ = "···÷••(¯`·._=-HyPO ver.¹·º 4 AOL 4 -=_.·´¯)••÷"

SendChat Tea$

IMKeyword SNFromLastChatLine, "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">Can you send " + Text2

End If
DoEvents
Loop
End Sub
