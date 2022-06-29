VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Room Finder"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form13"
   ScaleHeight     =   1050
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   1200
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   2760
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Room Name"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Room Finder"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Text3 = AOLFindRoom
End Sub

Private Sub Command3_Click()
Unload Form13
Form13.Hide

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form13
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Not FindChatRoom() = "" Then
SendChat "-æ-[ HyPO™ ]-æ-Found a " + Text1 + " Room"

Text2 = ""
Text1 = ""
Text3 = ""
Timer2.Enabled = False
Timer1.Enabled = False
Exit Sub
End If
KeyWord ("aol://2719:2-2-" + Text1 + Text2)
'''
SendKeys " "
SendKeys " "

''

''

Text2.Text = Val(Text2.Text) + 2
TimeOut 1

If Text2.Text = "2" Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
If Not FindChatRoom() = "" Then
SendChat "-æ-[ HyPO™ ]-æ-Found a " + Text1 + " Room"
Text1.Text = ""
Text2.Text = ""
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not FindChatRoom() = "" Then
SendChat "-æ-[ HyPO™ ]-æ-Found a " + Text1 + " Room"
 
Text2 = ""
Text1 = ""
Text3 = ""
Timer2.Enabled = False
Timer1.Enabled = False
Exit Sub
End If
Text3 = AOLFindRoom
On Error Resume Next
KeyWord ("aol://2719:2-2-" + Text1 + Text2)
AppActivate "America  Online"
TimeOut 1
''
SendKeys " "
SendKeys " "
 ''
 ''
 
Text2.Text = Val(Text2.Text) + 1
If Not FindChatRoom() = "" Then
SendChat "-æ-[ HyPO™ ]-æ-Found a " + Text1 + " Room"

Text2 = ""
Text1 = ""
Text3 = ""
Timer2.Enabled = False
Timer1.Enabled = False
Exit Sub
End If
End Sub
