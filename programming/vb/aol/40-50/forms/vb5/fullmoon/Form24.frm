VERSION 5.00
Begin VB.Form Form24 
   BorderStyle     =   0  'None
   Caption         =   "AFK Bot"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   LinkTopic       =   "Form14"
   Picture         =   "Form24.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   15000
      Left            =   6120
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   2280
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   1200
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5520
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "Gone To the Bathroom"
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear Lists"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "minutes"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "You've Been AFK for"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Names"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Messages :"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason AFK :"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AFK Bot"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop AFK"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Start AFK"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChat "Full Moons AFK bot"
SendChat "I'm AFK because : " + ("" + Text1.Text + "")
SendChat "I have been AFK for " + ("" + Label9.Caption + "") + " minutes"
SendChat "Type '/AFK' and a message to leave me a message"
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
SendChat "Full Moons AFK bot"
Call Timeout(0.1)
SendChat ("I'm AFK because : ") + ("" + Text1.Text + "")
Call Timeout(0.1)
SendChat ("I have been AFK for ") + ("" + Label9.Caption + "") + (" minutes")
Call Timeout(0.1)
SendChat "Type '/AFK' and a message to leave me a message"
Timer1.Enabled = True
Timer2.Enabled = True
Call Pause(5)
Timer3.Enabled = True
End Sub

Private Sub Label11_Click()
Form24.WindowState = 1
End Sub

Private Sub Label12_Click()

Unload Me
End Sub

Private Sub Label13_Click()
List1.Clear
List2.Clear
End Sub

Private Sub Label2_Click()
SendChat "Full moon's AFK bot is now off!"
SendChat "I was AFk for " + ("" + Label9.Caption + "") + " minutes"

Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Timer1_Timer()
Label9.Caption = Label9.Caption + 1
End Sub

Private Sub Timer2_Timer()
Text2 = LCase(LastChatLine)
Text2.Text = LCase(Text2.Text)
If Text2 Like ("/afk*") Then
List1.AddItem SNFromLastChatLine
List2.AddItem LastChatLine
SendChat ("" + SNFromLastChatLine + "") + (", Your message has been recorded")
End If
End Sub

Private Sub Timer3_Timer()
Command1_Click
End Sub
