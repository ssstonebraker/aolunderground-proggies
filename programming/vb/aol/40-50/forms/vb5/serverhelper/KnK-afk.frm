VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   " AFK Bot"
   ClientHeight    =   1635
   ClientLeft      =   1425
   ClientTop       =   1080
   ClientWidth     =   2160
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "KnK-afk.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "X"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   2160
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "For the hell of it!!!"
      Top             =   960
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1440
      Top             =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Stop"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Your SN or change to handle"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "KnK Founders AFK Bot"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   750
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub afk_Click()

End Sub

Private Sub Command1_Click()
Timer1 = True
Timer2.interval = 0
End Sub

Private Sub Command2_Click()
Timer2.interval = 1
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol4" Then
SendChat BlackGreenBlack("") + Text2.text + (" is now Back! I was away for «(") + Label1 + (" mins)»")
Timer1.Enabled = False
Label1.Caption = 0
End If
If aversion$ = "aol95" Then
AOLChatSend ("") + Text2.text + (" is now Back! I was away for «(") + Label1 + (" mins)»")
Timer1.Enabled = False
Label1.Caption = 0
End If
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text2.text = UserSN()
StayOnTop Me
End Sub


Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Timer1_Timer()
DoEvents
'AOL95
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol95" Then
AOLChatSend (Text2.text + " is Away «(Reason: ") + Text1.text + (")» From « ") + Label5.Caption + (" »")
Label1.Caption = Val(Label1) + 1
End If
'AOL 4.o
If aversion$ = "aol4" Then
Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
If Err Then
MsgBox "KnK Founders Server Helper could not find the file'KnK.ini'  Either it wasnt in the C:\Program Files\Server Helper folder or it was missing.  Please goto http://knk.tierranet.com/serv  to dl a full clean copy.", vbExclamation, "Error"
End If
If Color$ = "bgb" Then
SendChat BlackGreenBlack(Text2 + " is Away «(Reason: ") + Text1 + (")» From « ") + Label5.Caption + (" »")
Label1.Caption = Val(Label1) + 1
End If
If Color$ = "bbb" Then
SendChat BlackBlueBlack(Text2 + " is Away «(Reason: ") + Text1 + (")» From « ") + Label5.Caption + (" »")
Label1.Caption = Val(Label1) + 1
End If
End If
End Sub


Private Sub Timer2_Timer()
Label5.Caption = Format$(Now, "h:mm:ss AM/PM")
End Sub


Private Sub Timer3_Timer()

End Sub
