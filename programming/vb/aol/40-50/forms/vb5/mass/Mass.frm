VERSION 5.00
Object = "{30ACFC93-545D-11D2-A11E-20AE06C10000}#1.0#0"; "VB5CHAT2.OCX"
Begin VB.Form Mass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   8040
   ClientTop       =   2040
   ClientWidth     =   4335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   30
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   29
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4095
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   3855
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   16
         Text            =   "SN TO ADD"
         ToolTipText     =   "SN to add to the mm"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Text            =   "10"
         ToolTipText     =   "Mins till u mm"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "TRIGGER WERD"
         ToolTipText     =   "Trigger werd for mm"
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Mailz u are gonna mm"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         ToolTipText     =   "Lamers on the mm"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add Mailz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Adds your mailz to a list box"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add SN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Adds a SN to the MM"
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MM now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Starts the MM now"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear lists"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Claers mail list and peeps on the mm"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         ToolTipText     =   "Duh what u think"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Starts the fuggin Bot"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Trigger werd"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "SN to add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Min(s) Till MM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Mailz U gonna MM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Lamers U gonna MM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   3840
      Top             =   4200
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   3840
      Top             =   4200
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3840
      Top             =   4200
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3840
      Top             =   4200
   End
   Begin VB5Chat2.Chat Chat1 
      Left            =   2160
      Top             =   4320
      _ExtentX        =   3969
      _ExtentY        =   2170
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PeaCe OuT MMer By RaVaGe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FF80FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C000C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   26
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800080&
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400040&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Mass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chat1_ChatMsg(Screen_Name As String, What_Said As String)
If LCase(What_Said) Like LCase(Text1.text) Then
List1.AddItem "(" & Screen_Name & ")"
For Y = 0 To List1.ListCount - 1
tt$ = tt$ + List1.List(Y)
Next Y
Timeout (0.01)
Text4.text = tt$
KillDupes List1
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: BoldRedBlackRed Screen_Name & " Bet u think your lucky just cuz u r on the MM"
Case 2: BoldRedBlackRed Screen_Name & " Your name has been added to the MM"
Case 3: BoldRedBlackRed Screen_Name & " You will never get on the MM with an attitude like that"
Case 4: BoldRedBlackRed Screen_Name & " Man your lame but you're on the MM"
Case 5: BoldRedBlackRed Screen_Name & " Your momma had better luck than u did but u made on the list"
Case 6: BoldRedBlackRed Screen_Name & " Blah - blah - blah ... Your on the MM"
Case 7: BoldRedBlackRed Screen_Name & " GUESS WHAT! I don't care"
Case Else: BoldRedBlackRed Screen_Name & " U should be proud"
End Select
End If
If LCase(What_Said) Like LCase("remove me") Then
 For i = 0 To List1.ListCount - 1
If LCase(List1.List(i)) = LCase("(" & Screen_Name & ")") Then
List1.RemoveItem i
End If
Next i
 For Y = 0 To List1.ListCount - 1
tt$ = tt$ + List1.List(Y)
Next Y
Timeout (0.01)
Text4.text = tt$
Dim l003B As Variant
Randomize Timer
l003B = Int(Rnd * 8)
Select Case l003A
Case 1: BoldRedBlackRed Screen_Name & " Yeah well you suck anyway"
Case 2: BoldRedBlackRed Screen_Name & " Why did u ask to be on the mm to start with?"
Case 3: BoldRedBlackRed Screen_Name & " You will never get on the MM again with an attitude like that"
Case 4: BoldRedBlackRed Screen_Name & " Man you're lame "
Case 5: BoldRedBlackRed Screen_Name & " Your momma should whoop your A$$ for that"
Case 6: BoldRedBlackRed Screen_Name & " Blah - blah - blah ... You're off the MM"
Case 7: BoldRedBlackRed Screen_Name & " GUESS WHAT! I don't care"
Case Else: BoldRedBlackRed Screen_Name & " Up yours"
End Select
End If
End Sub

Private Sub Command1_Click()
Timer2.Enabled = True
Timer3.Enabled = True

'Starts the timmer so
'it can scroll when it should
Chat1.ScanOn
SendChat "•÷·· · ··÷• <U><Font color= #FF0000>Pe<Font color= #FF0450>ace O<Font color= #FF0800>ut MMer</u> "
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Mass mailing " & Label14.Caption & " mailz"
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• In " & Text2.text & " Min(s) "
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Say " & Chr(34) & Text1.text & Chr(34) & " to be added"
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Say " & Chr(34) & "remove me" & Chr(34) & " to be removed"
End Sub

Private Sub Command2_Click()
Timer2.Enabled = False 'Stops the timmer
Chat1.ScanOff
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Bot Wuz Stopped"
Timeout 0.4
SendChat "•÷·· · ··÷• <U><Font color= #FF0000>Pe<Font color= #FF0450>ace O<Font color= #FF0800>ut</u> "
End Sub

Private Sub Command3_Click()
List1.Clear
List2.Clear
Text4.text = ""
'Duh it clears the lists and text box
End Sub

Private Sub Command4_Click()
Timer4.Enabled = True
End Sub

Private Sub Command5_Click()
List1.AddItem "(" & Text3.text & ")"
For Y = 0 To List1.ListCount - 1
tt$ = tt$ + List1.List(Y)
Next Y
Timeout (0.01)
Text4.text = tt$
Text3.text = ""
'Ok text3 is the sn lamer is addin
'text4 is where we add it to the list and text
'so we can mass mail from the text
'By sayin all that we added a name to a list and a text
'the cleared it from the one that the user added the name to
End Sub

Private Sub Command6_Click()
AoL40_Mail_OpenMailBox
Timeout 0.3
AoL40_Mail_AddNewToListBox List2
Timeout 0.3
Label14.Caption = List2.ListCount
AoL40_Mail_Minimize_MailBox
'Hmmmm all that to add mail to a list
'Dam that's alota werk
End Sub

Private Sub Command7_Click()
Me.Hide
End Sub

Private Sub Command8_Click()
Me.WindowState = 1
End Sub

Private Sub Form_Load()
Timer3.Enabled = False
Timer4.Enabled = False
Text1.text = UserSN & " (V) (V)"
'gives us our trigger werd
Text3.text = UserSN
'Puts a name in the text where
'the user adds a sn to add
StayOnTop Me
'Keeps this form ontop
Timer2.Enabled = False
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Timer2_Timer()
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Mass mailing " & Label14.Caption & " mailz"
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• In " & Text2.text & " Min(s) "
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·÷• Say " & Chr(34) & Text1.text & Chr(34) & " to be added"
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Say " & Chr(34) & "remove me" & Chr(34) & " to be removed"
End Sub

Private Sub Timer3_Timer()
Label17.Caption = Label17.Caption + 1
If Label17.Caption = "60" Then
Label17.Caption = "0"
Text2.text = Text2.text - 1
If Text2.text = "0" Then
Timer4.Enabled = True
End If
End If
End Sub

Private Sub Timer4_Timer()
Timer3.Enabled = False
Timer5.Enabled = True
Timer2.Enabled = False
Chat1.ScanOff
IMBuddy "$IM_OFF", "IM OFF" 'shuts ims off
SendChat "•÷·· · ··÷• <U><Font color= #FF0000>Pe<Font color= #FF0450>ace O<Font color= #FF0800>ut MMer</u> "
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Mailing Now"
Start:
AoL40_Mail_OpenNewMailNumber Label18.Caption
Timeout 1
Call AoL40_Mail_CLicKForward
Timeout 1
CLicKSendAndForwardMail Text4.text
Timeout 1
AoL40_Mail_CloseEmail
CLicKKeepAsNew
Timeout 3.76767676
Picture1.Visible = True
Label18.Caption = Label18.Caption + 1
Call PercentBar(Picture1, Label18.Caption, Label14.Caption)
'Above shit is a bit hard to do in detail but read
'what they say there are and that is what it
'does
If Label14.Caption = Label18.Caption Then
Timer4.Enabled = False
Else
GoTo Start:
End If
End Sub

Private Sub Timer5_Timer()
If Label18.Caption = Label14.Caption Then
IMBuddy "$IM_ON", "IMS ON"
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Mass mail Wuz ended"
Timeout 0.4
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• Say thanxz to " & UserSN
Timeout 0.6
Bolditalic_BlackPurpleBlack "•÷·· · ··÷• For mass mailin all u lamerz!"
Timeout 0.6
SendChat "•÷·· · ··÷• <U><Font color= #FF0000>Pe<Font color= #FF0450>ace O<Font color= #FF0800>ut</u> "
If Timer4.Enabled = False Then
Picture1.Visible = False
Timer5.Enabled = False
End If
End If
End Sub
