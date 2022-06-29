VERSION 5.00
Begin VB.Form Info 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "X-Treme Server InFo"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   HelpContextID   =   40
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   360
      X2              =   4200
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   360
      X2              =   4200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click On The Form To Exit ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      WhatsThisHelpID =   40
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " -•´)•–  X-Treme Server '99 &InFo -•´)•– "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   40
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      WhatsThisHelpID =   40
      Width           =   4335
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
StayOnTop Server
End Sub

Private Sub Form_Load()
CenterForm Me
StayOnTop Me
NewLine$ = Chr$(13) & Chr$(10)
Label1 = NewLine$ & NewLine$ & "Made With Visual Basic 5.0 Enterprise Edition."
Label1 = Label1 & NewLine$ & "I Will Like To ""Thanks!"" To Everyone Who Helped By Sending In bug Reports, Suggestions, Etc." & NewLine$ & "Shout outs go to:"
Label1 = Label1 & NewLine$ & "DapMaster(Cool Guy),Thanks Dude For Everything" & NewLine$ & "Mad21Maxx ( For The Free AOL Service.)"
Label1 = Label1 & NewLine$ & "CrackZone1(The Best ""Hacker""&&""Cracker"")"
Label1 = Label1 & NewLine$ & "Tito :" & "  The Bytes Hunter"
End Sub

Private Sub Form_LostFocus()
StayOnTop Me
End Sub

Private Sub Label1_Click()
Unload Me
StayOnTop Server
End Sub

Private Sub Label3_Click()
Unload Me
StayOnTop Server
End Sub
