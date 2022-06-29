VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attention"
   ClientHeight    =   720
   ClientLeft      =   2160
   ClientTop       =   2685
   ClientWidth     =   2430
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "KnK-attention.frx":0000
   ScaleHeight     =   720
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1200
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
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
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Your text here"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If UserSN() = "" Then
MsgBox "Ummm..... For this program to work right,  I'd suguest that you sign on!  =)", vbInformation, "Need To sign on"
Exit Sub
End If

'AOL4.o command
aversion$ = GetFromINI("AOL", "aversion", App.Path + "\KnK.ini")
If aversion$ = "aol4" Then
SendChat BlackGreenBlack("«-×´¯`°  Attention  °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack(Text1)
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Attention  °´¯`×-»")
Timer1 = True
End If

'AOL95 command
If aversion$ = "aol95" Then
AOLChatSend ("«-×´¯`°  Attention  °´¯`×-»")
TimeOut (0.5)
AOLChatSend (Text1)
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Attention  °´¯`×-»")
Timer2 = True
End If
End Sub


Private Sub Command2_Click()
Timer1 = False
Timer2 = False
End Sub


Private Sub Form_Load()
Label1 = UserSN()
StayOnTop Me

End Sub

Private Sub Timer1_Timer()
DoEvents
SendChat BlackGreenBlack("«-×´¯`°  Attention  °´¯`×-»")
TimeOut (0.5)
SendChat BlackGreenBlack(Text1)
TimeOut (0.5)
SendChat BlackGreenBlack("«-×´¯`°  Attention  °´¯`×-»")

End Sub


Private Sub Timer2_Timer()

DoEvents
AOLChatSend ("«-×´¯`°  Attention  °´¯`×-»")
TimeOut (0.5)
AOLChatSend (Text1)
TimeOut (0.5)
AOLChatSend ("«-×´¯`°  Attention  °´¯`×-»")


End Sub
