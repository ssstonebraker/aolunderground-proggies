VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " AFK Bot"
   ClientHeight    =   510
   ClientLeft      =   1470
   ClientTop       =   1365
   ClientWidth     =   3330
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   510
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " Youve ben gone for"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   1815
      Begin VB.Label Label2 
         Caption         =   " Mins"
         DataSource      =   " "
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Height          =   195
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Stop"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1455
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Label1.Caption = 0
End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub


Private Sub Timer1_Timer()
DoEvents
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>AFK Bot</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I> I've ben gone for: ") + Label1 + (" Mins</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
Label1.Caption = Val(Label1) + 1
End Sub


Private Sub Timer2_Timer()

End Sub


