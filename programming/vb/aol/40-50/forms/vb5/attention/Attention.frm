VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Attention"
   ClientHeight    =   630
   ClientLeft      =   2160
   ClientTop       =   2685
   ClientWidth     =   2475
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   630
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
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
      Left            =   720
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
      Text            =   "Text1"
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1 = True
End Sub


Private Sub Command2_Click()
Timer1 = False
End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
DoEvents
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>Attention</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
SendChat Text1
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>Attention</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")

End Sub


