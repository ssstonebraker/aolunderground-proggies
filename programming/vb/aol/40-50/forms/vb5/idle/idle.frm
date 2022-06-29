VERSION 5.00
Begin VB.Form Form8 
   Caption         =   " Idle Bot"
   ClientHeight    =   525
   ClientLeft      =   2055
   ClientTop       =   2130
   ClientWidth     =   1605
   LinkTopic       =   "Form8"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   525
   ScaleWidth      =   1605
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   840
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Stop"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Start"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>KnK Founders Idle Bot</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")

Timer1.Enabled = True
End Sub


Private Sub Command2_Click()
Timer1.Enabled = False
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + KnK_L() + "" & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>   I'm Back!!   </I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + KnK_R())
End Sub


Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Timer1_Timer()
DoEvents
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & L$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>KnK Founders Idle Bot</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
End Sub


