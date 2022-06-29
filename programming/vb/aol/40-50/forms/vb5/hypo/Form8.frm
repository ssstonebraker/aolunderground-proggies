VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Room Freezer"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   900
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Text            =   $"Form8.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Text            =   $"Form8.frx":00F4
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Text            =   $"Form8.frx":01E8
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Text            =   $"Form8.frx":02DC
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Secret Freeze"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Freeze Room"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChat Text1 + Text2 + Text3 + Text4
  TimeOut 1
 SendChat Text1 + Text2 + Text3 + Text4
 TimeOut 1
 SendChat Text1 + Text2 + Text3 + Text4
 TimeOut 1
 SendChat Text1 + Text2 + Text3 + Text4
  TimeOut 1
 SendChat Text1 + Text2 + Text3 + Text4
  TimeOut 1
 SendChat Text1 + Text2 + Text3 + Text4
 Tea$ = "-æ-[ HyPO ™ ]-æ-ToaST's CraZy Lagger"
 fnt$ = "10"
A = Len(Tea$)
For w = 1 To A Step 4
    R$ = Mid$(Tea$, w, 1)
    u$ = Mid$(Tea$, w + 1, 1)
    S$ = Mid$(Tea$, w + 2, 1)
    T$ = Mid$(Tea$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup><FONT SIZE=" + fnt$ + "><b>" & R$ & "</sup></font></b>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRBb = P$
 TimeOut 1
SendChat WavYChaTRBb
End Sub

Private Sub Command2_Click()

 
Unload Form8

Form8.Hide

End Sub

Private Sub Command3_Click()
SendChat "<html>Dont Lag</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><font Color#FF0000><html></html><html></html><html></html><html></html><html></html>"
SendChat "<html>Please Stop<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
SendChat "<html>HEY Stop Laggin</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><font Color#FF0000><html></html><html></html><html></html><html></html><html></html>"
SendChat "<html> I hate Laggers</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><font Color#FF0000><html></html><html></html><html></html><html></html><html></html>"
 
End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form8

End Sub

