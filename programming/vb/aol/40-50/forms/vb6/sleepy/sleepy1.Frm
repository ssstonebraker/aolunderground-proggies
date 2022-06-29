VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node's Sleepy Time"
   ClientHeight    =   1575
   ClientLeft      =   1245
   ClientTop       =   1470
   ClientWidth     =   2505
   Icon            =   "sleepy1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1575
   ScaleWidth      =   2505
   Begin VB.CommandButton Command5 
      Caption         =   "IMz Off"
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "IMz On"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   1920
      Top             =   840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Reason To Idle"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
dos32.ChatSend ("--=(I'm In Idle Cause)=--")
dos32.ChatSend ("--=(" + Text1.Text + ")=--")
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub


Private Sub Command3_Click()
dos32.ChatSend ("--=(Node's Sleepy Time)=--")
Pause 0.51
dos32.ChatSend ("--=(By: Node)=--")
Pause 0.51
dos32.ChatSend ("--=(Unloaded)=--")
Form1.Hide
End
End Sub


Private Sub Command4_Click()
dos32.IMsOn
End Sub

Private Sub Command5_Click()
dos32.IMsOff
End Sub

Private Sub Form_Load()
dos32.ChatSend ("--=(Node's Sleepy Time)=--")
Pause 0.49
dos32.ChatSend ("--=(By: Node)=--")
Pause 0.49
dos32.ChatSend ("--=(Loaded)=--")
End Sub

Private Sub Form_Terminate()
dos32.ChatSend ("--=(Node's Sleepy Time)=--")
Pause 0.51
dos32.ChatSend ("--=(By: Node)=--")
Pause 0.51
dos32.ChatSend ("--=(Unloaded)=--")
End Sub


Private Sub Form_Unload(Cancel As Integer)
dos32.ChatSend ("--=(Node's Sleepy Time)=--")
Pause 0.51
dos32.ChatSend ("--=(By: Node)=--")
Pause 0.51
dos32.ChatSend ("--=(Unloaded)=--")
End Sub


Private Sub Timer1_Timer()
dos32.ChatSend ("--=(I'm In Idle Cause)=--")
dos32.ChatSend ("--=(" + Text1.Text + ")=--")
DoEvents
End Sub


