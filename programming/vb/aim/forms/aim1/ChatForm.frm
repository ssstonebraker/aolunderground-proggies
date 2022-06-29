VERSION 4.00
Begin VB.Form ChatForm 
   Caption         =   "ChatTest Form"
   ClientHeight    =   1125
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3405
   Height          =   1530
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3405
   Top             =   1170
   Width           =   3525
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "AddRoom"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Find SN"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Greeting"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Chat Scroll"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Chat Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Room ??"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cap Chnge"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ChatSend"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close Chat"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Chat"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "ChatForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call ChatClear
End Sub


Private Sub Command10_Click()
Call AddRoom(List1, True)
End Sub


Private Sub Command2_Click()
Call CloseChat
End Sub


Private Sub Command3_Click()
Call ChatSend("Hi")
End Sub


Private Sub Command4_Click()
Call CapChat("HI")
End Sub


Private Sub Command5_Click()
Call Greeting
End Sub


Private Sub Command6_Click()
Call SayRoom
End Sub

Private Sub Command7_Click()
MsgBox ChatName
End Sub

Private Sub Command8_Click()
Call ChatScroll(4, 0.4, "hi!")
End Sub

Private Sub Command9_Click()
If FindChatter("TimOstman") = True Then MsgBox ("Hes Here!")
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub


