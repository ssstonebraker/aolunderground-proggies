VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form15"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
 Text1 = AOLLastChatLine
End Sub
