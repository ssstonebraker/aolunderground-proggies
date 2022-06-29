VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Stop Watch"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stop Watch Example made by SlipKnot (erik)
'Email: EriK@hockeyfights.net  or  Slipknot8276@yahoo.com
'www.8op.com/slipknotpro
'This is a simple stop watch that i made
'It uses a timer, a label, and 2 command buttons
'This is to understand timers better!!
'Only uses 3 lines of code and NO .bas file (module)
'this is for beginning programmers, and easy to
'understand!!


Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Label1.Caption + 1
End Sub
