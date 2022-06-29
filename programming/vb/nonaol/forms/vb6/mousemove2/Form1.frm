VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "KJL's Example"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1965
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   1965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Smoth RND"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblx2 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lbly2 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbly1 
      Caption         =   "0"
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblx1 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim z As POINTAPI 'Declare variable

Private Sub Command1_Click()
'picks a random place on your screen
'then it puts it there
Call SetCursorPos(RandomNumber(800), RandomNumber(600))
End Sub

Private Sub Command2_Click()
'enables and disables depending if
'it is running or not
If Command2.Caption = "Smoth RND" Then
    lblx2.Caption = RandomNumber(700) + 100
    lbly2.Caption = RandomNumber(500) + 100
    Timer2.Interval = 1
    Command2.Caption = "Smoth Stop"
Else
    Timer2.Interval = 0
    Command2.Caption = "Smoth RND"
End If
End Sub

Private Sub Timer1_Timer()
'this finds where your mouse is
'and puts it in a label
GetCursorPos z 'Get Co-ordinets
lblx1.Caption = z.x 'Get x co-ordinets
lbly1.Caption = z.Y 'Get y co-ordinets
End Sub

Private Sub Timer2_Timer()
'This finds a place on the screen
'then it slowly gose to that place
'1 pixel at a time
'then it checks if it is on that spot
'and then it finds a new spot
GetCursorPos z 'Get Co-ordinets
lblx1.Caption = z.x 'Get x co-ordinets
lbly1.Caption = z.Y 'Get y co-ordinets
If lbly1.Caption > lbly2.Caption And lblx1.Caption > lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption - 1, lbly1.Caption - 1)
ElseIf lbly1.Caption < lbly2.Caption And lblx1.Caption < lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption + 1, lbly1.Caption + 1)
ElseIf lbly1.Caption > lbly2.Caption And lblx1.Caption < lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption + 1, lbly1.Caption - 1)
ElseIf lbly1.Caption < lbly2.Caption And lblx1.Caption > lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption - 1, lbly1.Caption + 1)
ElseIf lbly1.Caption < lbly2.Caption Then
    Call SetCursorPos(lblx1.Caption, lbly1.Caption + 1)
ElseIf lbly1.Caption > lbly2.Caption Then
    Call SetCursorPos(lblx1.Caption, lbly1.Caption - 1)
ElseIf lblx1.Caption < lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption + 1, lbly1.Caption)
ElseIf lblx1.Caption > lblx2.Caption Then
    Call SetCursorPos(lblx1.Caption - 1, lbly1.Caption)
End If
If lblx1 = lblx2 And lbly1 = lbly2 Then
    lblx2.Caption = RandomNumber(700) + 100
    lbly2.Caption = RandomNumber(500) + 100
End If
End Sub
