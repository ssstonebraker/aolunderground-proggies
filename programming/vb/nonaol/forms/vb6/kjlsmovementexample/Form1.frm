VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "KJL's Example"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.Shape Ball 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape Top1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   0
      Width           =   4455
   End
   Begin VB.Shape Bottom1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Shape House 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2040
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape Right1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   4440
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Left1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was made by KJL (http://kjl.cjb.net)
'You can not hold me responable if you do anything to
'your computer while using this code
'use this coding at your own risk

'This is so it moves while you have your key down
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'_____________________________________________________
'this is to move when you press the up arrow key
If KeyCode = vbKeyUp Then
'this is were it stops it from hitting the house
For u = 0 To (Ball.Width + House.Width)
If Ball.Top + Ball.Height >= House.Top And Ball.Left + Ball.Width - u = House.Left And Ball.Top + Ball.Height <= (House.Top + House.Height + House.Height) Then GoTo Sides
Next u
'this stops it when you hit the border
If Ball.Top <= Top1.Top + Top1.Height Then GoTo Sides
'makes the ball move
Ball.Top = Ball.Top - 10
'if it hits something it skips and makes it so it does not move
Sides:
End If
'_____________________________________________________
'this is to move when you press the down arrow key
If KeyCode = vbKeyDown Then
'this is were it stops it from hitting the house
For d = 0 To (Ball.Width + House.Width)
If Ball.Top + Ball.Height >= House.Top And Ball.Left + Ball.Width - d = House.Left And Ball.Top + Ball.Height <= (House.Top + House.Height) Then GoTo Sides1
Next d
'this stops it when you hit the border
If Ball.Top >= Bottom1.Top - Ball.Height Then GoTo Sides1
'makes the ball move
Ball.Top = Ball.Top + 10
'if it hits something it skips and makes it so it does not move
Sides1:
End If
'_____________________________________________________
'this is to move when you press the left arrow key
If KeyCode = vbKeyLeft Then
'this is were it stops it from hitting the house
For l = 0 To (Ball.Height + House.Height)
If Ball.Left + Ball.Width >= House.Left And Ball.Top + Ball.Height - l = House.Top And Ball.Left + Ball.Width <= (House.Left + House.Width + House.Width) Then GoTo Sides2
Next l
'this stops it when you hit the border
If Ball.Left <= Left1.Left + Left1.Width Then GoTo Sides2
'makes the ball move
Ball.Left = Ball.Left - 10
'if it hits something it skips and makes it so it does not move
Sides2:
End If
'_____________________________________________________
'this is to move when you press the right arrow key
If KeyCode = vbKeyRight Then
'this is were it stops it from hitting the house
For r = 0 To (Ball.Height + House.Height)
If Ball.Left + Ball.Width >= House.Left And Ball.Top + Ball.Height - r = House.Top And Ball.Left + Ball.Width <= (House.Left + House.Width) Then GoTo Sides3
Next r
'this stops it when you hit the border
If Ball.Left >= Right1.Left - Ball.Width Then GoTo Sides3
'makes the ball move
Ball.Left = Ball.Left + 10
'if it hits something it skips and makes it so it does not move
Sides3:
End If
End Sub
