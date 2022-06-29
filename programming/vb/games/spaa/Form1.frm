VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8580
      TabIndex        =   0
      Top             =   5475
      Width           =   8640
      Begin VB.Label Label6 
         Height          =   255
         Left            =   8400
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "By Paul Bryan, in 1999"
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Score"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Health"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   120
   End
   Begin VB.Image Image2 
      Height          =   960
      Index           =   0
      Left            =   5040
      Picture         =   "Form1.frx":08CA
      Top             =   1920
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   15
      Left            =   3240
      Picture         =   "Form1.frx":1F14
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   14
      Left            =   2760
      Picture         =   "Form1.frx":27DE
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   13
      Left            =   3600
      Picture         =   "Form1.frx":30A8
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   7320
      Picture         =   "Form1.frx":3972
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   11
      Left            =   3360
      Picture         =   "Form1.frx":423C
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   10
      Left            =   2400
      Picture         =   "Form1.frx":4B06
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   9
      Left            =   3000
      Picture         =   "Form1.frx":53D0
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   8
      Left            =   8040
      Picture         =   "Form1.frx":5C9A
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   7
      Left            =   7320
      Picture         =   "Form1.frx":6564
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   6
      Left            =   7320
      Picture         =   "Form1.frx":6E2E
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   5
      Left            =   2280
      Picture         =   "Form1.frx":76F8
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   4
      Left            =   1680
      Picture         =   "Form1.frx":7FC2
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   3
      Left            =   8040
      Picture         =   "Form1.frx":888C
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   2
      Left            =   3720
      Picture         =   "Form1.frx":9156
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   1
      Left            =   8040
      Picture         =   "Form1.frx":9A20
      Top             =   4200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":A2EA
      Top             =   4800
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'"SouthPark Alien Abduction" by Paul Bryan 1999
'                   email: pb2012@mad.scientist.com
'
'
' you are Carl the Visiting Alien, Floating around a sizable Playing Field.
' click next to SP people to abduct them, be sure to stay away from the succubus,
' cause he drains your health.The object is to abduct everyone, without being killed by the Succubus.
'
'                           Good Luck!
Dim r(15) As Integer 'Robot regeneration
Dim X(15) As Integer 'Robot X Position
Dim Y(15) As Integer 'Robot Y Position
Dim i(15) As Integer 'Robot Move Interval (Animation Speed) {Higher number faster movement)
Public s As Integer ' Total Score Constant
Public lv As Integer ' Level Constant
Public ls As Integer ' Level Score Constant
Dim h As Integer ' Health Constant
Private Sub Form_Load()
    Randomize Timer() ' Initialize Random # Generator
    lv = lv + 1: h = h + 100 ' Start Level 1
    ls = 0: Image1(0).Left = Int(Rnd(1) * (Me.Width - 500)): Image1(0).Top = Int(Rnd(1) * (Me.Height - 1600))
    For n = 0 To 15: r(n) = Int(Rnd(1) * lv) + 1: X(n) = Int(Rnd(1) * (Form1.Width - 500)) + 1
                'Object Positions and Defaults for Level, lv
    Y(n) = Int(Rnd(1) * (Form1.Height - 1600)) + 1: i(n) = Int(Rnd(1) * 70) + 1: Next
    d = Int(Rnd(1) * 10) + 1: j = Int(Rnd(1) * 5) + 1
    rb = d: For n = rb To rb + j: Image2(n).Visible = True: Next n
    Image2(0).Visible = True: Label3.Caption = h: Me.Caption = "SPAA Level " & lv

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        'Move Carl to where the form was clicked
    Call umove(X, Y)
End Sub

Public Sub umove(X As Single, Y As Single)
p = 10 ' carls Speed (interval)
u = (3 * lv)
lp:
    t = Image1(0).Top: l = Image1(0).Left 'get carls position
    If t = Y Then GoTo pt
    If t <= Y Then Image1(0).Top = Image1(0).Top + p: GoTo pt
    If t >= Y Then Image1(0).Top = Image1(0).Top - p
            'which way to move, up or down of current, and move.
pt:
    If (l <= X + p And l >= X - p) And (t <= Y + p And t >= Y - p) Then GoTo dn
            'see if carl is within interval of destination. if yes then dn
    If l = X Then GoTo py
    If l <= X Then Image1(0).Left = Image1(0).Left + p: GoTo py
    If l >= X Then Image1(0).Left = Image1(0).Left - p
            'which way to move, left or right of current, and move.
py:
    GoTo lp
dn: ' Landed
    For n = 0 To 15: If r(n) = 0 Or Image2(n).Visible = False Then GoTo nj
    X1 = Image2(n).Top: Y1 = Image2(n).Left
    If t >= X1 - 600 And t <= X1 + 600 Then GoTo ku
            'check for Objects within range
    GoTo nj
ku:
    If l >= Y1 - 600 And l <= Y1 + 600 Then GoTo scor
             'if object within range, then either score or get hurt
    GoTo nj
scor:
    If n = 0 Then s = s - 1: h = h - (lv * 2): Label3.Caption = h: GoTo nj
             'if object = 0 then get hurt
    s = s + u: ls = ls + u: r(n) = r(n) - 1: Image2(n).Visible = False
             'otherwise Score
    Image2(n).Top = Int(Rnd(1) * Me.Height): Image2(n).Left = Int(Rnd(1) * Me.Width)
    Label4.Caption = s
nj:
    Next ' check next object for it's position range, from carl.
    If h <= 0 Then h = 0: Label3.Caption = h: GoTo ulose
        'if health < 1 then U Lose
    GoTo er
ulose:
    MsgBox "Sorry but you Lost with " & s & " Points, on Level " & lv
        ' health > 1
    s = 0: lv = 0: h = 0
    Call Form_Load
  
er: ' destination reached, Scored or Hurt, or both has been complete.
End Sub
Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
X = Image2(Index).Top - 5: Y = Image2(Index).Left + 5
    'find the location of the robot clicked
Call umove(Y, X) ' send Carl to be moved there, and check for a score or a hit.
End Sub


Private Sub Label6_Click()
MsgBox "SouthPark Alien Abduct Game, by Paul Bryan 1999, email:pb2012@mad.scientist.com"
End Sub

Private Sub Timer1_Timer()
    'Move Robots
    For n = 0 To 15: If r(n) = 0 Or Image2(n).Visible = False Then GoTo dn
lp:
    t = Image2(n).Top: l = Image2(n).Left ' Current robot postion
    If t = Y(n) Then GoTo pt
    If t < Y(n) Then Image2(n).Top = Image2(n).Top + i(n): GoTo pt
    If t > Y(n) Then Image2(n).Top = Image2(n).Top - i(n)
                'which way to move,up or down of current robot position.
pt:
    If l = X(n) And t = Y(n) Then GoTo dn
    If l = X(n) Then GoTo py
    If l < X(n) Then Image2(n).Left = Image2(n).Left + i(n): GoTo py
    If l > X(n) Then Image2(n).Left = Image2(n).Left - i(n)
                'which way to move, left or right of current robot position.
py:
    If l = X(n) And t = Y(n) Then GoTo dn ' if robot to destination then skip to next robot
dn:
    If n = 0 Then GoTo pb
        'if n is the succubus then goto homing routine
    GoTo js
pb: ' homing routine
    Y(n) = Image1(0).Top - 100: X(n) = Image1(0).Left - 100
        'find carls position, and set the succubus Destination to it.
    i(n) = Int(ls) ' set carls speed (interval)
       
    If i(n) < (lv * 10) Then i(n) = (lv * 10) ' set min level speed of succubus
    If i(n) > 140 Then i(n) = 140 ' set max speed of succubus
    If Image1(0).Top >= t And Image1(0).Top <= t + 800 Then GoTo ku
        'if succubus within  top range of carl then continue, otherwise next robot
    GoTo js
ku:
    If Image1(0).Left >= l And Image1(0).Left <= l + 800 Then GoTo hit
        'if succubus within left range of carl then Hurt Carl
    GoTo js
hit:
    h = h - (lv * 2): Label3.Caption = h: GoTo js
        'succubus attack carl
js:
    Next n ' next robot move
End Sub
Public Sub newloc()
' Determine a new destination for a robot
be:
    k = 0 ' reset screen robot counter
    For n = 1 To 15: If r(n) <= 0 Or Image2(n).Visible = False Then GoTo gu
    k = k + 1 ' how many robots on the screen now
    lo = Int(Rnd(1) * 100) + 1: If lo < 50 Then GoTo gu
        '50/50 chance of actually making a change to the destination of each robot
        'per timer cycle.
    X(n) = Int(Rnd(1) * (Form1.Width - 500)) + 1: Y(n) = Int(Rnd(1) * (Form1.Height - 1600)) + 1
        'set new robot destination
    i(n) = Int(Rnd(1) * 70) + 1: If i(n) < 10 Then i(n) = 10
        'set new robot speed (interval)
gu:
    Next ' next robot 50/50 chance to change

    kp = 0 ' reset avail robot counter
    For n = 1 To 15: If r(n) <= 0 Then GoTo hg
    If r(n) > 0 Then kp = kp + 1 ' if robot still has more than 0 lives
                                 ' Then Count it as availible.
    If r(n) > 0 And k < 6 Then Image2(n).Visible = True: k = k + 1
            'if current robot life>0 and # of bots on screen<6 then
            'make the current robot visible to carl.
hg:
    Next ' next robot
    If kp = 0 And k <= 1 Then GoTo uwin
            ' no robots left availible or on the screen, Next Level.
    GoTo hy
uwin: ' go onto next level
    Beep
    Call Form_Load ' goes to next level
hy:
    End Sub
Private Sub Timer2_Timer()
        'select new destinations for robots
    newloc
End Sub
