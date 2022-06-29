VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Bounce 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BOOM !"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9000
   Icon            =   "bounce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton HelpBtn 
      Caption         =   "Help!"
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton ClearAll 
      Caption         =   "Clear Screen"
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin MSComctlLib.Slider NumBalls 
      Height          =   630
      Left            =   7440
      TabIndex        =   3
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   200
      SelStart        =   40
      TickFrequency   =   20
      Value           =   40
   End
   Begin VB.PictureBox Background 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7560
      Left            =   -120
      Picture         =   "bounce.frx":0882
      ScaleHeight     =   7500
      ScaleWidth      =   7500
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7560
   End
   Begin VB.Timer BallTimer 
      Interval        =   1
      Left            =   6240
      Top             =   7560
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3495
      Left            =   7680
      ScaleHeight     =   3435
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox BallPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   6960
      Picture         =   "bounce.frx":B7A74
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.Slider AbsorbPct 
      Height          =   630
      Left            =   7440
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   35
      TickFrequency   =   10
      Value           =   35
   End
   Begin MSComctlLib.Slider GravAccel 
      Height          =   630
      Left            =   7440
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Max             =   40
      SelStart        =   10
      TickFrequency   =   5
      Value           =   10
   End
   Begin MSComctlLib.Slider Spread 
      Height          =   630
      Left            =   7440
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin MSComctlLib.Slider Intensity 
      Height          =   630
      Left            =   7440
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   35
      TickFrequency   =   10
      Value           =   35
   End
   Begin MSComctlLib.Slider Speed 
      Height          =   630
      Left            =   7440
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1111
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   100
      SelStart        =   25
      TickFrequency   =   20
      Value           =   25
   End
   Begin VB.Label Label6 
      Caption         =   "Speed"
      Height          =   255
      Left            =   7560
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Boom 
      Caption         =   "BOOM !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   13
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Intensity"
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Spread"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Gravity"
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Particle Elasticity"
      Height          =   255
      Left            =   7560
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Particle Count"
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Bounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Boom! Particle Explosion Simulation by Jason Merlo  10/24/99
'jmerlo@austin.rr.com
'jason.merlo@frco.com
'http://home.austin.rr.com/smozzie

'I grabbed the map bitmap somewhere online.  At this point, I don't
'remember where I snagged it from, but the proper credit should go
'to the original artist.

Option Explicit
Option Base 1

'These are the screen dimensions.
Const WIDE = 500
Const TALL = 500

'The maximum number of balls that we can track.
Const BALLMAX = 500

'The balls array contains the vital statistics of each ball.
Dim Balls(BALLMAX) As Ball

Dim Xstart As Single
Dim Ystart As Single

Dim Absorb As Single
Dim Gravity As Single
Dim Xspread As Single
Dim Yspread As Single
Dim Xpos As Single
Dim Ypos As Single
Dim Yvel As Single
Dim Ytilt As Single
Dim Ytiltvel As Single
Dim Ystartloc As Single

Dim Yshadow As Long
Dim CurrentCount As Long

Dim i As Integer
Dim j As Integer
Dim u As Integer

Private Sub PrepareBuffer()

'This sub simply copies the background picture into the buffer.
'If there is any prep work to do to the background, this is where
'it would get done.

u% = BitBlt(PicBuffer.hDC, 0, 0, WIDE, TALL, Background.hDC, 0, 0, SRCCOPY)

End Sub

Private Function EraseBall(Dest As Long, Backg As Long, x As Long, y As Long, Oldshad As Long)

'This function erases a given ball and its shadow by pasting
'the appropriate section of background over it.
u% = BitBlt(Dest, x, y, 5, 5, Backg, x, y, SRCCOPY)
u% = BitBlt(Dest, x, Oldshad, 5, 3, Backg, x, Oldshad, SRCCOPY)

End Function

Private Function DrawBall(Source As Long, Dest As Long, Shape As Integer, x As Long, y As Long)

'Now we draw the new ball.
u% = BitBlt(Dest, x, y, 5, 5, Source, (Shape - 1) * 5, 5, SRCAND)
u% = BitBlt(Dest, x, y, 5, 5, Source, (Shape - 1) * 5, 0, SRCINVERT)

End Function

Private Function DrawShadow(Source As Long, Dest As Long, Shape As Integer, x As Long, Newshad As Long)

'Draw the shadow so that the ball will overlap it.
u% = BitBlt(Dest, x, Newshad, 5, 3, Source, (Shape - 1) * 5, 13, SRCAND)
u% = BitBlt(Dest, x, Newshad, 5, 3, Source, (Shape - 1) * 5, 10, SRCINVERT)

End Function

Private Sub BallTimer_Timer()

'Here's the meat of the program.  This is where we modify each
'ball's position based on the explosion's properties and global
'settings.

'Get the gravity and elasticity settings each time, since all
'balls are affected by it.  This makes it a little more interactive, too.
Absorb = AbsorbPct.Value / 100
Gravity = GravAccel.Value / 10
BallTimer.Interval = Speed.Value

For i = 1 To BALLMAX
  'Only operate on balls that are still bouncing.
  If Balls(i).Enabled = True Then
    'Get the class properties so we can work locally.  Now the values
    'in the class are the 'old' values, used for erasing balls
    'and for shadow calculations.
    Xpos = Balls(i).Xpos
    Ypos = Balls(i).Ypos
    Yvel = Balls(i).Yvel
    Ytilt = Balls(i).Ytilt
    Ytiltvel = Balls(i).Ytiltvel
    Ystartloc = Balls(i).Ystart
    'Apply gravity to the vertical velocity.
    Yvel = Yvel + Gravity
    'Adjust the tilt of the ball. This is used to skew the pattern for
    'a 3-d type of effect.  A z-axis modifier, if you will.
    Ytilt = Ytilt + Ytiltvel
    'Now check if we can kill this ball because it's out of bounds.
    If Xpos < 0 Or Xpos > WIDE Or (Balls(i).Yshadow) < 0 Or (Ypos + Ytilt + Ytiltvel) > TALL Then
      Balls(i).Enabled = False
      CurrentCount = CurrentCount - 1
    Else
      'Move ball.
      Xpos = Xpos + Balls(i).Xvel
      Ypos = Ypos + Yvel
      'If we went past our starting point, we need to rebound.
      If Ypos > Ystartloc Then
        'The ground absorbs some velocity and reverses the ball's direction.
        Yvel = Absorb * (-Yvel)
        Ypos = Ystartloc
        'If the ball has slowed down enough, we will stop it altogether.
        'The 0.5 factor here is arbitrary, you can experiment with different
        'values for different effects.  Note that if this value is too
        'low, the balls may *never* stop.
        If Abs(Yvel) < (0.5 * Gravity) Then
          'Take this ball out of service and free up a slot for a new one.
          Balls(i).Enabled = False
          CurrentCount = CurrentCount - 1
        End If
      End If
    End If
    'If this ball didn't die, we need to get its shadow position
    'and copy its stats over to the class properties.
    If Balls(i).Enabled = True Then
      'Calculate the shadow position and draw the ball.
      Yshadow = (Int(Balls(i).Ypos + Ytilt) + Int(Ystartloc - Balls(i).Ypos)) + 2
      'Update the values of all the parameters in the class.  These will
      'be the 'old' values on the next scan!
      Balls(i).Yshadow = Yshadow
      Balls(i).Xpos = Xpos
      Balls(i).Ypos = Ypos
      Balls(i).Yvel = Yvel
      Balls(i).Ytilt = Ytilt
      Balls(i).Ytiltvel = Ytiltvel
      Balls(i).Ystart = Ystartloc
    End If
  End If
Next i

'Update the video buffer.

'Step 1:  Copy the blank background over.
Call PrepareBuffer

'Step 2:  Draw all the shadows FIRST so that they don't overlap any balls.
For i = 1 To BALLMAX
  If Balls(i).Enabled = True Then
    Call DrawShadow(BallPic.hDC, PicBuffer.hDC, Balls(i).Shape, Int(Balls(i).Xpos), Balls(i).Yshadow)
  End If
Next i

'Step 3:  Draw all the balls.
For i = 1 To BALLMAX
  If Balls(i).Enabled = True Then
    Call DrawBall(BallPic.hDC, PicBuffer.hDC, Balls(i).Shape, Int(Balls(i).Xpos), Int(Balls(i).Ypos + Balls(i).Ytilt))
  End If
Next i

'Draw the screen once a scan.
Call Form_Paint

End Sub

Private Sub Boom_DblClick()

MsgBox ("Boom! was programmed by Jason Merlo in VB6.  Thanks for playing with it!")

End Sub

Private Sub ClearAll_Click()

'This sub cleans out the whole ball array, and erases all balls from
'the screen so you can start over.  This is here just in case the user
'sets the elasticity too high and the balls don't stop.   :)

For i = 1 To BALLMAX
  If Balls(i).Enabled = True Then
    Balls(i).Enabled = False
    Call EraseBall(PicBuffer.hDC, Background.hDC, Int(Balls(i).Xpos), Int(Balls(i).Ypos + Balls(i).Ytilt), Balls(i).Yshadow)
  End If
Next i

'Start the counter over.
CurrentCount = 0

'Clean the screen.
Call Form_Paint

End Sub

Private Sub Form_Load()

Randomize

'Set up the screen and all the controls according to our screen size.
PicBuffer.Height = TALL * Screen.TwipsPerPixelY
PicBuffer.Width = WIDE * Screen.TwipsPerPixelX
Bounce.Height = TALL * Screen.TwipsPerPixelY
Bounce.Width = (WIDE + 105) * Screen.TwipsPerPixelX

NumBalls.Left = (WIDE) * Screen.TwipsPerPixelX
AbsorbPct.Left = (WIDE) * Screen.TwipsPerPixelX
GravAccel.Left = (WIDE) * Screen.TwipsPerPixelX
Spread.Left = (WIDE) * Screen.TwipsPerPixelX
Intensity.Left = (WIDE) * Screen.TwipsPerPixelX
Speed.Left = (WIDE) * Screen.TwipsPerPixelX
Label1.Left = (WIDE + 8) * Screen.TwipsPerPixelX
Label2.Left = (WIDE + 8) * Screen.TwipsPerPixelX
Label3.Left = (WIDE + 8) * Screen.TwipsPerPixelX
Label4.Left = (WIDE + 8) * Screen.TwipsPerPixelX
Label5.Left = (WIDE + 8) * Screen.TwipsPerPixelX
Label6.Left = (WIDE + 8) * Screen.TwipsPerPixelX

NumBalls.Top = (63) * Screen.TwipsPerPixelY
AbsorbPct.Top = (123) * Screen.TwipsPerPixelY
GravAccel.Top = (183) * Screen.TwipsPerPixelY
Spread.Top = (243) * Screen.TwipsPerPixelY
Intensity.Top = (303) * Screen.TwipsPerPixelY
Speed.Top = (363) * Screen.TwipsPerPixelY
Label1.Top = (48) * Screen.TwipsPerPixelY
Label2.Top = (108) * Screen.TwipsPerPixelY
Label3.Top = (168) * Screen.TwipsPerPixelY
Label4.Top = (228) * Screen.TwipsPerPixelY
Label5.Top = (288) * Screen.TwipsPerPixelY
Label6.Top = (348) * Screen.TwipsPerPixelY

Boom.Top = 10 * Screen.TwipsPerPixelY
Boom.Left = (WIDE + 10) * Screen.TwipsPerPixelX

ClearAll.Left = (WIDE + 12) * Screen.TwipsPerPixelX
ClearAll.Top = (410) * Screen.TwipsPerPixelY
HelpBtn.Left = (WIDE + 12) * Screen.TwipsPerPixelX
HelpBtn.Top = (440) * Screen.TwipsPerPixelY

'We'll run every 1 millisecond, though it takes quite a bit longer than
'that to run each iteration of the timer code.
BallTimer.Interval = 22

'No balls currently onscreen.
CurrentCount = 0

'Create new instances of the ball class for each ball.
For i = 1 To BALLMAX
  Set Balls(i) = New Ball
  Balls(i).Enabled = False
Next i

'Brand new screen.
Call PrepareBuffer
Call Form_Paint

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'This sub starts an explosion.

'First, grab the mouse coordinates in pixels.
Xstart = x / Screen.TwipsPerPixelX
Ystart = y / Screen.TwipsPerPixelX

'Only start another explosion if we have enough available balls.
If CurrentCount + NumBalls.Value <= BALLMAX Then
  CurrentCount = CurrentCount + NumBalls.Value
  Call BuildBalls(NumBalls.Value)
End If
  
End Sub

Private Sub Form_Paint()

'This sub will keep the display freshened even if something gets
'in front of it.  We also call it from other functions.
'Basically, it just takes our buffered picture and copies it to
'the screen.

u% = BitBlt(hDC, 0, 0, WIDE, TALL, PicBuffer.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub BuildBalls(StartNum As Long)

'Grab the global settings.
Absorb = AbsorbPct.Value / 100
Gravity = GravAccel.Value / 10

'Grab the initial numbers for the spread calculation.
Xspread = Spread.Value / 5
Yspread = Spread.Value / 10

'Do this for each new ball.
For j = 1 To NumBalls.Value

  'Find an empty ball slot and fill it up with info.
  For i = 1 To BALLMAX
    If Balls(i).Enabled = False Then
      'Slot i is free.  Let's populate it.
      Set Balls(i) = New Ball
      Balls(i).Enabled = True
      Balls(i).Xpos = Xstart
      Balls(i).Ypos = Ystart
      Balls(i).Xvel = (Rnd(1) * Xspread) - (Xspread / 2)
      Balls(i).Yvel = -((Rnd(1) * Intensity.Value / 3) + (Intensity.Value / 20))
      Balls(i).Ystart = Ystart
      Balls(i).Ytilt = 0
      Balls(i).Ytiltvel = (Rnd(1) * Yspread) - (Yspread / 2)
      Balls(i).Yshadow = Ystart + 2
      Balls(i).Shape = Int(Rnd(1) * 4) + 1
      'Now fudge out of the loop and initialize the next ball.
      Exit For
    End If
  Next i

Next j

End Sub

Private Sub HelpBtn_Click()

Shell "Notepad.exe boom.txt", vbNormalFocus

End Sub
