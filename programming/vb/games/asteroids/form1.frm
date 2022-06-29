VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6015
   ClientLeft      =   1095
   ClientTop       =   1560
   ClientWidth     =   7110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   7110
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
'|This game was made by: GLeeT --  Copyright (C) 1996 - 1998              |
'|This may be a simple little game, but it helps you game programmers...  |
'|I Know this will help me in the near future!                            |
'|And I want to thank all of you who like all of my work!                 |
'|You may not get rid of any of this that I type out, because it is the   |
'|copyright, and I need all the leagle stuff in all of my stuff or else it|
'|is illeagle...  Hey, I didn't make the rules around here, the govt. did!|
'|                                Later,                                  |
'|                         GLeeT                                          |
'--------------------------------------------------------------------------

Dim cx As Double
Dim cy As Double
Dim OldCX As Double
Dim OldCY As Double
Dim X() As Double
Dim Y() As Double
Dim BulletsLoose As Integer
Dim da As Double
Dim db As Double
Dim Direction As Double
Dim Speed As Integer
Dim Pie As Double
Dim Forward As Boolean
Dim TurnLeft As Boolean
Dim TurnRight As Boolean
Dim Firing As Boolean
Dim ReallyFiring As Boolean
Dim RefreshWin As Boolean
Dim Radians As Double
Dim Shields As Boolean
Dim Projectile(100) As BulletData
Dim BaseStart As Double
Dim BaseEnd As Double
Private Type BulletData
    Active As Boolean
    DistanceMoved As Integer
    BulletX As Double
    BulletY As Double
    BulletDirection As Double
End Type
Private Declare Function GetActiveWindow Lib "User32" () As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'Halt the system for specified Time

Sub DrawShip()
Picture1.ForeColor = RGB(255, 255, 255)
X(1) = (cx + (Sin(Direction)) * 300)
Y(1) = (cy + (Cos(Direction)) * 300)
X(2) = (cx + (Sin(Direction + Degrees(140))) * 300)
Y(2) = (cy + (Cos(Direction + Degrees(140))) * 300)
X(3) = (cx + (Sin(Direction + Degrees(220))) * 300)
Y(3) = (cy + (Cos(Direction + Degrees(220))) * 300)
'Thrusters
X(4) = (cx + (Sin(Direction + Degrees(210))) * 320)
Y(4) = (cy + (Cos(Direction + Degrees(210))) * 320)
X(5) = (cx + (Sin(Direction + Degrees(150))) * 320)
Y(5) = (cy + (Cos(Direction + Degrees(150))) * 320)

X(6) = (cx + (Sin(Direction + Degrees(200))) * 330)
Y(6) = (cy + (Cos(Direction + Degrees(200))) * 330)
X(7) = (cx + (Sin(Direction + Degrees(160))) * 330)
Y(7) = (cy + (Cos(Direction + Degrees(160))) * 330)

X(8) = (cx + (Sin(Direction + Degrees(190))) * 340)
Y(8) = (cy + (Cos(Direction + Degrees(190))) * 340)
X(9) = (cx + (Sin(Direction + Degrees(170))) * 340)
Y(9) = (cy + (Cos(Direction + Degrees(170))) * 340)

X(10) = (cx + (Sin(Direction + Degrees(179))) * 350)
Y(10) = (cy + (Cos(Direction + Degrees(179))) * 350)
X(11) = (cx + (Sin(Direction + Degrees(181))) * 350)
Y(11) = (cy + (Cos(Direction + Degrees(181))) * 350)
X(12) = X(2)
Y(12) = Y(2)
X(13) = (cx + (Sin(Direction + Degrees(40))) * 230)
Y(13) = (cy + (Cos(Direction + Degrees(40))) * 230)
X(14) = (cx + (Sin(Direction + Degrees(120))) * 160)
Y(14) = (cy + (Cos(Direction + Degrees(120))) * 160)
X(15) = X(3)
Y(15) = Y(3)
X(16) = (cx + (Sin(Direction + Degrees(320))) * 230)
Y(16) = (cy + (Cos(Direction + Degrees(320))) * 230)
X(17) = (cx + (Sin(Direction + Degrees(240))) * 160)
Y(17) = (cy + (Cos(Direction + Degrees(240))) * 160)


'Sheilds

If Shields = True Then Picture1.Circle (cx, cy), 400

Picture1.Line (X(1), Y(1))-(X(14), Y(14))
Picture1.Line (X(17), Y(17))-(X(1), Y(1))

If 360 - UnDegrees(Direction) > 50 Then
    BaseStart = Direction + Degrees(50)
Else
    BaseStart = Direction + Degrees(50) - Degrees(360)
End If

If 360 - UnDegrees(Direction) > 130 Then
    BaseEnd = Direction + Degrees(130)
Else
    BaseEnd = Direction + Degrees(130) - Degrees(360)
End If

Picture1.Circle (cx, cy), 300, , BaseStart, BaseEnd

Picture1.Line (X(12), Y(12))-(X(13), Y(13))
Picture1.Line (X(13), Y(13))-(X(14), Y(14))
Picture1.Line (X(15), Y(15))-(X(16), Y(16))
Picture1.Line (X(16), Y(16))-(X(17), Y(17))

If Forward = True Then
    Picture1.Line (X(4), Y(4))-(X(5), Y(5))
    Picture1.Line (X(6), Y(6))-(X(7), Y(7))
    Picture1.Line (X(8), Y(8))-(X(9), Y(9))
    Picture1.Line (X(10), Y(10))-(X(11), Y(11))

End If

OldCX = cx
OldCY = cy
End Sub
Sub ClearShip()
Picture1.ForeColor = RGB(0, 0, 0)

Picture1.Line (X(1), Y(1))-(X(14), Y(14))
Picture1.Line (X(17), Y(17))-(X(1), Y(1))
Picture1.Line (X(4), Y(4))-(X(5), Y(5))
Picture1.Line (X(6), Y(6))-(X(7), Y(7))
Picture1.Line (X(8), Y(8))-(X(9), Y(9))
Picture1.Line (X(10), Y(10))-(X(11), Y(11))
Picture1.Line (X(12), Y(12))-(X(13), Y(13))
Picture1.Line (X(13), Y(13))-(X(14), Y(14))
Picture1.Line (X(15), Y(15))-(X(16), Y(16))
Picture1.Line (X(16), Y(16))-(X(17), Y(17))
Picture1.Circle (OldCX, OldCY), 300, , BaseStart, BaseEnd
Picture1.Circle (OldCX, OldCY), 400
End Sub

Private Sub Form_Load()
Startup = True
Pie = 3.14159265358979
Radians = (2 * Pie) / 360
Direction = Degrees(180)
Speed = 1
ReDim X(20)
ReDim Y(20)
ReallyFiring = True
End Sub

Private Sub Form_Resize()
Picture1.Width = Form1.Width
Picture1.Height = Form1.Height
cx = Picture1.Width / 2
cy = Picture1.Height / 2
Do Until 1 = 2
If GetActiveWindow <> Form1.hWnd Then
RefreshWin = True
End If
If RefreshWin = True Then
    If GetActiveWindow = Form1.hWnd Then
        RefreshWin = False
        DoEvents
        Picture1.Refresh
        Picture1.SetFocus
        DoEvents
    End If
End If
cx = cx + da
cy = cy + db
CheckWrap
ClearShip
DrawShip
If TurnLeft = True Then
Direction = Direction + Degrees(5)
End If
If TurnRight = True Then
Direction = Direction - Degrees(5)
End If

If Direction = Degrees(360) Then Direction = Degrees(0)
If Direction < Degrees(0) Then Direction = Degrees(355)

If Forward = True Then
da = da + ((Sin(Direction)) * Speed)
db = db + ((Cos(Direction)) * Speed)
End If
If Firing = True And ReallyFiring = True And Shields = False Then
Dim TempBullet As Integer
'First Bullet
TempBullet = FindFreeBullet
BulletsLoose = BulletsLoose + 1
Projectile(TempBullet).BulletX = X(13)
Projectile(TempBullet).BulletY = Y(13)
Projectile(TempBullet).Active = True
Projectile(TempBullet).BulletDirection = Direction
'Second Bullet
TempBullet = FindFreeBullet
BulletsLoose = BulletsLoose + 1
Projectile(TempBullet).BulletX = X(16)
Projectile(TempBullet).BulletY = Y(16)
Projectile(TempBullet).Active = True
Projectile(TempBullet).BulletDirection = Direction

ReallyFiring = False
End If
If BulletsLoose > 0 Then UpdateBullets
DoEvents
Sleep (10)
Loop
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
TurnLeft = True
ElseIf KeyCode = vbKeyRight Then
TurnRight = True
ElseIf KeyCode = vbKeyUp Then
Forward = True
ElseIf KeyCode = vbKeySpace Then
Firing = True
ElseIf KeyCode = vbKeyDown Then
Shields = True
ElseIf KeyCode = vbKeyEscape Then
End
End If
End Sub


Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
TurnLeft = False
ElseIf KeyCode = vbKeyRight Then
TurnRight = False
ElseIf KeyCode = vbKeyUp Then
Forward = False
ElseIf KeyCode = vbKeySpace Then
ReallyFiring = True
Firing = False
ElseIf KeyCode = vbKeyDown Then
Shields = False
End If
End Sub

Sub CheckWrap()
If cy > Picture1.Height Then
cy = 0
End If
If cy < 0 Then
cy = Picture1.Height
End If
If cx > Picture1.Width Then
cx = 0
End If
If cx < 0 Then
cx = Picture1.Width
End If
End Sub
Function Degrees(Number As Double) As Double
Degrees = (Number * Radians)
End Function
Function UnDegrees(Number As Double) As Double
UnDegrees = (Number / Radians)
End Function

Sub UpdateBullets()
Dim i As Integer
For i = 1 To UBound(Projectile)
If Projectile(i).Active = True Then
If Projectile(i).BulletX > Picture1.Width Or Projectile(i).BulletX < 0 Or Projectile(i).BulletY > Picture1.Height Or Projectile(i).BulletY < 0 And Projectile(i).Active = True Then
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (Projectile(i).BulletX, Projectile(i).BulletY)-(Projectile(i).BulletX + 5, Projectile(i).BulletY + 5), , BF
Projectile(i).Active = False
BulletsLoose = BulletsLoose - 1
Debug.Print BulletsLoose
End If
Picture1.ForeColor = RGB(0, 0, 0)
Picture1.Line (Projectile(i).BulletX, Projectile(i).BulletY)-(Projectile(i).BulletX + 5, Projectile(i).BulletY + 5), , BF
Projectile(i).BulletX = Projectile(i).BulletX + ((Sin(Projectile(i).BulletDirection)) * 200)
Projectile(i).BulletY = Projectile(i).BulletY + ((Cos(Projectile(i).BulletDirection)) * 200)
Picture1.ForeColor = RGB(255, 255, 255)
Picture1.Line (Projectile(i).BulletX, Projectile(i).BulletY)-(Projectile(i).BulletX + 5, Projectile(i).BulletY + 5), , BF
End If
Next i
End Sub
Function FindFreeBullet() As Integer
For i = 1 To UBound(Projectile)
If Projectile(i).Active = False Then
FindFreeBullet = i
Exit Function
End If
Next i
End Function

