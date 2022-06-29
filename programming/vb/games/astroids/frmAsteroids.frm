VERSION 5.00
Begin VB.Form frmAsteroids 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "aStErOiDs vErSiOn 1.0"
   ClientHeight    =   10635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14985
   Icon            =   "FRMAST~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   10560
      Picture         =   "FRMAST~1.frx":08CA
      ScaleHeight     =   1440
      ScaleWidth      =   4305
      TabIndex        =   14
      Top             =   5160
      Width           =   4305
   End
   Begin VB.PictureBox picBoomM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "FRMAST~1.frx":14D0E
      ScaleHeight     =   750
      ScaleWidth      =   5250
      TabIndex        =   13
      Top             =   9720
      Width           =   5250
   End
   Begin VB.PictureBox picBoomS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "FRMAST~1.frx":21ACA
      ScaleHeight     =   750
      ScaleWidth      =   5250
      TabIndex        =   12
      Top             =   9000
      Width           =   5250
   End
   Begin VB.PictureBox picMissleS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   10800
      Picture         =   "FRMAST~1.frx":2E886
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   6360
      Width           =   75
   End
   Begin VB.PictureBox picMissleM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   10800
      Picture         =   "FRMAST~1.frx":2E91A
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   6240
      Width           =   75
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7560
      Top             =   240
   End
   Begin VB.PictureBox picShipM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   120
      Picture         =   "FRMAST~1.frx":2E9AE
      ScaleHeight     =   600
      ScaleWidth      =   9600
      TabIndex        =   9
      Top             =   8160
      Width           =   9600
   End
   Begin VB.PictureBox picShipS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   120
      Picture         =   "FRMAST~1.frx":415F2
      ScaleHeight     =   600
      ScaleWidth      =   9600
      TabIndex        =   8
      Top             =   7560
      Width           =   9600
   End
   Begin VB.PictureBox PicSmlRockM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1680
      Picture         =   "FRMAST~1.frx":54236
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.PictureBox PicSmlRockS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1680
      Picture         =   "FRMAST~1.frx":58BEA
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.PictureBox PicMedRockM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "FRMAST~1.frx":5D59E
      ScaleHeight     =   750
      ScaleWidth      =   7500
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox PicMedRockS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   240
      Picture         =   "FRMAST~1.frx":6FADA
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox picBackGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   0
      Picture         =   "FRMAST~1.frx":82016
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox picBigRockM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   -480
      Picture         =   "FRMAST~1.frx":11481A
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   11250
   End
   Begin VB.PictureBox PicBigRockS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   -480
      Picture         =   "FRMAST~1.frx":13DC22
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   750
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   11250
   End
   Begin VB.PictureBox PicBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5940
      Left            =   7560
      ScaleHeight     =   400
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.PictureBox picTitleM 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   3240
         Picture         =   "FRMAST~1.frx":16702A
         ScaleHeight     =   1440
         ScaleWidth      =   4305
         TabIndex        =   15
         Top             =   3360
         Width           =   4305
      End
   End
End
Attribute VB_Name = "frmAsteroids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Asteroids - Yet another rendition of a classic
' version 1.0
' (c)1999 John Wilson
'
' This one was a challenge, keeping up with a player, as many as
' 16 asteroids and 30 missles at the same time!
' but maybe you can follow the code. I tried breaking the program up into
' as many subroutines as possible so it would be easier to read.
' oh well......
' comments and suggestions?
' email:
' jwilson@carpet.dalton.peachnet.edu
' wilsonj@vol.com
' webmaster@zeroflake.com
'

' ****************************************************************
' function used in determining cpu speed
' ****************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long

' ****************************************************************
' BITBLT Function
' ***************************************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' constants used by the BITBLT function
Private Const SRCCOPY = &HCC0020            ' (DWORD) dest = source
Private Const SRCINVERT = &H660046          ' (DWORD) dest = source XOR dest
Private Const SRCAND = &H8800C6             ' (DWORD) dest = source AND dest

' ****************************************************************
' Sound Function
' ***************************************************************

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' ***************************************************************
' Ship Variables
' ***************************************************************

' Ships X and Y position
Private ShipX As Single
Private ShipY As Single

' Player/Ship Lives
Private Lives As Integer

' Ships thrust
Private ShipThrust As Single

' Ships direction
Private ShipDirection As Integer

' Movement values for ship
Private Xvector As Single
Private Yvector As Single

' ***************************************************************
' Missle Variables
' ***************************************************************

' missle enabled
Private missle(30) As Integer

' missle location
Private missleX(30) As Single
Private missleY(30) As Single

' missle directions
Private missleDx(30) As Single
Private missleDy(30) As Single

'life span of missles
Private missleLife(30) As Integer

'number of missles
Private MissleNumber As Integer
' ***************************************************************
' Asteroid Variables
' ***************************************************************

' Asteroid(s) enabled/type
' 0 = no asteroid
' 1 = big asteroid
' 2 = med. asteroid
' 3 = small asteroid
Private Asteroid(16) As Integer

' missle location
Private AsteroidX(16) As Single
Private AsteroidY(16) As Single

' missle directions
Private AsteroidDx(16) As Single
Private AsteroidDy(16) As Single

' ***************************************************************
' General Use Variables
' ***************************************************************

' delay
Private CPUDelay As Integer

' is game over?
Private GameOver As Boolean

' explosion
Private Boom As Boolean

' General direction tables
Private SinTable(15) As Single
Private CosTable(15) As Single

' animation variable
Private aniCount As Single

' Good 'ole level and score variables
Private Level As Integer
Private Score As Integer
Private HighScore As Integer

Private Sub Form_Load()

' set the height/width of the form

frmAsteroids.Height = 400 * Screen.TwipsPerPixelX
frmAsteroids.Width = 500 * Screen.TwipsPerPixelY

'set up some initial goodies, lives, default number of missles, etc.
Lives = 3
Level = 1
MissleNumber = Level + 3
GameOver = False

On Error Resume Next

HighScore = GetSetting("Asteroids", "HighScore", "One")

' Create data table for sin and cos functions
' used in movement of ship and missles
' we could do this as an equation in code, but i am soooooo lazy

' sin table is for Y-Axis movement
SinTable(0) = 0
SinTable(1) = -0.383
SinTable(2) = -0.707
SinTable(3) = -0.924
SinTable(4) = -1
SinTable(5) = -0.924
SinTable(6) = -0.707
SinTable(7) = -0.383
SinTable(8) = -0
SinTable(9) = 0.383
SinTable(10) = 0.707
SinTable(11) = 0.924
SinTable(12) = 1
SinTable(13) = 0.924
SinTable(14) = 0.707
SinTable(15) = 0.383

' cos table is for X-Axis movement
CosTable(0) = 1
CosTable(1) = 0.924
CosTable(2) = 0.707
CosTable(3) = 0.383
CosTable(4) = 0
CosTable(5) = -0.383
CosTable(6) = -0.707
CosTable(7) = -0.924
CosTable(8) = -1
CosTable(9) = -0.924
CosTable(10) = -0.707
CosTable(11) = -0.383
CosTable(12) = 0
CosTable(13) = 0.383
CosTable(14) = 0.707
CosTable(15) = 0.924

Call initShip
Call initAsteroids

End Sub
Private Sub PaintBack()

' make a copy of the background into the buffer
u% = BitBlt(PicBuffer.hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, picBackGround.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub PaintScreen()

' paint screen
u% = BitBlt(hDC, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub initShip()

' Put the ship coordinates in the center of screen
ShipX = 230
ShipY = 200

' Zero the ships movement
ShipThrust = 0
Xvector = 0
Yvector = 0

' Reset the ships direction
ShipDirection = 12

End Sub

Private Sub initAsteroids()

' get the random number generator pumping!
Randomize

' clear out any stray asteroids
For X = 1 To 16
    Asteroid(X) = 0
Next X

' turn on four large asteroids
' and set some default properties for them
Asteroid(1) = 1
Asteroid(2) = 1
Asteroid(3) = 1
Asteroid(4) = 1

Dim n As Integer

Randomize

n = Int(Rnd * 15)

AsteroidDx(1) = CosTable(n) * (Level / 3)
AsteroidDy(1) = SinTable(n) * (Level / 3)

Randomize

n = Int(Rnd * 15)

AsteroidDx(2) = CosTable(n) * (Level / 3)
AsteroidDy(2) = SinTable(n) * (Level / 3)

Randomize

n = Int(Rnd * 15)

AsteroidDx(3) = CosTable(n) * (Level / 3)
AsteroidDy(3) = SinTable(n) * (Level / 3)

Randomize

n = Int(Rnd * 15)

AsteroidDx(4) = CosTable(n) * (Level / 3)
AsteroidDy(4) = SinTable(n) * (Level / 3)

AsteroidX(1) = Int(Rnd * 100)
AsteroidX(2) = Int(Rnd * 100)
AsteroidX(3) = Int(Rnd * 100) + 400
AsteroidX(4) = Int(Rnd * 100) + 400

AsteroidY(1) = Int(Rnd * 400)
AsteroidY(2) = Int(Rnd * 400)
AsteroidY(3) = Int(Rnd * 400)
AsteroidY(4) = Int(Rnd * 400)

End Sub

Private Sub PaintAsteroids()

For X = 1 To 16

    If Asteroid(X) = 1 Then
    
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 75, 75, picBigRockM.hDC, Int(aniCount) * 75, 0, SRCAND)
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 75, 75, PicBigRockS.hDC, Int(aniCount) * 75, 0, SRCINVERT)
    
    End If
    
    If Asteroid(X) = 2 Then
    
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 50, 50, PicMedRockM.hDC, Int(aniCount) * 50, 0, SRCAND)
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 50, 50, PicMedRockS.hDC, Int(aniCount) * 50, 0, SRCINVERT)
    
    End If
    
    If Asteroid(X) = 3 Then
    
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 25, 25, PicSmlRockM.hDC, Int(aniCount) * 25, 0, SRCAND)
        u% = BitBlt(PicBuffer.hDC, Int(AsteroidX(X)), Int(AsteroidY(X)), 25, 25, PicSmlRockS.hDC, Int(aniCount) * 25, 0, SRCINVERT)
    
    End If

Next X

End Sub
Private Sub PaintShip()

'paint rocketship
u% = BitBlt(PicBuffer.hDC, Int(ShipX), Int(ShipY), 40, 40, picShipM.hDC, ShipDirection * 40, 0, SRCAND)
u% = BitBlt(PicBuffer.hDC, Int(ShipX), Int(ShipY), 40, 40, picShipS.hDC, ShipDirection * 40, 0, SRCINVERT)

End Sub
Private Sub PaintTitle()

'paint Title
u% = BitBlt(PicBuffer.hDC, 107, 50, 287, 96, picTitleM.hDC, 0, 0, SRCAND)
u% = BitBlt(PicBuffer.hDC, 107, 50, 287, 96, picTitle.hDC, 0, 0, SRCINVERT)

PicBuffer.CurrentY = 150

Dim data As String

PicBuffer.FontSize = 8

data = "(c) 1999 - John Wilson"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

PicBuffer.Print

PicBuffer.FontSize = 12

data = "Press <SPACE> to Start"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

PicBuffer.FontSize = 8

PicBuffer.Print

PicBuffer.FontSize = 10

data = "High Score:" & HighScore
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

PicBuffer.FontSize = 8

PicBuffer.Print

data = "Controls:"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "<Right Arrow> - Rotate Right"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "<Left Arrow> - Rotate Left"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "<Up Arrow> - Thrust Forward"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "<Down Arrow> - Flip"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "<Left Control> - Fire"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

data = "or, Mouse Control"
PicBuffer.CurrentX = (500 - PicBuffer.TextWidth(data)) / 2
PicBuffer.Print data

End Sub


Private Sub PaintMissles()

For X = 1 To MissleNumber

    If missle(X) = 1 Then
    
        'paint missles
        u% = BitBlt(PicBuffer.hDC, Int(missleX(X)), Int(missleX(X)), 5, 5, picMissleM.hDC, 0, 0, SRCAND)
        u% = BitBlt(PicBuffer.hDC, Int(missleX(X)), Int(missleY(X)), 5, 5, picMissleS.hDC, 0, 0, SRCINVERT)
    
    End If
    
Next X

End Sub
Private Sub PaintScore()

PicBuffer.Font = "Terminal"
PicBuffer.FontSize = 11
PicBuffer.ForeColor = vbGreen

PicBuffer.CurrentX = 10
PicBuffer.CurrentY = 360
PicBuffer.Print "Score:" & Score

temp = "Lives:" & Lives
PicBuffer.CurrentX = 280 - PicBuffer.TextWidth(temp)
PicBuffer.CurrentY = 360
PicBuffer.Print temp

temp = "Level:" & Level
PicBuffer.CurrentX = 485 - PicBuffer.TextWidth(temp)
PicBuffer.CurrentY = 360
PicBuffer.Print temp

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then Call PicBuffer_KeyDown(vbKeyControl, 0)
If Button = 2 Then Call PicBuffer_KeyDown(vbKeyUp, 0)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

' player movement w/mouse
' Pretty lame, maybe if i restrict mouse movement to the form, i dunno
' seems to awkward, interesting to write, surely there is an easier way....

On Error Resume Next

Opp = X - ShipX
Adj = Y - ShipY
Hyp = Sqr(Adj ^ 2 + Opp ^ 2)

angle = Atn((Opp / Hyp) / Sqr(-(Opp / Hyp) * (Opp / Hyp) + 1))

angle = Abs(angle * 57.2957795)

    If angle >= 0 And angle <= 11.25 Then
        If Opp > 0 And Adj < 0 Then
            ShipDirection = 4
        End If
        If Opp > 0 And Adj > 0 Then
            ShipDirection = 12
        End If
        
         If Opp < 0 And Adj < 0 Then
            ShipDirection = 4
        End If
        If Opp < 0 And Adj > 0 Then
            ShipDirection = 12
        End If
        
    ElseIf angle > 11.25 And angle <= 33.75 Then
        If Opp > 0 And Adj < 0 Then
            ShipDirection = 3
        End If
        If Opp > 0 And Adj > 0 Then
            ShipDirection = 13
        End If
      
        If Opp < 0 And Adj < 0 Then
            ShipDirection = 5
        End If
        If Opp < 0 And Adj > 0 Then
            ShipDirection = 11
        End If
        
    ElseIf angle > 33.75 And angle <= 56.25 Then
        If Opp > 0 And Adj < 0 Then
            ShipDirection = 2
        End If
        If Opp > 0 And Adj > 0 Then
            ShipDirection = 14
        End If
        If Opp < 0 And Adj < 0 Then
            ShipDirection = 6
        End If
        If Opp < 0 And Adj > 0 Then
            ShipDirection = 10
        End If
        
    ElseIf angle > 56.25 And angle <= 78.75 Then
        If Opp > 0 And Adj < 0 Then
            ShipDirection = 1
        End If
        If Opp > 0 And Adj > 0 Then
            ShipDirection = 15
        End If
        If Opp < 0 And Adj < 0 Then
            ShipDirection = 7
        End If
        If Opp < 0 And Adj > 0 Then
            ShipDirection = 9
        End If
        
    ElseIf angle > 78.75 And angle <= 90 Then
        If Opp > 0 And Adj < 0 Then
            ShipDirection = 0
        End If
        If Opp > 0 And Adj > 0 Then
            ShipDirection = 0
        End If
        If Opp < 0 And Adj < 0 Then
            ShipDirection = 8
        End If
        If Opp < 0 And Adj > 0 Then
            ShipDirection = 8
        End If
    End If

End Sub

Private Sub PicBuffer_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 32 And Timer1.Enabled = True Then
    Timer1.Enabled = False
    GameOver = False
    Boom = False
    Score = 0
    Level = 1
    Lives = 3
    Call GameLoop
    Call initAsteroids
    Call initShip
End If

' turn the ship left and right
If KeyCode = vbKeyRight Then
    ShipDirection = ShipDirection - 1
    ShipThrust = 0
    If ShipDirection < 0 Then ShipDirection = 15
ElseIf KeyCode = vbKeyLeft Then
    ShipDirection = ShipDirection + 1
    ShipThrust = 0
    If ShipDirection > 15 Then ShipDirection = 0
End If

' rocket thrusters
If KeyCode = vbKeyUp Then
    ShipThrust = 1
End If

' flip
If KeyCode = vbKeyDown Then
    ShipDirection = ShipDirection + 8
    If ShipDirection > 15 Then ShipDirection = ShipDirection - 16
End If

If GameOver = False Then

    If KeyCode = vbKeyControl Then
        X = 1
mloop:
        If missle(X) = 0 Then
            missle(X) = 1
            missleDx(X) = CosTable(ShipDirection)
            missleDy(X) = SinTable(ShipDirection)
            missleX(X) = ShipX + 18
            missleY(X) = ShipY + 18
            missleLife(X) = 0
            X = MissleNumber
        End If
        X = X + 1
        If X <= MissleNumber Then GoTo mloop
    End If
End If

End Sub

Private Sub Timer1_Timer()

Static runonce As Boolean

PaintBack

If runonce = False Then
    initAsteroids
End If

MoveRocks

aniCount = aniCount + 1
If aniCount > 9 Then aniCount = 0

PaintAsteroids
PaintTitle
PaintScore
PaintScreen

runonce = True

End Sub

Private Sub GameLoop()

Dim cancel

RunLoop:

aniCount = aniCount + 0.25
If aniCount >= 10 Then aniCount = 1

PaintBack
MoveMissles
MoveRocks
Do
Loop Until TimeDelay(20) = True

If GameOver = False Then
    MoveShip
    DShipCollision
End If

DMslCollision
PaintMissles
PaintAsteroids

If GameOver = False Then
    PaintShip
End If

If Boom = True Then
    DrawExplosion
End If

PaintScore
PaintScreen
DoEvents

If Timer1.Enabled = False Then

    GoTo RunLoop

End If

End Sub
Private Sub MoveShip()

If ShipThrust > 0 Then
    ShipThrust = ShipThrust - 0.05
End If

'from the direction and thrust work out the movement variables

Xvector = Xvector + (ShipThrust * CosTable(ShipDirection))

Yvector = Yvector + (ShipThrust * SinTable(ShipDirection))

If Xvector > 3 Then
    Xvector = 3
ElseIf Xvector < -3 Then
    Xvector = -3
ElseIf Xvector > 0 Then
    Xvector = Xvector - 0.05
ElseIf Xvector < 0 Then
    Xvector = Xvector + 0.05
End If

If Yvector > 3 Then
    Yvector = 3
ElseIf Yvector < -3 Then
    Yvector = -3
ElseIf Yvector > 0 Then
    Yvector = Yvector - 0.05
ElseIf Yvector < 0 Then
    Yvector = Yvector + 0.05
End If

ShipX = ShipX + Xvector
ShipY = ShipY + Yvector

' keep the ship on the screen
If ShipX > 475 Then ShipX = -15
If ShipX < -15 Then ShipX = 475
If ShipY > 375 Then ShipY = -15
If ShipY < -15 Then ShipY = 375

End Sub

Private Sub MoveMissles()

For X = 1 To MissleNumber

    If missle(X) = 1 Then
    
        missleX(X) = missleX(X) + (missleDx(X) * 5)
        missleY(X) = missleY(X) + (missleDy(X) * 5)

        If missleX(X) > 500 Then
            missleX(X) = 0
        ElseIf missleX(X) < 0 Then
            missleX(X) = 500
        End If
        
        If missleY(X) > 400 Then
            missleY(X) = 0
        ElseIf missleY(X) < 0 Then
            missleY(X) = 400
        End If
        
        missleLife(X) = missleLife(X) + 1
        
        If missleLife(X) > 50 Then
            missle(X) = 0
        End If
        
    End If

Next X

End Sub

Private Sub MoveRocks()

For X = 1 To 16
    If Asteroid(X) <> 0 Then
        AsteroidX(X) = AsteroidX(X) + AsteroidDx(X)
        AsteroidY(X) = AsteroidY(X) + AsteroidDy(X)
        
        If Asteroid(X) = 1 Then Size = 75
        If Asteroid(X) = 2 Then Size = 50
        If Asteroid(X) = 3 Then Size = 25
                
        If AsteroidX(X) < 0 - (Size / 4) Then AsteroidX(X) = 500 - (Size / 4)
        If AsteroidY(X) < 0 - (Size / 4) Then AsteroidY(X) = 400 - (Size / 4)
        If AsteroidX(X) > 500 - (Size / 4) Then AsteroidX(X) = 0 - (Size / 4)
        If AsteroidY(X) > 400 - (Size / 4) Then AsteroidY(X) = 0 - (Size / 4)
        
    End If
Next X

End Sub

Private Sub DShipCollision()

For X = 1 To 16

    If Asteroid(X) <> 0 Then
    
        If Asteroid(X) = 1 Then Size = 50
        If Asteroid(X) = 2 Then Size = 30
        If Asteroid(X) = 3 Then Size = 10
        
        If (ShipX >= AsteroidX(X) And ShipX <= AsteroidX(X) + Size) Or (AsteroidX(X) >= ShipX And AsteroidX(X) <= ShipX + 30) Then
    
                If (ShipY >= AsteroidY(X) And ShipY <= AsteroidY(X) + Size) Or (AsteroidY(X) >= ShipY And AsteroidY(X) <= ShipY + 30) Then
        
                    Call sndPlaySound(App.Path & "\shipboom.wav", 1)
                    
                    If Score > HighScore Then
                        HighScore = Score
                        SaveSetting "Asteroids", "HighScore", "One", HighScore
                    End If
                    
                    If Lives <= 0 Then
                        GameOver = True
                    End If
                    
                    Boom = True
                    Exit For
                                    
                End If

        End If
                    
    End If

Next X

End Sub

Private Sub DMslCollision()

' well here we go, missle to asteroid collision detection
' 5 missles, up to 16 asteroids .... this gets very complicated

For X = 1 To MissleNumber

    For n = 1 To 16
    
        If Asteroid(n) = 1 Then Size = 75
        If Asteroid(n) = 2 Then Size = 50
        If Asteroid(n) = 3 Then Size = 25
        
        If missle(X) = 1 And Asteroid(n) <> 0 Then
        
            If missleX(X) > AsteroidX(n) And missleX(X) < AsteroidX(n) + Size Then
                If missleY(X) > AsteroidY(n) And missleY(X) < AsteroidY(n) + Size Then
                    
                    Score = Score + 100 - Size
                    
                    ' bonus lives, you'll have to be good to get these
                    
                    If Score / 10000 = 1 Then Lives = Lives + 1
                    If Score / 20000 = 1 Then Lives = Lives + 1
                    If Score / 40000 = 1 Then Lives = Lives + 1
                    
                    Call sndPlaySound(App.Path & "\rockboom.wav", 1)
                    ' asteroid hit
                    ' make the rock smaller and kill the missle
                    missle(X) = 0
                    Asteroid(n) = Asteroid(n) + 1
                    If Asteroid(n) = 4 Then
                        Asteroid(n) = 0
                    Else
                        For z = 1 To 16
                            ' look for an empty asteroid slot and add a new one
                            If Asteroid(z) = 0 Then
                                Asteroid(z) = Asteroid(n)
                                AsteroidX(z) = AsteroidX(n) - (Size / 4)
                                AsteroidY(z) = AsteroidY(n) - (Size / 4)
                                AsteroidDx(z) = AsteroidDy(n)
                                AsteroidDy(z) = AsteroidDx(n)
                                Exit For
                            End If
                        Next z
                    End If
                End If
            End If
        End If
    Next n
Next X

'check for no more asteroids
n = 0

For X = 1 To 16
    
    n = n + Asteroid(X)

Next X

If n = 0 Then
    
    'next level
    For X = 1 To MissleNumber
        missle(X) = 0
    Next X
    Level = Level + 1
    initAsteroids
    initShip
    Xvector = 0
    Yvector = 0
    ShipThrust = 0
    MissleNumber = Level + 3
    If MissleNumber >= 30 Then MissleNumber = 30
End If

End Sub


Private Sub DrawExplosion()

Static nboom As Single

u% = BitBlt(PicBuffer.hDC, Int(ShipX), Int(ShipY), 50, 50, picBoomM.hDC, Int(nboom) * 50, 0, SRCAND)
u% = BitBlt(PicBuffer.hDC, Int(ShipX), Int(ShipY), 50, 50, picBoomS.hDC, Int(nboom) * 50, 0, SRCINVERT)

nboom = nboom + 0.5

If nboom >= 7 Then
    Boom = False
    Lives = Lives - 1
    initShip
    If Lives = 0 Then
        Timer1.Enabled = True
    End If
    nboom = 0
End If

End Sub

Private Function TimeDelay(ByVal Delay As Long) As Boolean


    Static Start As Long
    Dim Elapsed As Long

    If Start = 0 Then
        Start = GetTickCount
    End If

    Elapsed = GetTickCount

    If (Elapsed - Start) >= Delay Then
        TimeDelay = True
        Start = 0
    Else: TimeDelay = False
    End If

End Function
