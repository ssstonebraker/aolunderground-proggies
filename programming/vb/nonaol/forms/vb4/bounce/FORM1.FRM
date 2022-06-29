VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   1365
   ClientTop       =   765
   ClientWidth     =   3840
   Height          =   5790
   Left            =   1305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   3840
   Top             =   420
   Width           =   3960
   Begin VB.Timer tmrBounce 
      Interval        =   50
      Left            =   1320
      Top             =   2400
   End
   Begin VB.PictureBox picHidden 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   1800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   5385
      Left            =   0
      Picture         =   "Form1.frx":8301
      ScaleHeight     =   355
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   251
      TabIndex        =   0
      Top             =   0
      Width           =   3825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Const MERGEPAINT = &HBB0226
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Variables for positioning the ball.
Private Const BallR = 10            ' Radius.
Private Const BallD = 2 * BallR + 1 ' Diameter.
Private CurX As Single              ' Position.
Private CurY As Single
Private OldX As Single
Private OldY As Single
Private VelX As Single              ' Velocity.
Private VelY As Single
Private Xmax As Single              ' Edge of canvas.
Private Ymax As Single
' Draw the ball at (CurX, CurY).
Private Sub DrawBall()
    ' Fix the part of the image that was covered.
    BitBlt picCanvas.hDC, _
        OldX - BallR, OldY - BallR, BallD, BallD, _
        picHidden.hDC, OldX - BallR, OldY - BallR, SRCCOPY
    OldX = CurX
    OldY = CurY

    ' Redraw the ball.
    picCanvas.Circle (CurX, CurY), BallR

    ' Update the display.
    picCanvas.Refresh
End Sub

' Initialize values, and draw the initial ball.
Private Sub Form_Load()
    ' Make the form fit the picture.
    Width = (Width - ScaleWidth) + picCanvas.Width
    Height = (Height - ScaleHeight) + picCanvas.Height

    Xmax = picCanvas.ScaleWidth - BallR
    Ymax = picCanvas.ScaleHeight - BallR

    ' Set initial position and velocity.
    Randomize
    CurX = Int((Xmax - BallR + 1) * Rnd + BallR)
    CurY = Int((Ymax - BallR + 1) * Rnd + BallR)
    OldX = CurX
    OldY = CurY
    VelX = Int((10 - 5 + 1) * Rnd + 5)
    VelY = Int((10 - 5 + 1) * Rnd + 5)

    DrawBall
End Sub

' Move the ball.
Private Sub tmrBounce_Timer()
    CurX = CurX + VelX
    If (CurX > Xmax) Then
        CurX = Xmax
        VelX = -VelX
'        Beep
    ElseIf (CurX < BallR) Then
        CurX = BallR
        VelX = -VelX
'        Beep
    End If

    CurY = CurY + VelY
    If (CurY > Ymax) Then
        CurY = Ymax
        VelY = -VelY
'        Beep
    ElseIf (CurY < BallR) Then
        CurY = BallR
        VelY = -VelY
'        Beep
    End If

    DrawBall
End Sub

