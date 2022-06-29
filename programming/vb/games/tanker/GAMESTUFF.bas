Attribute VB_Name = "GAMESTUFF"
Option Explicit

Public Type TankSettings
    Angle As Integer
    Power As Integer
End Type

Public Type OPair
    x As Double
    y As Double
End Type

Public Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const PNT = &HCC0020

Global Const TURN_RED = False
Global Const TURN_BLUE = True
Global Const LLength = 12

Global Const GREEN = 65280
Global Const SKY_BLUE = 16776960
Global Const DARK_GRAY = 8224125
Global Const WHITE = 16777215
Global Const RED = 255
Global Const BLUE = 16711680
Global Const ORANGE = 45055

Global Const OPENING = "PRESS F2 TO BEGIN A NEW GAME"

Global Const BORDER_INSET = 0
Global Const BORDER_RAISED = 1

Global Const RT_SCRIPT = "Red's Turn"
Global Const BT_SCRIPT = "Blue's Turn"
Global Const BulVel = 10
Global Const pi = 3.14159265358979

Global Const rDieCapt = "Red is destroyed!"
Global Const bDieCapt = "Blue is destroyed!"

Public XPos As Double, YPos As Double

Public InGame As Boolean
Public InFire As Boolean
Public InPause As Boolean

Public BTank As TankSettings, RTank As TankSettings

Global Const TxtFire = "FIRE!!!"
Global Const TxtReset = "RESET"

Global Const rSndFire = "REDGUN"
Global Const bSndFire = "BLUEGUN"
Global Const rDest = "REDHIT"
Global Const bDest = "BLUEHIT"
Global Const gDest = "GREENHIT"
Global Const Fire = "FIRE"
Global Const Er = "NEGATIVE"

Public XSpot As Double, YSpot As Double

Public tmpColor As Long

Public RTurn As Boolean
Public Mov As OPair
Public BulVis As Boolean

Public AppPath As String

Sub Make3D(pic As Form, ctl As Control, ByVal BorderStyle As Integer)
Dim AdjustX As Integer, AdjustY As Integer
Dim RightSide As Single
Dim BW As Integer, BorderWidth As Integer
Dim LeftTopColor As Long, RightBottomColor As Long
Dim i As Integer

    If Not ctl.Visible Then Exit Sub

    AdjustX = Screen.TwipsPerPixelX
    AdjustY = Screen.TwipsPerPixelY
    BorderWidth = 3
    Select Case BorderStyle
    Case 0:
        LeftTopColor = DARK_GRAY
        RightBottomColor = WHITE
    Case 1:
        LeftTopColor = WHITE
        RightBottomColor = DARK_GRAY
    End Select
    For BW = 1 To BorderWidth
        pic.CurrentX = ctl.Left - (AdjustX * BW)
        pic.CurrentY = ctl.Top - (AdjustY * BW)
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top - (AdjustY * BW)), LeftTopColor
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top - (AdjustY * BW)), LeftTopColor
    Next
End Sub

Sub MIDN()
Dim i%, j%, k&
    i% = Val(Format(Time, "NN"))
    j% = Val(Format(Time, "SS"))
    For k& = 1 To Abs(j% - i%)
        Randomize
    Next k&
End Sub

Public Function GetOPair() As OPair
Dim G As TankSettings
    If RTurn = True Then
        G = RTank
    Else
        G = BTank
    End If
    GetOPair.x = Sine(G.Angle) * G.Power
    GetOPair.y = Cosine(G.Angle) * G.Power
End Function

Public Function InverseSine(ByVal i As Double) As Double
    InverseSine = Atn(i / Sqr(-i * i + 1))
    InverseSine = InverseSine * (180 / pi)
End Function

Public Function Cosine(ByVal i As Double) As Double
    Cosine = Cos(i * (pi / 180))
End Function

Public Function Sine(ByVal i As Double) As Double
    Sine = Sin(i * (pi / 180))
End Function

Public Sub Wait(ByVal Seconds As Integer)
Dim Start As Double, TotalTime As Double, Finish As Double
    Start = Timer
    Do While Timer < Start + Seconds
        DoEvents
    Loop
    Finish = Timer
    TotalTime = Finish - Start
End Sub

Sub Main()
    AppPath = App.Path & "\"
    frmTank.Show vbModal
End Sub

Public Sub PlaySound(ByVal File As String)
Dim tFile As String
    tFile = AppPath & File & ".WAV"
    sndPlaySound tFile, SND_ASYNC
End Sub
