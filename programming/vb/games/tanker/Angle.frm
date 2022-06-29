VERSION 5.00
Begin VB.Form frmTank 
   BorderStyle     =   0  'None
   ClientHeight    =   8655
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   11880
   ControlBox      =   0   'False
   DrawWidth       =   5
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Coordinate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   945
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Timer tmTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   7200
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "FIRE!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MaskColor       =   &H000000FF&
      TabIndex        =   12
      Top             =   7680
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.HScrollBar scrAngle 
      Height          =   255
      LargeChange     =   10
      Left            =   360
      Max             =   90
      Min             =   -90
      TabIndex        =   2
      Top             =   8040
      Width           =   1575
   End
   Begin VB.HScrollBar scrPower 
      Height          =   255
      LargeChange     =   100
      Left            =   2160
      Max             =   1000
      TabIndex        =   1
      Top             =   8040
      Width           =   1575
   End
   Begin VB.PictureBox pctField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FF00&
      Height          =   6975
      Left            =   120
      MouseIcon       =   "Angle.frx":0000
      ScaleHeight     =   465
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   783
      TabIndex        =   0
      Tag             =   "/3D/"
      Top             =   240
      Width           =   11745
      Begin VB.PictureBox BDead 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   840
         Picture         =   "Angle.frx":0152
         ScaleHeight     =   90
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   5880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox RDead 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   480
         Picture         =   "Angle.frx":02B4
         ScaleHeight     =   90
         ScaleWidth      =   240
         TabIndex        =   17
         Top             =   5880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox BDie 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   840
         Picture         =   "Angle.frx":0416
         ScaleHeight     =   180
         ScaleWidth      =   240
         TabIndex        =   16
         Top             =   6000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox RDie 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   480
         Picture         =   "Angle.frx":0698
         ScaleHeight     =   180
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   6000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox TankBlue 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   840
         Picture         =   "Angle.frx":091A
         ScaleHeight     =   90
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   5760
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox TankRed 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   480
         Picture         =   "Angle.frx":0A7C
         ScaleHeight     =   90
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   5760
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblPause 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PAUSED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   3240
         Visible         =   0   'False
         Width           =   11745
      End
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   9240
      TabIndex        =   11
      Top             =   7770
      Width           =   2175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRedScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
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
      Height          =   420
      Left            =   8160
      TabIndex        =   10
      Top             =   7770
      Width           =   855
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RED:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7320
      TabIndex        =   9
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label lblBlueScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0000"
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
      Height          =   420
      Left            =   6240
      TabIndex        =   8
      Top             =   7770
      Width           =   855
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BLUE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   5280
      TabIndex        =   7
      Top             =   7800
      Width           =   885
   End
   Begin VB.Image imgScores 
      Height          =   615
      Left            =   5160
      Tag             =   "/3D/"
      Top             =   7680
      Width           =   6450
   End
   Begin VB.Image imgOutset 
      Height          =   855
      Left            =   240
      Tag             =   "/3DUP/"
      Top             =   7560
      Width           =   11505
   End
   Begin VB.Image imgInset 
      Height          =   1095
      Left            =   120
      Tag             =   "/3D/"
      Top             =   7440
      Width           =   11745
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      Caption         =   "Angle:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   7680
      Width           =   450
   End
   Begin VB.Label lblAng 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label lblPower 
      AutoSize        =   -1  'True
      Caption         =   "Power:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label lblPow 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1000"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   7680
      Width           =   855
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGamePause 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameEnd 
         Caption         =   "&End Game"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuGamePause1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameCredits 
         Caption         =   "&Credits"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuGameSettings 
         Caption         =   "&Settings..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuGamePause2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BLength As Double, RLength As Double
Dim BulX As Double, BulY As Double
Dim TankTurn As Boolean
Dim FreeBool As Boolean
Dim KDwn As Boolean

Sub DrawTank(ByVal Angle As Integer, ByVal Length As Integer)
Dim TmpX, TmpY, x, y, i
    x = Sine(Angle) * Length
    y = Cosine(Angle) * Length
    pctField.Line (XPos, YPos)-(XPos + x, YPos - y)
    Angle = Angle: Length = Length
End Sub

Private Sub cmdFire_Click()
Dim FireSndFile As String
Dim i As Integer
    If cmdFire.Caption = TxtReset Then
        SetupGame
        SetupTurn TankTurn
        cmdFire.Caption = TxtFire
        Exit Sub
    End If
    If InGame = False Then PlaySound Er: Exit Sub
    If InPause = True Then PlaySound Er: Exit Sub
    If InFire Then PlaySound Er: Exit Sub
    pctField.MousePointer = 99
    FreeBool = True
    i = Int(Rnd * 10): Debug.Print i
    If i = 7 Then PlaySound Fire
    Inscript "FIRE!!!"
    pctField.ScaleMode = 1
    If RTurn = True Then
        FireSndFile = rSndFire
        BulX = TankRed.Left + TankRed.Width / 2
        BulY = TankRed.Top - 5
    Else
        FireSndFile = bSndFire
        BulX = TankBlue.Left + TankBlue.Width / 2
        BulY = TankBlue.Top - 5
    End If
    PlaySound FireSndFile
    pctField.ScaleMode = 3
    BulVis = True
    Mov = GetOPair
    tmTimer.Enabled = True
    InFire = True
End Sub

Private Sub Form_Load()
    lblStat = OPENING
End Sub

Private Sub Form_Paint()
Dim i As Integer
    ScaleMode = 1
    DrawWidth = 1
    On Error Resume Next
    For i = Me.Controls.Count - 1 To 0 Step -1
        If InStr(UCase$(Me.Controls(i).Tag), "/3D/") Then
            Make3D Controls(i).Container, Me.Controls(i), BORDER_INSET
        ElseIf InStr(UCase$(Me.Controls(i).Tag), "/3DUP/") Then
            Make3D Controls(i).Container, Me.Controls(i), BORDER_RAISED
        End If
    Next
    DrawWidth = 5
    ScaleMode = 3
End Sub

Sub SetupGame()
Dim X1 As Double, Y1 As Double
Dim X2 As Double, Y2 As Double
Dim Obj As Control, Founded As Boolean
    InGame = True
    Inscript RT_SCRIPT
    RTurn = True
    InPause = False
    cmdFire.Caption = TxtFire
    MIDN
    pctField.DrawWidth = 2
    X1 = Rnd * pctField.Width \ 2 + 5
    X2 = pctField.Width - Rnd * pctField.Width \ 2 - 5
    Y1 = pctField.Height - (Rnd * pctField.Height \ 2) - 30
    Y2 = pctField.Height - (Rnd * pctField.Height \ 2) - 30
    TankRed.Move X1, Y1
    TankBlue.Move X2, Y2
    For X2 = 1 To 1000
        DoEvents
    Next X2
    X1 = 0
    Set Obj = TankRed
    Set pctField.Picture = Nothing
    pctField.BackColor = SKY_BLUE
    pctField.ForeColor = GREEN
    pctField.AutoRedraw = True
    Do
        X2 = X1: Y2 = Y1
        X1 = X1 + (Rnd * 8) + 15
        Y1 = pctField.Height - Rnd * (pctField.Height / 2) - 20
        DrawWidth = 1
        If X1 > Obj.Left Then
            If Obj.Name = "TankRed" Then
                Set Obj = TankBlue
                pctField.Line (X2, Y2)-(TankRed.Left, TankRed.Top + TankRed.Height)
                X2 = TankRed.Left: Y2 = TankRed.Top + TankRed.Height
                X1 = TankRed.Left + TankRed.Width
                Y1 = TankRed.Top + TankRed.Height
            ElseIf Founded = False Then
                Founded = True
                pctField.Line (X2, Y2)-(TankBlue.Left, TankBlue.Top + TankBlue.Height)
                X2 = TankBlue.Left: Y2 = TankBlue.Top + TankBlue.Height
                X1 = TankBlue.Left + TankBlue.Width
                Y1 = TankBlue.Top + TankBlue.Height
            End If
        End If
        pctField.Line (X1, Y1)-(X2, Y2)
    Loop Until X1 > pctField.Width
    DrawWidth = 2
    pctField.ForeColor = SKY_BLUE
    FloodFill pctField.hdc, 1, pctField.Height - 1, RGB(0, 255, 0)
    pctField.Line (TankRed.Left - 5, TankRed.Top - 300)-(TankRed.Left + TankRed.Width + 5, TankRed.Top + TankRed.Height + 1), , BF
    pctField.Line (TankBlue.Left - 5, TankBlue.Top - 300)-(TankBlue.Left + TankBlue.Width + 5, TankBlue.Top + TankBlue.Height + 1), , BF
    pctField.Picture = pctField.Image
    scrAngle = 30: scrPower = 300
    RTank.Angle = 30: RTank.Power = 300
    BTank.Angle = -30: BTank.Power = 300
    TankFresh
    ScaleMode = 3
    SetupTurn True
End Sub


Private Sub mnuGameCredits_Click()
    MsgBox "Cruddy System Software." & vbCrLf & vbCrLf & "Please send all complaints to:" & vbCrLf & "1526 29th Place S.E." & vbCrLf & "Puyallup, WA 98374"
End Sub

Private Sub mnuGameEnd_Click()
    pctField.Cls
    lblStat = OPENING
End Sub

Private Sub mnuGameExit_Click()
    End
End Sub

Private Sub mnuGameNew_Click()
    SetupGame
End Sub

Sub Flash(ByVal Obj As Control)
Dim X1, Y1
    X1 = Obj.Left + Obj.Width / 2
    Y1 = Obj.Top + Obj.Height / 2
    Line (Obj.Left - 2, Obj.Top - 2)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 2)
    pctField.FillColor = SKY_BLUE
    FloodFill pctField.hdc, X1, Y1, SKY_BLUE
    pctField.FillColor = GREEN
End Sub

Sub TankFresh()
    Flash TankRed: Flash TankBlue
    XPos = TankRed.Left + TankRed.Width / 2
    YPos = TankRed.Top + TankRed.Height / 2
    pctField.ForeColor = RED
    Call BitBlt(pctField.hdc, TankRed.Left, TankRed.Top, TankRed.Width, TankRed.Height, TankRed.hdc, 0, 0, PNT)
    DrawTank RTank.Angle, LLength
    XPos = TankBlue.Left + TankBlue.Width / 2
    YPos = TankBlue.Top + TankBlue.Height / 2
    pctField.ForeColor = BLUE
    Call BitBlt(pctField.hdc, TankBlue.Left, TankBlue.Top, TankBlue.Width, TankBlue.Height, TankBlue.hdc, 0, 0, PNT)
    DrawTank BTank.Angle, LLength
    If RTurn = True Then
        scrAngle.Value = RTank.Angle
        scrPower.Value = RTank.Power
    Else
        scrAngle.Value = BTank.Angle
        scrPower.Value = BTank.Power
    End If
End Sub

Private Sub mnuGamePause_Click()
    If InGame = False Then PlaySound Er: Exit Sub
    If InPause = True Then PlaySound Er: Exit Sub
    If InFire Then PlaySound Er: Exit Sub
    InPause = Not InPause
    lblPause.Visible = InPause
End Sub

Private Sub pctField_KeyDown(KeyCode As Integer, Shift As Integer)
Static n As Integer
Dim i As Integer
    On Error Resume Next
    If KDwn = True Then n = n + 1 Else: n = 0
    KDwn = True
    i = 1
    If KeyCode = vbKeyLeft Then
        scrAngle.Value = scrAngle.Value - i
    ElseIf KeyCode = vbKeyRight Then
        scrAngle.Value = scrAngle.Value + i
    End If
    If n >= 100 Then i = 5 Else: i = 1
    If KeyCode = vbKeyDown Then
        scrPower.Value = scrPower.Value - i
    ElseIf KeyCode = vbKeyUp Then
        scrPower.Value = scrPower.Value + i
    End If
    If KeyCode = vbKeySpace Then cmdFire_Click
End Sub

Private Sub pctField_KeyUp(KeyCode As Integer, Shift As Integer)
    KDwn = False
End Sub

Private Sub scrAngle_Change()
    lblAng.Caption = 90 - Abs(scrAngle.Value)
    SetData
End Sub

Private Sub scrPower_Change()
    If InGame = False Then Exit Sub
    lblPow.Caption = scrPower.Value
    SetData
End Sub

Sub Inscript(ByVal Message As String)
    lblStat = Message
End Sub

Sub SetData()
    If InGame = False Then PlaySound Er: Exit Sub
    If InFire Then PlaySound Er: Exit Sub
    If RTurn = True Then
        RTank.Angle = scrAngle.Value
        RTank.Power = scrPower.Value
    Else
        BTank.Angle = scrAngle.Value
        BTank.Power = scrPower.Value
    End If
    TankFresh
End Sub


Private Sub tmTimer_Timer()
Static Done As Boolean, i As Integer
Dim Pix As Long, Pix2 As Long, Die As Boolean
    If InFire = False Then Exit Sub
    pctField.ScaleMode = 1
    pctField.AutoRedraw = True
    pctField.FillColor = SKY_BLUE
    pctField.Circle (BulX, BulY), 10, SKY_BLUE
    BulX = BulX + (Mov.x / 5)
    BulY = BulY - (Mov.y / 5)
    pctField.FillColor = ORANGE
    pctField.Circle (BulX, BulY), 8, ORANGE
    pctField.FillColor = GREEN
    Mov.y = Mov.y - BulVel
    Coordinate.Cls
    Coordinate.Print Int(pctField.Width - BulX / 15) & ", " & Int(pctField.Height - BulY / 15)
    If BulX < 0 Then BulX = pctField.Width * 15
    If BulX / 15 > pctField.Width Then BulX = 0
    If BulY / 15 > pctField.Height Then Done = True
    pctField.ScaleMode = 3
    Pix = pctField.Point(BulX / 15, BulY / 15 + 3)
    Pix2 = pctField.Point(BulX / 15, BulY / 15 - 3)
    If FreeBool Then i = 0: FreeBool = Not FreeBool
    i = i + 1
    TankFresh
    If Pix = GREEN Or Pix2 = GREEN Then
        Done = True
        Inscript "Green"
        PlaySound gDest
        pctField.FillColor = SKY_BLUE
        pctField.Circle (BulX, BulY), 10, SKY_BLUE
        pctField.FillColor = GREEN
    ElseIf Pix = BLUE Or Pix2 = BLUE Then
        If i > 10 And CInt(Abs(Mov.y)) <> BulVel Then
            pctField.FillColor = SKY_BLUE
            pctField.Circle (BulX / 15, BulY / 15), 10, SKY_BLUE
            pctField.FillColor = GREEN
            KillTank False
            Done = True
            Inscript bDieCapt
            Die = True
        End If
    ElseIf Pix = RED Or Pix2 = RED Then
        If i > 10 And CInt(Abs(Mov.y)) <> BulVel Then
            pctField.FillColor = SKY_BLUE
            pctField.Circle (BulX / 15, BulY / 15), 10, SKY_BLUE
            pctField.FillColor = GREEN
            KillTank True
            Done = True
            Inscript rDieCapt
            Die = True
        End If
    End If
    If Done = True Then
        tmTimer.Enabled = False
        InFire = False
        Done = False
        pctField.Refresh
        pctField.MousePointer = 1
        If Not Die Then SetupTurn Not RTurn Else: Exit Sub
        pctField.FillColor = SKY_BLUE
        pctField.Circle (BulX / 15, BulY / 15), 10, SKY_BLUE
        pctField.Circle (BulX, BulY), 30, SKY_BLUE
        pctField.FillColor = GREEN
        TankFresh
    End If
    pctField.AutoRedraw = True
End Sub

Sub SetupTurn(ByVal RedTurn As Boolean)
Dim Col As Long, Tnk As TankSettings
    RTurn = RedTurn
    If RedTurn Then
        Tnk = RTank
        Col = RED
        Inscript "Red's Turn"
    Else
        Tnk = BTank
        Col = BLUE
        Inscript "Blue's Turn"
    End If
    lblStat.ForeColor = Col
    scrAngle.Value = Tnk.Angle
    scrPower.Value = Tnk.Power
End Sub

Sub KillTank(ByVal RedKill As Boolean)
Dim TankDie As Object, TankDead As Object
Dim XPosition As Double, YPosition As Double
Dim TankScoreBoard As Label
Dim KillSoundFile As String
    If RedKill Then
        Set TankScoreBoard = lblBlueScore
        Set TankDie = RDie
        Set TankDead = RDead
        XPos = TankRed.Left
        YPos = TankRed.Top
        KillSoundFile = rDest
    Else
        Set TankScoreBoard = lblRedScore
        Set TankDie = BDie
        Set TankDead = BDead
        XPos = TankBlue.Left
        YPos = TankBlue.Top
        KillSoundFile = bDest
    End If
    TankTurn = RedKill
    InGame = False
    TankScoreBoard = Format(TankScoreBoard + 1, "0000")
    pctField.ForeColor = SKY_BLUE
    pctField.Line (XPos, YPos - 300)-(XPos + TankRed.Width, YPos + TankRed.Height), , BF
    Call BitBlt(pctField.hdc, XPos, YPos - TankRed.Height, TankRed.Width, TankRed.Height * 2, TankDie.hdc, 0, 0, PNT)
    PlaySound KillSoundFile
    Wait 1
    pctField.Line (XPos, YPos - 15)-(XPos + TankRed.Width, YPos + TankRed.Height), , BF
    Call BitBlt(pctField.hdc, XPos, YPos, TankRed.Width, TankRed.Height * 2, TankDead.hdc, 0, 0, PNT)
    cmdFire.Caption = TxtReset
End Sub
