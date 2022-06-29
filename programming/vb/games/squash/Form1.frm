VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Squash Tennis - Second Edition by Jongmin Baek.  / You can't win against computer! If you can, you are a genius!"
   ClientHeight    =   6810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11355
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton SelectDevice 
      Caption         =   "Mouse"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   855
   End
   Begin VB.OptionButton SelectDevice 
      Caption         =   "Keyboard"
      Height          =   375
      Index           =   0
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CheckBox PCoption 
      Caption         =   "Play with PC"
      Height          =   255
      Left            =   9360
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Score1 
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Player2 : 0"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Score0 
      Height          =   375
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Player1 : 0"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   9360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0442
      Top             =   0
      Width           =   1935
   End
   Begin VB.PictureBox Board 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   7031.46
      ScaleMode       =   0  'User
      ScaleWidth      =   1003.263
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   5880
         Top             =   2280
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "My Term!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Term!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Shape mybar 
         BackColor       =   &H80000018&
         BackStyle       =   1  'Opaque
         Height          =   150
         Index           =   1
         Left            =   5880
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Shape mybar 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         Height          =   150
         Index           =   0
         Left            =   3120
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Shape Bomb 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   4598
         Shape           =   3  'Circle
         Top             =   238
         Width           =   495
      End
      Begin VB.Image Flame 
         Height          =   2040
         Left            =   4368
         Picture         =   "Form1.frx":048D
         Top             =   -1080
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Select Device (You can use Mouse when you play with PC)"
      Height          =   615
      Left            =   9360
      TabIndex        =   18
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum : 300"
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Now Speed is..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9360
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "By click, a new game will start."
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "(PC will play as player2)"
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label TermSet 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   7
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Now Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "To Start or Stop, Press F2 key."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Score "
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      Begin VB.Menu StartStop 
         Caption         =   "Start"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProgram 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Messaged As Boolean
Dim Target As Double
Dim S0 As Integer
Dim S1 As Integer
Dim Term As Integer
Dim Speed As Double
Dim SetRad As Double
Dim RedC As Integer
Dim Increase As Boolean
Dim OriginalKeycode As Integer
Dim Direction As Double
Private Sub Board_KeyDown(KeyCode As Integer, Shift As Integer)
If SelectDevice(1).Value = True Then Exit Sub
OriginalKeycode = KeyCode
KeyPressOperation
End Sub
Private Sub KeyPressOperation()
'If StartStop.Caption = "Start" Then Exit Sub
Select Case OriginalKeycode
    Case 65, 97
        If mybar(0).Left > 0 Then mybar(0).Left = mybar(0).Left - 50
    Case 68, 100
        If mybar(0).Left + mybar(0).Width < Board.ScaleWidth Then mybar(0).Left = mybar(0).Left + 50
    Case 83, 113
        If mybar(0).Top + 200 < 6420 Then mybar(0).Top = mybar(0).Top + 200
    Case 87, 117
        If mybar(0).Top - 200 > 5000 Then mybar(0).Top = mybar(0).Top - 200
End Select
If PCoption = 0 Then
    Select Case OriginalKeycode
        Case 74, 104
            If mybar(1).Left > 0 Then mybar(1).Left = mybar(1).Left - 50
        Case 76, 106
            If mybar(1).Left + mybar(1).Width < Board.ScaleWidth Then mybar(1).Left = mybar(1).Left + 50
        Case 73, 103
            If mybar(1).Top - 200 > 5000 Then mybar(1).Top = mybar(1).Top - 200
        Case 75, 105
            If mybar(1).Top + 200 < 6420 Then mybar(1).Top = mybar(1).Top + 200
    End Select
End If
End Sub
Private Sub Board_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SelectDevice(0).Value = True Then Exit Sub
B = 50
center = mybar(0).Left + mybar(0).Width / 2
If center > X + B / 2 And mybar(0).Left > B Then mybar(0).Left = mybar(0).Left - B
If center + B / 2 < X And mybar(0).Left + mybar(0).Width + B < Board.ScaleWidth Then mybar(0).Left = mybar(0).Left + B
'center2 = mybar(0).Top
'If center2 + 200 < 6420 And center2 < Y Then mybar(0).Top = mybar(0).Top + 200
'If center2 - 200 > 5000 And center2 > Y Then mybar(0).Top = mybar(0).Top - 200
End Sub
Private Sub ExitProgram_Click()
End
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If SelectDevice(1).Value = True Then Exit Sub
OriginalKeycode = KeyCode
KeyPressOperation
End Sub
Private Sub Form_Load()
mybar(0).Top = 6040
mybar(1).Top = 6418
Bomb.Left = 500
Flame.Left = 475
SetRad = 3.141592654 / 180
RedC = 0
Term = 0
Direction = 0
Speed = 300
Flame.Visible = False
Label12.Caption = LTrim$(Str$(Speed))
Increase = True
Bomb.Top = 0
Flame.Top = Bomb.Top - 1380
End Sub
Private Sub Optionmenu_Click()
StartStop.Caption = "Start"
Form1.Enabled = False
Form2.Show
End Sub
Private Sub PCoption_Click()
mybar(0).Top = 6040
mybar(1).Top = 6418
Bomb.Left = 500
Flame.Left = 475
SetRad = 3.141592654 / 180
RedC = 0
Term = 0
Direction = 0
Speed = 300
Label12.Caption = LTrim$(Str$(Speed))
Flame.Visible = False
Increase = True
Bomb.Top = 0
Flame.Top = Bomb.Top - 1380
S1 = 0: S0 = 0: Score0.Text = "Player1 : 0": Score1.Text = "Player2 : 0"
StartStop.Caption = "Start"
If PCoption.Value = 0 Then Label6.Visible = False: Label7.Visible = False: SelectDevice(1).Enabled = False: SelectDevice(0).Value = True
If PCoption.Value = 1 Then Label6.Visible = True: Label7.Visible = False: SelectDevice(1).Enabled = True
Board.SetFocus
End Sub
Private Sub Score1_GotFocus()
Board.SetFocus
End Sub
Private Sub Score0_GotFocus()
Board.SetFocus
End Sub
Private Sub SelectDevice_GotFocus(Index As Integer)
Board.SetFocus
End Sub
Private Sub StartStop_Click()
StartStop.Caption = IIf(StartStop.Caption = "Start", "Stop", "Start")
If StartStop.Caption = "Stop" Then GG = False
Board.SetFocus
End Sub
Private Sub Text1_GotFocus()
Board.SetFocus
End Sub
Private Sub Timer1_Timer()
If Messaged = False Then
    Messaged = True
    MsgBox "You are about to play my application! Enjoy it!", , "Jongmin Baek"
    MsgBox "E-mail me - chunjaemanse3@netzero.net", , "Jongmin Baek"
End If
RedC = RedC + IIf(Increase = True, 5, -5)
If RedC < 0 Then RedC = 0: Increase = True
If RedC > 255 Then RedC = 255: Increase = False
Bomb.BackColor = RGB(255, RedC, RedC)
mybar(0).BackColor = RGB(255 - RedC, 255 - RedC, 255)
mybar(1).BackColor = RGB(255 - RedC, 255, 255 - RedC)

If Bomb.Top > Board.ScaleHeight Then
If Term = 0 Then S1 = S1 + 1 Else S0 = S0 + 1
Score0.Text = "Player1 :" + Str$(S0)
Score1.Text = "Player2 :" + Str$(S1)
If Term = 0 Then Target = Bomb.Left + Bomb.Width / 2 Else If PCoption.Value = 1 Then MsgBox "Rats!", , "I like to say..."
Term = 1 - Term
If PCoption.Value = 1 Then
    If Term = 0 Then Label6.Visible = True: Label7.Visible = False Else Label6.Visible = False: Label7.Visible = True
End If
TermSet.Caption = LTrim$(Str$(Term + 1))
Direction = 0
Speed = 300
Label12.Caption = LTrim$(Str$(Speed))
Flame.Visible = False
Increase = True
Bomb.Top = 0
Flame.Top = Bomb.Top - 1380
End If

If StartStop.Caption = "Start" Then Exit Sub
If PCoption = 1 Then
    If Target - mybar(1).Width / 2 < 0 Then Target = mybar(1).Width / 2
    If Target + mybar(1).Width / 2 > Board.ScaleWidth Then Target = Board.ScaleWidth - mybar(1).Width / 2
    f = Target
    If Term = 0 Then f = Board.ScaleWidth / 2
    f2 = mybar(1).Left + mybar(1).Width / 2: xh = 0
    If f > f2 + 20 Then xh = 40
    If f < f2 - 20 Then xh = -40
    If Bomb.Left + xh < 0 Or Bomb.Left + Bomb.Width + xh > Board.ScaleWidth Then GoTo 23
    mybar(1).Left = mybar(1).Left + xh
23 End If
MX = Sin(Direction * SetRad) * Speed
MY = Cos(Direction * SetRad) * Speed
center = Bomb.Left + Bomb.Width / 2
center2 = Bomb.Top + Bomb.Height
If MX + Bomb.Left < 0 Or MX + Bomb.Left + Bomb.Width > Board.ScaleWidth Then GoTo 10
If Bomb.Top + MY < 0 Then GoTo 20
Bomb.Left = Bomb.Left + MX
Flame.Left = Bomb.Left - 25
Bomb.Top = Bomb.Top + MY
Flame.Top = Bomb.Top - 1380
If center > mybar(Term).Left + mybar(Term).Width Or center < mybar(Term).Left Then GoTo 30
If center2 < mybar(Term).Top - MY Or center2 > mybar(Term).Top Then GoTo 30
distant = Abs(mybar(Term).Left + mybar(Term).Width / 2 - center)
Speed = Speed + 10
Label12.Caption = LTrim$(Str$(Speed))
If Speed > 420 Then Flame.Visible = True
Max = mybar(Term).Width / 2
If Direction >= 270 Then
    Direction = Direction + distant / Max * 15
Else
    Direction = Direction - distant / Max * 15
End If
If Direction < 0 Then Direction = Direction + 360
If Direction > 360 Then Direction = Direction - 360
Direction = 180 - Direction
If Direction < 0 Then Direction = Direction + 360
If Direction <= 120 Then Direction = 120
If Direction >= 240 Then Direction = 240
Term = 1 - Term
If Term = 0 Then Label6.Visible = True: Label7.Visible = False Else Label6.Visible = False: Label7.Visible = True
TermSet.Caption = LTrim$(Str$(Term + 1))
30 GoTo SetTarget
10 Direction = 360 - Direction: Speed = Speed + 10
Label12.Caption = LTrim$(Str$(Speed))
If Speed > 420 Then Flame.Visible = True
GoTo SetTarget
20 Direction = 180 - Direction: Speed = Speed + 10
Label12.Caption = LTrim$(Str$(Speed))
If Speed > 420 Then Flame.Visible = True
If Direction < 0 Then Direction = Direction + 360

SetTarget: If Term = 1 Then
        Dist = mybar(1).Top - Bomb.Top - Bomb.Height
        If Direction >= 90 And Direction <= 270 Then Dist = -(mybar(1).Top + Bomb.Top - Bomb.Height)
        GG = Tan(Direction * SetRad) * Dist + center
137     If GG > Board.ScaleWidth Then GG = Board.ScaleWidth * 2 - GG: GoTo 137
        If GG < 0 Then GG = -GG: GoTo 137
        Target = GG
    End If
End Sub
