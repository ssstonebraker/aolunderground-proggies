VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kill Bill"
   ClientHeight    =   6705
   ClientLeft      =   180
   ClientTop       =   195
   ClientWidth     =   8955
   Icon            =   "form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "EASY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   6120
      Width           =   6495
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4920
      Picture         =   "form1.frx":030A
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "form1.frx":61CC
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1320
      Picture         =   "form1.frx":C08E
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   3240
      Picture         =   "form1.frx":11F50
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrSeconds 
      Left            =   600
      Top             =   960
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2400
      Picture         =   "form1.frx":17E12
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   6120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   480
      Picture         =   "form1.frx":1DCD4
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1920
      Picture         =   "form1.frx":23B96
      ScaleHeight     =   1575
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   240
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
           Option Explicit
           Dim Difficulty
           Dim DeltaX
           Dim DeltaY
           Dim gTimerSpeed
           Dim gGameOn As Boolean
           Dim gHit As Boolean
           Dim gSeconds
           Dim gShots
           Dim gHits
           Dim gTime


   Private Sub cmdExit_Click()


       Unload Form1
   End Sub



   Private Sub cmdReset_Click()


       Picture1.Picture = Picture2.Picture
       Screen.MousePointer = vbCrosshair
       Timer1.Interval = 0
       DeltaX = Difficulty * 100 ' Initialize variables.
       DeltaY = Difficulty * 100
       cmdStart.Visible = True
       cmdExit.Visible = True
       Command1.Visible = True
       cmdReset.Visible = False
   End Sub



   Private Sub cmdStart_Click()


       Picture1.Picture = Picture2.Picture
       Screen.MousePointer = vbCrosshair
       Timer1.Interval = 1
       gTimerSpeed = 1
       DeltaX = Difficulty * 100
       DeltaY = Difficulty * 100
       EnableTrap Form1
       gHits = 0
       cmdReset.Visible = False
       cmdStart.Visible = False
       cmdExit.Visible = False
       Command1.Visible = False
       gHit = False
       gShots = 0
       gGameOn = True
       gTime = 0
       tmrSeconds.Interval = 1000
   End Sub
Private Sub Command1_Click()
If Difficulty = 1 Then
Difficulty = 2
Command1.Caption = "MEDIUM"
ElseIf Difficulty = 2 Then
Difficulty = 3
Command1.Caption = "HARD"
ElseIf Difficulty = 3 Then
Difficulty = 6
Command1.Caption = "IMPOSSIBLE"
ElseIf Difficulty = 6 Then
Difficulty = 1
Command1.Caption = "EASY"
End If
End Sub

Private Sub Form_Load()
Difficulty = 1
End Sub

   Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


       gShots = gShots + 1
   End Sub



   Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)



       If gGameOn = True Then
       gShots = gShots + 1
       gHits = gHits + 1
       If gHits = 1 Then Picture1.Picture = Picture3.Picture
       If gHits = 2 Then Picture1.Picture = Picture4.Picture
       If gHits = 3 Then Picture1.Picture = Picture5.Picture
       If gHits = 4 Then Picture1.Picture = Picture6.Picture
       If gHits = 5 Then Picture1.Picture = Picture7.Picture
       
       If gHits = 5 Then
       DisableTrap Form1
       Timer1.Interval = 0
       gGameOn = False
       Screen.MousePointer = Default
       cmdReset.Visible = True
       cmdExit.Visible = True
       Command1.Visible = True
       gHit = True
       tmrSeconds.Interval = 0
       MsgBox "It took you " & gShots & " shots and " & gTime & " seconds to kill him!"
       
              End If
           


           
           
           Exit Sub
       ElseIf gGameOn = False Then
           Exit Sub
       End If


   End Sub



   Private Sub Timer1_Timer()



       If gHit = True Then
           Timer1.Interval = 0
           Exit Sub
       End If


       If gTimerSpeed < 50 Then gTimerSpeed = gTimerSpeed + 1
       Timer1.Interval = gTimerSpeed
       Picture1.Move Picture1.Left + DeltaX, Picture1.Top + DeltaY
       If Picture1.Left < ScaleLeft Then DeltaX = Difficulty * 100


       If Picture1.Left + Picture1.Width > ScaleWidth + ScaleLeft Then
           DeltaX = -Difficulty * 100
       End If


       If Picture1.Top < ScaleTop Then DeltaY = Difficulty * 100


       If Picture1.Top + Picture1.Height > ScaleHeight + ScaleTop Then
           DeltaY = -Difficulty * 100
       End If


   End Sub



   Private Sub tmrSeconds_Timer()


       gTime = gTime + 1
   End Sub
