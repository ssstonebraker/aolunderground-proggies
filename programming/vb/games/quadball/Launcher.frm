VERSION 5.00
Begin VB.Form Launcher 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quad-Ball Launcher, By Arvinder Sehmi 1999.  Arvinder@Sehmi.co.uk"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "Launcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar SpeedSetter 
      Height          =   195
      LargeChange     =   20
      Left            =   2385
      Max             =   200
      Min             =   10
      TabIndex        =   9
      Top             =   4680
      Value           =   10
      Width           =   3615
   End
   Begin VB.CheckBox LaunchExit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Exit After Launch?"
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
      Left            =   1890
      TabIndex        =   8
      Top             =   4050
      Value           =   1  'Checked
      Width           =   1950
   End
   Begin LaunchQuadball.Logo Logo 
      Height          =   1125
      Left            =   135
      TabIndex        =   5
      Top             =   3420
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1984
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "^\ Too Fast!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5805
      TabIndex        =   15
      Top             =   4995
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   735
      Left            =   90
      Top             =   4545
      Width           =   6900
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Higher Speed = Higher Score"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   4995
      Width           =   2625
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "^\ Expert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4140
      TabIndex        =   13
      Top             =   4950
      Width           =   870
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "^\ Novice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2475
      TabIndex        =   12
      Top             =   4950
      Width           =   960
   End
   Begin VB.Label SetSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Height          =   285
      Left            =   6255
      TabIndex        =   11
      Top             =   4635
      Width           =   555
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Comet Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   4635
      Width           =   2400
   End
   Begin VB.Image Title 
      Height          =   1575
      Left            =   90
      Picture         =   "Launcher.frx":09BA
      Top             =   180
      Width           =   6930
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFF00&
      X1              =   135
      X2              =   1710
      Y1              =   1890
      Y2              =   1890
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFF00&
      X1              =   3915
      X2              =   5400
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Label Train 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   2025
      TabIndex        =   3
      Top             =   3285
      Width           =   1485
   End
   Begin VB.Shape TrainShape 
      BorderColor     =   &H00FF0000&
      Height          =   555
      Left            =   2025
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   1500
   End
   Begin VB.Label Story 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   5265
      TabIndex        =   4
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Shape StoryShape 
      BorderColor     =   &H00FF0000&
      Height          =   555
      Left            =   5265
      Shape           =   2  'Oval
      Top             =   3825
      Width           =   1500
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      X1              =   3915
      X2              =   5265
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFF00&
      X1              =   3915
      X2              =   6705
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      X1              =   3915
      X2              =   3915
      Y1              =   1755
      Y2              =   4095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      X1              =   135
      X2              =   2025
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      X1              =   135
      X2              =   3195
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      X1              =   135
      X2              =   135
      Y1              =   1890
      Y2              =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quad-Ball Training Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   1935
      Width           =   3210
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quad-Ball Story Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   1845
      Width           =   3210
   End
   Begin VB.Label Info 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "If You Have Any Problems, Questions or Comments Please Contact Me: Arvinder@Sehmi.Co.Uk"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   5445
      Width           =   6795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Launcher.frx":5032
      ForeColor       =   &H0000FFFF&
      Height          =   1905
      Left            =   4005
      TabIndex        =   7
      Top             =   2430
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "In This Mode Of Quad-Ball You Must Try And Keep The Comet On The Screen For As Long As Possible To Beat The Top Score."
      ForeColor       =   &H0000FFFF&
      Height          =   825
      Left            =   270
      TabIndex        =   6
      Top             =   2430
      Width           =   2895
   End
End
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯'
' This Is Just A nice interface in which '
' to launch Quad-Ball, other then that is'
' has no special job.                    '
'________________________________________'
Dim StoryHighlight As Boolean
Dim TrainHighlight As Boolean
Private Sub Form_Load()
 StoryHighlight = False
 TrainHighlight = False
 Me.Show
 Me.Refresh
 Logo.Start  'start the spinning logo
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Logo.StopAnimation ' stop spinning logo
End Sub
'the Following 6 subs
'unhighlight the label buttons
'(when the mouse moves over then)
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub

Private Sub SpeedSetter_Change()
 SetSpeed = SpeedSetter.Value
End Sub
Private Sub SpeedSetter_Scroll()
 SetSpeed = SpeedSetter.Value
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If TrainHighlight = True Then Unhighlight False, True
 If StoryHighlight = True Then Unhighlight True, False
End Sub
'highlight the story button
Private Sub Story_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If TrainHighlight = True Then Unhighlight False, True
If StoryHighlight = True Then Exit Sub
With Story
.FontBold = True
.FontSize = 16
.Top = 3900
End With
StoryShape.BorderWidth = 3
StoryShape.BorderColor = vbCyan
WAVPlay "click.qbs"
StoryHighlight = True
End Sub
'highlight the train button
Private Sub Train_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If StoryHighlight = True Then Unhighlight True, False
If TrainHighlight = True Then Exit Sub
With Train
 .FontBold = True
 .FontSize = 16
 .Top = 3225
End With
TrainShape.BorderWidth = 3
TrainShape.BorderColor = vbCyan
WAVPlay "click.qbs"
TrainHighlight = True
End Sub
'load Q-Ball Training
Private Sub Train_Click()
 ThisDir
 Shell "QuadBall_Training.exe " & Trim(Str(SpeedSetter.Value)), vbNormalFocus
 If LaunchExit.Value = vbChecked Then Unload Me: End
End Sub
'Loa d Q-Ball Story
Private Sub Story_Click()
 ThisDir
 Shell "QuadBall.exe " & Trim(Str(SpeedSetter.Value)), vbNormalFocus
 If LaunchExit.Value = vbChecked Then Unload Me: End
End Sub
' unhilights specified button on call
Public Sub Unhighlight(StoryLabel As Boolean, TrainLabel As Boolean)
If TrainLabel = True Then
 With Train
  .FontBold = False
  .FontSize = 12
  .Top = 3285
 End With
 TrainShape.BorderWidth = 1
 TrainShape.BorderColor = vbBlue
 TrainHighlight = False
End If
If StoryLabel = True Then
 With Story
 .FontBold = False
 .FontSize = 12
 .Top = 3960
 End With
 StoryShape.BorderWidth = 1
 StoryShape.BorderColor = vbBlue
 StoryHighlight = False
End If
End Sub

