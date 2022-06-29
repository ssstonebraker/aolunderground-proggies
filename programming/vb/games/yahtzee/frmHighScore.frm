VERSION 5.00
Begin VB.Form frmHighScore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Score"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmHighScore.frx":0000
   ScaleHeight     =   3435
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1725
      Top             =   1140
   End
   Begin VB.Image imgReset 
      Height          =   330
      Left            =   360
      Picture         =   "frmHighScore.frx":1EB52
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image imgOK 
      Height          =   330
      Left            =   2340
      Picture         =   "frmHighScore.frx":1EE3F
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2310
      Width           =   240
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1650
      Width           =   240
   End
   Begin VB.Label labS1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   4
      Left            =   2760
      TabIndex        =   12
      Top             =   2640
      Width           =   675
   End
   Begin VB.Label labS1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   3
      Left            =   2760
      TabIndex        =   11
      Top             =   2310
      Width           =   675
   End
   Begin VB.Label labS1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   2
      Left            =   2760
      TabIndex        =   10
      Top             =   1980
      Width           =   675
   End
   Begin VB.Label labS1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   1650
      Width           =   675
   End
   Begin VB.Label labP1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   2085
   End
   Begin VB.Label labP1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   2310
      Width           =   2085
   End
   Begin VB.Label labP1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   1980
      Width           =   2085
   End
   Begin VB.Label labP1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   1650
      Width           =   2085
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label labS1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label labP1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2085
   End
   Begin VB.Line Line1 
      X1              =   1250
      X2              =   3775
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblBestPlayers 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Best Players Of All Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  '//User pressed the escape key
  If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
  
  '//Get highscores
  Call GetScores
  
  For i = 0 To 4 '//Move high scores name text off of form
   labP1(i).Left = -1500
   lbl1(i).Left = -1860
  Next i
  
  Timer1.Enabled = True '//Enable timer so they can start moving into view
 
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
  
  For i = 0 To 4
   lbl1(i).Left = lbl1(i).Left + 120 '//Move player rank numbers in
   labP1(i).Left = labP1(i).Left + 120 '//Move player names in
   
   '//They are in place, stop timer and exit timer sub
   If labP1(4).Left >= 480 Then Timer1.Enabled = False: Exit Sub
  Next i
  
End Sub

Private Sub imgOK_Click()

  '//Unload form
  Unload Me

End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture on mouse down
  imgOK.Picture = frmMain.ImageList1.ListImages(4).Picture

End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Reset picture on mouse up
  imgOK.Picture = frmMain.ImageList1.ListImages(3).Picture

End Sub

Private Sub imgReset_Click()
  
  '//Reset high scores
  Call ResetHighScores
  
  '//Get scores again to set the current screen to correct scores in registry
  Call GetScores
  
End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture on mouse down
  imgReset.Picture = frmMain.ImageList1.ListImages(6).Picture
  
End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Reset picture on mouse up
  imgReset.Picture = frmMain.ImageList1.ListImages(5).Picture

End Sub

Private Sub GetScores()

  '//Get all high scores from registry
  labP1(0).Caption = GetSetting(MyApp, "Settings", "HSName0", "Default")
  labP1(1).Caption = GetSetting(MyApp, "Settings", "HSName1", "Default")
  labP1(2).Caption = GetSetting(MyApp, "Settings", "HSName2", "Default")
  labP1(3).Caption = GetSetting(MyApp, "Settings", "HSName3", "Default")
  labP1(4).Caption = GetSetting(MyApp, "Settings", "HSName4", "Default")
  labS1(0).Caption = GetSetting(MyApp, "Settings", "HSScore0", "Default")
  labS1(1).Caption = GetSetting(MyApp, "Settings", "HSScore1", "Default")
  labS1(2).Caption = GetSetting(MyApp, "Settings", "HSScore2", "Default")
  labS1(3).Caption = GetSetting(MyApp, "Settings", "HSScore3", "Default")
  labS1(4).Caption = GetSetting(MyApp, "Settings", "HSScore4", "Default")

 
  If frmMain.pBar.Value = 13 Then '//Game is over highlight your highscore
    
    labP1(HSPosition).FontUnderline = True
    labS1(HSPosition).FontUnderline = True
    labP1(HSPosition).ForeColor = &H8000&
    labS1(HSPosition).ForeColor = &H8000&
  
  End If
  
End Sub
