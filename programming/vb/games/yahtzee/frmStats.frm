VERSION 5.00
Begin VB.Form frmStats 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Statistics"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmStats.frx":0000
   ScaleHeight     =   2055
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgReset 
      Height          =   330
      Left            =   480
      Picture         =   "frmStats.frx":1CF2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Image imgOK 
      Height          =   330
      Left            =   2760
      Picture         =   "frmStats.frx":1FDF
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTotalAvgScore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   120
   End
   Begin VB.Label lblTGPlayed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   120
   End
   Begin VB.Label lblAverageScores 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Average Score:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblTotalGames 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Games Played:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1305
      TabIndex        =   0
      Top             =   600
      Width           =   1470
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgOK_Click()
  
  '//Unload form
  Unload Me

End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture to down pic
  imgOK.Picture = frmMain.ImageList1.ListImages(4).Picture

End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change pic back to normal
  imgOK.Picture = frmMain.ImageList1.ListImages(3).Picture

End Sub

Private Sub imgReset_Click()
  
  Call ResetStats '//Reset was clicked, set back to default in registry
  Call InitStats '//Call sub to update the on screen stats to default

End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture to down pic
  imgReset.Picture = frmMain.ImageList1.ListImages(6).Picture

End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  '//Change picture to up pic
  imgReset.Picture = frmMain.ImageList1.ListImages(5).Picture

End Sub

Private Sub Form_Load()
  
  Call InitStats '//Call sub to update the on screen stats to registry values

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  '//User pressed the escape key
  If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub InitStats()
Dim Stat1 As Long '//Tracks total games played
Dim Stat2 As Long '//Tracks total score of all games played
  
  Stat1 = CLng(GetSetting(MyApp, "Settings", "GamesPlayed")) '//Get value from registry
  Stat2 = CLng(GetSetting(MyApp, "Settings", "TotalScore")) '//Get value from registry
  
  If Stat1 = 0 Then '//Games played was zero, just update onscreen values to zero
    
    lblTGPlayed.Caption = "0"
    lblTotalAvgScore.Caption = "0"
  
  Else
   
  '//Set viewable games played total to registy value
  lblTGPlayed.Caption = Stat1
  
  '//Average the score based on games played and total score
  lblTotalAvgScore.Caption = Round((Stat2 / Stat1), 0)

  End If
  
End Sub
