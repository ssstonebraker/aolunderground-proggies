VERSION 5.00
Begin VB.Form frmGameOver 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   Icon            =   "frmGameOver.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameOver.frx":000C
   ScaleHeight     =   2085
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   195
      TabIndex        =   8
      Top             =   1440
      Width           =   840
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1320
      X2              =   3600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label labRTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblRightTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Right Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblLeftTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Left Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1605
      TabIndex        =   5
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label lblBonusTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bonus Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1335
      TabIndex        =   4
      Top             =   900
      Width           =   1320
   End
   Begin VB.Label lblGameTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1380
      TabIndex        =   3
      Top             =   1680
      Width           =   1290
   End
   Begin VB.Label labLTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label labBTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BTotal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   900
      Width           =   855
   End
   Begin VB.Label labGTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  
  'Unload form
  Unload Me

End Sub

Private Sub Form_Load()
  
  'Play clapping wav sound
  If frmMain.mnuSound.Checked Then PlaySound App.Path & "\Clap.wav"
  
  'Fill totals
  labLTotal.Caption = lColTotal '//Left column total without bonus
  labBTotal.Caption = Bonus63 '//Left column bonus points
  labRTotal.Caption = rColTotal '//Right column total
  labGTotal.Caption = GameTotal '//Game total

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  
  '//User pressed the escape key
  If KeyCode = vbKeyEscape Then Unload Me

End Sub

