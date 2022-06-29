VERSION 5.00
Begin VB.Form frmVBtris 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Full Moon Tetris"
   ClientHeight    =   5055
   ClientLeft      =   1050
   ClientTop       =   1155
   ClientWidth     =   4410
   FillStyle       =   0  'Solid
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "VBtris32.frx":0000
   ScaleHeight     =   337
   ScaleMode       =   0  'User
   ScaleWidth      =   294
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBoard 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   4380
      Left            =   240
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   2
      Top             =   360
      Width           =   2460
   End
   Begin VB.PictureBox picBoard2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00C0C0C0&
      Height          =   4380
      Left            =   5640
      ScaleHeight     =   290
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   162
      TabIndex        =   1
      Top             =   120
      Width           =   2460
   End
   Begin VB.PictureBox picNext 
      Height          =   1260
      Left            =   2880
      ScaleHeight     =   1200
      ScaleWidth      =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   2880
      TabIndex        =   4
      Top             =   4080
      Width           =   1140
   End
   Begin VB.Label lblLines 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   1140
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   2880
      TabIndex        =   7
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Next Piece:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   3840
      Width           =   675
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines:"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   795
   End
End
Attribute VB_Name = "frmVBtris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------
'This form contains the user interface of the game
'-------------------------------------------------------

Private Sub Form_Activate()
'-------------------------------------------------------
'Refresh the board
'-------------------------------------------------------
Dim Temp
Temp = BitBlt(Board.BoardDC, 0, 0, 160, 288, Board.B2DC, 1, 1, SRCCOPY)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------
'Moves the pieces left or right, rotates them, or speeds
'there descent provided that a game is being played and
'that they haven't stopped moving down.  Also, pauses
'and unpauses the game
'-------------------------------------------------------
If Board.Game And NewPiece = False And PauseTheGame = False Then
    Select Case KeyCode
        Case vbKeyLeft
            MovePieceLeft 'Move the piece left
        Case vbKeyRight
            MovePieceRight 'Move the piece right
        Case vbKeyClear, vbKeyUp
            RotatePiece    'Rotate the piece
        Case vbKeyDown
            FallPiece = True 'Speed the descent
            'This records the Y position of the piece
            'when it starts its rapid descent for
            'score keeping purposes.  Each line it
            'falls after this is worth one point.
            If FallY = 0 Then FallY = Board.PieceY
        Case vbKeyP    'Pause the game
            PauseTheGame = True
    End Select
ElseIf Board.Game And PauseTheGame And KeyCode = vbKeyP Then
    PauseTheGame = False    'Unpause the game
End If


End Sub

Private Sub Form_Load()
'-------------------------------------------------------
'Retrieve a picture from frmPics to prevent delays
'during a game
'-------------------------------------------------------
picNext = frmPics.Next(1)
picNext = LoadPicture("")
'-------------------------------------------------------
'Run the procedures involved with moving the pieces
'so there won't be any delays when they are first
'called
'-------------------------------------------------------
Board.CurPiece = 5
Board.PieceX = 5
Board.PieceY = 5
MovePieceDown
MovePieceLeft
MovePieceRight
RotatePiece

'-------------------------------------------------------
'Load the high scores so they can be displayed before a
'game is played
'-------------------------------------------------------
GetScores
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Form_Resize()
'-------------------------------------------------------
'Pause the game if the form is minimized and a game is
'in progress and set AutoRedraw to true on picBoard so
'the gameboard will still appear there when the form is
'restored.
'-------------------------------------------------------
If frmVBtris.WindowState = 0 And Board.Game Then
    PauseTheGame = True
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'-------------------------------------------------------
'End the program
'-------------------------------------------------------
DoEvents
Dim Temp As Long

If PlaySounds Then
    If InStr(App.Path, " ") Then
        Temp = mciSendString("close " & Chr(34) & App.Path & "\" & MUSIC & Chr(34), 0&, 0, 0)
    Else
        Temp = mciSendString("close " & App.Path & "\" & MUSIC, 0&, 0, 0)
    End If
End If
Unload Me

End Sub

Private Sub mnuAbout_Click()
'-------------------------------------------------------
'Display the about form and pause the game if one is in
'progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmAbout.Show 1

End Sub

Private Sub mnuEndGame_Click()
'-------------------------------------------------------
'End the current game
'-------------------------------------------------------
GameOver = True

End Sub

Private Sub mnuExit_Click()
On Error Resume Next
'-------------------------------------------------------
'Unload the form to call the Form_Unload procedure.
'Attempting to end any other way has resulted in many
'a fatal error in VB32.EXE
'-------------------------------------------------------
Unload Me


End Sub

Private Sub mnuHighScore_Click()
'-------------------------------------------------------
'Display the high scores
'-------------------------------------------------------
DisplayHighScores

End Sub

Private Sub mnuInstruct_Click()
'-------------------------------------------------------
'Display the instructions and pause the game if one is
'in progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmInstruct.Show 1

End Sub

Private Sub mnuNewGame_Click()
'-------------------------------------------------------
'Disable the new game and high score menus while
'enabling the end game and pause menus.
'-------------------------------------------------------
mnuNewGame.Enabled = False
mnuEndGame.Enabled = True
mnuHighScore.Enabled = False
mnuPause.Enabled = True
'-------------------------------------------------------
'Start the game
'-------------------------------------------------------
NewGame

End Sub

Private Sub mnuOptions_Click()
'-------------------------------------------------------
'Display the options form, fills in the options
'accordingly, and pause the game if one is in progress
'-------------------------------------------------------
If Board.Game Then
    PauseTheGame = True
End If
frmOptions.txtStartingLevel = StartingLevel
If FillLines Then
    frmOptions.chkFillLines.Value = 1
Else
    frmOptions.chkFillLines.Value = 0
End If
If PlaySounds Then
    frmOptions.chkPlaySounds.Value = 1
Else
    frmOptions.chkPlaySounds.Value = 0
End If
If HideSplash Then
    frmOptions.chkSkipIntro.Value = 1
Else
    frmOptions.chkSkipIntro.Value = 0
End If
frmOptions.Show 1

End Sub

Private Sub mnuPause_Click()
'-------------------------------------------------------
'Pause or unpause the game
'-------------------------------------------------------
PauseTheGame = Not (PauseTheGame)

End Sub


Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()
Form2.PopupMenu Form2.Label4, 2
End Sub

Private Sub Label6_Click()
Form2.PopupMenu Form2.Label5, 2
End Sub

Private Sub Label7_Click()
Unload frmVBtris
End Sub

Private Sub Label8_Click()
frmVBtris.WindowState = 1
End Sub

Private Sub lblScore_Click()

End Sub
