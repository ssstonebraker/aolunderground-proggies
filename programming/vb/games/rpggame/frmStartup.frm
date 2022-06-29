VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6240
      Left            =   4320
      Picture         =   "frmStartup.frx":0000
      ScaleHeight     =   6240
      ScaleWidth      =   4620
      TabIndex        =   6
      Top             =   360
      Width           =   4620
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Save Game"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   2550
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resume Game"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   3345
   End
   Begin VB.Image imgTitle 
      Height          =   1230
      Left            =   120
      Picture         =   "frmStartup.frx":13A87
      Top             =   120
      Width           =   2685
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   4920
      Width           =   1830
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit Game"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   4
      Left            =   480
      TabIndex        =   2
      Top             =   6240
      Width           =   2520
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Load Game"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   2
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   2640
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Width           =   2550
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim current As Integer  'holds the currently selected label

Private Sub Form_Load()
    
    Dim slash As String
    
    'show that the origional screen settings have not yet been saved
    OldBPP = 0
    
    'change the settings to 640 x 480 x 16 for fastest possible game speeds
    InitializeRes
    
    'set the initial label to New
    current = 1
    
    'see if the end of the path has a back slash on it
    slash = ""
    If Right(App.Path, 1) <> "\" Then slash = "\"
    
    sndStep = App.Path & slash & "1.wav"
    sndButton = App.Path & slash & "2.wav"
    
End Sub

Private Sub imgTitle_Click()
        
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'disable this window
    frmStartup.Enabled = False
    'show the about window
    frmAbout.Show
    
End Sub

Private Sub lblButton_Click(Index As Integer)
    
    Call sndPlaySound(sndButton, &H1)
    'act based on which label is click on
    Select Case (Index)
    
    Case Is = 0 'the resume game button
        'check to see if there is a game in progress
        If GameinProgress = True Then
            'hide this form, and show the main one
            frmDisplay.Show
            frmStartup.Hide
        Else
            'my custom message box
            Call msg("No game in progress, cannot resume", frmStartup)
        End If
    Case Is = 1 'the new game button
        'show the resume game button, for later
        lblButton(0).Visible = True
        lblButton(0).ForeColor = QBColor(12)
        lblButton(1).ForeColor = QBColor(15)
        current = 0
        'hide this form, show the new game window
        frmNew.Show
        frmStartup.Hide
        'show that the game has started
    Case Is = 2
        Call msg("Not Yet Imlimented.", frmStartup)
    Case Is = 3 'the options button
        frmOptions.Visible = True
        frmStartup.Enabled = False
    Case Is = 4 'the quit label
        Call restoreRes
        End
    End Select

End Sub

Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'see if a new label has been selected
    If current <> Index Then
        
        'if it has, then turn the old label white
        lblButton(current).ForeColor = QBColor(15)
        'set the current label
        current = Index
        'make the current label red
        lblButton(Index).ForeColor = QBColor(12)
    End If


End Sub
