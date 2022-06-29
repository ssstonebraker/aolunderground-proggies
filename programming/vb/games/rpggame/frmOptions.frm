VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblReturn 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Return"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   2085
      Width           =   1635
   End
   Begin VB.Label lblDecreaseGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Decrease"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label lblIncreaseGame 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Increase"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lblNew 
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1830
   End
   Begin VB.Label lblStatic 
      BackColor       =   &H00000000&
      Caption         =   "Game Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Shape shpBarG 
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   2400
      Top             =   960
      Width           =   1815
   End
   Begin VB.Shape shpSlideG 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   375
      Left            =   4080
      Top             =   960
      Width           =   135
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'make all the labels white
lblIncreaseWalk.ForeColor = QBColor(15)
lblIncreaseGame.ForeColor = QBColor(15)
lblDecreaseWalk.ForeColor = QBColor(15)
lblDecreaseGame.ForeColor = QBColor(15)
lblReturn.ForeColor = QBColor(15)

End Sub


'decrease game speed
Private Sub lblDecreaseGame_Click()
    
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'move the sliding bar to the right
    shpSlideG.Left = shpSlideG.Left - 30
    
    'see if the slider is at min
    If shpSlideG.Left > shpBarG.Left Then
        'if not, make the wait time larger (longer wait = decreased game speed)
        wait = wait + 0.005
    Else
        'if it is, make the slider equal to the max value, in case it went over
        shpSlideG.Left = shpBarG.Left
    End If
    
    shpBarG.Refresh
End Sub

Private Sub lblDecreaseGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDecreaseGame.ForeColor = QBColor(12)
End Sub

'decrease the walking speed
Private Sub lblDecreaseWalk_Click()
    
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'move the sliding bar to the right
    shpSlideW.Left = shpSlideW.Left - 60
    
    'see if the slider is at min
    If shpSlideW.Left > shpBarW.Left Then
        'if not, make the wait time larger (longer wait = decreased game speed)
        Speed = Speed - 1
    Else
        'if it is, make the slider equal to the max value, in case it went over
        shpSlideW.Left = shpBarW.Left
    End If

    shpBarW.Refresh
End Sub

Private Sub lblDecreaseWalk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDecreaseWalk.ForeColor = QBColor(12)
End Sub

'increase game speed
Private Sub lblIncreaseGame_Click()

    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'move the sliding bar to the right
    shpSlideG.Left = shpSlideG.Left + 30
    'see if the slider is at max
    If shpSlideG.Left + shpSlideG.Width <= shpBarG.Left + shpBarG.Width Then
        'if not, make the waiting time smaller (shorter wait, increased game speed)
        wait = wait - 0.005
        If wait < 0 Then wait = 0
    Else
        'if it is, make the slider equal to the max value, in case it went over
        shpSlideG.Left = shpBarG.Left + shpBarG.Width - shpSlideG.Width
        wait = 0
    End If

    shpBarG.Refresh

End Sub

Private Sub lblIncreaseGame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblIncreaseGame.ForeColor = QBColor(12)
End Sub

'increase walking speed
Private Sub lblIncreaseWalk_Click()
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    'move the sliding bar to the right
    shpSlideW.Left = shpSlideW.Left + 60
    'see if the slider is at max
    If shpSlideW.Left + shpSlideW.Width < shpBarW.Left + shpBarW.Width Then
        'if not, increase the speed
        Speed = Speed + 1
    Else
        'if it is, make the slider equal to the max value, in case it went over
        shpSlideW.Left = shpBarW.Left + shpBarW.Width - shpSlideW.Width
    End If

    shpBarW.Refresh
End Sub

Private Sub lblIncreaseWalk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblIncreaseWalk.ForeColor = QBColor(12)
End Sub

Private Sub lblReturn_Click()

    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    'show the main window, and hide the options window
    frmOptions.Visible = False
    frmStartup.Enabled = True
    frmStartup.SetFocus

End Sub

Private Sub lblReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblReturn.ForeColor = QBColor(12)
End Sub
