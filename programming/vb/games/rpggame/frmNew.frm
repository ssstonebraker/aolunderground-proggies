VERSION 5.00
Begin VB.Form frmNew 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picChar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5325
      Index           =   1
      Left            =   5520
      Picture         =   "frmNew.frx":0000
      ScaleHeight     =   5325
      ScaleWidth      =   4080
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   4080
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Home"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin VB.PictureBox picChar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5190
      Index           =   2
      Left            =   5280
      Picture         =   "frmNew.frx":9B5E
      ScaleHeight     =   5190
      ScaleWidth      =   4500
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.PictureBox picChar 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5670
      Index           =   3
      Left            =   5160
      Picture         =   "frmNew.frx":138A3
      ScaleHeight     =   5670
      ScaleWidth      =   4740
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mage"
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
      Index           =   1
      Left            =   4560
      TabIndex        =   27
      Top             =   600
      Width           =   705
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barbarian"
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
      Index           =   3
      Left            =   7200
      TabIndex        =   32
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label lblChar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warrior"
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
      Index           =   2
      Left            =   5760
      TabIndex        =   28
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a Class:"
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
      Index           =   2
      Left            =   4560
      TabIndex        =   26
      Top             =   240
      Width           =   1905
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   3000
      TabIndex        =   25
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label lblButton 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Home"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   6480
      Width           =   1050
   End
   Begin VB.Label lblDecrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
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
      Index           =   4
      Left            =   2160
      TabIndex        =   23
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label lblDecrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
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
      Index           =   3
      Left            =   2160
      TabIndex        =   22
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label lblDecrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
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
      Index           =   2
      Left            =   2160
      TabIndex        =   21
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label lblDecrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
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
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label lblIncrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
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
      Index           =   4
      Left            =   3120
      TabIndex        =   19
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Index           =   4
      Left            =   2520
      TabIndex        =   18
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label lblIncrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   17
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Index           =   3
      Left            =   2520
      TabIndex        =   16
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label lblIncrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   15
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   14
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label lblIncrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
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
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label lblDecrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
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
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lblIncrease 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vitality:"
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
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Label lblVal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "50"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligence:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1545
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stamina:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1125
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Character Points Remaining:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   4095
      Width           =   3615
   End
   Begin VB.Label lblPtsRemain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   4095
      Width           =   495
   End
   Begin VB.Label lblStatic 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter your Hero's Name:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3195
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gotClass As Boolean

Private Sub Form_Load()
    Character.itsClass = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If gotClass = False Then
        'make all the labels white
        lblChar(MAGE).ForeColor = QBColor(15)
        lblChar(WARRIOR).ForeColor = QBColor(15)
        lblChar(BARBARIAN).ForeColor = QBColor(15)
    ElseIf gotClass = True Then
        'determine which of the character labels to make red
        If Character.itsClass = MAGE Then
            lblChar(MAGE).ForeColor = QBColor(12)
            lblChar(WARRIOR).ForeColor = QBColor(15)
            lblChar(BARBARIAN).ForeColor = QBColor(15)
        ElseIf Character.itsClass = WARRIOR Then
            lblChar(MAGE).ForeColor = QBColor(15)
            lblChar(WARRIOR).ForeColor = QBColor(12)
            lblChar(BARBARIAN).ForeColor = QBColor(15)
        ElseIf Character.itsClass = BARBARIAN Then
            lblChar(MAGE).ForeColor = QBColor(15)
            lblChar(WARRIOR).ForeColor = QBColor(15)
            lblChar(BARBARIAN).ForeColor = QBColor(12)
        End If
    End If
    
    'make all the labels for increasing stats etc white
    For X = 0 To 4
        lblIncrease(X).ForeColor = QBColor(15)
        lblDecrease(X).ForeColor = QBColor(15)
        lblStat(X).ForeColor = QBColor(15)
    Next X
    
    'make the two buttons white
    lblButton(0).ForeColor = QBColor(15)
    lblButton(1).ForeColor = QBColor(15)
End Sub

Private Sub lblButton_Click(Index As Integer)

    If Index = CANCEL Then 'the cancel button
        
        'play the button sound
        Call sndPlaySound(sndButton, &H1)
        frmNew.Hide 'hide the New Game menu
        frmStartup.Show 'show the startup menu
    
    ElseIf Index = OK Then  'the proceed button
    
        'check to see that the character's name is valid (no numbers)
        For i = 1 To Len(txtName.Text)
            If IsNumeric(Mid(txtName.Text, i, 1)) = True Then
                Call msg("You must name your character.", frmNew)
                Exit Sub
            End If
        Next i
        
        If txtName.Text = "" Then   'if no name has been entered
            Call msg("You must name your character.", frmNew)
            
        ElseIf lblPtsRemain.Caption <> 0 Then 'if there's points remaining
            Call msg("You must use all your character points.", frmNew)
            
        Else    'if everything's ok
            'indicate that there's a game in progress
            GameInProgress = True
            frmDisplay.Show
            frmNew.Hide
            
        End If
    End If

End Sub

'make the current label red
Private Sub lblButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblButton(Index).ForeColor = QBColor(12)
End Sub

'the user clicks on a character class
Private Sub lblChar_Click(Index As Integer)

    gotClass = True
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'make both labels white and hide both pictures
    For X = 1 To 3
        lblChar(X).ForeColor = QBColor(15)
        picChar(X).Visible = False
    Next X
    
    'show the proper picture
    picChar(Index).Visible = True
    lblChar(Index).ForeColor = QBColor(12)
    
    
    'set the character's base stats based on the class
    If Index = MAGE Then
        'set the class
        Character.itsClass = MAGE
        'set the base stats
        Call Character.setBaseStat(STR, 40)
        Call Character.setBaseStat(STA, 50)
        Call Character.setBaseStat(MAG, 80)
        Call Character.setBaseStat(INTEL, 70)
        Call Character.setBaseStat(VITAL, 50)
    ElseIf Index = WARRIOR Then
        'set the class
        Character.itsClass = WARRIOR
        'set the base stats
        Call Character.setBaseStat(STR, 65)
        Call Character.setBaseStat(STA, 70)
        Call Character.setBaseStat(MAG, 20)
        Call Character.setBaseStat(INTEL, 55)
        Call Character.setBaseStat(VITAL, 70)
    ElseIf Index = BARBARIAN Then
        'set the class
        Character.itsClass = BARBARIAN
        'set the base stats
        Call Character.setBaseStat(STR, 85)
        Call Character.setBaseStat(STA, 75)
        Call Character.setBaseStat(MAG, 5)
        Call Character.setBaseStat(INTEL, 45)
        Call Character.setBaseStat(VITAL, 75)
    End If
    
    Dim i As Integer
    For i = 0 To 4
        lblVal(i).Caption = Character.getBaseStat(i)
    Next i
    
    lblPtsRemain.Caption = 100

End Sub

Private Sub lblChar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'make the label's forecolor red
    lblChar(Index).ForeColor = QBColor(12)
End Sub

Private Sub lblDecrease_Click(Index As Integer)
    
    'play the sound effect
    Call sndPlaySound(sndButton, &H1)
        
    'see if the user has selected a class yet
    If gotClass = True Then
        'decrease the stat
        If lblVal(Index).Caption > Character.getBaseStat(Index) Then
            lblVal(Index).Caption = lblVal(Index).Caption - 5
            lblPtsRemain.Caption = lblPtsRemain.Caption + 5
        End If
    Else
        'show a message indicating that a class has not been shown
        Call msg("You must first choose a Class", frmNew)
    End If
    
End Sub

Private Sub lblDecrease_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'make all the labels white
    For X = 0 To 4
        lblIncrease(X).ForeColor = QBColor(15)
        lblDecrease(X).ForeColor = QBColor(15)
        lblStat(X).ForeColor = QBColor(15)
    Next X

    'make the currently selected label red
    lblStat(Index).ForeColor = QBColor(12)
    lblDecrease(Index).ForeColor = QBColor(12)
    
End Sub

Private Sub lblIncrease_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'set all the labels color's to white
    For X = 0 To 4
        lblIncrease(X).ForeColor = QBColor(15)
        lblDecrease(X).ForeColor = QBColor(15)
        lblStat(X).ForeColor = QBColor(15)
    Next X
    
    'make the selected stat red
    lblIncrease(Index).ForeColor = QBColor(12)
    lblStat(Index).ForeColor = QBColor(12)
End Sub

Private Sub lblIncrease_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'see if the user has selected a class yet
    If gotClass = True Then
        'if so, increase the amount of points for this attribute
        If lblPtsRemain > 0 And lblVal(Index).Caption < 100 Then
            lblVal(Index).Caption = lblVal(Index).Caption + 5
            lblPtsRemain.Caption = lblPtsRemain.Caption - 5
        End If
    Else
        'show a message indicating that a class has not been shown
        Call msg("You must first choose a Class.", frmNew)
    End If
    
End Sub
