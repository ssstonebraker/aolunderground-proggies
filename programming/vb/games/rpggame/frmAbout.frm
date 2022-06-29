VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Matthew Eagar"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   2325
   End
   Begin VB.Image imgTitle 
      Height          =   1230
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Top             =   0
      Width           =   2685
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic RPG"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   2205
   End
   Begin VB.Label lblButton 
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
      Left            =   4800
      TabIndex        =   0
      Top             =   3840
      Width           =   1635
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'put the current version into the version label
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'make the label white
    lblButton.ForeColor = QBColor(15)

End Sub


Private Sub lblButton_Click()
    'play the button sound
    Call sndPlaySound(sndButton, &H1)
    
    'hide this form and re-inable the main form
    frmAbout.Hide
    frmStartup.Enabled = True
    frmStartup.SetFocus
End Sub

Private Sub lblButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'make the lable red
    lblButton.ForeColor = QBColor(12)
End Sub
