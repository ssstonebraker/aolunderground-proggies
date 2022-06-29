VERSION 5.00
Begin VB.Form frmInstruct 
   BorderStyle     =   0  'None
   Caption         =   "Instructions"
   ClientHeight    =   2985
   ClientLeft      =   1155
   ClientTop       =   2415
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "Instruct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Instruct.frx":030A
   ScaleHeight     =   2985
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To pause and unpause the game, press P or us the Pause menu."
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
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1830
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To make the pieces fall right to the bottom, press the down key."
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
      TabIndex        =   3
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To rotate the pieces, press either the up key or 5 on the numpad."
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To move the pieces left and right, use the left and right arrow keys.  Make sure numlock is off if you are using the numpad)"
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tetris is very simple to play."
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "Instruct.frx":3202E
      Top             =   2160
      Width           =   480
   End
End
Attribute VB_Name = "frmInstruct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()

End Sub


Private Sub Form_Load()
Label2.ForeColor = QBColor(15)
StayOnTop Me
'Position the form
frmInstruct.Left = (Screen.Width - frmInstruct.Width) / 2
frmInstruct.Top = (Screen.Height - frmInstruct.Height) / 2
'Center the OK command button on the form
Label1.Left = (frmInstruct.Width - Label1.Width) / 2

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label6_Click()
'Hide the form
frmInstruct.Hide

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

)
End Sub
