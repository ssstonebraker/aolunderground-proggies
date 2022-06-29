VERSION 5.00
Begin VB.Form Mesages 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   Picture         =   "Mesages.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   315
      MaxLength       =   20
      TabIndex        =   1
      Top             =   990
      Width           =   2625
   End
   Begin VB.TextBox Text1 
      Height          =   600
      Left            =   315
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   270
      Width           =   2625
   End
   Begin VB.Label startbut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   720
      MouseIcon       =   "Mesages.frx":297B
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1305
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   1890
      MouseIcon       =   "Mesages.frx":2ACD
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1305
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The ^ will display amount finished"
      ForeColor       =   &H80000018&
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   0
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   0
      Picture         =   "Mesages.frx":2C1F
      Top             =   0
      Width           =   3240
   End
End
Attribute VB_Name = "Mesages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
FormOnTop Me    'sets the form to the top most
Text1 = MailMsg 'sets text1 the the mail message
Text2 = SentMsg 'sets text2 to the sent message
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me 'enables the user to move the form by dragging it
End Sub
Private Sub Label1_Click()
Me.Hide 'hides the form
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me 'see above
End Sub
Private Sub startbut_Click()
MailMsg = Text1 'sets mail message to the current text1
SentMsg = Text2 'sets sent message to the current text2
Me.Hide 'hides the form
End Sub

