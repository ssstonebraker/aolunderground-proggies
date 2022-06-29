VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About Tetris"
   ClientHeight    =   2010
   ClientLeft      =   3600
   ClientTop       =   3555
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "About.frx":030A
   ScaleHeight     =   2010
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "About..."
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
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This Was Made With VB5,Thank You Livid For Your Help."
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
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "About.frx":167CA
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()

End Sub

Private Sub Form_Load()
Label2.ForeColor = QBColor(15)
'Position the form
frmAbout.Left = (Screen.Width - frmAbout.Width) / 2
frmAbout.Top = (Screen.Height - frmAbout.Height) / 3
'Center the OK command button on the form
Label1.Left = (frmAbout.Width - Label1.Width) / 2
StayOnTop Me
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
'Hide the form
frmAbout.Hide

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = QBColor(12)
End Sub
