VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000007&
   Caption         =   "Mail Punter"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   1245
   ScaleWidth      =   3810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Text            =   $"Form4.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   525
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form4.frx":0102
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "Subject"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Who to Mail Punt"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close Me"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendMail Text1, Text2, Text4
End Sub

Private Sub Command2_Click()
Unload Form4

Form4.Hide

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form4

End Sub
