VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   Caption         =   "Punter"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   ScaleHeight     =   1860
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   720
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "100"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   3
      Text            =   "SteveCase"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "Punt Em"
      Height          =   495
      Left            =   0
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
 
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False

End Sub

Private Sub Command3_Click()
Unload Form5

Form5.Hide

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form5
End Sub

Private Sub Timer1_Timer()
Call IMKeyword(Text1, "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><h3>ToaST says BYE BYE")
TimeOut 0.5
Text2.Text = Val(Text2.Text) - 1
If Text2.Text = 0 Then
Timer1.Enabled = False
End If

End Sub
