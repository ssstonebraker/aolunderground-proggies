VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Punter"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   $"Form9.frx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "100"
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "Who to Punt"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Punt"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
Playwav App.Path + "\punt.toast"
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False

End Sub

Private Sub Command3_Click()
Unload Form9
Form9.Hide

End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form9
End Sub

Private Sub Timer1_Timer()
Call IMKeyword(Text1, Text3)
Text2.Text = Val(Text2.Text) - 1
If Text2.Text = 0 Then
Timer1.Enabled = False
End If

End Sub

