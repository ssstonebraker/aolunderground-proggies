VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H80000007&
   Caption         =   "Fake Room"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form14"
   ScaleHeight     =   975
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fake Room"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Steve Case 4 gay Males"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form14.Hide
Unload Form14


End Sub

Private Sub Command2_Click()


If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If
SendChat "                                                                  "
TimeOut 0.002
SendChat "*** You are in """ + Text1 + """. ***"
TimeOut 0.002
SendChat "                                                                  "
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Timer1_Timer()


End Sub

Private Sub Form_Activate()
FadeFormPurple Me
End Sub

Private Sub Form_Load()
StayOnTop Form14
End Sub
