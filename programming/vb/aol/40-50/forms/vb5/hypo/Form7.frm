VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Anti Log-Off"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2535
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   1605
   ScaleWidth      =   2535
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keeps You Online Longer"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Closes the window that asks you if you want to get offline"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
KeyWord ("aol://2719:2-2-idle")
Timer2.Enabled = True

End Sub

Private Sub Command3_Click()

Form7.Hide

Timer1.Enabled = False
Timer2.Enabled = False
Unload Form7
End Sub

Private Sub Form_Load()
StayOnTop Form7
End Sub

Private Sub Timer1_Timer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Private Sub Timer2_Timer()
Timeout 10
SendChat ("Imagine 98 Idle")
End Sub
