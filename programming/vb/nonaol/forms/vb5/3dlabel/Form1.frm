VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line16 
      X1              =   1080
      X2              =   1080
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line15 
      X1              =   480
      X2              =   1080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line14 
      BorderColor     =   &H80000009&
      X1              =   480
      X2              =   480
      Y1              =   2280
      Y2              =   2760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H80000009&
      X1              =   480
      X2              =   1080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      Caption         =   "EXIT"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      X1              =   480
      X2              =   480
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      X1              =   480
      X2              =   1080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000001&
      X1              =   1080
      X2              =   1080
      Y1              =   720
      Y2              =   240
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   480
      X2              =   1080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Help"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   1440
      X2              =   1440
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   120
      X2              =   120
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   120
      X2              =   1440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   1440
      X2              =   1440
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   120
      X2              =   1440
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   120
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1440
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Line13.Visible = False
Line14.Visible = False
Line15.Visible = False
Line16.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Line13.Visible = False
Line14.Visible = False
Line15.Visible = False
Line16.Visible = False
End Sub

Private Sub Label1_Click()

Line2.Visible = False
Line1.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = True
Line6.Visible = True
Line7.Visible = True
Line8.Visible = True

StartTime = Timer
Do While Timer - StartTime < 0.1
DoEvents
Loop

Line2.Visible = True
Line1.Visible = True
Line3.Visible = True
Line4.Visible = True
StartTime = Timer
Do While Timer - StartTime < 0.1
DoEvents
Loop

Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
End Sub

Private Sub Label2_Click()
MsgBox "Click The thing down there to see what the 3d box around a label looks like!  Also not that ther IS A LABEL!", , "Help"
MsgBox "U can use the code aslong as u don't say u made it", , "I made it"
MsgBox "Code By XBrôS", , "Me"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line9.Visible = True
Line10.Visible = True
Line11.Visible = True
Line12.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line13.Visible = True
Line14.Visible = True
Line15.Visible = True
Line16.Visible = True
MsgBox "See ya!", , "See ya!"
Unload Me
End Sub
