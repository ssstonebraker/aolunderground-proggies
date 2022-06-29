VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Animations Example                                By $i|\|\|eR"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   70
      Left            =   4320
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2160
      Top             =   1320
   End
   Begin VB.CommandButton title 
      Caption         =   "Click HeRe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "       BlinK Example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Line Line16 
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   360
   End
   Begin VB.Line Line15 
      X1              =   3240
      X2              =   3240
      Y1              =   360
      Y2              =   720
   End
   Begin VB.Line Line14 
      X1              =   6120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line13 
      X1              =   6120
      X2              =   3240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line12 
      X1              =   6120
      X2              =   6120
      Y1              =   3600
      Y2              =   720
   End
   Begin VB.Line Line11 
      X1              =   5280
      X2              =   6120
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line10 
      X1              =   3240
      X2              =   5880
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line9 
      X1              =   3240
      X2              =   3240
      Y1              =   720
      Y2              =   3600
   End
   Begin VB.Image search_pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   4560
      Picture         =   "frmMain.frx":030A
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "     Motion Example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line8 
      X1              =   3000
      X2              =   3000
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line6 
      X1              =   3000
      X2              =   240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      X1              =   3000
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   3000
      X2              =   3000
      Y1              =   3360
      Y2              =   480
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   3000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   2760
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   3360
   End
   Begin VB.Image aol9 
      Height          =   480
      Left            =   2280
      Picture         =   "frmMain.frx":0921
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image aol8 
      Height          =   480
      Left            =   2040
      Picture         =   "frmMain.frx":11EB
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image aol7 
      Height          =   480
      Left            =   1800
      Picture         =   "frmMain.frx":1AB5
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image aol6 
      Height          =   480
      Left            =   1560
      Picture         =   "frmMain.frx":237F
      Top             =   1920
      Width           =   480
   End
   Begin VB.Image aol5 
      Height          =   480
      Left            =   1320
      Picture         =   "frmMain.frx":2C49
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image aol3 
      Height          =   480
      Left            =   840
      Picture         =   "frmMain.frx":3513
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image aol2 
      Height          =   480
      Left            =   600
      Picture         =   "frmMain.frx":3DDD
      Top             =   960
      Width           =   480
   End
   Begin VB.Image aol4 
      Height          =   480
      Left            =   1080
      Picture         =   "frmMain.frx":46A7
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image aol1 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":4F71
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
aol2.Visible = False
aol3.Visible = False
aol4.Visible = False
aol5.Visible = False
aol6.Visible = False
aol7.Visible = False
aol8.Visible = False
aol9.Visible = False
End Sub

Private Sub Timer1_Timer()
aol1.Visible = False
Pause 0.1
aol2.Visible = True
Pause 0.1
aol2.Visible = False
Pause 0.1
aol3.Visible = True
Pause 0.1
aol3.Visible = False
Pause 0.1
aol4.Visible = True
Pause 0.1
aol4.Visible = False
Pause 0.1
aol5.Visible = True
Pause 0.1
aol5.Visible = False
Pause 0.1
aol6.Visible = True
Pause 0.1
aol6.Visible = False
Pause 0.1
aol7.Visible = True
Pause 0.1
aol7.Visible = False
Pause 0.1
aol8.Visible = True
Pause 0.1
aol8.Visible = False
Pause 0.1
aol9.Visible = True
Pause 0.1
aol9.Visible = False
Pause 0.1
aol1.Visible = True

End Sub

Private Sub Timer2_Timer()
search_pic.Visible = False
Pause 0.1
search_pic.Visible = True
End Sub

Private Sub title_Click()
Me.Caption = "A"
Pause 0.2
Me.Caption = "An"
Pause 0.2
Me.Caption = "Ani"
Pause 0.2
Me.Caption = "Anim"
Pause 0.2
Me.Caption = "Anima"
Pause 0.2
Me.Caption = "Animat"
Pause 0.2
Me.Caption = "Animati"
Pause 0.2
Me.Caption = "Animatio"
Pause 0.2
Me.Caption = "Animation"
Pause 0.2
Me.Caption = "Animation E"
Pause 0.2
Me.Caption = "Animation Ex"
Pause 0.2
Me.Caption = "Animation Exa"
Pause 0.2
Me.Caption = "Animation Exam"
Pause 0.2
Me.Caption = "Animation Examp"
Pause 0.2
Me.Caption = "Animation Exampl"
Pause 0.2
Me.Caption = "Animation Example"
End Sub
