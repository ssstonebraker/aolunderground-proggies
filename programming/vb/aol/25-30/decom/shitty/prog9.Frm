VERSION 4.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Redialer Stats"
   ClientHeight    =   975
   ClientLeft      =   1545
   ClientTop       =   1605
   ClientWidth     =   2340
   Height          =   1380
   Left            =   1485
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2340
   ShowInTaskbar   =   0   'False
   Top             =   1260
   Width           =   2460
   Begin VB.CommandButton Command1 
      Caption         =   "STOP!!!"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Tries"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form9"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
c% = FindWindow("Thunderform", "Redialer Stats")
Z = aolhwnd = c%
SetWinOnTop = SetWindowPos(c%, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


