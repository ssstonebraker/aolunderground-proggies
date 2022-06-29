VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   LinkTopic       =   "Form3"
   ScaleHeight     =   495
   ScaleWidth      =   1350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "   Áñ†ï ÇrîSïS"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form3.Hide

End Sub

Private Sub Form_Load()
Call StayOnTop(Form3.hwnd, True)
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Timer1_Timer()
Command1.Caption = Time
End Sub
