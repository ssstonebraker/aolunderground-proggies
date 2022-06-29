VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   LinkTopic       =   "Form7"
   ScaleHeight     =   1245
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   1935
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "What 2 Change it 2"
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Áñ†ï ÇrîSïS"
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form7.Top <= -5000
Form7.Top = Trim(str(Int(Form7.Top) - 175))
Loop
Unload Form7
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Form_Load()
Call StayOnTop(Form7.hwnd, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Label2_Click()
Call Chat_Caption(Text1)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub
