VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form12"
   ScaleHeight     =   1605
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KiLLz"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Áñ†ï ÇrîSïS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form12.Top <= -5000
Form12.Top = Trim(Str(Int(Form12.Top) - 175))
Loop
Unload Form12
End Sub

Private Sub Form_Load()
Call StayOnTop(Form12.hwnd, True)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"

End Sub

Private Sub Label2_Click()
PopupMenu Form2.SCROLL
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Label3_Click()
PopupMenu Form2.KILLZ
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
Label3.ForeColor = "&H3333FF"

End Sub
