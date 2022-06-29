VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   LinkTopic       =   "Form4"
   ScaleHeight     =   1635
   ScaleWidth      =   1650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Who 2 Ignore"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ignore Them"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Ignore0r"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form4.Top <= -5000
Form4.Top = Trim(str(Int(Form4.Top) - 175))
Loop
Unload Form4
End Sub

Private Sub Form_Load()
Call StayOnTop(Form4.hwnd, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&HFF0000"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label2_Click()
Call chat_ignore(Text1)
Call Chat_Send("<font color=black><B>•</B>´¯`·../)" + Text1.text + " <B>W</B>as [<B>X</B>]ed(' ·.·<B>•</B>")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub
