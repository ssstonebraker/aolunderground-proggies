VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2430
   LinkTopic       =   "Form8"
   ScaleHeight     =   1650
   ScaleWidth      =   2430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Send"
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
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
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
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form8.Top <= -5000
Form8.Top = Trim(str(Int(Form8.Top) - 175))
Loop
Unload Form8
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"

End Sub

Private Sub Form_Load()
Call StayOnTop(Form8.hwnd, True)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"

End Sub

Private Sub Label2_Click()
Call Chat_InvitetoRoom(Text1, "Come to my room")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"

End Sub
