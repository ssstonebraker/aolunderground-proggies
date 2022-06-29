VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   LinkTopic       =   "Form10"
   ScaleHeight     =   1530
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2895
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GO"
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
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   195
      Left            =   2640
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
      Width           =   2655
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form10.Top <= -5000
Form10.Top = Trim(Str(Int(Form10.Top) - 175))
Loop
Unload Form10
End Sub

Private Sub Form_Load()
Call StayOnTop(Form10.hwnd, True)
Label2.ForeColor = "&H3333FF"
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

Call Chat_QuickRoom2(Text1)
TimeOut (3)
Call Chat_Send("<font color=black><B>•</B>´¯`·../)<B>Á</B>ñ†ï <B>Ç</B>rîSïS(' ·.·<B>•</B>")
Call Chat_Send("<font color=black><B>•</B>´¯`·../)<B>Q</B><S>uick</S> <B>R</B><S>oom</S>(' ·.·<B>•</B>")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
End Sub
