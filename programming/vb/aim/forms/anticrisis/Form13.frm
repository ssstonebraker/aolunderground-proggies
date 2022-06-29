VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form13"
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   LinkTopic       =   "Form13"
   ScaleHeight     =   1635
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   600
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2400
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
      Width           =   2415
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form13.Top <= -5000
Form13.Top = Trim(Str(Int(Form13.Top) - 175))
Loop
Unload Form13
End Sub

Private Sub Form_Load()
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
Timer1.Enabled = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
End Sub

Private Sub Timer1_Timer()
Call Chat_QuickRoom2(Text1)
TimeOut (3)
Call Chat_Send("<font color=black><B>•</B>´¯`·../)<B>Á</B>ñ†ï <B>Ç</B>rîSïS  <B>Q</B>uick <B>R</B>oom")
Unload Form13
End Sub
