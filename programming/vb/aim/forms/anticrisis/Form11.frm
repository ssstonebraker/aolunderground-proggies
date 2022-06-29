VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form11"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LinkTopic       =   "Form11"
   ScaleHeight     =   1440
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2415
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2160
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
      Width           =   2175
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form11.Top <= -5000
Form11.Top = Trim(Str(Int(Form11.Top) - 175))
Loop
Unload Form11
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H3333FF"
End Sub

Private Sub Form_Load()
Call StayOnTop(Form11.hwnd, True)
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
Call Chat_ScrollUserInfo(Text1)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = "&H00FF00"
End Sub

Private Sub Text1_Change()
Label2.ForeColor = "&H3333FF"
End Sub
