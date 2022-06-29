VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   LinkTopic       =   "Form5"
   ScaleHeight     =   1500
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton Command2 
         Caption         =   "Options"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "http://"
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chat Linker"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form5.Top <= -5000
Form5.Top = Trim(str(Int(Form5.Top) - 175))
Loop
Unload Form5
End Sub

Private Sub Command2_Click()
PopupMenu Form6.Options
End Sub

Private Sub Form_Load()
Call StayOnTop(Form5.hwnd, True)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub
