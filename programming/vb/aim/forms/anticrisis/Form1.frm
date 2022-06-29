VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   225
      Left            =   2640
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "IM"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Chat"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Label Label3 
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
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do Until Form1.Top <= -5000
Form1.Top = Trim(str(Int(Form1.Top) - 175))
Loop
Unload Form1
End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub Form_Load()
Call StayOnTop(Form1.hwnd, True)
Timer1.Enabled = True
TimeOut (1)
Call Chat_Send("<font color=black><B>•</B>´¯`·../)<B>Á</B>ñ†ï <B>Ç</B>rîSïS(' ·.·<B>•</B>")
TimeOut (0.2)
Call Chat_Send("<FONT COLOR=BLACK><B>•</B>´¯`·../)<B>V</B><S>ersion</S> 1.0(' ·.·•")
TimeOut (0.2)
Call Chat_Send("<FONT COLOR=BLACK><B>•</B>´¯`·../)  <a href=""http://www.anticrisis.net"">AntiCrisis</a></B>  (' ·.·<B>•</B>")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = "&HFF0000"
Label5.ForeColor = "&HFF0000"
Label4.ForeColor = "&HFF0000"
Label6.ForeColor = "&HFF0000"

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = "&H3333FF"
Label4.ForeColor = "&H3333FF"
Label5.ForeColor = "&H3333FF"
Label6.ForeColor = "&H3333FF"
End Sub

Private Sub Label1_Click()
PopupMenu Form2.File
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = "&H00FF00"
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_Move(Me)
End Sub

Private Sub Label4_Click()
PopupMenu Form2.Chat
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = "&H00FF00"
End Sub

Private Sub Timer2_Timer()
Label2.Caption = Time
End Sub

Private Sub Label5_Click()
PopupMenu Form2.IM
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = "&H00FF00"
End Sub

Private Sub Label6_Click()

PopupMenu Form2.OTHER
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = "&H00FF00"
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Time
End Sub
