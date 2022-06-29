VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Mouse Move Example"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Form"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   0  'Manual
      Class           =   "SoundRec"
      Height          =   375
      Left            =   -360
      SourceDoc       =   "C:\My Documents\KnK-buttonsvolum1\Lclick.wav"
      TabIndex        =   7
      Top             =   -120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"mouse move.frx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   2175
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "mouse move.frx":00AE
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = ""
Label4.Visible = False
Label7.Caption = ""
End Sub





Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = Time
End Sub

Private Sub Label5_Click()
OLE1.DoVerb
WindowState = 1
End Sub

Private Sub Label6_Click()
OLE1.DoVerb
End
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Visible = True
End Sub



Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.Caption = "OK"

End Sub
