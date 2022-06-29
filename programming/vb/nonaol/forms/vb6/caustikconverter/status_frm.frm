VERSION 5.00
Begin VB.Form status_frm 
   BorderStyle     =   0  'None
   Caption         =   "Generating Preview..."
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox stat1 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox stat2 
      BackColor       =   &H00808080&
      Height          =   255
      Left            =   720
      ScaleHeight     =   13
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   720
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   2160
      X2              =   2400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   600
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   2400
      X2              =   2400
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   2400
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color Conversion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   2400
      X2              =   2400
      Y1              =   960
      Y2              =   480
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   2040
      X2              =   2400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2400
      X2              =   600
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   600
      Y1              =   960
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Creation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generating Preview.."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   1530
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   3000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "status_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = mshop_frm.Image2.Picture
Image2.Picture = mshop_frm.Image1.Picture
StayOnTop Me
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
scanning = 0
Me.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Showactive Label4
End Sub
