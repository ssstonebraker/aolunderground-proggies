VERSION 5.00
Begin VB.Form FRMMain 
   BorderStyle     =   0  'None
   ClientHeight    =   3720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   375
      Left            =   1320
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   1320
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2880
      X2              =   2880
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   840
      X2              =   2880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   840
      X2              =   2880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   840
      X2              =   840
      Y1              =   480
      Y2              =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How do I know what is by F®A//\\//TIC?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F®A//\\//TIC'S Mouse Move Ex."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3750
      Left            =   0
      Picture         =   "FRMMain.frx":0000
      Top             =   0
      Width           =   3750
   End
End
Attribute VB_Name = "FRMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call SetWindowPos(FRMMAIN.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "How do I know what is by F®A//\\//TIC?"
Label3.Caption = ""
Label2.ForeColor = &H0
Label4.Caption = "Time?"
Label4.ForeColor = &H0
Label5.Caption = ""
Label6.Caption = "Date?"
Label6.ForeColor = &H0
Label7.Caption = ""
Label1.ForeColor = &H0&
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&
Label3.Caption = "Thanks for downloading my example and for help or information e-mail F®A//\\//TIC at Frantic554@aol.com."
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFF&
Label3.Caption = "The program will have F®A//\\//TIC spelled like that somewhere on the form usually with faded colors."

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFF&
Label5.Caption = Time
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFF&
Label7.Caption = Date
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
    Call SendMessage(FRMMAIN.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
