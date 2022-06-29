VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3660
   ClientLeft      =   3060
   ClientTop       =   3525
   ClientWidth     =   5310
   ForeColor       =   &H00000000&
   Height          =   4065
   Icon            =   "Form1.frx":0000
   Left            =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Top             =   3180
   Width           =   5430
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   720
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1035
      ItemData        =   "Form1.frx":0442
      Left            =   120
      List            =   "Form1.frx":0444
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "Person you want to ignore."
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":0446
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "stop"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "start"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "im answer message"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   1200
      X2              =   1320
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1440
      X2              =   1320
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      DrawMode        =   6  'Mask Pen Not
      Index           =   1
      X1              =   1320
      X2              =   1320
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   3600
      X2              =   3720
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   3840
      X2              =   3720
      Y1              =   840
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   2  'Dash
      DrawMode        =   6  'Mask Pen Not
      Index           =   0
      X1              =   3720
      X2              =   3720
      Y1              =   720
      Y2              =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "ignore message"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " _"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "- turk'z im answerer\ignorer - "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Call FormOnTop(Form1)
Timer1.Enabled = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub


Private Sub Label1_Click()
End
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub


Private Sub Label3_Click()
Form1.WindowState = 1
End Sub


Private Sub Label6_Click()
Timer1.Enabled = True
End Sub

Private Sub Label7_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()


Dim Text As String
If FindIM <> 0 Then
GoTo Hello:

Hello:
If IMSender = Text2 Then
Call InstantMessage(IMSender, Text3)
TimeOut 1
Call CloseWindow(FindIM&)
Else: GoTo Programmers:

Programmers:
List1.AddItem IMSender & ": " & IMLastMsg
Call InstantMessage(IMSender, Text1)
TimeOut 1
Call CloseWindow(FindIM&)
End If
End If
End Sub

