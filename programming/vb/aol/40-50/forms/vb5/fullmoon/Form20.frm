VERSION 5.00
Begin VB.Form Form20 
   BorderStyle     =   0  'None
   Caption         =   "Warez Requester"
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   LinkTopic       =   "Form20"
   Picture         =   "Form20.frx":0000
   ScaleHeight     =   1350
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "  _"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Warez Requester"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
Call SendChat("Full Moon Warez Requester")
Timeout 0.5
Call SendChat("Requesting """ + Text2 + """")
Timeout 0.5
Call SendChat("If you have it type""" + Text2 + """")

Timer1.Enabled = True
End Sub

Private Sub Label2_Click()
Call SendChat("Full Moon Warez Requester")
Timeout 0.5
Call SendChat("Request For  """ + Text2 + """")
Timeout 0.5
Timer1.Enabled = False
Call SendChat("Status:Canceled")
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Form20.WindowState = 1
End Sub

Private Sub Timer1_Timer()
Do
Text1 = LastChatLine
Text2.Text = LCase(Text2)
Text1.Text = LCase(Text1)
If Text1 = Text2 Then
   

Call SendChat(" SNFromLastChatLine " + " Can you send " + Text2)
Timeout 1
'
'
SendChat Tea$

End If
DoEvents
Loop
End Sub
