VERSION 5.00
Object = "{92EDEF56-A415-11D2-BBA6-BA26EE701995}#4.0#0"; "QUIRKAIM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Izekials AiM Phader 1.o"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "My Page"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Mail Me"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "iM Me"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "Who..."
      Top             =   720
      Width           =   1455
   End
   Begin QuiRKAIM.AIM AIM1 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   1920
      Top             =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "X"
      Height          =   285
      Left            =   3960
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¯"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox Check3 
      Caption         =   "WavY"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Advertise"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Open iM"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Chat Room"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "­"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Type Here..."
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click()
 Form1.Height = 1350
End Sub

Private Sub Command1_Click()
If Check1.Value = 0 And Check2.Value = 0 Then MsgBox "Hit The Down Arrow And Choose Either Chat Or IM", vbOKOnly, "Error"
If Check2.Value = 1 Then
    sString$ = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLUE, Text1, Check3.Value)
    Call AIM1.IM(Text2, sString$)
End If
If Check1.Value = 1 Then
    sString$ = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLUE, Text1, Check3.Value)
    AIM1.ChatSend (sString$)
End If
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Form1.Height = 675
End Sub

Private Sub Command3_Click()
Ad1$ = FadeByColor3(FADE_GREEN, FADE_BLUE, FADE_GREEN, "--AiM Phader", False)
Ad2$ = FadeByColor3(FADE_GREEN, FADE_BLUE, FADE_GREEN, "--xIzekial83", False)
    AIM1.ChatSend (Ad1$)
    TimeOut 0.5
    AIM1.ChatSend (Ad2$)
End Sub

Private Sub Command4_Click()
Form1.Height = 1050
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Call AIM1.IM("izekial83", "Sup")
End Sub

Private Sub Command7_Click()
Call Mail_Send("Funkdemon@yahoo.com", "AiM Phader", "I Am Testing It")
End Sub

Private Sub Command8_Click()
Shell ("http://members.xoom.com/izekial83/")
End Sub

Private Sub Form_Load()
Ad1$ = FadeByColor3(FADE_GREEN, FADE_BLUE, FADE_GREEN, "--AiM Phader", False)
Ad2$ = FadeByColor3(FADE_GREEN, FADE_BLUE, FADE_GREEN, "--xIzekial83", False)
    AIM1.ChatSend (Ad1$)
    TimeOut 0.5
    AIM1.ChatSend (Ad2$)
StayOnTop Me
Call TopLeftForm(Form1)
Form1.Height = 675
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub

Private Sub Timer1_Timer()
If Check1.Value = 1 Then Check2.Value = 0
If Check2.Value = 1 Then Check1.Value = 0
End Sub
