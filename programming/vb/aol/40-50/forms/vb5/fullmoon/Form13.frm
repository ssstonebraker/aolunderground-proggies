VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Voter Bot"
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   4440
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   480
   End
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
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
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
      Left            =   2880
      TabIndex        =   13
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " X"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Voter Bot"
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
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
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
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Question Here"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Vote"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Vote"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Answered No"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Answered Yes"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChat "Type Yes or No"
'
SendChat "And The Question IS.. """ + Text2 + """"

Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer1.Enabled = False
Form19.Hide
Unload Form19




End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label10_Click()
Form13.WindowState = 1
End Sub

Private Sub Label3_Click()
SendChat "Type Yes or No"
'
SendChat "And The Question IS.. """ + Text2 + """"

Timer1.Enabled = True
End Sub

Private Sub Label4_Click()
Timer1.Enabled = False
counterz$ = List2.ListCount
Counter$ = List1.ListCount
SendChat "FullMoon" + counterz$ + " Votes for no"
SendChat "FullMoon" + Counter + " Votes for yes"
Timer1.Enabled = False
Timer1.Enabled = False


End Sub

Private Sub Label9_Click()
Unload Form13
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call RemoveNameFromList(List1)

End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call RemoveNameFromList(List2)
End Sub

Private Sub Timer1_Timer()
Text1 = LastChatLine
Text1.Text = LCase(Text1)
If Text1 = ("no") Then
SendChat SNFromLastChatLine + "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "> Vote for No Counted"
List2.AddItem SNFromLastChatLine
Timeout 0.9
End If

If Text1 = ("yes") Then
SendChat SNFromLastChatLine + "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "> Vote for Yes Counted"
Timeout 0.9
List1.AddItem SNFromLastChatLine
End If
DoEvents
End Sub

Private Sub Timer2_Timer()
Label6.Caption = Str(List1.ListCount)

End Sub

Private Sub Timer3_Timer()
Label7.Caption = Str(List2.ListCount)
End Sub
