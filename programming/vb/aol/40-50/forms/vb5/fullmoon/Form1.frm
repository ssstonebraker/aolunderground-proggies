VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "."
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1875
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Kill Wait"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Buster"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "IMs OFF"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IMs ON"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Full  Moon"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " _"
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
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Other"
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
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IM Stuff"
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
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat"
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
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   3360
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll Up"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call SendChat("<font face=""comic sans ms""><FONT COLOR=#fffffe>{S IM}")
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()

StayOnTop Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = QBColor(15)
Label4.ForeColor = QBColor(15)
Label5.ForeColor = QBColor(15)
Label6.ForeColor = QBColor(15)
Label9.ForeColor = QBColor(15)
Label1.ForeColor = QBColor(15)
Label7.ForeColor = QBColor(15)
Label3.ForeColor = QBColor(15)
End Sub

Private Sub Label1_Click()
PlayWAV ("C:/DOLAN")
If Label1.Caption = "Roll Up" Then
Form1.Height = 645
Label1.Caption = " Roll Down"
Else
Label1.Caption = "Roll Up"
Form1.Height = 1894
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = QBColor(12)
End Sub

Private Sub Label11_Click()
Form17.Show
End Sub

Private Sub Label12_Click()
Form22.Show
Unload Form1
End Sub

Private Sub Label13_Click()
Unload Me
End Sub

Private Sub Label14_Click()
Call AOLKillWait
End Sub

Private Sub Label2_Click()
PlayWAV ("C:/DOLAN")
Form2.PopupMenu Form2.Label, 1
'This one uses a 1 since its the first pull down menue on Form2
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = QBColor(12)
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = QBColor(12)
End Sub

Private Sub Label4_Click()
PlayWAV ("C:/DOLAN")
Form2.PopupMenu Form2.Label1, 2
'This one uses a 1 since its the first pull down menue on Form2
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = QBColor(12)
End Sub

Private Sub Label5_Click()
PlayWAV ("C:/DOLAN")
Form2.PopupMenu Form2.Label3, 2
'This one uses a 1 since its the first pull down menue on Form2
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label5.ForeColor = QBColor(12)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Form2.PopupMenu Form2.Label2, 2
'This one uses a 1 since its the first pull down menue on Form2Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = QBColor(12)
End Sub

Private Sub Label7_Click()
PlayWAV ("C:/DOLAN")
Form1.WindowState = 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = QBColor(12)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Timer1_Timer()
o
End Sub
