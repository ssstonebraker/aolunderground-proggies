VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "New User Reset"
   ClientHeight    =   3240
   ClientLeft      =   2265
   ClientTop       =   2190
   ClientWidth     =   2700
   ForeColor       =   &H00FFFFFF&
   HelpContextID   =   4340
   LinkTopic       =   "Form41"
   MaxButton       =   0   'False
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   3240
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   2640
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3960
      ScaleHeight     =   375
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label9 
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
      Left            =   2280
      TabIndex        =   14
      Top             =   -120
      Width           =   255
   End
   Begin VB.Label Label10 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Name Reset"
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
      Left            =   480
      TabIndex        =   12
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "seconds..."
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
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Now"
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
      Height          =   240
      Left            =   960
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   4200
      Top             =   1920
      Width           =   285
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Screen Name"
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
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Screen Name"
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
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your AOL Directory"
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
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFF00
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub

Private Sub Label10_Click()
Unload Form7
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub

Private Sub Label4_Click()
Picture1.Visible = True
Timeout (0.5)
Picture1.Visible = False
If UserSN > "" Then
MsgBox "Please exit America Online before using this feature.", 64
Exit Sub
End If
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Timer1.Enabled = True
Label6.Caption = "10"
Call AOLSNReset(Text2.Text, Text1.Text, Text3.Text)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub

Private Sub Label9_Click()
Form7.WindowState = 1
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub
Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub
Private Sub Text3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFFFF
End Sub
Private Sub Timer1_Timer()
If Label6.Caption = "0" Then
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Timer1.Enabled = False
Exit Sub
End If
Label6.Caption = Label6.Caption - 1
End Sub
