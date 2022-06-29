VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   0  'None
   Caption         =   "Room Anoy"
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   1875
   ScaleWidth      =   2025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Abadi MT Condensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Text            =   "Sounds"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Anoy"
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
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   1560
      TabIndex        =   4
      Top             =   -120
      Width           =   255
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
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Stop"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Start"
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
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
StayOnTop Me
Combo1.AddItem ("GoodBye")
Combo1.AddItem ("FileDone")
Combo1.AddItem ("Welcome")
Combo1.AddItem ("IM")
Combo1.AddItem ("DROP")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
If FindChatRoom() = "" Then
Kazoo = MsgBox("You must be in a chat room to use this function", vbCritical, "HyPO")
Exit Sub
End If

SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5

SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
Timeout 0.5
SendChat "{S " + Combo1
End Sub

Private Sub Label2_Click()
Combo1 = ""
End Sub

Private Sub Label3_Click()
Unload Form16
End Sub

Private Sub Label4_Click()
Form16.WindowState = 1
End Sub
