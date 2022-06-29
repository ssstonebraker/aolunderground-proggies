VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   0  'None
   Caption         =   "Invisible Talker"
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   2025
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Width           =   1215
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
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
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
      Left            =   1680
      TabIndex        =   6
      Top             =   -120
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "What To Say"
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
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Invisible Talker"
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
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Exit"
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
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Send"
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
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendChat("<font face=""comic sans ms""><FONT COLOR=#fffffe> " + Text1 + "")
Text1 = ""
End Sub

Private Sub Form_Load()
SendChat BlueBlack("<font face=""comic sans ms"">/`·-Full Moon")
Timeout (0.13)
SendChat BlueBlack("<font face=""comic sans ms"">/`·-Invisible Talker Loaded")
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub Label1_Click()
Call SendChat("<font face=""comic sans ms""><FONT COLOR=#fffffe> " + Text1 + "")
Text1 = ""
End Sub

Private Sub Label2_Click()
SendChat BlueBlack("<font face=""comic sans ms"">" + "/`·-Full Moon")
Timeout (0.13)
SendChat BlueBlack("<font face=""comic sans ms"">/`·-Invisible Talker Un-Loaded")
Unload Form14
End Sub

Private Sub Label3_Click()
SendChat BlueBlack("<font face=""comic sans ms"">/`·-Full Moon")
Timeout (0.13)
SendChat BlueBlack("<font face=""comic sans ms"">/`·-Invisible Talker Un-Loaded")
Unload Form14
End Sub

Private Sub Label6_Click()
Form14.WindowState = 1
End Sub

