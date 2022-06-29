VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Greetz"
   ClientHeight    =   4785
   ClientLeft      =   2490
   ClientTop       =   1065
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Jinax"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "and  Jinax"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Cryo        Jolt"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cryo and Jolt for making killer .bas files."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ki\X/:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   $"KnK-Greetz.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KrAzY:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "KrAzY: Thanx for bein a good friend and helping me with VB.  Hurry up and get back on!"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "EnD:"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ReVeNgE:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"KnK-Greetz.frx":0094
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"KnK-Greetz.frx":0138
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Greetz"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
StayOnTop Me
End Sub
