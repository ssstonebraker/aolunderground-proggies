VERSION 5.00
Begin VB.Form APISpy 
   Caption         =   "ApI SpYeR Example By Izekial83"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4365
   Icon            =   "ApI SpY.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "iM Me"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mail Me"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "ApI SpY.frx":030A
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3840
      Top             =   1680
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "APISpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Command1.Caption = "Start" Then Command1.Caption = "Stop": Timer1.Enabled = True: Exit Sub
    If Command1.Caption = "Stop" Then Command1.Caption = "Start": Timer1.Enabled = False: Exit Sub
End Sub

Private Sub Command2_Click()
    Call SendMail("Funkdemon@yahoo.com", "ApI SpY", "Sup,  I Was Checking Your API SpY Example")
End Sub

Private Sub Command3_Click()
    Call SendIM("izekial83", "Sup, I Am Checking your Spy")
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Form_Load()
    CenterForm Me
    StayOnTop Me
End Sub

Private Sub Form_Resize()
    StayOnTop Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Timer1_Timer()
        Call WindowSPY(Text1, Text2, Text3, Text4, Text5, Text6, Text7, Text8, Text9)
End Sub
