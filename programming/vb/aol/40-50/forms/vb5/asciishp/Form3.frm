VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Water Rapids"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   840
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "        Stop"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "       Start"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "  X"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "  -"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Water Rapids Idle"
      BeginProperty Font 
         Name            =   "Ebola"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
StayOnTop Me
End Sub

Private Sub Label2_Click()
Me.WindowState = 1
End Sub

Private Sub Label3_Click()
Form3.Hide
End Sub

Private Sub Label6_Click()
Call IM_Keyword("$IM_OFF", " FeaR OwnZ ")
ChatSend "" & (" ")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Water Rapids Idle Bot")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ " & UserSN + " is now idle")
TimeOut 0.3
ChatSend "" & (" ")
Timer1.Enabled = True
End Sub

Private Sub Label7_Click()
Label4.caption = "0"
Call IM_Keyword("$IM_ON", " FeaR OwnZ ")
ChatSend "" & (" ")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Water Rapids Idle Bot")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ " & UserSN + " is now back")
TimeOut 0.3
ChatSend "" & (" ")
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
DoEvents
Label4.caption = Val(Label4) + 1
ChatSend "" & (" ")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ Water Rapids Idle Bot")
TimeOut 0.3
ChatSend "<font face=""Arial Narrow""></B></I></U></S>" & BlueGreen("‹v^•][ I've ben gone for [") + Label4 + ("] Min(s)")
TimeOut 0.3
ChatSend "" & (" ")
End Sub
