VERSION 4.00
Begin VB.Form IMForm 
   Caption         =   "IMTest Form"
   ClientHeight    =   1125
   ClientLeft      =   1140
   ClientTop       =   1530
   ClientWidth     =   4395
   Height          =   1530
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   4395
   Top             =   1185
   Width           =   4515
   Begin VB.CommandButton Command16 
      Caption         =   "OnTop"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Close IMs"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Close IM"
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Chnge Cap"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Mini IMs"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "IM Text"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Send IM 2"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Send IM"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Reply  IM"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Open IM"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Last Text"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Last Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Last Line"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "IM From.."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Count IMs"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear IM"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "IMForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call ImClear
End Sub



Private Sub Command10_Click()
Call ImSend2("TimOstman", "Hi!", True, False, False, False)
End Sub


Private Sub Command11_Click()
MsgBox ImText
End Sub

Private Sub Command12_Click()
Call MiniIMs
End Sub

Private Sub Command13_Click()
Call CapIm("Hi!")
End Sub

Private Sub Command14_Click()
Call CloseIM
End Sub

Private Sub Command15_Click()
Call CloseIMs
End Sub

Private Sub Command16_Click()
TopIm (True)
End Sub

Private Sub Command17_Click()
Call FindInfo
End Sub

Private Sub Command2_Click()
MsgBox ImCount
End Sub


Private Sub Command3_Click()
MsgBox ImFromWho
End Sub


Private Sub Command4_Click()
MsgBox ImLastLine
End Sub


Private Sub Command5_Click()
MsgBox ImLastName
End Sub


Private Sub Command6_Click()
MsgBox ImLastText2
End Sub


Private Sub Command7_Click()
Call ImOpen
End Sub


Private Sub Command8_Click()
Call ImReply("Hi")
End Sub


Private Sub Command9_Click()
Call ImSend("TimOstman", "Hi!")
End Sub


Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub


