VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Secret Area"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2490
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Macro Kill"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sign On As Guest"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Spiral Scroller"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Elite Talker"
      Height          =   255
      Left            =   600
      MaskColor       =   &H000000FF&
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ChatSend EliteText("" & Text1.text + "")
Text1.text = ""
End Sub

Private Sub Command2_Click()
SpiralScroll ("" & Text1.text + "")
Text1.text = ""
End Sub

Private Sub Command3_Click()
Call SignOff
Call SignOnAsGuest
End Sub

Private Sub Command4_Click()
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
TimeOut 0.3
ChatSend "<font face=""Comic Sans MS"">" & ("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
StayOnTop Me
CenterForm Me
End Sub

Private Sub Form_Paint()
FadeFormBlue Me
End Sub
