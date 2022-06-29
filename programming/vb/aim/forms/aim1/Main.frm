VERSION 4.00
Begin VB.Form Main 
   Caption         =   "Main "
   ClientHeight    =   1830
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   1560
   Height          =   2235
   Icon            =   "Main.frx":0000
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   1560
   Top             =   1170
   Width           =   1680
   Begin VB.CommandButton Command4 
      Caption         =   "Main AiM Form"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Invite Form"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "IM Form"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Chat Form"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   0
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "Main"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Main.Hide
ChatForm.Show
End Sub

Private Sub Command2_Click()
Main.Hide
IMForm.Show
End Sub


Private Sub Command3_Click()
Main.Hide
InviteForm.Show
End Sub

Private Sub Command4_Click()
Main.Hide
MainForm.Show
End Sub

Private Sub Form_Load()
Label1.Caption = UserName
End Sub


