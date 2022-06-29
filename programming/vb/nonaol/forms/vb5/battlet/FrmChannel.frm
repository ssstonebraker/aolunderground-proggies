VERSION 5.00
Begin VB.Form FrmChannel 
   Caption         =   "Channel"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   Picture         =   "FrmChannel.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00004080&
      Caption         =   "Quit"
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00004080&
      Caption         =   "ok"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2970
      ItemData        =   "FrmChannel.frx":1EBC2
      Left            =   240
      List            =   "FrmChannel.frx":1EBF3
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtchannel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Channel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
BattleMain.Text1.Text = "/Channel " & txtchannel.Text
BattleMain.Send_Click
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub List1_Click()
txtchannel.Text = List1.List(List1.ListIndex)
End Sub
