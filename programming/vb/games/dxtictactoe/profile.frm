VERSION 5.00
Begin VB.Form profile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Users Name"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   2940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Update Profile"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   435
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1530
      MaxLength       =   6
      TabIndex        =   0
      Top             =   435
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter Your name"
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   75
      Width           =   2205
   End
End
Attribute VB_Name = "profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
SaveSetting App.Title, "Settings", "profilename", profilename
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting App.Title, "Settings", "profilename", profilename
End Sub
