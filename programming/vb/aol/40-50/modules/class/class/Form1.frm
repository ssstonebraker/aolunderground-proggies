VERSION 5.00
Begin VB.Form GetInfo 
   Caption         =   "Class Module examp by Ðigital-Flame"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get names"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If your havein a hard time getting stuff to call from the Class, take a look in the General section of this form!"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is a Class Module example made by Ðigital-Flame"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "User name:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Comp name:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "GetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
Dim Digital As New UserInfo
'
'

Private Sub Command1_Click()
Text1.Text = Digital.ComputerName
Text2.Text = Digital.UserNamez
End Sub

