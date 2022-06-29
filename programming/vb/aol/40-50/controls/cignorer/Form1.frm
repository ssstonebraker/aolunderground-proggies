VERSION 5.00
Object = "{F778346B-2D27-11D3-8F9E-FB93F8432133}#1.0#0"; "CIGNORER.OCX"
Begin VB.Form Form1 
   Caption         =   "test's ignore example"
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "unignore"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ignore"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "person"
      Top             =   120
      Width           =   975
   End
   Begin ignorer.chatx chatx1 
      Height          =   1500
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   2646
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
chatx1.Person = "" + Text1 + "" 'sets person
chatx1.Ignore = True ' ignores
End Sub

Private Sub Command2_Click()
chatx1.Person = "" + Text1 + "" ' sets person
chatx1.Ignore = False 'unignore
End Sub

Private Sub Text1_Change()
'dont have to type full sn
'say the persons screen name is monkey893
'you can type 893 or monky either way it will ignore them
'have fun =)
End Sub
