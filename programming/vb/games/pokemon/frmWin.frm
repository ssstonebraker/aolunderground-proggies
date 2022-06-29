VERSION 5.00
Begin VB.Form frmWin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrExit 
      Interval        =   5000
      Left            =   3225
      Top             =   780
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   1185
      Width           =   2325
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   4245
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Win!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1005
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   4140
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FormOnTop Me
End Sub
Private Sub tmrExit_Timer()
    If frmBattle.TurnA = 0 Then
        Unload frmInternetConnect
    Else
        Unload frmInternetListen
    End If
    Unload frmBattle
    Unload Me
End Sub
