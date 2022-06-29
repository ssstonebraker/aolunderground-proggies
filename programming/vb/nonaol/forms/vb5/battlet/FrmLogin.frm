VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "Login"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   Picture         =   "FrmLogin.frx":0000
   ScaleHeight     =   4710
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox UserNamet 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "Guest"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox PassWordt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Login 
      Caption         =   "Login"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   3360
      Picture         =   "FrmLogin.frx":2090E
      Top             =   2040
      Width           =   2700
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Login_Click()
  If UserNamet.Text = "" Then
     MsgBox "Must Enter a Username to Login", vbOKOnly
     Exit Sub
  End If
  Username = UserNamet.Text
  Password = PassWordt.Text
  Select Case UserNamet.Text
    Case "Anonymous", "anonymous", "Guest", "guest":
         BattleMain.ProceedLogin
         Unload Me
         Exit Sub
  End Select
  If Password <> "" Then
    BattleMain.ProceedLogin
  Else
    MsgBox "Must Enter a Password to Login", vbOKOnly
    Exit Sub
  End If
Unload Me
End Sub

Private Sub Quit_Click()
Unload Me
End Sub
