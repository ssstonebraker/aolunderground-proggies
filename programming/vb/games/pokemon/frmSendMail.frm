VERSION 5.00
Begin VB.Form frmSendIM 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBorder 
      Height          =   60
      Left            =   -15
      TabIndex        =   7
      Top             =   3180
      Width           =   4545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Top             =   3315
      Width           =   1245
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   255
      TabIndex        =   3
      Top             =   3300
      Width           =   1245
   End
   Begin VB.TextBox txtMsg 
      Height          =   2445
      Left            =   1110
      TabIndex        =   2
      Top             =   660
      Width           =   3300
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1110
      TabIndex        =   1
      Top             =   285
      Width           =   3300
   End
   Begin VB.Label lblMsg 
      Alignment       =   1  'Right Justify
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label lblDrag 
      BackStyle       =   0  'Transparent
      Caption         =   " Send Instant Message"
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
      Height          =   225
      Left            =   225
      TabIndex        =   5
      Top             =   0
      Width           =   4020
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   300
      Width           =   915
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Left            =   0
      Picture         =   "frmSendMail.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   225
   End
   Begin VB.Image c1 
      Height          =   225
      Left            =   0
      Picture         =   "frmSendMail.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4530
   End
   Begin VB.Image c3 
      Height          =   90
      Left            =   3330
      Picture         =   "frmSendMail.frx":127C
      Stretch         =   -1  'True
      Top             =   4590
      Width           =   1065
   End
   Begin VB.Image c2 
      Height          =   90
      Left            =   3330
      Picture         =   "frmSendMail.frx":1C2E
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   1065
   End
   Begin VB.Image c4 
      Height          =   90
      Left            =   3330
      Picture         =   "frmSendMail.frx":256A
      Stretch         =   -1  'True
      Top             =   4650
      Width           =   1065
   End
End
Attribute VB_Name = "frmSendIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    FormOnTop Me
End Sub
Private Sub Form_GotFocus()
    c1.Picture = c3.Picture
End Sub
Private Sub Form_LostFocus()
    c1.Picture = c4.Picture
End Sub
Private Sub lblDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    c1.Picture = c2.Picture
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, WM_MOVE, 0
    c1.Picture = c3.Picture
End Sub
