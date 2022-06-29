VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "about..."
   ClientHeight    =   1530
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4770
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1056.033
   ScaleMode       =   0  'User
   ScaleWidth      =   4479.276
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "about.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "vb guides and examples by sonic"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4394.761
      Y1              =   610.843
      Y2              =   610.843
   End
   Begin VB.Label lblTitle 
      Caption         =   "min/max/restore example"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4394.761
      Y1              =   621.196
      Y2              =   621.196
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdOK_Click()
  End
End Sub

