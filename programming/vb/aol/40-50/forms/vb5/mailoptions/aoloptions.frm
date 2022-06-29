VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "mail options"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "aoloptions.frx":0000
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "WHO DID THIS"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "write mail"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new mail"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AppActivate "AMERICA  ONLINE"  'activates aol
SendKeys "^(r)"            'tells aol to open new mail
End Sub

Private Sub Command2_Click()
AppActivate "AMERICA  ONLINE"  'activates aol
SendKeys "^(m)"            'tells aol to write mail
End Sub

Private Sub Command4_Click()
MsgBox "THEMAN", , "PROGRAMMER"  'sends a messagbox
'theman is the message,,,programmer is the title
End Sub
