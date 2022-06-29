VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Accept Aim Im"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   2250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Accept Win"
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   1125
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   $"FrmAccept.frx":0000
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   2265
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call AcceptAimIm
End Sub
