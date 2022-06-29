VERSION 2.00
Begin Form Form1 
   Caption         =   "Focus Problems"
   ClientHeight    =   2010
   ClientLeft      =   1650
   ClientTop       =   2625
   ClientWidth     =   3630
   Height          =   2415
   Left            =   1590
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   3630
   Top             =   2280
   Width           =   3750
   Begin CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin TextBox txtPassword 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin TextBox txtUserid 
      Height          =   285
      Left            =   1620
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin Label Label2 
      Caption         =   "Password"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   660
      Width           =   975
   End
   Begin Label Label1 
      Caption         =   "Userid"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   735
   End
End
Option Explicit

Sub Command1_Click ()

    Unload Me
    
End Sub

Sub txtUserid_LostFocus ()

    If txtUserid.Text = "" Then txtUserid.SetFocus

End Sub

