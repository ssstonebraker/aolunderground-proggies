VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shell IE"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   2355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Shell IE"
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   855
      Width           =   2235
   End
   Begin VB.Label Label1 
      Caption         =   $"ShellIE.frx":0000
      Height          =   810
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2610
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
nResult = Shell("start.exe http://teamparadox.cjb.net, vbHide)
End Sub
