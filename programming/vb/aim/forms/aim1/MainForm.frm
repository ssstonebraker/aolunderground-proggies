VERSION 4.00
Begin VB.Form MainForm 
   Caption         =   "MainTest Form"
   ClientHeight    =   1170
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3360
   Height          =   1575
   Left            =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3360
   Top             =   1170
   Width           =   3480
   Begin VB.CommandButton Command10 
      Caption         =   "Online?"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Load AiM"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit AiM"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Show Main"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Hide Main"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Find Prof-"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Find Main"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OnTop"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mini Main"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "User SN"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox UserName
End Sub


Private Sub Command10_Click()
MsgBox IsOnline
End Sub


Private Sub Command2_Click()
Call MiniMain
End Sub


Private Sub Command3_Click()
TopMain (True)
End Sub


Private Sub Command4_Click()
Call FindMain
End Sub


Private Sub Command5_Click()
Call FindProfile
End Sub


Private Sub Command6_Click()
HideMain (True)
End Sub


Private Sub Command7_Click()
HideMain (False)
End Sub


Private Sub Command8_Click()
Call MainClose
End Sub


Private Sub Command9_Click()
Call MainOpen("C:\Program Files\Aim95\aim.exe")
End Sub


Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub


