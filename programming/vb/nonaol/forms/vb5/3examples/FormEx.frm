VERSION 5.00
Begin VB.Form FormEx 
   Caption         =   "min/max/restore example"
   ClientHeight    =   495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "maximize"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "restore"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "minimize"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "FormEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.WindowState = 1
End Sub

Private Sub Command2_Click()
    Me.WindowState = 0
    Command2.Enabled = False
    Command3.Enabled = True
End Sub

Private Sub Command3_Click()
    Me.WindowState = 2
    Command2.Enabled = True
    Command3.Enabled = False
End Sub

