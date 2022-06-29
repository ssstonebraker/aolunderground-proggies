VERSION 5.00
Begin VB.Form frmChatJoin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Join Chat"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "frmChatJoin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmChatJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdJoin_Click()
  If frmSignOn.wskAIM.State = sckConnected And txtRoom.Text <> "" Then
    Call SendProc(2, "toc_chat_join 4 " & Chr(34) & txtRoom.Text & Chr(34) & Chr(0))
    Unload Me
  End If
End Sub
