VERSION 5.00
Begin VB.Form frmInvitation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Invitation"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmInvitation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDecline 
      Caption         =   "Decline"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Chat Invitation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1770
   End
End
Attribute VB_Name = "frmInvitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()
  Call SendProc(2, "toc_chat_accept " & Chr(34) & Right(Me.Tag, Len(Me.Tag) - 1) & Chr(34) & Chr(0))
  Unload Me
End Sub

Private Sub cmdDecline_Click()
  Unload Me
End Sub

