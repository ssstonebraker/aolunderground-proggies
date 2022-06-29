VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SN Reveal"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reveal"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000008&
      ForeColor       =   &H8000000E&
      Height          =   325
      Left            =   600
      TabIndex        =   1
      Text            =   "Revealed SN."
      ToolTipText     =   "There screen name, REVEALED!"
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H8000000E&
      Height          =   325
      Left            =   600
      TabIndex        =   0
      Text            =   "SN to reveal."
      ToolTipText     =   "Copy and paste there screen name here!"
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1 = True Then Text2 = LCase(Text1)
End Sub

