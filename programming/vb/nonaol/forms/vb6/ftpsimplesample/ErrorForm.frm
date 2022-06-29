VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extended Error Message"
   ClientHeight    =   2928
   ClientLeft      =   2940
   ClientTop       =   3372
   ClientWidth     =   6996
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2928
   ScaleWidth      =   6996
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Dismiss"
      Height          =   372
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   2292
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6732
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Form2
End Sub
