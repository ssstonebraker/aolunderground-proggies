VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UnACE Test"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDest 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "UnACE"
      Height          =   450
      Left            =   1800
      TabIndex        =   1
      Top             =   885
      Width           =   1215
   End
   Begin VB.TextBox txtACE 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "C:\Voodoo.ace"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Destination:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   510
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archive:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click()

Dim lResult As Long
    
    lResult = ACEExtract(txtACE.Text, txtDest.Text)
    
    MsgBox "ACEExtract returned: " & lResult

End Sub


