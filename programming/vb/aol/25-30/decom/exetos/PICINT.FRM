VERSION 2.00
Begin Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   3375
   ClientLeft      =   1320
   ClientTop       =   1605
   ClientWidth     =   6375
   Height          =   3780
   Icon            =   PICINT.FRX:0000
   Left            =   1260
   LinkTopic       =   "Form4"
   ScaleHeight     =   3375
   ScaleWidth      =   6375
   Top             =   1260
   Width           =   6495
   Begin SSPanel Panel3D1 
      BevelInner      =   1  'Inset
      BevelWidth      =   3
      BorderWidth     =   2
      Caption         =   "Panel3D1"
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin Image Image1 
         Height          =   3150
         Left            =   120
         Picture         =   PICINT.FRX:0302
         Top             =   120
         Width           =   6120
      End
   End
End
' PICINT.FRM
Option Explicit

Sub Form_Load ()
subCD0 Me
End Sub

Sub Image1_DblClick ()
Unload Me
Form2.Show
End Sub
