VERSION 2.00
Begin Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   480
   ClientLeft      =   9150
   ClientTop       =   6750
   ClientWidth     =   1560
   ControlBox      =   0   'False
   Height          =   885
   Left            =   9090
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   FORM4.FRX:0000
   ScaleHeight     =   480
   ScaleWidth      =   1560
   Top             =   6405
   Width           =   1680
   Begin PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      Picture         =   FORM4.FRX:0302
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
' FORM4.FRM
Option Explicit

Sub Form_Load ()
subE20 Me
End Sub

Sub Picture1_Click ()
Form4.Hide
Unload Me
Form1.Show
End Sub
