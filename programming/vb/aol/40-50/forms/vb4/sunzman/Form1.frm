VERSION 4.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   -7035
   ClientTop       =   165
   ClientWidth     =   7905
   Height          =   495
   Left            =   -7095
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   90
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Top             =   -180
   Width           =   8025
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7440
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Do
Form1.Left = Form1.Left + 50
Loop Until Form1.Left > 1890
Do
Form1.Height = Form1.Height + 1
Loop Until Form1.Height > 6825
Do
Form1.Top = Form1.Top + 1
Loop Until Form1.Top > 3120
Do
Form1.Left = Form1.Left + 50
Loop Until Form1.Left > 11750
Form2.Show
Unload Form1
End Sub


