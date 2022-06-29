VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   Caption         =   "Intro Pic"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   2655
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
StayOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.Show
Unload Form11

End Sub
