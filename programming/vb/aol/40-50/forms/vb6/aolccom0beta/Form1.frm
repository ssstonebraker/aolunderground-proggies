VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   2220
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
main.Show
Unload Me
End Sub

Private Sub Form_Load()
dos32.FormTop Me
End Sub
