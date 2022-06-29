VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2145
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   2145
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Frame1_Click()

End Sub

Private Sub Form_Load()
Call Send_Text("<font color=#123526ff>Baby MIMER")
Call Send_Text("<font color=#9654318ff>By<font color=#986464318ff> Flyman")
Call Send_Text("<font color=#123526ff>First 3.5 MIMER")
Me.Show
Pause 1.2
Form1.Show
End Sub
