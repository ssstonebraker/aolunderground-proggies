VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fucked Up Mp3 Example - By Xen"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
Form2.Show
End Sub

