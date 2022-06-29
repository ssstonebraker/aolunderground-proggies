VERSION 4.00
Begin VB.Form Form2 
   Caption         =   "Popupmenu & Draw a line example By DeVil"
   ClientHeight    =   4140
   ClientLeft      =   150
   ClientTop       =   5445
   ClientWidth     =   6690
   Height          =   4830
   Left            =   90
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   Top             =   4815
   Width           =   6810
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu dot 
         Caption         =   "Dotted line"
      End
      Begin VB.Menu line 
         Caption         =   "Line"
      End
      Begin VB.Menu dash 
         Caption         =   "Dash"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub dash_Click()
Form1.DrawStyle = 1
End Sub

Private Sub dot_Click()
Form1.DrawStyle = 2
End Sub


Private Sub Line_Click()
Form1.DrawStyle = 0
End Sub


