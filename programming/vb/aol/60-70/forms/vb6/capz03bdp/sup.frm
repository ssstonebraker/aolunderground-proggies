VERSION 5.00
Begin VB.Form sup 
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "sup.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   5400
   End
End
Attribute VB_Name = "sup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
main.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
FormTop Me
End Sub
