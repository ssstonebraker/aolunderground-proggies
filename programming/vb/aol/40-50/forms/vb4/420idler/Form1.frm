VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   2415
   ClientTop       =   3435
   ClientWidth     =   7095
   Height          =   7590
   Left            =   2355
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Top             =   3090
   Width           =   7215
   Begin VB.Line Line4 
      BorderWidth     =   12
      X1              =   120
      X2              =   6960
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line3 
      BorderWidth     =   12
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderWidth     =   13
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderWidth     =   13
      X1              =   120
      X2              =   6960
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Form_Load()
Left = (Screen.Width - Width) \ 2
Top = (Screen.Height - Height) \ 2
Form1.Show
Call Pause("3")
Form2.Show
Call FormOnTop(Me)
Call Pause("1.2")
Unload Form1
End Sub


