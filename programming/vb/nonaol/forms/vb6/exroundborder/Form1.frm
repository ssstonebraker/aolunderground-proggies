VERSION 5.00
Begin VB.Form frmRound 
   BorderStyle     =   0  'None
   Caption         =   "Rounded Border Example"
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   286
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmRound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rounded Border Example
'Written by Chris Neuner
'in Visual Basic 6.0


Dim RoundIt As New clsRounder 'Create a new instance of the clsRounder class

Private Sub Form_Click()
    'Unload the form when it is clicked
    Unload Me
End Sub

Private Sub Form_Paint()
    'Paints the rounded rectangular region
    RoundIt.RoundedBorder Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Destroys the class
    Set RoundIt = Nothing
    'Ends the program
    End
End Sub
