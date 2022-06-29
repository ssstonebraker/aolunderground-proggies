VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Game"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image4 
      Height          =   780
      Left            =   1800
      Picture         =   "Form1.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Image Image3 
      Height          =   570
      Left            =   960
      Picture         =   "Form1.frx":067A
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   120
      Picture         =   "Form1.frx":0CB1
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   4080
      Picture         =   "Form1.frx":12E8
      Top             =   5520
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image1.Top = Image1.Top - 50
Image2.Left = Image1.Left
Image3.Left = Image1.Left
Image4.Left = Image1.Left
End If





If KeyCode = vbKeyRight Then
Image2.Top = Image1.Top
Image2.Left = Image1.Left
Image1.Visible = False
Image3.Visible = False
Image2.Visible = True
Image4.Visible = False
Image2.Left = Image2.Left + 50
Image1.Left = Image2.Left
Image3.Left = Image2.Left
Image4.Left = Image2.Left
End If

If KeyCode = vbKeyLeft Then

Image3.Top = Image1.Top
Image3.Left = Image1.Left
Image1.Visible = False
Image2.Visible = False
Image3.Visible = True
Image4.Visible = False
Image3.Left = Image3.Left - 50
Image1.Left = Image3.Left
Image2.Left = Image3.Left
Image4.Left = Image3.Left

End If

If KeyCode = vbKeyDown Then

Image4.Top = Image1.Top
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = True
Image4.Top = Image4.Top + 50
Image3.Top = Image4.Top
Image2.Top = Image4.Top
Image1.Top = Image4.Top
End If
End Sub
