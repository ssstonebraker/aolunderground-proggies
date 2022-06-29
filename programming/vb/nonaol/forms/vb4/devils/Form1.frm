VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Popupmenu & Draw a line example By DeVil"
   ClientHeight    =   4140
   ClientLeft      =   735
   ClientTop       =   2580
   ClientWidth     =   6690
   DrawStyle       =   1  'Dash
   Height          =   4545
   Left            =   675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Top             =   2235
   Width           =   6810
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose a different style"
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.PopupMenu Form2.Menu, 1
End Sub

Private Sub Command2_Click()
Form3.Visible = True
Unload Me
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Line -(X, Y)
End Sub


