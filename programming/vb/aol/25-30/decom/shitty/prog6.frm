VERSION 4.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Min"
   ClientHeight    =   210
   ClientLeft      =   6330
   ClientTop       =   555
   ClientWidth     =   1560
   ControlBox      =   0   'False
   Height          =   615
   Left            =   6270
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Top             =   210
   Width           =   1680
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   30
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Deicide"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form6.Hide
Form4.Show
End Sub


Private Sub Command2_Click()
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
End Sub


Private Sub Form_Load()
Call StayOnTop(Form6)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub


