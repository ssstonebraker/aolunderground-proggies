VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Send Mail"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   ScaleHeight     =   3285
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Text            =   "Your msg"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Subject"
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendMail Text1.Text, Text2.Text, Text3.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Hide
Form1.Show
End Sub
