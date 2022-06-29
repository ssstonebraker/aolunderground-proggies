VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Time"
   ClientHeight    =   1215
   ClientLeft      =   2160
   ClientTop       =   1530
   ClientWidth     =   1815
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   1815
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "1"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Custom Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Text1.Text - "1"
If Text1.Text = "0" Then
Text1.Text = "20"
End If
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text - -"1"
If Text1.Text = "21" Then
Text1.Text = "1"
End If
End Sub

Private Sub Command3_Click()
Form1.Text2.Text = "t"
Form1.Label8.Caption = Form2.Text1.Text
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Text2.Text = "t"
Form1.Label8.Caption = Form2.Text1.Text
End Sub
