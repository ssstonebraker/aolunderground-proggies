VERSION 4.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Example By DeVil"
   ClientHeight    =   840
   ClientLeft      =   675
   ClientTop       =   2475
   ClientWidth     =   6450
   DrawStyle       =   4  'Dash-Dot-Dot
   Height          =   1245
   Left            =   615
   LinkTopic       =   "Form3"
   MouseIcon       =   "Form3.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   840
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Top             =   2130
   Width           =   6570
   Begin VB.CommandButton Command3 
      BackColor       =   &H00000000&
      Caption         =   "Exit"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "INI reader and Writer"
      Height          =   495
      Left            =   3720
      MouseIcon       =   "Form3.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Popupmenu and Draw a line example"
      Height          =   495
      Left            =   0
      MouseIcon       =   "Form3.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "WELCOME to the Popupmenu and Draw a line example and .INI file writer and reader examples"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Visible = True
Unload Me
End Sub


Private Sub Command2_Click()
Form4.Visible = True
End Sub


Private Sub Command3_Click()
MsgBox "leaving? ok bye", 0, "BYEEEE!"
Unload Me
End Sub


