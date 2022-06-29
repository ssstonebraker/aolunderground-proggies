VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Make Msgbox"
      Height          =   495
      Left            =   2820
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Thread"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2730
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   $"thread.frx":0000
      Height          =   1875
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
en = False
Call Create
End Sub

Private Sub Command2_Click()
en = True
End Sub

Private Sub Command3_Click()
MsgBox "wow, it still runs"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Command2_Click
End Sub


