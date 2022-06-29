VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Add Score"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   1635
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort"
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1635
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "Winner"
      Top             =   1080
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "Score"
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add  Score + Sort"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1635
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   60
      List            =   "Form1.frx":0016
      TabIndex        =   0
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
AddScore List1, Text1.Text, Text2.Text, True
End Sub

Private Sub Command2_Click()
SortScore List1
End Sub

Private Sub Command3_Click()
List1.Clear
End Sub


Private Sub Command4_Click()
AddScore List1, Text1.Text, Text2.Text, False
End Sub


