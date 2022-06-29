VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "using Val codes"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Words From Sublime Neo"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      Caption         =   "Commands"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   2775
      Begin VB.CommandButton Command4 
         Caption         =   "Divide"
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Mult."
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Subtract"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Answer"
      Height          =   855
      Left            =   1800
      TabIndex        =   2
      Top             =   0
      Width           =   975
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Number2"
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number 1"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = Val(Text1.Text) + Val(Text2.Text)

End Sub

Private Sub Command2_Click()
Text3.Text = Val(Text1.Text) - Val(Text2.Text)

End Sub

Private Sub Command3_Click()
Text3.Text = Val(Text1.Text) * Val(Text2.Text)

End Sub

Private Sub Command4_Click()
Text3.Text = Val(Text1.Text) / Val(Text2.Text)

End Sub

Private Sub Command5_Click()
MsgBox "I Role Play on AOL.And i first used the val codes to help the prog add the damage to the character, first i made a dice bot,then i made the val codes add the damage it rolled to the current damage it had now.So just so you know,the val codes does come in handy"

End Sub
