VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IM Link Sender Example"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Link URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Persons SN:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu imlinks 
      Caption         =   "IM Links"
      Visible         =   0   'False
      Begin VB.Menu regular 
         Caption         =   "Regular"
      End
      Begin VB.Menu nonunderlined 
         Caption         =   "Non-Underlined"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.PopupMenu Form2.imlinks
End Sub

Private Sub Command4_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
FormOnTop Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form2.Hide
Form1.Show
End Sub

Private Sub nonunderlined_Click()
Call InstantMessage("" + Text1.Text + "", "< a href=" & Text2.Text & "></u>" & Text3.Text & "</a>")
End Sub

Private Sub regular_Click()
Call InstantMessage("" + Text1.Text + "", "< a href=" & Text2.Text & ">" & Text3.Text & "</a>")
End Sub
