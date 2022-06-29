VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Listbox Search by nick(txt)"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3675
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox add 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "add"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox String1 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "Search String"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox lstfound 
      Height          =   840
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox lstsearch 
      Height          =   840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Found 
      Caption         =   "total strings found: 0"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "This is a listbox search example useing InStr made by Nick(txt)"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
'adds the text in the textbox add to lstsearch
If KeyCode = 13 Then
If add.Text = "" Then Exit Sub
lstsearch.AddItem add.Text
add.Text = ""
End If
End Sub

Private Sub Command1_Click()
'search for string
 For x = 0 To lstsearch.ListCount - 1
   If InStr(LCase(lstsearch.List(x)), LCase(String1.Text)) Then
     lstfound.AddItem lstsearch.List(x)
     Found.Caption = "total strings found: " & lstfound.ListCount & " "
     End If
 Next x

End Sub
