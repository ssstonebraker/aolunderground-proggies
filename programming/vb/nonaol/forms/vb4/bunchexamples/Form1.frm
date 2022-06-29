VERSION 4.00
Begin VB.Form Form1 
   ClientHeight    =   2100
   ClientLeft      =   2325
   ClientTop       =   1605
   ClientWidth     =   4245
   Height          =   2505
   Left            =   2265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4245
   Top             =   1260
   Width           =   4365
   Begin VB.CommandButton Command7 
      Caption         =   "count"
      Height          =   240
      Left            =   2100
      TabIndex        =   6
      Top             =   1350
      Width           =   525
   End
   Begin VB.CommandButton Command4 
      Caption         =   "remove from list"
      Height          =   270
      Left            =   30
      TabIndex        =   5
      Top             =   1830
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "add to list checking for duplicates"
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Top             =   1590
      Width           =   2595
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear List"
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   1350
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   165
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add to list"
      Height          =   240
      Left            =   1035
      TabIndex        =   1
      Top             =   1350
      Width           =   1065
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   4020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command5_Click()

End Sub

Private Sub Command1_Click()
'Iz representing eXcel 2001
If Text1.Text = "" Then Exit Sub ' tells vb to quit if the textbox is empty
List1.AddItem Text1.Text ' adds the text to the list
Text1.Text = "" ' clears the textbox
End Sub


Private Sub Command2_Click()
'Iz representing eXcel 2001
List1.Clear ' clears the list. simple, eh?
End Sub


Private Sub Command3_Click()
'Iz representing eXcel 2001
'this one may be harder to understand
'this code will add an item to a list, but not before it checks to make sure
'that its not already there. if it is, it will not add the item
'very useful
For i = 0 To List1.ListCount ' declares an intiger from 0 to how ever many items are on the list
X = List1.List(i) 'the number for i will change each time it trys a list item up until how ever
'many items are on the list. what this does is tell vb that each time x = a different list
'item
If LCase(Text1.Text) = LCase(X) Then Exit Sub 'i used the LCASE function to make the search
'not case sensitive
Next i 'returns to i to check the next item
List1.AddItem Text1.Text
Text1.Clear
End Sub

Private Sub Command4_Click()
'Iz representing eXcel 2001
'once a user clicks an item on the list
'this code will remove it from the list
'code also works nice if you put it in List_DBLCLICK
'so they can double click an item to remove it
List1.RemoveItem (List1.ListIndex) ' removes the selected item from listbox
End Sub

Private Sub Command7_Click()
'Iz representing eXcel 2001
MsgBox "there are " & List1.ListCount & " items on this list" '<displays a message box telling
'the user how many items are on the list using the ListCount property
End Sub


