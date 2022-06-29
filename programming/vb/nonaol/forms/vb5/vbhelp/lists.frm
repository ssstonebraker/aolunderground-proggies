VERSION 5.00
Begin VB.Form lists 
   Caption         =   "Comboboxes Exposed by:K¡m0"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C l e a r   C o m b o b o x "
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<-Refresh"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-Remove"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "I am here now"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Add below text to the list"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   4800
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   4800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Highlighted Text:  "
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "List count:  "
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "List Index Number:  "
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "lists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Definitions
'line number: ListIndex
'ListIndex: is pertaning to a line of text in a listbox
'List (line number): is the text of that line number
'AddItem: add the text ou want
'RemiveItem (line number):will remove a ListIndex
'ListCount: will get the total number of lines in a listbox
'Clear: will clear the lsit (ie.list1.clear)
'
'All funtions are a copyright

Private Sub Command1_Click()
'by defult text4 is "I am here now"
If Text4 = "I am here now" Then
'Line below will add "Now I am here" only
'if text4 equals "I am here now"
List1.AddItem "Now I am here"
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
'Exit sub prevents the program from
'going past this point
Exit Sub
End If
'Text4 is the text you want to add
List1.AddItem Text4
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
End Sub

Private Sub Command2_Click()
'if there is no text in the list
'then dont remove anything
If List1.ListCount = "0" Then
'Exit sub prevents the program from
'going past this point
Exit Sub
End If
'This will remove the line number
'that you clicked on or the line
'number that you put in the box
List1.RemoveItem Text1
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
End Sub

Private Sub Command3_Click()
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
End Sub



Private Sub Command4_Click()
'clear will clear list1
List1.Clear
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
End Sub

Private Sub Command5_Click()
MsgBox "Don't STEAL form as is by doing so you are stealing ''21st Centry Software'' Copyrighted software. you can email me at Jay_Leno@hotmail.com", 64, "About"
End Sub

Private Sub Form_Load()
'clear will clear List1
List1.Clear
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
'This is the text i wanted to add
List1.AddItem "Made"
List1.AddItem "by"
List1.AddItem "K¡m0"
List1.AddItem "Employed By:"
List1.AddItem "21st Centry Software"
'This will get the index number or the
'line number of text that you clicked on
a = List1.ListIndex
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
End Sub

Private Sub List1_Click()
'This will get the index number or the
'line number of text that you clicked on
a = List1.ListIndex
'This tells you what line you clicked on
Text1 = a
'This will get how many lines of
'text there are in the label
Text2 = List1.ListCount
'This is that actual text of the
'line you clicked on
Text3 = List1.List(a)
End Sub
