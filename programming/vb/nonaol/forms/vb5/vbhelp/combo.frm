VERSION 5.00
Begin VB.Form Combobox 
   Caption         =   "Comboboxes Exposed by:K¡m0"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   3045
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C l e a r   C o m b o b o x "
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<-Refresh"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<-Remove"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do it"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "I am here now"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Add below text to the list"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Highlighted Text:  "
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "List count:  "
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "List Index Number:  "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Combobox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Definitions
'line number: ListIndex
'ListIndex: is pertaning to a line of text in a Combobox
'List (line number): is the text of that line number
'AddItem: add the text ou want
'RemiveItem (line number):will remove a ListIndex
'ListCount: will get the total number of lines in a Combobox
'Clear: will clear the lsit (ie.Combo1.clear)
'
'All funtions are a copyright of Microsoft i am Just helping
'you to learn them

Private Sub Combo1_Click()
'This will get the index number or the
'line number of text that you clicked on
a = Combo1.ListIndex
'This tells you what line you clicked on
Text1 = a
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
'This is that actual text of the
'line you clicked on
Text3 = Combo1.Text
End Sub

Private Sub Command1_Click()
'by defult text4 is "I am here now"
If Text4 = "I am here now" Then
'Line below will add "Now I am here" only
'if text4 equals "I am here now"
Combo1.AddItem "Now I am here"
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
'Exit sub prevents the program from
'going past this point
Exit Sub
End If
'Text4 is the text you want to add
Combo1.AddItem Text4
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
End Sub

Private Sub Command2_Click()
'if there is no text in the list
'then dont remove anything
If Combo1.ListCount = "0" Then
'Exit sub prevents the program from
'going past this point
Exit Sub
End If
'This will remove the line number
'that you clicked on or the line
'number that you put in the box
Combo1.RemoveItem Text1
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
End Sub

Private Sub Command3_Click()
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
End Sub



Private Sub Command4_Click()
'clear will clear list1
Combo1.Clear
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
End Sub

Private Sub Command5_Click()
MsgBox "Don't STEAL form as is by doing so you are stealing ''21st Centry Software'' Copyrighted software. you can email me at Jay_Leno@hotmail.com", 64, "About"
End Sub

Private Sub Form_Load()
'clear will clear Combo1
Combo1.Clear
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
'This is the text i wanted to add
Combo1.AddItem "Made"
Combo1.AddItem "by"
Combo1.AddItem "K¡m0"
Combo1.AddItem "Employed By:"
Combo1.AddItem "21st Centry Software"
'This will get the index number or the
'line number of text that you clicked on
a = Combo1.ListIndex
'This will get how many lines of
'text there are in the label
Text2 = Combo1.ListCount
End Sub

Private Sub List1_Click()

End Sub

