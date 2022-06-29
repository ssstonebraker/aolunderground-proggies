VERSION 5.00
Begin VB.Form Lists_Exposed_v2 
   Caption         =   "Lists Exposed v2 By Kim0"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "About"
      Height          =   555
      Left            =   1785
      TabIndex        =   20
      Top             =   2415
      Width           =   885
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Reset list"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   2730
      Width           =   1725
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear List"
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   2460
      Width           =   1725
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add to list any point"
      Height          =   285
      Left            =   2790
      TabIndex        =   16
      Top             =   2400
      Width           =   1800
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2985
      TabIndex        =   15
      Text            =   "Line 3"
      Top             =   2040
      Width           =   1635
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Look 4 --->"
      Height          =   285
      Left            =   1785
      TabIndex        =   14
      Top             =   2040
      Width           =   1065
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Look n delete ->"
      Height          =   300
      Left            =   1770
      TabIndex        =   13
      Top             =   1680
      Width           =   1440
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3270
      TabIndex        =   12
      Text            =   "Line 1"
      Top             =   1680
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add --->"
      Height          =   285
      Left            =   1755
      TabIndex        =   11
      Top             =   1335
      Width           =   885
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2685
      TabIndex        =   10
      Top             =   1320
      Width           =   1950
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2700
      TabIndex        =   9
      Top             =   990
      Width           =   1920
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3990
      TabIndex        =   7
      Top             =   660
      Width           =   630
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2430
      TabIndex        =   5
      Top             =   630
      Width           =   630
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add #1-10 With asdf"
      Height          =   255
      Left            =   2865
      TabIndex        =   3
      Top             =   15
      Width           =   1770
   End
   Begin VB.CommandButton Command2 
      Caption         =   "List All ascii charactors (255)"
      Height          =   300
      Left            =   1740
      TabIndex        =   2
      Top             =   300
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add #1-10"
      Height          =   255
      Left            =   1710
      TabIndex        =   1
      Top             =   15
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   1710
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2790
      TabIndex        =   17
      Top             =   2655
      Width           =   1785
   End
   Begin VB.Label Label3 
      Caption         =   "List Caption"
      Height          =   240
      Left            =   1800
      TabIndex        =   8
      Top             =   1035
      Width           =   885
   End
   Begin VB.Label Label2 
      Caption         =   "List Count"
      Height          =   210
      Left            =   3150
      TabIndex        =   6
      Top             =   690
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Index #"
      Height          =   225
      Left            =   1815
      TabIndex        =   4
      Top             =   675
      Width           =   615
   End
End
Attribute VB_Name = "Lists_Exposed_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'This is used to clear the list ever
'time you click on this button
List1.Clear
'You will be seeing this in other buttons
'look for the changes-----
'This is a loop made for looping with
'in a certain numbers of times specified.
'In this case it will only run 10 times.
'If you put 0 to 10 this would loop 11 times
'Besides loopping it also make changes the
'value of I and each time this will add to
'the I
For I = 1 To 10
'This will "AddItem" in this case (I) But
'do you remember what we said (I) was? (I) is a
'number 1 to 10
List1.AddItem I
'with out the Next not only will VB tell
'you are wrong but your computer will not
'no to go back to the for I = 1 to 10
Next I

End Sub

Private Sub Command10_Click()
Dim NL As String
NL = Chr(13) + Chr(10)
MsgBox "Hello Thank you for downloading Lists Exposed v2." + NL + _
       "Second I did not take the time to do all the . , in suff" + NL + _
       "Third you can do the same things with a combobox I used" + NL + _
       "lists because they are used more often" + NL + _
       "Thanks Dos " + NL + "http://www.hider.com/dos" + NL + _
       "and to the people on " + NL + "KnK's Web board" + NL + _
       "http://www.nwozone.com/KnK4o" + NL + _
       "I Spent a while on this and do plain on a v3" + NL + _
       "Please don't use form as is. If you have any aditional" + NL + _
       "questions then email me at Jay_Leno@hotmail.com", vbInformation, "About Lists Exposed v2 by Kim0"
End Sub

Private Sub Command2_Click()
'This is used to clear the list ever
'time you click on this button
List1.Clear
'You will be seeing this in other buttons
'look for the changes-----
'This is a loop made for looping with
'in a certain numbers of times specified.
'In this case it will only run 10 times.
'If you put 0 to 10 this would loop 11 times
'Besides loopping it also make changes the
'value of I and each time this will add to
'the I
For I = 0 To 255
'This will "AddItem" in this case I But
'do you remember what we said I was? I is a
'number 1 to 10
'adding on
'you see & " = " & Chr(I)
'lets say I = 15 this along
'with adding I(the number) it will add ¤
'it would look like this (15 = ¤) inside
'the listbox
List1.AddItem I & " = " & Chr(I)
'with out the Next not only will VB tell
'you are wrong but your computer will not
'no to go back to the For I = 1 To 10
Next I
End Sub

Private Sub Command3_Click()
'This is used to clear the list ever
'time you click on this button
List1.Clear
'You will be seeing this in other buttons
'look for the changes-----
'This is a loop made for looping with
'in a certain numbers of times specified.
'In this case it will only run 10 times.
'If you put 0 to 10 this would loop 11 times
'Besides loopping it also make changes the
'value of I and each time this will add to
'the I
For I = 1 To 10
'This will "AddItem" in this case I But
'do you remember what we said I was? I is a
'number 1 to 10
'adding on
'you see & "asdf"
'along with ading I(the number) it will add
'asdf
List1.AddItem I & "asdf"
'with out the Next not only will VB tell
'you are wrong but your computer will not
'no to go back to the For I = 1 To 10
Next I
End Sub

Private Sub Command4_Click()
'This line will add to the
'text from text4 to the end of the list
List1.AddItem Text4
End Sub

Private Sub Command5_Click()
'This tells your computer all it needs
'is this letter representing a string
'If you dont do this you will use up
'reasorces that not everone has
Dim B As String, C As String, D As String
'IF text5 equals nothing then this will
'tell you that it is blank and stop the
'funcion from carring on
If Text5 = "" Then
MsgBox "You must have something in the textbox"
Exit Sub
End If
'if there is nothing in the list a
'message box will popup saying whay you
'see here
If List1.ListCount = 0 Then MsgBox "You must have at lease one item on the list"
'This is a loop made for looping with
'in a certain numbers of times specified.
'In this case it will only run 10 times.
'If you put 0 to 10 this would loop 11 times
'Besides loopping it also make changes the
'value of A and each time this will add to
'the A
For A = 0 To List1.ListCount
'We take text5 and even if you look for
'LiNe 5 it will turn it into line5
'this is the optimal way to check to
'see if strings match thanks (Dos)
B = LCase(Replace(Text5, " ", ""))
'This line dose the same thing but checks the
'list caption on each line it will goto each
'line because of the For/Next loop
C = LCase(Replace(List1.List(A), " ", ""))
'Now we are ready to check and see if
'these match
If B = C Then
'If B does equal c then D = 1
D = 1
'now that we found what we were
'looking for we need to get out of the loop
GoTo line
'yes the Else statment
Else
'if b doesn't equal c then D = 0
D = 0
'end the if
End If
'this will send you back throu the loop
Next A
'remember the goto statment well it was
'targeted right here witch is outside of
'the loop that we wanted to stop
line:
'if D equals 1 then that is good because
'we want more RICE yum yum o no no were
'was I? O yea then we can remove it.
If D = 1 Then
'This line will remove the index number
'of the text you wanted to find remember
'that you didn't change the a value of A
'so it still equal what the For/Next loop
'left it
List1.RemoveItem A
'yes the Else statment....
'AGAIN? O THE INSANITY!!! hehehe
Else
'If the text was not found this line
'will tell you. I am using text5 insted
'of B because B is the no spases no caps (string)
MsgBox Text5 + " was not found"
'and end the if statment
End If
'this updates the List count
Text2 = List1.ListCount
End Sub

Private Sub Command6_Click()
'This tells your computer all it needs
'is this letter representing a string
'If you dont do this you will use up
'reasorces that not everone has
Dim B As String, C As String, D As String
'IF text6 equals nothing then this will
'tell you that it is blank and stop the
'funcion from carring on
If Text6 = "" Then
MsgBox "You must have something in the textbox"
Exit Sub
End If
'if there is nothing in the list a
'message box will popup saying whay you
'see here
If List1.ListCount = 0 Then MsgBox "You must have at lease one item on the list"
'This is a loop made for looping with
'in a certain numbers of times specified.
'In this case it will only run 10 times.
'If you put 0 to 10 this would loop 11 times
'Besides loopping it also make changes the
'value of A and each time this will add to
'the A
For A = 0 To List1.ListCount
'We take text5 and even if you look for
'LiNe 5 it will turn it into line5
'this is the optimal way to check to
'see if strings match thanks (Dos)
B = LCase(Replace(Text6, " ", ""))
'This line dose the same thing but checks the
'list caption on each line it will goto each
'line because of the For/Next loop
C = LCase(Replace(List1.List(A), " ", ""))
'Now we are ready to check and see if
'these match
If B = C Then
'If B does equal c then D = 1
D = 1
'now that we found what we were
'looking for we need to get out of the loop
GoTo line
'yes the Else statment
Else
'if b doesn't equal c then D = 0
D = 0
'end the if
End If
'this will send you back throu the loop
Next A
'remember the goto statment well it was
'targeted right here witch is outside of
'the loop that we wanted to stop
line:
'if D equals 1 then that is good because
'we want more RICE yum yum o no no were
'was I? O yea then we can remove it.
If D = 1 Then
'If D equals 1 then this will tell yes
'it was found?
MsgBox Text6 + " was found YES"
'yes the Else statment....
'AGAIN? O THE INSANITY!!! hehehe
Else
'If the text was not found this line
'will tell you. I am using text5 insted
'of B because B is the no spases no caps (string)
MsgBox Text6 + " was not found"
'and end the if statment
End If
'this updates the List count
Text2 = List1.ListCount
End Sub

Private Sub Command7_Click()
'This tells your computer all it needs
'is this letter representing a string
'If you dont do this you will use up
'reasorces that not everone has
Dim A As String
'You meet Input Box, Input Box meet you
'this is one of my favorite functions
'this allow the user to change the
'outcome.
A = InputBox("Were do you want to add ''" & Text7 & "'' you have a choise of 0 to " & Text2 & "?", "Were to add it")
'If you were to add this to a the list
'to more spaces then are avalible you
'would get an error (run time error 5)
If A > Text2 Then
'This tells you if the number you put
'in was too big
MsgBox A & " is to big of a number."
'There is that darn Else statment
'if this does not equal then do this
'(that is the best way I can explain
'the else statment)
Else
'Well this is a little different the is
'an , A at the end of it what do you think
'this does if you said something like
'telling VB what line to add the text to then
'you are right
List1.AddItem Text7, A
'this updates the List count
Text2 = List1.ListCount
'and end the if statment
End If
End Sub

Private Sub Command8_Click()
'this clears the list
List1.Clear
Text2 = ""
End Sub

Private Sub Command9_Click()
List1.Clear
For A = 0 To 5
List1.AddItem "Line " & A
Next A
Text2 = List1.ListCount
End Sub

Private Sub Form_Load()
For A = 0 To 5
List1.AddItem "Line " & A
Next A
Text2 = List1.ListCount
End Sub

Private Sub List1_Click()
'This tells your computer all it needs
'is this letter representing a string
'If you dont do this you will use up
'reasorces that not everone has
Dim A As String
'This will set the A string to the
'number line that you click on the list
'I can use this in other ways like the
'next line of code that says Text1 = A
A = List1.ListIndex
'This will tell Text1 to equal what A equals
Text1 = A
'This will not be needed for any other
'purpose so i will jsut have it equal text2
Text2 = List1.ListCount
'This is a code that you can set a string
'to equal what is in a listbox
'that enables you to use the text
'from a line in the list
Text3 = List1.List(A)

End Sub

Private Sub List1_DblClick()
'This tells your computer all it needs
'is this letter representing a string
'If you dont do this you will use up
'reasorces that not everone has
Dim A As String, B As String
'This will set the A string to the
'number line that you click on the list
'I can use this in other ways like the
'next line of code that says Text1 = A
A = List1.ListIndex
'This is going to ask you a question
B = MsgBox("Are you sure you want to delete " & A & "?", vbYesNo, "Are you sure?")
'if b = vbyes then it will Remove
'the item you dubble clicked on
If B = vbYes Then
'this code will delete what ever
'line you dubble click on
List1.RemoveItem A
'Else will tell your computer that
'if b equals anthing ELSE but vbyes
'then it will do the function after
'the Else statment
Else
'This will tell you computer to stop
'what you told it to do
Exit Sub
'This tells the if statment to end
End If
End Sub
