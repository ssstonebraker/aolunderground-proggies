VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3075
   ScaleWidth      =   6570
   Begin VB.CheckBox Check1 
      Caption         =   "Remove Lists After Transfered"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Delete Buddy List"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Min."
      Height          =   195
      Left            =   5160
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Close"
      Height          =   195
      Left            =   5880
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Copy 
      Caption         =   "Create New List"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List5 
      Height          =   870
      Left            =   4440
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox List4 
      Height          =   870
      Left            =   2280
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox List3 
      Height          =   870
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   870
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   870
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Remove List #:"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim aol As Long, mdi As Long, Buddy As Long, IM As Long, Locate As Long, Setup As Long, Counting As Long
Dim BuddyList As Long, name As String, Create As Long, Edit As Long, EditListed As Long, Count As Long
Dim EditList As Long, listname As Long, buffer As String, TextLength As Long, ListNames As Long
Dim Category As String, Add As Long, Remove As Long, Name1 As Long, NewGroup As Long, Named As String
Dim Buffer1 As String, Combo As Long, Named1 As String, ExBuffer As String, ExBuddyList As Long, ExTextLength As Long

If Command1.Caption = "Start" Then

'First we look for the final stage. If its present then we can skip all the other BS
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TextLength& = GetWindowTextLength(EditList&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(EditList&, buffer$, TextLength& + 1)
 If buffer$ Like ("Edit List *") Then
  EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
   If EditListed& > 0 Then
    MsgBox "Now, all you have to do is select a category in the " & Chr(34) & "Buddy List Group Name" & Chr(34) & " box. Then press copy.", vbInformation, ""
    Combo& = FindWindowEx(EditListed&, 0&, "_AOL_Combobox", vbNullString)
    Call SendMessage(Combo&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Combo&, WM_LBUTTONUP, 0&, 0&)
    Command1.Caption = "Copy"
   Exit Sub
   End If
 End If

'We are going backwards so we dont mess up on ne thing. Plus its faster
'This looks for the second window. If its not found, then we go from scratch
BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
ExTextLength& = GetWindowTextLength(BuddyList&)
ExBuffer$ = String(ExTextLength&, 0&)
Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 If ExBuffer$ Like (GetUser + "'s Buddy *") Then
  ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
  Create& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Icon", vbNullString)
  Edit& = FindWindowEx(ExBuddyList&, Create&, "_AOL_Icon", vbNullString)
   If ExBuddyList& > 0 Then
    Pause (0.5)
    Call SendMessage(Edit&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Edit&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(Edit&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Edit&, WM_LBUTTONUP, 0&, 0&)
   End If
   Do
    EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    TextLength& = GetWindowTextLength(EditList&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(EditList&, buffer$, TextLength& + 1)
   Loop Until buffer$ Like ("Edit List *")
  EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
   If EditListed& > 0 Then
    MsgBox "Now, all you have to do is select a category in the " & Chr(34) & "Buddy List Group Name" & Chr(34) & " box. Then press copy.", vbInformation, ""
    Combo& = FindWindowEx(EditListed&, 0&, "_AOL_Combobox", vbNullString)
    Call SendMessage(Combo&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Combo&, WM_LBUTTONUP, 0&, 0&)
    Command1.Caption = "Copy"
    Exit Sub
   End If
 End If

'This looks for the buddy list and clicks the setup button
Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
 If Buddy& = 0 Then
  Call KeyWord("BuddyView")
   Do
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
   Loop Until Buddy& > 0
  Pause (1)
 End If
IM& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
Locate& = FindWindowEx(Buddy&, IM&, "_AOL_Icon", vbNullString)
Setup& = FindWindowEx(Buddy&, Locate&, "_AOL_Icon", vbNullString)
Call SendMessage(Setup&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Setup&, WM_LBUTTONUP, 0&, 0&)
'Loops until the second window is found
 Do
  BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
  ExTextLength& = GetWindowTextLength(BuddyList&)
  ExBuffer$ = String(ExTextLength&, 0&)
  Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 Loop Until ExBuffer$ Like (GetUser + "'s Buddy *")
ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
 If ExBuddyList& > 0 Then
  Create& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Icon", vbNullString)
  Edit& = FindWindowEx(ExBuddyList&, Create&, "_AOL_Icon", vbNullString)
  Pause (0.5)
  Call SendMessage(Edit&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Edit&, WM_LBUTTONUP, 0&, 0&)
 End If
'Going to the third window, once again a loop
 Do
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
  TextLength& = GetWindowTextLength(EditList&)
  buffer$ = String(TextLength&, 0&)
  Call GetWindowText(EditList&, buffer$, TextLength& + 1)
 Loop Until buffer$ Like ("Edit List *")
EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
 If EditListed& > 0 Then
  MsgBox "Now, all you have to do is select a category in the " & Chr(34) & "Buddy List Group Name" & Chr(34) & " box. Then press copy.", vbInformation, ""
  Combo& = FindWindowEx(EditListed&, 0&, "_AOL_Combobox", vbNullString)
  Call SendMessage(Combo&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Combo&, WM_LBUTTONUP, 0&, 0&)
  Command1.Caption = "Copy"
 Exit Sub
 End If
End If

If Command1.Caption = "Copy" Then
'First of course we have to find the window
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TextLength& = GetWindowTextLength(EditList&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(EditList&, buffer$, TextLength& + 1)
 If buffer$ Like ("Edit List *") Then
  EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
 End If
'Names the category
Buffer1$ = "Edit List"
Category$ = Mid(buffer$, 11, Len(Buffer1$))
'This checks if the category has been copied already
 If Text1.text = Category Then
  MsgBox "This category has already been copied", vbCritical, "Already Copied"
  Exit Sub
 Else
  If Text2.text = Category Then
   MsgBox "This category has already been copied", vbCritical, "Already Copied"
   Exit Sub
  Else
   If Text3.text = Category Then
    MsgBox "This category has already been copied", vbCritical, "Already Copied"
    Exit Sub
   Else
    If Text4.text = Category Then
     MsgBox "This category has already been copied", vbCritical, "Already Copied"
     Exit Sub
    Else
     If Text5.text = Category Then
      MsgBox "This category has already been copied", vbCritical, "Already Copied"
      Exit Sub
     End If
    End If
   End If
  End If
 End If
'Back to naming the category
 If Text1.text = "" Then
  Text1.text = Category$
  Label2.Caption = "1"
 Else
  If Text2.text = "" Then
   Text2.text = Category$
   Label2.Caption = "2"
  Else
   If Text3.text = "" Then
    Text3.text = Category$
    Label2.Caption = "3"
   Else
    If Text4.text = "" Then
     Text4.text = Category$
     Label2.Caption = "4"
    Else
     If Text5.text = "" Then
      Text5.text = Category$
      Label2.Caption = "5"
     Else
      MsgBox "Sorry, there is no more room to copy lists." & Chr(13) & Chr(10) & "If you would like to copy more, transfer all you have now and then come back and get the remaining lists.", vbCritical, "Room Limited"
      Exit Sub
     End If
    End If
   End If
  End If
 End If
'Finds the list box and counts it
ListNames& = FindWindowEx(EditListed&, 0, "_AOL_Listbox", vbNullString)
Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
Count = Count&
'Buttons
Add& = FindWindowEx(EditListed&, 0, "_AOL_Icon", vbNullString)
Remove& = FindWindowEx(EditListed&, Add&, "_AOL_Icon", vbNullString)
'Finds the text boxes
NewGroup& = FindWindowEx(EditListed&, 0, "_AOL_Edit", vbNullString)
Name1& = FindWindowEx(EditListed&, NewGroup&, "_AOL_Edit", vbNullString)
'Now we remove that names one by one and puts it into our program
 If Label2.Caption = "1" Then
  For X = 0 To Count
   Call SendMessage(Remove&, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessage(Remove&, WM_LBUTTONUP, 0&, 0&)
   Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
    If Count& > 0 Then
     Do
      Counting& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
     Loop Until Count& = Counting& + 1
     Named$ = GetText(Name1&)
     List1.AddItem Named$
    End If
  Next X
  Label3.Caption = Val(Label3.Caption) + 1
 Else
  If Label2.Caption = "2" Then
   For X = 0 To Count
    Call SendMessage(Remove&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Remove&, WM_LBUTTONUP, 0&, 0&)
    Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
     If Count& > 0 Then
      Do
       Counting& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
      Loop Until Count& = Counting& + 1
      Named$ = GetText(Name1&)
      List2.AddItem Named$
     End If
   Next X
   Label3.Caption = Val(Label3.Caption) + 1
  Else
   If Label2.Caption = "3" Then
    For X = 0 To Count
     Call SendMessage(Remove&, WM_LBUTTONDOWN, 0&, 0&)
     Call SendMessage(Remove&, WM_LBUTTONUP, 0&, 0&)
     Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
      If Count& > 0 Then
       Do
        Counting& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
       Loop Until Count& = Counting& + 1
       Named$ = GetText(Name1&)
       List3.AddItem Named$
      End If
    Next X
    Label3.Caption = Val(Label3.Caption) + 1
   Else
    If Label2.Caption = "4" Then
     For X = 0 To Count
      Call SendMessage(Remove&, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessage(Remove&, WM_LBUTTONUP, 0&, 0&)
      Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
       If Count& > 0 Then
        Do
         Counting& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
        Loop Until Count& = Counting& + 1
        Named$ = GetText(Name1&)
        List4.AddItem Named$
       End If
     Next X
     Label3.Caption = Val(Label3.Caption) + 1
    Else
     If Label2.Caption = "5" Then
      For X = 0 To Count
       Call SendMessage(Remove&, WM_LBUTTONDOWN, 0&, 0&)
       Call SendMessage(Remove&, WM_LBUTTONUP, 0&, 0&)
       Count& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
        If Count& > 0 Then
         Do
          Counting& = SendMessage(ListNames&, LB_GETCOUNT, 0&, 0&)
         Loop Until Count& = Counting& + 1
         Named$ = GetText(Name1&)
         List5.AddItem Named$
        End If
      Next X
      Label3.Caption = Val(Label3.Caption) + 1
     End If
    End If
   End If
  End If
 End If
reply = MsgBox("Press OK for a new category or press cancel if you are done", vbOKCancel, "")
 If reply = vbOK Then
  Combo& = FindWindowEx(EditListed&, 0&, "_AOL_Combobox", vbNullString)
  Call SendMessage(Combo&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Combo&, WM_LBUTTONUP, 0&, 0&)
 End If
 If reply = vbCancel Then
  Call SendMessage(EditListed&, WM_CLOSE, 0&, 0&)
  Command1.Caption = "Start"
  Pause (0.5)
  BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
  ExTextLength& = GetWindowTextLength(BuddyList&)
  ExBuffer$ = String(ExTextLength&, 0&)
  Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
  If ExBuffer$ Like (GetUser + "'s Buddy *") Then
   ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
   Call SendMessage(ExBuddyList&, WM_CLOSE, 0&, 0&)
  End If
 End If
End If
End Sub

Private Sub Command10_Click()
If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" Then
 MsgBox "There are no lists to save", vbCritical, ""
 Exit Sub
End If
If Text1 > "" Then
 Text6 = Dir1 + "\" + Text1 + ".prlt"
 Call SaveListBox(Text6, List1)
End If
If Text2 > "" Then
 Text6 = Dir1 + "\" + Text2 + ".prlt"
 Call SaveListBox(Text6, List1)
End If
If Text3 > "" Then
 Text6 = Dir1 + "\" + Text3 + ".prlt"
 Call SaveListBox(Text6, List1)
End If
If Text4 > "" Then
 Text6 = Dir1 + "\" + Text4 + ".prlt"
 Call SaveListBox(Text6, List1)
End If
If Text5 > "" Then
 Text6 = Dir1 + "\" + Text5 + ".prlt"
 Call SaveListBox(Text6, List1)
End If
End Sub

Private Sub Command11_Click()
Dim Group As String, MyString As String, FindString As String, Spot As Long
If File1.ListCount = 0 Then
 MsgBox "There are no lists to load", vbCritical, ""
 Exit Sub
End If
For X = 0 To File1.ListCount
 If X = "" Then
  Exit Sub
 End If
 If Text1 = "" Then
  Text6 = Dir1 + "\" + File1.List(X)
  Group$ = InStr(Right(Text6, 5), Text6)
  MsgBox Group$
  Text1 = File1.List(X)
  Call Loadlistbox(Text6, List1)
 Else
  If Text2 = "" Then
   Text6 = Dir1 + "\" + File1.List(X)
   Text2 = File1.List(X)
   Call Loadlistbox(Text6, List1)
  Else
   If Text3 = "" Then
    Text6 = Dir1 + "\" + File1.List(X)
    Text3 = File1.List(X)
    Call Loadlistbox(Text6, List1)
   Else
    If Text4 = "" Then
     Text6 = Dir1 + "\" + File1.List(X)
     Text4 = File1.List(X)
     Call Loadlistbox(Text6, List1)
    Else
     If Text5 = "" Then
      Text6 = Dir1 + "\" + File1.List(X)
      Text5 = File1.List(X)
      Call Loadlistbox(Text6, List1)
     End If
    End If
   End If
  End If
 End If
Next X
End Sub

Private Sub Command2_Click()
If Text1 = "" Then
Exit Sub
Else
reply = MsgBox("Are you sure you want to delete the list " & Text1.text & "?", 4, "Delete?")
If reply = vbYes Then
Text1.text = ""
List1.Clear
Label3.Caption = Val(Label3.Caption) - 1
Else
Exit Sub
End If
End If
End Sub

Private Sub Command3_Click()
If Text2 = "" Then
Exit Sub
Else
reply = MsgBox("Are you sure you want to delete the list " & Text2.text & "?", 4, "Delete?")
If reply = vbYes Then
Text2.text = ""
List2.Clear
Label3.Caption = Val(Label3.Caption) - 1
Else
Exit Sub
End If
End If
End Sub

Private Sub Command4_Click()
If Text3 = "" Then
Exit Sub
Else
reply = MsgBox("Are you sure you want to delete the list " & Text3.text & "?", 4, "Delete?")
If reply = vbYes Then
Text3.text = ""
List3.Clear
Label3.Caption = Val(Label3.Caption) - 1
Else
Exit Sub
End If
End If
End Sub

Private Sub Command5_Click()
If Text4 = "" Then
Exit Sub
Else
reply = MsgBox("Are you sure you want to delete the list " & Text4.text & "?", 4, "Delete?")
If reply = vbYes Then
Text4.text = ""
List4.Clear
Label3.Caption = Val(Label3.Caption) - 1
Else
Exit Sub
End If
End If
End Sub

Private Sub Command6_Click()
If Text5 = "" Then
Exit Sub
Else
reply = MsgBox("Are you sure you want to delete the list " & Text5.text & "?", 4, "Delete?")
If reply = vbYes Then
Text5.text = ""
List5.Clear
Label3.Caption = Val(Label3.Caption) - 1
Else
Exit Sub
End If
End If
End Sub

Private Sub Command7_Click()
Form2.Hide
End Sub

Private Sub Command8_Click()
Form2.WindowState = 1
End Sub

Private Sub Command9_Click()
Dim aol As Long, mdi As Long, Buddy As Long, BuddyList As Long, EditList As Long, EditListed As Long
Dim IM As Long, Locate As Long, Setup As Long, Create As Long, Add As Long, Remove As Long, Save As Long
Dim GroupBox As Long, name As String, NameBox As Long
Dim List As Long, Count As Long, Yes As Long, No As Long, MessageBox As Long
Dim Editor As Long, Buffer1 As String, TextLength As Long
Dim GroupNamed1 As String, ListNamed1 As String, OkBox As Long, Ok As Long

reply = MsgBox("Are you sure you want to delete your entire buddylist?", 36, "")
 If reply = vbNo Then
  Exit Sub
 End If
'First just for purposes of selcting a window, we are going to look
'For the windows and then close them all. Then we re-open them
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TextLength& = GetWindowTextLength(EditList&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(EditList&, buffer$, TextLength& + 1)
 If buffer$ Like ("Edit List *") Then
  EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
  If EditListed& > 0 Then
   Call SendMessage(EditListed&, WM_CLOSE, 0&, 0&)
  End If
 End If
Pause (0.5)
BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
ExTextLength& = GetWindowTextLength(BuddyList&)
ExBuffer$ = String(ExTextLength&, 0&)
Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 If ExBuffer Like (GetUser + "'s Buddy *") Then
  ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
  If ExBuddyList& > 0 Then
   Call SendMessage(ExBuddyList&, WM_CLOSE, 0&, 0&)
  End If
 End If
Pause (0.5)
'Finds BuddyList window
Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
'If buddylist is not there we get it!
 If Buddy& = 0 Then
  Call KeyWord("BuddyView")
   Do
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
   Loop Until Buddy& > 0
  Pause (0.7)
 End If
IM& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
Locate& = FindWindowEx(Buddy&, IM&, "_AOL_Icon", vbNullString)
Setup& = FindWindowEx(Buddy&, Locate&, "_AOL_Icon", vbNullString)
Call SendMessage(Setup&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Setup&, WM_LBUTTONUP, 0&, 0&)
'Loops after the setup is pressed until the second window is found
 Do
  BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
  ExTextLength& = GetWindowTextLength(BuddyList&)
  ExBuffer$ = String(ExTextLength&, 0&)
  Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 Loop Until ExBuffer$ Like (GetUser + "'s Buddy *")
ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
List& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Listbox", vbnullsting)
Count& = SendMessage(List&, LB_GETCOUNT, 0&, 0&)
Create& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Icon", vbNullString)
Edit& = FindWindowEx(ExBuddyList&, Create&, "_AOL_Icon", vbNullString)
Delete& = FindWindowEx(ExBuddyList&, Edit&, "_AOL_Icon", vbNullString)
'Counts the list and then removes all categories there already
 If Count& > 0 Then
  If Count& = 1 Then
   Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
    Do
     MessageBox& = FindWindow("_AOL_Modal", vbNullString)
    Loop Until MessageBox& > 0
   Pause (0.6)
   Yes& = FindWindowEx(MessageBox&, 0&, "_AOL_Icon", vbNullString)
   Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
   Pause (0.9)
  Else
   For X = 0 To Count& - 1
    Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
     Do
      MessageBox& = FindWindow("_AOL_Modal", vbNullString)
     Loop Until MessageBox& > 0
    Pause (0.6)
    Yes& = FindWindowEx(MessageBox&, 0&, "_AOL_Icon", vbNullString)
    Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  End If
 End If
End Sub
Private Sub Copy_Click()
Dim aol As Long, mdi As Long, Buddy As Long, BuddyList As Long, EditList As Long, EditListed As Long
Dim IM As Long, Locate As Long, Setup As Long, Create As Long, Add As Long, Remove As Long, Save As Long
Dim GroupBox As Long, name As String, NameBox As Long
Dim List As Long, Count As Long, Yes As Long, No As Long, MessageBox As Long
Dim Editor As Long, Buffer1 As String, TextLength As Long
Dim GroupNamed1 As String, ListNamed1 As String, OkBox As Long, Ok As Long
Dim Counter As Long, CountList As Long, Listing As Long, View As Long

'First we see if there are any names to transfer
 If Text1 = "" Then
  If Text2 = "" Then
   If Text3 = "" Then
    If Text4 = "" Then
     If Text5 = "" Then
      MsgBox "There are no names to transfer", vbCritical, "No Transfer"
      Exit Sub
     End If
    End If
   End If
  End If
 End If

'First just for purposes of selcting a window, we are going to look
'For the windows and then close them all. Then we re-open them
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
EditList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TextLength& = GetWindowTextLength(EditList&)
buffer$ = String(TextLength&, 0&)
Call GetWindowText(EditList&, buffer$, TextLength& + 1)
 If buffer$ Like ("Edit List *") Then
  EditListed& = FindWindowEx(mdi&, 0&, "AOL Child", buffer$)
  If EditListed& > 0 Then
   Call SendMessage(EditListed&, WM_CLOSE, 0&, 0&)
  End If
 End If
Pause (0.5)
BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
ExTextLength& = GetWindowTextLength(BuddyList&)
ExBuffer$ = String(ExTextLength&, 0&)
Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 If ExBuffer Like (GetUser + "'s Buddy *") Then
  ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
  If ExBuddyList& > 0 Then
   Call SendMessage(ExBuddyList&, WM_CLOSE, 0&, 0&)
  End If
 End If
Pause (0.5)
'Finds BuddyList window
Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
'If buddylist is not there we get it!
 If Buddy& = 0 Then
  Call KeyWord("BuddyView")
   Do
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Buddy& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
   Loop Until Buddy& > 0
  Pause (0.7)
 End If
IM& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
Locate& = FindWindowEx(Buddy&, IM&, "_AOL_Icon", vbNullString)
Setup& = FindWindowEx(Buddy&, Locate&, "_AOL_Icon", vbNullString)
Call SendMessage(Setup&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Setup&, WM_LBUTTONUP, 0&, 0&)
Call SendMessage(Setup&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Setup&, WM_LBUTTONUP, 0&, 0&)
'Loops after the setup is pressed until the second window is found
 Do
  BuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
  ExTextLength& = GetWindowTextLength(BuddyList&)
  ExBuffer$ = String(ExTextLength&, 0&)
  Call GetWindowText(BuddyList&, ExBuffer$, ExTextLength& + 1)
 Loop Until ExBuffer$ Like (GetUser + "'s Buddy *")
ExBuddyList& = FindWindowEx(mdi&, 0&, "AOL Child", ExBuffer$)
List& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Listbox", vbnullsting)
Count& = SendMessage(List&, LB_GETCOUNT, 0&, 0&)
Create& = FindWindowEx(ExBuddyList&, 0&, "_AOL_Icon", vbNullString)
Edit& = FindWindowEx(ExBuddyList&, Create&, "_AOL_Icon", vbNullString)
Delete& = FindWindowEx(ExBuddyList&, Edit&, "_AOL_Icon", vbNullString)
'Counts the list and then removes all categories there already
 If Count& > 0 Then
  If Count& = 1 Then
   Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
    Do
     MessageBox& = FindWindow("_AOL_Modal", vbNullString)
    Loop Until MessageBox& > 0
   Pause (0.6)
   Yes& = FindWindowEx(MessageBox&, 0&, "_AOL_Icon", vbNullString)
   Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
   Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
   Pause (0.9)
  Else
   For X = 0 To Count& - 1
    Call SendMessage(Delete&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Delete&, WM_LBUTTONUP, 0&, 0&)
     Do
      MessageBox& = FindWindow("_AOL_Modal", vbNullString)
     Loop Until MessageBox& > 0
    Pause (0.6)
    Yes& = FindWindowEx(MessageBox&, 0&, "_AOL_Icon", vbNullString)
    Call SendMessage(Yes&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Yes&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  End If
 End If
'Now to start adding the names to the buddylist
 If Text1.text > "" Then
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Pause (1)
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
  'Another loop
   Do
   Loop Until EditList& > 0
  'These are all the textboxes and buttons
  GroupBox& = FindWindowEx(EditList&, 0&, "_AOL_Edit", vbNullString)
  NameBox& = FindWindowEx(EditList&, GroupBox&, "_AOL_Edit", vbNullString)
  Add& = FindWindowEx(EditList&, 0&, "_AOL_Icon", vbNullString)
  Remove& = FindWindowEx(EditList&, Add&, "_AOL_Icon", vbNullString)
  Save& = FindWindowEx(EditList&, Remove&, "_AOL_Icon", vbNullString)
  GroupNamed1$ = Text1.text
  Call SendMessageByString(GroupBox&, WM_SETTEXT, 0&, GroupNamed1$)
  Pause (0.5)
   For X = 0 To List1.ListCount - 1
    ListNamed1$ = List1.List(X)
    Call SendMessageByString(NameBox&, WM_SETTEXT, 0&, ListNamed1$)
    Pause (0.2)
    Call SendMessage(Add&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Add&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  Call SendMessage(Save&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Save&, WM_LBUTTONUP, 0&, 0&)
   Do
    OkBox& = FindWindow("#32770", "America Online")
   Loop Until OkBox& > 0
   Pause (0.6)
  OkBox& = FindWindow("#32770", "America Online")
  Ok& = FindWindowEx(OkBox&, 0&, "Button", vbNullString)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Label3.Caption = Val(Label3.Caption) - 1
   If Check1.Value = True Then
    Text1 = ""
    List1.Clear
   End If
  Pause (1)
 End If
 If Text2.text > "" Then
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Pause (1)
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
  'Another loop
   Do
   Loop Until EditList& > 0
  'These are all the textboxes and buttons
  GroupBox& = FindWindowEx(EditList&, 0&, "_AOL_Edit", vbNullString)
  NameBox& = FindWindowEx(EditList&, GroupBox&, "_AOL_Edit", vbNullString)
  Add& = FindWindowEx(EditList&, 0&, "_AOL_Icon", vbNullString)
  Remove& = FindWindowEx(EditList&, Add&, "_AOL_Icon", vbNullString)
  Save& = FindWindowEx(EditList&, Remove&, "_AOL_Icon", vbNullString)
  GroupNamed1$ = Text2.text
  Call SendMessageByString(GroupBox&, WM_SETTEXT, 0&, GroupNamed1$)
  Pause (0.5)
   For X = 0 To List2.ListCount - 1
    ListNamed1$ = List2.List(X)
    Call SendMessageByString(NameBox&, WM_SETTEXT, 0&, ListNamed1$)
    Pause (0.2)
    Call SendMessage(Add&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Add&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  Call SendMessage(Save&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Save&, WM_LBUTTONUP, 0&, 0&)
   Do
    OkBox& = FindWindow("#32770", "America Online")
   Loop Until OkBox& > 0
   Pause (0.6)
  OkBox& = FindWindow("#32770", "America Online")
  Ok& = FindWindowEx(OkBox&, 0&, "Button", vbNullString)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Label3.Caption = Val(Label3.Caption) - 1
   If Check1.Value = True Then
    Text2 = ""
    List2.Clear
   End If
  Pause (1)
 End If
 If Text3.text > "" Then
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Pause (1)
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
  'Another loop
   Do
   Loop Until EditList& > 0
  'These are all the textboxes and buttons
  GroupBox& = FindWindowEx(EditList&, 0&, "_AOL_Edit", vbNullString)
  NameBox& = FindWindowEx(EditList&, GroupBox&, "_AOL_Edit", vbNullString)
  Add& = FindWindowEx(EditList&, 0&, "_AOL_Icon", vbNullString)
  Remove& = FindWindowEx(EditList&, Add&, "_AOL_Icon", vbNullString)
  Save& = FindWindowEx(EditList&, Remove&, "_AOL_Icon", vbNullString)
  GroupNamed1$ = Text3.text
  Call SendMessageByString(GroupBox&, WM_SETTEXT, 0&, GroupNamed1$)
  Pause (0.5)
   For X = 0 To List3.ListCount - 1
    ListNamed1$ = List3.List(X)
    Call SendMessageByString(NameBox&, WM_SETTEXT, 0&, ListNamed1$)
    Pause (0.2)
    Call SendMessage(Add&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Add&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  Call SendMessage(Save&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Save&, WM_LBUTTONUP, 0&, 0&)
   Do
    OkBox& = FindWindow("#32770", "America Online")
   Loop Until OkBox& > 0
   Pause (0.6)
  OkBox& = FindWindow("#32770", "America Online")
  Ok& = FindWindowEx(OkBox&, 0&, "Button", vbNullString)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Label3.Caption = Val(Label3.Caption) - 1
   If Check1.Value = True Then
    Text3 = ""
    List3.Clear
   End If
  Pause (1)
 End If
 If Text4.text > "" Then
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Pause (1)
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
  'Another loop
   Do
   Loop Until EditList& > 0
  'These are all the textboxes and buttons
  GroupBox& = FindWindowEx(EditList&, 0&, "_AOL_Edit", vbNullString)
  NameBox& = FindWindowEx(EditList&, GroupBox&, "_AOL_Edit", vbNullString)
  Add& = FindWindowEx(EditList&, 0&, "_AOL_Icon", vbNullString)
  Remove& = FindWindowEx(EditList&, Add&, "_AOL_Icon", vbNullString)
  Save& = FindWindowEx(EditList&, Remove&, "_AOL_Icon", vbNullString)
  GroupNamed1$ = Text4.text
  Call SendMessageByString(GroupBox&, WM_SETTEXT, 0&, GroupNamed1$)
  Pause (0.5)
   For X = 0 To List4.ListCount - 1
    ListNamed1$ = List4.List(X)
    Call SendMessageByString(NameBox&, WM_SETTEXT, 0&, ListNamed1$)
    Pause (0.2)
    Call SendMessage(Add&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Add&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  Call SendMessage(Save&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Save&, WM_LBUTTONUP, 0&, 0&)
   Do
    OkBox& = FindWindow("#32770", "America Online")
   Loop Until OkBox& > 0
   Pause (0.6)
  OkBox& = FindWindow("#32770", "America Online")
  Ok& = FindWindowEx(OkBox&, 0&, "Button", vbNullString)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Label3.Caption = Val(Label3.Caption) - 1
   If Check1.Value = True Then
    Text4 = ""
    List4.Clear
   End If
  Pause (1)
 End If
 If Text5.text > "" Then
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Create&, WM_LBUTTONUP, 0&, 0&)
  Pause (1)
  aol& = FindWindow("AOL Frame25", vbNullString)
  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
  EditList& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
  'Another loop
   Do
   Loop Until EditList& > 0
  'These are all the textboxes and buttons
  GroupBox& = FindWindowEx(EditList&, 0&, "_AOL_Edit", vbNullString)
  NameBox& = FindWindowEx(EditList&, GroupBox&, "_AOL_Edit", vbNullString)
  Add& = FindWindowEx(EditList&, 0&, "_AOL_Icon", vbNullString)
  Remove& = FindWindowEx(EditList&, Add&, "_AOL_Icon", vbNullString)
  Save& = FindWindowEx(EditList&, Remove&, "_AOL_Icon", vbNullString)
  GroupNamed1$ = Text5.text
  Call SendMessageByString(GroupBox&, WM_SETTEXT, 0&, GroupNamed1$)
  Pause (0.5)
   For X = 0 To List5.ListCount - 1
    ListNamed1$ = List5.List(X)
    Call SendMessageByString(NameBox&, WM_SETTEXT, 0&, ListNamed1$)
    Pause (0.2)
    Call SendMessage(Add&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Add&, WM_LBUTTONUP, 0&, 0&)
    Pause (0.9)
   Next X
  Call SendMessage(Save&, WM_LBUTTONDOWN, 0&, 0&)
  Call SendMessage(Save&, WM_LBUTTONUP, 0&, 0&)
   Do
    OkBox& = FindWindow("#32770", "America Online")
   Loop Until OkBox& > 0
   Pause (0.6)
  OkBox& = FindWindow("#32770", "America Online")
  Ok& = FindWindowEx(OkBox&, 0&, "Button", vbNullString)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYDOWN, VK_SPACE, 0&)
  Call SendMessage(Ok&, WM_KEYUP, VK_SPACE, 0&)
  Label3.Caption = Val(Label3.Caption) - 1
   If Check1.Value = True Then
    Text5 = ""
    List5.Clear
   End If
  Pause (1)
 End If
 If Label3.Caption = "0" Then
  Call SendMessage(ExBuddyList&, WM_CLOSE, 0&, 0&)
  MsgBox "BuddyList Transfer Complete", vbInformation, ""
 End If
End Sub

Private Sub Form_Load()
Dim aol As Long
FormOnTop Me
aol& = FindWindow("AOL Frame25", vbNullString)
 If aol& = 0 Then
  MsgBox "Please load AOL before using this option", vbCritical, "Load AOL"
  End
 End If
End Sub

Private Sub List1_DblClick()
 If List1.text = "" Then
  Exit Sub
 End If
reply = MsgBox("Do you want to remove " & List1.text & " from the list?", 36, "Remove")
 If reply = vbYes Then
  List1.RemoveItem List1.ListIndex
 Else
  Exit Sub
 End If
End Sub

Private Sub List2_DblClick()
 If List2.text = "" Then
  Exit Sub
 End If
reply = MsgBox("Do you want to remove " & List2.text & " from the list?", 36, "Remove")
 If reply = vbYes Then
  List1.RemoveItem List2.ListIndex
 Else
  Exit Sub
 End If
End Sub

Private Sub List3_DblClick()
 If List3.text = "" Then
  Exit Sub
 End If
reply = MsgBox("Do you want to remove " & List3.text & " from the list?", 36, "Remove")
 If reply = vbYes Then
  List1.RemoveItem List3.ListIndex
 Else
  Exit Sub
 End If
End Sub

Private Sub List4_DblClick()
 If List4.text = "" Then
  Exit Sub
 End If
reply = MsgBox("Do you want to remove " & List4.text & " from the list?", 36, "Remove")
 If reply = vbYes Then
  List1.RemoveItem List4.ListIndex
 Else
  Exit Sub
 End If
End Sub

Private Sub List5_DblClick()
 If List5.text = "" Then
  Exit Sub
 End If
reply = MsgBox("Do you want to remove " & List5.text & " from the list?", 36, "Remove")
 If reply = vbYes Then
  List1.RemoveItem List5.ListIndex
 Else
  Exit Sub
 End If
End Sub
