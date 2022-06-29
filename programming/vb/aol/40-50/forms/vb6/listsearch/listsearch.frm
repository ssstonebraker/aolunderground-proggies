VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "make mail list"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "clear"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "find"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1530
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim z As Long 'defines the variable
List2.Clear 'clears list for new search
For z = 0 To List1.ListCount - 1 'this tells it to check the item on a list
If InStr(LCase(List1.List(z)), LCase(Text1)) Then List2.AddItem (List1.List(z)) 'this checks for the string
Next z 'this tells it to goto the next item
End Sub

Private Sub Command2_Click()
List1.AddItem (Text1) 'adds item to list
Text1 = "" 'clears text
End Sub

Private Sub Command3_Click()
List1.Clear 'clears lists
List2.Clear
End Sub

Private Sub Command4_Click()
    Dim X As Long, Y As Long, z As Long
    List1.Clear
    List2.Clear
MailOpenFlash 'opens mail box
Do
X = List1.ListCount
Pause 2
Y = List1.ListCount
Loop Until X = Y 'waits for mail to load by continously checking the numbers of mails
Call MailListFlash(List1, True) 'adds mail to a list
For z = 0 To List1.ListCount - 1 'starts going through items on a list
List2.AddItem (List1.List(z)) 'adds each item in list 1 to list2
If List2.ListCount = 500 Then Call MailSend(AOLUser(), "-^v^--• silver mail list", ListToString(List1)) 'checks to see if list2.list count is 500.  if it is it mails it to you.  this is done so mail doesnt take forever to load.
Next z 'goes to next list item
Call MailSend(AOLUser(), "-^v^--• silver mail list", ListToString(List1)) 'this sends list ifunder 500 mails or sends left overs
'just like a mass imer heh?
End Sub

