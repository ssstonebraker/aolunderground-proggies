VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "enter handle example · zb"
   ClientHeight    =   1035
   ClientLeft      =   2445
   ClientTop       =   1830
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3360
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "about"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "enter"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "handle:"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then  'checks to see if the enter handle field
'has text in it.
MsgBox "please enter a handle.", vbExclamation, "error" 'tells
'the user in a message box to enter a handle.
Else 'else

Call WriteToINI("enter handle example by zb", "handle", Text1.Text, "c:\windows\desktop\handleEX.ini")


han$ = GetFromINI("enter handle example by zb", "handle", "c:\windows\desktop\handleEX.ini")


Form2.Label1.Caption = han$ 'this will make a label the hanlde
'(you maybe wanna hide this in ur program.)


Form2.Show
Form1.Hide



End If 'ends if.
End Sub

Private Sub Command2_Click()
MsgBox "this is a quick little example i wipped up to show how to set a handle in an ini.", vbInformation, "about"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
'SetOnTop Me 'lil32.bas (by zb) call to set form on top.

If FileExists("c:\windows\desktop\handleEX.ini") = False Then 'checks to
'see if the ini is found.


Form1.Show 'shows the enter handle form.

Else

han$ = GetFromINI("enter handle example by zb", "handle", "c:\windows\desktop\handleEX.ini")


Form2.Label1.Caption = "" & han$ & ""  'this will make a label the hanlde
'(you maybe wanna hide this in ur program.)


Form2.Show  'shows the main form.
Form1.Hide 'unloads the enter handle form because the
'handle has already been enter'd.


End If 'ends if.




End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'unloads project.
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'enter key
If KeyCode = 13 Then
If Text1.Text = "" Then  'checks to see if the enter handle field
'has text in it.
MsgBox "please enter a handle.", vbExclamation, "error" 'tells
'the user in a message box to enter a handle.
Else 'else

Call WriteToINI("enter handle example by zb", "handle", Text1.Text, "c:\windows\desktop\handleEX.ini")


han$ = GetFromINI("enter handle example by zb", "handle", "c:\windows\desktop\handleEX.ini")


Form2.Label1.Caption = han$ 'this will make a label the hanlde
'(you maybe wanna hide this in ur program.)


Form2.Show
Form1.Hide



End If 'ends if.
Else
End If
End Sub
