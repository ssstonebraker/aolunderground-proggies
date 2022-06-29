VERSION 5.00
Begin VB.Form window 
   Caption         =   "Window hider"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   Icon            =   "window.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "New Caption"
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Changet its caption"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close it"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show it"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Caption"
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide it"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declares our variable *window* as a long variable, or a window.
Dim window As Long

Private Sub Command1_Click()
'This is setting the value of the variable window as having the caption we type in
'to the text box.
window& = FindWindow(vbNullString, "" + Text1.text + "")
'This is saying Show the window *window* as 0, or not visible.
Call ShowWindow(window, 0)
End Sub

Private Sub Command2_Click()
'This is setting the value of the variable window as having the caption we type in
'to the text box.
window& = FindWindow(vbNullString, "" + Text1.text + "")
'This is saying Show the window *window* as 1, or visible.
Call ShowWindow(window, 1)
End Sub

Private Sub Command3_Click()
'This is setting the value of the variable window as having the caption we type in
'to the text box.
window& = FindWindow(vbNullString, "" + Text1.text + "")
'This is sending a message to the window *window* saying "Use the mouse to close it."
Call SendMessageLong(window, WM_CLOSE, 0, 0)
End Sub

Private Sub Command4_Click()
'This is setting the value of the variable window as having the caption we type in
'to the text box.
window& = FindWindow(vbNullString, "" + Text1.text + "")
'This is using the WinSetText in the bas It is saying change the text on the window *window*
'to the text we put in textbox #2
Call WinSetText(window, "" + Text2.text + "")
End Sub

Private Sub Form_Load()

End Sub
