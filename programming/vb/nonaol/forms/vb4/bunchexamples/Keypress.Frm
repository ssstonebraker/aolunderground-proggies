VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "KeyPress"
   ClientHeight    =   1470
   ClientLeft      =   3345
   ClientTop       =   1740
   ClientWidth     =   2010
   Height          =   1875
   Left            =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   2010
   Top             =   1395
   Width           =   2130
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   30
      TabIndex        =   1
      Top             =   75
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   1155
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
'ok here is how you take information from a textbox to a list box
'using the textbox's Key_Press function
'Iz Representing eXcel 2001
If KeyAscii = 13 Then 'waits for the user to press enter(Chr(13))
List1.AddItem Text1.Text ' Adds the text to a list box
Text1.Text = "" 'Clears the textbox
End If ' closes the IF statement
End Sub


