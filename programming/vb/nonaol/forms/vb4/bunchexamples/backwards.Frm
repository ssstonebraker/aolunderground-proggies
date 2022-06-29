VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Whats your name?"
   ClientHeight    =   390
   ClientLeft      =   2745
   ClientTop       =   1665
   ClientWidth     =   1800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   795
   Left            =   2685
   LinkTopic       =   "Form1"
   ScaleHeight     =   390
   ScaleWidth      =   1800
   Top             =   1320
   Width           =   1920
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


Private Sub Form_Load()
MsgBox "please enter your name and press enter"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'Iz representing eXcel 2001

If KeyAscii = 13 Then ' waits for the user to press enter [remember you have to end an if statement]
a = Text1.Text 'sets text1 to 'a' so you wont have to type text1.text, you can just use a
b = Len(a) ' counts the amount of characters in text1 (or a)
For i = 1 To b 'theres that changing variable i again
lastchar = Mid(a, b, 1) ' Mid(stringvar, start, length)
'stringval is your text1
'start is telling where vb to start, here its at the end
'1 is how many characters to go back
this = this + lastchar 'stores each character
b = b - 1 'counts back 1 to record the next one
Next ' goes back to i ;D
MsgBox a & " backwards is " & this ' displays text1(a) and some words plus
Text1.Text = "" ' clears text1
End If ' ends the if :)
'now isnt that easy?
End Sub


