Attribute VB_Name = "pwguard"
'pwguard.bas · by zb

'info:
'made in·         visual basic 6.0 pro  [32 bit]
'works in·         visual basic 32 bit
'completed·      may 27, 1999

'umm, this bas isn't to complicated. oh well, it works, vist
'my site for new vb files, programs, etc....  http://come.to/zbo
'oh yeah, this is 100% my coding.


'contacts:

'e-mail·         xzbx@hotmail.com
'web site·      http://come.to/zbo
'aim sn·        aim  zb


'here is the bas...


Public Const ENTER_KEY = 13

Sub PW_Guard(Password As String, Texty As TextBox, NextForm As Form, CorrectMsg As String, WrongMsg As String, FormToUnload As Form)
'put this in the enter button

'example:
'Call PW_Guard("Your Password", Text1, Form1, "you have entered the correct password!", "you have entered the incorrect password.", Form1)

Pvv$ = Password
If Texty.Text = "" Then
'add a message if ya want that indicates that nothing is in the text box.
Else
If Texty.Text = Pvv$ Then
MsgBox CorrectMsg, vbInformation, "correct"
Unload FormToUnload
NextForm.Show
Else
MsgBox WrongMsg, vbExclamation, "incorrect"
End If
End If
End Sub

Sub Set_PW_Box(TextBox As TextBox)
'put this in the form load event.

'example:
'Call Set_PW_Box (Text1)   -  (in the form load event.)

TextBox.PasswordChar = "*"
End Sub


'-zb


