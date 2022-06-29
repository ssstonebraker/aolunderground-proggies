' PW.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then MsgBox "Type in a Password", 16, "Error": Exit Sub
If  UCase$(Text1) = "YUBAN OWNS ME" Then
Unload Me
TOSer.Show
Else
MsgBox "Sorry You Bastard.  Maybe Next Time.", 16, "Execution": Exit Sub
Text1.Text = ""
End If
End Sub

Sub Form_Load ()
subCD0 Me
End Sub

Sub Text1_KeyPress (KeyAscii As Integer)
If  KeyAscii = 13 Then
If  Text1.Text = "" Then MsgBox "Type in a Password", 16, "Error": Exit Sub
If  UCase$(Text1) = "YUBAN OWNS ME" Then
Unload Me
TOSer.Show
Else
MsgBox "Sorry You Bastard.  Maybe Next Time.", 16, "Execution": Exit Sub
Text1.Text = ""
End If
End If
End Sub
