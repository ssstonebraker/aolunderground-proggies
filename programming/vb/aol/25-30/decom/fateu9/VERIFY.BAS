' VERIFY.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0026 As Variant
Dim l002A As Variant
Dim l002E As Variant
Dim l0032 As Variant
Dim l0036 As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As Variant
Dim l0046 As Variant
If  Text1.Text = "" Then
MsgBox "Please enter in their SN!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "Please enter in their string!", 16
Exit Sub
End If
Text1.Text = LCase(Text1.Text)
l0026 = 1
On Error Resume Next
For l002A = 1 To 16
If  Err Then
MsgBox "This is -NOT- MaGuS!", 16
Exit Sub
End If
l002E = Mid(Text2.Text, l002A, 1)
l0032 = Asc(l002E)
l0036 = Mid(Text1.Text, l0026, 1)
l0026 = l0026 + 1
If  l0026 > Len(Text1.Text) Then l0026 = 1
l003A = Asc(l0036)
l003E = l0032 + l003A
If  l003E > 255 Then
MsgBox "This is -NOT- MaGuS!", 16
Exit Sub
End If
l0042 = Chr(l003E)
l0046 = l0046 + l0042
DoEvents
Next l002A
If  l0046 = Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) + Chr(255) Then
MsgBox "This is the -REAL- MaGuS!", 16
Else
MsgBox "This is -NOT- MaGuS!", 16
End If
End Sub

Sub Command3D2_Click ()
Unload verify
End Sub

Sub Form_Load ()
extsub1050 verify.hWnd, -1, 0, 0, 0, 0, 3
If  verify.Caption <> fnC60() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub
