' Form3.frm
Option Explicit

Sub Command3D1_Click ()
' Const c0024 = 3 ' &H3%
If  Text1.Text = "Spawn" Then
Me.Hide
Form1.Show
Exit Sub
Else :
End
End If
End Sub

Sub Form_Load ()
' Const c0024 = 3 ' &H3%
sub530 Me, c0024
subD80 Me
End Sub
