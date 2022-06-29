' UNLOAD.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
l001C = extfn1018("AOL FRAME25", 0&)
If  l001C > 0 Then
MsgBox "Please unload America Online!", 16
Exit Sub
End If
l001C = Mid(Text1.Text, Len(Text1.Text), 1)
If  l001C = "\" Then
l001C = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
Else
l001C = Text1.Text
End If
On Error Resume Next
Open l001C + "\tool\manager.aol" For Binary As #1
If  Err Then
Close #1
MsgBox "That directory doesn't exist!", 16
Exit Sub
End If
Close #1
On Error Resume Next
Name l001C + "\tool\aolsux.aol" As l001C + "\tool\aolsux.bak"
If  Err Then
MsgBox "Master.aol is already unloaded!", 16
Else
MsgBox "Master.aol has been unloaded succesfully!", 16
End If
End Sub

Sub Command3D2_Click ()
Dim l0026 As Variant
l0026 = extfn1018("AOL FRAME25", 0&)
If  l0026 > 0 Then
MsgBox "Please unload America Online!", 16
Exit Sub
End If
l0026 = Mid(Text1.Text, Len(Text1.Text), 1)
If  l0026 = "\" Then
l0026 = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
Else
l0026 = Text1.Text
End If
On Error Resume Next
Open l0026 + "\tool\manager.aol" For Binary As #1
If  Err Then
Close #1
MsgBox "That directory doesn't exist!", 16
Exit Sub
End If
Close #1
On Error Resume Next
Name l0026 + "\tool\aolsux.bak" As l0026 + "\tool\aolsux.aol"
If  Err Then
MsgBox "Master.aol is already loaded!", 16
Else
MsgBox "Master.aol has been loaded succesfully!", 16
End If
End Sub

Sub Form_Load ()
extsub1050 unloads.hWnd, -1, 0, 0, 0, 0, 3
End Sub
