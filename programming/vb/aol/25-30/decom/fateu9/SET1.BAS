' SET1.FRM
Option Explicit

Sub Form_Load ()
Dim l001A As Variant
For l001A = 33 To 255
List1.AddItem Chr(l001A)
Next l001A
MsgBox "Click once on the character that you wish to copy to the clipboard!", 64
extsub1050 set1.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub List1_Click ()
Clipboard.Clear
Clipboard.SetText List1.List(List1.ListIndex)
End Sub
