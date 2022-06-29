' CUSTOMKI.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l0022 As Variant
Dim l0028 As Variant
Dim l002C As String
Dim l0030 As Variant
Dim l0034 As Variant
Dim l003A As Variant
If  Text1.Text = "" Then
MsgBox "Please enter in something to scroll!", 16
Exit Sub
End If
l0022 = extfn1018("AOL FRAME25", 0&)
l0028 = extfn10F8(l0022, "Welcome")
l002C$ = String(30, 0)
l0030 = extfnFA8(l0028, l002C$, 250)
If  l0030 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
For l0034 = 1 To 4
l003A = fnA68(Chr(13) + Chr(10) + Text1)
Next l0034
End Sub

Sub Form_Load ()
extsub1050 customkill.hWnd, -1, 0, 0, 0, 0, 3
End Sub
