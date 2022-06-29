' SEARCH.FRM
Option Explicit

Sub Command3D1_Click ()
If  Text1.Text = "" Then
MsgBox "You must enter in a text to search for!", 16
Exit Sub
End If
If  Command3D1.Caption = "Find" Then
Timer1.Enabled = True
Timer1.Interval = 1
End If
If  Command3D1.Caption = "Find Next" Then
Command3D1.Caption = "Find Next "
Else
End If
End Sub

Sub Command3D2_Click ()
search.WindowState = 1
End Sub

Sub Form_Load ()
extsub1050 search.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Form_Unload (p1 As Integer)
Timer1.Interval = 0
End Sub

Sub Timer1_Timer ()
Dim l0034 As Variant
Dim l003A As Variant
Dim l0040 As Variant
Dim l0046 As Variant
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As String
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0066 As Variant
l0034 = extfn1018("AOL FRAME25", 0&)
l003A = extfn10C0(l0034, "_AOL_TREE")
If  l003A = 0 Then
MsgBox "Please have your mailbox open", 16
Timer1.Interval = 0
GoTo L5C58E
End If
l0040 = extfn1088(l003A)
l0046 = extfnFE0(l003A, 1036, 0, 0)
l004A = 2948 / l0046
For l004E = 0 To l0046 - 1
If  Timer1.Interval = 0 Then GoTo L5C58E
l0052$ = String(255, 0)
l0056 = extfn11A0(l003A, 1034, l004E, l0052$)
l005A = InStr(1, l0052$, Text1.Text, 1)
If  l005A Then
l005E = extfnFE0(l003A, 1031, l004E, 0)
l0062 = l0062 + 1
Command3D1.Caption = "Find Next"
Do
DoEvents
If  Timer1.Interval = 0 Then GoTo L5C58E
Loop Until Command3D1.Caption <> "Find Next"
End If
l0066 = l0066 + l004A
Picture1.Line (x,y)(0, 0)-(l0066, 500), QBColor(3), BF
Next l004E
If  l0062 = "" Then
MsgBox "0 Mails were found"
Else
MsgBox l0062 + " Mails were found"
End If
Picture1.Cls
Timer1.Interval = 0
Command3D1.Caption = "Find"
L5C58E:
End Sub
