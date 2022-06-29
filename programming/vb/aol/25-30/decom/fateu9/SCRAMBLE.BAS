' SCRAMBLE.FRM
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
Dim l004A As Variant
If  Text1.Text = "" Then
MsgBox "Please enter in a word to scramble!", 16
Exit Sub
End If
Text7.Text = ""
l0026 = Text1.Text
1 :
l002A = InStr(l0026, Chr(32))
If  l002A Then
l002E = Mid(l0026, 1, l002A - 1)
l0026 = Mid(l0026, l002A + 1, Len(l0026))
l0032 = 1
Else
l002E = l0026
l0036 = 1
l0032 = 1
End If
l003A = l002E
For l003E = 1 To Len(l003A)
l0042 = Mid(l003A, l003E, 1)
Randomize
l0046 = Int(Len(l003A) * Rnd) + 1
Randomize
l004A = Mid(l003A, l0046, 1)
Mid(l003A, l003E, 1) = l004A
Mid(l003A, l0046, 1) = l0042
Next l003E
If  l0032 = 1 Then
If  l0036 = 1 Then
Text7.Text = Text7.Text + l003A
GoTo 5
End If
Text7.Text = Text7.Text + l003A + Chr(32)
Else
Text7.Text = l003A
End If
If  l0036 = 1 Then GoTo 5
If  l0032 = 1 Then
l0032 = 0
GoTo 1
End If
5 :
End Sub

Sub Command3D2_Click ()
Dim l0050 As Variant
Dim l0056 As Variant
Dim l005A As String
Dim l005E As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l007C As Variant
l0050 = extfn1018("AOL FRAME25", 0&)
l0056 = extfn10F8(l0050, "Welcome")
l005A$ = String(30, 0)
l005E = extfnFA8(l0056, l005A$, 250)
If  l005E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "Please enter in a word to scramble!", 16
Exit Sub
End If
If  Text7.Text = "" Then
MsgBox "Please scramble the word before begining!", 16
Exit Sub
End If
If  Command3D2.Caption = "Start" Then
Command3D2.Caption = "Stop"
l006E = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(83) + Chr(99) + Chr(114) + Chr(97) + Chr(109) + Chr(98) + Chr(108) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l0072 = fnA68(fn330() + "Word: " + Text7.Text + fn368())
DoEvents
l007C = fnA68(fn330() + "Clue: " + Text5.Text + fn368())
DoEvents
Timer1.Enabled = True
Timer2.Enabled = True
Timer1.Interval = 1
Timer2.Interval = 60000&
Else
Text1.Enabled = True
Text5.Enabled = True
Command3D2.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Interval = 0
Timer2.Enabled = False
End If
End Sub

Sub Command3D3_Click ()
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l00A0 As Variant
Dim l00AA As Variant
l008C = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(83) + Chr(99) + Chr(114) + Chr(97) + Chr(109) + Chr(98) + Chr(108) + Chr(101) + Chr(114) + Chr(32) + Chr(83) + Chr(99) + Chr(111) + Chr(114) + Chr(101) + Chr(32) + Chr(66) + Chr(111) + Chr(97) + Chr(114) + Chr(100) + Chr(58) + fn368())
l0090 = l0090 + 1
For l0094 = 0 To List2.ListCount - 1
If  l0090 = 4 Then
l0090 = 0
l00A0 = fn13A0(2.7#)
End If
l00AA = fnA68("«‹- -– " + List2.List(l0094) + " - " + List1.List(l0094) + " point(s) –-- -›»")
l0090 = l0090 + 1
DoEvents
Next l0094
End Sub

Sub Command3D4_Click ()
List2.Clear
List1.Clear
List3.Clear
End Sub

Sub Command3D5_Click ()
scram.WindowState = 1
End Sub

Sub Command3D6_Click ()
Unload scram
End Sub

Sub Form_Load ()
extsub1050 scram.hWnd, -1, 0, 0, 0, 0, 3
If  scram.Caption <> fn988() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub Timer1_Timer ()
Dim l00C4 As Variant
If  Timer1.Interval = 0 Then GoTo 4
Text3.Text = fn170()
l00C4 = fn9C0()
4 :
End Sub

Sub Timer2_Timer ()
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
If  Timer2.Interval = 0 Then GoTo 3
l00C8 = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(83) + Chr(99) + Chr(114) + Chr(97) + Chr(109) + Chr(98) + Chr(108) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l00CC = fnA68(fn330() + "Word: " + Text7.Text + fn368())
DoEvents
l00D0 = fnA68(fn330() + "Clue: " + Text5.Text + fn368())
DoEvents
3 :
End Sub
