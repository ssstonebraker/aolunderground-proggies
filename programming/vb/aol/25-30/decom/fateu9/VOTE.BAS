' VOTE.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  Text1.Text = Text2.Text Then
MsgBox "Agree and Disagree must be different!", 16
Exit Sub
End If
If  Text1.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text2.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text3.Text = "" Then
MsgBox "You must enter in something to vote!", 16
Exit Sub
End If
If  Text4.Text = "" Then
MsgBox "You must enter in when the results will be scrolled!", 16
Exit Sub
End If
If  Not IsNumeric(Text4.Text) Then
MsgBox "You must enter in a numeric number for when the results will be scrolled!", 16
Exit Sub
End If
If  Timer1.Interval = 0 Then
l0052 = fnA30(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(86) + Chr(111) + Chr(116) + Chr(101) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l0056 = fnA30(fn330() + "Question: " + Text1.Text + fn368())
DoEvents
l005A = fnA30(fn330() + "Agree: """ + Text2.Text + """" + fn368())
DoEvents
l005A = fnA30(fn330() + "Disagree: """ + Text3.Text + """" + fn368())
DoEvents
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label3.Caption = "Yes: 0"
Label5.Caption = "No: 0"
List1.Clear
Timer1.Enabled = True
Timer2.Enabled = True
Timer1.Interval = 1
Timer2.Interval = 60000&
Command3D1.Caption = "&Stop Bot"
Else
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Interval = 0
Timer2.Enabled = False
Command3D1.Caption = "&Start Bot"
End If
End Sub

Sub Command3D2_Click ()
Label3.Caption = "Yes: 0"
List2.Clear
End Sub

Sub Command3D3_Click ()
Label5.Caption = "No: 0"
List3.Clear
End Sub

Sub Command3D4_Click ()
vote.WindowState = 1
End Sub

Sub Form_Load ()
If  vote.Caption <> fn1560() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
extsub1050 vote.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l00A4 As Variant
Text5.Text = fn170()
l00A4 = fn1528()
End Sub

Sub Timer2_Timer ()
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Text4.Text = Text4.Text - 1
If  Text4.Text = "0" Then
l00A8 = fnA30(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(86) + Chr(111) + Chr(116) + Chr(101) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l00AC = fnA30(fn330() + "Question: " + Text1.Text + fn368())
DoEvents
l00B0 = Mid(Label3.Caption, 6, Len(Label3.Caption))
l00B4 = Mid(Label5.Caption, 5, Len(Label5.Caption))
l00B8 = fnA30(fn330() + "Final Results: Agree: " + l00B0 + " /Disagree: " + l00B4 + fn368())
DoEvents
Command3D1.Caption = "&Start Bot"
Timer1.Interval = 0
Timer2.Interval = 0
GoTo 20
End If
l00A8 = fnA30(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(86) + Chr(111) + Chr(116) + Chr(101) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + fn368())
DoEvents
l00AC = fnA30(fn330() + "Question: " + Text1.Text + fn368())
DoEvents
l00B8 = fnA30(fn330() + "Agree: """ + Text2.Text + """" + fn368())
DoEvents
l00B8 = fnA30(fn330() + "Disagree: """ + Text3.Text + """" + fn368())
DoEvents
List2.Clear
List3.Clear
GoTo 30
20 :
Timer2.Interval = 0
30 :
End Sub
