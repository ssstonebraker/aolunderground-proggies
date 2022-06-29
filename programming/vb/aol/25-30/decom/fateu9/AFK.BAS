' AFK.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
Dim l002E As Variant
Dim l0032 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l0052 As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l002E = Mid(l0026$, 10, 10)
l0032 = InStr(l002E, "!")
If  l0032 Then
l002E = Mid(l002E, 1, l0032 - 1)
End If
If  Text1.Text = "" Then
MsgBox "Please enter in the reason of AFK!", 16
Exit Sub
End If
If  Command3D1.Caption = "Start" Then
Command3D1.Caption = "Stop"
l0044 = fnA68("«‹- -– Fate Ultra Away From the Keyboard Bot –-- -›»")
l0048 = fnA68("«‹- -– Reason: " + Text1.Text + " –-- -›»")
DoEvents
If  Check3D1.Value = True Then
l0052 = fnA68("«‹- -– Type " + Chr(34) + "/" + l002E + Chr(34) + " and The Message –-- -›»")
End If
If  Check3D1.Value = True Then
Timer2.Enabled = True
Timer2.Interval = 1
End If
Timer1.Enabled = True
Timer1.Interval = 60000&
Else
If  Check3D1.Value = True Then
Timer2.Enabled = False
End If
l0044 = fnA68("«‹- -– Fate Ultra - No Longer AFK –-- -›»")
Command3D1.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
Text2.Text = "0"
End If
End Sub

Sub Command3D2_Click ()
afkb.WindowState = 1
End Sub

Sub Command3D3_Click ()
List1.Clear
List2.Clear
Text3.Text = ""
End Sub

Sub Form_Load ()
Check3D1.Value = True
Text2.Text = "0"
extsub1050 afkb.hWnd, -1, 0, 0, 0, 0, 3
If  afkb.Caption <> fnC8() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub List1_Click ()
Dim l0080 As Variant
l0080 = List1.ListIndex
Text3.Text = List2.List(l0080)
End Sub

Sub List1_DblClick ()
List2.RemoveItem List1.ListIndex
List1.RemoveItem List1.ListIndex
End Sub

Sub Timer1_Timer ()
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As String
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
Dim l00A2 As Variant
Dim l00A6 As Variant
If  Timer1.Interval = 0 Then GoTo 10
l0084 = extfn1018("AOL FRAME25", 0&)
l0088 = extfn10F8(l0084, "Welcome")
l008C$ = String(30, 0)
l008E = extfnFA8(l0088, l008C$, 250)
If  l008E <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
l0092 = Mid(l008C$, 10, 10)
l0096 = InStr(l0092, "!")
If  l0096 Then
l0092 = Mid(l0092, 1, l0096 - 1)
End If
l009A = Text2.Text
l009A = l009A + 1
Text2.Text = l009A
l009A = fnA68("«‹- -– Fate Ultra Away From the Keyboard Bot –-- -›»")
DoEvents
DoEvents
l009E = fnA68("«‹- -– Away Since: " + Text2.Text + " Minute(s) –-- -›»")
DoEvents
DoEvents
l00A2 = fnA68("«‹- -– Reason: " + Text1.Text + " –-- -›»")
DoEvents
DoEvents
If  Check3D1.Value = True Then
l00A6 = fnA68("«‹- -– Type " + Chr(34) + "/" + l0092 + Chr(34) + " and The Message –-- -›»")
DoEvents
DoEvents
End If

10 :
End Sub

Sub Timer2_Timer ()
Dim l00B4 As Variant
Text4.Text = fn170()
l00B4 = fn1330()
End Sub
