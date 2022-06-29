' LAMESENS.FRM
Option Explicit

Sub Command3D1_Click ()
If  Command3D1.Caption = "Start" Then
Command3D1.Caption = "Stop"
Label1.Caption = Text3.Text + " Seconds"
Timer2.Interval = 1
Timer1.Interval = 1000
Else
Command3D1.Caption = "Start"
Timer2.Interval = 0
Timer1.Interval = 0
End If
End Sub

Sub Form_Load ()
extsub1050 Lamesense.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub Timer1_Timer ()
Dim l003C As Variant
Dim l0040 As Variant
Dim l0044 As Variant
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l0060 As Variant
Dim l0064 As Variant
If  Timer1.Interval = 0 Then
Label1.Caption = "0 Seconds"
GoTo 123
End If
Label1.Caption = Trim(Str(Int(Mid(Label1.Caption, 1, 2)) - 1)) + " Seconds"
If  Label1.Caption = "0 Seconds" Then
Command3D1.Caption = "Start"
l003C = ""
l0040 = 0
Timer1.Interval = 0
For l0044 = 0 To List1.ListCount - 1
l004E = InStr(List1.List(l0044), "-")
l0052 = Mid(List1.List(l0044), 1, l004E - 1)
l0056 = Mid(List1.List(l0044), l004E + 1, Len(List1.List(l0044)))
If  Int(l0056) > l0040 Then
l0040 = Int(l0056)
l003C = l0052
End If
Next l0044
l0044 = extfn1018("AOL FRAME25", 0&)
l004E = extfn10C0(l0044, "_AOL_LISTBOX")
l0052 = extfnFE0(l004E, 1036, 0, 0)
l0056 = 100 / l0052
l0060 = l0056 * List1.ListCount
l0060 = 100 - Int(l0060)
l0064 = Trim(Str(Fix(l0060))) + "%"
If  l003C = "" Then l003C = "Nobody!"
If  List1.ListCount = 0 Then l0064 = "100%"
l004E = fnA68("«‹- -– Fate Ultra Lamer Sensor –-- -›»")
l004E = fnA68("«‹- -– Lamest Person: " + l003C + " –-- -›»")
l004E = fnA68("«‹- -– This room is " + l0064 + " Lamer free –-- -›»")
Timer1.Interval = 0
GoTo 123
End If
123 :
End Sub

Sub Timer2_Timer ()
Dim l0074 As Variant
Text1.Text = fn170()
l0074 = fn4B8()
End Sub
