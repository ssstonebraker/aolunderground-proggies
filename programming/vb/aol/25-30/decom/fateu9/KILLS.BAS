' KILLS.FRM
Option Explicit

Sub Command3D1_Click ()
Dim l001C As Variant
Dim l0022 As Variant
Dim l0026 As String
Dim l002A As Variant
l001C = extfn1018("AOL FRAME25", 0&)
l0022 = extfn10F8(l001C, "Welcome")
l0026$ = String(30, 0)
l002A = extfnFA8(l0022, l0026$, 250)
If  l002A <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
killstatus.Show
kills.Hide
Timer1.Interval = 1
End Sub

Sub Form_Load ()
extsub1050 kills.hWnd, -1, 0, 0, 0, 0, 3
End Sub

Sub HScroll1_Change ()
Dim l003C As Variant
Label1.Caption = Str(l003C)
End Sub

Sub sub4420 ()
Dim l0046 As Variant
Label2.Caption = Str(l0046)
End Sub

Sub Timer1_Timer ()
Dim l005C As Variant
Dim l0060 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
If  Not IsNumeric(Text1.Text) Then
MsgBox "It must be a number!", 16
Timer1.Interval = 0
GoTo L1A0E6
End If
If  Not IsNumeric(Text2.Text) Then
MsgBox "It must be a number!", 16
Timer1.Interval = 0
GoTo L1A0E6
End If
If  Text1.Text <= 0 Then
MsgBox "The value must range between 1 and 1000!", 16
GoTo L1A0E6
End If
If  Text1.Text >= 1001 Then
MsgBox "The value must range between 1 and 1000!", 16
GoTo L1A0E6
End If
If  Timer1.Interval = 0 Then GoTo L1A0E6
l005C = Int(Text1.Text)
For l0060 = 1 To l005C
killstatus.Label1.Caption = "Sending"
If  Timer1.Interval = 0 Then GoTo L1A0E6
l0066 = fnA30("———————————————F    ———————————————A    ———————————————T    ———————————————E    ")
DoEvents
l006A = fnA30("———————————————F    ———————————————A    ———————————————T    ———————————————E    ")
DoEvents
l006E = fnA30("———————————————F    ———————————————A    ———————————————T    ———————————————E    ")
DoEvents
l0072 = fnA30("———————————————F    ———————————————A    ———————————————T    ———————————————E    ")
l0076 = Timer
l007A = Int(Text2.Text) * .1#
While Timer - l0076 < l007A
If  Timer1.Interval = 0 Then GoTo L1A0E6
killstatus.Label1.Caption = "Pausing"
DoEvents
Wend
Next l0060
L1A0E6:
Timer1.Interval = 0
kills.Show
Unload killstatus
End Sub


Sub sub9980 ()
Unload kills
End Sub
