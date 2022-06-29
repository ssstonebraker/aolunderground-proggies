' ROOMRUN.FRM
Option Explicit
Dim m001A As Variant

Sub Command3D1_Click ()
Dim l0032 As Variant
Dim l0038 As Variant
Dim l003C As String
Dim l0040 As Variant
Dim l0044 As Variant
Dim l0048 As Variant
Dim l004C As Variant
If  Text1.Text = "" Then
MsgBox "Please enter a room to run!", 16
Exit Sub
End If
l0032 = extfn1018("AOL FRAME25", 0&)
l0038 = extfn10F8(l0032, "Welcome")
l003C$ = String(30, 0)
l0040 = extfnFA8(l0038, l003C$, 250)
If  l0040 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
m001A = Text1.Text
Do Until 2 > 3
l0044 = InStr(m001A, Chr(32))
If  l0044 Then
l0048 = Mid(m001A, 1, l0044 - 1)
l004C = Mid(m001A, l0044 + 1, Len(m001A))
m001A = l0048 + l004C
Else
GoTo 70
End If
Loop
70 :
m001A = UCase(m001A)
If  fn288() = m001A Then Exit Sub
If  Command3D1.Caption = "Start" Then
Command3D1.Caption = "Stop"
Timer1.Enabled = True
Timer1.Interval = 1
Timer2.Enabled = True
Timer2.Interval = 1000
Else
Command3D1.Caption = "Start"
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Interval = 0
Timer2.Enabled = False
End If
End Sub

Sub Command3D2_Click ()
roomrun.WindowState = 1
End Sub

Sub Command3D3_Click ()
Unload roomrun
End Sub

Sub Form_Load ()
extsub1050 roomrun.hWnd, -1, 0, 0, 0, 0, 3
If  roomrun.Caption <> fn950() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
End Sub

Sub Text1_Change ()
Label3.Caption = "Checking: " + Text1.Text
End Sub

Sub Timer1_Timer ()
Dim l0070 As Variant
Dim l0074 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008E As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00BA As Variant
Dim l00BE As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D6 As Variant
l0070 = ""
L3AF2C:
If  Timer1.Interval = 0 Then GoTo L3B4DA
l0074 = extfn1018("AOL FRAME25", 0&)
l007A = fn5D0("&Go To", "Keyword...")
L3AFC8:
DoEvents
If  Timer1.Interval = 0 Then GoTo L3B4DA
l007E = extfn10F8(l0074, "Keyword")
If  l007E = 0 Then GoTo L3AFC8
l0084 = extfn10C0(l007E, "_AOL_EDIT")
l0088 = extfn10C0(l007E, "_AOL_ICON")
l008E = extfn11A0(l0084, 12, 0, "aol://2719:2-2-" + m001A + l0070)
l0094 = fn1480(l0088)
l0098 = UCase(Mid(m001A, 1, 1))
l009C = LCase(Mid(m001A, 2, Len(m001A)))
l00A0 = l0098 + l009C
Label3.Caption = "Checking: " + l00A0 + l0070
Do Until 2 > 3
If  Timer1.Interval = 0 Then GoTo L3B4DA
l00A4 = extfn1018("#32770", "America Online")
If  l00A4 > 0 Then
l00A8 = extfn10F8(l00A4, "OK")
l00AC = fn1480(l00A8)
If  l0070 = "" Then
l0070 = "2"
Else
l0070 = Trim(Str(Int(l0070) + 1))
End If
l00B0 = fn288()
If  l00B0 = m001A + l0070 Then
l0070 = ""
End If
GoTo L3AF2C
End If
l00B0 = fn288()
If  l00B0 = m001A + l0070 Then
l00BA = Label2.Caption
If  Int(Label2.Caption) > 1 Then
l00BE = "s"
Else
l00BE = ""
End If
l00C4 = fn13A0(.7#)
Label3.Caption = "Checking: " + Text1.Text
l00C8 = Mid(l00B0, 1, 1)
l00CC = Mid(l00B0, 2, Len(l00B0))
l00D0 = LCase(l00CC)
l00B0 = l00C8 + l00D0
l00D6 = fnA68("«‹- -– Fate Ultra Found " + l00B0 + " In " + l00BA + " Second" + l00BE + " –-- -›»")
GoTo L3B4DA
End If
If  Timer1.Interval = 0 Then GoTo L3B4DA
DoEvents
Loop
L3B4DA:
Timer1.Interval = 0
Timer1.Enabled = False
Timer2.Interval = 0
Timer2.Enabled = False
Label2.Caption = "0"
Command3D1.Caption = "Start"
End Sub

Sub Timer2_Timer ()
If  Timer2.Interval = 0 Then GoTo 50
Label2.Caption = Int(Label2.Caption) + 1
50 :
End Sub
