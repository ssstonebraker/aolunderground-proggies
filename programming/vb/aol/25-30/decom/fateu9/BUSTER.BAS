' BUSTER.FRM
Option Explicit
Dim m001A As Variant
Dim m003A As Variant

Sub Combo1_Click ()
Dim l0050 As Variant
Dim l0054 As String
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l006C As Variant
Dim l0070 As Variant
Dim l0074 As String
Dim l0076 As Variant
Dim l007A As String
Dim l007C As Variant
l0050 = Combo1.Text
On Error Resume Next
If  l0050 = "Private" Then
m003A = "1"
Open "fate.ini" For Binary As #1
l0054$ = String(LOF(1), 0)
Get #1, 1, l0054$
Close #1
l0056 = InStr(1, l0054$, "[Public Room]", 1)
l005A = Mid(l0054$, 1, l0056 - 5)
l005E = InStr(1, l005A, "[Private Room]", 1)
l0062 = Mid(l005A, l005E + 16, Len(l005A))
text1.Clear
Do Until 2 > 3
l006C = InStr(l0062, Chr(10))
If  l006C = 0 Then GoTo 50
l0070 = Mid(l0062, 1, l006C - 2)
If  l0070 = "" Then GoTo 567
text1.AddItem l0070
567 :
l0062 = Mid(l0062, l006C + 1, Len(l0062))
DoEvents
Loop
50 :
If  l0062 = "" Then
Else
text1.AddItem l0062
End If
text1.Text = text1.List(0)
End If
If  l0050 = "Member" Then
m003A = "2"
Open "fate.ini" For Binary As #1
l0074$ = String(LOF(1), 0)
Get #1, 1, l0074$
Close #1
l005E = InStr(1, l0074$, "[Member Room]", 1)
l0062 = Mid(l0074$, l005E + 15, Len(l0074$))
l0076 = InStr(1, l0062, "[FAC]", 1)
l0062 = Mid(l0062, 1, l0076 - 5)
text1.Clear
Do Until 2 > 3
l006C = InStr(l0062, Chr(10))
If  l006C = 0 Then GoTo 40
l0070 = Mid(l0062, 1, l006C - 2)
If  l0070 = "" Then GoTo 123
text1.AddItem l0070
123 :
l0062 = Mid(l0062, l006C + 1, Len(l0062))
DoEvents
Loop
40 :
If  l0062 = "" Then GoTo 1231
text1.AddItem l0062
1231 :
text1.Text = text1.List(0)
End If
If  l0050 = "Public" Then
m003A = "3"
Open "fate.ini" For Binary As #1
l007A$ = String(LOF(1), 0)
Get #1, 1, l007A$
Close #1
l0056 = InStr(1, l007A$, "[Member Room]", 1)
On Error Resume Next
l005A = Mid(l007A$, 1, l0056 - 5)
If  Err Then
Open "fate.ini" For Output As #1
l007C = Chr(13) + Chr(10)
Print #1, "[Private Room]" + l007C + l007C + "[Public Room]" + l007C + l007C + "[Member Room]" + l007C + l007C
Close #1
Exit Sub
End If
l005E = InStr(1, l005A, "[Public Room]", 1)
l0062 = Mid(l005A, l005E + 13, Len(l005A))
text1.Clear
Do Until 2 > 3
l006C = InStr(l0062, Chr(10))
If  l006C = 0 Then GoTo 60
l0070 = Mid(l0062, 1, l006C - 2)
If  l0070 = "" Then GoTo 345
text1.AddItem l0070
345 :
l0062 = Mid(l0062, l006C + 1, Len(l0062))
DoEvents
Loop
60 :
If  l0062 = "" Then GoTo 456
text1.AddItem l0062
456 :
text1.Text = text1.List(0)
End If
If  Err Then sub3A0
End Sub

Sub sub99F8 ()
End Sub

Sub Command3D1_Click ()
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As String
Dim l008A As Variant
Dim l008E As Variant
Dim l0092 As Variant
Dim l0096 As Variant
Dim l009A As Variant
Dim l009E As Variant
On Error Resume Next
If  text1.Text = "" Then GoTo 56
SetAttr "fate.ini", 0
For l0080 = 0 To text1.ListCount - 1
l0084 = InStr(1, text1.List(l0080), text1.Text, 1)
If  l0084 Then
If  Len(text1.List(l0080)) = Len(text1.Text) Then GoTo 56
End If
Next l0080
text1.AddItem text1.Text
If  m003A = "1" Then
Open "fate.ini" For Binary As #1
l0088$ = String(LOF(1), 0)
Get #1, 1, l0088$
Close #1
For l008A = 0 To text1.ListCount - 1
l008E = l008E + text1.List(l008A) + Chr(13) + Chr(10)
Next l008A
l0092 = InStr(1, l0088$, "[Private Room]", 1)
l0096 = Mid(l0088$, 1, l0092 + 15)
l009A = InStr(1, l0088$, "[Public Room]", 1)
l009E = Mid(l0088$, l009A - 2, Len(l0088$))
Open "fate.ini" For Output As #1
Print #1, l0096 + l008E + l009E
Close #1
End If
If  m003A = "3" Then
Open "fate.ini" For Binary As #1
l0088$ = String(LOF(1), 0)
Get #1, 1, l0088$
Close #1
For l008A = 0 To text1.ListCount - 1
l008E = l008E + text1.List(l008A) + Chr(13) + Chr(10)
Next l008A
l0092 = InStr(1, l0088$, "[Public Room]", 1)
l0096 = Mid(l0088$, 1, l0092 + 14)
l009A = InStr(1, l0088$, "[Member Room]", 1)
l009E = Mid(l0088$, l009A - 2, Len(l0088$))
Open "fate.ini" For Output As #1
Print #1, l0096 + l008E + l009E
Close #1
End If
If  m003A = "2" Then
Open "fate.ini" For Binary As #1
l0088$ = String(LOF(1), 0)
Get #1, 1, l0088$
Close #1
For l008A = 0 To text1.ListCount - 1
l008E = l008E + text1.List(l008A) + Chr(13) + Chr(10)
Next l008A
l0092 = InStr(1, l0088$, "[Member Room]", 1)
l0096 = Mid(l0088$, 1, l0092 + 14)
l009A = InStr(1, l0088$, "[FAC]", 1)
l009E = Mid(l0088$, l009A - 2, Len(l0088$))
Open "fate.ini" For Output As #1
Print #1, l0096 + l008E + l009E
Close #1
End If

56 :
If  Err Then sub3A0
End Sub

Sub Command3D2_Click ()
Dim l00A2 As Variant
Dim l00A6 As Variant
Dim l00AA As String
Dim l00AC As Variant
Dim l00B0 As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
If  text1.Text = "" Then GoTo 67
On Error Resume Next
SetAttr "fate.ini", 0
For l00A2 = 0 To text1.ListCount - 1
l00A6 = InStr(1, text1.List(l00A2), text1.Text, 1)
If  l00A6 Then
If  Len(text1.List(l00A2)) = Len(text1.Text) Then text1.RemoveItem l00A2
End If
Next l00A2
If  m003A = "1" Then
Open "fate.ini" For Binary As #1
l00AA$ = String(LOF(1), 0)
Get #1, 1, l00AA$
Close #1
For l00AC = 0 To text1.ListCount - 1
l00B0 = l00B0 + text1.List(l00AC) + Chr(13) + Chr(10)
Next l00AC
l00B4 = InStr(1, l00AA$, "[Private Room]", 1)
l00B8 = Mid(l00AA$, 1, l00B4 + 15)
l00BC = InStr(1, l00AA$, "[Public Room]", 1)
l00C0 = Mid(l00AA$, l00BC - 2, Len(l00AA$))
Open "fate.ini" For Output As #1
Print #1, l00B8 + l00B0 + l00C0
Close #1
End If
If  m003A = "3" Then
Open "fate.ini" For Binary As #1
l00AA$ = String(LOF(1), 0)
Get #1, 1, l00AA$
Close #1
For l00AC = 0 To text1.ListCount - 1
l00B0 = l00B0 + text1.List(l00AC) + Chr(13) + Chr(10)
Next l00AC
l00B4 = InStr(1, l00AA$, "[Public Room]", 1)
l00B8 = Mid(l00AA$, 1, l00B4 + 14)
l00BC = InStr(1, l00AA$, "[Member Room]", 1)
l00C0 = Mid(l00AA$, l00BC - 2, Len(l00AA$))
Open "fate.ini" For Output As #1
Print #1, l00B8 + l00B0 + l00C0
Close #1
End If
If  m003A = "2" Then
Open "fate.ini" For Binary As #1
l00AA$ = String(LOF(1), 0)
Get #1, 1, l00AA$
Close #1
For l00AC = 0 To text1.ListCount - 1
l00B0 = l00B0 + text1.List(l00AC) + Chr(13) + Chr(10)
Next l00AC
l00B4 = InStr(1, l00AA$, "[Member Room]", 1)
l00B8 = Mid(l00AA$, 1, l00B4 + 14)
l00BC = InStr(1, l00AA$, "[FAC]", 1)
l00C0 = Mid(l00AA$, l00BC - 2, Len(l00AA$))
Open "fate.ini" For Output As #1
Print #1, l00B8 + l00B0 + l00C0
Close #1
End If
text1.Text = text1.List(0)
67 :
If  Err Then sub3A0
End Sub

Sub Command3D3_Click ()
Dim l00C6 As Variant
Dim l00CC As Variant
Dim l00D0 As String
Dim l00D4 As Variant
l00C6 = extfn1018("AOL FRAME25", 0&)
l00CC = extfn10F8(l00C6, "Welcome")
l00D0$ = String(30, 0)
l00D4 = extfnFA8(l00CC, l00D0$, 250)
If  l00D4 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  text1 = "" Then
MsgBox "Please Enter In The Name Of The Room To Bust In", 16
Exit Sub
End If
If  okay.Caption = "Start" Then
okay.Caption = "Stop"
Timer1.Interval = 1
Else
okay.Caption = "Start"
Timer1.Interval = 0
End If
End Sub

Sub Form_Load ()
Dim l00F0 As String
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0106 As Variant
Combo1.AddItem "Private"
Combo1.AddItem "Member"
Combo1.AddItem "Public"
Combo1.ListIndex = 0
Combo2.AddItem "Menu"
Combo2.AddItem "Keyword"
Combo2.ListIndex = 0
extsub1050 buster.hWnd, -1, 0, 0, 0, 0, 3
If  buster.Caption <> fn218() Then
MsgBox "Program contains unauthorized changes." + Chr(13) + Chr(10) + "Terminating now!", 16
End
End If
m003A = "1"
On Error Resume Next
Open "fate.ini" For Binary As #1
l00F0$ = String(LOF(1), 0)
Get #1, 1, l00F0$
Close #1
l00F2 = InStr(1, l00F0$, "[Public Room]", 1)
l00F6 = Mid(l00F0$, 1, l00F2 - 5)
l00FA = InStr(1, l00F6, "[Private Room]", 1)
l00FE = Mid(l00F6, l00FA + 16, Len(l00F6))
text1.Clear
Do Until 2 > 3
l0102 = InStr(l00FE, Chr(10))
If  l0102 = 0 Then GoTo 300
l0106 = Mid(l00FE, 1, l0102 - 2)
If  l0106 = "" Then GoTo 678
text1.AddItem l0106
678 :
l00FE = Mid(l00FE, l0102 + 1, Len(l00FE))
DoEvents
Loop
300 :
If  l00FE = "" Then
Else
text1.AddItem l00FE
End If
text1.Text = text1.List(0)
If  Err Then sub3A0
End Sub

Sub HScroll1_Change ()
End Sub

Sub sub66C8 ()
End Sub

Sub sub6700 ()
End Sub

Sub okay_Click ()
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As String
Dim l0114 As Variant
Dim l011A As Variant
Dim l0120 As Variant
Dim l0124 As String
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
l010A = extfn1018("AOL FRAME25", 0&)
l010E = extfn10F8(l010A, "Welcome")
l0112$ = String(30, 0)
l0114 = extfnFA8(l010E, l0112$, 250)
If  l0114 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  text1.Text = "" Then
MsgBox "Please enter in the name of the room to bust in!", 16
Exit Sub
End If
If  okay.Caption = "Start" Then
l010A = extfn1018("AOL FRAME25", 0&)
l010E = extfn10F8(l010A, "AOL Live!")
If  l010E = 0 Then GoTo 45
l011A = extfn1088(l010E)
l0114 = extfn10C0(l011A, "_AOL_EDIT")
l0120 = extfn1088(l0114)
l0124$ = String(100, 0)
l0126 = extfnFA8(l0120, l0124$, 250)
l012A = l0124$
l0124$ = ""
l012E = InStr(l012A, Chr(0))
l012A = Mid(l012A, 1, l012E - 1)
Do Until 2 > 3
l0132 = InStr(l012A, Chr(32))
If  l0132 Then
l0136 = Mid(l012A, 1, l0132 - 1)
l013A = Mid(l012A, l0132 + 1, Len(l012A))
l012A = l0136 + l013A
Else
GoTo 70
End If
Loop
70 :
l012A = UCase(l012A)

45 :
m001A = text1.Text
Do Until 2 > 3
l0132 = InStr(m001A, Chr(32))
If  l0132 Then
l0136 = Mid(m001A, 1, l0132 - 1)
l013A = Mid(m001A, l0132 + 1, Len(m001A))
m001A = l0136 + l013A
Else
GoTo 80
End If
Loop
80 :
m001A = UCase(m001A)
If  m003A <> "1" Then
If  "TOWNSQUARE-" + m001A = l012A Then
Exit Sub
End If
Else
If  m001A = l012A Then
Exit Sub
End If
End If
okay.Caption = "Stop"
If  Combo2.Text = "Keyword" Then
Timer2.Enabled = True
Timer2.Interval = 1
Else
Timer1.Enabled = True
Timer1.Interval = 1
End If
Else
okay.Caption = "Start"
If  Combo2.Text = "Keyword" Then
Timer2.Interval = 0
Timer2.Enabled = False
Else
Timer1.Interval = 0
Timer1.Enabled = False
End If
End If
End Sub

Sub sub6770 ()
End Sub

Sub text1_KeyPress (KeyAscii As Integer)
Dim l0146 As Variant
Dim l014A As Variant
Dim l014E As Variant
Dim l0152 As String
Dim l0154 As Variant
Dim l0158 As Variant
Dim l015C As Variant
Dim l0160 As String
Dim l0162 As Variant
Dim l0166 As Variant
Dim l016A As Variant
Dim l016E As Variant
Dim l0172 As Variant
l0146 = Chr(KeyAscii)
If  l0146 = Chr(13) Then
l014A = extfn1018("AOL FRAME25", 0&)
l014E = extfn10F8(l014A, "Welcome")
l0152$ = String(30, 0)
l0154 = extfnFA8(l014E, l0152$, 250)
If  l0154 <= 7 Then
MsgBox "An error occurred: Not Signed On!", 16
Exit Sub
End If
If  text1 = "" Then
MsgBox "Please Enter In The Name Of The Room To Bust In", 16
Exit Sub
End If
If  okay.Caption = "Start" Then
l014A = extfn1018("AOL FRAME25", 0&)
l014E = extfn10F8(l014A, "AOL Live!")
If  l014E = 0 Then GoTo 450
l0158 = extfn1088(l014E)
l0154 = extfn10C0(l0158, "_AOL_EDIT")
l015C = extfn1088(l0154)
l0160$ = String(l015C, 0)
l0162 = extfnFA8(l015C, l0160$, 250)
l0166 = l0160$
l016A = InStr(l0166, Chr(0))
l0166 = Mid(l0166, 1, l016A - 1)
Do Until 2 > 3
l0146 = InStr(l0166, Chr(32))
If  l0146 Then
l016E = Mid(l0166, 1, l0146 - 1)
l0172 = Mid(l0166, l0146 + 1, Len(l0166))
l0166 = l016E + l0172
Else
GoTo 700
End If
Loop
700 :
l0166 = UCase(l0166)

450 :
m001A = text1.Text
Do Until 2 > 3
l0146 = InStr(m001A, Chr(32))
If  l0146 Then
l016E = Mid(m001A, 1, l0146 - 1)
l0172 = Mid(m001A, l0146 + 1, Len(m001A))
m001A = l016E + l0172
Else
GoTo 800
End If
Loop
800 :
m001A = UCase(m001A)
If  m001A = l0166 Then
Exit Sub
End If
okay.Caption = "Stop"
If  Combo2.Text = "Keyword" Then
Timer2.Enabled = True
Timer2.Interval = 1
Else
Timer1.Enabled = True
Timer1.Interval = 1
End If
Else
okay.Caption = "Start"
If  Combo2.Text = "Keyword" Then
Timer2.Interval = 0
Timer2.Enabled = False
Else
Timer1.Interval = 0
Timer1.Enabled = False
End If
End If
End If
End Sub

Sub Timer1_Timer ()
Dim l0176 As Variant
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Variant
Dim l018A As Variant
Dim l0190 As Variant
Dim l0194 As Variant
Dim l0198 As Variant
Dim l019E As Variant
Dim l01A2 As Variant
Dim l01AC As String
Dim l01AE As Variant
Dim l01B2 As Variant
Dim l01B6 As Variant
Dim l01BA As Variant
Dim l01BE As Variant
Dim l01C2 As Variant
Dim l01C6 As Variant
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D2 As Variant
Dim l01D6 As Variant
Dim l01DA As Variant
Dim l01DE As String
Dim l01E0 As Variant
Dim l01E4 As Variant
Dim l01E8 As Variant
Dim l01EC As Variant
Dim l01F0 As Variant
Dim l01F4 As Variant
Dim l01F8 As Variant
Dim l01FE As Variant
Dim l0206 As Variant
Dim l020A As Variant
Dim l020E As Variant
Dim l0212 As Variant
Dim l0216 As Variant
Dim l021A As Variant
Dim l021E As Variant
l0176 = extfn1018("AOL FRAME25", 0&)
l017C = fn5D0("&Go To", "Edit Go To Menu")

10 :
DoEvents
l0180 = extfn1018("_AOL_MODAL", "Favorite Places")
If  l0180 = 0 Then GoTo 10
l0184 = extfn10C0(l0180, "_AOL_EDIT")
l018A = extfn11A0(l0184, 12, 0, text1)
l0190 = extfn1168(l0184, 2)
If  m003A = "1" Then l0194 = extfn11A0(l0190, 12, 0, "aol://2719:2-2-" + text1.Text)
If  m003A = "2" Then l0194 = extfn11A0(l0190, 12, 0, "aol://2719:61-2-" + text1.Text)
If  m003A = "3" Then l0194 = extfn11A0(l0190, 12, 0, "aol://2719:21-2-" + text1.Text)
l0198 = extfn10F8(l0180, "Save Changes")
l019E = fn1480(l0198)
20 :
DoEvents
l01A2 = extfn1018("_AOL_MODAL", "Favorite Places")
If  l01A2 > 0 Then GoTo 20
stat.Caption = "1"
30 :
l01AC$ = text1.Text
l01AE = fn5D0("&Go To", l01AC$)
Do Until 2 > 3
If  m003A = "3" Then
l01B2 = extfn10F8(l0176, "Untitled")
If  l01B2 > 0 Then
l01B6 = extfn10F8(l01B2, "NO")
l01BA = fn1480(l01B6)
stat.Caption = stat.Caption + 1
DoEvents
If  Timer1.Interval = 0 Then GoTo L4AE1E
GoTo 30
End If
End If
If  m003A = "2" Then GoTo 90
If  m003A = "1" Then
90 :
l01BE = extfn1018("#32770", "America Online")
If  l01BE > 0 Then
l01C2 = extfn10F8(l01BE, "OK")
l01C6 = fn1480(l01C2)
stat.Caption = stat.Caption + 1
If  Timer1.Interval = 0 Then GoTo L4AE1E
DoEvents
GoTo 30
End If
End If
l01CA = extfn1018("AOL FRAME25", 0&)
l01CE = extfn10F8(l01CA, "AOL Live!")
If  l01CE > 0 Then
l01D2 = extfn1088(l01CE)
l01D6 = extfn10C0(l01D2, "_AOL_EDIT")
l01DA = extfn1088(l01D6)
l01DE$ = String(255, 0)
l01E0 = extfnFA8(l01DA, l01DE$, 250)
l01E4 = l01DE$
l01E8 = InStr(l01E4, Chr(0))
l01E4 = Mid(l01E4, 1, l01E8 - 1)
Do Until 2 > 3
l01EC = InStr(l01E4, Chr(32))
If  l01EC Then
l01F0 = Mid(l01E4, 1, l01EC - 1)
l01F4 = Mid(l01E4, l01EC + 1, Len(l01E4))
l01E4 = l01F0 + l01F4
Else
GoTo 100
End If
Loop
100 :
l01E4 = UCase(l01E4)
If  m003A <> "1" Then
l01F8 = "TOWNSQUARE-"
Else
l01F8 = ""
End If
If  l01E4 = l01F8 + m001A Then
If  stat.Caption = "1" Then
l01FE = fn13A0(.7#)
l0206 = extfn11A0(l01D6, 12, 0, fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(66) + Chr(117) + Chr(115) + Chr(116) + Chr(32) + Chr(73) + Chr(110) + Chr(32) + Chr(84) + Chr(111) + Chr(111) + Chr(107) + Chr(32) + Chr(49) + Chr(32) + Chr(65) + Chr(116) + Chr(116) + Chr(101) + Chr(109) + Chr(112) + Chr(116) + fn368())
l020A = extfn1168(l01D6, 2)
l020E = fn1480(l020A)
DoEvents
GoTo L4AE1E
End If
l0212 = fn13A0(.7#)
l0216 = extfn11A0(l01D6, 12, 0, fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(66) + Chr(117) + Chr(115) + Chr(116) + Chr(32) + Chr(73) + Chr(110) + Chr(32) + Chr(84) + Chr(111) + Chr(111) + Chr(107) + Chr(32) + stat.Caption + Chr(32) + Chr(65) + Chr(116) + Chr(116) + Chr(101) + Chr(109) + Chr(112) + Chr(116) + Chr(115) + fn368())
l021A = extfn1168(l01D6, 2)
l021E = fn1480(l021A)
DoEvents
GoTo L4AE1E
End If
End If
DoEvents
If  Timer1.Interval = 0 Then GoTo L4AE1E
Loop
L4AE1E:
Timer1.Interval = 0
Timer1.Enabled = False
okay.Caption = "Start"
stat.Caption = "0"
End Sub


Sub Timer2_Timer ()
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
Dim l022E As Variant
Dim l0232 As Variant
Dim l0236 As Variant
Dim l023A As Variant
Dim l023E As Variant
Dim l0242 As Variant
Dim l0246 As Variant
Dim l024A As Variant
Dim l024E As Variant
Dim l0252 As Variant
Dim l0258 As Variant
Dim l025C As Variant
Dim l0260 As Variant
Dim l0264 As Variant
Dim l026A As Variant
stat.Caption = "1"
L4AE90:
l0222 = extfn1018("AOL FRAME25", 0&)
l0226 = fn5D0("&Go To", "Keyword...")
L4AF14:
DoEvents
If  Timer2.Interval = 0 Then GoTo L4B626
l022A = extfn10F8(l0222, "Keyword")
If  l022A = 0 Then GoTo L4AF14
l022E = extfn10C0(l022A, "_AOL_EDIT")
l0232 = extfn10C0(l022A, "_AOL_ICON")
If  m003A = "1" Then l0236 = extfn11A0(l022E, 12, 0, "aol://2719:2-2-" + text1.Text)
If  m003A = "2" Then l0236 = extfn11A0(l022E, 12, 0, "aol://2719:61-2-" + text1.Text)
If  m003A = "3" Then l0236 = extfn11A0(l022E, 12, 0, "aol://2719:21-2-" + text1.Text)
l023A = fn1480(l0232)
Do Until 2 > 3
If  Timer2.Interval = 0 Then GoTo L4B626
If  m003A = "1" Then
If  Timer2.Interval = 0 Then GoTo L4B626
l023E = extfn1018("#32770", "America Online")
If  l023E > 0 Then
l0242 = extfn10F8(l023E, "OK")
l0246 = fn1480(l0242)
stat.Caption = Trim(Str(Int(stat.Caption) + 1))
GoTo L4AE90
End If
End If
If  m003A = "2" Then
If  Timer2.Interval = 0 Then GoTo L4B626
l023E = extfn1018("#32770", "America Online")
If  l023E > 0 Then
l0242 = extfn10F8(l023E, "OK")
l0246 = fn1480(l0242)
stat.Caption = Trim(Str(Int(stat.Caption) + 1))
GoTo L4AE90
End If
End If
If  m003A = "3" Then
If  Timer2.Interval = 0 Then GoTo L4B626
l024A = extfn10F8(l0222, "Untitled")
If  l024A > 0 Then
l024E = extfn10F8(l024A, "NO")
l0252 = fn1480(l024E)
stat.Caption = Trim(Str(Int(stat.Caption) + 1))
GoTo L4AE90
End If
End If
l0258 = fn288()
If  m003A <> "1" Then
l025C = "TOWNSQUARE-"
Else
l025C = ""
End If
If  l0258 = l025C + m001A Then
If  stat.Caption = "1" Then
l0260 = ""
Else
l0260 = "s"
End If
l0264 = fn13A0(.6#)
l026A = fnA30("«‹- -– Fate Ultra Bust In Took " + Trim(stat.Caption) + " attempt" + l0260 + " –-- -›»")
GoTo L4B626
End If
If  Timer2.Interval = 0 Then GoTo L4B626
DoEvents
Loop
L4B626:
stat.Caption = "0"
Timer2.Interval = 0
okay.Caption = "Start"
End Sub

Sub sub9980 ()
Unload buster
End Sub
