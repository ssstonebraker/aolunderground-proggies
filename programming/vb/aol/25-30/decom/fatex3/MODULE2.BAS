' Module2
Option Explicit

Function fn1640 () As Variant
Dim l0040 As Variant
Dim l0046 As Variant
Dim l004C As Variant
Dim l0050 As Variant
Dim l0056 As Variant
Dim l005C As String
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As String
Dim l0092 As String
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
Dim l00A0 As Variant
Dim l00A4 As Variant
Dim l00AA As Variant
Dim l00B4 As Variant
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
l0040 = 0
l0046 = findwindow("AOL FRAME25", 0&)
l004C = fn2F8()
l0050 = l004C
l0056 = Findchildbyclass(l0050, "_AOL_VIEW")
l005C$ = getapitext(l0056)
afkb.Text4.Text = l005C$
For l0062 = 0 To Len(afkb.Text4.Text) - 1
l0066 = Mid(afkb.Text4.Text, Len(afkb.Text4.Text) - l0062, 1)
If l0066 = Chr(10) Then Exit For
l006A = l0066 + l006A
DoEvents
Next l0062
l006E = afkb.Text4.Text
Do Until 2 > 3
l0072 = InStr(l006E, l006A)
If l0072 Then
afkb.Text5.Text = Trim(Str(Int(afkb.Text5.Text + 1)))
l006E = Mid(l006E, l0072 + Len(l006A) + 2, Len(l006E))
Else
Exit Do
End If
DoEvents
Loop
l0076 = findwindow("AOL FRAME25", 0&)
l007A = fn2F8()
l007E = l007A
l0082 = Findchildbyclass(l007E, "_AOL_VIEW")
l0088 = SendMessageByNum(l0082, 14, 0, 0)
101 :
If l0040 = 1 Then
l0088 = l008C
End If
Do Until 2 > 3
l007E = fn2F8()
l0082 = Findchildbyclass(l007E, "_AOL_VIEW")
l008C = SendMessageByNum(l0082, 14, 0, 0)
If l0088 <> l008C Then Exit Do
If afkb.Command3D1.Caption = "Start" Then GoTo L12AB6
If fn2F8() = 0 Then
MsgBox "Chat room has been closed, shutting down bot!", 16
If afkb.Check1.Value = 1 Then
afkb.Timer2.Enabled = False
End If
afkb.Command3D1.Caption = "Start"
afkb.Timer1.Interval = 0
afkb.Timer1.Enabled = False
afkb.Text2.Text = "0"
GoTo L12AB6
End If
DoEvents
Loop
l0090$ = getapitext(l0082)
afkb.Text4.Text = l0090$
l0092$ = afkb.Text4.Text
For l0094 = 1 To Int(afkb.Text5.Text)
DoEvents
l0098 = InStr(l0092$, l006A)
If l0098 Then
l0092$ = Mid(l0092$, l0098 + Len(l006A) + 2, Len(l0092$))
Else
GoTo 155
End If
Next l0094

155 :
l009C = l0092$

300 :
DoEvents
l00A0 = InStr(l009C, Chr(13) + Chr(10))
If l00A0 Then
l00A4 = Mid(l009C, 1, l00A0 - 1)
afkb.Text6.Text = l00A4
l0072 = InStr(1, l00A4, ":" + Chr(9) + "/" + fnA68() + " ", 1)
If l0072 Then
l006E = InStr(l00A4, ":")
If l006E = 0 Then GoTo 5
l0076 = Mid(l00A4, 1, l006E - 1)
If l0076 = fnA68() Then GoTo 5
l007A = Mid(l00A4, l0072 + 3 + Len(fnA68()), Len(l00A4))
For l007E = 0 To afkb.List1.ListCount - 1
l0082 = InStr(1, afkb.List1.List(l007E), l0076, 1)
If l0082 Then
If Len(afkb.List1.List(l007E)) = Len(l0076) Then GoTo 5
End If
Next l007E
afkb.List1.AddItem l0076
afkb.List2.AddItem l007A
afkb.Command3D4.Enabled = True
afkb.Label2.Caption = Trim(Str(afkb.List1.ListCount))
l00AA = Timeout(2.5)
l00B4 = fnAD8(fn448() + "<" + l0076 + ">" + " - Message saved." + fn480())
afkb.List1.Selected(0) = True
End If
5 :
l009C = Mid(l009C, l00A0 + 2, Len(l009C))
GoTo 300
Else
afkb.Text6.Text = l009C
l0072 = InStr(1, l009C, ":" + Chr(9) + "/" + fnA68() + " ", 1)
If l0072 Then
l006E = InStr(l009C, ":")
If l006E = 0 Then GoTo 4
l0076 = Mid(l009C, 1, l006E - 1)
If l0076 = fnA68() Then GoTo 4
l007A = Mid(l009C, l0072 + 3 + Len(fnA68()), Len(l009C))
For l007E = 0 To afkb.List1.ListCount - 1
l0082 = InStr(1, afkb.List1.List(l007E), l0076, 1)
If l0082 Then
If Len(afkb.List1.List(l007E)) = Len(l0076) Then GoTo 4
End If
Next l007E
afkb.List1.AddItem l0076
afkb.List2.AddItem l007A
afkb.Command3D4.Enabled = True
l00B8 = Timeout(2.5)
afkb.Label2.Caption = Trim(Str(afkb.List1.ListCount))
l00B4 = fnAD8(fn448() + "<" + l0076 + ">" + " - Message saved." + fn480())
afkb.List1.Selected(0) = True
End If
4 :
End If
l006A = afkb.Text6.Text
l00BC = afkb.Text4.Text
afkb.Text5.Text = "0"
Do Until 2 > 3
l00C0 = InStr(l00BC, l006A)
If l00C0 Then
afkb.Text5.Text = Trim(Str(Int(afkb.Text5.Text + 1)))
l00BC = Mid(l00BC, l00C0 + Len(l006A) + 2, Len(l00BC))
Else
Exit Do
End If
Loop
l0040 = 1
If afkb.Command3D1.Caption = "Start" Then GoTo L12AB6
DoEvents
GoTo 101

L12AB6:
afkb.Text4.Text = ""
afkb.Text5.Text = "0"
afkb.Text6.Text = ""
End Function

Function fn1678 () As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As String
Dim l00E2 As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As Variant
Dim l00FE As Variant
Dim l0102 As Variant
Dim l0106 As Variant
Dim l010A As Variant
Dim l010E As Variant
Dim l0112 As String
Dim l0114 As String
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As Variant
Dim l0142 As Variant
Dim l0146 As Variant
Dim l014A As Variant
l00C8 = 0
l00CC = findwindow("AOL FRAME25", 0&)
l00D0 = fn2F8()
l00D4 = l00D0
l00D8 = Findchildbyclass(l00D4, "_AOL_VIEW")
l00DC$ = getapitext(l00D8)
ECHO.text1.Text = l00DC$
For l00E2 = 0 To Len(ECHO.text1.Text) - 1
l00E6 = Mid(ECHO.text1.Text, Len(ECHO.text1.Text) - l00E2, 1)
If l00E6 = Chr(10) Then Exit For
l00EA = l00E6 + l00EA
DoEvents
Next l00E2
l00EE = ECHO.text1.Text
Do Until 2 > 3
l00F2 = InStr(l00EE, l00EA)
If l00F2 Then
ECHO.text2.Text = Trim(Str(Int(ECHO.text2.Text + 1)))
l00EE = Mid(l00EE, l00F2 + Len(l00EA) + 2, Len(l00EE))
Else
Exit Do
End If
DoEvents
Loop
l00F6 = findwindow("AOL FRAME25", 0&)
l00FA = fn2F8()
l00FE = l00FA
l0102 = Findchildbyclass(l00FE, "_AOL_VIEW")
l0106 = SendMessageByNum(l0102, 14, 0, 0)
1015 :
If l00C8 = 1 Then
l0106 = l010A
End If
Do Until 2 > 3
l00FE = fn2F8()
l0102 = Findchildbyclass(l00FE, "_AOL_VIEW")
l010A = SendMessageByNum(l0102, 14, 0, 0)
If l0106 <> l010A Then Exit Do
If ECHO.Command3D1.Caption = "Echo" Then GoTo L1366E
If fn2F8() = 0 Then
l010E = 0
ECHO.Command3D1.Caption = "Echo"
ECHO.Timer1.Enabled = False
ECHO.Timer1.Interval = 0
MsgBox "Chat window has been closed, shutting down bot!", 16
GoTo L1366E
End If
DoEvents
Loop
l0112$ = getapitext(l0102)
ECHO.text1.Text = l0112$
l0114$ = ECHO.text1.Text
For l0116 = 1 To Int(ECHO.text2.Text)
DoEvents
l011A = InStr(l0114$, l00EA)
If l011A Then
l0114$ = Mid(l0114$, l011A + Len(l00EA) + 2, Len(l0114$))
Else
GoTo 1555
End If
Next l0116

1555 :
l011E = l0114$

3005 :
DoEvents
If ECHO.Command3D1.Caption = "Echo" Then GoTo L1366E
l0122 = InStr(l011E, Chr(13) + Chr(10))
If l0122 Then
l0126 = Mid(l011E, 1, l0122 - 1)
ECHO.text3.Text = l0126
l00EE = InStr(l0126, ":")
If l00EE = 0 Then GoTo 45
l00F6 = Mid(l0126, 1, l00EE - 1)
If l00F6 = fnA68() Then GoTo 45
For l012A = 0 To ECHO.List1.ListCount - 1
l012E = InStr(1, ECHO.List1.List(l012A), l00F6, 1)
If l012E Then
l00FE = Mid(l0126, l00EE + 2, Len(l0126))
For l0132 = 0 To ECHO.List2.ListCount - 1
l0136 = InStr(1, ECHO.List2.List(l0132), l00FE, 1)
If l0136 Then GoTo 45
Next l0132
If ECHO.Check1.Value = 1 Then
l013A = l00FE
l00FE = ""
For l013E = 0 To Len(l013A) - 1
l0142 = Mid(l013A, Len(l013A) - l013E, 1)
l00FE = l00FE + l0142
Next l013E
End If
l0102 = fnAD8(l00FE)
End If
DoEvents
Next l012A

45 :
If ECHO.Command3D1.Caption = "Echo" Then GoTo L1366E
l011E = Mid(l011E, l0122 + 2, Len(l011E))
GoTo 3005
Else
ECHO.text3.Text = l011E
l00EE = InStr(l011E, ":")
If l00EE = 0 Then GoTo 47
l00F6 = Mid(l011E, 1, l00EE - 1)
If l00F6 = fnA68() Then GoTo 47
For l012A = 0 To ECHO.List1.ListCount - 1
l012E = InStr(1, ECHO.List1.List(l012A), l00F6, 1)
If l012E Then
l00FE = Mid(l011E, l00EE + 2, Len(l011E))
For l0132 = 0 To ECHO.List2.ListCount - 1
l0136 = InStr(1, ECHO.List2.List(l0132), l00FE, 1)
If l0136 Then GoTo 47
Next l0132
If ECHO.Check1.Value = 1 Then
l013A = l00FE
l00FE = ""
For l013E = 0 To Len(l013A) - 1
l0142 = Mid(l013A, Len(l013A) - l013E, 1)
l00FE = l00FE + l0142
Next l013E
End If
l0102 = fnAD8(l00FE)
End If
DoEvents
Next l012A
47 :
End If
l00EA = ECHO.text3.Text
l0146 = ECHO.text1.Text
ECHO.text2.Text = "0"
Do Until 2 > 3
l014A = InStr(l0146, l00EA)
If l014A Then
ECHO.text2.Text = Trim(Str(Int(ECHO.text2.Text + 1)))
l0146 = Mid(l0146, l014A + Len(l00EA) + 2, Len(l0146))
Else
Exit Do
End If
Loop
l00C8 = 1
If ECHO.Command3D1.Caption = "Echo" Then GoTo L1366E
DoEvents
GoTo 1015

L1366E:
ECHO.text1.Text = ""
ECHO.text2.Text = "0"
ECHO.text3.Text = ""
End Function

Function fn16E8 () As Variant
Dim l015C As Variant
Dim l0162 As Variant
Dim l0166 As String
Dim l016A As Variant
Dim l016E As Variant
Dim l0172 As Variant
Dim l017A As Variant
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As String
Dim l018A As Variant
Dim l018E As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As Variant
Dim l019E As Variant
Dim l01A2 As Variant
Dim l01A6 As Variant
Dim l01AA As Variant
Dim l01AE As Variant
Dim l01B2 As Variant
Dim l01B6 As Variant
Dim l01BC As Variant
l015C = findwindow("AOL FRAME25", 0&)
l0162 = Findchildbytitle(l015C, "Welcome")
l0166$ = String(30, 0)
l016A = GetWindowText(l0162, l0166$, 250)
l016E = Mid(l0166$, 10, 10)
l0172 = InStr(l016E, "!")
If l0172 Then
l016E = Mid(l016E, 1, l0172 - 1)
End If

320 :
If fightbot.Command3D2.Caption = "Start" Then GoTo L142EA
DoEvents
l017A = findwindow("AOL FRAME25", 0&)
l017E = fn2F8()
l017A = l017E
l017E = Findchildbyclass(l017A, "_AOL_VIEW")
If l017E = 0 Then
Exit Function
End If
l0182 = SendMessageByNum(l017E, 14, 0, 0)
If l0182 = Len(fightbot.Text4.Text) Then GoTo 320
l0186$ = String(l0182, 0)
l018A = SendMessageByString(l017E, 13, l0182, l0186$)
l018E = Len(fightbot.Text4.Text)
l0192 = Len(l0186$) - l018E
On Error Resume Next
l0196 = Right(l0186$, l0192)
l0196 = Mid(l0196, 3, Len(l0196))
340 :
If fightbot.Command3D2.Caption = "Start" Then GoTo L142EA
l019A = InStr(l0196, Chr(10))
If l019A Then
l019E = Mid(l0196, 1, l019A - 2)
fightbot.Text5.Text = l019E
l01A2 = InStr(fightbot.Text5.Text, Chr(0))
If l01A2 Then
fightbot.Text5.Text = Mid(fightbot.Text5.Text, 1, l01A2 - 1)
End If
l01A6 = InStr(1, fightbot.Text5.Text, fightbot.Text3.Text, 1)
If l01A6 Then
l01AA = Mid(fightbot.Text5.Text, l01A6 - 2, 2)
If l01AA <> ":" + Chr(9) Then
GoTo 360
End If
l01AA = InStr(fightbot.Text5.Text, ":")
l01AE = Mid(fightbot.Text5.Text, 1, l01AA - 1)
l01B2 = Mid(fightbot.Text5.Text, l01AA + 2, Len(fightbot.Text5.Text))
If Len(l01B2) <> Len(fightbot.Text3.Text) Then GoTo 360
If l01AE = l016E Then GoTo 360
l01B6 = Timeout(.6)
If fightbot.Text1.Text = "" And fightbot.Text2.Text > "" Then
fightbot.Text1.Text = l01AE
l01BC = fnAA0(fn448() + l01AE + " will challenge " + fightbot.Text2.Text + fn480())
GoTo L142EA
End If
If fightbot.Text2.Text = "" And fightbot.Text1.Text > "" Then
fightbot.Text2.Text = l01AE
l01BC = fnAA0(fn448() + l01AE + " will challenge " + fightbot.Text1.Text + fn480())
GoTo L142EA
End If
If fightbot.Text1.Text = "" And fightbot.Text2.Text = "" Then
fightbot.Text1.Text = l01AE
l01BC = fnAA0(fn480() + l01AE + " will be the first fighter!" + fn480())
GoTo 360
End If
End If
360 :
l0196 = Mid(l0196, l019A + 1, Len(l0196))
GoTo 340
End If
fightbot.Text5.Text = l0196
If fightbot.Text5.Text <> "" Then fightbot.Text5.Text = fightbot.Text5.Text
l01A2 = InStr(fightbot.Text5.Text, Chr(0))
If l01A2 Then
fightbot.Text5.Text = Mid(fightbot.Text5.Text, 1, l01A2 - 1)
End If
l01A6 = InStr(1, fightbot.Text5.Text, fightbot.Text3.Text, 1)
If l01A6 Then
l01AA = Mid(fightbot.Text5.Text, l01A6 - 2, 2)
If l01AA <> ":" + Chr(9) Then
GoTo 350
End If
l01AA = InStr(fightbot.Text5.Text, ":")
l01AE = Mid(fightbot.Text5.Text, 1, l01AA - 1)
l01B2 = Mid(fightbot.Text5.Text, l01AA + 2, Len(fightbot.Text5.Text))
If Len(l01B2) <> Len(fightbot.Text3.Text) Then GoTo 350
If l01AE = l016E Then GoTo 350
l01B6 = Timeout(.6)
If fightbot.Text1.Text = "" And fightbot.Text2.Text > "" Then
l01BC = fnAA0(fn448() + l01AE + " will challenge " + fightbot.Text2.Text + fn480())
fightbot.Text1.Text = l01AE
GoTo L142EA
End If
If fightbot.Text2.Text = "" And fightbot.Text1.Text > "" Then
l01BC = fnAA0(fn448() + l01AE + " will challenge " + fightbot.Text1.Text + fn480())
fightbot.Text2.Text = l01AE
GoTo L142EA
End If
If fightbot.Text1.Text = "" And fightbot.Text2.Text = "" Then
fightbot.Text1.Text = l01AE
l01BC = fnAA0(fn448() + l01AE + " will be the first fighter!" + fn480())
GoTo 350
End If
End If
350 :
If fightbot.Command3D2.Caption = "Start" Then GoTo L142EA
fightbot.Text4.Text = l0186$
DoEvents
GoTo 320
L142EA:
fightbot.Command3D2.Caption = "Start"
fightbot.Timer2.Interval = 0
fightbot.Timer2.Enabled = False
fightbot.Text4.Text = ""
fightbot.Text5.Text = ""
End Function

Function fn1720 () As Variant
Dim l01C4 As Variant
Dim l01C8 As Variant
Dim l01CC As Variant
Dim l01D0 As Variant
Dim l01D4 As Variant
Dim l01D8 As String
Dim l01DE As Variant
Dim l01E2 As Variant
Dim l01E6 As Variant
Dim l01EA As Variant
Dim l01EE As Variant
Dim l01F2 As Variant
Dim l01F6 As Variant
Dim l01FA As Variant
Dim l01FE As Variant
Dim l0202 As Variant
Dim l0206 As Variant
Dim l020A As String
Dim l020C As String
Dim l020E As Variant
Dim l0212 As Variant
Dim l0216 As Variant
Dim l021A As Variant
Dim l021E As Variant
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
Dim l022E As Variant
Dim l0232 As Variant
Dim l0236 As Variant
Dim l023A As Variant
l01C4 = 0
l01C8 = findwindow("AOL FRAME25", 0&)
l01CC = fn2F8()
l01D0 = l01CC
l01D4 = Findchildbyclass(l01D0, "_AOL_VIEW")
l01D8$ = getapitext(l01D4)
getmm.Text1.Text = l01D8$
For l01DE = 0 To Len(getmm.Text1.Text) - 1
l01E2 = Mid(getmm.Text1.Text, Len(getmm.Text1.Text) - l01DE, 1)
If l01E2 = Chr(10) Then Exit For
l01E6 = l01E2 + l01E6
DoEvents
Next l01DE
l01EA = getmm.Text1.Text
Do Until 2 > 3
l01EE = InStr(l01EA, l01E6)
If l01EE Then
getmm.Text2.Text = Trim(Str(Int(getmm.Text2.Text + 1)))
l01EA = Mid(l01EA, l01EE + Len(l01E6) + 2, Len(l01EA))
Else
Exit Do
End If
DoEvents
Loop
l01F2 = findwindow("AOL FRAME25", 0&)
l01F6 = fn2F8()
l01FA = l01F6
l01FE = Findchildbyclass(l01FA, "_AOL_VIEW")
l0202 = SendMessageByNum(l01FE, 14, 0, 0)
1016 :
If l01C4 = 1 Then
l0202 = l0206
End If
Do Until 2 > 3
l01FA = fn2F8()
l01FE = Findchildbyclass(l01FA, "_AOL_VIEW")
l0206 = SendMessageByNum(l01FE, 14, 0, 0)
If l0202 <> l0206 Then Exit Do
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
If fn2F8() = 0 Then
getmm.Timer1.Interval = 0
getmm.Timer1.Enabled = False
getmm.Command3D1.Caption = "Start"
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo L150E6
End If
DoEvents
Loop
l020A$ = getapitext(l01FE)
getmm.Text1.Text = l020A$
l020C$ = getmm.Text1.Text
For l020E = 1 To Int(getmm.Text2.Text)
DoEvents
l0212 = InStr(l020C$, l01E6)
If l0212 Then
l020C$ = Mid(l020C$, l0212 + Len(l01E6) + 2, Len(l020C$))
Else
GoTo 1556
End If
Next l020E
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
1556 :
l0216 = l020C$

3006 :
DoEvents
l021A = InStr(l0216, Chr(13) + Chr(10))
If l021A Then
l021E = Mid(l0216, 1, l021A - 1)
getmm.Text3.Text = l021E
l01EA = InStr(l021E, ":")
If l01EA = 0 Then GoTo 11
l01F2 = Mid(l021E, 1, l01EA - 1)
l01F6 = Mid(l021E, l01EA + 2, Len(l021E))
l01FA = InStr(1, l01F6, "type", 1)
If l01FA Then
If l01F2 = fnA68() Then GoTo 11
12 :
l01FE = InStr(l01F6, Chr(34))
If l01FE Then
l0202 = Mid(l01F6, l01FE + 1, Len(l01F6))
l0206 = InStr(l0202, Chr(34))
If l0206 Then
l0222 = Mid(l0202, 1, l0206 - 1)
For l0226 = 0 To getmm.List1.ListCount - 1
l022A = InStr(1, getmm.List1.List(l0226), l01F2, 1)
If l022A Then GoTo 11
Next l0226
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
For l022E = 0 To getmm.List2.ListCount - 1
l0232 = InStr(1, l0222, getmm.List2.List(l022E), 1)
If l0232 Then GoTo 11
DoEvents
Next l022E
getmm.List1.AddItem l01F2
getmm.Command3D2.Enabled = True
getmm.Label2.Caption = Trim(Str(getmm.List1.ListCount))
l0226 = fnAD8(l0222)
End If
End If
Else
l01FA = InStr(1, l01F6, "say", 1)
If l01FA Then GoTo 12
End If
11 :
l0216 = Mid(l0216, l021A + 2, Len(l0216))
GoTo 3006
Else
getmm.Text3.Text = l0216
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
l01EA = InStr(l0216, ":")
If l01EA = 0 Then GoTo 13
l01F2 = Mid(l0216, 1, l01EA - 1)
l01F6 = Mid(l01F2, 1, 1)
If l01F6 = Chr(10) Then l01F2 = Mid(l01F2, 2, Len(l01F2))
l01F6 = Mid(l0216, l01EA + 2, Len(l0216))
l01FA = InStr(1, l01F6, "type", 1)
If l01FA Then
If l01F2 = fnA68() Then GoTo 13
14 :
l01FE = InStr(l01F6, Chr(34))
If l01FE Then
l0202 = Mid(l01F6, l01FE + 1, Len(l01F6))
l0206 = InStr(l0202, Chr(34))
If l0206 Then
l0222 = Mid(l0202, 1, l0206 - 1)
For l0226 = 0 To getmm.List1.ListCount - 1
l022A = InStr(1, getmm.List1.List(l0226), l01F2, 1)
If l022A Then GoTo 13
Next l0226
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
For l022E = 0 To getmm.List2.ListCount - 1
l0232 = InStr(1, l0222, getmm.List2.List(l022E), 1)
If l0232 Then GoTo 13
DoEvents
Next l022E
getmm.List1.AddItem l01F2
getmm.Command3D2.Enabled = True
getmm.Label2.Caption = Trim(Str(getmm.List1.ListCount))
l0226 = fnAD8(l0222)
End If
End If
Else
l01FA = InStr(1, l01F6, "say", 1)
If l01FA Then GoTo 14
End If
13 :
End If
l01E6 = getmm.Text3.Text
l0236 = getmm.Text1.Text
getmm.Text2.Text = "0"
Do Until 2 > 3
l023A = InStr(l0236, l01E6)
If l023A Then
getmm.Text2.Text = Trim(Str(Int(getmm.Text2.Text + 1)))
l0236 = Mid(l0236, l023A + Len(l01E6) + 2, Len(l0236))
Else
Exit Do
End If
Loop
l01C4 = 1
If getmm.Command3D1.Caption = "Start" Then GoTo L150E6
DoEvents
GoTo 1016

L150E6:
getmm.Text1.Text = ""
getmm.Text2.Text = "0"
getmm.Text3.Text = ""
End Function

Function fn1758 () As Variant
End Function

Function fn1790 () As Variant
Dim l0246 As Variant
Dim l024A As Variant
Dim l024E As Variant
Dim l0252 As Variant
Dim l0256 As Variant
Dim l025A As String
Dim l0260 As Variant
Dim l0264 As Variant
Dim l0268 As Variant
Dim l026C As Variant
Dim l0270 As Variant
Dim l0274 As Variant
Dim l0278 As Variant
Dim l027C As Variant
Dim l0280 As Variant
Dim l0284 As Variant
Dim l0288 As Variant
Dim l028C As String
Dim l028E As String
Dim l0290 As Variant
Dim l0294 As Variant
Dim l0298 As Variant
Dim l029C As Variant
Dim l02A0 As Variant
Dim l02A4 As Variant
Dim l02A8 As Variant
l0246 = 0
l024A = findwindow("AOL FRAME25", 0&)
l024E = fn2F8()
l0252 = l024E
l0256 = Findchildbyclass(l0252, "_AOL_VIEW")
l025A$ = getapitext(l0256)
Lamesense.Text1.Text = l025A$
For l0260 = 0 To Len(Lamesense.Text1.Text) - 1
l0264 = Mid(Lamesense.Text1.Text, Len(Lamesense.Text1.Text) - l0260, 1)
If l0264 = Chr(10) Then Exit For
l0268 = l0264 + l0268
DoEvents
Next l0260
l026C = Lamesense.Text1.Text
Do Until 2 > 3
l0270 = InStr(l026C, l0268)
If l0270 Then
Lamesense.Text2.Text = Trim(Str(Int(Lamesense.Text2.Text + 1)))
l026C = Mid(l026C, l0270 + Len(l0268) + 2, Len(l026C))
Else
Exit Do
End If
DoEvents
Loop
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
l0274 = findwindow("AOL FRAME25", 0&)
l0278 = fn2F8()
l027C = l0278
l0280 = Findchildbyclass(l027C, "_AOL_VIEW")
l0284 = SendMessageByNum(l0280, 14, 0, 0)
1014 :
If l0246 = 1 Then
l0284 = l0288
End If
Do Until 2 > 3
l027C = fn2F8()
l0280 = Findchildbyclass(l027C, "_AOL_VIEW")
l0288 = SendMessageByNum(l0280, 14, 0, 0)
If l0284 <> l0288 Then Exit Do
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
If fn2F8() = 0 Then
Lamesense.Command3D1.Caption = "Start"
Lamesense.Timer2.Interval = 0
Lamesense.Timer1.Interval = 0
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo L15CA2
End If
DoEvents
Loop
l028C$ = getapitext(l0280)
Lamesense.Text1.Text = l028C$
l028E$ = Lamesense.Text1.Text
For l0290 = 1 To Int(Lamesense.Text2.Text)
DoEvents
l0294 = InStr(l028E$, l0268)
If l0294 Then
l028E$ = Mid(l028E$, l0294 + Len(l0268) + 2, Len(l028E$))
Else
GoTo 1554
End If
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
Next l0290

1554 :
l0298 = l028E$

3004 :
DoEvents
l029C = InStr(l0298, Chr(13) + Chr(10))
If l029C Then
l02A0 = Mid(l0298, 1, l029C - 1)
Lamesense.Text4.Text = l02A0
l026C = InStr(l02A0, ":")
If l026C = 0 Then GoTo 236
l0274 = Mid(l02A0, 1, l026C - 1)
If l0274 = fnA68() Then GoTo 236
For l027C = 0 To Lamesense.List1.ListCount - 1
l0280 = InStr(1, Lamesense.List1.List(l027C), "-", 1)
l0284 = Mid(Lamesense.List1.List(l027C), 1, l0280 - 1)
If l0284 = l0274 Then
l0288 = Int(Mid(Lamesense.List1.List(l027C), l0280 + 1, Len(Lamesense.List1.List(l027C))))
l0288 = l0288 + 1
Lamesense.List1.List(l027C) = l0274 + "-" + Trim(Str(l0288))
GoTo 236
End If
Next l027C
Lamesense.List1.AddItem l0274 + "-1"

236 :
l0298 = Mid(l0298, l029C + 2, Len(l0298))
GoTo 3004
Else
Lamesense.Text4.Text = l0298
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
l026C = InStr(l0298, ":")
If l026C = 0 Then GoTo 235
l0274 = Mid(l0298, 1, l026C - 1)
If l0274 = fnA68() Then GoTo 235
For l027C = 0 To Lamesense.List1.ListCount - 1
l0280 = InStr(1, Lamesense.List1.List(l027C), "-", 1)
l0284 = Mid(Lamesense.List1.List(l027C), 1, l0280 - 1)
If l0284 = l0274 Then
l0288 = Int(Mid(Lamesense.List1.List(l027C), l0280 + 1, Len(Lamesense.List1.List(l027C))))
l0288 = l0288 + 1
Lamesense.List1.List(l027C) = l0274 + "-" + Trim(Str(l0288))
GoTo 235
End If
Next l027C
Lamesense.List1.AddItem l0274 + "-1"
235 :
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
End If
l0268 = Lamesense.Text4.Text
l02A4 = Lamesense.Text1.Text
Lamesense.Text2.Text = "0"
Do Until 2 > 3
l02A8 = InStr(l02A4, l0268)
If l02A8 Then
Lamesense.Text2.Text = Trim(Str(Int(Lamesense.Text2.Text + 1)))
l02A4 = Mid(l02A4, l02A8 + Len(l0268) + 2, Len(l02A4))
Else
Exit Do
End If
If Lamesense.Command3D1.Caption = "Start" Then GoTo L15CA2
Loop
l0246 = 1
DoEvents
GoTo 1014
L15CA2:
Lamesense.Text1.Text = ""
Lamesense.Text2.Text = "0"
Lamesense.Text4.Text = ""
End Function

Function fn17C8 (ByVal pv02B0 As Integer, ByVal pv02B2 As Integer) As Variant
Dim l02B4 As Variant
Dim l02B8 As Variant
Dim l02BC As Variant
Dim l02C0 As Variant
Dim l02C4 As Variant
l02B4 = findwindow("AOL FRAME25", 0&)
l02B8 = getmenu(l02B4)
l02BC = getsubmenu(l02B8, pv02B0)
l02C0 = getmenuitemid(l02BC, pv02B2)
l02C4 = SendMessageByNum(l02B4, 273, l02C0, 0)
End Function

Function fn1800 (p02CC As String, p02CE As String) As Variant
Dim l02D0 As Variant
Dim l02D4 As Variant
Dim l02D8 As Variant
Dim l02DC As Variant
Dim l02E0 As String
Dim l02E4 As Variant
Dim l02EA As Variant
Dim l02EE As Variant
Dim l02F4 As Variant
Dim l02F8 As String
Dim l02FA As Variant
Dim l02FE As Variant
Dim l0302 As Variant
l02D0 = findwindow("AOL FRAME25", 0&)
l02D4 = getmenu(l02D0)
l02D8 = 0
Do Until 2 > 3
l02DC = getsubmenu(l02D4, l02D8)
l02E0$ = String(50, 0)
l02E4 = getmenustring(l02D4, l02DC, l02E0$, 100, 100)
l02E0$ = fnC28(l02E0$)
If Len(l02E0$) = 0 And l02E0$ = "" Then
MsgBox "An error occured: Can't find America Online Menu: " + p02CC + ". Please restart America Online in order for Fate to work.", 16
Exit Function
End If
If l02E0$ = p02CC Then GoTo 51
l02D8 = l02D8 + 1
DoEvents
Loop
51 :
l02EA = 0
Do Until 2 > 3
l02EE = getmenuitemid(l02DC, l02EA)
l02F4 = getmenuitemcount(l02DC) - 1
l02F8$ = String(55, 0)
l02FA = getmenustring(l02DC, l02EE, l02F8$, 100, 100)
l02F8$ = fnC28(l02F8$)
If Len(l02F8$) = 0 And l02F8$ = "" And l02EA > l02F4 Then
MsgBox "An error occured: Can't find America Online SubMenu: " + p02CE + " under " + p02CC + ". Please restart America Online in order for Fate to work.", 16
Exit Function
End If
l02FE = InStr(l02F8$, Chr(9))
If l02FE Then l02F8$ = Mid(l02F8$, 1, l02FE - 1)
If l02F8$ = p02CE Then GoTo 52
l02EA = l02EA + 1
DoEvents
Loop
52 :
l0302 = fn17C8(l02D8, l02EA)
End Function

Function fn1838 () As Variant
Dim l030A As Variant
Dim l030E As Variant
Dim l0312 As Variant
Dim l0316 As Variant
Dim l031A As Variant
Dim l031E As Variant
Dim l0322 As String
Dim l0328 As Variant
Dim l032C As Variant
Dim l0330 As Variant
Dim l0334 As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
Dim l0354 As String
Dim l0356 As String
Dim l0358 As Variant
Dim l035C As Variant
Dim l0360 As Variant
Dim l0364 As Variant
Dim l0368 As Variant
Dim l036C As Variant
Dim l0370 As Variant
l030A = 0
l030E = 0
l0312 = findwindow("AOL FRAME25", 0&)
l0316 = fn2F8()
l031A = l0316
l031E = Findchildbyclass(l031A, "_AOL_VIEW")
l0322$ = getapitext(l031E)
area51.Text3.Text = l0322$
For l0328 = 0 To Len(area51.Text3.Text) - 1
l032C = Mid(area51.Text3.Text, Len(area51.Text3.Text) - l0328, 1)
If l032C = Chr(10) Then Exit For
l0330 = l032C + l0330
DoEvents
Next l0328
l0334 = area51.Text3.Text
Do Until 2 > 3
l0338 = InStr(l0334, l0330)
If l0338 Then
area51.Text4.Text = Trim(Str(Int(area51.Text4.Text + 1)))
l0334 = Mid(l0334, l0338 + Len(l0330) + 2, Len(l0334))
Else
Exit Do
End If
DoEvents
Loop
l033C = findwindow("AOL FRAME25", 0&)
l0340 = fn2F8()
l0344 = l0340
l0348 = Findchildbyclass(l0344, "_AOL_VIEW")
l034C = SendMessageByNum(l0348, 14, 0, 0)
L164C0: :
If l030E = 1 Then
l034C = l0350
End If
Do Until 2 > 3
l0344 = fn2F8()
l0348 = Findchildbyclass(l0344, "_AOL_VIEW")
l0350 = SendMessageByNum(l0348, 14, 0, 0)
If l034C <> l0350 Then Exit Do
If fn2F8() = 0 Then
End If
DoEvents
Loop
l0354$ = getapitext(l0348)
area51.Text3.Text = l0354$
l0356$ = area51.Text3.Text
For l0358 = 1 To Int(area51.Text4.Text)
DoEvents
l035C = InStr(l0356$, l0330)
If l035C Then
l0356$ = Mid(l0356$, l035C + Len(l0330) + 2, Len(l0356$))
Else
GoTo L16662
End If
Next l0358

L16662: :
l0360 = l0356$

L16678: :
DoEvents
l0364 = InStr(l0360, Chr(13) + Chr(10))
If l0364 Then
l0368 = Mid(l0360, 1, l0364 - 1)
area51.Text5.Text = l0368
l0338 = InStr(1, l0368, "OnlineHost:" + Chr(9) + area51.Text1.Text, 1)
If l0338 Then
l030A = l030A + 1
If l030A = 4 Then
l030A = 0
Exit Function
End If
l0360 = Mid(l0360, l0364 + 2, Len(l0360))
GoTo L16678
End If
Else
area51.Text5.Text = l0360
l0338 = InStr(1, l0360, "OnlineHost:" + Chr(9) + area51.Text1.Text, 1)
If l0338 Then
l030A = l030A + 1
If l030A = 4 Then
l030A = 0
Exit Function
End If
End If
End If
l0330 = area51.Text5.Text
l036C = area51.Text3.Text
area51.Text4.Text = "0"
Do Until 2 > 3
l0370 = InStr(l036C, l0330)
If l0370 Then
area51.Text4.Text = Trim(Str(Int(area51.Text4.Text + 1)))
l036C = Mid(l036C, l0370 + Len(l0330) + 2, Len(l036C))
Else
Exit Do
End If
Loop
l030E = 1
DoEvents
GoTo L164C0
L16916:
area51.Text3.Text = ""
area51.Text4.Text = "0"
area51.Text5.Text = ""
End Function

Function fn1870 (p0378 As Variant) As Variant
Dim l037C As Variant
Dim l0380 As Variant
l037C = SendMessageByNum(p0378, &H201, 0, 0)
l0380 = SendMessageByNum(p0378, &H202, 0, 0)
DoEvents
DoEvents
End Function

Function fn18A8 () As Variant
Dim l0388 As Variant
Dim l038C As Variant
Dim l0390 As Variant
Dim l0394 As Variant
Dim l0398 As Variant
Dim l039C As Variant
Dim l03A0 As String
Dim l03A2 As Variant
Dim l03A6 As Variant
Dim l03AA As Variant
Dim l03AE As Variant
Dim l03B2 As Variant
Dim l03B6 As Variant
Dim l03BA As Variant
Dim l03BE As Variant
Dim l03C2 As Variant
Dim l03C6 As Variant
Dim l03CA As Variant
Dim l03CE As String
Dim l03D0 As String
Dim l03D2 As Variant
Dim l03D6 As Variant
Dim l03DA As Variant
Dim l03DE As Variant
Dim l03E2 As Variant
Dim l03E6 As Variant
Dim l03EA As Variant
l0388 = 0
l038C = 0
l0390 = findwindow("AOL FRAME25", 0&)
l0394 = fn2F8()
l0398 = l0394
l039C = Findchildbyclass(l0398, "_AOL_VIEW")
l03A0$ = getapitext(l039C)
area51.Text3.Text = l03A0$
For l03A2 = 0 To Len(area51.Text3.Text) - 1
l03A6 = Mid(area51.Text3.Text, Len(area51.Text3.Text) - l03A2, 1)
If l03A6 = Chr(10) Then Exit For
l03AA = l03A6 + l03AA
DoEvents
Next l03A2
l03AE = area51.Text3.Text
Do Until 2 > 3
l03B2 = InStr(l03AE, l03AA)
If l03B2 Then
area51.Text4.Text = Trim(Str(Int(area51.Text4.Text + 1)))
l03AE = Mid(l03AE, l03B2 + Len(l03AA) + 2, Len(l03AE))
Else
Exit Do
End If
DoEvents
Loop
l03B6 = findwindow("AOL FRAME25", 0&)
l03BA = fn2F8()
l03BE = l03BA
l03C2 = Findchildbyclass(l03BE, "_AOL_VIEW")
l03C6 = SendMessageByNum(l03C2, 14, 0, 0)
10179 :
If l038C = 1 Then
l03C6 = l03CA
End If
Do Until 2 > 3
l03BE = fn2F8()
l03C2 = Findchildbyclass(l03BE, "_AOL_VIEW")
l03CA = SendMessageByNum(l03C2, 14, 0, 0)
If l03C6 <> l03CA Then Exit Do
If fn2F8() = 0 Then
End If
DoEvents
Loop
l03CE$ = getapitext(l03C2)
area51.Text3.Text = l03CE$
l03D0$ = area51.Text3.Text
For l03D2 = 1 To Int(area51.Text4.Text)
DoEvents
l03D6 = InStr(l03D0$, l03AA)
If l03D6 Then
l03D0$ = Mid(l03D0$, l03D6 + Len(l03AA) + 2, Len(l03D0$))
Else
GoTo 15579
End If
Next l03D2

15579 :
l03DA = l03D0$

30079 :
DoEvents
l03DE = InStr(l03DA, Chr(13) + Chr(10))
If l03DE Then
l03E2 = Mid(l03DA, 1, l03DE - 1)
area51.Text5.Text = l03E2
l03B2 = InStr(1, l03E2, fnA68() + ":" + Chr(9), 1)
If l03B2 Then
l0388 = l0388 + 1
If l0388 = 2 Then
l0388 = 0
Exit Function
End If
l03DA = Mid(l03DA, l03DE + 2, Len(l03DA))
GoTo 30079
End If
Else
area51.Text5.Text = l03DA
l03B2 = InStr(1, l03DA, fnA68() + ":" + Chr(9), 1)
If l03B2 Then
l0388 = l0388 + 1
If l0388 = 2 Then
l0388 = 0
Exit Function
End If
End If
End If
l03AA = area51.Text5.Text
l03E6 = area51.Text3.Text
area51.Text4.Text = "0"
Do Until 2 > 3
l03EA = InStr(l03E6, l03AA)
If l03EA Then
area51.Text4.Text = Trim(Str(Int(area51.Text4.Text + 1)))
l03E6 = Mid(l03E6, l03EA + Len(l03AA) + 2, Len(l03E6))
Else
Exit Do
End If
Loop
l038C = 1
DoEvents
GoTo 10179
L170D4:
area51.Text3.Text = ""
area51.Text4.Text = "0"
area51.Text5.Text = ""
End Function

Function fn18E0 () As Variant
Dim l03F2 As Variant
Dim l03F6 As Variant
Dim l03FA As Variant
Dim l03FE As Variant
l03F2 = findwindow("AOL FRAME25", 0&)
l03F6 = Findchildbytitle(l03F2, "Send")
l03FA = SendMessageByNum(l03F6, &H201, 0, 0)
l03FE = SendMessageByNum(l03F6, &H202, 0, 0)
End Function

Function fn1918 () As Variant
Dim l0406 As Variant
Dim l040A As Variant
Dim l040E As Variant
Dim l0412 As Variant
Dim l0416 As Variant
Dim l041A As String
Dim l0420 As Variant
Dim l0424 As Variant
Dim l0428 As Variant
Dim l042C As Variant
Dim l0430 As Variant
Dim l0434 As Variant
Dim l0438 As Variant
Dim l043C As Variant
Dim l0440 As Variant
Dim l0444 As Variant
Dim l0448 As Variant
Dim l044C As String
Dim l044E As String
Dim l0450 As Variant
Dim l0454 As Variant
Dim l0458 As Variant
Dim l045C As Variant
Dim l0460 As Variant
Dim l0464 As Variant
Dim l0468 As Variant
Dim l046C As Variant
Dim l0470 As Variant
Dim l0474 As Variant
l0406 = 0
l040A = findwindow("AOL FRAME25", 0&)
l040E = fn2F8()
l0412 = l040E
l0416 = Findchildbyclass(l0412, "_AOL_VIEW")
l041A$ = getapitext(l0416)
vote.Text5.Text = l041A$
For l0420 = 0 To Len(vote.Text5.Text) - 1
l0424 = Mid(vote.Text5.Text, Len(vote.Text5.Text) - l0420, 1)
If l0424 = Chr(10) Then Exit For
l0428 = l0424 + l0428
DoEvents
Next l0420
l042C = vote.Text5.Text
Do Until 2 > 3
l0430 = InStr(l042C, l0428)
If l0430 Then
vote.Text6.Text = Trim(Str(Int(vote.Text6.Text + 1)))
l042C = Mid(l042C, l0430 + Len(l0428) + 2, Len(l042C))
Else
Exit Do
End If
DoEvents
Loop
If vote.Command3D1.Caption = "Start" Then GoTo L18312
l0434 = findwindow("AOL FRAME25", 0&)
l0438 = fn2F8()
l043C = l0438
l0440 = Findchildbyclass(l043C, "_AOL_VIEW")
l0444 = SendMessageByNum(l0440, 14, 0, 0)
1017 :
If l0406 = 1 Then
l0444 = l0448
End If
Do Until 2 > 3
l043C = fn2F8()
l0440 = Findchildbyclass(l043C, "_AOL_VIEW")
l0448 = SendMessageByNum(l0440, 14, 0, 0)
If vote.Command3D1.Caption = "Start" Then GoTo L18312
If l0444 <> l0448 Then Exit Do
If fn2F8() = 0 Then
vote.Timer1.Interval = 0
vote.Timer1.Enabled = False
vote.Timer2.Interval = 0
vote.Timer2.Enabled = False
vote.Command3D1.Caption = "Start"
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo L18312
End If
DoEvents
Loop
If vote.Command3D1.Caption = "Start" Then GoTo L18312
l044C$ = getapitext(l0440)
vote.Text5.Text = l044C$
l044E$ = vote.Text5.Text
For l0450 = 1 To Int(vote.Text6.Text)
DoEvents
l0454 = InStr(l044E$, l0428)
If l0454 Then
l044E$ = Mid(l044E$, l0454 + Len(l0428) + 2, Len(l044E$))
Else
GoTo 1557
End If
Next l0450
If vote.Command3D1.Caption = "Start" Then GoTo L18312
1557 :
l0458 = l044E$

3007 :
DoEvents
l045C = InStr(l0458, Chr(13) + Chr(10))
If l045C Then
l0460 = Mid(l0458, 1, l045C - 1)
vote.Text7.Text = l0460
l0430 = InStr(1, l0460, ":" + Chr(9) + vote.Text2.Text, 1)
If l0430 Then
l042C = InStr(l0460, ":")
If l042C = 0 Then GoTo 36
l0434 = Mid(l0460, 1, l042C - 1)
l0438 = Mid(l0434, 1, 1)
If l0438 = Chr(10) Then l0434 = Mid(l0434, 2, Len(l0434))
For l043C = 0 To vote.List1.ListCount - 1
l0440 = InStr(1, vote.List1.List(l043C), l0434, 1)
If l0440 Then GoTo 36
Next l043C
If l0434 = fnA68() Then GoTo 36
vote.List1.AddItem l0434
vote.List2.AddItem l0434
vote.Command3D4.Enabled = True
l0464 = Timeout(.6)
l0468 = fnAD8(fn448() + l0434 + " - Vote has been added" + fn480())
vote.Label2.Caption = Trim(Str(vote.List2.ListCount))
End If
l0430 = InStr(1, l0460, ":" + Chr(9) + vote.Text3.Text, 1)
If l0430 Then
l042C = InStr(l0460, ":")
If l042C = 0 Then GoTo 36
l0434 = Mid(l0460, 1, l042C - 1)
l0438 = Mid(l0434, 1, 1)
If l0438 = Chr(10) Then l0434 = Mid(l0434, 2, Len(l0434))
For l043C = 0 To vote.List1.ListCount - 1
l0440 = InStr(1, vote.List1.List(l043C), l0434, 1)
If l0440 Then GoTo 36
Next l043C
If l0434 = l046C Then GoTo 36
vote.List1.AddItem l0434
vote.List3.AddItem l0434
vote.Command3D2.Enabled = True
l0464 = Timeout(.6)
l0468 = fnAD8(fn448() + l0434 + " - Vote has been added" + fn480())
vote.Label4.Caption = Trim(Str(vote.List3.ListCount))
End If
36 :
l0458 = Mid(l0458, l045C + 2, Len(l0458))
GoTo 3007
Else
vote.Text7.Text = l0458
l0430 = InStr(1, l0458, ":" + Chr(9) + vote.Text2.Text, 1)
If l0430 Then
l042C = InStr(l0458, ":")
If l042C = 0 Then GoTo 37
l0434 = Mid(l0458, 1, l042C - 1)
l0438 = Mid(l0434, 1, 1)
If l0438 = Chr(10) Then l0434 = Mid(l0434, 2, Len(l0434))
For l043C = 0 To vote.List1.ListCount - 1
l0440 = InStr(1, vote.List1.List(l043C), l0434, 1)
If l0440 Then GoTo 37
Next l043C
If l0434 = fnA68() Then GoTo 37
vote.List1.AddItem l0434
vote.List2.AddItem l0434
vote.Command3D4.Enabled = True
l0464 = Timeout(.6)
l0468 = fnAD8(fn448() + l0434 + " - Vote has been added" + fn480())
vote.Label2.Caption = Trim(Str(vote.List2.ListCount))
End If
l0430 = InStr(1, l0458, ":" + Chr(9) + vote.Text3.Text, 1)
If l0430 Then
l042C = InStr(l0458, ":")
If l042C = 0 Then GoTo 37
l0434 = Mid(l0458, 1, l042C - 1)
l0438 = Mid(l0434, 1, 1)
If l0438 = Chr(10) Then l0434 = Mid(l0434, 2, Len(l0434))
For l043C = 0 To vote.List1.ListCount - 1
l0440 = InStr(1, vote.List1.List(l043C), l0434, 1)
If l0440 Then GoTo 37
Next l043C
If l0434 = l046C Then GoTo 37
vote.List1.AddItem l0434
vote.List3.AddItem l0434
vote.Command3D2.Enabled = True
l0464 = Timeout(.6)
l0468 = fnAD8(fn448() + l0434 + " - Vote has been added" + fn480())
vote.Label4.Caption = Trim(Str(vote.List3.ListCount))
End If
37 :
End If
If vote.Command3D1.Caption = "Start" Then GoTo L18312
l0428 = vote.Text7.Text
l0470 = vote.Text5.Text
vote.Text6.Text = "0"
Do Until 2 > 3
l0474 = InStr(l0470, l0428)
If l0474 Then
vote.Text6.Text = Trim(Str(Int(vote.Text6.Text + 1)))
l0470 = Mid(l0470, l0474 + Len(l0428) + 2, Len(l0470))
Else
Exit Do
End If
Loop
l0406 = 1
If vote.Command3D1.Caption = "Start" Then GoTo L18312
DoEvents
GoTo 1017
L18312:
vote.Text5.Text = ""
vote.Text6.Text = "0"
vote.Text7.Text = ""
End Function

Function Timeout (duration As Integer) As Variant
Dim starttime As Variant
starttime = Timer
While Timer - starttime < duration
DoEvents
Wend
End Function

