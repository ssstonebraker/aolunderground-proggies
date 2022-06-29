' Module2
Option Explicit
Declare Sub extsub1678 Lib "User" Alias "bringwindowtotop" (ByVal p1%)
Declare Function extfn16B0 Lib "User" Alias "enablewindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1640 Lib "fate01.dll" Alias "AOLgetcombo" () As Integer
Declare Function extfn1608 Lib "fate01.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1790 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn1800 Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn18E0 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn17C8 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1838 Lib "User" Alias "gettopwindow" (ByVal p1%) As Integer
Declare Function extfn1720 Lib "User" Alias "iswindowenabled" (ByVal p1%) As Integer
Declare Function extfn16E8 Lib "User" Alias "iswindowvisible" (ByVal p1%) As Integer
Declare Function extfn1870 Lib "Kernel" Alias "_lread" () As Integer
Declare Function extfn18A8 Lib "Kernel" Alias "_lwrite" () As Integer
Declare Function extfn1598 Lib "User" Alias "messagebox" () As Integer
Declare Function extfn15D0 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn1758 Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer

Function fn1330 () As Variant
Dim l0040 As Variant
Dim l0046 As Variant
Dim l004A As String
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005E As Variant
Dim l0062 As Variant
Dim l006C As Variant
Dim l0070 As String
Dim l0074 As Variant
Dim l0078 As Variant
Dim l007C As Variant
Dim l0080 As Variant
Dim l0084 As Variant
Dim l0088 As Variant
Dim l008C As Variant
Dim l0090 As Variant
Dim l0094 As Variant
Dim l0098 As Variant
Dim l009C As Variant
l0040 = extfn1018("AOL FRAME25", 0&)
l0046 = extfn10F8(l0040, "Welcome")
l004A$ = String(30, 0)
l004E = extfnFA8(l0046, l004A$, 250)
l0052 = Mid(l004A$, 10, 10)
l0056 = InStr(l0052, "!")
If  l0056 Then
l0052 = Mid(l0052, 1, l0056 - 1)
End If

422 :
If  afkb.Command3D1.Caption = "Start" Then GoTo LEFBC
DoEvents
l005E = extfn1018("AOL FRAME25", 0&)
l0062 = extfn10F8(l005E, "AOL Live!")
l005E = extfn1088(l0062)
l0062 = extfn10C0(l005E, "_AOL_VIEW")
If  l0062 = 0 Then
Exit Function
End If
l006C = extfnFE0(l0062, 14, 0, 0)
If  l006C = Len(afkb.Text4.Text) Then GoTo 422
l0070$ = String(l006C, 0)
l0074 = extfn11A0(l0062, 13, l006C, l0070$)
l0078 = Len(afkb.Text4.Text)
l007C = Len(l0070$) - l0078
On Error Resume Next
l0080 = Right(l0070$, l007C)
l0080 = Mid(l0080, 3, Len(l0080))
444 :
If  afkb.Command3D1.Caption = "Start" Then GoTo LEFBC
l0084 = InStr(l0080, Chr(10))
If  l0084 Then
l0088 = Mid(l0080, 1, l0084 - 2)
afkb.Text5.Text = l0088
If  afkb.Text5.Text <> "" Then afkb.Text5.Text = afkb.Text5.Text
l008C = InStr(afkb.Text5.Text, Chr(0))
If  l008C Then
afkb.Text5.Text = Mid(afkb.Text5.Text, 1, l008C - 1)
End If
l0090 = InStr(1, afkb.Text5.Text, "/" + l0052 + " ", 1)
If  l0090 Then
l0094 = Mid(afkb.Text5.Text, l0090 - 2, 2)
If  l0094 <> ":" + Chr(9) Then
GoTo 46
End If
l0094 = InStr(afkb.Text5.Text, ":")
l0098 = Mid(afkb.Text5.Text, 1, l0094 - 1)
If  l0098 = l0052 Then GoTo 46
l009C = Mid(afkb.Text5.Text, l0090 + 2 + Len(l0052), Len(afkb.Text5.Text))
afkb.List1.AddItem l0098
afkb.List2.AddItem l009C
afkb.List1.Selected(0) = True
End If

46 :
l0080 = Mid(l0080, l0084 + 1, Len(l0080))
GoTo 444
End If
afkb.Text5.Text = l0080
If  afkb.Text5.Text <> "" Then afkb.Text5.Text = afkb.Text5.Text
l008C = InStr(afkb.Text5.Text, Chr(0))
If  l008C Then
afkb.Text5.Text = Mid(afkb.Text5.Text, 1, l008C - 1)
End If
l0090 = InStr(1, afkb.Text5.Text, "/" + l0052 + " ", 1)
If  l0090 Then
l0094 = Mid(afkb.Text5.Text, l0090 - 2, 2)
If  l0094 <> ":" + Chr(9) Then
GoTo 455
End If
l0094 = InStr(afkb.Text5, ":")
l0098 = Mid(afkb.Text5.Text, 1, l0094 - 1)
If  l0098 = l0052 Then GoTo 455
l009C = Mid(afkb.Text5.Text, l0090 + 2 + Len(l0052), Len(afkb.Text5.Text))
afkb.List1.AddItem l0098
afkb.List2.AddItem l009C
afkb.List1.Selected(0) = True
End If

455 :
If  afkb.Command3D1.Caption = "Start" Then GoTo LEFBC
afkb.Text4.Text = l0070$
DoEvents
GoTo 422

LEFBC:
afkb.Text4.Text = ""
afkb.Text5.Text = ""
afkb.Text5.Text = ""
End Function

Function fn1368 () As Variant
Dim l00A4 As Variant
Dim l00A8 As Variant
Dim l00AC As String
Dim l00AE As Variant
Dim l00B2 As Variant
Dim l00B6 As Variant
Dim l00BE As Variant
Dim l00C2 As Variant
Dim l00C6 As Variant
Dim l00CA As String
Dim l00CC As Variant
Dim l00D0 As Variant
Dim l00D4 As Variant
Dim l00D8 As Variant
Dim l00DC As Variant
Dim l00E0 As Variant
Dim l00E4 As Variant
Dim l00E8 As Variant
Dim l00EC As Variant
Dim l00F0 As Variant
Dim l00F4 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Dim l0100 As Variant
Dim l0106 As Variant
l00A4 = extfn1018("AOL FRAME25", 0&)
l00A8 = extfn10F8(l00A4, "Welcome")
l00AC$ = String(30, 0)
l00AE = extfnFA8(l00A8, l00AC$, 250)
l00B2 = Mid(l00AC$, 10, 10)
l00B6 = InStr(l00B2, "!")
If  l00B6 Then
l00B2 = Mid(l00B2, 1, l00B6 - 1)
End If

42 :
If  ECHO.Command3D1.Caption = "Echo" Then GoTo LF86C
DoEvents
l00BE = extfn1018("AOL FRAME25", 0&)
l00C2 = extfn10F8(l00BE, "AOL Live!")
l00BE = extfn1088(l00C2)
l00C2 = extfn10C0(l00BE, "_AOL_VIEW")
If  l00C2 = 0 Then
Exit Function
End If
l00C6 = extfnFE0(l00C2, 14, 0, 0)
If  l00C6 = Len(ECHO.text1.Text) Then GoTo 42
l00CA$ = String(l00C6, 0)
l00CC = extfn11A0(l00C2, 13, l00C6, l00CA$)
l00D0 = Len(ECHO.text1.Text)
l00D4 = Len(l00CA$) - l00D0
On Error Resume Next
l00D8 = Right(l00CA$, l00D4)
l00D8 = Mid(l00D8, 3, Len(l00D8))
44 :
If  ECHO.Command3D1.Caption = "Echo" Then GoTo LF86C
l00DC = InStr(l00D8, Chr(10))
If  l00DC Then
l00E0 = Mid(l00D8, 1, l00DC - 2)
ECHO.text2.Text = l00E0
If  ECHO.text2.Text <> "" Then ECHO.text3.Text = ECHO.text2.Text
l00E4 = InStr(ECHO.text2.Text, Chr(0))
If  l00E4 Then
ECHO.text2.Text = Mid(ECHO.text2.Text, 1, l00E4 - 1)
End If
l00E8 = InStr(ECHO.text2.Text, ":")
l00EC = Mid(ECHO.text2.Text, 1, l00E8 - 1)
If  l00EC = l00B2 Then GoTo 45
l00F0 = InStr(1, l00EC, ECHO.echo1.Text, 1)
If  l00F0 Then
l00F4 = Mid(ECHO.text2.Text, l00E8 + 2, Len(ECHO.text2.Text))
If  ECHO.Check3D1.Value = True Then
l00F8 = l00F4
l00F4 = ""
For l00FC = 0 To Len(l00F8) - 1
l0100 = Mid(l00F8, Len(l00F8) - l00FC, 1)
l00F4 = l00F4 + l0100
Next l00FC
End If
l0106 = fnA68(l00F4)
End If
45 :
l00D8 = Mid(l00D8, l00DC + 1, Len(l00D8))
GoTo 44
End If
ECHO.text2.Text = l00D8
If  ECHO.text2.Text <> "" Then ECHO.text3.Text = ECHO.text2.Text
l00E4 = InStr(ECHO.text2.Text, Chr(0))
If  l00E4 Then
ECHO.text2.Text = Mid(ECHO.text2.Text, 1, l00E4 - 1)
End If
l00E8 = InStr(ECHO.text2.Text, ":")
l00EC = Mid(ECHO.text2.Text, 1, l00E8 - 1)
If  l00EC = l00B2 Then GoTo 59
l00F0 = InStr(1, l00EC, ECHO.echo1.Text, 1)
If  l00F0 Then
l00F4 = Mid(ECHO.text2.Text, l00E8 + 2, Len(ECHO.text2.Text))
If  ECHO.Check3D1.Value = True Then
l00F8 = l00F4
l00F4 = ""
For l00FC = 0 To Len(l00F8) - 1
l0100 = Mid(l00F8, Len(l00F8) - l00FC, 1)
l00F4 = l00F4 + l0100
Next l00FC
End If
l0106 = fnA68(l00F4)
End If
59 :
If  ECHO.Command3D1.Caption = "Echo" Then GoTo LF86C
ECHO.text1.Text = l00CA$
ECHO.text2.Text = ""
DoEvents
GoTo 42
LF86C:
ECHO.text1.Text = ""
ECHO.text2.Text = ""
ECHO.text3.Text = ""
End Function

Function fn13A0 (p010E As Integer) As Variant
Dim l0110 As Variant
l0110 = Timer
While Timer - l0110 < p010E
DoEvents
Wend
End Function

Function fn13D8 () As Variant
Dim l0118 As Variant
Dim l011C As Variant
Dim l0120 As String
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As String
Dim l0140 As Variant
Dim l0144 As Variant
Dim l0148 As Variant
Dim l014C As Variant
Dim l0150 As Variant
Dim l0154 As Variant
Dim l0158 As Variant
Dim l015C As Variant
Dim l0160 As Variant
Dim l0164 As Variant
Dim l0168 As Variant
Dim l016C As Variant
Dim l0172 As Variant
l0118 = extfn1018("AOL FRAME25", 0&)
l011C = extfn10F8(l0118, "Welcome")
l0120$ = String(30, 0)
l0122 = extfnFA8(l011C, l0120$, 250)
l0126 = Mid(l0120$, 10, 10)
l012A = InStr(l0126, "!")
If  l012A Then
l0126 = Mid(l0126, 1, l012A - 1)
End If

320 :
If  fightbot.Command3D2.Caption = "Start" Then GoTo L1058E
DoEvents
l0132 = extfn1018("AOL FRAME25", 0&)
l0136 = extfn10F8(l0132, "AOL Live!")
l0132 = extfn1088(l0136)
l0136 = extfn10C0(l0132, "_AOL_VIEW")
If  l0136 = 0 Then
Exit Function
End If
l013A = extfnFE0(l0136, 14, 0, 0)
If  l013A = Len(fightbot.Text4.Text) Then GoTo 320
l013E$ = String(l013A, 0)
l0140 = extfn11A0(l0136, 13, l013A, l013E$)
l0144 = Len(fightbot.Text4.Text)
l0148 = Len(l013E$) - l0144
On Error Resume Next
l014C = Right(l013E$, l0148)
l014C = Mid(l014C, 3, Len(l014C))
340 :
If  fightbot.Command3D2.Caption = "Start" Then GoTo L1058E
l0150 = InStr(l014C, Chr(10))
If  l0150 Then
l0154 = Mid(l014C, 1, l0150 - 2)
fightbot.Text5.Text = l0154
l0158 = InStr(fightbot.Text5.Text, Chr(0))
If  l0158 Then
fightbot.Text5.Text = Mid(fightbot.Text5.Text, 1, l0158 - 1)
End If
l015C = InStr(1, fightbot.Text5.Text, fightbot.Text3.Text, 1)
If  l015C Then
l0160 = Mid(fightbot.Text5.Text, l015C - 2, 2)
If  l0160 <> ":" + Chr(9) Then
GoTo 360
End If
l0160 = InStr(fightbot.Text5.Text, ":")
l0164 = Mid(fightbot.Text5.Text, 1, l0160 - 1)
l0168 = Mid(fightbot.Text5.Text, l0160 + 2, Len(fightbot.Text5.Text))
If  Len(l0168) <> Len(fightbot.Text3.Text) Then GoTo 360
If  l0164 = l0126 Then GoTo 360
l016C = fn13A0(.6#)
If  fightbot.Text1.Text = "" And fightbot.Text2.Text > "" Then
fightbot.Text1.Text = l0164
l0172 = fnA30("«‹- -– " + l0164 + " will challenge " + fightbot.Text2.Text + " –-- -›»")
GoTo L1058E
End If
If  fightbot.Text2.Text = "" And fightbot.Text1.Text > "" Then
fightbot.Text2.Text = l0164
l0172 = fnA30("«‹- -– " + l0164 + " will challenge " + fightbot.Text1.Text + " –-- -›»")
GoTo L1058E
End If
If  fightbot.Text1.Text = "" And fightbot.Text2.Text = "" Then
fightbot.Text1.Text = l0164
l0172 = fnA30("«‹- -– " + l0164 + " will be the first fighter! –-- -›»")
GoTo 360
End If
End If
360 :
l014C = Mid(l014C, l0150 + 1, Len(l014C))
GoTo 340
End If
fightbot.Text5.Text = l014C
If  fightbot.Text5.Text <> "" Then fightbot.Text5.Text = fightbot.Text5.Text
l0158 = InStr(fightbot.Text5.Text, Chr(0))
If  l0158 Then
fightbot.Text5.Text = Mid(fightbot.Text5.Text, 1, l0158 - 1)
End If
l015C = InStr(1, fightbot.Text5.Text, fightbot.Text3.Text, 1)
If  l015C Then
l0160 = Mid(fightbot.Text5.Text, l015C - 2, 2)
If  l0160 <> ":" + Chr(9) Then
GoTo 350
End If
l0160 = InStr(fightbot.Text5.Text, ":")
l0164 = Mid(fightbot.Text5.Text, 1, l0160 - 1)
l0168 = Mid(fightbot.Text5.Text, l0160 + 2, Len(fightbot.Text5.Text))
If  Len(l0168) <> Len(fightbot.Text3.Text) Then GoTo 350
If  l0164 = l0126 Then GoTo 350
l016C = fn13A0(.6#)
If  fightbot.Text1.Text = "" And fightbot.Text2.Text > "" Then
l0172 = fnA30("«‹- -– " + l0164 + " will challenge " + fightbot.Text2.Text + " –-- -›»")
fightbot.Text1.Text = l0164
GoTo L1058E
End If
If  fightbot.Text2.Text = "" And fightbot.Text1.Text > "" Then
l0172 = fnA30("«‹- -– " + l0164 + " will challenge " + fightbot.Text1.Text + " –-- -›»")
fightbot.Text2.Text = l0164
GoTo L1058E
End If
If  fightbot.Text1.Text = "" And fightbot.Text2.Text = "" Then
fightbot.Text1.Text = l0164
l0172 = fnA30("«‹- -– " + l0164 + " will be the first fighter! –-- -›»")
GoTo 350
End If
End If
350 :
If  fightbot.Command3D2.Caption = "Start" Then GoTo L1058E
fightbot.Text4.Text = l013E$
DoEvents
GoTo 320
L1058E:
fightbot.Command3D2.Caption = "Start"
fightbot.Timer2.Interval = 0
fightbot.Timer2.Enabled = False
fightbot.Text4.Text = ""
fightbot.Text5.Text = ""
End Function

Function fn1410 () As Variant
Dim l017A As Variant
Dim l017E As Variant
Dim l0182 As String
Dim l0184 As Variant
Dim l0188 As Variant
Dim l018C As Variant
Dim l0194 As Variant
Dim l0198 As Variant
Dim l019C As Variant
Dim l01A0 As String
Dim l01A2 As Variant
Dim l01A6 As Variant
Dim l01AA As Variant
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
Dim l01DE As Variant
Dim l01E2 As Variant
Dim l01E6 As Variant
Dim l01EA As Variant
l017A = extfn1018("AOL FRAME25", 0&)
l017E = extfn10F8(l017A, "Welcome")
l0182$ = String(30, 0)
l0184 = extfnFA8(l017E, l0182$, 250)
l0188 = Mid(l0182$, 10, 10)
l018C = InStr(l0188, "!")
If  l018C Then
l0188 = Mid(l0188, 1, l018C - 1)
End If

52 :
If  getmm.Command3D1.Caption = "Start Bot" Then GoTo L1104A
DoEvents
l0194 = extfn1018("AOL FRAME25", 0&)
l0198 = extfn10F8(l0194, "AOL Live!")
l0194 = extfn1088(l0198)
l0198 = extfn10C0(l0194, "_AOL_VIEW")
If  l0198 = 0 Then
Exit Function
End If
l019C = extfnFE0(l0198, 14, 0, 0)
If  l019C = Len(getmm.Text1.Text) Then GoTo 52
l01A0$ = String(l019C, 0)
l01A2 = extfn11A0(l0198, 13, l019C, l01A0$)
l01A6 = Len(getmm.Text1.Text)
l01AA = Len(l01A0$) - l01A6
On Error Resume Next
l01AE = Right(l01A0$, l01AA)
l01AE = Mid(l01AE, 3, Len(l01AE))
54 :
If  getmm.Command3D1.Caption = "Start Bot" Then GoTo L1104A
l01B2 = InStr(l01AE, Chr(10))
If  l01B2 Then
l01B6 = Mid(l01AE, 1, l01B2 - 2)
getmm.Text2.Text = l01B6
l01BA = InStr(getmm.Text2.Text, Chr(0))
If  l01BA Then
getmm.Text2.Text = Mid(getmm.Text2.Text, 1, l01BA - 1)
End If
l01BE = InStr(getmm.Text2.Text, ":")
l01C2 = Mid(getmm.Text2.Text, 1, l01BE - 1)
l01C6 = Mid(getmm.Text2.Text, l01BE + 2, Len(getmm.Text2.Text))
l01CA = InStr(1, l01C6, "type", 1)
If  l01CA Then
If  l01C2 = l0188 Then GoTo 55
58 :
l01CE = InStr(l01C6, Chr(34))
If  l01CE Then
l01D2 = Mid(l01C6, l01CE + 1, Len(l01C6))
l01D6 = InStr(l01D2, Chr(34))
If  l01D6 Then
l01DA = Mid(l01D2, 1, l01D6 - 1)
For l01DE = 0 To getmm.List1.ListCount - 1
l01E2 = InStr(1, getmm.List1.List(l01DE), l01C2, 1)
If  l01E2 Then GoTo 55
Next l01DE
l01E6 = Mid(getmm.Frame3D1.Caption, 11, Len(getmm.Frame3D1.Caption))
l01EA = l01E6 + 1
getmm.Frame3D1.Caption = "Added to: " + l01EA
getmm.List1.AddItem l01C2
l01DE = fnA68(l01DA)
End If
End If
Else
l01CA = InStr(1, l01C6, "say", 1)
If  l01CA Then GoTo 58
End If
55 :
l01AE = Mid(l01AE, l01B2 + 1, Len(l01AE))
GoTo 54
End If
getmm.Text2.Text = l01AE
l01BA = InStr(getmm.Text2.Text, Chr(0))
If  l01BA Then
getmm.Text2.Text = Mid(getmm.Text2.Text, 1, l01BA - 1)
End If
l01BE = InStr(getmm.Text2.Text, ":")
l01C2 = Mid(getmm.Text2.Text, 1, l01BE - 1)
l01C6 = Mid(getmm.Text2.Text, l01BE + 2, Len(getmm.Text2.Text))
l01CA = InStr(1, l01C6, "type", 1)
If  l01CA Then
If  l01C2 = l0188 Then GoTo 56
57 :
l01CE = InStr(l01C6, Chr(34))
If  l01CE Then
l01D2 = Mid(l01C6, l01CE + 1, Len(l01C6))
l01D6 = InStr(l01D2, Chr(34))
If  l01D6 Then
l01DA = Mid(l01D2, 1, l01D6 - 1)
For l01DE = 0 To getmm.List1.ListCount - 1
l01E2 = InStr(1, getmm.List1.List(l01DE), l01C2, 1)
If  l01E2 Then GoTo 56
Next l01DE
getmm.List1.AddItem l01C2
l01E6 = Mid(getmm.Frame3D1.Caption, 11, Len(getmm.Frame3D1.Caption))
l01EA = l01E6 + 1
getmm.Frame3D1.Caption = "Added to: " + l01EA
l01DE = fnA68(l01DA)
End If
End If
Else
l01CA = InStr(1, l01C6, "say", 1)
If  l01CA Then GoTo 57
End If

56 :
If  getmm.Command3D1.Caption = "Start Bot" Then GoTo L1104A
getmm.Text1.Text = l01A0$
DoEvents
GoTo 52
L1104A:
getmm.Text1.Text = ""
getmm.Text2.Text = ""
End Function

Function fn1448 (ByVal pv01F2 As Integer, ByVal pv01F4 As Integer) As Variant
Dim l01F6 As Variant
Dim l01FA As Variant
Dim l01FE As Variant
Dim l0202 As Variant
Dim l0206 As Variant
l01F6 = extfn1018("AOL FRAME25", 0&)
l01FA = extfn1790(l01F6)
l01FE = extfn17C8(l01FA, pv01F2)
l0202 = extfn1800(l01FE, pv01F4)
l0206 = extfnFE0(l01F6, 273, l0202, 0)
End Function

Function fn1480 (p020E As Variant) As Variant
Dim l0212 As Variant
Dim l0216 As Variant
l0212 = extfnFE0(p020E, &H201, 0, 0)
l0216 = extfnFE0(p020E, &H202, 0, 0)
DoEvents
DoEvents
End Function

Function fn14B8 () As Variant
Dim l021E As Variant
Dim l0222 As Variant
Dim l0226 As Variant
Dim l022A As Variant
l021E = extfn1018("AOL FRAME25", 0&)
l0222 = extfn10F8(l021E, "Send")
l0226 = extfnFE0(l0222, &H201, 0, 0)
l022A = extfnFE0(l0222, &H202, 0, 0)
End Function

Function fn14F0 () As Variant
fn14F0 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(87) + Chr(97) + Chr(114) + Chr(101) + Chr(122) + Chr(32) + Chr(83) + Chr(101) + Chr(114) + Chr(118) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn1528 () As Variant
Dim l0236 As Variant
Dim l023A As Variant
Dim l023E As String
Dim l0240 As Variant
Dim l0244 As Variant
Dim l0248 As Variant
Dim l0250 As Variant
Dim l0254 As Variant
Dim l0258 As Variant
Dim l025C As String
Dim l025E As Variant
Dim l0262 As Variant
Dim l0266 As Variant
Dim l026A As Variant
Dim l026E As Variant
Dim l0272 As Variant
Dim l0276 As Variant
Dim l027A As Variant
Dim l027E As Variant
Dim l0282 As Variant
Dim l0286 As Variant
Dim l028A As Variant
Dim l028E As Variant
Dim l0292 As Variant
Dim l0296 As Variant
Dim l029A As Variant
l0236 = extfn1018("AOL FRAME25", 0&)
l023A = extfn10F8(l0236, "Welcome")
l023E$ = String(30, 0)
l0240 = extfnFA8(l023A, l023E$, 250)
l0244 = Mid(l023E$, 10, 10)
l0248 = InStr(l0244, "!")
If  l0248 Then
l0244 = Mid(l0244, 1, l0248 - 1)
End If

32 :
If  vote.Command3D1.Caption = "&Start Bot" Then GoTo L122D2
DoEvents
l0250 = extfn1018("AOL FRAME25", 0&)
l0254 = extfn10F8(l0250, "AOL Live!")
l0250 = extfn1088(l0254)
l0254 = extfn10C0(l0250, "_AOL_VIEW")
If  l0254 = 0 Then
Exit Function
End If
l0258 = extfnFE0(l0254, 14, 0, 0)
If  l0258 = Len(vote.Text5.Text) Then GoTo 32
l025C$ = String(l0258, 0)
l025E = extfn11A0(l0254, 13, l0258, l025C$)
l0262 = Len(vote.Text5.Text)
l0266 = Len(l025C$) - l0262
On Error Resume Next
l026A = Right(l025C$, l0266)
l026A = Mid(l026A, 3, Len(l026A))
34 :
If  vote.Command3D1.Caption = "&Start Bot" Then GoTo L122D2
DoEvents
l026E = InStr(l026A, Chr(10))
If  l026E Then
l0272 = Mid(l026A, 1, l026E - 2)
vote.Text5.Text = l0272
If  vote.Text5.Text <> "" Then vote.Text7.Text = vote.Text5.Text
l0276 = InStr(vote.Text7.Text, Chr(0))
If  l0276 Then
vote.Text7.Text = Mid(vote.Text7.Text, 1, l0276 - 1)
End If
l027A = InStr(1, vote.Text7.Text, vote.Text2.Text, 1)
If  l027A Then
l027E = Mid(vote.Text7.Text, l027A - 2, 2)
If  l027E <> ":" + Chr(9) Then
GoTo 36
End If
l027E = InStr(vote.Text7, ":")
l0282 = Mid(vote.Text7.Text, 1, l027E - 1)
For l0286 = 0 To vote.List1.ListCount - 1
l028A = InStr(1, vote.List1.List(l0286), l0282, 1)
If  l028A Then GoTo 36
Next l0286
If  l0282 = l0244 Then GoTo 36
vote.List1.AddItem l0282
vote.List2.AddItem l0282
l028E = fn13A0(.6#)
l0292 = fnA68("«‹- -– " + l0282 + " - Vote Has Been Added! –-- -›»")
l0296 = Mid(vote.Label3.Caption, 6, Len(vote.Label3.Caption))
l029A = l0296 + 1
vote.Label3.Caption = "Yes: " + l029A
End If
l027A = InStr(1, vote.Text7.Text, vote.Text3.Text, 1)
If  l027A Then
l027E = Mid(vote.Text7.Text, l027A - 2, 2)
If  l027E <> ":" + Chr(9) Then
GoTo 36
End If
l027E = InStr(vote.Text7, ":")
l0282 = Mid(vote.Text7.Text, 1, l027E - 1)
For l0286 = 0 To vote.List1.ListCount - 1
l028A = InStr(1, vote.List1.List(l0286), l0282, 1)
If  l028A Then GoTo 36
Next l0286
If  l0282 = l0244 Then GoTo 36
vote.List1.AddItem l0282
vote.List3.AddItem l0282
l028E = fn13A0(.6#)
l0292 = fnA68("«‹- -– " + l0282 + " - Vote Has Been Added! –-- -›»")
l0296 = Mid(vote.Label5.Caption, 5, Len(vote.Label5.Caption))
l029A = l0296 + 1
vote.Label5.Caption = "No: " + l029A
End If

36 :
l026A = Mid(l026A, l026E + 1, Len(l026A))
GoTo 34
End If
vote.Text5.Text = l026A
If  vote.Text5.Text <> "" Then vote.Text7.Text = vote.Text5.Text
l0276 = InStr(vote.Text7.Text, Chr(0))
If  l0276 Then
vote.Text7.Text = Mid(vote.Text7.Text, 1, l0276 - 1)
End If
l027A = InStr(1, vote.Text7.Text, vote.Text2.Text, 1)
If  l027A Then
l027E = Mid(vote.Text7.Text, l027A - 2, 2)
If  l027E <> ":" + Chr(9) Then
GoTo 35
End If
l027E = InStr(vote.Text7, ":")
l0282 = Mid(vote.Text7.Text, 1, l027E - 1)
For l0286 = 0 To vote.List1.ListCount - 1
l028A = InStr(1, vote.List1.List(l0286), l0282, 1)
If  l028A Then GoTo 35
Next l0286
If  l0282 = l0244 Then GoTo 35
vote.List1.AddItem l0282
vote.List2.AddItem l0282
l028E = fn13A0(.6#)
l0292 = fnA68("«‹- -– " + l0282 + " - Vote Has Been Added! –-- -›»")
l0296 = Mid(vote.Label3.Caption, 6, Len(vote.Label3.Caption))
l029A = l0296 + 1
vote.Label3.Caption = "Yes: " + l029A
End If
l027A = InStr(1, vote.Text7.Text, vote.Text3.Text, 1)
If  l027A Then
l027E = Mid(vote.Text7.Text, l027A - 2, 2)
If  l027E <> ":" + Chr(9) Then
GoTo 35
End If
l027E = InStr(vote.Text7, ":")
l0282 = Mid(vote.Text7.Text, 1, l027E - 1)
For l0286 = 0 To vote.List1.ListCount - 1
l028A = InStr(1, vote.List1.List(l0286), l0282, 1)
If  l028A Then GoTo 35
Next l0286
If  l0282 = l0244 Then GoTo 35
vote.List1.AddItem l0282
vote.List3.AddItem l0282
l028E = fn13A0(.6#)
l0292 = fnA68("«‹- -– " + l0282 + " - Vote Has Been Added! –-- -›»")
l0296 = Mid(vote.Label5.Caption, 5, Len(vote.Label5.Caption))
l029A = l0296 + 1
vote.Label5.Caption = "No: " + l029A
End If

35 :
If  vote.Command3D1.Caption = "&Start Bot" Then GoTo L122D2
vote.Text5.Text = l025C$
DoEvents
GoTo 32
L122D2:
vote.Text5.Text = ""
vote.Text5.Text = ""
vote.Text7.Text = ""
End Function


Function fn1560 () As Variant
fn1560 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(86) + Chr(111) + Chr(116) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function
