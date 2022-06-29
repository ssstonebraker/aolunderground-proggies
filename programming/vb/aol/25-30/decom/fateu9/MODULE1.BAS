' Module1
Option Explicit
Declare Function extfnEC8 Lib "User" Alias "appendmenu" () As Integer
Declare Function extfn10C0 Lib "fate02.dll" Alias "FindChildByClass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnD78 Lib "User" Alias "getclassname" () As Integer
Declare Function extfnF00 Lib "User" Alias "deletemenu" () As Integer
Declare Function extfnDB0 Lib "User" Alias "DestroyWindow" () As Integer
Declare Sub extsubE58 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn10F8 Lib "fate02.dll" Alias "FindChildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1018 Lib "User" Alias "findwindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn12B8 Lib "User" Alias "GetClipboardData" () As Integer
Declare Function extfnCD0 Lib "User" Alias "GetDesktopHwnd" () As Integer
Declare Function extfn1280 Lib "User" Alias "getmenuitemcount" (ByVal p1%) As Integer
Declare Function extfnF38 Lib "User" Alias "getmenustring" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfnD08 Lib "Kernel" Alias "getversion" () As Long
Declare Function extfn1248 Lib "User" Alias "GetWindow" () As Integer
Declare Function extfnDE8 Lib "Kernel" Alias "GetWindowsDirectory" () As Integer
Declare Function extfnE90 Lib "User" Alias "insertmenu" () As Integer
Declare Function extfnFE0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn11A0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfnE20 Lib "User" Alias "modifymenu" () As Integer
Declare Sub extsubF70 Lib "User" Alias "movewindow" ()
Declare Function extfn1168 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1088 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1210 Lib "mmsystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn1130 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfnD40 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn11D8 Lib "User" Alias "setparent" () As Integer
Declare Sub extsub1050 Lib "User" Alias "setwindowpos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfnFA8 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer

Function fn100 () As Variant
Dim l00A6 As Variant
Dim l00AA As Variant
Dim l00AE As Variant
Dim l00B6 As String
Dim l00B8 As Variant
Dim l00BC As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00CC As Variant
22 :
DoEvents
l00A6 = extfn1018("AOL FRAME25", 0&)
l00AA = extfn10F8(l00A6, "AOL Live!")
l00A6 = extfn1088(l00AA)
l00AA = extfn10C0(l00A6, "_AOL_VIEW")
If l00AA = 0 Then
Exit Function
End If
l00AE = extfnFE0(l00AA, 14, 0, 0)
If l00AE = Len(mmER.Text8.Text) Then GoTo 22
l00B6$ = String(l00AE, 0)
l00B8 = extfn11A0(l00AA, 13, l00AE, l00B6$)
l00BC = Len(mmER.Text8.Text)
l00C0 = Len(l00B6$) - l00BC
On Error Resume Next
l00C4 = Right(l00B6$, l00C0)
l00C4 = Mid(l00C4, 3, Len(l00C4))
24 :
l00C8 = InStr(l00C4, Chr(10))
If l00C8 Then
l00CC = Mid(l00C4, 1, l00C8 - 2)
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
mmER.Text9.Text = l00CC
l00C4 = Mid(l00C4, l00C8 + 1, Len(l00C4))
GoTo 24
End If
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
mmER.Text9.Text = l00C4
mmER.Text8.Text = l00B6$
DoEvents
GoTo 22
End Function

Function fn138 () As Integer
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
l00D2 = extfn1018("AOL FRAME25", 0&)
l00D6 = extfn10C0(l00D2, "AOL TOOLBAR")
l00DA = extfn10C0(l00D6, "_AOL_ICON")
l00DE = extfn1168(l00DA, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
l00DE = extfn1168(l00DE, 2)
If l00DE > 0 Then
fn138 = 5
End If
If l00DE = 0 Then
fn138 = 8
End If
End Function

Function fn170 () As Variant
Dim l00E6 As Variant
Dim l00EA As Variant
Dim l00EE As Variant
Dim l00F2 As Variant
Dim l00F6 As Variant
Dim l00FA As String
Dim l00FC As Variant
l00E6 = extfn1018("AOL FRAME25", 0&)
l00EA = extfn10F8(l00E6, "AOL Live!")
l00EE = extfn1088(l00EA)
l00F2 = extfn10C0(l00EE, "_AOL_VIEW")
l00F6 = extfnFE0(l00F2, 14, 0, 0)
l00FA$ = String(l00F6, 0)
l00FC = extfn11A0(l00F2, 13, l00F6, l00FA$)
fn170 = l00FA$
End Function

Function fn1A8 () As Variant
fn1A8 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(77) + Chr(97) + Chr(105) + Chr(108) + Chr(32) + Chr(66) + Chr(111) + Chr(109) + Chr(98) + Chr(101) + Chr(114) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn1E0 () As Variant
Dim l0108 As Variant
Dim l010C As Variant
Dim l0110 As String
Dim l0112 As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
Dim l0126 As Variant
Dim l012A As String
Dim l012C As Variant
Dim l0130 As Variant
Dim l0134 As Variant
Dim l0138 As Variant
Dim l013C As Variant
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
Dim l016E As Variant
Dim l0172 As Variant
l0108 = extfn1018("AOL FRAME25", 0&)
l010C = extfn10F8(l0108, "Welcome")
l0110$ = String(30, 0)
l0112 = extfnFA8(l010C, l0110$, 250)
l0116 = Mid(l0110$, 10, 10)
l011A = InStr(l0116, "!")
If l011A Then
l0116 = Mid(l0116, 1, l011A - 1)
End If

32 :
If mmER.Command3D1.Caption = "&Start Bot" Then GoTo L28B4
DoEvents
l011E = extfn1018("AOL FRAME25", 0&)
l0122 = extfn10F8(l011E, "AOL Live!")
l011E = extfn1088(l0122)
l0122 = extfn10C0(l011E, "_AOL_VIEW")
If l0122 = 0 Then
Exit Function
End If
l0126 = extfnFE0(l0122, 14, 0, 0)
If l0126 = Len(mmER.Text8.Text) Then GoTo 32
l012A$ = String(l0126, 0)
l012C = extfn11A0(l0122, 13, l0126, l012A$)
l0130 = Len(mmER.Text8.Text)
l0134 = Len(l012A$) - l0130
On Error Resume Next
l0138 = Right(l012A$, l0134)
l0138 = Mid(l0138, 3, Len(l0138))
34 :
If mmER.Command3D1.Caption = "&Start Bot" Then GoTo L28B4
l013C = InStr(l0138, Chr(10))
If l013C Then
l0140 = Mid(l0138, 1, l013C - 2)
mmER.Text9.Text = l0140
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
l0144 = InStr(mmER.Text10.Text, Chr(0))
If l0144 Then
mmER.Text10.Text = Mid(mmER.Text10.Text, 1, l0144 - 1)
End If
l0148 = InStr(1, mmER.Text10.Text, mmER.Text4.Text, 1)
If l0148 Then
l014C = Mid(mmER.Text10.Text, l0148 - 2, 2)
If l014C <> ":" + Chr(9) Then
GoTo 36
End If
l014C = InStr(mmER.Text10, ":")
l0150 = Mid(mmER.Text10.Text, 1, l014C - 1)
l0154 = Mid(mmER.Text10.Text, l014C + 2, Len(mmER.Text10.Text))
If Len(l0154) <> Len(mmER.Text4.Text) Then GoTo 36
If l0150 = l0116 Then GoTo 36
For l0158 = 0 To mmER.List1.ListCount - 1
l015C = InStr(1, mmER.List1.List(l0158), l0150, 1)
If l015C Then
If Len(mmER.List1.List(l0158)) = Len(l0150) Then GoTo 36
End If
Next l0158
If l0150 = l0116 Then GoTo 36
mmER.List1.AddItem l0150
mmER.Text1.Text = mmER.Text1.Text + l0150 + ","
            l0160 = InStr(mmER.Frame3D1.Caption, "People to MM: ")
            l0164 = Mid(mmER.Frame3D1.Caption, l0160 + 14, Len(mmER.Frame3D1.Caption))
            l0168 = l0164 + 1
            mmER.Frame3D1.Caption = "People to MM: " + l0168
If mmER.Check3D3.Value = True Then
l016E = fn13A0(.6)
            l0172 = fnA30("‹—" + Chr(94) + Chr(118) + Chr(94) + Chr(91) + Chr(32) + Chr(60) + l0150 + Chr(62) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(111) + Chr(110) + Chr(32) + Chr(116) + Chr(104) + Chr(101) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + Chr(32) + Chr(93) + Chr(94) + Chr(118) + Chr(94) + "—›")
End If
End If
36 :
l0138 = Mid(l0138, l013C + 1, Len(l0138))
GoTo 34
End If
mmER.Text9.Text = l0138
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
l0144 = InStr(mmER.Text10.Text, Chr(0))
If l0144 Then
mmER.Text10.Text = Mid(mmER.Text10.Text, 1, l0144 - 1)
End If
l0148 = InStr(1, mmER.Text10.Text, mmER.Text4.Text, 1)
If l0148 Then
l014C = Mid(mmER.Text10.Text, l0148 - 2, 2)
If l014C <> ":" + Chr(9) Then
GoTo 35
End If
l014C = InStr(mmER.Text10, ":")
l0150 = Mid(mmER.Text10.Text, 1, l014C - 1)
l0154 = Mid(mmER.Text10.Text, l014C + 2, Len(mmER.Text10.Text))
If Len(l0154) <> Len(mmER.Text4.Text) Then GoTo 35
If l0150 = l0116 Then GoTo 35
For l0158 = 0 To mmER.List1.ListCount - 1
l015C = InStr(1, mmER.List1.List(l0158), l0150, 1)
If l015C Then
If Len(mmER.List1.List(l0158)) = Len(l0150) Then GoTo 35
End If
Next l0158
If l0150 = l0116 Then GoTo 35
mmER.List1.AddItem l0150
mmER.Text1.Text = mmER.Text1.Text + l0150 + ","
            l0160 = InStr(mmER.Frame3D1.Caption, "People to MM: ")
            l0164 = Mid(mmER.Frame3D1.Caption, l0160 + 14, Len(mmER.Frame3D1.Caption))
            l0168 = l0164 + 1
            mmER.Frame3D1.Caption = "People to MM: " + l0168
If mmER.Check3D3.Value = True Then
l016E = fn13A0(.6)
            l0172 = fnA30("‹—" + Chr(94) + Chr(118) + Chr(94) + Chr(91) + Chr(32) + Chr(60) + l0150 + Chr(62) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(111) + Chr(110) + Chr(32) + Chr(116) + Chr(104) + Chr(101) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + Chr(32) + Chr(93) + Chr(94) + Chr(118) + Chr(94) + "—›")
End If
End If
35 :
If mmER.Command3D1.Caption = "&Start Bot" Then GoTo L28B4
mmER.Text8.Text = l012A$
DoEvents
GoTo 32
L28B4:
mmER.Text8.Text = ""
mmER.Text9.Text = ""
mmER.Text10.Text = ""
End Function

Function fn218 () As Variant
fn218 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(66) + Chr(117) + Chr(115) + Chr(116) + Chr(32) + Chr(73) + Chr(110) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn288 () As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As Variant
Dim l019E As Variant
Dim l01A2 As Variant
Dim l01A6 As String
Dim l01A8 As Variant
Dim l01AC As Variant
Dim l01B0 As Variant
Dim l01B4 As Variant
l0192 = extfn1018("AOL FRAMe25", 0&)
l0196 = extfn10F8(l0192, "AOL Live!")
l019A = extfn1088(l0196)
l019E = extfn10C0(l019A, "_AOL_EDIT")
l01A2 = extfn1088(l019E)
l01A6$ = String(255, 0)
l01A8 = extfnFA8(l01A2, l01A6$, 250)
l01AC = l01A6$
l01B0 = InStr(l01AC, Chr(0))
If l01B0 Then
l01AC = Mid(l01AC, 1, l01B0 - 1)
End If
l01B4 = fnBF0(l01AC)
fn288 = l01B4
End Function

Function fn2C0 () As Variant
Dim l01BC As Variant
Dim l01C0 As Variant
Dim l01C4 As String
Dim l01C6 As Variant
Dim l01CA As Variant
Dim l01CE As Variant
Dim l01D6 As Variant
Dim l01DA As Variant
Dim l01DE As Variant
Dim l01E2 As String
Dim l01E4 As Variant
Dim l01E8 As Variant
Dim l01EC As Variant
Dim l01F0 As Variant
Dim l01F4 As Variant
Dim l01F8 As Variant
Dim l01FC As Variant
Dim l0200 As Variant
Dim l0204 As Variant
Dim l0208 As Variant
Dim l020C As Variant
Dim l0210 As Variant
Dim l0214 As Variant
Dim l0218 As Variant
Dim l021C As Variant
Dim l0220 As Variant
Dim l0224 As Variant
Dim l0228 As Variant
Dim l022C As Variant
Dim l0230 As Variant
Dim l0234 As Variant
Dim l0238 As Variant
Dim l023C As Variant
Dim l0240 As Variant
Dim l0244 As Variant
Dim l0248 As Variant
l01BC = extfn1018("AOL FRAME25", 0&)
l01C0 = extfn10F8(l01BC, "Welcome")
l01C4$ = String(30, 0)
l01C6 = extfnFA8(l01C0, l01C4$, 250)
l01CA = Mid(l01C4$, 10, 10)
l01CE = InStr(l01CA, "!")
If l01CE Then
l01CA = Mid(l01CA, 1, l01CE - 1)
End If
10003 :
DoEvents
If encryptchat.Timer1.Interval = 0 Then GoTo L3A4C
l01D6 = extfn1018("AOL FRAME25", 0&)
l01DA = extfn10F8(l01D6, "AOL Live!")
l01D6 = extfn1088(l01DA)
l01DA = extfn10C0(l01D6, "_AOL_VIEW")
If l01DA = 0 Then
Exit Function
End If
l01DE = extfnFE0(l01DA, 14, 0, 0)
If l01DE = Len(encryptchat.Text5.Text) Then GoTo 10003
l01E2$ = String(l01DE, 0)
l01E4 = extfn11A0(l01DA, 13, l01DE, l01E2$)
l01E8 = Len(encryptchat.Text5.Text)
l01EC = Len(l01E2$) - l01E8
On Error Resume Next
l01F0 = Right(l01E2$, l01EC)
l01F0 = Mid(l01F0, 3, Len(l01F0))
10002 :
l01F4 = InStr(l01F0, Chr(10))
If l01F4 Then
l01F8 = Mid(l01F0, 1, l01F4 - 2)
l01FC = InStr(l01F8, ":")
l0200 = Mid(l01F8, 1, l01FC - 1)
If l0200 = l01CA Then GoTo 10000
For l0204 = 0 To encryptchat.List1.ListCount - 1
l0208 = InStr(1, encryptchat.List1.List(l0204), l0200, 1)
If l0208 Then
If Len(encryptchat.List1.List(l0204)) = Len(l0200) Then
GoTo 10000
End If
End If
Next l0204
GoTo 10001
10000 :
l020C = Mid(l01F8, l01FC + 2, Len(l01F8))
l0210 = InStr(l020C, Chr(0))
If l0210 Then
l020C = Mid(l020C, 1, l0210 - 1)
End If

687 :
DoEvents
l0214 = InStr(l020C, Chr(9))
If l0214 Then
l0218 = Mid(l020C, 1, l0214 - 1)
l021C = Mid(l020C, l0214 + 1, Len(l020C))
l020C = l0218 + l021C
GoTo 687
End If
If l020C = "" Then GoTo 10001
If l020C = "«‹- -– Fate Ultra Encryption Chat Enabled –-- -›»" Then GoTo 10001
If l020C = "«‹- -– Fate Ultra Encryption Chat Disabled –-- -›»" Then GoTo 10001
l0220 = ""
l0224 = 1
For l0228 = 1 To Len(l020C)
l022C = Mid(l020C, l0228, 1)
l0230 = Asc(l022C)
l0234 = Mid(encryptchat.Text3.Text, l0224, 1)
l0224 = l0224 + 1
If l0224 > Len(encryptchat.Text3.Text) Then l0224 = 1
l0238 = Asc(l0234)
l023C = l0230 - l0238
If l023C < 1 Then
l0240 = l023C + 255
l0244 = Chr(l0240)
Else
l0244 = Chr(l023C)
End If
l0220 = l0220 + l0244
DoEvents
Next l0228
l0248 = Chr(13) + Chr(10)
encryptchat.Text1.Text = encryptchat.Text1.Text + l0248 + l0200 + ":" + Chr(9) + l0220
encryptchat.Text1.SelStart = Len(encryptchat.Text1.Text)
10001 :
l01F0 = Mid(l01F0, l01F4 + 1, Len(l01F0))
GoTo 10002
End If
If encryptchat.Timer1.Interval = 0 Then GoTo L3A4C
encryptchat.Text6.Text = l01F0
l01F8 = l01F0
l01FC = InStr(l01F8, ":")
l0200 = Mid(l01F8, 1, l01FC - 1)
If l0200 = l01CA Then GoTo 10005
For l0204 = 0 To encryptchat.List1.ListCount - 1
l0208 = InStr(1, encryptchat.List1.List(l0204), l0200, 1)
If l0208 Then
If Len(encryptchat.List1.List(l0204)) = Len(l0200) Then
GoTo 10005
End If
End If
Next l0204
GoTo 10004
10005 :
l020C = Mid(l01F8, l01FC + 2, Len(l01F8))
l0210 = InStr(l020C, Chr(0))
If l0210 Then
l020C = Mid(l020C, 1, l0210 - 1)
End If

697 :
DoEvents
l0214 = InStr(l020C, Chr(9))
If l0214 Then
l0218 = Mid(l020C, 1, l0214 - 1)
l021C = Mid(l020C, l0214 + 1, Len(l020C))
l020C = l0218 + l021C
GoTo 697
End If
If l020C = "" Then GoTo 10004
If l020C = "«‹- -– Fate Ultra Encryption Chat Enabled –-- -›»" Then GoTo 10004
If l020C = "«‹- -– Fate Ultra Encryption Chat Disabled –-- -›»" Then GoTo 10004
l0220 = ""
l0224 = 1
For l0228 = 1 To Len(l020C)
l022C = Mid(l020C, l0228, 1)
l0230 = Asc(l022C)
l0234 = Mid(encryptchat.Text3.Text, l0224, 1)
l0224 = l0224 + 1
If l0224 > Len(encryptchat.Text3.Text) Then l0224 = 1
l0238 = Asc(l0234)
l023C = l0230 - l0238
If l023C < 1 Then
l0240 = l023C + 255
l0244 = Chr(l0240)
Else
l0244 = Chr(l023C)
End If
l0220 = l0220 + l0244
DoEvents
Next l0228
l0248 = Chr(13) + Chr(10)
encryptchat.Text1.Text = encryptchat.Text1.Text + l0248 + l0200 + ":" + Chr(9) + l0220
encryptchat.Text1.SelStart = Len(encryptchat.Text1.Text)

10004 :
encryptchat.Text5.Text = l01E2$
DoEvents
If encryptchat.Timer1.Interval = 0 Then GoTo L3A4C
GoTo 10003

L3A4C:
encryptchat.Text5.Text = ""
encryptchat.Text6.Text = ""
encryptchat.Text7.Text = ""
End Function

Function fn2F8 () As Variant
fn2F8 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(69) + Chr(110) + Chr(99) + Chr(114) + Chr(121) + Chr(112) + Chr(116) + Chr(105) + Chr(111) + Chr(110) + Chr(32) + Chr(67) + Chr(104) + Chr(97) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn330 () As Variant
fn330 = "«‹- -– "
End Function

Function fn368 () As Variant
fn368 = " –-- -›»"
End Function

Function fn3D8 () As Variant
End Function

Function fn410 () As Variant
fn410 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(77) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(73) + Chr(110) + Chr(115) + Chr(116) + Chr(97) + Chr(110) + Chr(116) + Chr(32) + Chr(77) + Chr(101) + Chr(115) + Chr(115) + Chr(97) + Chr(103) + Chr(101) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn448 () As Variant
Dim l026A As Variant
Dim l0270 As Variant
Dim l0276 As Variant
Dim l027A As Variant
Dim l027E As Variant
Dim l0282 As String
Dim l0284 As Variant
Dim l0288 As Variant
Dim l028C As Variant
Dim l0290 As Variant
Dim l0294 As Variant
Dim l029A As Variant
Dim l029E
Dim l02A2 As Variant
Dim l02A6 As Variant
Dim l02AA$
Dim l02AC
Dim l02B0
Dim l02B4

l026A = extfn1018("AOL FRAME25", 0&)
l0270 = extfn1790(l026A)
l0276 = extfn17C8(l0270, 6)
l027A = fn138()
If l027A = 5 Then
l026A = extfn1018("AOL FRAME25", 0&)
l0270 = fn5D0("Mem&bers", "Send an Instant Message")
78 :
l0276 = extfn10F8(l026A, "Send Instant Message")
DoEvents
If l0276 = 0 Then GoTo 78
l027A = extfn10C0(l0276, "_AOL_EDIT")
l027E = extfn11A0(l027A, 12, 0, "$im_Off")
l0282$ = String(Len("$im_Off"), 0)
'l0284 = extfn11A0(l027A, 13, Len("$im_off") + 1, l0282$) + l0282$ <> "$im_Off"


l0288 = extfn10F8(l0276, "Send")
l028C = extfn1168(l0288, 3)
l0290 = extfn11A0(l028C, 12, 0, "Fate Ultra 9.0")
l0294 = extfn10F8(l0276, "Send")
l029A = fn1480(l0294)
2001 :
DoEvents
l029E = extfn1018("#32770", "America Online")
If l029E = 0 Then GoTo 2001
l02A2 = extfn10F8(l029E, "OK")
l02A6 = extfn1168(l02A2, 1)
l02AA$ = String(11, 0)
l02AC = extfn11A0(l02A6, 13, 12, l02AA$)
If l02AA$ = "You are now" Then
l02B0 = fn1480(l02A2)
l02B4 = extfnFE0(l0276, 16, 0, 0)
End If
Else
l026A = extfn1018("AOL FRAME25", 0&)
l0270 = fn5D0("Mem&bers", "Send an Instant Message")
110 :
DoEvents
l0276 = extfn10F8(l026A, "Send Instant Message")
If l0276 = 0 Then GoTo 110
l027A = extfn10C0(l0276, "_AOL_EDIT")
l027E = extfn11A0(l027A, 12, 0, "$im_Off")
l0282$ = String(Len("$im_Off"), 0)
l0284 = extfn11A0(l027A, 13, Len("$im_Off") + 1, l0282$)



l028C = extfn10C0(l0276, "RICHCNTL")
l0290 = extfn11A0(l028C, 12, 0, "Fate Ultra 9.0")
l0294 = extfn1168(l028C, 2)
l029A = fn1480(l0294)
2002 :
DoEvents
l029E = extfn1018("#32770", "America Online")
If l029E = 0 Then GoTo 2002
l02A2 = extfn10F8(l029E, "OK")
l02B0 = fn1480(l02A2)
l02B4 = extfnFE0(l0276, 16, 0, 0)
End If
End Function

Function fn480 () As Variant
Dim l02BC As Variant
Dim l02C0 As Variant
Dim l02C4 As Variant
Dim l02C8 As Variant
Dim l02CC As Variant
Dim l02D0 As String
Dim l02D2 As Variant
Dim l02D6
Dim l02DA
Dim l02DE
Dim l02E2
Dim l02E6
Dim l02EA
Dim l02EE
Dim l02F2
Dim l02F6$
Dim l02F8
Dim l02FC
Dim l0300
l02BC = extfn1018("AOL FRAME25", 0&)
l02C0 = extfn1790(l02BC)
l02C4 = extfn17C8(l02C0, 6)
l02C8 = fn138()
If l02C8 = 5 Then
l02BC = extfn1018("AOL FRAME25", 0&)
l02C0 = fn5D0("Mem&bers", "Send an Instant Message")
88 :
l02C4 = extfn10F8(l02BC, "Send Instant Message")
DoEvents
If l02C4 = 0 Then GoTo 88
l02C8 = extfn10C0(l02C4, "_AOL_EDIT")
l02CC = extfn11A0(l02C8, 12, 0, "$im_On")
l02D0$ = String(Len("$im_on"), 0)
l02D2 = extfn11A0(l02C8, 13, Len("$im_on") + 1, l02D0$)



l02D6 = extfn10F8(l02C4, "Send")
l02DA = extfn1168(l02D6, 3)
l02DE = extfn11A0(l02DA, 12, 0, "Fate Ultra 9.0")
l02E2 = extfn10F8(l02C4, "Send")
l02E6 = fn1480(l02E2)
2003 :
DoEvents
l02EA = extfn1018("#32770", "America Online")
If l02EA = 0 Then GoTo 2003
l02EE = extfn10F8(l02EA, "OK")
l02F2 = extfn1168(l02EE, 1)
l02F6$ = String(11, 0)
l02F8 = extfn11A0(l02F2, 13, 12, l02F6$)
If l02F6$ = "You are no " Then
l02FC = fn1480(l02EE)
l0300 = extfnFE0(l02C4, 16, 0, 0)
End If
Else
l02BC = extfn1018("AOL FRAME25", 0&)
l02C0 = fn5D0("Mem&bers", "Send an Instant Message")
112 :
DoEvents
l02C4 = extfn10F8(l02BC, "Send Instant Message")
If l02C4 = 0 Then GoTo 112
l02C8 = extfn10C0(l02C4, "_AOL_EDIT")
l02CC = extfn11A0(l02C8, 12, 0, "$im_On")
l02D0$ = String(Len("$im_on"), 0)
l02D2 = extfn11A0(l02C8, 13, Len("$im_on") + 1, l02D0$)
'Print l02D0$ <> "$im_On" +<missed arg> > <missed arg>
l02DA = extfn10C0(l02C4, "RICHCNTL")
l02DE = extfn11A0(l02DA, 12, 0, "Fate Ultra 9.0")
l02E2 = extfn1168(l02DA, 2)
l02E6 = fn1480(l02E2)
2007 :
DoEvents
l02EA = extfn1018("#32770", "America Online")
If l02EA = 0 Then GoTo 2007
l02EE = extfn10F8(l02EA, "OK")
l02FC = fn1480(l02EE)
l0300 = extfnFE0(l02C4, 16, 0, 0)
End If
End Function

Function fn4B8 () As Variant
Dim l0308 As Variant
Dim l030C As Variant
Dim l0310 As String
Dim l0312 As Variant
Dim l0316 As Variant
Dim l031A As Variant
Dim l0322 As Variant
Dim l0326 As Variant
Dim l032A As Variant
Dim l032E As String
Dim l0330 As Variant
Dim l0334 As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As Variant
Dim l0350 As Variant
Dim l0354 As Variant
Dim l0358 As Variant
Dim l035C As Variant
Dim l0360 As Variant
l0308 = extfn1018("AOL FRAME25", 0&)
l030C = extfn10F8(l0308, "Welcome")
l0310$ = String(30, 0)
l0312 = extfnFA8(l030C, l0310$, 250)
l0316 = Mid(l0310$, 10, 10)
l031A = InStr(l0316, "!")
If l031A Then
l0316 = Mid(l0316, 1, l031A - 1)
End If

232 :
If Lamesense.Command3D1.Caption = "Start" Then GoTo L5AC0
DoEvents
l0322 = extfn1018("AOL FRAME25", 0&)
l0326 = extfn10F8(l0322, "AOL Live!")
l0322 = extfn1088(l0326)
l0326 = extfn10C0(l0322, "_AOL_VIEW")
If l0326 = 0 Then
Exit Function
End If
l032A = extfnFE0(l0326, 14, 0, 0)
If l032A = Len(Lamesense.Text1.Text) Then GoTo 232
l032E$ = String(l032A, 0)
l0330 = extfn11A0(l0326, 13, l032A + 1, l032E$)
l0334 = Len(Lamesense.Text1.Text)
l0338 = Len(l032E$) - l0334
On Error Resume Next
l033C = Right(l032E$, l0338)
l033C = Mid(l033C, 3, Len(l033C))

234 :
If Lamesense.Command3D1.Caption = "Start" Then GoTo L5AC0
l0340 = InStr(l033C, Chr(10))
If l0340 Then
l0344 = Mid(l033C, 1, l0340 - 2)
Lamesense.Text2.Text = l0344
l0348 = InStr(Lamesense.Text2.Text, Chr(0))
If l0348 Then
Lamesense.Text2.Text = Mid(Lamesense.Text2.Text, 1, l0348 - 1)
End If
l034C = InStr(Lamesense.Text2, ":")
l0350 = Mid(Lamesense.Text2.Text, 1, l034C - 1)
If l0350 = l0316 Then GoTo 236
If l0350 = "" Then GoTo 232
For l0354 = 0 To Lamesense.List1.ListCount - 1
l0358 = InStr(1, Lamesense.List1.List(l0354), "-", 1)
l035C = Mid(Lamesense.List1.List(l0354), 1, l0358 - 1)
If l035C = l0350 Then
l0360 = Int(Mid(Lamesense.List1.List(l0354), l0358 + 1, Len(Lamesense.List1.List(l0354))))
l0360 = l0360 + 1
Lamesense.List1.List(l0354) = l0350 + "-" + Trim(Str(l0360))
GoTo 236
End If
Next l0354
Lamesense.List1.AddItem l0350 + "-1"

236 :
l033C = Mid(l033C, l0340 + 1, Len(l033C))
GoTo 234
End If
Lamesense.Text2.Text = l033C
l0348 = InStr(Lamesense.Text2.Text, Chr(0))
If l0348 Then
Lamesense.Text2.Text = Mid(Lamesense.Text2.Text, 1, l0348 - 1)
End If
l034C = InStr(Lamesense.Text2, ":")
l0350 = Mid(Lamesense.Text2.Text, 1, l034C - 1)
If l0350 = "" Then GoTo 232
If l0350 = l0316 Then GoTo 235
For l0354 = 0 To Lamesense.List1.ListCount - 1
l0358 = InStr(1, Lamesense.List1.List(l0354), "-", 1)
l035C = Mid(Lamesense.List1.List(l0354), 1, l0358 - 1)
If l035C = l0350 Then
l0360 = Int(Mid(Lamesense.List1.List(l0354), l0358 + 1, Len(Lamesense.List1.List(l0354))))
l0360 = l0360 + 1
Lamesense.List1.List(l0354) = l0350 + "-" + Trim(Str(l0360))
GoTo 235
End If
Next l0354
Lamesense.List1.AddItem l0350 + "-1"
235 :
If Lamesense.Command3D1.Caption = "Start" Then GoTo L5AC0
Lamesense.Text1.Text = l032E$
Lamesense.Text2.Text = ""
DoEvents
GoTo 232
L5AC0:
Lamesense.Text1.Text = ""
Lamesense.Text2.Text = ""
End Function

Function fn4F0 () As Variant
fn4F0 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(91) + Chr(32) + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(57) + Chr(46) + Chr(48) + Chr(32) + Chr(66) + Chr(121) + Chr(32) + Chr(77) + Chr(97) + Chr(71) + Chr(117) + Chr(83) + Chr(32) + Chr(38) + Chr(32) + Chr(70) + Chr(117) + Chr(78) + Chr(71) + Chr(105) + Chr(105) + Chr(32) + Chr(93) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn528 () As Variant
fn528 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(77) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(77) + Chr(97) + Chr(105) + Chr(108) + Chr(101) + Chr(114) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn560 () As Variant
Dim l0370 As Variant
Dim l0374 As Variant
Dim l0378 As Variant
Dim l037E As Variant
Dim l0382 As Variant
Dim l0386 As Variant
Dim l038A As String
Dim l038C As Variant
l0370 = extfn1018("AOL FRAME25", 0&)
l0374 = extfn1790(l0370)
l0378 = extfn17C8(l0374, 0)
l037E = extfn1800(l0378, 0)
l0382 = extfn1280(l0374)
l0386 = extfn1280(l0378)
l038A$ = String(1, 0)
l038C = extfnF38(l0374, 0, l038A$, 100, &H400)
If l0382 = 8 And l038A$ = "*" Then
fn560 = "1"
Else
fn560 = "0"
End If
End Function

Function fn5D0 (ByVal p0394 As String, p0396 As String) As Variant
Dim l0398 As Variant
Dim l039C As Variant
Dim l03A0 As String
Dim l03A2 As Variant
Dim l03A6 As Variant
Dim l03AA As Variant
Dim l03AE As Variant
Dim l03B2 As Variant
Dim l03B6 As String
Dim l03B8 As Variant
Dim l03BC As Variant
Dim l03C0 As Variant
Dim l03C4 As Variant
Dim l03C8 As Variant
Dim l03CC As Variant
l0398 = extfn1018("AOL FRAME25", 0&)
l039C = extfn1790(l0398)
l03A0$ = String(255, 0)
l03A2 = 0
Do Until 2 > 3
l03A6 = extfnF38(l039C, l03A2, l03A0$, 200, &H400)
l03AA = Mid(l03A0$, 1, l03A6)
If l03AA = p0394$ Then
Exit Do
Else
l03A2 = l03A2 + 1
End If
DoEvents
Loop
l03AE = extfn17C8(l039C, l03A2)
l03B2 = 0
l03B6$ = String(255, 0)
Do Until 2 > 3
l03B8 = extfnF38(l03AE, l03B2, l03B6$, 200, &H400)
l03BC = Mid(l03B6$, 1, l03B8)
l03C0 = InStr(l03BC, Chr(9))
If l03C0 Then
l03BC = Mid(l03BC, 1, l03C0 - 1)
End If
If l03BC = p0396$ Then
Exit Do
Else
l03B2 = l03B2 + 1
End If
Loop
l03C4 = extfn17C8(l039C, l03A2)
l03C8 = extfn1800(l03C4, l03B2)
l03CC = extfnFE0(l0398, 273, l03C8, 0)
End Function

Function fn608 () As Variant
fn608 = fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(121) + Chr(32) + Chr(77) + Chr(97) + Chr(71) + Chr(117) + Chr(83) + fn368()
End Function

Function fn640 () As Variant
fn640 = fn330() + Chr(84) + Chr(104) + Chr(101) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(103) + Chr(111) + Chr(105) + Chr(110) + Chr(103) + Chr(32) + Chr(111) + Chr(117) + Chr(116) + Chr(32) + Chr(105) + Chr(110) + Chr(32) + Chr(40)
End Function

Function fn678 () As Variant
fn678 = Chr(41) + Chr(32) + Chr(109) + Chr(105) + Chr(110) + Chr(115) + fn368()
End Function

Function fn6B0 () As Variant
fn6B0 = fn330() + Chr(84) + Chr(121) + Chr(112) + Chr(101) + Chr(32) + Chr(34)
End Function

Function fn6E8 () As Variant
fn6E8 = Chr(34) + Chr(32) + Chr(84) + Chr(111) + Chr(32) + Chr(103) + Chr(101) + Chr(116) + Chr(32) + Chr(111) + Chr(110) + fn368()
End Function

Function fn720 () As Variant
fn720 = fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(77) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(77) + Chr(97) + Chr(105) + Chr(108) + Chr(32) + Chr(72) + Chr(97) + Chr(115) + Chr(32) + Chr(66) + Chr(101) + Chr(101) + Chr(110) + Chr(32) + Chr(80) + Chr(97) + Chr(117) + Chr(115) + Chr(101) + Chr(100) + fn368()
End Function

Function fn758 () As Variant
fn758 = fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + Chr(32) + Chr(104) + Chr(97) + Chr(115) + Chr(32) + Chr(115) + Chr(116) + Chr(97) + Chr(114) + Chr(116) + Chr(101) + Chr(100) + Chr(33) + fn368()
End Function

Function fn790 () As Variant
fn790 = fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(40) + Chr(86) + Chr(41) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(121) + Chr(32) + Chr(77) + Chr(97) + Chr(71) + Chr(117) + Chr(83) + fn368()
End Function

Function fn7C8 () As Variant
Dim l03F4 As Variant
Dim l03F8 As String
Dim l03FA As Variant
Dim l03FE As Variant
l03F4 = extfn1018("_AOL_MODAL", 0&)
l03F8$ = String(255, 0)
l03FA = extfnFA8(l03F4, l03F8$, 250)
l03FE = InStr(l03F8$, Chr(0))
If l03FE Then
l03F8$ = Mid(l03F8$, 1, l03FE - 1)
End If
fn7C8 = l03F8$
End Function

Function fn800 () As Variant
Dim l0406 As Variant
Dim l040A As Variant
Dim l040E As String
Dim l0410 As Variant
l0406 = extfn1018("AOL FRAME25", 0&)
l040A = extfn10F8(l0406, "Welcome")
l040E$ = String$(30, 0)
l0410 = extfnFA8(l040A, l040E$, 250)
If l0410 <= 7 Then
MsgBox "Not Signed On!", 16, "Fate"
Exit Function
End If
End Function

Function fn8A8 () As Variant
fn8A8 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(73) + Chr(110) + Chr(115) + Chr(116) + Chr(97) + Chr(110) + Chr(116) + Chr(32) + Chr(77) + Chr(101) + Chr(115) + Chr(115) + Chr(97) + Chr(103) + Chr(101) + Chr(32) + Chr(66) + Chr(111) + Chr(109) + Chr(98) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn8E0 () As Variant
Dim l041C As Variant
Dim l0420 As Variant
Dim l0424 As String
Dim l0426 As Variant
Dim l042A As Variant
Dim l042E As Variant
Dim l0436 As Variant
Dim l043A As Variant
Dim l043E As Variant
Dim l0442 As String
Dim l0444 As Variant
Dim l0448 As Variant
Dim l044C As Variant
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
Dim l0478 As Variant
Dim l047C As Variant
Dim l0480 As Variant
Dim l0484 As Variant
Dim l0488 As Variant
Dim l048C As String
Dim l048E As Variant
Dim l0492 As Variant
Dim l0496 As Variant
Dim l049A As Variant
Dim l049E As Variant
Dim l04A2 As Variant
Dim l04A6 As Variant
Dim l04AA As Variant
Dim l04AE As Variant
Dim l04B2 As Variant
Dim l04B6 As Variant
Dim l04BA As Variant
l041C = extfn1018("AOL FRAME25", 0&)
l0420 = extfn10F8(l041C, "Welcome")
l0424$ = String(30, 0)
l0426 = extfnFA8(l0420, l0424$, 250)
l042A = Mid(l0424$, 10, 10)
l042E = InStr(l042A, "!")
If l042E Then
l042A = Mid(l042A, 1, l042E - 1)
End If

42 :
If warezre.Command3D1.Caption = "Start" Then GoTo L85F2
DoEvents
l0436 = extfn1018("AOL FRAME25", 0&)
l043A = extfn10F8(l0436, "AOL Live!")
l0436 = extfn1088(l043A)
l043A = extfn10C0(l0436, "_AOL_VIEW")
If l043A = 0 Then
Exit Function
End If
l043E = extfnFE0(l043A, 14, 0, 0)
If l043E = Len(warezre.Text1.Text) Then GoTo 42
l0442$ = String(l043E, 0)
l0444 = extfn11A0(l043A, 13, l043E, l0442$)
l0448 = Len(warezre.Text1.Text)
l044C = Len(l0442$) - l0448
On Error Resume Next
l0450 = Right(l0442$, l044C)
l0450 = Mid(l0450, 3, Len(l0450))
44 :
If warezre.Command3D1.Caption = "Start" Then GoTo L85F2
l0454 = InStr(l0450, Chr(10))
If l0454 Then
l0458 = Mid(l0450, 1, l0454 - 2)
warezre.Text2.Text = l0458
If warezre.Text2.Text <> "" Then warezre.Text3.Text = warezre.Text2.Text
l045C = InStr(warezre.Text3.Text, Chr(0))
If l045C Then
warezre.Text3.Text = Mid(warezre.Text3.Text, 1, l045C - 1)
End If
l0460 = InStr(1, warezre.Text3.Text, warezre.got.Text, 1)
If l0460 Then
l0464 = Mid(warezre.Text3.Text, l0460 - 2, 2)
If l0464 <> ":" + Chr(9) Then
GoTo 46
End If
l0464 = InStr(warezre.Text3.Text, ":")
l0468 = Mid(warezre.Text3.Text, 1, l0464 - 1)
If l0468 = l042A Then GoTo 46
For l046C = 0 To warezre.List1.ListCount - 1
l0470 = InStr(1, warezre.List1.List(l046C), l0468, 1)
If l0470 Then GoTo 46
Next l046C
If l0468 = l042A Then GoTo 46
warezre.List1.AddItem l0468
If warezre.Option3D1.Value = True Then
l0474 = fn138()
If l0474 = 5 Then
l0478 = fn5D0("Mem&bers", "Send an Instant Message")
l047C = extfn1018("AOL FRAME25", 0&)
89 :
DoEvents
l0480 = extfn10F8(l047C, "Send Instant Message")
If l0480 = 0 Then GoTo 89
l0484 = extfn10C0(l0480, "_AOL_EDIT")
l0488 = extfn11A0(l0484, 12, 0, l0468)
l048C$ = String(Len(l0468), 0)
l048E = extfn11A0(l0484, 13, Len(l0468) + 1, l048C$)
If l048C$ <> l0468 Then GoTo 89
l0492 = extfn1168(l0484, 2)
l0496 = extfn11A0(l0492, 12, 0, warezre.Text4.Text)
l049A = extfn1168(l0492, 2)
l049E = fn1480(l049A)
67 :
DoEvents
l04A2 = extfn10F8(l0480, "Available?")
If l04A2 > 0 Then
l04A6 = extfn1018("#32770", "America Online")
If l04A6 > 0 Then
l04AA = extfn10F8(l04A6, "OK")
l04AE = fn1480(l04AA)
l04B2 = extfnFE0(extfn1088(l04A2), 16, 0, 0)
GoTo 46
End If
GoTo 67
Else
GoTo 46
End If
Else
l0478 = fn5D0("Mem&bers", "Send an Instant Message")
l047C = extfn1018("AOL FRAME25", 0&)
83 :
DoEvents
l0480 = extfn10F8(l047C, "Send Instant Message")
If l0480 = 0 Then GoTo 83
l0484 = extfn10C0(l0480, "_AOL_EDIT")
l0488 = extfn11A0(l0484, 12, 0, l0468)
l048C$ = String(Len(l0468), 0)
l048E = extfn11A0(l0484, 13, Len(l0468) + 1, l048C$)
If l048C$ <> l0468 Then GoTo 83
l0492 = extfn10C0(l0480, "RICHCNTL")
l0496 = extfn11A0(l0492, 12, 0, warezre.Text4.Text)
l049A = extfn1168(l0492, 2)
l049E = fn1480(l049A)
68 :
DoEvents
l04A2 = extfn10F8(l0480, "To:")
If l04A2 > 0 Then
l04A6 = extfn1018("#32770", "America Online")
If l04A6 > 0 Then
l04AA = extfn10F8(l04A6, "OK")
l04AE = fn1480(l04AA)
l04B2 = extfnFE0(extfn1088(l04A2), 16, 0, 0)
GoTo 46
End If
GoTo 68
End If
End If
Else
l04B6 = fn13A0(.6)
l04BA = fnA68("«‹- -– <" + l0468 + "> " + warezre.Text4.Text + " –-- -›»")
End If
End If
46 :
l0450 = Mid(l0450, l0454 + 1, Len(l0450))
GoTo 44
End If
warezre.Text2.Text = l0450
If warezre.Text2.Text <> "" Then warezre.Text3.Text = warezre.Text2.Text
l045C = InStr(warezre.Text3.Text, Chr(0))
If l045C Then
warezre.Text3.Text = Mid(warezre.Text3.Text, 1, l045C - 1)
End If
l0460 = InStr(1, warezre.Text3.Text, warezre.got.Text, 1)
If l0460 Then
l0464 = Mid(warezre.Text3.Text, l0460 - 2, 2)
If l0464 <> ":" + Chr(9) Then
GoTo 45
End If
l0464 = InStr(warezre.Text3, ":")
l0468 = Mid(warezre.Text3.Text, 1, l0464 - 1)
If l0468 = l042A Then GoTo 45
For l046C = 0 To warezre.List1.ListCount - 1
l0470 = InStr(1, warezre.List1.List(l046C), l0468, 1)
If l0470 Then GoTo 45
Next l046C
If l0468 = l042A Then GoTo 45
warezre.List1.AddItem l0468
If warezre.Option3D1.Value = True Then
l0474 = fn138()
If l0474 = 5 Then
l0478 = fn5D0("Mem&bers", "Send an Instant Message")
l047C = extfn1018("AOL FRAME25", 0&)
899 :
DoEvents
l0480 = extfn10F8(l047C, "Send Instant Message")
If l0480 = 0 Then GoTo 899
l0484 = extfn10C0(l0480, "_AOL_EDIT")
l0488 = extfn11A0(l0484, 12, 0, l0468)
l048C$ = String(Len(l0468), 0)
l048E = extfn11A0(l0484, 13, Len(l0468) + 1, l048C$)
If l048C$ <> l0468 Then GoTo 899
l0492 = extfn1168(l0484, 2)
l0496 = extfn11A0(l0492, 12, 0, warezre.Text4.Text)
l049A = extfn1168(l0492, 2)
l049E = fn1480(l049A)
677 :
DoEvents
l04A2 = extfn10F8(l0480, "Available?")
If l04A2 > 0 Then
l04A6 = extfn1018("#32770", "America Online")
If l04A6 > 0 Then
l04AA = extfn10F8(l04A6, "OK")
l04AE = fn1480(l04AA)
l04B2 = extfnFE0(extfn1088(l04A2), 16, 0, 0)
GoTo 46
End If
GoTo 677
Else
GoTo 46
End If
Else
l0478 = fn5D0("Mem&bers", "Send an Instant Message")
l047C = extfn1018("AOL FRAME25", 0&)
833 :
DoEvents
l0480 = extfn10F8(l047C, "Send Instant Message")
If l0480 = 0 Then GoTo 833
l0484 = extfn10C0(l0480, "_AOL_EDIT")
l0488 = extfn11A0(l0484, 12, 0, l0468)
l048C$ = String(Len(l0468), 0)
l048E = extfn11A0(l0484, 13, Len(l0468) + 1, l048C$)
If l048C$ <> l0468 Then GoTo 833
l0492 = extfn10C0(l0480, "RICHCNTL")
l0496 = extfn11A0(l0492, 12, 0, warezre.Text4.Text)
l049A = extfn1168(l0492, 2)
l049E = fn1480(l049A)
688 :
DoEvents
l04A2 = extfn10F8(l0480, "To:")
If l04A2 > 0 Then
l04A6 = extfn1018("#32770", "America Online")
If l04A6 > 0 Then
l04AA = extfn10F8(l04A6, "OK")
l04AE = fn1480(l04AA)
l04B2 = extfnFE0(extfn1088(l04A2), 16, 0, 0)
GoTo 46
End If
GoTo 688
End If
End If
Else
l04B6 = fn13A0(.6)
l04BA = fnA68("«‹- -– <" + l0468 + "> " + warezre.Text4.Text + " –-- -›»")
End If
End If
45 :
If warezre.Command3D1.Caption = "Start" Then GoTo L85F2
warezre.Text1.Text = l0442$
DoEvents
GoTo 42

L85F2:
warezre.Text1.Text = ""
warezre.Text2.Text = ""
warezre.Text3.Text = ""
End Function

Function fn918 () As Variant
End Function

Function fn950 () As Variant
fn950 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(82) + Chr(111) + Chr(111) + Chr(109) + Chr(32) + Chr(70) + Chr(105) + Chr(110) + Chr(100) + Chr(101) + Chr(114) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn988 () As Variant
fn988 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(83) + Chr(99) + Chr(114) + Chr(97) + Chr(109) + Chr(98) + Chr(108) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fn9C0 () As Variant
Dim l04CE As Variant
Dim l04D2 As Variant
Dim l04D6 As String
Dim l04D8 As Variant
Dim l04DC As Variant
Dim l04E0 As Variant
Dim l04E8 As Variant
Dim l04EC As Variant
Dim l04F0 As Variant
Dim l04F4 As String
Dim l04F6 As Variant
Dim l04FA As Variant
Dim l04FE As Variant
Dim l0502 As Variant
Dim l0506 As Variant
Dim l050A As Variant
Dim l050E As Variant
Dim l0512 As Variant
Dim l0516 As Variant
Dim l051A As Variant
Dim l051E As Variant
Dim l0522 As Variant
Dim l0526 As Variant
Dim l052A As Variant
Dim l052E As Variant
Dim l0532 As Variant
Dim l0536 As Variant
Dim l053A As Variant
Dim l053E As Variant
Dim l0542 As Variant
Dim l0546 As Variant
Dim l054A As Variant
Dim l054E As Variant
Dim l0552 As Variant
l04CE = extfn1018("AOL FRAME25", 0&)
l04D2 = extfn10F8(l04CE, "Welcome")
l04D6$ = String(30, 0)
l04D8 = extfnFA8(l04D2, l04D6$, 250)
l04DC = Mid(l04D6$, 10, 10)
l04E0 = InStr(l04DC, "!")
If l04E0 Then
l04DC = Mid(l04DC, 1, l04E0 - 1)
End If

322 :
If scram.Command3D2.Caption = "Start" Then GoTo L9C76
DoEvents
l04E8 = extfn1018("AOL FRAME25", 0&)
l04EC = extfn10F8(l04E8, "AOL Live!")
l04E8 = extfn1088(l04EC)
l04EC = extfn10C0(l04E8, "_AOL_VIEW")
If l04EC = 0 Then
Exit Function
End If
l04F0 = extfnFE0(l04EC, 14, 0, 0)
If l04F0 = Len(scram.Text3.Text) Then GoTo 322
l04F4$ = String(l04F0, 0)
l04F6 = extfn11A0(l04EC, 13, l04F0, l04F4$)
l04FA = Len(scram.Text3.Text)
l04FE = Len(l04F4$) - l04FA
On Error Resume Next
l0502 = Right(l04F4$, l04FE)
l0502 = Mid(l0502, 3, Len(l0502))
344 :
If scram.Command3D2.Caption = "Start" Then GoTo L9C76
l0506 = InStr(l0502, Chr(10))
If l0506 Then
l050A = Mid(l0502, 1, l0506 - 2)
scram.Text4.Text = l050A
l050E = InStr(scram.Text4.Text, Chr(0))
If l050E Then
scram.Text4.Text = Mid(scram.Text4.Text, 1, l050E - 1)
End If
l0512 = InStr(1, scram.Text4.Text, scram.Text1.Text, 1)
If l0512 Then
l0516 = Mid(scram.Text4.Text, l0512 - 2, 2)
If l0516 <> ":" + Chr(9) Then GoTo 366
l0516 = InStr(scram.Text4, ":")
l051A = Mid(scram.Text4.Text, 1, l0516 - 1)
l051E = Mid(scram.Text4.Text, l0516 + 2, Len(scram.Text4.Text))
If Len(l051E) <> Len(scram.Text1.Text) Then GoTo 366
If l051A = l04DC Then GoTo 366
l0522 = fnA30("«‹- -– <" + l051A + "> Answered Correctly –-- -›»")
For l0526 = 0 To scram.List2.ListCount - 1
l052A = InStr(1, scram.List2.List(l0526), l051A, 1)
If l052A Then
If Len(l051A) = Len(scram.List2.List(l0526)) Then
l052E = scram.List1.List(l0526)
l0532 = Int(l052E) + 1
l0536 = scram.List2.List(l0526)
If scram.List1.ListCount = 0 Then
GoTo 96
End If
If scram.List1.ListCount = 1 Then
scram.List1.List(0) = l0532
GoTo 91
End If
If l0526 = 0 Then
scram.List1.List(l0526) = l0532
GoTo 91
End If
scram.List2.RemoveItem l0526
scram.List1.RemoveItem l0526
l053A = 1
For l053E = 0 To scram.List1.ListCount - 1
l0542 = scram.List1.List(l053E)
If Int(l0532) = Int(l0542) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 91
End If
If Int(l0532) > Int(scram.List1.List(l053E)) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 91
End If
l053A = 0
92 :
DoEvents
On Error Resume Next
If Int(l0532) < Int(scram.List1.List(l053E + l053A)) Then
If Int(l0532) > Int(scram.List1.List(l053E)) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 91
End If
If Err Then
scram.List1.AddItem Int(l0532), scram.List1.ListCount
scram.List2.AddItem l0536, scram.List2.ListCount
GoTo 91
End If
l053A = l053A + 1
GoTo 92
End If
Next l053E
End If
Else
GoTo 90
End If
90 :
Next l0526

96 :
scram.List2.AddItem l051A
l0546 = 1
scram.List1.AddItem "1"

91 :
scram.List3.Clear
For l054A = 0 To scram.List2.ListCount - 1
l054E = scram.List2.List(l054A)
l0552 = scram.List1.List(l054A)
scram.List3.AddItem l054E + "   -   " + l0552
Next l054A
scram.Command3D2.Caption = "Start"
scram.Timer1.Interval = 0
scram.Timer1.Enabled = False
scram.Timer2.Interval = 0
scram.Timer2.Enabled = False
GoTo L9C76
End If
366 :
l0502 = Mid(l0502, l0506 + 1, Len(l0502))
GoTo 344
End If
scram.Text4.Text = l0502
l050E = InStr(scram.Text4.Text, Chr(0))
If l050E Then
scram.Text4.Text = Mid(scram.Text4.Text, 1, l050E - 1)
End If
l0512 = InStr(1, scram.Text4.Text, scram.Text1.Text, 1)
If l0512 Then
l0516 = Mid(scram.Text4.Text, l0512 - 2, 2)
If l0516 <> ":" + Chr(9) Then GoTo 355
l0516 = InStr(scram.Text4, ":")
l051A = Mid(scram.Text4.Text, 1, l0516 - 1)
l051E = Mid(scram.Text4.Text, l0516 + 2, Len(scram.Text4.Text))
If Len(l051E) <> Len(scram.Text1.Text) Then GoTo 355
If l051A = l04DC Then GoTo 355
l0522 = fnA30("«‹- -– <" + l051A + "> Answered Correctly –-- -›»")
For l0526 = 0 To scram.List2.ListCount - 1
l052A = InStr(1, scram.List2.List(l0526), l051A, 1)
If l052A Then
If Len(l051A) = Len(scram.List2.List(l0526)) Then
l052E = scram.List1.List(l0526)
l0532 = Int(l052E) + 1
If scram.List1.ListCount = 0 Then
GoTo 97
End If
If scram.List1.ListCount = 1 Then
scram.List1.List(0) = l0532
GoTo 93
End If
If l0526 = 0 Then
scram.List1.List(l0526) = l0532
GoTo 93
End If
l0536 = scram.List2.List(l0526)
scram.List2.RemoveItem l0526
scram.List1.RemoveItem l0526
l053A = 1
For l053E = 0 To scram.List1.ListCount - 1
l0542 = scram.List1.List(l053E)
If Int(l0532) = Int(l0542) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 93
End If
If Int(l0532) > Int(scram.List1.List(l053E)) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 93
End If
l053A = 0
95 :
DoEvents
On Error Resume Next
If Int(l0532) < Int(scram.List1.List(l053E + l053A)) Then
If Int(l0532) > Int(scram.List1.List(l053E)) Then
scram.List1.AddItem Int(l0532), l053E
scram.List2.AddItem l0536, l053E
GoTo 93
End If
If Err Then
scram.List1.AddItem Int(l0532), scram.List1.ListCount
scram.List2.AddItem l0536, scram.List2.ListCount
GoTo 93
End If
l053A = l053A + 1
GoTo 95
End If
Next l053E
End If
Else
GoTo 94
End If
94 :
Next l0526
97 :
scram.List2.AddItem l051A
l0546 = 1
scram.List1.AddItem "1"

93 :
scram.List3.Clear
For l054A = 0 To scram.List2.ListCount - 1
l054E = scram.List2.List(l054A)
l0552 = scram.List1.List(l054A)
scram.List3.AddItem l054E + "   -   " + l0552
Next l054A
DoEvents
scram.Command3D2.Caption = "Start"
scram.Timer1.Interval = 0
scram.Timer2.Interval = 0
scram.Timer1.Enabled = False
scram.Timer2.Enabled = False
GoTo L9C76
End If
355 :
If scram.Command3D2.Caption = "Start" Then GoTo L9C76
scram.Text3.Text = l04F4$
DoEvents
GoTo 322
L9C76:
scram.Text3.Text = ""
scram.Text4.Text = ""
scram.Text4.Text = ""
End Function

Function fn9F8 () As Variant
Dim l055A As Variant
Dim l055E As Variant
Dim l0562 As String
Dim l0564 As Variant
Dim l0568 As Variant
Dim l056C As Variant
l055A = extfn1018("AOL FRAME25", 0&)
l055E = extfn10F8(l055A, "Welcome")
l0562$ = String(30, 0)
l0564 = extfnFA8(l055E, l0562$, 250)
l0568 = Mid(l0562$, 10, 10)
l056C = InStr(l0568, "!")
If l056C Then
l0568 = Mid(l0568, 1, l056C - 1)
End If
fn9F8 = l0568
End Function

Function fnA30 (ByVal p0574 As String) As Variant
Dim l0576 As Variant
Dim l057A As Variant
Dim l057E As Variant
Dim l0582 As Variant
Dim l0586 As String
Dim l0588 As Variant
Dim l058C As Variant
Dim l0590 As Variant
Dim l0594 As Variant
Dim l0598 As Variant
l0576 = extfn1018("AOL FRAME25", 0&)
l057A = extfn10F8(l0576, "AOL Live!")
l057E = extfn1088(l057A)
l0582 = extfn10C0(l057E, "_AOL_EDIT")
l0586$ = String(255, 0)
l0588 = extfn11A0(l0582, 13, 255, l0586$)
l058C = extfn11A0(l0582, 12, 0, p0574)
l0590 = extfn1168(l0582, 2)
l0594 = fn1480(l0590)
l0598 = extfn11A0(l0582, 12, 0, l0586$)
DoEvents
End Function

Function fnA68 (p05A0 As Variant) As Variant
Dim l05A4 As Variant
Dim l05A8 As Variant
Dim l05AC As Variant
Dim l05B0 As Variant
Dim l05B4 As String
Dim l05B6 As Variant
Dim l05BA As Variant
Dim l05BE As Variant
Dim l05C2 As Variant
Dim l05C6 As Variant
l05A4 = extfn1018("AOL FRAME25", 0&)
l05A8 = extfn10F8(l05A4, "AOL Live!")
l05AC = extfn1088(l05A8)
l05B0 = extfn10C0(l05AC, "_AOL_EDIT")
l05B4$ = String(255, 0)
l05B6 = extfn11A0(l05B0, 13, 255, l05B4$)
l05BA = extfn11A0(l05B0, 12, 0, p05A0)
l05BE = extfn1168(l05B0, 2)
l05C2 = fn1480(l05BE)
l05C6 = extfn11A0(l05B0, 12, 0, l05B4$)
DoEvents
End Function

Function fnAA0 (p05CE As String) As Variant
Dim l05D0 As Variant
Dim l05D4 As Variant
Dim l05D8 As Variant
Dim l05DC As Variant
Dim l05E0 As Variant
Dim l05E4 As Variant
Dim l05E8 As Variant
l05D0 = extfn1018("AOL FRAME25", 0&)
l05D4 = extfn10F8(l05D0, "AOL Live!")
l05D8 = extfn1088(l05D4)
l05DC = extfn10C0(l05D8, "_AOL_EDIT")
l05E0 = extfn11A0(l05DC, 12, 0, p05CE)
l05E4 = extfn1168(l05DC, 2)
l05E8 = fn1480(l05E4)
End Function

Function fnAD8 (p05F0 As Variant) As Variant
Dim l05F4 As Variant
Dim l05F8 As Variant
Dim l05FC As Variant
Dim l0600 As Variant
Dim l0604 As Variant
Dim l0608 As Variant
Dim l060C As Variant
l05F4 = extfn1018("AOL FRAME25", 0&)
l05F8 = extfn10F8(l05F4, "AOL Live!")
l05FC = extfn1088(l05F8)
l0600 = extfn10C0(l05FC, "_AOL_EDIT")
l0604 = extfn11A0(l0600, 12, 0, p05F0)
l0608 = extfn1168(l0600, 2)
l060C = fn1480(l0608)
End Function

Function fnB10 () As Variant
Dim l0614 As Variant
Dim l0618 As Variant
Dim l061C As Variant
Dim l0620 As String
Dim l0622 As Variant
Dim l0626 As Variant
Dim l062A As Variant
Dim l0632 As Variant
Dim l0636 As Variant
Dim l063A As Variant
Dim l063E As String
Dim l0640 As Variant
Dim l0644 As Variant
Dim l0648 As Variant
Dim l064C As Variant
Dim l0652 As Variant
Dim l0656 As Variant
Dim l065A As Variant
Dim l065E As Variant
Dim l0662 As Variant
Dim l0666 As Variant
Dim l066A As Variant
Dim l066E As Variant
Dim l0672 As Variant
Dim l0676 As Variant
Dim l067A As Variant
Dim l067E As Variant
Dim l0682 As Variant
Dim l0686 As Variant
Dim l068A As Variant
Dim l068E As Variant
Dim l0692 As Variant
Dim l0696 As Variant
Dim l069A As Variant
Dim l069E As Variant
Dim l06A2 As Variant
Dim l06A6 As Variant
Dim l06AA As Variant
Dim l06AE As Variant
Dim l06B2 As Variant
Dim l06B6 As Variant
Dim l06BA As Variant
Dim l06BE As Variant
Dim l06C2 As Variant
Dim l06C6 As Variant
Dim l06CA As Variant
Dim l06CE As Variant
Dim l06D2 As Variant
Dim l06D6 As Variant
Dim l06DA As Variant
Dim l06DE As Variant
Dim l06E2 As Variant
Dim l06E6 As Variant
Dim l06EA As Variant
Dim l06EE As Variant
Dim l06F2 As Variant
Dim l06F6 As Variant
Dim l06FA As Variant
Dim l06FE As Variant
Dim l0702 As String
Dim l0704 As Variant
Dim l0708 As Variant
Dim l070C As Variant
Dim l0710 As Variant
Dim l0714 As Variant
Dim l0718 As Variant
Dim l071C As Variant
Dim l0720 As Variant
Dim l0724 As Variant
Dim l0728 As Variant
Dim l072C As Variant
Dim l0730 As Variant
Dim l0734 As Variant
Dim l0738 As Variant
Dim l073C As Variant
Dim l0740 As Variant
Dim l0744 As Variant
Dim l0748 As Variant
Dim l074C As Variant
Dim l0750 As Variant
Dim l0754 As Variant
Dim l0758 As String
Dim l075A As Variant
Dim l075E As Variant
Dim l0762 As Variant
l0614 = Chr(13) + Chr(10)
l0618 = extfn1018("AOL FRAME25", 0&)
l061C = extfn10F8(l0618, "Welcome")
l0620$ = String(30, 0)
l0622 = extfnFA8(l061C, l0620$, 250)
l0626 = Mid(l0620$, 10, 10)
l062A = InStr(l0626, "!")
If l062A Then
l0626 = Mid(l0626, 1, l062A - 1)
End If

132 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
DoEvents
l0632 = extfn1018("AOL FRAME25", 0&)
l0636 = extfn10F8(l0632, "AOL Live!")
l0632 = extfn1088(l0636)
l0636 = extfn10C0(l0632, "_AOL_VIEW")
If l0636 = 0 Then
Exit Function
End If
l063A = extfnFE0(l0636, 14, 0, 0)
If l063A = Len(server.Text1.Text) Then GoTo 132
l063E$ = String(l063A, 0)
l0640 = extfn11A0(l0636, 13, l063A, l063E$)
l0644 = Len(server.Text1.Text)
l0648 = Len(l063E$) - l0644
On Error Resume Next
l064C = Right(l063E$, l0648)
l064C = Mid(l064C, 3, Len(l064C))
134 :
If server.Caption <> fn14F0() Then Unload server
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
DoEvents
l0652 = InStr(l064C, Chr(10))
If l0652 Then
l0656 = Mid(l064C, 1, l0652 - 2)
server.Text2.Text = l0656
l065A = InStr(1, server.Text2.Text, "/" + l0626 + " Send List", 1)
If l065A Then
l065E = Mid(server.Text2.Text, l065A - 2, 2)
If l065E <> ":" + Chr(9) Then
GoTo 136
End If
l065E = InStr(server.Text2, ":")
l0662 = Mid(server.Text2.Text, 1, l065E - 1)
For l0666 = 0 To server.List2.ListCount - 1
l066A = InStr(1, l0662, server.List2.List(l0666), 1)
If l066A Then
If Len(l0662) = Len(server.List2.List(l0666)) Then GoTo 136
End If
Next l0666
If l0662 = l0626 Then GoTo 136

1336 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l0636 = extfn1018("AOL FRAME25", 0&)
l063A = fn5D0("&Mail", "&Compose Mail")
10 :
DoEvents
l066E = extfn10F8(l0636, "Compose Mail")
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
If l066E = 0 Then GoTo 10
l0640 = extfn10C0(l066E, "_AOL_EDIT")
l0644 = extfn11A0(l0640, 12, 0, l0662)
l0648 = extfn10F8(l066E, "Subject:")
l0672 = extfn1168(l0648, 2)
If server.Caption <> fn14F0() Then Unload server
l0676 = extfn11A0(l0672, 12, 0, fn330() + "Fate Ultra Warez Server List" + fn368())
l067A = extfn10F8(l066E, "To:")
l065A = extfn1168(l067A, 1)
l065E = extfn11A0(l065A, 12, 0, server.Text5.Text)
l0662 = extfn10F8(l066E, "Send")
l067E = extfn1168(l0662, 2)
l0682 = fn1480(l067E)
500 :
DoEvents
l0686 = extfn1088(l0662)
l068A = extfn10F8(l0636, l0686)
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l0636 = extfn1018("AOL FRAME25", 0&)
l068E = extfn10F8(l0636, "Error")
If l068E > 0 Then
l061C = extfn1018("AOL FRAME25", 0&)
l0692 = extfn10F8(l061C, "Error")
l0696 = extfn10F8(l0692, "OK")
l069A = fn1480(l0696)
l069E = extfn10F8(l061C, "Compose Mail")
l06A2 = extfnFE0(l069E, 16, 0, 0)
GoTo LAF38
If server.Caption <> fn14F0() Then Unload server
End If
If l068A > 0 Then GoTo 500
l06A6 = extfn10F8(l068A, "OK")
l06AA = fn1480(l06A6)
LAF38:
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
End If
l065A = InStr(1, server.Text2.Text, "/" + l0626 + " Send ", 1)
If l065A Then
l065E = Mid(server.Text2.Text, l065A - 2, 2)
If l065E <> ":" + Chr(9) Then
DoEvents
GoTo 136
End If
l065E = InStr(server.Text2, ":")
l0662 = Mid(server.Text2.Text, 1, l065E - 1)
For l0666 = 0 To server.List2.ListCount - 1
l066A = InStr(1, l0662, server.List2.List(l0666), 1)
If l066A Then
If Len(l0662) = Len(server.List2.List(l0666)) Then GoTo 136
End If
Next l0666
l067E = Mid(server.Text2.Text, l065A + Len(l0626) + 7, Len(server.Text2.Text))
l06AE = InStr(1, l067E, "x", 1)
If l06AE Then GoTo 136
l06AE = InStr(1, l067E, "list", 1)
If l06AE Then GoTo 136
If IsNumeric(l067E) Then
For l0682 = 0 To server.List1.ListCount - 1
l068A = InStr(1, server.List1.List(l0682), l0662 + "-" + l067E, 1)
If l068A Then
If Len(server.List1.List(l0682)) = Len(l0662 + "-" + l067E) Then GoTo 1326
End If
Next l0682
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
If l0662 = l0626 Then GoTo 136
If server.Caption <> fn14F0() Then Unload server
server.List1.AddItem l0662 + "-" + l067E
1326 :
l06B2 = Mid(server.Frame3D1.Caption, 9, Len(server.Frame3D1.Caption))
l06B6 = l06B2 + 1
server.Frame3D1.Caption = "Served: " + l06B6
l068E = extfn1018("AOL FRAME25", 0&)
If server.Option3D1.Value = True Then l06BA = "New Mail"
If server.Option3D2.Value = True Then l06BA = "Old Mail"
If server.Option3D3.Value = True Then l06BA = "Incoming FlashMail"
l06BE = extfn10F8(l068E, l06BA)
l06C2 = extfn10C0(l06BE, "_AOL_TREE")
l06C6 = extfnFE0(l06C2, 1036, 0, 0)
If server.Option3D3.Value = True Then
l06CA = extfnFE0(l06C2, 1031, l067E, 0)
l06CE = extfn10F8(l06BE, "Open")
l06D2 = fn1480(l06CE)
GoTo 340
End If
l06CA = extfnFE0(l06C2, 1031, l067E, 0)
l06CE = extfn10F8(l06BE, "Read")
l06D2 = fn1480(l06CE)
340 :
150 :
DoEvents
l06D6 = extfn10F8(l068E, "Forward")
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
If l06D6 = 0 Then GoTo 150
151 :
DoEvents
l06DA = extfn1168(l06D6, 2)
l06DE = fn1480(l06DA)
l06E2 = Timer
While Timer - l06E2 < .7
l06E6 = extfn10F8(l068E, "Address")
If l06E6 > 0 Then GoTo 162
DoEvents
Wend
l06EA = extfn10F8(l068E, "Send Now")
l06E6 = extfn10F8(l068E, "Address")
If l06E6 = 0 Then GoTo 151
162 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l06EA = extfn10F8(l068E, "Send Now")
l06EE = extfn1088(l06E6)
l06F2 = extfn10C0(l06EE, "_AOL_EDIT")
l06F6 = extfn11A0(l06F2, 12, 0, l0662)
l06FA = extfn10F8(l06EE, "Subject:")
l06FE = extfn1168(l06FA, 2)
l0702$ = String(255, 0)
l0704 = extfn11A0(l06FE, 13, 255, l0702$)
l0708 = Mid(l0702$, 6, Len(l0702$))
l070C = extfn11A0(l06FE, 12, 0, l0708)
l0710 = extfn10F8(l06EE, "To:")
l0714 = extfn1168(l0710, 1)
l0718 = extfn1168(l0714, 3)
l071C = extfn11A0(l0718, 12, 0, fn330() + "Fate Ultra WareZ Server By MaGuS" + fn368() + Chr(13) + Chr(10) + "«‹- -– Requested WareZ Index#: " + l067E + " –-- -›»" + l0614 + l0614 + server.Text3.Text)
l0720 = extfn1168(l06EA, 2)
l0724 = fn1480(l0720)
152 :
DoEvents
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l0728 = extfn1088(l06EA)
l072C = extfn10F8(l068E, l0728)
l0730 = extfn10F8(l068E, "Error")
If l0730 > 0 Then
l0734 = extfn10F8(l0730, "OK")
l0738 = fn1480(l0734)
l073C = extfnFE0(l06EE, 16, 0, 0)
l0740 = extfnFE0(extfn1088(l06D6), 16, 0, 0)
GoTo 136
End If
If l072C > 0 Then GoTo 152
l0744 = extfn10F8(l072C, "OK")
l0748 = fn1480(l0744)
l074C = extfnFE0(extfn1088(l06D6), 16, 0, 0)
l0750 = extfn10F8(l06BE, "Keep As New")
l0754 = fn1480(l0750)
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
Else
DoEvents
GoTo 136
End If
End If

136 :
l064C = Mid(l064C, l0652 + 1, Len(l064C))
GoTo 134
End If
server.Text2.Text = l064C
l065A = InStr(1, server.Text2.Text, "/" + l0626 + " Send List", 1)
If l065A Then
l065E = Mid(server.Text2.Text, l065A - 2, 2)
If l065E <> ":" + Chr(9) Then
GoTo 135
End If
l065E = InStr(server.Text2, ":")
l0662 = Mid(server.Text2.Text, 1, l065E - 1)
For l0666 = 0 To server.List2.ListCount - 1
l066A = InStr(1, l0662, server.List2.List(l0666), 1)
If l066A Then
If Len(l0662) = Len(server.List2.List(l0666)) Then GoTo 135
End If
Next l0666
If l0662 = l0626 Then GoTo 135
1335 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l0636 = extfn1018("AOL FRAME25", 0&)
l063A = fn5D0("&Mail", "&Compose Mail")
12 :
DoEvents
l066E = extfn10F8(l0636, "Compose Mail")
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
If l066E = 0 Then GoTo 12
l0640 = extfn10C0(l066E, "_AOL_EDIT")
l0644 = extfn11A0(l0640, 12, 0, l0662)
l0648 = extfn10F8(l066E, "Subject:")
l0672 = extfn1168(l0648, 2)
l0676 = extfn11A0(l0672, 12, 0, fn330() + "Fate Ultra Warez Server List" + fn368())
l067A = extfn10F8(l066E, "To:")
l065A = extfn1168(l067A, 1)
l065E = extfn11A0(l065A, 12, 0, server.Text5.Text)
l0662 = extfn10F8(l066E, "Send")
l067E = extfn1168(l0662, 2)
l0682 = fn1480(l067E)
501 :
DoEvents
l0686 = extfn1088(l0662)
l068A = extfn10F8(l0636, l0686)
l0636 = extfn1018("AOL FRAME25", 0&)
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l068E = extfn10F8(l0636, "Error")
If l068E > 0 Then
l061C = extfn1018("AOL FRAME25", 0&)
l0692 = extfn10F8(l061C, "Error")
l0696 = extfn10F8(l0692, "OK")
l069A = fn1480(l0696)
l069E = extfn10F8(l061C, "Compose Mail")
l06A2 = extfnFE0(l069E, 16, 0, 0)
GoTo LC538
End If
If l068A > 0 Then GoTo 501
l06A6 = extfn10F8(l068A, "OK")
l06AA = fn1480(l06A6)
LC538:
End If
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l065A = InStr(1, server.Text2.Text, "/" + l0626 + " Send ", 1)
If l065A Then
l065E = Mid(server.Text2.Text, l065A - 2, 2)
If l065E <> ":" + Chr(9) Then
GoTo 135
End If
l065E = InStr(server.Text2, ":")
l0662 = Mid(server.Text2.Text, 1, l065E - 1)
For l0666 = 0 To server.List2.ListCount - 1
l066A = InStr(1, l0662, server.List2.List(l0666), 1)
If l066A Then
If Len(l0662) = Len(server.List2.List(l0666)) Then GoTo 135
End If
Next l0666
l067E = Mid(server.Text2.Text, l065A + Len(l0626) + 7, Len(server.Text2.Text))
l06AE = InStr(1, l067E, "x", 1)
If l06AE Then GoTo 135
l06AE = InStr(1, l067E, "list", 1)
If l06AE Then GoTo 135
If IsNumeric(l067E) Then
For l0682 = 0 To server.List1.ListCount - 1
l068A = InStr(1, server.List1.List(l0682), l0662 + "-" + l067E, 1)
If l068A Then
If Len(server.List1.List(l0682)) = Len(l0662 + "-" + l067E) Then GoTo 1325
End If
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
Next l0682
If l0662 = l0626 Then GoTo 135
server.List1.AddItem l0662 + "-" + l067E
1325 :
l06B2 = Mid(server.Frame3D1.Caption, 9, Len(server.Frame3D1.Caption))
l06B6 = l06B2 + 1
server.Frame3D1.Caption = "Served: " + l06B6
l068E = extfn1018("AOL FRAME25", 0&)
If server.Option3D1.Value = True Then l06BA = "New Mail"
If server.Option3D2.Value = True Then l06BA = "Old Mail"
If server.Option3D3.Value = True Then l06BA = "Incoming FlashMail"
l06BE = extfn10F8(l068E, l06BA)
l06C2 = extfn10C0(l06BE, "_AOL_TREE")
l06C6 = extfnFE0(l06C2, 1036, 0, 0)
If server.Option3D3.Value = True Then
l06CA = extfnFE0(l06C2, 1031, l067E, 0)
l06CE = extfn10F8(l06BE, "Open")
l06D2 = fn1480(l06CE)
GoTo 341
End If
l06CA = extfnFE0(l06C2, 1031, l067E, 0)
l06CE = extfn10F8(l06BE, "Read")
l06D2 = fn1480(l06CE)
341 :
156 :
DoEvents
l06D6 = extfn10F8(l068E, "Forward")
If l06D6 = 0 Then GoTo 156
157 :
DoEvents
l06DA = extfn1168(l06D6, 2)
l06DE = fn1480(l06DA)
l06E2 = Timer
While Timer - l06E2 < .7
l06E6 = extfn10F8(l068E, "Address")
If l06E6 > 0 Then GoTo 163
DoEvents
Wend
l06EA = extfn10F8(l068E, "Send Now")
l06E6 = extfn10F8(l068E, "Address")
If l06E6 = 0 Then GoTo 157
163 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l06EA = extfn10F8(l068E, "Send Now")
l06EE = extfn1088(l06E6)
l06F2 = extfn10C0(l06EE, "_AOL_EDIT")
l06F6 = extfn11A0(l06F2, 12, 0, l0662)
l06FA = extfn10F8(l06EE, "Subject:")
l06FE = extfn1168(l06FA, 2)
l0702$ = String(255, 0)
l0704 = extfn11A0(l06FE, 13, 255, l0702$)
l0708 = Mid(l0702$, 6, Len(l0702$))
l070C = extfn11A0(l06FE, 12, 0, l0708)
l0710 = extfn10F8(l06EE, "To:")
l0714 = extfn1168(l0710, 1)
l0718 = extfn1168(l0714, 3)
l071C = extfn11A0(l0718, 12, 0, fn330() + "Fate Ultra WareZ Server By MaGuS" + fn368() + Chr(13) + Chr(10) + "«‹- -– Requested WareZ Index#: " + l067E + " –-- -›»" + l0614 + l0614 + server.Text3.Text)
l0720 = extfn1168(l06EA, 2)
l0724 = fn1480(l0720)
158 :
DoEvents
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
l0728 = extfn1088(l06EA)
l072C = extfn10F8(l068E, l0728)
l0730 = extfn10F8(l068E, "Error")
If l0730 > 0 Then
l0734 = extfn10F8(l0730, "OK")
l0738 = fn1480(l0734)
l073C = extfnFE0(l06EE, 16, 0, 0)
l0740 = extfnFE0(extfn1088(l06D6), 16, 0, 0)
GoTo 135
End If
If l072C > 0 Then GoTo 158
l0744 = extfn10F8(l072C, "OK")
l0748 = fn1480(l0744)
l074C = extfnFE0(extfn1088(l06D6), 16, 0, 0)
l0750 = extfn10F8(l06BE, "Keep As New")
l0754 = fn1480(l0750)
Else
GoTo 135
End If
End If

135 :
If server.Command3D3.Caption = "Start" Then GoTo LD7D0
server.Text1.Text = l063E$
DoEvents
If server.Caption <> fn14F0() Then Unload server
If server.tell.Text = "1" Then
If server.Caption <> fn14F0() Then Unload server
l068E = extfn1018("AOL FRAME25", 0&)
l0618 = extfn10F8(l068E, "Welcome")
l0758$ = String(30, 0)
l075A = extfnFA8(l0618, l0758$, 250)
l075E = Mid(l0758$, 10, 10)
l0762 = InStr(l075E, "!")
If l0762 Then
l075E = Mid(l075E, 1, l0762 - 1)
End If
l0632 = fnA68(fn330() + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(87) + Chr(97) + Chr(114) + Chr(101) + Chr(122) + Chr(32) + Chr(83) + Chr(101) + Chr(114) + Chr(118) + Chr(101) + Chr(114) + Chr(32) + Chr(66) + Chr(121) + Chr(32) + Chr(77) + Chr(97) + Chr(71) + Chr(117) + Chr(83) + fn368())
DoEvents
DoEvents
l0636 = fnA68(fn330() + Chr(84) + Chr(121) + Chr(112) + Chr(101) + Chr(32) + Chr(34) + "/" + l075E + " Send X" + Chr(34) + " X Is The Index #" + fn368())
DoEvents
DoEvents
l063A = fnA68(fn330() + Chr(84) + Chr(121) + Chr(112) + Chr(101) + Chr(32) + Chr(34) + "/" + l075E + " Send List" + Chr(34) + " To Get List" + fn368())
DoEvents
DoEvents
server.tell.Text = ""
End If
DoEvents
GoTo 132
LD7D0:
server.Text1.Text = ""
server.Text2.Text = ""
End Function

Function fnB48 () As Variant
Dim l076A As Variant
Dim l076E As Variant
Dim l0776 As Variant
Dim l077A As String
Dim l077C As Variant
Dim l0780 As Variant
Dim l0784 As Variant
Dim l0788 As Variant
Dim l078C As Variant
Dim l0790 As Variant
Dim l0794 As Variant
Dim l0798 As Variant
Dim l079C As Variant
Dim l07A0 As Variant
Dim l07A4 As Variant
Dim l07A8 As Variant
Dim l07AC As Variant
Dim l07B0 As Variant
Dim l07B4 As Variant
Dim l07B8 As Variant
Dim l07BC As Variant
Dim l07C0 As Variant
Dim l07C4 As Variant
Dim l07C8 As Variant
50 :
DoEvents
l076A = extfn1018("AOL FRAME25", 0&)
l076E = extfn10F8(l076A, "AOL Live!")
l076A = extfn1088(l076E)
l076E = extfn10C0(l076A, "_AOL_VIEW")
If l076E = 0 Then
Exit Function
End If
If smartbot.Timer1.Interval = 0 Then GoTo LE1DC
l0776 = extfnFE0(l076E, 14, 0, 0)
If l0776 = Len(smartbot.Text1.Text) Then GoTo 50
l077A$ = String(l0776, 0)
l077C = extfn11A0(l076E, 13, l0776, l077A$)
l0780 = Len(smartbot.Text1.Text)
l0784 = Len(l077A$) - l0780
On Error Resume Next
l0788 = Right(l077A$, l0784)
l0788 = Mid(l0788, 3, Len(l0788))
51 :
If smartbot.Timer1.Interval = 0 Then GoTo LE1DC
l078C = InStr(l0788, Chr(10))
If l078C Then
l0790 = Mid(l0788, 1, l078C - 2)
smartbot.Text2.Text = l0790
If smartbot.Text2.Text <> "" Then smartbot.Text3.Text = smartbot.Text2.Text
l0794 = InStr(smartbot.Text2.Text, Chr(0))
If l0794 Then
smartbot.Text2.Text = Mid(smartbot.Text2.Text, 1, l0794 - 1)
End If
l0798 = InStr(1, smartbot.Text3.Text, "has entered the room", 1)
If l0798 Then
l079C = InStr(smartbot.Text3, ":")
l07A0 = Mid(smartbot.Text3.Text, 1, l079C - 1)
If l07A0 <> "OnlineHost" Then GoTo 522
l07A4 = Mid(smartbot.Text3.Text, l079C + 2, Len(smartbot.Text3.Text))
l07A8 = InStr(1, l07A4, " has entered the room", 1)
l07AC = Mid(l07A4, 1, l07A8 - 1)
If smartbot.Check3D1.Value = True Then
l07B0 = Mid(l07AC, 1, 8)
l07B4 = InStr(1, l07B0, "Catwatch", 1)
If l07B4 Then
If Len(l07B0) = Len("Catwatch") Then
l07B8 = extfn10F8(l076A, "AOL live!")
l07BC = extfn1088(l07B8)
l07C0 = extfnFE0(l07BC, 16, 0, 0)
End If
End If
End If
l07C4 = fn13A0(.4)
l07C8 = fnA30(smartbot.Text4.Text + l07AC + smartbot.Text5.Text)
End If

522 :
l0788 = Mid(l0788, l078C + 1, Len(l0788))
If smartbot.Timer1.Interval = 0 Then GoTo LE1DC
GoTo 51
End If
smartbot.Text2.Text = l0788
If smartbot.Text2.Text <> "" Then smartbot.Text3.Text = smartbot.Text2.Text
l0794 = InStr(smartbot.Text2.Text, Chr(0))
If l0794 Then
smartbot.Text2.Text = Mid(smartbot.Text2.Text, 1, l0794 - 1)
End If
l0798 = InStr(1, smartbot.Text3.Text, "has entered the room", 1)
If l0798 Then
l079C = InStr(smartbot.Text3, ":")
l07A0 = Mid(smartbot.Text3.Text, 1, l079C - 1)
If l07A0 <> "OnlineHost" Then GoTo 53
l07A4 = Mid(smartbot.Text3.Text, l079C + 2, Len(smartbot.Text3.Text))
l07A8 = InStr(1, l07A4, " has entered the room", 1)
l07AC = Mid(l07A4, 1, l07A8 - 1)
If smartbot.Check3D1.Value = True Then
l07B0 = Mid(l07AC, 1, 8)
l07B4 = InStr(1, l07B0, "Catwatch", 1)
If l07B4 Then
If Len(l07B0) = Len("Catwatch") Then
l07B8 = extfn10F8(l076A, "AOL live!")
l07BC = extfn1088(l07B8)
l07C0 = extfnFE0(l07BC, 16, 0, 0)
End If
End If
End If
l07C4 = fn13A0(.4)
l07C8 = fnA30(smartbot.Text4.Text + l07AC + smartbot.Text5.Text)
End If

53 :
smartbot.Text1.Text = l077A$
smartbot.Text2.Text = ""
smartbot.Text3.Text = ""
If smartbot.Timer1.Interval = 0 Then GoTo LE1DC
DoEvents
GoTo 50
LE1DC:
smartbot.Text1.Text = ""
smartbot.Text2.Text = ""
smartbot.Text3.Text = ""
End Function

Function fnB80 () As Variant
End Function

Function fnBB8 () As Variant
fnBB8 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(70) + Chr(97) + Chr(116) + Chr(101) + Chr(32) + Chr(85) + Chr(108) + Chr(116) + Chr(114) + Chr(97) + Chr(32) + Chr(84) + Chr(79) + Chr(83) + Chr(101) + Chr(114) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fnBF0 (p07D8 As Variant) As Variant
Dim l07DC As Variant
Dim l07E0 As Variant
Dim l07E4 As Variant
Do Until 2 > 3
l07DC = InStr(p07D8, Chr(32))
If l07DC Then
l07E0 = Mid(p07D8, 1, l07DC - 1)
l07E4 = Mid(p07D8, l07DC + 1, Len(p07D8))
p07D8 = l07E0 + l07E4
Else
GoTo 70
End If
Loop
70 :
p07D8 = UCase(p07D8)
fnBF0 = p07D8
End Function

Function fnC28 () As Variant
fnC28 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(65) + Chr(99) + Chr(99) + Chr(111) + Chr(117) + Chr(110) + Chr(116) + Chr(32) + Chr(67) + Chr(104) + Chr(101) + Chr(99) + Chr(107) + Chr(101) + Chr(114) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fnC60 () As Variant
fnC60 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(77) + Chr(97) + Chr(71) + Chr(117) + Chr(83) + Chr(32) + Chr(86) + Chr(101) + Chr(114) + Chr(105) + Chr(102) + Chr(121) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fnC8 () As Variant
fnC8 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(65) + Chr(70) + Chr(75) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Function fnC98 () As Variant
fnC98 = Chr(45) + Chr(32) + Chr(45) + Chr(45) + Chr(87) + Chr(97) + Chr(114) + Chr(101) + Chr(122) + Chr(32) + Chr(82) + Chr(101) + Chr(113) + Chr(117) + Chr(101) + Chr(115) + Chr(116) + Chr(32) + Chr(66) + Chr(111) + Chr(116) + Chr(45) + Chr(45) + Chr(32) + Chr(45)
End Function

Sub sub250 ()
Dim l017E As Variant
Dim l0182 As Variant
Dim l0186 As Variant
Dim l018A As Variant
l017E = credits.Picture1.Width / 2
l0182 = credits.Picture1.Height / 2
l0186 = 3500
l018A = l0182
Do Until 2 > 3
DoEvents
'credits.Picture1.Line (x, y)-(l017E, l0182)
If l0186 <= l017E Then Exit Do
l0186 = l0186 - 4
l018A = l018A - 4
Loop
Do Until 2 > 3
DoEvents
'credits.Picture1.Line (x, y)-(l017E, l0182)
If l018A >= l0182 Then Exit Do
l0186 = l0186 - 4
l018A = l018A + 4
Loop
Do Until 2 > 3
DoEvents
'credits.Picture1.Line (x, y)-(l017E, l0182)
If l0186 >= l017E Then Exit Do
l0186 = l0186 + 4
l018A = l018A + 4
Loop
Do Until 2 > 3
DoEvents
'credits.Picture1.Line (x, y)-(l017E, l0182)
If l018A <= l0182 Then Exit Do
l0186 = l0186 + 4
l018A = l018A - 4
Loop
End Sub

Sub sub3A0 ()
Dim l0258 As Variant
Dim l025C As String
SetAttr "fate.ini", 0
l0258 = Chr(13) + Chr(10)
l025C$ = "[Main]" + l0258 + "WAOLDir=" + l0258 + "Intro=Yes" + l0258 + l0258 + "[Options]" + l0258 + "Times=1" + l0258 + "Load=No" + l0258 + l0258 + "[Private Room]" + l0258 + "Coldice" + l0258 + "Coldice2" + l0258 + "Coldice3" + l0258 + "Coldice4" + l0258 + l0258 + "[Public Room]" + l0258 + "Lobby" + l0258 + "Lobby2" + l0258 + l0258 + "[Member Room]" + l0258 + "Alaska" + l0258 + "Oklahoma" + l0258 + l0258 + "[FAC]" + l0258 + "CC Number=" + l0258 + "First Name=" + l0258 + "Last Name=" + l0258 + "Address=" + l0258 + "City=" + l0258 + "State=" + l0258 + "Zipcode=" + l0258 + "Phone Number=" + l0258 + "Certificate Number=" + l0258 + "Certificate Password=" + l0258 + "Expiration=" + l0258 + "Bank Name=" + l0258 + "Prefix="
Open "fate.ini" For Output As #1
Print #1, l025C$
Close #1
End Sub

Sub sub598 ()
End Sub

Sub sub838 ()
End Sub

Sub sub870 ()
End Sub

