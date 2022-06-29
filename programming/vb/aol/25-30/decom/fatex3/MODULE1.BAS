' Module1
Option Explicit

Declare Sub bringwindowtotop Lib "User" Alias "bringwindowtotop" (ByVal p1%)

Declare Function enablewindow Lib "User" Alias "enablewindow" (ByVal p1%, ByVal p2%) As Integer

'Declare Function FindChildByClass Lib "fate02.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
'Declare Function FindChildByTitle Lib "fate02.dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function Findchildbyclass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function Findchildbytitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function findwindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer

Declare Function getapitext Lib "fate00.dll" Alias "getapitext" (ByVal p1%) As String
Declare Function getmenu Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function getmenuitemcount Lib "User" Alias "getmenuitemcount" (ByVal p1%) As Integer
Declare Function getmenuitemid Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer
Declare Function getmenustring Lib "User" Alias "getmenustring" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function GETnextwindow Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function GetParent Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function getsubmenu Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function gettopwindow Lib "User" Alias "gettopwindow" (ByVal p1%) As Integer
Declare Function getwindowsdirectory Lib "Kernel" Alias "getwindowsdirectory" (ByVal p1$, ByVal p2%) As Integer
Declare Function GetWindowText Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer

Declare Function iswindowenabled Lib "User" Alias "iswindowenabled" (ByVal p1%) As Integer
Declare Function iswindowvisible Lib "User" Alias "iswindowvisible" (ByVal p1%) As Integer

Declare Sub movewindow Lib "User" Alias "movewindow" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%)

Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)

Declare Function showwindow Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function sndplaysound Lib "mmsystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Sub setwindowpos Lib "User" Alias "setwindowpos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Sub SetWindowText Lib "User" Alias "SetWindowText" ()
Declare Function setparent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer

Declare Function trackpopupmenu Lib "User" Alias "trackpopupmenu" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, p7 As Any) As Integer

Declare Function vbgetaollist Lib "fate00.dll" Alias "vbgetaollist" (ByVal p1%, ByVal p2%) As String

Function AOLVersion () As Integer
Dim l0102 As Variant
Dim l0106 As Variant
Dim l010A As Variant
Dim l010E As Variant
l0102 = findwindow("AOL FRAME25", 0&)
l0106 = Findchildbyclass(l0102, "AOL TOOLBAR")
l010A = Findchildbyclass(l0106, "_AOL_ICON")
l010E = GETnextwindow(l010A, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
l010E = GETnextwindow(l010E, 2)
If l010E > 0 Then
AOLVersion = 5
End If
If l010E = 0 Then
AOLVersion = 8
End If
End Function

Function fn100 () As Variant
Dim l00EE As String
Dim l00F0 As Variant
Dim l00F4 As Variant
Dim l00F8 As Variant
Dim l00FC As Variant
Open fnD78() + "\Fate.ini" For Binary As #1
l00EE$ = String(LOF(1), 0)
Get #1, 1, l00EE$
Close #1
l00F0 = InStr(1, l00EE$, "AOL Directory=", 1)
If l00F0 Then
l00F4 = Mid(l00EE$, l00F0 + 14, Len(l00EE$))
l00F8 = InStr(l00F4, Chr(13))
l00FC = Mid(l00F4, 1, l00F8 - 1)
fn100 = l00FC
Else
l00F4 = fn3A0()
End If
End Function

Function fn170 () As Variant
Dim l0116 As Variant
Dim l011A As Variant
Dim l011E As Variant
Dim l0122 As Variant
l0116 = findwindow("AOL FRAME25", 0&)
l011A = fn2F8()
l011E = Findchildbyclass(l011A, "_AOL_VIEW")
l0122 = getapitext(l011E)
fn170 = l0122
End Function

Function fn1A8 () As Variant
Dim l012A As Variant
Dim l012E As Variant
Dim l0132 As Variant
Dim l0136 As Variant
Dim l013A As Variant
Dim l013E As String
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
Dim l0170 As String
Dim l0172 As String
Dim l0174 As Variant
Dim l0178 As Variant
Dim l017C As Variant
Dim l0180 As Variant
Dim l0184 As Variant
Dim l0188 As Variant
Dim l018C As Variant
Dim l0192 As Variant
Dim l0196 As Variant
Dim l019A As String
Dim l019C As Variant
Dim l01A0 As Variant
Dim l01A4 As Variant
Dim l01A8 As Variant
Dim l01AC As Variant
Dim l01B0 As Variant
Dim l01B4 As Variant
Dim l01BA As Variant
Dim l01BE As Variant
Dim l01C2 As Variant
Dim l01C6 As String
Dim l01C8 As Variant
Dim l01CC As Variant
Dim l01D0 As Variant
Dim l01D4 As Variant
Dim l01D8 As Variant
Dim l01DC As Variant
Dim l01E0 As Variant
Dim l01E4 As Variant
l012A = 0
l012E = findwindow("AOL FRAME25", 0&)
l0132 = fn2F8()
l0136 = l0132
l013A = Findchildbyclass(l0136, "_AOL_VIEW")
l013E$ = getapitext(l013A)
autoignore.Text1.Text = l013E$
For l0144 = 0 To Len(autoignore.Text1.Text) - 1
l0148 = Mid(autoignore.Text1.Text, Len(autoignore.Text1.Text) - l0144, 1)
If l0148 = Chr(10) Then Exit For
l014C = l0148 + l014C
DoEvents
Next l0144
l0150 = autoignore.Text1.Text
Do Until 2 > 3
l0154 = InStr(l0150, l014C)
If l0154 Then
autoignore.Text2.Text = Trim(Str(Int(autoignore.Text2.Text + 1)))
l0150 = Mid(l0150, l0154 + Len(l014C) + 2, Len(l0150))
Else
Exit Do
End If
DoEvents
Loop
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
l0158 = findwindow("AOL FRAME25", 0&)
l015C = fn2F8()
l0160 = l015C
l0164 = Findchildbyclass(l0160, "_AOL_VIEW")
l0168 = SendMessageByNum(l0164, 14, 0, 0)
1019 :
If l012A = 1 Then
l0168 = l016C
End If
Do Until 2 > 3
l0160 = fn2F8()
l0164 = Findchildbyclass(l0160, "_AOL_VIEW")
l016C = SendMessageByNum(l0164, 14, 0, 0)
If l0168 <> l016C Then Exit Do
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
If fn2F8() = 0 Then
MsgBox "Chat room has been closed, shutting down bot!", 16
autoignore.Command3D2.Caption = "Start"
autoignore.Timer1.Interval = 0
GoTo L34DC
End If
DoEvents
Loop
l0170$ = getapitext(l0164)
autoignore.Text1.Text = l0170$
l0172$ = autoignore.Text1.Text
For l0174 = 1 To Int(autoignore.Text2.Text)
DoEvents
l0178 = InStr(l0172$, l014C)
If l0178 Then
l0172$ = Mid(l0172$, l0178 + Len(l014C) + 2, Len(l0172$))
Else
GoTo 1559
End If
Next l0174
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
1559 :
l017C = l0172$

3009 :
DoEvents
l0180 = InStr(l017C, Chr(13) + Chr(10))
If l0180 Then
l0184 = Mid(l017C, 1, l0180 - 1)
autoignore.Text3.Text = l0184
l0188 = InStr(l0184, ":" + Chr(9))
If l0188 Then
l018C = Mid(l0184, 1, l0188 - 1)
If l018C = fnA68() Then GoTo 115
If l018C = "OnlineHost" Then GoTo 115
If LCase(l018C) = LCase(autoignore.Text5.Text) Then
autoignore.Text6.Text = Trim(Str(Int(autoignore.Text6.Text + 1)))
If Int(autoignore.Text6.Text) = Int(autoignore.Text4.Text) Then
autoignore.List1.AddItem l018C
l0192 = Timeout(.6)
l0192 = fnAD8(fn448() + "Fate X³ - " + l018C + " - Ignored" + fn480())
autoignore.Label2.Caption = Trim(Str(autoignore.List1.ListCount))
114 :
DoEvents
For l0196 = 0 To SendMessageByNum(Findchildbyclass(fn2F8(), "_AOL_LISTBOX"), 1036, 0, 0)
 bringwindowtotop fn2F8()
l019A$ = vbgetaollist(l0196, Findchildbyclass(fn2F8(), "_AOLLISTBOX"))
If fnC28(LCase(l019A$)) = fnC28(LCase(l018C)) Then
l019C = SendMessageByNum(Findchildbyclass(fn2F8(), "_AOL_LISTBOX"), 1031, l0196, 0)
l01A0 = findwindow("AOL FRAME25", 0&)
l01A4 = Findchildbyclass(l01A0, "_AOL_LISTBOX")
l01A8 = SendMessageByNum(l01A4, 515, 0, 0)
Do Until 2 > 3
l01AC = Findchildbytitle(l01A0, "Get Info")
If l01AC <> 0 Then Exit Do
DoEvents
Loop
l01B0 = GetParent(l01AC)
l01B4 = Findchildbytitle(l01B0, "Message")
l01BA = fn1870(l01B4)
Do Until 2 > 3
DoEvents
l01BE = Findchildbytitle(l01A0, "Send Instant Message")
l01C2 = Findchildbyclass(l01BE, "_AOL_EDIT")
l01C6$ = String(255, 0)
l01C8 = SendMessageByString(l01C2, 13, 255, l01C6$)
If fnC28(l01C6$) = fnC28(l018C) Then Exit Do
Loop
l01CC = SendMessageByNum(l01BE, 16, 0, 0)
l01D0 = Findchildbytitle(l01B0, "Get Info")
l01D4 = GETnextwindow(l01D0, 1)
l01D8 = GETnextwindow(l01D4, 3)
l01DC = GETnextwindow(l01D8, 3)
For l01E0 = 1025 To 1029
l0154 = SendMessageByNum(l01DC, l01E0, 1, 0)
Next l01E0
Do Until 2 > 3
DoEvents
l01B4 = Findchildbytitle(l01B0, "Message")
l01BA = fn1870(l01B4)
l01BE = Findchildbytitle(l01A0, "Send Instant Message")
l01C2 = Findchildbyclass(l01BE, "_AOL_EDIT")
l01C6$ = String(255, 0)
l01C8 = SendMessageByString(l01C2, 13, 255, l01C6$)
If fnC28(l01C6$) = fnC28(l018C) Then Exit Do
Loop
l01E0 = SendMessageByNum(l01B0, 16, 0, 0)
l01E0 = SendMessageByNum(l01BE, 16, 0, 0)
GoTo 115
End If
Next l0196
GoTo 114
End If
Else
autoignore.Text5.Text = l018C
autoignore.Text6.Text = "1"
End If
End If
115 :
l017C = Mid(l017C, l0180 + 2, Len(l017C))
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
GoTo 3009
Else
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
autoignore.Text3.Text = l017C
l0188 = InStr(l017C, ":" + Chr(9))
If l0188 Then
l018C = Mid(l017C, 1, l0188 - 1)
If l018C = fnA68() Then GoTo 117
If l018C = "OnlineHost" Then GoTo 117
If LCase(l018C) = LCase(autoignore.Text5.Text) Then
autoignore.Text6.Text = Trim(Str(Int(autoignore.Text6.Text + 1)))
If Int(autoignore.Text6.Text) = Int(autoignore.Text4.Text) Then
autoignore.List1.AddItem l018C
l0192 = Timeout(.6)
l0192 = fnAD8(fn448() + "Fate X³ - " + l018C + " - Ignored" + fn480())
autoignore.Label2.Caption = Trim(Str(autoignore.List1.ListCount))
116 :
DoEvents
For l0196 = 0 To SendMessageByNum(Findchildbyclass(fn2F8(), "_AOL_LISTBOX"), 1036, 0, 0)
 bringwindowtotop fn2F8()
l019A$ = vbgetaollist(l0196, Findchildbyclass(fn2F8(), "_AOL_LISTBOX"))
If fnC28(LCase(l019A$)) = fnC28(LCase(l018C)) Then
l019C = SendMessageByNum(Findchildbyclass(fn2F8(), "_AOL_LISTBOX"), 1031, l0196, 0)
l01A0 = findwindow("AOL FRAME25", 0&)
l01A4 = Findchildbyclass(l01A0, "_AOL_LISTBOX")
l01A8 = SendMessageByNum(l01A4, 515, 0, 0)
Do Until 2 > 3
l01AC = Findchildbytitle(l01A0, "Get Info")
If l01AC <> 0 Then Exit Do
DoEvents
Loop
l01B0 = GetParent(l01AC)
l01B4 = Findchildbytitle(l01B0, "Message")
l01BA = fn1870(l01B4)
Do Until 2 > 3
DoEvents
l01BE = Findchildbytitle(l01A0, "Send Instant Message")
l01C2 = Findchildbyclass(l01BE, "_AOL_EDIT")
l01C6$ = String(255, 0)
l01C8 = SendMessageByString(l01C2, 13, 255, l01C6$)
If fnC28(l01C6$) = fnC28(l018C) Then Exit Do
Loop
l01CC = SendMessageByNum(l01BE, 16, 0, 0)
l01D0 = Findchildbytitle(l01B0, "Get Info")
l01D4 = GETnextwindow(l01D0, 1)
l01D8 = GETnextwindow(l01D4, 3)
l01DC = GETnextwindow(l01D8, 3)
For l01E0 = 1025 To 1029
l0154 = SendMessageByNum(l01DC, l01E0, 1, 0)
Next l01E0
Do Until 2 > 3
DoEvents
l01B4 = Findchildbytitle(l01B0, "Message")
l01BA = fn1870(l01B4)
l01BE = Findchildbytitle(l01A0, "Send Instant Message")
l01C2 = Findchildbyclass(l01BE, "_AOL_EDIT")
l01C6$ = String(255, 0)
l01C8 = SendMessageByString(l01C2, 13, 255, l01C6$)
If fnC28(l01C6$) = fnC28(l018C) Then Exit Do
Loop
l01E0 = SendMessageByNum(l01B0, 16, 0, 0)
l01E0 = SendMessageByNum(l01BE, 16, 0, 0)
GoTo 117
End If
Next l0196
GoTo 116
End If
Else
autoignore.Text5.Text = l018C
autoignore.Text6.Text = "1"
End If
End If
End If

117 :
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
l014C = autoignore.Text3.Text
l0192 = autoignore.Text1.Text
autoignore.Text2.Text = "0"
Do Until 2 > 3
l01E4 = InStr(l0192, l014C)
If l01E4 Then
autoignore.Text2.Text = Trim(Str(Int(autoignore.Text2.Text + 1)))
l0192 = Mid(l0192, l01E4 + Len(l014C) + 2, Len(l0192))
Else
Exit Do
End If
Loop
l012A = 1
DoEvents
If autoignore.Command3D2.Caption = "Start" Then GoTo L34DC
GoTo 1019

L34DC:
autoignore.Text1.Text = ""
autoignore.Text2.Text = "0"
autoignore.Text3.Text = ""
autoignore.Text6.Text = "0"
End Function

Function fn1E0 () As Variant
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
Dim l0218 As Variant
Dim l021C As Variant
Dim l0220 As Variant
Dim l0224 As Variant
Dim l0228 As Variant
Dim l022C As String
Dim l022E As String
Dim l0230 As Variant
Dim l0234 As Variant
Dim l0238 As Variant
Dim l023C As Variant
Dim l0240 As Variant
Dim l0244 As Variant
Dim l0248 As Variant
Dim l024C As Variant
Dim l0250 As Variant
Dim l0254 As Variant
Dim l0258 As Variant
Dim l025C As Variant
Dim l0260 As Variant
Dim l0264 As Variant
Dim l0268 As Variant
l01EC = 0
l01F0 = findwindow("AOL FRAME25", 0&)
l01F4 = fn2F8()
l01F8 = Findchildbyclass(l01F4, "_AOL_VIEW")
l01FC = getapitext(l01F8)
mmER.Text8.Text = l01FC
For l0200 = 0 To Len(mmER.Text8.Text) - 1
l0204 = Mid(mmER.Text8.Text, Len(mmER.Text8.Text) - l0200, 1)
If l0204 = Chr(10) Then Exit For
l0208 = l0204 + l0208
DoEvents
Next l0200
l020C = mmER.Text8.Text
Do Until 2 > 3
l0210 = InStr(l020C, l0208)
If l0210 Then
mmER.Text10.Text = Trim(Str(Int(mmER.Text10.Text + 1)))
l020C = Mid(l020C, l0210 + Len(l0208) + 2, Len(l020C))
Else
Exit Do
End If
DoEvents
Loop
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E
l0218 = findwindow("AOL FRAME25", 0&)
l021C = fn2F8()
l0220 = Findchildbyclass(l021C, "_AOL_VIEW")
l0224 = SendMessageByNum(l0220, 14, 0, 0)
10111 :
If l01EC = 1 Then
l0224 = l0228
End If
Do Until 2 > 3
l021C = fn2F8()
l0220 = Findchildbyclass(l021C, "_AOL_VIEW")
l0228 = SendMessageByNum(l0220, 14, 0, 0)
If l0224 <> l0228 Then Exit Do
If fn2F8() = 0 Then
mmbot.Command3D1.Caption = "Start"
mmER.Timer5.Interval = 0
mmER.Timer5.Enabled = False
mmER.TIMER4.Interval = 0
mmER.TIMER4.Enabled = False
mmER.Timer1.Interval = 0
mmER.Timer1.Enabled = False
mmER.Timer3.Interval = 0
mmER.Timer3.Enabled = False
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo L454E
End If
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E
DoEvents
Loop
l022C$ = getapitext(l0220)
mmER.Text8.Text = l022C$
l022E$ = mmER.Text8.Text
For l0230 = 1 To Int(mmER.Text10.Text)
DoEvents
l0234 = InStr(l022E$, l0208)
If l0234 Then
l022E$ = Mid(l022E$, l0234 + Len(l0208) + 2, Len(l022E$))
Else
GoTo 1555
End If
Next l0230
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E

1555 :
l0238 = l022E$

3000 :
DoEvents
l023C = InStr(l0238, Chr(13) + Chr(10))
If l023C Then
l0240 = Mid(l0238, 1, l023C - 1)
mmER.Text9.Text = l0240
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E
l0244 = InStr(1, l0240, ":" + Chr(9) + mmbot.Text4.Text, 1)
If l0244 Then
l0248 = Mid(l0240, l0244, 2)
l024C = InStr(l0240, ":")
l0250 = Mid(l0240, 1, l024C - 1)
If l0250 = fnA68() Then GoTo 17
l0254 = Mid(l0240, l024C + 2, Len(l0240))
If Len(l0254) <> Len(mmbot.Text4.Text) Then GoTo 17
For l0258 = 0 To mmER.List1.ListCount - 1
l0220 = InStr(1, mmER.List1.List(l0258), l0250, 1)
If l0220 Then
If Len(mmER.List1.List(l0258)) = Len(l0250) Then GoTo 17
End If
Next l0258
mmER.List1.AddItem l0250
mmbot.List1.AddItem l0250
mmER.Text1.Text = mmER.Text1.Text + l0250 + ","
mmER.frame3d1.Caption = "People On: " + Trim(Str(mmER.List1.ListCount))
mmbot.frame3d1.Caption = "People On: " + Trim(Str(mmER.List1.ListCount))
If mmbot.Check3D1.Value = True Then
l025C = Timeout(.6)
l0260 = fnAA0(fn448() + Chr(60) + l0250 + Chr(62) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(111) + Chr(110) + Chr(32) + Chr(116) + Chr(104) + Chr(101) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + fn480())
End If
End If
17 :
l0238 = Mid(l0238, l023C + 2, Len(l0238))
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E
GoTo 3000
Else
mmER.Text9.Text = l0238
l0244 = InStr(1, l0238, ":" + Chr(9) + mmbot.Text4.Text, 1)
If l0244 Then
l0248 = Mid(l0238, l0244, 2)
l024C = InStr(l0238, ":")
l0250 = Mid(l0238, 1, l024C - 1)
If l0250 = fnA68() Then GoTo 177
l0254 = Mid(l0238, l024C + 2, Len(l0238))
If Len(l0254) <> Len(mmbot.Text4.Text) Then GoTo 177
For l0258 = 0 To mmER.List1.ListCount - 1
l0220 = InStr(1, mmER.List1.List(l0258), l0250, 1)
If l0220 Then
If Len(mmER.List1.List(l0258)) = Len(l0250) Then GoTo 177
End If
Next l0258
mmER.List1.AddItem l0250
mmbot.List1.AddItem l0250
mmER.Text1.Text = mmER.Text1.Text + l0250 + ","
mmER.frame3d1.Caption = "People On: " + Trim(Str(mmER.List1.ListCount))
mmbot.frame3d1.Caption = "People On: " + Trim(Str(mmER.List1.ListCount))
If mmbot.Check3D1.Value = True Then
l025C = Timeout(.6)
l0260 = fnAA0(fn448() + Chr(60) + l0250 + Chr(62) + Chr(32) + Chr(105) + Chr(115) + Chr(32) + Chr(111) + Chr(110) + Chr(32) + Chr(116) + Chr(104) + Chr(101) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(115) + Chr(115) + Chr(32) + Chr(40) + Chr(86) + Chr(41) + Chr(97) + Chr(105) + Chr(108) + fn480())
End If
End If
177 :
If mmbot.Command3D1.Caption = "Start" Then GoTo L454E
End If
l0208 = mmER.Text9.Text
l0264 = mmER.Text8.Text
mmER.Text10.Text = "0"
Do Until 2 > 3
l0268 = InStr(l0264, l0208)
If l0268 Then
mmER.Text10.Text = Trim(Str(Int(mmER.Text10.Text + 1)))
l0264 = Mid(l0264, l0268 + Len(l0208) + 2, Len(l0264))
Else
Exit Do
End If
Loop
l01EC = 1
DoEvents
GoTo 10111

L454E:
mmER.Text8.Text = ""
mmER.Text9.Text = ""
mmER.Text10.Text = "0"
End Function

Function fn218 () As Variant
Dim l0270 As Variant
Dim l0274 As Variant
Dim l0278 As Variant
Dim l027C As Variant
Dim l0280 As String
Dim l0282 As String
Dim l0284 As Variant
Dim l0288 As Variant
Dim l028C As Variant
If l0270 = Chr(10) Then GoTo 1666
l0274 = l0270 + l0274

1666 :
Do Until 2 > 3
l027C = InStr(l0278, l0274)
If l027C Then
l0278 = Mid(l0278, l027C + Len(l0274) + 2, Len(l0278))
Else
Exit Do
End If
DoEvents
Loop

11111 :
DoEvents
l0280$ = fn170()
l0282$ = l0280$
DoEvents
If l0284 Then
If l0280$ = "" Then GoTo 11111
Else
GoTo 15555
End If

15555 :
DoEvents
Do Until 2 > 3
DoEvents
l0288 = InStr(l0280$, Chr(13) + Chr(10))
If l0288 Then
l028C = Mid(l0280$, 1, l0288 - 1)
l028C = fnC28(l028C)
l0280$ = Mid(l0280$, Len(l028C) + 2, Len(l0280$))
GoTo 15555
Else
l0280$ = fnC28(l0280$)
If l0280$ = "" Then GoTo 11111
MsgBox l0280$
l0274 = l0280$
l0278 = l0282$
Do Until 2 > 3
l027C = InStr(l0278, l0274)
If l027C Then
l0278 = Mid(l0278, l027C + Len(l0274) + 2, Len(l0278))
Else
Exit Do
End If
DoEvents
Loop
GoTo 11111
End If
Loop
End Function

Function fn250 (p0294 As Variant, p0298 As String) As Variant
Dim l029E As Variant
Dim l02A2 As Variant
Dim l02A6 As Variant
Dim l02AA As Variant
sugg.Show
l029E = Chr(13) + Chr(10)
Do Until l02A2 > 0
l02A2 = findwindow(0&, "- --Suggestion/Bugs/Comment Box-- -")
DoEvents
Loop
sugg.sugger.Text = "Bug Report -" + l029E + l029E + "Form Name - " + p0298 + l029E + "Error - #" + Trim(Str(p0294))
l02A6 = Findchildbytitle(l02A2, "bug")
l02AA = fn1870(l02A6)
MsgBox "A bug has been found and the programmer will be notified!", 32
End Function

Function fn2F8 () As Variant
Dim l02BE As Variant
Dim l02C2 As Variant
Dim l02C6 As Variant
Dim l02CA As Variant
Dim l02CE As Variant
Dim l02D2 As Variant
Dim l02D6 As Variant
Dim l02DA As String
Dim l02DC As Variant
l02BE = findwindow("AOL FRAME25", 0&)
l02C2 = Findchildbyclass(l02BE, "_AOL_VIEW")
l02C6 = GetParent(l02C2)
Do Until 2 > 3
l02CA = Findchildbyclass(l02C6, "_AOL_VIEW")
l02CE = Findchildbyclass(l02C6, "_AOL_LISTBOX")
l02D2 = Findchildbyclass(l02C6, "_AOL_ICON")
l02D6 = Findchildbyclass(l02C6, "_AOL_STATIC")
If l02CA > 0 Then
If l02CE > 0 Then
If l02D2 > 0 Then
If l02D6 > 0 Then
l02DA$ = String(255, 0)
l02DC = SendMessageByString(l02D6, 13, 255, l02DA$)
fn2F8 = GetParent(l02D6)
Exit Function
End If
End If
End If
End If
l02C6 = GETnextwindow(l02C6, 2)
If l02C6 = 0 Then
fn2F8 = 0
Exit Function
End If
DoEvents
Loop
End Function

Function fn330 () As Variant
Dim l02E4 As Variant
Dim l02E8 As Variant
Dim l02EE As Variant
Dim l02F2 As Variant
l02E4 = 0
l02E8 = findwindow("AOL FRAME25", 0&)
l02E8 = Findchildbyclass(l02E8, "MDICLIENT")
l02EE = gettopwindow(l02E8)
If l02EE = 0 Then l02E4 = 0
3 :
DoEvents
l02F2 = GETnextwindow(l02EE, 2)
If l02F2 > 0 Then
l02E4 = l02E4 + 1
Else
fn330 = l02E4
Exit Function
End If
l02EE = l02F2
GoTo 3
End Function

Function fn368 (ByVal p02FA As String, ByVal p02FC As String) As Variant
Dim l02FE As Variant
Dim l0302 As Variant
Dim l0306 As Variant
Dim l030A As String
Dim l030C As Variant
Dim l0310 As Variant
Dim l0314 As String
Dim l0316 As Variant
l02FE = findwindow("AOL FRAME25", 0&)
l0302 = Findchildbyclass(l02FE, p02FA)
If l0302 = 0 Then
fn368 = 0
Exit Function
End If
l0306 = GetParent(l0302)
l030A$ = String(255, 0)
l030C = GetWindowText(l0306, l030A$, 255)
l030A$ = fnC28(l030A$)
If l030A$ = p02FC Then
fn368 = l0306
Exit Function
Else
L4F58: :
DoEvents
l0310 = GETnextwindow(l0306, 2)
l0314$ = String(255, 0)
l0316 = GetWindowText(l0310, l0314$, 255)
l0314$ = fnC28(l0314$)
If l0314$ = p02FC Then
l0316 = Findchildbyclass(l0310, p02FA)
If l0316 > 0 Then
fn368 = l0310
Exit Function
End If
End If
If l0314$ = "" Then
fn368 = 0
Exit Function
Else
l0306 = l0310
GoTo L4F58
End If
End If
End Function

Function fn3A0 () As Variant
Dim l0322 As Variant
Dim l0326 As Variant
Dim l032A As String
Dim l032C As String
Dim l032E As String
pwait.Show
MsgBox "The FATE.INI file in your windows directory is either missing or invalid. Click on ok to fix the problem.", 16
l0322 = Chr(13) + Chr(10)
l0326 = l0322 + l0322 + "[Advertise Options]" + l0322 + "Advertise=Yes" + l0322 + l0322 + "[Room Finder]" + l0322 + "Fate" + l0322 + "Freeshit" + l0322 + "Warez" + l0322 + "FreeWarez" + l0322 + "Dew" + l0322 + "Phreestuff" + l0322 + "Freestuff"
l032A$ = "[FAC]" + l0322 + "CC Number:" + l0322 + "First Name:" + l0322 + "Last Name:" + l0322 + "Address:" + l0322 + "City:" + l0322 + "State:" + l0322 + "Zipcode:" + l0322 + "Phone Number:" + l0322 + "Certificate Number:" + l0322 + "Certificate Password:" + l0322 + "Expiration:" + l0322 + "Bank Name:" + l0322 + "Prefix:" + l0322 + l0322 + "[Mass Online Checker]" + l0326
l032C$ = "[Private Room]" + l0322 + l0322 + "[Public Room]" + l0322 + l0322 + "[Member Room]" + l0322 + l0322 + l032A$
l032E$ = "[Main]" + l0322 + "AOL Directory=" + l0322 + "Intro=Yes" + l0322 + l0322 + "[Support]" + l0322 + "Times=0" + l0322 + "Load=No" + l0322 + l0322 + l032C$
Open fnD78() + "\fate.ini" For Output As #1
Print #1, l032E$
Close #1
Unload pwait
options.Show
End Function

Function fn3D8 () As Variant
Dim l0338 As Variant
Dim l033C As Variant
Dim l0340 As Variant
Dim l0344 As Variant
Dim l0348 As Variant
Dim l034C As String
Dim l034E As Variant
Dim l0352 As Variant
Dim l0356 As Variant
l0338 = findwindow("AOL FRAME25", 0&)
l033C = fn2F8()
l0340 = l033C
l0344 = Findchildbyclass(l0340, "_AOL_EDIT")
l0348 = GetParent(l0344)
l034C$ = String(255, 0)
l034E = GetWindowText(l0348, l034C$, 250)
l0352 = l034C$
l0356 = fnC28(l0352)
fn3D8 = l0356
End Function

Function fn410 () As Variant
Dim l035E As Variant
Dim l0362 As Variant
Dim l0366 As Variant
Dim l036A As Variant
Dim l036E As Variant
Dim l0372 As String
Dim l0378 As Variant
Dim l037C As Variant
Dim l0380 As Variant
Dim l0384 As Variant
Dim l0388 As Variant
Dim l038C As Variant
Dim l0390 As Variant
Dim l0394 As Variant
Dim l0398 As Variant
Dim l039C As Variant
Dim l03A0 As Variant
Dim l03A4 As String
Dim l03A6 As String
Dim l03A8 As Variant
Dim l03AC As Variant
Dim l03B0 As Variant
Dim l03B4 As Variant
Dim l03B8 As Variant
Dim l03BC As Variant
Dim l03C0 As Variant
Dim l03C4 As Variant
Dim l03C8 As Variant
Dim l03CC As Variant
Dim l03D0 As Variant
Dim l03D4 As Variant
Dim l03D8 As Variant
Dim l03DC As Variant
Dim l03E0 As Variant
Dim l03E4 As Variant
Dim l03E8 As Variant
Dim l03EC As Variant
Dim l03F0 As Variant
Dim l03F4 As Variant
Dim l03F8 As Variant
Dim l03FC As Variant
l035E = 0
l0362 = findwindow("AOL FRAME25", 0&)
l0366 = fn2F8()
l036A = l0366
l036E = Findchildbyclass(l036A, "_AOL_VIEW")
l0372$ = getapitext(l036E)
encryptchat.Text5.Text = l0372$
For l0378 = 0 To Len(encryptchat.Text5.Text) - 1
l037C = Mid(encryptchat.Text5.Text, Len(encryptchat.Text5.Text) - l0378, 1)
If l037C = Chr(10) Then Exit For
l0380 = l037C + l0380
DoEvents
Next l0378
l0384 = encryptchat.Text5.Text
Do Until 2 > 3
l0388 = InStr(l0384, l0380)
If l0388 Then
encryptchat.Text6.Text = Trim(Str(Int(encryptchat.Text6.Text + 1)))
l0384 = Mid(l0384, l0388 + Len(l0380) + 2, Len(l0384))
Else
Exit Do
End If
DoEvents
Loop
l038C = findwindow("AOL FRAME25", 0&)
l0390 = fn2F8()
l0394 = l0390
l0398 = Findchildbyclass(l0394, "_AOL_VIEW")
l039C = SendMessageByNum(l0398, 14, 0, 0)
1012 :
If l035E = 1 Then
l039C = l03A0
End If
Do Until 2 > 3
l0394 = fn2F8()
l0398 = Findchildbyclass(l0394, "_AOL_VIEW")
l03A0 = SendMessageByNum(l0398, 14, 0, 0)
If l039C <> l03A0 Then Exit Do
DoEvents
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
If fn2F8() = 0 Then
encryptchat.Command3D5.Caption = "Start"
encryptchat.Timer1.Interval = 0
encryptchat.Timer1.Enabled = False
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo L6818
End If
Loop
l03A4$ = getapitext(l0398)
encryptchat.Text5.Text = l03A4$
l03A6$ = encryptchat.Text5.Text
For l03A8 = 1 To Int(encryptchat.Text6.Text)
DoEvents
l03AC = InStr(l03A6$, l0380)
If l03AC Then
l03A6$ = Mid(l03A6$, l03AC + Len(l0380) + 2, Len(l03A6$))
Else
GoTo 1552
End If
Next l03A8

1552 :
l03B0 = l03A6$

3002 :
DoEvents
l03B4 = InStr(l03B0, Chr(13) + Chr(10))
If l03B4 Then
l03B8 = Mid(l03B0, 1, l03B4 - 1)
encryptchat.Text7.Text = l03B8
l0388 = InStr(l03B8, ":")
If l0388 = 0 Then GoTo 10000
l0384 = Mid(l03B8, 1, l0388 - 1)
If l0384 = fnA68() Then GoTo 10000
For l038C = 0 To encryptchat.List1.ListCount - 1
l0390 = InStr(1, encryptchat.List1.List(l038C), l0384, 1)
If l0390 Then
If Len(encryptchat.List1.List(l038C)) = Len(l0384) Then
GoTo 10000
End If
End If
Next l038C
GoTo 10001
10000 :
l0394 = Mid(l03B8, l0388 + 2, Len(l03B8))
l0394 = fnC28(l0394)
l03BC = Mid(l0394, 1, 2)
If l03BC <> Chr(255) + Chr(208) Then
GoTo 10001
Else
l0394 = Mid(l0394, 3, Len(l0394))
End If
687 :
DoEvents
l03C0 = InStr(l0394, Chr(9))
If l03C0 Then
l03C4 = Mid(l0394, 1, l03C0 - 1)
l03C8 = Mid(l0394, l03C0 + 1, Len(l0394))
l0394 = l03C4 + l03C8
GoTo 687
End If
If l0394 = "" Then GoTo 10001
If l0394 = fn448() + "Fate X³ Encryption Chat Enabled" + fn480() Then GoTo 10001
If l0394 = fn448() + "Fate X³ Encryption Chat Disabled" + fn480() Then GoTo 10001
l03CC = ""
l03D0 = 1
For l03D4 = 1 To Len(l0394)
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
l03D8 = Mid(l0394, l03D4, 1)
l03DC = Asc(l03D8)
l03E0 = Mid(encryptchat.Text3.Text, l03D0, 1)
l03D0 = l03D0 + 1
If l03D0 > Len(encryptchat.Text3.Text) Then l03D0 = 1
l03E4 = Asc(l03E0)
l03E8 = l03DC - l03E4
If l03E8 < 1 Then
l03EC = l03E8 + 255
l03F0 = Chr(l03EC)
Else
l03F0 = Chr(l03E8)
End If
l03CC = l03CC + l03F0
DoEvents
Next l03D4
l03F4 = Chr(13) + Chr(10)
encryptchat.Text1.Text = encryptchat.Text1.Text + l03F4 + l0384 + ":" + Chr(9) + l03CC
encryptchat.Text1.SelStart = Len(encryptchat.Text1.Text)
10001 :
l03B0 = Mid(l03B0, l03B4 + 2, Len(l03B0))
GoTo 3002
Else
encryptchat.Text7.Text = l03B0
l0388 = InStr(l03B0, ":")
If l0388 = 0 Then GoTo L62CC
l0384 = Mid(l03B0, 1, l0388 - 1)
If l0384 = fnA68() Then GoTo L62CC
For l038C = 0 To encryptchat.List1.ListCount - 1
l0390 = InStr(1, encryptchat.List1.List(l038C), l0384, 1)
If l0390 Then
If Len(encryptchat.List1.List(l038C)) = Len(l0384) Then
GoTo L62CC
End If
End If
Next l038C
GoTo L66CC
L62CC: :
l0394 = Mid(l03B0, l0388 + 2, Len(l03B0))
l0394 = fnC28(l0394)
l03BC = Mid(l0394, 1, 2)
If l03BC <> Chr(255) + Chr(208) Then
GoTo L66CC
Else
l0394 = Mid(l0394, 3, Len(l0394))
End If
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
6871 :
DoEvents
l03C0 = InStr(l0394, Chr(9))
If l03C0 Then
l03C4 = Mid(l0394, 1, l03C0 - 1)
l03C8 = Mid(l0394, l03C0 + 1, Len(l0394))
l0394 = l03C4 + l03C8
GoTo 6871
End If
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
If l0394 = "" Then GoTo L66CC
If l0394 = fn448() + "Fate X³ Encryption Chat Enabled" + fn480() Then GoTo L66CC
If l0394 = fn448() + "Fate X³ Encryption Chat Disabled" + fn480() Then GoTo L66CC
l03CC = ""
l03D0 = 1
For l03D4 = 1 To Len(l0394)
l03D8 = Mid(l0394, l03D4, 1)
l03DC = Asc(l03D8)
l03E0 = Mid(encryptchat.Text3.Text, l03D0, 1)
l03D0 = l03D0 + 1
If l03D0 > Len(encryptchat.Text3.Text) Then l03D0 = 1
l03E4 = Asc(l03E0)
l03E8 = l03DC - l03E4
If l03E8 < 1 Then
l03EC = l03E8 + 255
l03F0 = Chr(l03EC)
Else
l03F0 = Chr(l03E8)
End If
l03CC = l03CC + l03F0
DoEvents
Next l03D4
l03F4 = Chr(13) + Chr(10)
encryptchat.Text1.Text = encryptchat.Text1.Text + l03F4 + l0384 + ":" + Chr(9) + l03CC
encryptchat.Text1.SelStart = Len(encryptchat.Text1.Text)

L66CC: :
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
End If
l0380 = encryptchat.Text7.Text
l03F8 = encryptchat.Text5.Text
encryptchat.Text6.Text = "0"
Do Until 2 > 3
l03FC = InStr(l03F8, l0380)
If l03FC Then
encryptchat.Text6.Text = Trim(Str(Int(encryptchat.Text6.Text + 1)))
l03F8 = Mid(l03F8, l03FC + Len(l0380) + 2, Len(l03F8))
Else
Exit Do
End If
Loop
l035E = 1
DoEvents
If encryptchat.Command3D5.Caption = "Start" Then GoTo L6818
GoTo 1012

L6818:
encryptchat.Text5.Text = ""
encryptchat.Text6.Text = ""
encryptchat.Text7.Text = ""
End Function

Function fn448 () As Variant
fn448 = "«—•¤[ "
End Function

Function fn480 () As Variant
fn480 = " ]¤•—»"
End Function

Function fn4B8 (ByVal p040C As String) As Variant
Dim l040E As Variant
Dim l0412 As Variant
Dim l0416 As Variant
Dim l041A As Variant
Dim l041E As String
Dim l0420 As Variant
Dim l0424 As Variant
l040E = findwindow("AOL FRAME25", 0&)
l0412 = Findchildbyclass(l040E, "MDICLIENT")
l0412 = gettopwindow(l0412)

2 :
DoEvents
If l0416 = 0 Then
l0416 = 1
l041A = l0412
Else
l041A = GETnextwindow(l0412, 2)
End If
l041E$ = String(255, 0)
l0420 = GetWindowText(l041A, l041E$, 255)
l0424 = fnC28(l041E$)
If l0424 = p040C Then
fn4B8 = l041A
Exit Function
End If
If l0412 = 0 Then
fn4B8 = 0
Exit Function
End If
l0412 = l041A
GoTo 2
End Function

Function fn4F0 (p042C As String) As Variant
Dim l042E As Variant
Dim l0432 As String
Dim l0434 As Variant
Do Until 2 > 3
For l042E = 1 To 17000
l0432$ = String(255, 0)
l0434 = GetWindowText(l042E, l0432$, 255)
If fnC28(l0432$) = p042C Then
fn4F0 = l042E
Exit Function
End If
Next l042E
Loop
End Function

Function fn528 (ByVal p043C As String, p043E As String) As Variant
Dim l0440 As Variant
Dim l0444 As Variant
Dim l0448 As Variant
Dim l044C As String
Dim l044E As Variant
Dim l0452 As Variant
Dim l0456 As String
Dim l0458 As Variant
l0440 = findwindow("AOL FRAME25", 0&)
l0444 = Findchildbytitle(l0440, p043C)
If l0444 = 0 Then
fn528 = 0
Exit Function
End If
l0448 = GetParent(l0444)
l044C$ = String(255, 0)
l044E = GetWindowText(l0448, l044C$, 255)
l044C$ = fnC28(l044C$)
If l044C$ = p043E Then
fn528 = l0448
Exit Function
Else
12345 :
DoEvents
l0452 = GETnextwindow(l0448, 2)
l0456$ = String(255, 0)
l0458 = GetWindowText(l0452, l0456$, 255)
l0456$ = fnC28(l0456$)
If l0456$ = p043E Then
l0458 = Findchildbytitle(l0452, p043C)
If l0458 > 0 Then
fn528 = l0452
Exit Function
End If
End If
If l0456$ = "" Then
fn528 = 0
Exit Function
Else
l0448 = l0452
GoTo 12345
End If
End If
End Function

Function fn560 () As Variant
Dim l0460 As Variant
Dim l0464 As Variant
Dim l0468 As String
Dim l046A As Variant
Dim l046E As Variant
Dim l0472 As Variant
Dim l0476 As Variant
l0460 = findwindow("AOL FRAME25", 0&)
l0460 = Findchildbyclass(l0460, "MDICLIENT")
l0464 = gettopwindow(l0460)
l0468$ = String(255, 0)
l046A = GetWindowText(l0464, l0468$, 255)
l046E = GETnextwindow(l0464, 2)
l0468$ = String(255, 0)
l046A = GetWindowText(l046E, l0468$, 255)
l0468$ = fnC28(l0468$)
If Len(l0468$) = 0 Then
fn560 = 1
Exit Function
End If
l0472 = 2
13 :
DoEvents
l0476 = GETnextwindow(l046E, 2)
l0468$ = String(255, 0)
l046A = GetWindowText(l0476, l0468$, 255)
l0468$ = fnC28(l0468$)
If Len(l0468$) = 0 Then
fn560 = l0472
Exit Function
Else
l0472 = l0472 + 1
fn560 = l0472
l046E = l0476
GoTo 13
End If
End Function

Function fn598 () As Variant
Dim l047E As Variant
Dim l0482 As Variant
Dim l0486 As Variant
Dim l048A As Variant
Dim l048E As Variant
Dim l0492 As Variant
Dim l0496 As Variant
Dim l049A As Variant
Dim l049E As Variant
11 :
l047E = findwindow("AOL FRAME25", 0&)
l0482 = fn2F8()
l0486 = Findchildbyclass(l0482, "_AOL_LISTBOX")
l048A = SendMessageByNum(l0486, 1036, 0, 0)
For l048E = 0 To l048A - 1
l0492 = vbgetaollist(l048E, l0486)
l0496 = l0496 + l0492 + ","
Next l048E
l047E = findwindow("AOL FRAME25", 0&)
l0482 = fn2F8()
l0486 = Findchildbyclass(l0482, "_AOL_LISTBOX")
l048A = SendMessageByNum(l0486, 1036, 0, 0)
For l048E = 0 To l048A - 1
l0492 = vbgetaollist(l048E, l0486)
l0496 = InStr(l0496, l0492 + ",")
If l0496 Then
Else
l049A = fnAD8(l0492 + ", Entered the room!")
End If
l049E = l049E + l0492 + ","
Next l048E
GoTo 11
End Function

Function fn5D0 (ByVal p04A6 As String) As Variant
Dim l04AC As Variant
help1.Show
For l04AC = 0 To help1.List1.ListCount - 1
If help1.List1.List(l04AC) = p04A6 Then
help1.List1.Selected(l04AC) = True
Exit Function
End If
Next l04AC
End Function

Function fn608 () As Variant
Dim l04B4 As Variant
Dim l04BC As Variant
Dim l04C0 As Variant
Dim l04C4 As Variant
Dim l04C8 As Variant
Dim l04CC As Variant
Dim l04D0 As Variant
Dim l04D4 As Variant
Dim l04D8 As Variant
Dim l04DC As Variant
Dim l04E0 As Variant
Dim l04E4 As Variant
Dim l04E8 As Variant
Dim l04EC As Variant
Do Until 2 > 3
For l04B4 = 0 To imbot.List1.ListCount - 1
l04BC = imbot.List1.List(l04B4)
If l04B4 = 0 Then
l04C0 = l04C0 + "," + l04BC + ","
Else
l04C0 = l04C0 + l04BC + ","
End If
Next l04B4
If fn3D8() <> imbot.Text2.Text Then GoTo 2898
l04C4 = fn2F8()
l04C8 = Findchildbyclass(l04C4, "_AOL_LISTBOX")
l04CC = SendMessageByNum(l04C8, 1036, 0, 0)
For l04B4 = 0 To l04CC - 1
l04BC = vbgetaollist(l04B4, l04C8)
If l04B4 = 0 Then
l04D0 = l04D0 + "," + l04BC + ","
Else
l04D0 = l04D0 + l04BC + ","
End If
l04D4 = InStr(l04C0, "," + l04BC + ",")
If l04D4 = 0 Then
If l04BC = "" Then GoTo 129
l04D8 = l04BC
l04DC = imbot.Text1.Text
l04E0 = InStr(1, l04DC, "{%S}", 1)
If l04E0 Then
If l04E0 = 1 Then
l04E4 = ""
Else
l04E4 = Mid(l04DC, 1, l04E0 - 1)
End If
l04E8 = Mid(l04DC, l04E0 + 4, Len(l04DC))
l04DC = l04E4 + l04D8 + l04E8
Else
l04DC = l04DC
End If
If imbot.Timer1.Interval = 0 Then Exit Function
l04EC = fnB10(l04D8, l04DC)
End If
129 :
If fn3D8() <> imbot.Text2.Text Then GoTo 2898
Next l04B4
If l04D0 <> l04C0 Then
2898 :
imbot.List1.Clear
For l04B4 = 0 To l04CC - 1
l04BC = vbgetaollist(l04B4, l04C8)
imbot.List1.AddItem l04BC
Next l04B4
imbot.Text2.Text = fn3D8()
End If
If imbot.Timer1.Interval = 0 Then Exit Function
DoEvents
l04C0 = ""
l04D0 = ""
If fn3D8() <> imbot.Text2.Text Then GoTo 2898
Loop
End Function

Function fn6B0 () As Variant
Dim l0590 As Variant
l0590 = findwindow("AOL FRAME25", 0&)
End Function

Function fn6E8 (ByVal p0598 As String) As Variant
Dim l059A As Variant
Dim l059E As Variant
Dim l05A2 As Variant
Dim l05A6 As Variant
Dim l05AA As Variant
Dim l05AE As Variant
Dim l05B2 As Variant
l059A = findwindow("AOL FRAME25", 0&)
l059E = fn7C8("Keyword...")
101 :
DoEvents
l05A2 = Findchildbytitle(l059A, "Keyword")
If l05A2 = 0 Then GoTo 101
l05A6 = Findchildbyclass(l05A2, "_AOL_EDIT")
l05AA = Findchildbyclass(l05A2, "_AOL_ICON")
If l05AA = 0 Then GoTo 101
l05AE = SendMessageByString(l05A6, 12, 0, p0598)
113 :
DoEvents
l05B2 = fn1870(l05AA)
l05AE = Findchildbytitle(l05A2, "_AOL_EDIT")
If l05AE > 0 Then GoTo 113
End Function

Function fn720 (ByVal p05BA As String) As Variant
Dim l05C0 As Variant
Dim l05C4 As Variant
Dim l05C8 As Variant
Dim l05CC As Variant
mac.Show
1 :
DoEvents
l05C0 = findwindow(0&, "- --Macros-- -")
If l05C0 = 0 Then GoTo 1
For l05C4 = 0 To mac.List1.ListCount - 1
If mac.List1.List(l05C4) = p05BA$ Then
mac.List1.Selected(l05C4) = True
DoEvents
l05C8 = Findchildbytitle(l05C0, "Scroll")
l05CC = fn1870(l05C8)
Exit Function
End If
DoEvents
Next l05C4
End Function

Function fn758 () As Variant
Dim l05D4 As Variant
Dim l05D8 As Variant
Dim l05DC As Variant
Dim l05E0 As String
Dim l05E2 As Variant
l05D4 = findwindow("AOL FRAME25", 0&)
l05D8 = getmenu(l05D4)
l05DC = getsubmenu(l05D8, 0)
l05E0$ = String(255, 0)
l05E2 = getmenustring(l05D8, l05DC, l05E0$, 100, 100)
l05E0$ = fnC28(l05E0$)
fn758 = l05E0$
End Function

Function fn790 () As Variant
Dim l05EA As Variant
Dim l05EE As Variant
Dim l05F2 As Variant
Dim l05F6 As Variant
Dim l05FA As Variant
Dim l05FE As Variant
Dim l060A As Variant
Dim l060E As Variant
Dim l0612 As Variant
Dim l0616 As Variant
Dim l061A As Variant
Dim l061E As Variant
Dim l0622 As Variant
Dim l0626 As Variant
Dim l062A As Variant
Dim l062E As Variant
Dim l0632 As Variant
Dim l0636 As Variant
Dim l063A As Variant
Dim l063E As Variant
Dim l0642 As Variant
Dim l0646 As Variant
Dim l064A As Variant
Dim l064E As Variant
l05EA = 0
l05EE = findwindow("AOL FRAME25", 0&)
l05F2 = fn2F8()
l05F6 = Findchildbyclass(l05F2, "_AOL_VIEW")
l05FA = SendMessageByNum(l05F6, 14, 0, 0)
Do Until 2 > 3
If l05EA = 1 Then
l05FA = l05FE
End If
Do Until 2 > 3
l05F2 = fn2F8()
l05F6 = Findchildbyclass(l05F2, "_AOL_VIEW")
l05FE = SendMessageByNum(l05F6, 14, 0, 0)
If mchat.acw.Checked = True Then
If iswindowvisible(fn2F8()) = 1 Then
l060A = showwindow(fn2F8(), 0)
mchat.SetFocus
End If
End If
If l05FA <> l05FE Then Exit Do
If mchat.Timer1.Interval = 0 Then GoTo L9544
l060E = fn988()
If fn2F8() = 0 Then
If mchat.Label2.Caption = "** Not In a Room **" Then
Else
mchat.Label2.Caption = "** Not In a Room **"
End If
End If
DoEvents
Loop
mchat.Text1.Text = fn170()
mchat.Text1.SelStart = Len(mchat.Text1.Text)
l0612 = fn170()
If mchat.gomm.Checked = True Then
127 :
DoEvents
l0616 = InStr(1, l0612, "Type", 1)
If l0616 Then
l061A = Mid(l0612, l0616 + 4, Len(l0612))
l061E = InStr(l061A, Chr(34))
If l061E Then
l0622 = Mid(l061A, l061E + 1, Len(l061A))
l0626 = InStr(l0622, Chr(34))
If l0626 Then
l062A = Mid(l0622, 1, l0626 - 1)
For l062E = 1 To Len(l061A)
l0632 = Mid(l0612, l0616 - l062E, 1)
If l0632 = Chr(10) Then
GoTo 149
Else
l0636 = l0632 + l0636
End If
Next l062E

149 :
l063A = InStr(l0636, ":")
l0636 = Mid(l0636, 1, l063A - 1)
For l063E = 0 To mchat.List2.ListCount - 1
If InStr(1, mchat.List2.List(l063E), l0636, 1) Then
If Len(mchat.List2.List(l063E)) = Len(l0636) Then GoTo 159
End If
Next l063E
If l0636 = fnA68() Then GoTo 159
mchat.List2.AddItem l0636
l0642 = getmenu(mchat.hWnd)
l0646 = getsubmenu(l0642, 0)
l064A = getmenuitemcount(l0646)
l064E = l064A - 11
Load mchat.sepa1(l064E)
mchat.sepa1(l064E).Caption = l0636
l0632 = Timeout(.5)
l062E = fnAD8(l062A)
mchat.SetFocus

159 :
l0612 = l061A
GoTo 127
End If
End If
l0616 = InStr(1, l0612, "say", 1)
If l0616 Then GoTo 127
End If
End If
If mchat.acw.Checked = True Then
If iswindowvisible(fn2F8()) = 1 Then
l060A = showwindow(fn2F8(), 0)
mchat.SetFocus
End If
End If
If mchat.lc.Checked = True Then
Open "log.txt" For Output As #1
Print #1, mchat.Text1.Text
Close #1
End If
l05EA = 1
DoEvents
If mchat.Timer1.Interval = 0 Then GoTo L9544
Loop
L9544:
End Function

Function fn7C8 (p0656 As String) As Variant
Dim l0658 As Variant
Dim l065C As Variant
Dim l0660 As Variant
Dim l0664 As Variant
Dim l0668 As Variant
Dim l066C As Variant
Dim l0672 As String
Dim l0674 As Variant
Dim l067A As Variant
l0658 = findwindow("AOL FRAME25", 0&)
l065C = getmenu(l0658)
l0660 = 0
l0664 = 0
l0668 = 1
l066C = 1
Do Until l0668 <= 0
l0668 = getsubmenu(l065C, l0660)
Do Until l066C <= 0
l066C = getmenuitemid(l0668, l0664)
l0672$ = String(50, 0)
l0674 = getmenustring(l065C, l066C, l0672$, 100, 100)
l0672$ = fnC28(l0672$)
l0674 = InStr(l0672$, Chr(9))
If l0674 Then
l0672$ = Mid(l0672$, 1, l0674 - 1)
End If
If l0672$ = p0656 Then
l067A = fn17C8(l0660, l0664)
Exit Function
End If
l0664 = l0664 + 1
DoEvents
Loop
l066C = 1
l0664 = 0
l0660 = l0660 + 1
DoEvents
Loop
End Function

Function fn800 (ByVal p0682 As String) As Variant
Dim l0684 As Variant
Dim l0688 As Variant
Dim l068C As Variant
Dim l0690 As Variant
Dim l0694 As Variant
Dim l0698 As Variant
Dim l069C As String
Dim l069E As Variant
Dim l06A2 As Variant
l0684 = findwindow("AOL FRAME25", 0&)
l0688 = getmenu(l0684)
l068C = 0
l0690 = 0
l0694 = 1
l0698 = 1
Do Until l0694 <= 0
l0694 = getsubmenu(l0688, l068C)
Do Until l0698 <= 0
l0698 = getmenuitemid(l0694, l0690)
l069C$ = String(50, 0)
l069E = getmenustring(l0688, l0698, l069C$, 100, 100)
l069C$ = fnC28(l069C$)
If l069C$ = p0682 Then
l06A2 = fn17C8(l068C, l0690)
Exit Function
End If
l0690 = l0690 + 1
Loop
l0698 = 1
l0690 = 0
l068C = l068C + 1
Loop
End Function

Function fn838 () As Variant
Dim l06AA As Variant
Dim l06AE As String
Dim l06B0 As Variant
Dim l06B4 As Variant
l06AA = findwindow("_AOL_MODAL", 0&)
l06AE$ = String(255, 0)
l06B0 = GetWindowText(l06AA, l06AE$, 250)
l06B4 = InStr(l06AE$, Chr(0))
If l06B4 Then
l06AE$ = Mid(l06AE$, 1, l06B4 - 1)
End If
fn838 = l06AE$
End Function

Function fn870 () As Variant
Dim l06BC As Variant
Dim l06C0 As Variant
Dim l06C4 As String
Dim l06C6 As Variant
l06BC = findwindow("AOL FRAME25", 0&)
l06C0 = Findchildbytitle(l06BC, "Welcome")
l06C4$ = String$(l06C0, 0)
l06C6 = GetWindowText(l06C0, l06C4$, 250)
If l06C6 <= 7 Then
fn870 = 0
Else
fn870 = 1
End If
End Function

Function fn8A8 () As Variant
Dim l06CE As Variant
Dim l06D2 As Variant
Dim l06D6 As String
Dim l06D8 As Variant
l06CE = findwindow("AOL FRAME25", 0&)
l06D2 = Findchildbytitle(l06CE, "Welcome")
l06D6$ = String(30, 0)
l06D8 = GetWindowText(l06D2, l06D6$, 255)
If l06D8 <= 7 Then
fn8A8 = 0
Else
fn8A8 = 1
End If
End Function

Function fn8E0 (p06E0 As String, p06E2 As String) As Variant
Dim l06E4 As Variant
Dim l06E8 As Variant
Dim l06EC As Variant
Dim l06F0 As String
Dim l06F2 As Variant
Dim l06F6 As Variant
Dim l06FA As Variant
Dim l06FE As Variant
Dim l0702 As Variant
Dim l0706 As String
Dim l0708 As Variant
Dim l070C As Variant
Dim l0710 As Variant
l06E4 = findwindow("AOL FRAME25", 0&)
l06E8 = Findchildbytitle(l06E4, p06E0)
l06EC = GetParent(l06E8)
l06F0$ = String(255, 0)
l06F2 = GetWindowText(l06EC, l06F0$, 250)
l06F0$ = fnC28(l06F0$)
If l06F0$ = p06E2 Then
l06F6 = GETnextwindow(l06E8, 2)
l06FA = fn1870(l06F6)
Exit Function
Else
134 :
DoEvents
l06FE = GETnextwindow(l06EC, 2)
l0702 = Findchildbytitle(l06FE, p06E0)
l06EC = l06FE
If l0702 = 0 Then GoTo 134
l0706$ = String(255, 0)
l0708 = GetWindowText(l06FE, l0706$, 250)
l0706$ = fnC28(l0706$)
If l0706$ = p06E2 Then
l0708 = Findchildbytitle(l06FE, p06E0)
l070C = GETnextwindow(l0708, 2)
l0710 = fn1870(l070C)
Exit Function
Else
l06EC = l06FE
GoTo 134
End If
End If
End Function

Function fn918 () As Variant
Dim l0718 As Variant
Dim l071C As Variant
Dim l0720 As Variant
Dim l0724 As Variant
Dim l0728 As Variant
Dim l072C As String
Dim l0732 As Variant
Dim l0736 As Variant
Dim l073A As Variant
Dim l073E As Variant
Dim l0742 As Variant
Dim l0746 As Variant
Dim l074A As Variant
Dim l074E As Variant
Dim l0752 As Variant
Dim l0756 As Variant
Dim l075A As Variant
Dim l075E As String
Dim l0760 As String
Dim l0762 As Variant
Dim l0766 As Variant
Dim l076A As Variant
Dim l076E As Variant
Dim l0772 As Variant
Dim l0776 As Variant
Dim l077A As Variant
Dim l077E As Variant
Dim l0782 As Variant
Dim l0786 As Variant
l0718 = 0
l071C = findwindow("AOL FRAME25", 0&)
l0720 = fn2F8()
l0724 = l0720
l0728 = Findchildbyclass(l0724, "_AOL_VIEW")
l072C$ = getapitext(l0728)
warezre.Text1.Text = l072C$
For l0732 = 0 To Len(warezre.Text1.Text) - 1
l0736 = Mid(warezre.Text1.Text, Len(warezre.Text1.Text) - l0732, 1)
If l0736 = Chr(10) Then Exit For
l073A = l0736 + l073A
DoEvents
Next l0732
l073E = warezre.Text1.Text
Do Until 2 > 3
l0742 = InStr(l073E, l073A)
If l0742 Then
warezre.Text2.Text = Trim(Str(Int(warezre.Text2.Text + 1)))
l073E = Mid(l073E, l0742 + Len(l073A) + 2, Len(l073E))
Else
Exit Do
End If
DoEvents
Loop
l0746 = findwindow("AOL FRAME25", 0&)
l074A = fn2F8()
l074E = l074A
l0752 = Findchildbyclass(l074E, "_AOL_VIEW")
l0756 = SendMessageByNum(l0752, 14, 0, 0)
1011 :
If l0718 = 1 Then
l0756 = l075A
End If
Do Until 2 > 3
l074E = fn2F8()
l0752 = Findchildbyclass(l074E, "_AOL_VIEW")
l075A = SendMessageByNum(l0752, 14, 0, 0)
If l0756 <> l075A Then Exit Do
DoEvents
If fn2F8() = 0 Then
warezre.Text6.Text = "1"
warezre.Timer2.Interval = 0
warezre.Timer2.Enabled = False
warezre.Timer1.Interval = 0
warezre.Timer1.Enabled = False
warezre.Command3D1.Caption = "Start"
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo LA7AC
End If
If warezre.Command3D1.Caption = "Start" Then GoTo LA7AC
Loop
l075E$ = getapitext(l0752)
warezre.Text1.Text = l075E$
l0760$ = warezre.Text1.Text
For l0762 = 1 To Int(warezre.Text2.Text)
DoEvents
l0766 = InStr(l0760$, l073A)
If l0766 Then
l0760$ = Mid(l0760$, l0766 + Len(l073A) + 2, Len(l0760$))
Else
GoTo 155
End If
Next l0762
If warezre.Command3D1.Caption = "Start" Then GoTo LA7AC

155 :
l076A = l0760$

300 :
DoEvents
l076E = InStr(l076A, Chr(13) + Chr(10))
If l076E Then
l0772 = Mid(l076A, 1, l076E - 1)
warezre.Text3.Text = l0772
l0776 = InStr(1, warezre.Text3.Text, ":" + Chr(9) + warezre.got.Text, 1)
If l0776 Then
l077A = InStr(warezre.Text3.Text, ":")
If l077A = 0 Then GoTo 4
l077E = Mid(warezre.Text3.Text, 1, l077A - 1)
If l077E = fnA68() Then GoTo 4
l0782 = fn950(l077E)
End If
4 :
l076A = Mid(l076A, l076E + 2, Len(l076A))
If warezre.Command3D1.Caption = "Start" Then GoTo LA7AC
GoTo 300
Else
warezre.Text3.Text = l076A
l0776 = InStr(1, warezre.Text3.Text, ":" + Chr(9) + warezre.got.Text, 1)
If l0776 Then
l077A = InStr(warezre.Text3.Text, ":")
If l077A = 0 Then GoTo 44
l077E = Mid(warezre.Text3.Text, 1, l077A - 1)
If l077E = fnA68() Then GoTo 44
l0782 = fn950(l077E)
End If
44 :
End If
l073A = warezre.Text3.Text
l0776 = warezre.Text1.Text
warezre.Text2.Text = "0"
Do Until 2 > 3
l0786 = InStr(l0776, l073A)
If l0786 Then
warezre.Text2.Text = Trim(Str(Int(warezre.Text2.Text + 1)))
l0776 = Mid(l0776, l0786 + Len(l073A) + 2, Len(l0776))
Else
Exit Do
End If
If warezre.Command3D1.Caption = "Start" Then GoTo LA7AC
Loop
l0718 = 1
If warezre.Command3D1.Caption = "Start" Then GoTo LA7AC
DoEvents
GoTo 1011

LA7AC:
warezre.Text1.Text = ""
warezre.Text2.Text = "2"
warezre.Text3.Text = ""
End Function

Function fn950 (p078E As Variant) As Variant
Dim l0792 As Variant
Dim l0796 As Variant
Dim l079A As Variant
Dim l079E As Variant
Dim l07A2 As Variant
Dim l07A6 As Variant
Dim l07AA As Variant
Dim l07AE As Variant
Dim l07B2 As Variant
Dim l07B6 As Variant
Dim l07BA As Variant
Dim l07BE As Variant
Dim l07C2 As String
Dim l07C4 As Variant
Dim l07C8 As Variant
Dim l07CC As Variant
Dim l07D0 As Variant
Dim l07D4 As Variant
Dim l07D8 As Variant
Dim l07DC As Variant
Dim l07E0 As Variant
Dim l07E4 As Variant
Dim l07E8 As Variant
Dim l07EC As Variant
Dim l07F0 As Variant
Dim l07F4 As Variant
Dim l07F8 As String
Dim l07FA As Variant
On Error GoTo 1003
For l0792 = 0 To warezre.List1.ListCount - 1
l0796 = InStr(1, warezre.List1.List(l0792), p078E, 1)
If l0796 Then GoTo 45
Next l0792
warezre.List1.AddItem p078E
warezre.Label2.Caption = Trim(Str(warezre.List1.ListCount))
warezre.Command3D3.Enabled = True
If warezre.Option3D1.Value = True Then
l079A = InStr(1, warezre.Text4.Text, "{%S}", 1)
If l079A Then
If l079A = 1 Then
l079E = ""
Else
l079E = Mid(warezre.Text4.Text, 1, l079A - 1)
End If
l07A2 = Mid(warezre.Text4.Text, l079A + 4, Len(warezre.Text4.Text))
l07A6 = l079E + p078E + l07A2
Else
l07A6 = warezre.Text4.Text
End If
l07AA = AOLVersion()
If l07AA = 5 Then
l07AE = fn7C8("Send an Instant Message")
l07B2 = findwindow("AOL FRAME25", 0&)
89 :
DoEvents
l07B6 = Findchildbytitle(l07B2, "Send Instant Message")
If l07B6 = 0 Then GoTo 89
l07BA = Findchildbyclass(l07B6, "_AOL_EDIT")
l07BE = SendMessageByString(l07BA, 12, 0, p078E)
l07C2$ = String(Len(p078E), 0)
l07C4 = SendMessageByString(l07BA, 13, Len(p078E) + 1, l07C2$)
If l07C2$ <> p078E Then GoTo 89
l07C8 = GETnextwindow(l07BA, 2)
l07CC = SendMessageByString(l07C8, 12, 0, l07A6)
l07D0 = GETnextwindow(l07C8, 2)
l07D4 = fn1870(l07D0)
67 :
DoEvents
l07D8 = Findchildbytitle(l07B6, "Available?")
If l07D8 > 0 Then
l07DC = findwindow("#32770", "America Online")
If l07DC > 0 Then
l07E0 = Findchildbytitle(l07DC, "OK")
l07E4 = fn1870(l07E0)
l07E8 = SendMessageByNum(GetParent(l07D8), 16, 0, 0)
GoTo 45
End If
GoTo 67
Else
GoTo 45
End If
Else
l07AE = fn7C8("Send an Instant Message")
l07B2 = findwindow("AOL FRAME25", 0&)
83 :
DoEvents
l07B6 = Findchildbytitle(l07B2, "Send Instant Message")
If l07B6 = 0 Then GoTo 83
l07BA = Findchildbyclass(l07B6, "_AOL_EDIT")
l07BE = SendMessageByString(l07BA, 12, 0, p078E)
l07C2$ = String(Len(p078E), 0)
l07C4 = SendMessageByString(l07BA, 13, Len(p078E) + 1, l07C2$)
If l07C2$ <> p078E Then GoTo 83
l07C8 = Findchildbyclass(l07B6, "RICHCNTL")
l07CC = SendMessageByString(l07C8, 12, 0, l07A6)
l07D0 = GETnextwindow(l07C8, 2)
l07D4 = fn1870(l07D0)
68 :
DoEvents
l07D8 = Findchildbytitle(l07B6, "To:")
If l07D8 > 0 Then
l07DC = findwindow("#32770", "America Online")
If l07DC > 0 Then
l07E0 = Findchildbytitle(l07DC, "OK")
l07E4 = fn1870(l07E0)
l07E8 = SendMessageByNum(GetParent(l07D8), 16, 0, 0)
GoTo 45
End If
GoTo 68
End If
End If
Else
l07EC = Timeout(.6)
l07F0 = fnAD8(fn448() + "<" + p078E + "> " + warezre.Text4.Text + fn480())
End If
45 :
GoTo 1004
1003 :
l07F4 = Err
l07F8$ = warezre.Caption
l07FA = fn250(l07F4, l07F8$)
Exit Function

1004 :
End Function

Function fn988 () As Variant
Dim l0802 As Variant
Dim l0806 As Variant
Dim l080A As Variant
Dim l080E As Variant
Dim l0812 As Variant
Dim l0816 As Variant
Dim l081A As Variant
Dim l081E As Variant
Dim l0822 As Variant
Dim l0826 As Variant
Dim l082A As Variant
Dim l082E As Variant
Dim l0832 As Variant
l0802 = fn2F8()
l0806 = Findchildbyclass(l0802, "_AOL_LISTBOX")
l080A = SendMessageByNum(l0806, 1036, 0, 0)
For l080E = 0 To l080A - 1
l0812 = vbgetaollist(l080E, l0806)
If l080E = 0 Then
l0816 = l0816 + "," + l0812 + ","
Else
l0816 = l0816 + l0812 + ","
End If
Next l080E
If fn3D8() <> mchat.Label2.Caption Then GoTo 289
For l080E = 0 To mchat.List1.ListCount - 1
l0812 = mchat.List1.List(l080E)
If l080E = 0 Then
l081A = l081A + "," + l0812 + ","
Else
l081A = l081A + l0812 + ","
End If
l081E = InStr(l0816, "," + l0812 + ",")
If l0816 = "" Then GoTo LB90C
If l081E = 0 Then
If mchat.nwml.Checked = False Then GoTo 239
If l0812 = "" Then GoTo 239
If fn3D8() <> mchat.Label2.Caption Then GoTo 289
l0822 = fn2F8()
l0826 = Findchildbyclass(l0822, "_AOL_VIEW")
l082A = SendMessageByString(l0826, 12, 0, Chr(13) + Chr(10) + "FateHost:" + Chr(9) + l0812 + " Has left the room.")
End If
239 :
Next l080E
l0802 = fn2F8()
l0806 = Findchildbyclass(l0802, "_AOL_LISTBOX")
l080A = SendMessageByNum(l0806, 1036, 0, 0)
For l080E = 0 To l080A - 1
l0812 = vbgetaollist(l080E, l0806)
l081E = InStr(l081A, "," + l0812 + ",")
If l081E = 0 Then
If l081A = "" Then GoTo LB90C
If fn3D8() <> mchat.Label2.Caption Then GoTo 289
If mchat.nwma.Checked = False Then GoTo 23
If l0812 = "" Then GoTo 23
If mchat.cwa.Checked = True Then
If InStr(1, Mid(l0812, 1, 8), "Catwatch", 1) Then l0802 = SendMessageByNum(fn2F8(), 16, 0, 0)
End If
If mchat.gim.Checked = True Then
l082E = Timeout(.4)
l0832 = fnAD8(fn448() + "Fate X³ - " + l0812 + " has entered the room" + fn480())
mchat.SetFocus
End If
l0822 = fn2F8()
l0826 = Findchildbyclass(l0822, "_AOL_VIEW")
l082A = SendMessageByString(l0826, 12, 0, Chr(13) + Chr(10) + "FateHost:" + Chr(9) + l0812 + " Has entered the room.")
End If
23 :
Next l080E
LB90C:
If l0816 <> l081A Then
289 :
mchat.List1.Clear
l0802 = fn2F8()
l0806 = Findchildbyclass(l0802, "_AOL_LISTBOX")
For l080E = 0 To l080A - 1
l0812 = vbgetaollist(l080E, l0806)
mchat.List1.AddItem l0812
Next l080E
If fn2F8() = 0 Then
If mchat.Label2.Caption = "** Not In a Room **" Then
Else
mchat.Label2.Caption = "** Not In a Room **"
End If
Else
mchat.Label2.Caption = fn3D8()
End If
mchat.Label1.Caption = "People Here: " + mchat.List1.ListCount
End If
l081A = ""
l0816 = ""
End Function

Function fn9C0 () As Variant
Dim l083A As Variant
Dim l083E As Variant
Dim l0842 As Variant
Dim l0846 As Variant
Dim l084A As String
Dim l084C As Variant
Dim l0850 As Variant
Dim l0854 As Variant
Dim l0858 As Variant
Dim l085C As Variant
Dim l0860 As Variant
l083A = fn2F8()
If l083A > 0 Then
l083E = l083A
l0842 = Findchildbyclass(l083E, "_AOL_EDIT")
l0846 = GetParent(l0842)
l084A$ = String(255, 0)
l084C = GetWindowText(l0846, l084A$, 250)
l0850 = l084A$
l0854 = InStr(l0850, Chr(0))
l0850 = Mid(l0850, 1, l0854 - 1)
Do Until 2 > 3
l0858 = InStr(l0850, Chr(32))
If l0858 Then
l085C = Mid(l0850, 1, l0858 - 1)
l0860 = Mid(l0850, l0858 + 1, Len(l0850))
l0850 = l085C + l0860
Else
GoTo 100
End If
Loop
100 :
l0850 = UCase(l0850)
End If
fn9C0 = l0850
End Function

Function fn9F8 () As Variant
Dim l0868 As Variant
Dim l086C As Variant
Dim l0870 As Variant
Dim l0874 As Variant
Dim l0878 As Variant
Dim l087C As String
Dim l0882 As Variant
Dim l0886 As Variant
Dim l088A As Variant
Dim l088E As Variant
Dim l0892 As Variant
Dim l0896 As Variant
Dim l089A As Variant
Dim l089E As Variant
Dim l08A2 As Variant
Dim l08A6 As Variant
Dim l08AA As Variant
Dim l08AE As String
Dim l08B0 As String
Dim l08B2 As Variant
Dim l08B6 As Variant
Dim l08BA As Variant
Dim l08BE As Variant
Dim l08C2 As Variant
Dim l08C6 As Variant
Dim l08CA As Variant
Dim l08CE As Variant
Dim l08D2 As Variant
l0868 = 0
l086C = findwindow("AOL FRAME25", 0&)
l0870 = fn2F8()
l0874 = l0870
l0878 = Findchildbyclass(l0874, "_AOL_VIEW")
l087C$ = getapitext(l0878)
scram.Text11.Text = l087C$
For l0882 = 0 To Len(scram.Text11.Text) - 1
l0886 = Mid(scram.Text11.Text, Len(scram.Text11.Text) - l0882, 1)
If l0886 = Chr(10) Then Exit For
l088A = l0886 + l088A
DoEvents
Next l0882
l088E = scram.Text11.Text
Do Until 2 > 3
l0892 = InStr(l088E, l088A)
If l0892 Then
scram.Text12.Text = Trim(Str(Int(scram.Text12.Text + 1)))
l088E = Mid(l088E, l0892 + Len(l088A) + 2, Len(l088E))
Else
Exit Do
End If
DoEvents
Loop
l0896 = findwindow("AOL FRAME25", 0&)
l089A = fn2F8()
l089E = l089A
l08A2 = Findchildbyclass(l089E, "_AOL_VIEW")
l08A6 = SendMessageByNum(l08A2, 14, 0, 0)

1013 :
If l0868 = 1 Then
l08A6 = l08AA
End If
Do Until 2 > 3
l089E = fn2F8()
l08A2 = Findchildbyclass(l089E, "_AOL_VIEW")
l08AA = SendMessageByNum(l08A2, 14, 0, 0)
If l08A6 <> l08AA Then Exit Do
If scram.Command3D2.Caption = "Start" Then GoTo LC818
If fn2F8() = 0 Then
scram.Text1.Enabled = True
scram.Text5.Enabled = True
scram.Command3D2.Caption = "Start"
scram.Timer1.Interval = 0
scram.Timer1.Enabled = False
scram.Timer2.Interval = 0
scram.Timer2.Enabled = False
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo LC818
End If
DoEvents
Loop
l08AE$ = getapitext(l08A2)
scram.Text11.Text = l08AE$
l08B0$ = scram.Text11.Text
If scram.Command3D2.Caption = "Start" Then GoTo LC818
For l08B2 = 1 To Int(scram.Text12.Text)
DoEvents
l08B6 = InStr(l08B0$, l088A)
If l08B6 Then
l08B0$ = Mid(l08B0$, l08B6 + Len(l088A) + 2, Len(l08B0$))
Else
GoTo 1553
End If
Next l08B2

1553 :
l08BA = l08B0$

3003 :
DoEvents
l08BE = InStr(l08BA, Chr(13) + Chr(10))
If l08BE Then
l08C2 = Mid(l08BA, 1, l08BE - 1)
scram.Text13.Text = l08C2
l0892 = InStr(1, l08C2, ":" + Chr(9) + scram.Text1.Text, 1)
If l0892 Then
l088E = InStr(l08C2, ":")
If l088E = 0 Then GoTo 9
l0896 = Mid(l08C2, 1, l088E - 1)
l089A = Mid(l08C2, l088E + 2, Len(l08C2))
If Len(l089A) <> Len(scram.Text1.Text) Then GoTo 9
l089A = Mid(l0896, 1, 1)
If l089A = Chr(10) Then l0896 = Mid(l0896, 2, Len(l0896))
If l0896 = fnA68() Then GoTo 9
l08C6 = fnAA0(fn448() + "<" + l0896 + "> Answered Correctly" + fn480())
l08CA = fnA30(l0896)
GoTo LC818
End If
If scram.Command3D2.Caption = "Start" Then GoTo LC818
9 :
l08BA = Mid(l08BA, l08BE + 2, Len(l08BA))
GoTo 3003
Else
scram.Text13.Text = l08BA
l0892 = InStr(1, l08BA, ":" + Chr(9) + scram.Text1.Text, 1)
If l0892 Then
l088E = InStr(l08BA, ":")
If l088E = 0 Then GoTo 14
l0896 = Mid(l08BA, 1, l088E - 1)
l089A = Mid(l08BA, l088E + 2, Len(l08BA))
For l086C = 1 To Len(l089A)
l0870 = Mid(l089A, l086C, 1)
l0874 = Asc(l0870)
Next l086C
If Len(l089A) <> Len(scram.Text1.Text) Then GoTo 14
If l0896 = fnA68() Then GoTo 14
l08C6 = fnAA0(fn448() + "<" + l0896 + "> Answered Correctly" + fn480())
l08CA = fnA30(l0896)
GoTo LC818
End If
14 :
 If scram.Command3D2.Caption = "Start" Then GoTo LC818
End If
l088A = scram.Text13.Text
l08CE = scram.Text11.Text
scram.Text12.Text = "0"
Do Until 2 > 3
l08D2 = InStr(l08CE, l088A)
If l08D2 Then
scram.Text12.Text = Trim(Str(Int(scram.Text12.Text + 1)))
l08CE = Mid(l08CE, l08D2 + Len(l088A) + 2, Len(l08CE))
Else
Exit Do
End If
Loop
l0868 = 1
If scram.Command3D2.Caption = "Start" Then GoTo LC818
DoEvents
GoTo 1013

LC818:
scram.Text11.Text = ""
scram.Text12.Text = "0"
scram.Text13.Text = ""
End Function

Function fnA30 (p08DA As Variant) As Variant
Dim l08DE As Variant
Dim l08E2 As Variant
Dim l08E6 As Variant
Dim l08EA As Variant
Dim l08EE As Variant
Dim l08F2 As Variant
Dim l08F6 As Variant
Dim l08FA As Variant
Dim l08FE As Variant
Dim l0902 As Variant
Dim l0906 As Variant
Dim l090A As Variant
If scram.List1.ListCount = 0 Then
scram.List2.AddItem p08DA
l08DE = 1
scram.List1.AddItem scram.Combo1.Text
GoTo 91
End If
For l08E2 = 0 To scram.List2.ListCount - 1
l08E6 = InStr(1, scram.List2.List(l08E2), p08DA, 1)
If l08E6 Then
If Len(p08DA) = Len(scram.List2.List(l08E2)) Then
l08EA = scram.List1.List(l08E2)
l08EE = Int(l08EA) + Int(scram.Combo1.Text)
l08F2 = scram.List2.List(l08E2)
If scram.List1.ListCount = 0 Then
GoTo 96
End If
If scram.List1.ListCount = 1 Then
scram.List1.List(0) = l08EE
GoTo 91
End If
If l08E2 = 0 Then
scram.List1.List(l08E2) = l08EE
GoTo 91
End If
scram.List2.RemoveItem l08E2
scram.List1.RemoveItem l08E2
l08F6 = 1
For l08FA = 0 To scram.List1.ListCount - 1
l08FE = scram.List1.List(l08FA)
If Int(l08EE) = Int(l08FE) Then
scram.List1.AddItem Int(l08EE), l08FA
scram.List2.AddItem l08F2, l08FA
GoTo 91
End If
If Int(l08EE) > Int(scram.List1.List(l08FA)) Then
scram.List1.AddItem Int(l08EE), l08FA
scram.List2.AddItem l08F2, l08FA
GoTo 91
End If
l08F6 = 0
92 :
DoEvents
On Error Resume Next
If Int(l08EE) < Int(scram.List1.List(l08FA + l08F6)) Then
If Int(l08EE) > Int(scram.List1.List(l08FA)) Then
scram.List1.AddItem Int(l08EE), l08FA
scram.List2.AddItem l08F2, l08FA
GoTo 91
End If
If Err Then
scram.List1.AddItem Int(l08EE), scram.List1.ListCount
scram.List2.AddItem l08F2, scram.List2.ListCount
GoTo 91
End If
l08F6 = l08F6 + 1
GoTo 92
End If
Next l08FA
End If
Else
GoTo 90
End If
90 :
Next l08E2

96 :
scram.List2.AddItem p08DA
l08DE = 1
scram.List1.AddItem scram.Combo1.Text
If scram.List2.ListCount <> 0 Then
scram.List2.RemoveItem Int(scram.List2.ListCount - 1)
scram.List1.RemoveItem Int(scram.List1.ListCount - 1)
l08F6 = 1
For l08FA = 0 To scram.List1.ListCount - 1
l08FE = scram.List1.List(l08FA)
If Int(scram.Combo1.Text) = Int(l08FE) Then
scram.List1.AddItem Int(scram.Combo1.Text), l08FA
scram.List2.AddItem p08DA, l08FA
GoTo 91
End If
If Int(scram.Combo1.Text) > Int(scram.List1.List(l08FA)) Then
scram.List1.AddItem Int(scram.Combo1.Text), l08FA
scram.List2.AddItem p08DA, l08FA
GoTo 91
End If
l08F6 = 0
922 :
DoEvents
On Error Resume Next
If Int(scram.Combo1.Text) < Int(scram.List1.List(l08FA + l08F6)) Then
If Int(scram.Combo1.Text) > Int(scram.List1.List(l08FA)) Then
scram.List1.AddItem Int(scram.Combo1.Text), l08FA
scram.List2.AddItem p08DA, l08FA
GoTo 91
End If
If Err Then
scram.List1.AddItem Int(scram.Combo1.Text), scram.List1.ListCount
scram.List2.AddItem p08DA, scram.List2.ListCount
GoTo 91
End If
l08F6 = l08F6 + 1
GoTo 922
End If
Next l08FA
End If

91 :
scram.List3.Clear
For l0902 = 0 To scram.List2.ListCount - 1
l0906 = scram.List2.List(l0902)
l090A = scram.List1.List(l0902)
scram.List3.AddItem l0906 + "   -   " + l090A
Next l0902
scram.Label1.Caption = scram.List2.List(0)
scram.Command3D2.Caption = "Start"
scram.Timer1.Interval = 0
scram.Timer1.Enabled = False
scram.Timer2.Interval = 0
scram.Timer2.Enabled = False
scram.Text2.Text = ""
scram.Text1.Enabled = True
scram.Text5.Enabled = True
End Function

Function fnA68 () As Variant
Dim l0912 As Variant
Dim l0916 As Variant
Dim l091A As String
Dim l091C As Variant
Dim l0920 As Variant
Dim l0924 As Variant
l0912 = findwindow("AOL FRAME25", 0&)
l0916 = Findchildbytitle(l0912, "Welcome")
l091A$ = String(30, 0)
l091C = GetWindowText(l0916, l091A$, 250)
l0920 = Mid(l091A$, 10, 10)
l0924 = InStr(l0920, "!")
If l0924 Then
l0920 = Mid(l0920, 1, l0924 - 1)
End If
fnA68 = l0920
End Function

Function fnAA0 (ByVal p092C As String) As Variant
Dim l092E As Variant
Dim l0932 As Variant
Dim l0936 As Variant
Dim l093A As Variant
Dim l093E As String
Dim l0940 As Variant
Dim l0944 As Variant
Dim l0948 As Variant
Dim l094C As Variant
Dim l0950 As Variant
l092E = findwindow("AOL FRAME25", 0&)
l0932 = fn2F8()
l0936 = l0932
l093A = Findchildbyclass(l0936, "_AOL_EDIT")
l093E$ = String(255, 0)
l0940 = SendMessageByString(l093A, 13, 255, l093E$)
l0944 = SendMessageByString(l093A, 12, 0, p092C)
l0948 = GETnextwindow(l093A, 2)
l094C = fn1870(l0948)
l0950 = SendMessageByString(l093A, 12, 0, l093E$)
DoEvents
End Function

Function fnAD8 (p0958 As Variant) As Variant
Dim l095C As Variant
Dim l0960 As Variant
Dim l0964 As Variant
Dim l0968 As Variant
Dim l096C As String
Dim l096E As Variant
Dim l0972 As Variant
Dim l0976 As Variant
Dim l097A As Variant
Dim l097E As Variant
l095C = findwindow("AOL FRAME25", 0&)
l0960 = fn2F8()
l0964 = l0960
l0968 = Findchildbyclass(l0964, "_AOL_EDIT")
l096C$ = String(255, 0)
l096E = SendMessageByString(l0968, 13, 255, l096C$)
l0972 = SendMessageByString(l0968, 12, 0, p0958)
l0976 = GETnextwindow(l0968, 2)
l097A = fn1870(l0976)
l097E = SendMessageByString(l0968, 12, 0, l096C$)
DoEvents
End Function

Function fnB10 (p0986 As Variant, p098A As Variant) As Variant
Dim l098E As Variant
Dim l0992 As Variant
Dim l0996 As Variant
Dim l099A As Variant
Dim l099E As Variant
Dim l09A2 As Variant
Dim l09A6 As String
Dim l09A8 As Variant
Dim l09AC As Variant
Dim l09B0 As Variant
Dim l09B4 As String
Dim l09B6 As Variant
Dim l09BA As Variant
Dim l09BE As Variant
Dim l09C2 As Variant
Dim l09C6 As Variant
Dim l09CA As Variant
Dim l09CE As Variant
Dim l09D2 As Variant
Dim l09D6 As Variant
Dim l09DA As Variant
Dim l09DE As Variant
Dim l09E2 As Variant
Dim l09E6 As Variant
l098E = findwindow("AOL FRAME25", 0&)
34 :
l0992 = fn7C8("Send an Instant Message")
l0996 = AOLVersion()
If l0996 = 8 Then
36 :
DoEvents
l099A = Findchildbytitle(l098E, "Send Instant Message")
If l099A = 0 Then GoTo 36
Else
l099A = Findchildbytitle(l098E, "Send Instant Message")
If l099A = 0 Then GoTo 34
End If
l099E = Findchildbyclass(l099A, "_AOL_EDIT")
l09A2 = SendMessageByString(l099E, 12, 0, p0986)
l09A6$ = String(Len(p0986), 0)
l09A8 = SendMessageByString(l099E, 13, Len(p0986) + 1, l09A6$)
If l09A6$ <> p0986 Then GoTo 36

56 :
DoEvents
l09AC = Findchildbytitle(l098E, "To:")
l09B0 = GetParent(l09AC)
l09B4$ = String(50, 0)
l09B6 = GetWindowText(l09B0, l09B4$, 250)
If l09B4$ Like "*Send Instant Message*" Then
Else
GoTo 56
End If
l09BA = Findchildbytitle(l099A, "Send")
l09BE = GETnextwindow(l09BA, 3)
If l0996 = 5 Then
Else
l09BE = Findchildbyclass(l099A, "RICHCNTL")
End If
l09C2 = Chr(13) + Chr(10)
l09C6 = SendMessageByString(l09BE, 12, 0, p098A)
If l0996 = 5 Then
l09CA = fn1870(l09BA)
Else
l09CE = GETnextwindow(l09BE, 2)
l09D2 = fn1870(l09CE)
End If
Do Until 2 > 3
If l0996 = 5 Then
l09D6 = Findchildbytitle(l098E, "Available?")
Else
l09AC = Findchildbytitle(l098E, "To:")
l09B0 = GetParent(l09AC)
l09B4$ = String(50, 0)
l09B6 = GetWindowText(l09B0, l09B4$, 250)
If l09B4$ Like "*Send Instant Message*" Then
l09D6 = Findchildbytitle(l09B0, "To:")
Else
l09D6 = 0
End If
End If
If l09D6 = 0 Then GoTo 1344
l09DA = findwindow("#32770", "America Online")
If l09DA > 0 Then
l09DE = Findchildbytitle(l09DA, "OK")
l09E2 = fn1870(l09DE)
l09E6 = SendMessageByNum(GetParent(l09D6), 16, 0, 0)
GoTo 1344
End If
DoEvents
Loop
1344 :
End Function

Function fnB48 () As Variant
Dim l09EE As Variant
Dim l09F2 As Variant
Dim l09F6 As Variant
Dim l09FA As Variant
Dim l09FE As Variant
Dim l0A02 As String
Dim l0A08 As Variant
Dim l0A0C As Variant
Dim l0A10 As Variant
Dim l0A14 As Variant
Dim l0A18 As Variant
Dim l0A1C As Variant
Dim l0A20 As Variant
Dim l0A24 As Variant
Dim l0A28 As Variant
Dim l0A2C As Variant
Dim l0A30 As Variant
Dim l0A34 As String
Dim l0A36 As String
Dim l0A38 As Variant
Dim l0A3C As Variant
Dim l0A40 As Variant
Dim l0A44 As Variant
Dim l0A48 As Variant
Dim l0A4C As Variant
Dim l0A50 As Variant
Dim l0A54 As Variant
Dim l0A58 As Variant
Dim l0A5C As Variant
Dim l0A60 As Variant
l09EE = 0
l09F2 = findwindow("AOL FRAME25", 0&)
l09F6 = fn2F8()
l09FA = l09F6
l09FE = Findchildbyclass(l09FA, "_AOL_VIEW")
l0A02$ = getapitext(l09FE)
server.Text11.Text = l0A02$
For l0A08 = 0 To Len(server.Text11.Text) - 1
l0A0C = Mid(server.Text11.Text, Len(server.Text11.Text) - l0A08, 1)
If l0A0C = Chr(10) Then Exit For
l0A10 = l0A0C + l0A10
DoEvents
Next l0A08
l0A14 = server.Text11.Text
Do Until 2 > 3
l0A18 = InStr(l0A14, l0A10)
If l0A18 Then
server.Text12.Text = Trim(Str(Int(server.Text12.Text + 1)))
l0A14 = Mid(l0A14, l0A18 + Len(l0A10) + 2, Len(l0A14))
Else
Exit Do
End If
DoEvents
Loop
l0A1C = findwindow("AOL FRAME25", 0&)
l0A20 = fn2F8()
l0A24 = l0A20
l0A28 = Findchildbyclass(l0A24, "_AOL_VIEW")
l0A2C = SendMessageByNum(l0A28, 14, 0, 0)
1014 :
If l09EE = 1 Then
l0A2C = l0A30
End If
Do Until 2 > 3
l0A24 = fn2F8()
l0A28 = Findchildbyclass(l0A24, "_AOL_VIEW")
l0A30 = SendMessageByNum(l0A28, 14, 0, 0)
If server.tell.Text = "1" Then
server.tell.Text = ""
l09F2 = fnAD8(fn448() + "Fate X³ Warez Server By MaGuS" + fn480())
DoEvents
DoEvents
l09F6 = fnAD8(fn448() + "Type " + Chr(34) + "/" + fnA68() + " Send X" + Chr(34) + " X - Index #" + fn480())
DoEvents
DoEvents
l09FA = fnAD8(fn448() + "Type " + Chr(34) + "/" + fnA68() + " Send List" + Chr(34) + " To Get List" + fn480())
DoEvents
DoEvents
End If
If l0A2C <> l0A30 Then Exit Do
If server.Command3D3.Caption = "Start" Then GoTo LF104
If fn2F8() = 0 Then
server.Command3D3.Caption = "Start"
server.Timer1.Interval = 0
server.Timer1.Enabled = False
MsgBox "The chat window has been closed, shutting down bot!", 16
GoTo LF104
End If
DoEvents
Loop
l0A34$ = getapitext(l0A28)
server.Text11.Text = l0A34$
l0A36$ = server.Text11.Text
For l0A38 = 1 To Int(server.Text12.Text)
DoEvents
l0A3C = InStr(l0A36$, l0A10)
If l0A3C Then
l0A36$ = Mid(l0A36$, l0A3C + Len(l0A10) + 2, Len(l0A36$))
Else
GoTo 1554
End If
Next l0A38

1554 :
l0A40 = l0A36$

3004 :
DoEvents
l0A44 = InStr(l0A40, Chr(13) + Chr(10))
If l0A44 Then
l0A48 = Mid(l0A40, 1, l0A44 - 1)
server.Text13.Text = l0A48
l0A18 = InStr(1, l0A48, ":" + Chr(9) + "/" + fnA68() + " Send List", 1)
If l0A18 Then
l0A14 = InStr(l0A48, ":")
If l0A14 = 0 Then GoTo 25
l0A1C = Mid(l0A48, 1, l0A14 - 1)
For l0A4C = 1 To server.List2.ListCount - 1
l0A50 = InStr(1, l0A1C, server.List2.List(l0A4C), 1)
If l0A50 Then
If Len(l0A1C) = Len(server.List2.List(l0A4C)) Then GoTo 25
End If
Next l0A4C
If l0A1C = fnA68() Then GoTo 25
l0A24 = fnB80(l0A1C)
End If
l0A18 = InStr(1, l0A48, ":" + Chr(9) + "/" + fnA68() + " Send ", 1)
If l0A18 Then
l0A14 = InStr(l0A48, ":")
If l0A14 = 0 Then GoTo 25
l0A1C = Mid(l0A48, 1, l0A14 - 1)
For l0A4C = 1 To server.List2.ListCount - 1
l0A50 = InStr(1, l0A1C, server.List2.List(l0A4C), 1)
If l0A50 Then
If Len(l0A1C) = Len(server.List2.List(l0A4C)) Then GoTo 25
End If
Next l0A4C
l0A20 = Mid(l0A48, l0A18 + Len(fnA68()) + 9, Len(l0A48))
l0A54 = InStr(1, l0A20, "x", 1)
If l0A54 Then GoTo 25
l0A54 = InStr(1, l0A20, "list", 1)
If l0A54 Then GoTo 25
If IsNumeric(l0A20) Then
For l0A24 = 0 To server.List1.ListCount - 1
l0A28 = InStr(1, server.List1.List(l0A24), l0A1C + "-" + l0A20, 1)
If l0A28 Then
If Len(server.List1.List(l0A24)) = Len(l0A1C + "-" + l0A20) Then GoTo 25
End If
Next l0A24
If server.Command3D3.Caption = "Start" Then GoTo LF104
If l0A1C = fnA68() Then GoTo 25
server.List1.AddItem l0A1C + "-" + l0A20
server.Label3.Caption = Trim(Str(server.List1.ListCount))
l0A58 = fnBB8(l0A1C, Int(l0A20))
End If
End If
25 :
l0A40 = Mid(l0A40, l0A44 + 2, Len(l0A40))
GoTo 3004
Else
server.Text13.Text = l0A40
If server.Command3D3.Caption = "Start" Then GoTo LF104
l0A18 = InStr(1, l0A40, ":" + Chr(9) + "/" + fnA68() + " Send List", 1)
If l0A18 Then
l0A14 = InStr(l0A40, ":")
If l0A14 = 0 Then GoTo 26
l0A1C = Mid(l0A40, 1, l0A14 - 1)
For l0A4C = 1 To server.List2.ListCount - 1
l0A50 = InStr(1, l0A1C, server.List2.List(l0A4C), 1)
If l0A50 Then
If Len(l0A1C) = Len(server.List2.List(l0A4C)) Then GoTo 26
End If
Next l0A4C
If l0A1C = fnA68() Then GoTo 26
l0A24 = fnB80(l0A1C)
End If
If server.Command3D3.Caption = "Start" Then GoTo LF104
l0A18 = InStr(1, l0A40, ":" + Chr(9) + "/" + fnA68() + " Send ", 1)
If l0A18 Then
l0A14 = InStr(l0A40, ":")
If l0A14 = 0 Then GoTo 26
l0A1C = Mid(l0A40, 1, l0A14 - 1)
For l0A4C = 1 To server.List2.ListCount - 1
l0A50 = InStr(1, l0A1C, server.List2.List(l0A4C), 1)
If l0A50 Then
If Len(l0A1C) = Len(server.List2.List(l0A4C)) Then GoTo 26
End If
Next l0A4C
If server.Command3D3.Caption = "Start" Then GoTo LF104
l0A20 = Mid(l0A40, l0A18 + Len(fnA68()) + 9, Len(l0A40))
l0A54 = InStr(1, l0A20, "x", 1)
If l0A54 Then GoTo 26
l0A54 = InStr(1, l0A20, "list", 1)
If l0A54 Then GoTo 26
If IsNumeric(l0A20) Then
For l0A24 = 0 To server.List1.ListCount - 1
l0A28 = InStr(1, server.List1.List(l0A24), l0A1C + "-" + l0A20, 1)
If l0A28 Then
If Len(server.List1.List(l0A24)) = Len(l0A1C + "-" + l0A20) Then GoTo 26
End If
Next l0A24
If server.Command3D3.Caption = "Start" Then GoTo LF104
If l0A1C = fnA68() Then GoTo 26
server.List1.AddItem l0A1C + "-" + l0A20
server.Label3.Caption = Trim(Str(server.List1.ListCount))
l0A58 = fnBB8(l0A1C, Int(l0A20))
End If
End If
26 :
End If
If server.Command3D3.Caption = "Start" Then GoTo LF104
l0A10 = server.Text13.Text
l0A5C = server.Text11.Text
server.Text12.Text = "0"
Do Until 2 > 3
l0A60 = InStr(l0A5C, l0A10)
If l0A60 Then
server.Text12.Text = Trim(Str(Int(server.Text12.Text + 1)))
l0A5C = Mid(l0A5C, l0A60 + Len(l0A10) + 2, Len(l0A5C))
Else
Exit Do
End If
Loop
l09EE = 1
If server.Command3D3.Caption = "Start" Then GoTo LF104
DoEvents
GoTo 1014

LF104:
server.Text11.Text = ""
server.Text12.Text = "0"
server.Text13.Text = ""
server.tell.Text = ""
End Function

Function fnB80 (p0A68 As Variant) As Variant
Dim l0A6C As Variant
Dim l0A70 As Variant
Dim l0A74 As Variant
Dim l0A78 As Variant
Dim l0A7C As Variant
Dim l0A80 As Variant
Dim l0A84 As Variant
Dim l0A88 As Variant
Dim l0A8C As Variant
Dim l0A90 As Variant
Dim l0A94 As Variant
Dim l0A98 As Variant
Dim l0A9C As Variant
Dim l0AA0 As Variant
Dim l0AA4 As Variant
Dim l0AA8 As Variant
Dim l0AAC As Variant
Dim l0AB0 As Variant
Dim l0AB4 As Variant
Dim l0AB8 As Variant
Dim l0ABC As Variant
Dim l0AC0 As Variant
Dim l0AC4 As Variant
Dim l0AC8 As Variant
Dim l0ACC As Variant
Dim l0AD0 As Variant
l0A6C = p0A68
l0A70 = findwindow("AOL FRAME25", 0&)
l0A74 = fn7C8("&Compose Mail")
12 :
DoEvents
l0A78 = Findchildbytitle(l0A70, "Compose Mail")
If server.Command3D3.Caption = "Start" Then GoTo LFE58
If l0A78 = 0 Then GoTo 12
l0A7C = Findchildbyclass(l0A78, "_AOL_EDIT")
l0A80 = SendMessageByString(l0A7C, 12, 0, p0A68)
l0A84 = Findchildbytitle(l0A78, "Subject:")
l0A88 = GETnextwindow(l0A84, 2)
l0A8C = SendMessageByString(l0A88, 12, 0, fn448() + "Fate X³ Warez Server List" + fn480())
l0A90 = Findchildbytitle(l0A78, "To:")
l0A94 = GETnextwindow(l0A90, 1)
l0A98 = Chr(13) + Chr(10)
l0A9C = "  _ |       |/     ¯¯/       /';| '\|   _'  | |       |  ¯¯¯¯    |       |/   |'¯¯` |;|" + l0A98 + "     |       |        '|       |; '/ ' | _'    | |       |   |¯¯¯¯| '|\       \   |[    ;|/" + l0A98 + "     |       |        '|\____\/    |\____\'\___ '\ /      '/| '|;'\____\  ¯¯¯" + l0A98 + "    /____/|        |;|'¯¯` ;|    '|;|'¯`'¯' |'|'¯¯`  / ___'/;'|  \;'|'¯`'¯` |" + l0A98 + "    |'¯`'¯`'|'|     _  \|[      | _'  \|[     ;|'|[     ]|'¯¯`' |;'/    \|[ '    ;|i88i™" + l0A98 + "    |[     ;|/        '  ¯¯¯¯        ¯¯¯¯  ¯¯¯¯|[   ' ;|/        ¯¯¯¯" + l0A98 + "     ¯¯¯¯                                             ¯¯¯" + l0A98 + l0A98 + server.Text3.Text
l0AA0 = "      ____    ____" + l0A98 + "     /       /| /        \" + l0A98 + "    '|       |;||\_____'|    " + l0A98 + "    '|       |/'|;|'¯`'¯¯`|          " + l0A98 + "   ' |       |  \|;      ;'|                       ____" + l0A98 + "    '|       |    ¯¯¯¯¯                   '  /       /|____" + l0A98 + "    '|       |   /¯¯¯¯¯\              _    |       |/        \       _____ __" + l0A98 + "   ' |\____\ '|\_____'|         ____   |       |\____ /|    '/       /|\     ' \" + l0A98 + "    ''/¯¯¯¯/| |;|'¯¯`'¯`|____ |\       \ '|       |;|'¯¯`  |'|   /       /';|;|       |" + l0A98 + " _  |       |;| \'|;    ';/  _   /' |;|       | |       |\|[     ;|/  '|       |;'/'\/___ /|" + l0A98 + l0A9C
l0AA4 = SendMessageByString(l0A94, 12, 0, l0AA0 + l0A98 + fn448() + "Fate X³ Warez Server" + fn480() + l0A98 + fn448() + "Type """ + "/" + fnA68() + " Send List" + """ for a list" + fn480() + l0A98 + fn448() + "Type """ + "/" + fnA68() + " Send X" + """ where X is the Request Index" + fn480() + l0A98 + l0A98 + server.Text5.Text)
p0A68 = Findchildbytitle(l0A78, "Send")
l0AA8 = GETnextwindow(p0A68, 2)
l0AAC = fn1870(l0AA8)
501 :
DoEvents
l0AB0 = GetParent(p0A68)
l0A70 = findwindow("AOL FRAME25", 0&)
If server.Command3D3.Caption = "Start" Then GoTo LFE58
l0AB4 = Findchildbytitle(l0A70, "Error")
If l0AB4 > 0 Then
l0AB8 = findwindow("AOL FRAME25", 0&)
l0ABC = Findchildbytitle(l0AB8, "Error")
l0AC0 = Findchildbytitle(l0ABC, "OK")
l0AC4 = fn1870(l0AC0)
l0AC8 = Findchildbytitle(l0AB8, "Compose Mail")
l0ACC = SendMessageByNum(l0AC8, 16, 0, 0)
GoTo LFE58
End If
If l0AB0 > 0 Then GoTo 501
l0AD0 = Timeout(.5)
l0AD0 = fnAD8(fn448() + "Fate X³ <" + l0A6C + "> - List was sent" + fn480())
LFE58:
End Function

Function fnBB8 (p0AD8 As Variant, p0ADC As Integer) As Variant
Dim l0ADE As Variant
Dim l0AE2 As Variant
Dim l0AE6 As String
Dim l0AE8 As Variant
Dim l0AEC As Variant
Dim l0AF0 As Variant
Dim l0AF4 As Variant
Dim l0AF8 As Variant
Dim l0AFC As Variant
Dim l0B00 As Variant
Dim l0B04 As Variant
Dim l0B08 As Variant
Dim l0B0C As Variant
Dim l0B10 As Variant
Dim l0B14 As Variant
Dim l0B18 As Variant
Dim l0B1C As Variant
Dim l0B20 As Variant
Dim l0B24 As Variant
Dim l0B28 As Variant
Dim l0B2C As String
Dim l0B2E As Variant
Dim l0B32 As Variant
Dim l0B36 As Variant
Dim l0B3A As Variant
Dim l0B3E As Variant
Dim l0B42 As Variant
Dim l0B46 As Variant
Dim l0B4A As Variant
Dim l0B4E As Variant
Dim l0B52 As Variant
Dim l0B56 As Variant
Dim l0B5A As Variant
Dim l0B5E As Variant
Dim l0B62 As Variant
Dim l0B66 As Variant
Dim l0B6A As Variant
Dim l0B6E As Variant
Dim l0B72 As Variant
Dim l0B76 As Variant
Dim l0B7A As Variant
l0ADE = Chr(13) + Chr(10)
l0AE2 = findwindow("AOL FRAME25", 0&)
If server.Option3D1.Value = True Then l0AE6$ = "New Mail"
If server.Option3D2.Value = True Then l0AE6$ = "Old Mail"
If server.Option3D3.Value = True Then l0AE6$ = "Incoming FlashMail"
l0AE8 = fn528("Delete", l0AE6$)
l0AEC = Findchildbyclass(l0AE8, "_AOL_TREE")
l0AF0 = SendMessageByNum(l0AEC, 1036, 0, 0)
If server.Option3D3.Value = True Then
l0AF4 = SendMessageByNum(l0AEC, 1031, p0ADC, 0)
l0AF8 = Findchildbytitle(l0AE8, "Open")
l0AFC = fn1870(l0AF8)
GoTo 340
End If
l0AF4 = SendMessageByNum(l0AEC, 1031, p0ADC, 0)
l0AF8 = Findchildbytitle(l0AE8, "Read")
l0AFC = fn1870(l0AF8)
340 :
150 :
DoEvents
l0B00 = Findchildbytitle(l0AE2, "Forward")
If server.Command3D3.Caption = "Start" Then GoTo L10D84
If l0B00 = 0 Then GoTo 150
151 :
DoEvents
l0B04 = GETnextwindow(l0B00, 2)
l0B08 = fn1870(l0B04)
l0B0C = Timer
While Timer - l0B0C < .7
l0B10 = Findchildbytitle(l0AE2, "Address")
If l0B10 > 0 Then GoTo 162
DoEvents
Wend
l0B14 = Findchildbytitle(l0AE2, "Send Now")
l0B10 = Findchildbytitle(l0AE2, "Address")
If l0B10 = 0 Then GoTo 151
162 :
If server.Command3D3.Caption = "Start" Then GoTo L10D84
l0B14 = Findchildbytitle(l0AE2, "Send Now")
l0B18 = GetParent(l0B10)
l0B1C = Findchildbyclass(l0B18, "_AOL_EDIT")
l0B20 = SendMessageByString(l0B1C, 12, 0, p0AD8)
l0B24 = Findchildbytitle(l0B18, "Subject:")
l0B28 = GETnextwindow(l0B24, 2)
l0B2C$ = String(255, 0)
l0B2E = SendMessageByString(l0B28, 13, 255, l0B2C$)
l0B32 = Mid(l0B2C$, 6, Len(l0B2C$))
l0B36 = SendMessageByString(l0B28, 12, 0, l0B32)
l0B3A = Findchildbytitle(l0B18, "To:")
l0B3E = GETnextwindow(l0B3A, 1)
l0B42 = GETnextwindow(l0B3E, 3)
l0ADE = Chr(13) + Chr(10)
l0B46 = "  _ |       |/     ¯¯/       /';| '\|   _'  | |       |  ¯¯¯¯    |       |/   |'¯¯` |;|" + l0ADE + "     |       |        '|       |; '/ ' | _'    | |       |   |¯¯¯¯| '|\       \   |[    ;|/" + l0ADE + "     |       |        '|\____\/    |\____\'\___ '\ /      '/| '|;'\____\  ¯¯¯" + l0ADE + "    /____/|        |;|'¯¯` ;|    '|;|'¯`'¯' |'|'¯¯`  / ___'/;'|  \;'|'¯`'¯` |" + l0ADE + "    |'¯`'¯`'|'|     _  \|[      | _'  \|[     ;|'|[     ]|'¯¯`' |;'/    \|[ '    ;|i88i™" + l0ADE + "    |[     ;|/        '  ¯¯¯¯        ¯¯¯¯  ¯¯¯¯|[   ' ;|/        ¯¯¯¯" + l0ADE + "     ¯¯¯¯                                             ¯¯¯" + l0ADE + l0ADE + server.Text3.Text
l0B4A = "      ____    ____" + l0ADE + "     /       /| /        \" + l0ADE + "    '|       |;||\_____'|    " + fn448() + "Fate X³ Warez Server By MaGuS" + fn480() + l0ADE + "    '|       |/'|;|'¯`'¯¯`|          " + fn448() + "Requested Index#: " + Trim(Str(p0ADC)) + fn480() + l0ADE + "   ' |       |  \|;      ;'|                       ____" + l0ADE + "    '|       |    ¯¯¯¯¯                   '  /       /|____" + l0ADE + "    '|       |   /¯¯¯¯¯\              _    |       |/        \       _____ __" + l0ADE + "   ' |\____\ '|\_____'|         ____   |       |\____ /|    '/       /|\     ' \" + l0ADE + "    ''/¯¯¯¯/| |;|'¯¯`'¯`|____ |\       \ '|       |;|'¯¯`  |'|   /       /';|;|       |" + l0ADE + " _  |       |;| \'|;    ';/  _   /' |;|       | |       |\|[     ;|/  '|       |;'/'\/___ /|" + l0ADE + l0B46
l0B4E = SendMessageByString(l0B42, 12, 0, l0B4A)
l0B52 = GETnextwindow(l0B14, 2)
l0B56 = fn1870(l0B52)
152 :
DoEvents
If server.Command3D3.Caption = "Start" Then GoTo L10D84
l0B5A = GetParent(l0B24)
l0B5E = Findchildbytitle(l0AE2, "Error")
If l0B5E > 0 Then
l0B62 = Findchildbytitle(l0B5E, "OK")
l0B66 = fn1870(l0B62)
l0B6A = SendMessageByNum(l0B18, 16, 0, 0)
l0B6E = SendMessageByNum(GetParent(l0B00), 16, 0, 0)
GoTo L10D84
End If
If l0B5A > 0 Then GoTo 152
l0B72 = SendMessageByNum(GetParent(l0B00), 16, 0, 0)
l0B76 = Findchildbytitle(l0AE8, "Keep As New")
l0B7A = fn1870(l0B76)
L10D84:
End Function

Function fnBF0 () As Variant
Dim l0B82 As Variant
Dim l0B8A As Variant
Dim l0B8E As Variant
Dim l0B92 As Variant
Dim l0B96 As Variant
Dim l0B9A As Variant
Dim l0B9E As Variant
Dim l0BA2 As Variant
Dim l0BA6 As Variant
Dim l0BAA As Variant
Do Until 2 > 3
For l0B82 = 0 To smartbot.List1.ListCount - 1
l0B8A = smartbot.List1.List(l0B82)
If l0B82 = 0 Then
l0B8E = l0B8E + "," + l0B8A + ","
Else
l0B8E = l0B8E + l0B8A + ","
End If
If fn3D8() <> smartbot.Text2.Text Then GoTo 2899
Next l0B82
l0B92 = fn2F8()
l0B96 = Findchildbyclass(l0B92, "_AOL_LISTBOX")
l0B9A = SendMessageByNum(l0B96, 1036, 0, 0)
For l0B82 = 0 To l0B9A - 1
l0B8A = vbgetaollist(l0B82, l0B96)
If l0B82 = 0 Then
l0B9E = l0B9E + "," + l0B8A + ","
Else
l0B9E = l0B9E + l0B8A + ","
End If
l0BA2 = InStr(l0B8E, "," + l0B8A + ",")
If l0BA2 = 0 Then
If fn3D8() <> smartbot.Text2.Text Then GoTo 2899
If l0B8A = "" Then GoTo 128
If smartbot.Timer1.Interval = 0 Then Exit Function
smartbot.List2.AddItem l0B8A
l0BA6 = Timeout(.4)
If smartbot.Caption = "Greeter Bot" Then
If smartbot.Command3D1.Caption = "Start" Then GoTo L1133A
l0BAA = fnAA0(fn448() + l0B8A + " - Has Entered The Room " + fn480())
Else
l0BA6 = InStr(1, l0B8A, "Catwatch", 1)
If l0BA6 Then
l0BA6 = SendMessageByNum(fn2F8(), 16, 0, 0)
MsgBox "Catwatch has been detected and chat room was closed!", 16
End If
End If
End If
128 :
If fn3D8() <> smartbot.Text2.Text Then GoTo 2899
Next l0B82
If l0B9E <> l0B8E Then
2899 :
smartbot.List1.Clear
For l0B82 = 0 To l0B9A - 1
l0B8A = vbgetaollist(l0B82, l0B96)
If smartbot.Timer1.Interval = 0 Then Exit Function
smartbot.List1.AddItem l0B8A
Next l0B82
End If
smartbot.Text2.Text = fn3D8()
If smartbot.Command3D1.Caption = "Start" Then GoTo L1133A
DoEvents
l0B8E = ""
l0B9E = ""
If fn3D8() <> smartbot.Text2.Text Then GoTo 2899
Loop
L1133A:
End Function

Function fnC28 (p0BB2 As Variant) As Variant
Dim l0BB6 As Variant
l0BB6 = InStr(p0BB2, Chr(0))
If l0BB6 Then
fnC28 = Mid(p0BB2, 1, l0BB6 - 1)
Else
fnC28 = p0BB2
End If
End Function

Function fnC60 () As Variant
End Function

Function fnC8 () As Variant
Dim l00C0 As Variant
Dim l00C4 As Variant
Dim l00C8 As Variant
Dim l00D0 As String
Dim l00D2 As Variant
Dim l00D6 As Variant
Dim l00DA As Variant
Dim l00DE As Variant
Dim l00E2 As Variant
Dim l00E6 As Variant
22 :
DoEvents
l00C0 = findwindow("AOL FRAME25", 0&)
l00C4 = fn2F8()
l00C0 = l00C4
l00C4 = Findchildbyclass(l00C0, "_AOL_VIEW")
If l00C4 = 0 Then
Exit Function
End If
l00C8 = SendMessageByNum(l00C4, 14, 0, 0)
If l00C8 = Len(mmER.Text8.Text) Then GoTo 22
l00D0$ = String(l00C8, 0)
l00D2 = SendMessageByString(l00C4, 13, l00C8, l00D0$)
l00D6 = Len(mmER.Text8.Text)
l00DA = Len(l00D0$) - l00D6
On Error Resume Next
l00DE = Right(l00D0$, l00DA)
l00DE = Mid(l00DE, 3, Len(l00DE))
24 :
l00E2 = InStr(l00DE, Chr(10))
If l00E2 Then
l00E6 = Mid(l00DE, 1, l00E2 - 2)
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
mmER.Text9.Text = l00E6
l00DE = Mid(l00DE, l00E2 + 1, Len(l00DE))
GoTo 24
End If
If mmER.Text9.Text <> "" Then mmER.Text10.Text = mmER.Text9.Text
mmER.Text9.Text = l00DE
mmER.Text8.Text = l00D0$
DoEvents
GoTo 22
End Function

Function fnCD0 (p0BF0 As Variant) As Variant
Dim l0BF4 As Variant
Dim l0BF8 As Variant
Dim l0BFC As Variant
Do Until 2 > 3
l0BF4 = InStr(p0BF0, Chr(32))
If l0BF4 Then
l0BF8 = Mid(p0BF0, 1, l0BF4 - 1)
l0BFC = Mid(p0BF0, l0BF4 + 1, Len(p0BF0))
p0BF0 = l0BF8 + l0BFC
Else
GoTo 70
End If
Loop
70 :
p0BF0 = UCase(p0BF0)
fnCD0 = p0BF0
End Function

Function fnD08 () As Variant
Dim l0C04 As Variant
Open "fatex30.exe" For Input As #1
l0C04 = LOF(1)
If l0C04 <> 2161069 Then
MsgBox "This program has been modified!", 16
End
End If
Close #1
End Function

Function fnD40 (p0C0C As Variant) As Variant
Dim l0C10 As Variant
Dim l0C14 As Variant
Dim l0C18 As Variant
Dim l0C1C As Variant
Dim l0C20 As Variant
Dim l0C24 As String
Dim l0C26 As Variant
Dim l0C2A As Variant
Dim l0C2E As Variant
Dim l0C32 As Variant
Dim l0C36 As Variant
Dim l0C3A As Variant
Dim l0C3E As Variant
15 :
DoEvents
l0C10 = findwindow("AOL FRAME25", 0&)
l0C10 = Findchildbytitle(l0C10, p0C0C)
l0C14 = Findchildbyclass(l0C10, "_AOL_TREE")
l0C18 = findwindow("#32770", "America Online")
l0C1C = Findchildbytitle(l0C18, "OK")
l0C20 = GETnextwindow(l0C1C, 1)
l0C24$ = String(255, 0)
l0C26 = SendMessageByString(l0C20, 13, 255, l0C24$)
l0C2A = fnC28(l0C24$)
l0C2E = Mid(l0C24$, 1, 11)
If l0C2E = "You have no" Then
l0C32 = fn1870(l0C1C)
MsgBox "You have no " + p0C0C + "!", 16
fnD40 = 1
Exit Function
End If
If l0C14 = 0 Then GoTo 15
Do Until 2 > 3
pwait.Label1.Caption = "Waiting for " + p0C0C + " to load up..."
l0C36 = SendMessageByNum(l0C14, 1036, 0, 0)
l0C3A = Timeout(.6)
l0C3E = SendMessageByNum(l0C14, 1036, 0, 0)
If l0C36 = l0C3E Then Exit Do
Loop
pwait.Label1.Caption = "Please Wait..."
End Function

Function fnD78 () As Variant
Dim l0C46 As String
Dim l0C48 As Variant
l0C46$ = String(255, 0)
l0C48 = getwindowsdirectory(l0C46$, 255)
fnD78 = fnC28(l0C46$)
End Function

Sub FormCenterLeft (f As Form)
f.Left = Screen.Width / 2 - f.Width / 2
End Sub

Sub FormCenterTop (f As Form)
f.Top = Screen.Height / 2 - f.Height / 2
End Sub

Function ZSelfCheck (ByVal p0BC2 As String) As Variant
Dim l0BC8 As Variant
Dim l0BCC As String
Dim l0BCE As Variant
Dim l0BD2 As String
Dim l0BD4 As Variant
Dim l0BD8 As Variant
Dim l0BDC As Variant
Dim l0BE0 As Variant
Dim l0BE4 As Variant
Dim l0BE8 As Variant
SelfCheck.Show
l0BC8 = 0
l0BCC$ = p0BC2
l0BCE = 1
Do Until 2 > 3
DoEvents
l0BD2$ = ""
DoEvents
On Error Resume Next
Open "fatex30.exe" For Binary As #1
If Err Then
MsgBox "Either that file doesn't exist, or an unexpected error occured!", 16
Exit Function
End If
If l0BCE = 1 Then
l0BD4 = LOF(1)
l0BD8 = l0BD4 / 32000
l0BD8 = 3068 / l0BD8
l0BDC = l0BD8
'selfcheck.Picture1.Line (x,y)(0, 0)-(l0BDC, 500), QBColor(1), BF
Else
l0BDC = l0BDC + l0BD8
'selfcheck.Picture1.Line (x,y)(0, 0)-(l0BDC, 500), QBColor(1), BF
End If
l0BD2$ = String(32000, 0)
Get #1, l0BCE, l0BD2$
Close #1
Open "fatex30.exe" For Binary As #2
l0BE0 = InStr(1, l0BD2$, l0BCC$, 1)
If l0BE0 Then
Mid(l0BD2$, l0BE0) = String(Len(p0BC2), 1)
l0BC8 = l0BC8 + 1
40 :
DoEvents
l0BE4 = InStr(1, l0BD2$, l0BCC$, 1)
If l0BE4 Then
Mid(l0BD2$, l0BE4) = String(Len(p0BC2), 1)
l0BC8 = l0BC8 + 1
GoTo 40
End If
End If
l0BCE = l0BCE + 32000
l0BE8 = LOF(2)
Close #2
If l0BCE > l0BE8 Then GoTo 30
Loop
30 :
Unload SelfCheck
SelfCheck.Picture1.Cls
If l0BC8 <> 117 Then
MsgBox "The Fate that you have right now has been edited in some way. Goto: http://www.geocities.com/Area51/Corridor/8621 to download a SAFE version of Fate."
Else
MsgBox "The Fate that you have is completely safe."
End If
End Function

