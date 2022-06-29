' Module2
Option Explicit
Declare Sub extsub1F38 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub1F70 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Function extfn22F0 Lib "APIGuide.Dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfn2328 Lib "APIGuide.Dll" Alias "agDeviceMode" () As Integer
Declare Sub extsub1FA8 Lib "APIGuide.Dll" Alias "agDWordTo2Integers" ()
Declare Function extfn2360 Lib "APIGuide.Dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn20F8 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn2130 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn2168 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn20C0 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn21A0 Lib "APIGuide.Dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn2050 Lib "APIGuide.Dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn21D8 Lib "APIGuide.Dll" Alias "agGetControlName" () As String
Declare Function extfn2088 Lib "APIGuide.Dll" Alias "agGetInstance" () As Integer
Declare Function extfn1CD0 Lib "APIGuide.Dll" Alias "AgGetStringFromLPSTR" () As String
Declare Function extfn2408 Lib "APIGuide.Dll" Alias "agHugeOffset" () As Long
Declare Function extfn2398 Lib "APIGuide.Dll" Alias "agInp" () As Integer
Declare Function extfn23D0 Lib "APIGuide.Dll" Alias "agInpw" () As Integer
Declare Sub extsub1FE0 Lib "APIGuide.Dll" Alias "agOutp" ()
Declare Sub extsub2018 Lib "APIGuide.Dll" Alias "agOutpw" ()
Declare Function extfn2440 Lib "APIGuide.Dll" Alias "agVBGetVersion" () As Integer
Declare Function extfn2478 Lib "APIGuide.Dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfn24B0 Lib "APIGuide.Dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn2210 Lib "APIGuide.Dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfn2280 Lib "APIGuide.Dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfn2248 Lib "APIGuide.Dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfn22B8 Lib "APIGuide.Dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfn30B8 Lib "311.Dll" Alias "AOLGetcombo" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3010 Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1448 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn1480 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn18E0 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfn1E90 Lib "GDI" Alias "BitBlt" () As Integer
Declare Sub extsub10C8 Lib "User" Alias "clienttoscreen" ()
Declare Sub extsub2C90 Lib "User" Alias "closewindow" ()
Declare Function extfn1410 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn2B40 Lib "GDI" Alias "CreateSolidBrush" (ByVal p1&) As Integer
Declare Function extfn2F68 Lib "bubble.dll" Alias "DeleteBubble" () As Integer
Declare Sub extsub1D78 Lib "GDI" Alias "deleteobject" (ByVal p1%)
Declare Function extfn2C58 Lib "User" Alias "destroywindow" () As Integer
Declare Sub extsub11E0 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn24E8 Lib "APIGuide.Dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn1170 Lib "User" Alias "enablewindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn19F8 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfn13A0 Lib "User" Alias "exitwindow" () As Integer
Declare Function extfn1F00 Lib "vbwfind.dll" Alias "findchild" () As Integer
Declare Function extfn2FD8 Lib "vbwfind.dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn2FA0 Lib "vbwfind.dll" Alias "findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3048 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn1330 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn1368 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn1E20 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn1640 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn17C8 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1288 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn2C20 Lib "User" Alias "getcurrenttime" () As Long
Declare Function extfn1790 Lib "User" Alias "getcursor" () As Integer
Declare Function extfn2B78 Lib "User" Alias "getdc" (ByVal p1%) As Integer
Declare Function extfn1988 Lib "User" Alias "GetDesktopWindow" () As Integer
Declare Function extfn1DB0 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn2BB0 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn1B10 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn1250 Lib "User" Alias "GetFreeSystemResources" () As Integer
Declare Function extfn2E18 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn2CC8 Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn2D70 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1918 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn2D00 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn13D8 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn2980 Lib "311.Dll" Alias "AOLGetList" () As Integer
Declare Function extfn1838 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn12F8 Lib "User" Alias "getnextwindow" () As Integer
Declare Function extfn3080 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1C98 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn1BB8 Lib "Kernel" Alias "GetPrivateProfileString" () As Integer
Declare Function extfn1BF0 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn1C28 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn1758 Lib "User" Alias "GetScrollPos" () As Integer
Declare Sub extsub2BE8 Lib "User" Alias "GetScrollRange" (ByVal p1%, ByVal p2%, p3%, p4%)
Declare Function extfn2DA8 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn16B0 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn1B48 Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn1800 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn1950 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn18A8 Lib "User" Alias "Gettopwindow" () As Integer
Declare Function extfn1AD8 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn2E88 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2AD0 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn1A68 Lib "Kernel" Alias "GetWindowDirectory" () As Integer
Declare Sub extsub2A60 Lib "User" Alias "GetWindowRect" (ByVal p1%, p2 As Any)
Declare Function extfn1598 Lib "User" Alias "getwindowtext" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1870 Lib "User" Alias "GetWindowTextLength" () As Integer
Declare Function extfn15D0 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1AA0 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn14B8 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn1720 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfn1A30 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn1EC8 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Sub extsub1090 Lib "User" Alias "movewindow" ()
Declare Function extfn2868 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Sub extsub2558 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Function extfn2910 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfn2948 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfn26A8 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn26E0 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn2670 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2718 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn2638 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2750 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn2600 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Sub extsub2520 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Function extfn25C8 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn27F8 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfn27C0 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfn2788 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfn2830 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfn28A0 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfn28D8 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Sub extsub2590 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Declare Sub extsub1D40 Lib "GDI" Alias "Rectangle" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%)
Declare Sub extsub3248 Lib "User" Alias "releaseCapture" ()
Declare Function extfn2A98 Lib "User" Alias "ReleaseDC" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1058 Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn2B08 Lib "GDI" Alias "SelectObject" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2EC0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4&) As Long
Declare Function extfn2EF8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn2F30 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn1678 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Sub extsub1D08 Lib "GDI" Alias "SetBkColor" ()
Declare Sub extsub1100 Lib "User" Alias "setcapture" ()
Declare Function extfn12C0 Lib "User" Alias "SetClassWord" () As Integer
Declare Sub extsub1218 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn2D38 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn1138 Lib "User" Alias "setparent" () As Integer
Declare Function extfn16E8 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn1E58 Lib "GDI" Alias "settextcolor" () As Long
Declare Function extfn2E50 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfn1608 Lib "User" Alias "SetWindowText" () As Integer
Declare Sub extsub11A8 Lib "User" Alias "ShowOwnedPopups" ()
Declare Function extfnFE8 Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1020 Lib "MMSystem" Alias "sndPlaySound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn19C0 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn2DE0 Lib "APIGuide.Dll" Alias "AgGetStringFromLPSTR" () As String
Declare Function extfn1DE8 Lib "GDI" Alias "TextOut" () As Integer
Declare Sub extsub2A28 Lib "User" Alias "UpdateWindow" ()
Declare Function extfn29B8 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfn29F0 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn14F0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn1560 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn1528 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn1B80 Lib "Kernel" Alias "WritePrivateProfileString" () As Integer
Declare Function extfn1C60 Lib "Kernel" Alias "WriteProfileString" () As Integer
Const c0768 = 150 ' &H96%

Sub sub1B0 (p05BA As ListBox)
Dim l05BC As Variant
Dim l05C0 As Integer
Dim l05C2 As String
Dim l05C4 As Variant
Dim l05C8 As Integer
Dim l05CA As String
Dim l05CC As String
Dim l05CE As Variant
l05BC = extfn3048%("AOL Frame25", "America  Online")
l05C0% = extfn2FA0%(l05BC, "Welcome,")
l05C2$ = fn6F0((l05C0%))
If  InStr(l05C2$, "!") = 0 Then Exit Sub
l05C2$ = Left$(l05C2$, InStr(l05C2$, "!") - 1)
l05C4 = Right$(l05C2$, Len(l05C2$) - InStr(l05C2$, ",") - 1)
For l05C8% = 0 To 25
l05CA$ = String$(256, " ")
l05CE = extfn3010(l05C8%, l05CA$) & l05CC$
If  Len(Trim$(l05CA$)) <= 1 Then Exit For
l05CA$ = Left$(Trim$(l05CA$), Len(Trim(l05CA$)) - 1)
If  l05CA$ = l05C4 Then GoTo L34FA
p05BA.AddItem (l05CA$)
L34FA:
Next l05C8%
End Sub

Sub sub1E8 (p05D4 As ListBox)
Dim l05D6 As Variant
Dim l05DA As Integer
Dim l05DC As String
Dim l05DE As Variant
Dim l05E2 As Integer
Dim l05E4 As String
Dim l05E6 As String
Dim l05E8 As Variant
l05D6 = extfn3048%("AOL Frame25", "America  Online")
l05DA% = extfn2FA0%(l05D6, "Welcome,")
l05DC$ = fn6F0((l05DA%))
If  InStr(l05DC$, "!") = 0 Then Exit Sub
l05DC$ = Left$(l05DC$, InStr(l05DC$, "!") - 1)
l05DE = Right$(l05DC$, Len(l05DC$) - InStr(l05DC$, ",") - 1)
For l05E2% = 0 To 25
l05E4$ = String$(256, " ")
l05E8 = extfn3010(l05E2%, l05E4$) & l05E6$
If  Len(Trim$(l05E4$)) <= 1 Then Exit For
l05E4$ = Left$(Trim$(l05E4$), Len(Trim(l05E4$)) - 1)
If  l05E4$ = l05DE Then GoTo L3778
If  InStr(LCase$(l05E4$), "tos") Then GoTo L3778
If  InStr(LCase$(l05E4$), "guide") Then GoTo L3778
If  InStr(LCase$(l05E4$), "host") Then GoTo L3778
p05D4.AddItem (l05E4$)
L3778:
Next l05E2%
End Sub


Function fn220 (p05F0 As String) As Variant
Dim l05F2 As String
Dim l05F4 As String
Dim l05F6 As String
Do
l05F2$ = Right$(p05F0$, 1)
l05F4$ = l05F4$ & l05F2$
l05F6$ = Left$(p05F0$, Len(p05F0$) - 1)
p05F0$ = l05F6$
DoEvents
Loop Until Len(p05F0$) = 0
fn220 = l05F4$
End Function

Function fn258 (p05FC As Variant, p0600 As Variant) As Variant
Dim l0604 As Variant
l0604 = Mid(p05FC, Val(p0600))
fn258 = Left(l0604, 1)
End Function

Function fn290 () As Variant
Dim l060C As Variant
Dim l0610 As Variant
Dim l0614 As Variant
Dim l0618 As Integer
Dim l061A As Variant
Dim l061E As Variant
l060C = extfn3048("AOL Frame25", 0&)
l0610 = extfn2FA0(l060C, "List Rooms")
l0614 = extfn3080(l0610)
l0618% = extfn2FD8(l0614, "_AOL_View")
l061A = fn6F0(l0618%)
Do
l060C = extfn3048("AOL Frame25", 0&)
l0610 = extfn2FA0(l060C, "List Rooms")
l0614 = extfn3080(l0610)
l0618% = extfn2FD8(l0614, "_AOL_View")
l061E = fn6F0(l0618%)
DoEvents
Loop Until l061E <> l061A
fn290 = Mid(l061E, Len(l061A) + 3)
End Function

Function fn2C8 (p0626 As String) As Variant
Dim l0628 As Integer
Dim l062A As Variant
Dim l062E As Variant
Dim l0632 As Integer
Dim l0634 As Integer
Dim l0636 As Variant
Dim l063A As Integer
Dim l063C As Integer
Dim l063E As Integer
Dim l0640 As Integer
Dim l0642 As Integer
Dim l0644 As Integer
Dim l0646 As Integer
Dim l0648 As Integer
Dim l064A As Integer
Dim l064C As Integer
Dim l064E As Integer
Dim l0650 As Integer
Dim l0652 As Integer
Dim l0654 As Integer
Dim l0656 As Variant
Dim l065A As Integer
Dim l065C As Variant
Dim l0660 As Integer
Dim l0662 As Variant
Do
l0628% = extfn3048("AOL Frame25", 0&)
l062A = extfn2FA0(l0628%, "Compose Mail")
subC30 l062A, "Compose Mail"
DoEvents
Loop Until l062A = 0
l0628% = extfn3048("AOL Frame25", 0&)
subBC0 "Compose"
Do
l062E = DoEvents()
l0632% = extfn2FA0(l0628%, "Compose Mail")
l0634% = extfn2FD8(l0632%, "_AOL_Edit")
l0636 = extfnFE8(l0632%, gv0066)
Loop Until l0634% <> 0
l063A% = extfn3080(l0632%)
l063C% = extfn2FD8(l0632%, "_AOL_Icon")
l063E% = extfn2FD8(l0632%, "_AOL_View")
l0634% = extfn2FD8(l0632%, "_AOL_Edit")
l0640% = extfn2F30(l0634%, gc05E0, 0, p0626$ & ", '")
l0642% = extfn2E88(l0634%, gc05D6)
l0644% = extfn2E88(l0642%, gc05D6)
l0646% = extfn2E88(l0644%, gc05D6)
l0648% = extfn2E88(l0646%, gc05D6)
l0640% = extfn2F30(l0648%, gc05E0, 0, "TOS check")
l064A% = extfn2E88(l0648%, gc05D6)
l064C% = extfn2E88(l064A%, gc05D6)
l064E% = extfn2E88(l064C%, gc05D6)
l0650% = extfn2E88(l064E%, gc05D6)
l0652% = extfn2E88(l0650%, gc05D6)
l0654% = extfn2E88(l0652%, gc05D6)
l0640% = extfn2F30(l064E%, gc05E0, 0, "TOS Check")
l0656 = extfn2FD8(l0632%, "RICHCNTL")
l0640% = extfn2F30(l0656, gc05E0, 0, "TOS Check")
subDB8 (.1#)
l065A% = extfn2EF8(l063C%, gc05CE, &HD, 0)
l065A% = extfn2EF8(l063C%, gc05D0, &HD, 0)
subDB8 (.1#)
Do
l065C = extfn2FA0(l0628%, "Error")
subDB8 (.001#)
Loop Until l065C <> 0
l0660% = extfn2FD8(l065C, "_AOL_View")
l0662 = fn6F0(l0660%)
If  InStr(LCase$(l0662), LCase$(p0626$)) Then
MsgBox "Account is dead"
fn2C8 = 1
End If
If  InStr(LCase$(l0662), LCase$("mailbox")) Then
MsgBox "Account is alive"
fn2C8 = 0
End If
sub958 (l065C)
sub958 (l0632%)
Do
l0628% = extfn3048("AOL Frame25", 0&)
l062A = extfn2FA0(l0628%, "Compose Mail")
subC30 l062A, "Compose Mail"
DoEvents
Loop Until l062A = 0
End Function

Function fn300 (p066A As String) As Variant
Dim l066C As String
Dim l066E As String
Dim l0670 As String
Dim l0672 As String
l066C$ = LCase$(p066A$)
Do
l066E$ = Left$(l066C$, 1)
l0670$ = l066E$
If  l066E$ = "a" Then l0670$ = "z"
If  l066E$ = "b" Then l0670$ = "a"
If  l066E$ = "c" Then l0670$ = "b"
If  l066E$ = "d" Then l0670$ = "c"
If  l066E$ = "e" Then l0670$ = "d"
If  l066E$ = "f" Then l0670$ = "e"
If  l066E$ = "g" Then l0670$ = "f"
If  l066E$ = "h" Then l0670$ = "g"
If  l066E$ = "i" Then l0670$ = "h"
If  l066E$ = "j" Then l0670$ = "i"
If  l066E$ = "k" Then l0670$ = "j"
If  l066E$ = "l" Then l0670$ = "k"
If  l066E$ = "m" Then l0670$ = "l"
If  l066E$ = "n" Then l0670$ = "m"
If  l066E$ = "o" Then l0670$ = "n"
If  l066E$ = "p" Then l0670$ = "o"
If  l066E$ = "q" Then l0670$ = "p"
If  l066E$ = "r" Then l0670$ = "q"
If  l066E$ = "s" Then l0670$ = "r"
If  l066E$ = "t" Then l0670$ = "s"
If  l066E$ = "u" Then l0670$ = "t"
If  l066E$ = "v" Then l0670$ = "u"
If  l066E$ = "w" Then l0670$ = "v"
If  l066E$ = "x" Then l0670$ = "w"
If  l066E$ = "y" Then l0670$ = "x"
If  l066E$ = "z" Then l0670$ = "y"
If  l066E$ = "0" Then l0670$ = "9"
If  l066E$ = "1" Then l0670$ = "0"
If  l066E$ = "2" Then l0670$ = "1"
If  l066E$ = "3" Then l0670$ = "2"
If  l066E$ = "4" Then l0670$ = "3"
If  l066E$ = "5" Then l0670$ = "4"
If  l066E$ = "6" Then l0670$ = "5"
If  l066E$ = "7" Then l0670$ = "6"
If  l066E$ = "8" Then l0670$ = "7"
If  l066E$ = "9" Then l0670$ = "8"
l0672$ = l0672$ & l0670$
l066C$ = Mid$(l066C$, 2)
DoEvents
Loop Until Len(l066C$) = 0
fn300 = l0672$
End Function

Function fn338 (p0678 As String) As Variant
Dim l067A As String
Dim l067C As String
Dim l067E As String
Dim l0680 As String
l067A$ = LCase$(p0678$)
Do
l067C$ = Left$(l067A$, 1)
l067E$ = l067C$
If  l067C$ = "a" Then l067E$ = "b"
If  l067C$ = "b" Then l067E$ = "c"
If  l067C$ = "c" Then l067E$ = "d"
If  l067C$ = "d" Then l067E$ = "e"
If  l067C$ = "e" Then l067E$ = "f"
If  l067C$ = "f" Then l067E$ = "g"
If  l067C$ = "g" Then l067E$ = "h"
If  l067C$ = "h" Then l067E$ = "i"
If  l067C$ = "i" Then l067E$ = "j"
If  l067C$ = "j" Then l067E$ = "k"
If  l067C$ = "k" Then l067E$ = "l"
If  l067C$ = "l" Then l067E$ = "m"
If  l067C$ = "m" Then l067E$ = "n"
If  l067C$ = "n" Then l067E$ = "o"
If  l067C$ = "o" Then l067E$ = "p"
If  l067C$ = "p" Then l067E$ = "q"
If  l067C$ = "q" Then l067E$ = "r"
If  l067C$ = "r" Then l067E$ = "s"
If  l067C$ = "s" Then l067E$ = "t"
If  l067C$ = "t" Then l067E$ = "u"
If  l067C$ = "u" Then l067E$ = "v"
If  l067C$ = "v" Then l067E$ = "w"
If  l067C$ = "w" Then l067E$ = "x"
If  l067C$ = "x" Then l067E$ = "y"
If  l067C$ = "y" Then l067E$ = "z"
If  l067C$ = "z" Then l067E$ = "a"
If  l067C$ = "0" Then l067E$ = "1"
If  l067C$ = "1" Then l067E$ = "2"
If  l067C$ = "2" Then l067E$ = "3"
If  l067C$ = "3" Then l067E$ = "4"
If  l067C$ = "4" Then l067E$ = "5"
If  l067C$ = "5" Then l067E$ = "6"
If  l067C$ = "6" Then l067E$ = "7"
If  l067C$ = "7" Then l067E$ = "8"
If  l067C$ = "8" Then l067E$ = "9"
If  l067C$ = "9" Then l067E$ = "0"
l0680$ = l0680$ & l067E$
l067A$ = Mid$(l067A$, 2)
DoEvents
Loop Until Len(l067A$) = 0
fn338 = l0680$
End Function

Sub sub370 (p0682 As Integer)
Dim l0684 As Integer
l0684% = extfn2EF8(p0682%, gc05CE, &HD, 0)
l0684% = extfn2EF8(p0682%, gc05D0, &HD, 0)
End Sub

Function fn3A8 (p068A As Variant) As Variant
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
Do
l068E = p068A
If  Len(l068E) > 16 Then Exit Function
Do Until Len(l068E) = 16
Randomize Timer
l068E = l068E & Int(10 * Rnd)
Loop
l0692 = fn258(l068E, 1)
l0696 = fn258(l068E, 2)
l069A = fn258(l068E, 3)
l069E = fn258(l068E, 4)
l06A2 = fn258(l068E, 5)
l06A6 = fn258(l068E, 6)
l06AA = fn258(l068E, 7)
l06AE = fn258(l068E, 8)
l06B2 = fn258(l068E, 9)
l06B6 = fn258(l068E, 10)
l06BA = fn258(l068E, 11)
l06BE = fn258(l068E, 12)
l06C2 = fn258(l068E, 13)
l06C6 = fn258(l068E, 14)
l06CA = fn258(l068E, 15)
l06CE = fn258(l068E, 16)
l06D2 = Val(l0696) + Val(l069E) + Val(l06A6) + Val(l06AE) + Val(l06B6) + Val(l06BE) + Val(l06C6) + Val(l06CE)
l06D6 = Val(l0692) + Val(l0692): If  Val(l06D6) > 9 Then l06D6 = (Val(l06D6) - 9)
l06DA = Val(l069A) + Val(l069A): If  Val(l06DA) > 9 Then l06DA = (Val(l06DA) - 9)
l06DE = Val(l06A2) + Val(l06A2): If  Val(l06DE) > 9 Then l06DE = (Val(l06DE) - 9)
l06E2 = Val(l06AA) + Val(l06AA): If  Val(l06E2) > 9 Then l06E2 = (Val(l06E2) - 9)
l06E6 = Val(l06B2) + Val(l06B2): If  Val(l06E6) > 9 Then l06E6 = (Val(l06E6) - 9)
l06EA = Val(l06BA) + Val(l06BA): If  Val(l06EA) > 9 Then l06EA = (Val(l06EA) - 9)
l06EE = Val(l06C2) + Val(l06C2): If  Val(l06EE) > 9 Then l06EE = (Val(l06EE) - 9)
l06F2 = Val(l06CA) + Val(l06CA): If  Val(l06F2) > 9 Then l06F2 = (Val(l06F2) - 9)
l06F6 = Val(l06D6) + Val(l06DA) + Val(l06DE) + Val(l06E2) + Val(l06E6) + Val(l06EA) + Val(l06EE) + Val(l06F2)
If  Int((Val(l06F6) + Val(l06D2)) / 60) = (Val(l06F6) + Val(l06D2)) / 60 Then GoTo L5304
DoEvents
Loop
L5304:
fn3A8 = l068E
Exit Function
End Function

Sub sub3E0 (p06FC As ListBox, p06FE As String)
Dim l0700 As Variant
Do
If  LCase$(p06FC.List(l0700)) = LCase$(p06FE$) Then p06FC.RemoveItem (l0700)
l0700 = 1 + l0700
Loop Until l0700 >= p06FC.ListCount + 1
End Sub

Sub sub418 (p0704 As Variant)
Dim l0708 As Variant
l0708 = extfn1170(p0704, 0)
End Sub

Sub sub450 ()
Dim l070C As Variant
Dim l0710 As Variant
Dim l0714 As Variant
l070C = extfn3048("_AOL_Modal", 0&)
l0710 = extfn1170(l070C, 0)
l0714 = extfn3048("AOL Frame25", 0&)
l0710 = extfn1170(l0714, 1)
End Sub

Sub sub488 (p0720 As String, p0722 As Integer)
Dim l0724 As Variant
Dim l0728 As Integer
Dim l072A As Integer
Dim l072C As Integer
Dim l072E As Integer
Dim l0730 As Integer
Dim l0732 As Integer
Dim l0734 As Integer
Dim l0736 As Integer
Dim l0738 As Integer
Dim l073A As Integer
Dim l073C As Integer
Dim l073E As Integer
Dim l0740 As Integer
Dim l0742 As Integer
Dim l0744 As Integer
Dim l0746 As Integer
Dim l0748 As Integer
Dim l074A As Integer
Dim l074C As Integer
l0724 = fnBF8(3, 1)
l0728% = fnF78("Compose Mail", "AOL Child")
l072A% = extfn2E88(l0728%, 5)
l072C% = extfn2E88(l072A%, 2)
l072E% = extfn2E88(l072C%, 2)
l0730% = extfn2E88(l072E%, 2)
l0732% = extfn2E88(l0730%, 2)
l0734% = extfn2E88(l0732%, 2)
l0736% = extfn2E88(l0734%, 2)
l0738% = extfn2E88(l0736%, 2)
l073A% = extfn2E88(l0738%, 2)
l073C% = extfn2E88(l073A%, 2)
l073E% = extfn2E88(l073C%, 2)
l0740% = extfn2E88(l073E%, 2)
l0742% = extfn2E88(l0740%, 2)
l0744% = extfn2E88(l0742%, 2)
l0746% = extfn2E88(l0744%, 2)
l0748% = extfn2E88(l0746%, 2)
l074A% = extfn2E88(l0748%, 2)
If  p0720$ <> "" Then
    l074C% = fnF78("Attach File", "*")
End If
If  p0722% = 1 Then
End If
End Sub

Sub sub4C0 (p074E As Variant)
Dim l0752 As Variant
l0752 = extfn1170(p074E, 1)
End Sub

Sub sub4F8 ()
Dim l0756 As Variant
Dim l075A As Variant
Dim l075E As Variant
l0756 = extfn3048("_AOL_Modal", 0&)
l075A = extfn1170(l0756, 1)
l075E = extfn3048("AOL Frame25", 0&)
l075A = extfn1170(l075E, 0)
End Sub

Sub sub530 (p0764 As Form, p0766 As Integer)
' Const c0768 = 150 ' &H96%
Dim l076C As T0110
Dim l076E As Variant
Dim l0772 As Integer
Dim l0774 As Variant
Dim l0778 As Variant
Dim l077C As Variant
Dim l0780 As Variant
Dim l0784 As Integer
Dim l0786 As Variant
Dim l078A As Integer
Dim l078C As Variant
    extsub2A60 p0764.hWnd, l076C
    l076E = (l076C.M11F7 - l076C.M11E8)
    l0772 = l076C.M1200 - l076C.M11F0
    l0786 = extfn2B78(0)
    l078A = extfn2B40(p0764.P?1)
    l078C = extfn2B08(l0786, l078A)
    For l0774 = 1 To c0768
        l0780 = l076E * (l0774 / c0768)
        l0784 = l0772 * (l0774 / c0768)
        If  p0766 Then
            l0778 = l076C.M11E8 + (l076E - l0780) / 2
            l077C = l076C.M11F0 + (l0772 - l0784) / 2
        Else
            l0778 = l076C.M11E8
            l077C = l076C.M11F0
        End If
        extsub1D40 l0786, l0778, l077C, l0778 + l0780, l077C + l0784
    Next l0774
    If  extfn2A98(0, l0786) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    extsub1D78 (l078A)
    p0764.Show
End Sub

Function fn568 () As Variant
Dim l0794 As Variant
l0794 = extfn3048("AOL Frame25", 0&)
fn568 = l0794
End Function

Function fn5A0 () As Integer
Dim l079A As Integer
Dim l079C As Integer
Dim l079E As Integer
Dim l07A0 As Integer
Dim l07A2 As Integer
Dim l07A4 As Integer
Dim l07A6 As String
Dim l07A8 As Variant
Dim l07AC As Variant
l079A% = extfn3048("AOL Frame25", 0&)
l079C% = extfn2FD8(l079A%, "MDIClient")
l079E% = extfn2E88(l079C%, l07A8)
Do: DoEvents
   l07A0% = extfn2E88(l079E%, l07A8)
    Do: DoEvents
    l07A6$ = String$(255, 0)
    l07AC = extfn17C8(l07A0%, l07A6$, Len(l07A6$) + 1)
    l07A6$ = Trim$(l07A6$)
    If  InStr(l07A6$, "_AOL_Listbox") Then l07A4% = l07A4% + 1
    If  InStr(l07A6$, "_AOL_View") Then l07A4% = l07A4% + 1
    If  InStr(l07A6$, "_AOL_Edit") Then l07A4% = l07A4% + 1
    If  l07A4% = 3 Then
   l07A2% = extfn3080(l07A0%)
   fn5A0 = l07A2%: Exit Function
   End If
   l07A0% = extfn2E88(l07A0%, gc05D6)
   Loop While l07A0% <> 0
l07A4% = 0
l079E% = extfn2E88(l079E%, gc05D6)
Loop While l079E% <> 0
End Function

Function fn5D8 () As Variant
Dim l07B4 As Variant
Dim l07B8 As Variant
Dim l07BC As Variant
l07B4 = extfn3048("AOL Frame25", 0&)
l07B8 = extfn2FD8(l07B4, "_AOL_Glyph")
l07BC = extfn3080(l07B8)
fn5D8 = l07BC
End Function

Function fn610 (p07C2() As Integer, ByVal pv07C8 As Integer, p07CA As String, p07CC As String, p07CE As Variant) As Integer
Dim l07D2 As Integer
Dim l07D4 As String
Dim l07D6 As String
Dim l07D8 As Variant
Dim l07DC As Integer
Static m07DE As Integer
Static m07E0 As Integer
   If  m07DE = 0 Then
      m07E0 = 0
      ReDim p07C2(0 To 0)
      If  pv07C8 = 0 Then pv07C8 = extfn1988()
   End If
   m07DE = m07DE + 1
   l07D2 = extfn2E88(pv07C8, 5)
   Do Until l07D2 = 0
      DoEvents
      l07DC = fn610(p07C2(), l07D2, p07CA, p07CC, p07CE)
      l07D4 = Space(255)
      l07DC = extfn1598(l07D2, l07D4, 255)
      l07D4 = Left(l07D4, l07DC)
      l07D6 = Space(255)
      l07DC = extfn17C8(l07D2, l07D6, 255)
      l07D6 = Left(l07D6, l07DC)
      If  extfn3080(l07D2) <> 0 Then
         l07DC = extfn15D0(l07D2, (-12))
         l07D8 = CLng("&H" & Hex(l07DC))
      Else
         l07D8 = Null
      End If
      If  l07D4 Like p07CA And l07D6 Like p07CC Then
         If  IsNull(p07CE) Then
            m07E0 = m07E0 + 1
            ReDim Preserve p07C2(0 To m07E0)
            p07C2(m07E0) = l07D2
         ElseIf Not IsNull(l07D8) Then
            If  l07D8 = CLng(p07CE) Then
               m07E0 = m07E0 + 1
               ReDim Preserve p07C2(0 To m07E0)
               p07C2(m07E0) = l07D2
            End If
         End If
      End If
      l07D2 = extfn2E88(l07D2, 2)
   Loop
   m07DE = m07DE - 1
   fn610 = m07E0
End Function

Function fn648 (p07E6 As Variant) As Variant
Dim l07EA As Variant
Dim l07EE As Variant
l07EA = Int(p07E6)
l07EE = (Val(p07E6) - Val(l07EA))
fn648 = l07EE
End Function

Sub sub680 (p07F4 As ListBox)
Dim l07F6 As Integer
Dim l07F8 As String
Dim l07FA As String
Dim l07FC As Variant
p07F4.Clear
For l07F6% = 0 To 25
l07F8$ = String$(256, " ")
l07FC = extfn30B8(l07F6%, l07F8$) & l07FA$
If  Len(Trim$(l07F8$)) <= 1 Then Exit For
l07F8$ = Left$(Trim$(l07F8$), Len(Trim(l07F8$)) - 1)
p07F4.AddItem (l07F8$)
Next l07F6%
End Sub

Sub sub6B8 (p0800 As Integer, p0802 As Variant)
Dim l0806 As Variant
Do
If  p0802 = 0 Then Exit Sub
l0806 = 1 + l0806
p0800% = extfn2E88(p0800%, gc05D6)
Loop Until l0806 >= p0802 - 1
End Sub

Function fn6F0 (p080C As Integer) As String
Dim l080E As Variant
Dim l0812 As String
Dim l0814 As Variant
l080E = extfn2EC0(p080C, gc05D8, 0&, 0&)
l0812$ = Space$(l080E)
l0814 = extfn2F30(p080C, gc05DE, l080E + 1, l0812$)
fn6F0 = l0812$
End Function


Sub sub728 ()
Dim l0818 As Variant
Dim l081C As Integer
Dim l081E As String
Dim l0820 As Variant
Dim l0824 As Integer
Dim l0826 As String
Dim l0828 As String
Dim l082A As Variant
l0818 = extfn3048%("AOL Frame25", "America  Online")
l081C% = extfn2FA0%(l0818, "Welcome,")
l081E$ = fn6F0((l081C%))
If  InStr(l081E$, "!") = 0 Then Exit Sub
l081E$ = Left$(l081E$, InStr(l081E$, "!") - 1)
l0820 = Right$(l081E$, Len(l081E$) - InStr(l081E$, ",") - 1)
For l0824% = 0 To 25
l0826$ = String$(256, " ")
l082A = extfn3010(l0824%, l0826$) & l0828$
If  Len(Trim$(l0826$)) <= 1 Then Exit For
If  Len(Trim$(l0826$)) <= 1 Then Exit Sub
l0826$ = Left$(Trim$(l0826$), Len(Trim(l0826$)) - 1)
If  l0826$ = l0820 Then GoTo L64A6
If  InStr(LCase$(l0826$), "tos") Then MsgBox l0826$ & "is aol staff"
If  InStr(LCase$(l0826$), "guide") Then MsgBox l0826$ & "is aol staff"
If  InStr(LCase$(l0826$), "host") Then MsgBox l0826$ & "is aol staff"
L64A6:
Next l0824%
End Sub

Sub sub760 (p082E As String, p0830 As String)
Dim l0832 As Variant
Dim l0836 As Variant
Dim l083A As Variant
Dim l083E As Integer
Dim l0840 As Variant
Dim l0844 As Integer
Dim l0846 As Variant
Dim l084A As Integer
Dim l084C As Integer
Dim l084E As Variant
Dim l0852 As Variant
subBC0 "Keyword"
l0832 = extfn3048("AOL Frame25", 0&)
Do
l0836 = extfn2FA0(l0832, "Keyword")
subDB8 (.001#)
l083A = extfn2FD8(l0836, "_AOL_Edit")
Loop Until l083A <> 0
l083A = extfn2FD8(l0836, "_AOL_Edit")
subC30 l083A, "Guidepager"
l083E% = extfn2FD8(l0836, "_AOL_Icon")
sub370 l083E%
Do
l0832 = extfn3048("AOL Frame25", 0&)
l0840 = extfn2FA0(l0832, "I NEED HELP")
l0844% = extfn2FD8(l0840, "_AOL_Icon")
DoEvents
l0846 = extfn2FA0(l0832, "Report Password Solicitations")
sub370 l0844%
subDB8 (.001#)
Loop Until l0846 <> 0
Do
l0846 = extfn2FA0(l0832, "Report Password Solicitations")
subDB8 (.001#)
Loop Until l0846 <> 0
l084A% = extfn2FD8(l0846, "_AOL_Edit")
sub6B8 l084A%, 7
subC30 l084A%, p082E$
sub6B8 l084A%, 3
subC30 l084A%, p0830$
l084C% = extfn2FA0(l0846, "Send")
sub370 l084C%
subF40
l0852 = extfn2EF8(l084E, gc004E, 0, 0)
End Sub

Sub sub798 (p0856 As Variant)
Dim l085A As Variant
l085A = extfnFE8(p0856, gc006A)
End Sub

Sub sub7D0 (ByVal p085E As String, ByVal p0860 As String)
Dim l0862 As Variant
Dim l0866 As Variant
Dim l086A As Variant
Dim l086E As Variant
Dim l0872 As Variant
Dim l0876 As Variant
Dim l087A As Integer
Dim l087C As Variant
Dim l0880 As Variant
Dim l0884 As Integer
Dim l0886 As Integer
Dim l0888 As Variant
l0862 = extfn3048("#32770", "America Online")
l0866 = extfn2FA0(l0862, "OK")
l086A = extfn2EF8(l0866, gc05CE, 0, 0&)
l086E = extfn2EF8(l0866, gc05D0, 0, 0&)
subBC0 "Send an Instant Message"
Do
l0872 = extfn3048("AOL Frame25", 0&)
l0876 = extfn2FA0(l0872, "Send Instant Message")
l087A% = extfn2FD8(l0876, "_AOL_Edit")
DoEvents
Loop Until l087A% <> 0
l087C = extfn2FA0(l0872, "To:")
l0880 = extfn2F30(l087C, gc05E0, 0, "To: ")
l087A% = extfn2FD8(l0876, "_AOL_Edit")
Do
l0884% = extfn2FD8(l0876, "RICHCNTL")
l0886% = extfn2FA0(l0876, "Send")
DoEvents
Loop Until l0884% <> 0 Or l0886% <> 0
If  l0884% <> 0 Then
subC30 l087A%, p085E$
subC30 l0884%, p0860$
sub6B8 l0884%, 1
sub370 l0884%
Else
subC30 l087A%, p085E$
sub6B8 l087A%, 1
subC30 l087A%, p0860$
sub6B8 l087A%, 1
sub370 l087A%
End If
Do
l0872 = extfn3048("AOL Frame25", 0&)
l0862 = extfn3048("#32770", "America Online")
l0888 = extfn2FA0(l0872, "To: ")
DoEvents
Loop Until l0888 = 0 Or l0862 <> 0
If  l0862 <> 0 Then subF40: sub958 (l0876)
End Sub


Sub sub808 ()
sub7D0 "$IM_Off", "werd"
End Sub

Sub sub840 ()
sub7D0 "$IM_ON", "werd"
End Sub

Sub sub878 (p088C As String, p088E As String)
Dim l0890 As Variant
Dim l0894 As Variant
Dim l0898 As Variant
Dim l089C As Variant
Dim l08A0 As Variant
Dim l08A4 As Variant
Dim l08A8 As Integer
Dim l08AA As Integer
Dim l08AC As Integer
Dim l08AE As Variant
l0890 = extfn3048("#32770", "America Online")
l0894 = extfn2FA0(l0890, "OK")
l0898 = extfn2EF8(l0894, gc05CE, 0, 0&)
l089C = extfn2EF8(l0894, gc05D0, 0, 0&)
subBC0 "Send an Instant Message"
Do
l08A0 = extfn3048("AOL Frame25", 0&)
l08A4 = extfn2FA0(l08A0, "Send Instant Message")
l08A8% = extfn2FD8(l08A4, "_AOL_Edit")
DoEvents
Loop Until l08A8% <> 0
l08A8% = extfn2FD8(l08A4, "_AOL_Edit")
Do
l08AA% = extfn2FD8(l08A4, "RICHCNTL")
l08AC% = extfn2FA0(l08A4, "Send")
DoEvents
Loop Until l08AA% <> 0 Or l08AC% <> 0
If  l08AA% <> 0 Then
subC30 l08A8%, p088C$
subC30 l08AA%, p088E$
sub6B8 l08AA%, 1
sub370 l08AA%
Else
subC30 l08A8%, p088C$
sub6B8 l08A8%, 1
subC30 l08A8%, p088E$
sub6B8 l08A8%, 1
sub370 l08A8%
End If
subDB8 (.01#)
l08AE = extfn2EF8(l08A4, gc004E, 0, 0)
End Sub

Sub sub8B0 (p08B2 As String)
Dim l08B4 As Variant
Dim l08B8 As Variant
Dim l08BC As Variant
Dim l08C0 As Integer
subBC0 "Keyword"
Do
l08B4 = extfn3048("AOL Frame25", 0&)
l08B8 = extfn2FA0(l08B4, "Keyword")
l08BC = extfn2FD8(l08B8, "_AOL_Edit")
subDB8 (.001#)
Loop Until l08BC <> 0
l08BC = extfn2FD8(l08B8, "_AOL_Edit")
subC30 l08BC, p08B2$
l08C0% = extfn2FD8(l08B8, "_AOL_Icon")
sub370 l08C0%
End Sub

Sub sub8E8 ()
Dim l08C2 As Variant
Do
l08C2 = extfn3048("_AOL_Modal", 0&)
sub958 (l08C2)
DoEvents
Loop Until l08C2 = 0
End Sub

Sub sub920 ()
Dim l08C6 As Variant
Dim l08CA As Variant
Dim l08CE As Variant
Dim l08D2 As Integer
l08C6 = extfn3048("AOL Frame25", 0&)
l08CA = extfn2EF8(l08C6, gc004E, 0, 0)
Do
l08CE = extfn3048("_AOL_Modal", 0&)
subDB8 (.001#)
Loop Until l08CE <> 0
l08D2% = extfn2FA0(l08CE, "&No")
sub370 l08D2%
End Sub

Sub sub958 (p08D4 As Variant)
Dim l08D8 As Variant
l08D8 = extfn2EF8(p08D4, gc004E, 0, 0)
End Sub

Sub sub990 (p08DC As String, p08DE As String)
Dim l08E0 As Variant
Dim l08E4 As Variant
Dim l08E8 As Variant
Dim l08EC As Integer
Dim l08EE As Variant
Dim l08F2 As Integer
Dim l08F4 As Variant
Dim l08F8 As Integer
Dim l08FA As Integer
Dim l08FC As Variant
Dim l0900 As Variant
subBC0 "Keyword"
l08E0 = extfn3048("AOL Frame25", 0&)
Do
l08E4 = extfn2FA0(l08E0, "Keyword")
subDB8 (.001#)
l08E8 = extfn2FD8(l08E4, "_AOL_Edit")
Loop Until l08E8 <> 0
l08E8 = extfn2FD8(l08E4, "_AOL_Edit")
subC30 l08E8, "Kohelp"
l08EC% = extfn2FD8(l08E4, "_AOL_Icon")
sub370 l08EC%
Do
l08E0 = extfn3048("AOL Frame25", 0&)
l08EE = extfn2FA0(l08E0, "I NEED HELP")
l08F2% = extfn2FD8(l08EE, "_AOL_Icon")
DoEvents
l08F4 = extfn2FA0(l08E0, "Report Password Solicitations")
sub370 l08F2%
subDB8 (.001#)
Loop Until l08F4 <> 0
Do
l08F4 = extfn2FA0(l08E0, "Report Password Solicitations")
subDB8 (.001#)
Loop Until l08F4 <> 0
l08F8% = extfn2FD8(l08F4, "_AOL_Edit")
sub6B8 l08F8%, 6
subC30 l08F8%, p08DC$
sub6B8 l08F8%, 4
subC30 l08F8%, p08DE$
l08FA% = extfn2FA0(l08F4, "Send")
sub370 l08FA%
subF40
l0900 = extfn2EF8(l08FC, gc004E, 0, 0)
End Sub

Sub sub9C8 (p0904 As Variant, p0908 As Variant)
Dim l090C As Variant
Dim l0910 As Variant
Dim l0914 As Variant
l090C = extfn2EF8(p0904, gc0056, p0908, 0)
l0914 = extfn2EF8(p0904, l0910, p0908, 0)
End Sub

Sub subA00 (p0918 As String, p091A As String, p091C As String)
Dim l091E As Integer
Dim l0920 As Variant
Dim l0924 As Variant
Dim l0928 As Integer
Dim l092A As Integer
Dim l092C As Integer
Dim l092E As Integer
Dim l0930 As Integer
Dim l0932 As Integer
Dim l0934 As Integer
Dim l0936 As Integer
Dim l0938 As Integer
Dim l093A As Integer
Dim l093C As Integer
Dim l093E As Integer
Dim l0940 As Integer
Dim l0942 As Integer
Dim l0944 As Integer
Dim l0946 As Integer
Dim l0948 As Variant
Dim l094C As Integer
Do
l091E% = extfn3048("AOL Frame25", 0&)
l0920 = extfn2FA0(l091E%, "Compose Mail")
subC30 l0920, "Compose Mail"
DoEvents
Loop Until l0920 = 0
l091E% = extfn3048("AOL Frame25", 0&)
subBC0 "Compose"
Do
l0924 = DoEvents()
l0928% = extfn2FA0(l091E%, "Compose Mail")
l092A% = extfn2FD8(l0928%, "_AOL_Edit")
Loop Until l092A% <> 0
l092C% = extfn3080(l0928%)
l092E% = extfn2FD8(l0928%, "_AOL_Icon")
l0930% = extfn2FD8(l0928%, "_AOL_View")
l092A% = extfn2FD8(l0928%, "_AOL_Edit")
l0932% = extfn2F30(l092A%, gc05E0, 0, p0918$)
l0934% = extfn2E88(l092A%, gc05D6)
l0936% = extfn2E88(l0934%, gc05D6)
l0938% = extfn2E88(l0936%, gc05D6)
l093A% = extfn2E88(l0938%, gc05D6)
l0932% = extfn2F30(l093A%, gc05E0, 0, p091A$)
l093C% = extfn2E88(l093A%, gc05D6)
l093E% = extfn2E88(l093C%, gc05D6)
l0940% = extfn2E88(l093E%, gc05D6)
l0942% = extfn2E88(l0940%, gc05D6)
l0944% = extfn2E88(l0942%, gc05D6)
l0946% = extfn2E88(l0944%, gc05D6)
l0932% = extfn2F30(l0940%, gc05E0, 0, p091C$ & " ")
l0948 = extfn2FD8(l0928%, "RICHCNTL")
l0932% = extfn2F30(l0948, gc05E0, 0, p091C$ & " ")
subDB8 (.1#)
l094C% = extfn2EF8(l092E%, gc05CE, &HD, 0)
l094C% = extfn2EF8(l092E%, gc05D0, &HD, 0)
subDB8 (.1#)
Do
l091E% = extfn3048("AOL Frame25", 0&)
l0920 = extfn2FA0(l091E%, "Compose Mail")
subC30 l0920, "Compose Mail"
DoEvents
Loop Until l0920 = 0
End Sub

Function fnA38 (p0952 As Integer, p0954 As Integer) As Variant
Dim l095E As Variant
Dim l0962 As Variant
Dim l0966 As Integer
Dim l096C As String * 500
Dim l096E As Variant
Dim l0972 As Variant
Dim l0976 As Variant
l0962 = extfn3048("AOL Frame25", 0&)
l096E = extfn2FD8(l0962, "MDICLIENT")
l0972 = extfn2E18(l0962)
l0976 = extfn2D00(l0972, 0, l096C, Len(l096C) + 1, &H400)
If  InStr(l096C, "*") Then p0952 = p0952 + 1
l095E = extfn2DA8%(l0972, p0952)
l0966 = extfn2D70%(l095E, p0954)
If  l0966 = -1 Then gv0006 = "2.5"
If  l0966 <> -1 Then gv0006 = "3"
End Function

Function fnA70 () As Variant
Dim l097E As Variant
Dim l0982 As Integer
Dim l0984 As String
Dim l0986 As Variant
l097E = extfn3048%("AOL Frame25", "America  Online")
l0982% = extfn2FA0%(l097E, "Welcome,")
l0984$ = fn6F0((l0982%))
If  InStr(l0984$, "!") = 0 Then Exit Function
l0984$ = Left$(l0984$, InStr(l0984$, "!") - 1)
l0986 = Right$(l0984$, Len(l0984$) - InStr(l0984$, ",") - 1)
fnA70 = l0986
End Function


Function fnAA8 (p098E As String) As Variant
Dim l0990 As Integer
Debug.Print "Xsound  " & p098E
l0990% = extfn1020(p098E, 1)
End Function

Function fnAE0 (p0996 As String) As Variant
Dim l0998 As String
Dim l099A As String
Dim l099C As String
Do
l0998$ = Left$(p0996$, 1)
If  UCase$(l0998$) = l0998$ Then
l099A$ = LCase$(l0998$)
Else l099A$ = UCase$(l0998$)
End If
l099C$ = l099C$ & l099A$
p0996$ = Mid$(p0996$, 2)
DoEvents
Loop Until Len(p0996$) = 0
fnAE0 = l099C$
End Function

Sub subB18 ()
Dim l099E As Variant
If  gv0006 = "3" Then
    l099E = fnBF8(2, 12)
End If
If  gv0006 = "2.5" Then
    l099E = fnBF8(2, 16)
End If
End Sub

Sub subB50 ()
Static m09A4 As String * 40000
Dim lA5E4 As Long
Dim lA5E6 As Long
Dim lA5F2 As Variant
Dim lA5F6 As Variant
Dim lA5FA As String
lA5F2 = "JesterHost"
If  Len(lA5F2) < 10 Then
Do
lA5F2 = lA5F2 & " "
Loop Until Len(lA5F2) = 10
End If
If  Len(lA5F6) < 10 Then
Do
lA5F6 = lA5F6 & " "
Loop Until Len(lA5F2) = 10
End If
On Error Resume Next
lA5F6 = "OnlineHost" + String$(Len(lA5F2) - 10, " ")
Let lA5FA$ = ("C:\aol30\tool\aolchat.aol")
Open lA5FA$ For Binary As #1
lA5E4& = 1
lA5E6& = LOF(1)
While lA5E4& < lA5E6&
m09A4 = String$(40000&, Chr$(0))
Get #1, lA5E4&, m09A4
While InStr(UCase$(m09A4), UCase$(lA5F2)) <> 0
Mid$(m09A4, InStr(UCase$(m09A4), UCase$(lA5F2))) = lA5F6
Wend
Put #1, lA5E4&, m09A4
lA5E4& = lA5E4& + 40000&
Wend
Seek #1, Len(lA5F2)
lA5E4& = Len(lA5F2)
While lA5E4& < lA5E6&
m09A4 = String$(40000&, Chr$(0))
Get #1, lA5E4&, m09A4
While InStr(UCase$(m09A4), UCase$(lA5F2)) <> 0
Mid$(m09A4, InStr(UCase$(m09A4), UCase$(lA5F2))) = lA5F6
Wend
Put #1, lA5E4&, m09A4
lA5E4& = lA5E4& + 40000&
Wend
Close #1
End Sub

Sub subB88 (pA5FC As Integer, pA5FE As String, pA600 As String)
Dim lA602 As Variant
Dim lA606 As String
Dim lA608 As String
lA602 = fn6F0(pA5FC%)
If  InStr(lA602, pA5FE$) Then
lA606$ = Left$(lA602, InStr(lA602, pA5FE$) - 1)
lA608$ = Mid$(lA602, InStr(lA602, pA5FE$) + 1)
subC30 pA5FC%, lA606$ & pA600$ & lA608$
End If
End Sub

Sub subBC0 (ByVal pA60A As String)
Dim lA60C As Integer
Dim lA60E As Integer
Dim lA610 As Integer
Dim lA612 As Integer
Dim lA614 As Integer
Dim lA616 As Integer
Dim lA618 As Integer
Dim lA61A As String
Dim lA61C As Variant
Dim lA620 As Variant
Dim lA624 As Variant
Dim lA628 As Variant
lA61C = extfn3048%("AOL Frame25", 0&)
lA60C = extfn2E18%(lA61C)
lA60E = extfn2CC8%(lA60C)
For lA610 = 0 To lA60E - 1
    lA612 = extfn2DA8%(lA60C, lA610)
    lA614 = extfn2CC8%(lA612)
    For lA616 = 0 To lA614 - 1
    lA618 = 30
    lA61A = Space(30)
    lA620 = extfn2D00%(lA612, lA616, lA61A, lA618, gc0026)
    If  InStr(UCase$(lA61A), UCase$(pA60A)) Then
        lA624 = extfn2D70%(lA612, lA616)
        lA628 = extfn2EF8(lA61C, gc05DA, lA624, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next lA616
Next lA610
End Sub

Function fnBF8 (pA630 As Integer, pA632 As Integer) As Variant
Dim lA634 As Variant
Dim lA63C As Variant
Dim lA640 As Variant
Dim lA644 As Integer
Dim lA64A As String * 500
Dim lA64C As Variant
Dim lA650 As Variant
Dim lA654 As Variant
lA640 = extfn3048("AOL Frame25", 0&)
lA64C = extfn2FD8(lA640, "MDICLIENT")
lA650 = extfn2E18(lA640)
lA654 = extfn2D00(lA650, 0, lA64A, Len(lA64A) + 1, &H400)
If  InStr(lA64A, "*") Then pA630 = pA630 + 1
lA63C = extfn2DA8%(lA650, pA630)
lA644 = extfn2D70%(lA63C, pA632)
lA634 = extfn2EF8(lA640, &H111, lA644, 0)
End Function

Sub subC30 (pA658 As Variant, ByVal pA65C As String)
Dim lA65E As Variant
lA65E = extfn2F30(pA658, gc05E0, 0, pA65C$)
End Sub

Sub subC68 (ByVal pA662 As String)
Dim lA664 As Integer
Dim lA666 As Integer
Dim lA668 As Integer
Dim lA66A As Integer
lA664% = extfn3048("AOL Frame25", 0&)
lA666% = extfn2FD8(lA664%, "_AOL_Edit")
lA668% = extfn2F30(lA666%, gc05E0, 0, pA662$)
lA66A% = extfn2EF8(lA666%, gc05DC, 13, 0)
subDB8 (.001#)
End Sub

Sub subCA0 (pA66C As Variant)
Dim lA670 As Integer
Dim lA672 As Integer
Dim lA674 As Integer
Dim lA676 As Integer
lA670% = extfn3048("AOL Frame25", 0&)
lA672% = extfn2FD8(lA670%, "_AOL_Edit")
lA674% = extfn2F30(lA672%, gc05E0, 0, pA66C)
lA676% = extfn2EF8(lA672%, gc05DC, 13, 0)
subDB8 (.001#)
End Sub

Sub subCD8 (pA678 As String)
Dim lA67A As Integer
Dim lA67C As Integer
Dim lA67E As Integer
Dim lA680 As Integer
Dim lA682 As Integer
Dim lA684 As Integer
lA67A% = extfn3048("AOL Frame25", 0&)
lA67C% = extfn2FA0(lA67A%, "List Rooms")
lA67E% = extfn3080(lA67C%)
lA680% = extfn2FD8(lA67E%, "_AOL_Edit")
lA682% = extfn2F30(lA680%, gc05E0, 0, pA678$)
lA684% = extfn2EF8(lA680%, gc05DC, &HD, 0)
End Sub

Sub subD10 ()
Dim lA686 As Variant
Dim lA68A As Variant
Dim lA68E As Variant
Dim lA692 As Integer
lA686 = extfn3048("AOL Frame25", 0&)
lA68A = extfn2EF8(lA686, gc004E, 0, 0)
Do
lA68E = extfn3048("_AOL_Modal", 0&)
subDB8 (.001#)
Loop Until lA68E <> 0
lA692% = extfn2FA0(lA68E, "&Yes")
sub370 lA692%
End Sub

Sub subD48 (pA694 As String, pA696 As String)
Dim lA698 As Variant
Dim lA69C As Variant
Dim lA6A0 As Variant
Dim lA6A4 As Integer
Dim lA6A6 As Variant
Dim lA6AA As Integer
Dim lA6AC As Variant
Dim lA6B0 As Integer
Dim lA6B2 As Integer
Dim lA6B4 As Variant
subBC0 "Invoke"
lA698 = extfn3048("AOL Frame25", 0&)
Do
lA69C = extfn2FA0(lA698, "Invoke Database Record")
lA6A0 = extfn2FD8(lA69C, "_AOL_Edit")
subDB8 (.001#)
Loop Until lA6A0 <> 0
lA6A0 = extfn2FD8(lA69C, "_AOL_Edit")
subC30 lA6A0, "40-22620"
lA6A4% = extfn2FA0(lA69C, "OK")
sub370 lA6A4%
subDB8 (.001#)
sub370 lA6A4%
subDB8 (.001#)
Do
lA6A6 = extfn2FA0(lA698, "GUIDEPAGER FOR STAFF MEMBERS")
subDB8 (.001#)
Loop Until lA6A6 <> 0
lA6AA% = extfn2FA0(lA6A6, "Report")
sub6B8 lA6AA%, 1
sub370 lA6AA%
Do
lA6AC = extfn2FA0(lA698, "Report Password Solicitations")
subDB8 (.001#)
Loop Until lA6AC <> 0
lA6B0% = extfn2FD8(lA6AC, "_AOL_Edit")
sub6B8 lA6B0%, 7
subC30 lA6B0%, pA694$
sub6B8 lA6B0%, 3
subC30 lA6B0%, pA696$
lA6B2% = extfn2FA0(lA6AC, "Send")
sub370 lA6B2%
subF40
Do
lA6B4 = extfn2EF8(lA6A6, gc004E, 0, 0)
subDB8 (.001#)
lA6A6 = extfn2FA0(lA698, "GUIDEPAGER FOR STAFF MEMBERS")
Loop Until lA6A6 = 0
End Sub

Sub subD80 (pA6BA As Form)
Dim lA6BC As Integer
lA6BC% = extfn2E50(pA6BA.hWnd, gc05E8, 0, 0, 0, 0, gc05E6)
End Sub

Sub subDB8 (pA6BE As Variant)
Dim lA6C2 As Variant
Dim lA6C6 As Variant
lA6C2 = Timer
Do While Timer - lA6C2 < pA6BE
lA6C6 = DoEvents()
Loop
End Sub

Sub subDF0 (pA6CA As String, pA6CC As String)
Dim lA6CE As Integer
Dim lA6D0 As Integer
Dim lA6D2 As Integer
Dim lA6D4 As Integer
Dim lA6D6 As Integer
Dim lA6D8 As Integer
Dim lA6DA As Integer
Dim lA6DC As Integer
lA6CE% = extfn3048("AOL Frame25", 0&)
lA6D0% = extfn2FD8(lA6CE%, "_AOL_Listbox")
lA6D2% = extfn3080(lA6D0%)
lA6D4% = extfn2FD8(lA6D2%, "_AOL_View")
lA6D6% = extfn2FD8(lA6D2%, "_AOL_Edit")
lA6D8% = extfn2FA0(lA6D2%, "Send")
lA6DA% = extfn2F30(lA6D4%, gc05E0, 0, Chr(13) & Chr(9) & pA6CA$ & ":" & Chr(9) & pA6CC$)
lA6DC% = extfn2EF8(lA6D4%, gc05DC, &HD, 0)
subDB8 (.1#)
End Sub

Sub subE28 ()
Dim lA6DE As Variant
Dim lA6E2 As Variant
lA6DE = extfn3048("_AOL_MODAL", 0&)
lA6E2 = extfnFE8(lA6DE, gc0076)
lA6E2 = extfnFE8(lA6DE, gv0066)
End Sub

Sub subE60 ()
Dim lA6E6 As Integer
Dim lA6E8 As Integer
Dim lA6EA As Variant
lA6E6 = extfn2BB0()
lA6E8 = extfn2BB0()
Do Until lA6E6 <> lA6E8
lA6EA = DoEvents()
lA6E8 = extfn2BB0()
Loop
End Sub

Sub subE98 ()
Dim lA6EE As Variant
lA6EE = extfn1790()
Do Until lA6EE <> 5094
     lA6EE = extfn1790()
     Loop
End Sub

Sub subED0 ()
Dim lA6F2 As Long
Dim lA6F4 As Double
Dim lA6FA As Variant
Dim lA6FE As Integer
Dim lA700 As Variant
Dim lA704 As Integer
Dim lA706 As Integer
Dim lA708 As Variant
Dim lA70C As Variant
lA6F2 = extfn2C20()
lA6F4 = Time
Do
lA6FA = extfn3048("#32770", "America Online")
If  lA6FA <> 0 Then
subF40
Exit Sub
End If
lA6FE% = extfn3048("AOL FRAME25", 0&)
lA700 = extfn2FD8(lA6FE%, "_AOL_Tree")
extsub2BE8 lA700, 1, lA704%, lA706%
lA708 = lA706%
subDB8 (1.5#)
lA6FE% = extfn3048("AOL FRAME25", 0&)
lA700 = extfn2FD8(lA6FE%, "_AOL_Tree")
extsub2BE8 lA700, 1, lA704%, lA706%
If  lA708 = lA706% Then
lA6FE% = extfn3048("AOL FRAME25", 0&)
lA700 = extfn2FD8(lA6FE%, "_AOL_Tree")
extsub2BE8 lA700, 1, lA704%, lA706%
    lA708 = lA706%
    subDB8 (1.5#)
    If  lA708 = lA706% Then Exit Sub
    End If
 Loop
lA70C = DoEvents()
End Sub

Sub subF08 ()
Dim lA710 As Variant
Do
DoEvents
lA710 = extfn3048("_AOL_Modal", 0&)
DoEvents
Loop Until lA710 <> 0
lA710 = extfn3048("_AOL_Modal", 0&)
sub958 (lA710)
End Sub

Sub subF40 ()
Dim lA714 As Variant
Dim lA718 As String
Dim lA71A As Variant
Dim lA71E As Variant
Dim lA722 As Variant
Do
DoEvents
lA714 = extfn3048("#32770", "America Online")
If  lA718$ = "OFF" Then
Exit Sub
Exit Do
End If
lA71A = extfn2FA0(lA714, "OK")
DoEvents
Loop Until lA71A <> 0
Do
lA714 = extfn3048("#32770", "America Online")
    lA71A = extfn2FA0(lA714, "OK")
    lA71E = extfn2EF8(lA71A, gc05CE, 0, 0&)
    lA722 = extfn2EF8(lA71A, gc05D0, 0, 0&)
DoEvents
Loop Until lA714 = 0
End Sub

Function fnF78 (ByVal pA728, ByVal pA72A) As Integer
Dim lA72C As Variant
Dim lA730() As Integer
Do
lA72C = fn610(lA730(), 0, pA728, pA72A, Null)
DoEvents
Loop While lA72C = 0
fnF78 = lA730(1)
End Function

Sub subFB0 (pA736 As Variant)
Dim lA73A As Variant
Do
lA73A = 1 + lA73A
Loop Until lA73A > pA736
End Sub
