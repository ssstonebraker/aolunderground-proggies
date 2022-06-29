' Module1
Option Explicit
Declare Sub extsub2550 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub2588 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Function extfn2908 Lib "APIGuide.Dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfn2940 Lib "APIGuide.Dll" Alias "agDeviceMode" () As Integer
Declare Sub extsub25C0 Lib "APIGuide.Dll" Alias "agDWordTo2Integers" ()
Declare Function extfn2978 Lib "APIGuide.Dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn2710 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn2748 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn2780 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn26D8 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn27B8 Lib "APIGuide.Dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn2668 Lib "APIGuide.Dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn27F0 Lib "APIGuide.Dll" Alias "agGetControlName" () As String
Declare Function extfn26A0 Lib "APIGuide.Dll" Alias "agGetInstance" () As Integer
Declare Function extfn21D0 Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" () As String
Declare Function extfn2A20 Lib "APIGuide.Dll" Alias "agHugeOffset" () As Long
Declare Function extfn29B0 Lib "APIGuide.Dll" Alias "agInp" () As Integer
Declare Function extfn29E8 Lib "APIGuide.Dll" Alias "agInpw" () As Integer
Declare Sub extsub25F8 Lib "APIGuide.Dll" Alias "agOutp" ()
Declare Sub extsub2630 Lib "APIGuide.Dll" Alias "agOutpw" ()
Declare Function extfn2A58 Lib "APIGuide.Dll" Alias "agVBGetVersion" () As Integer
Declare Function extfn2A90 Lib "APIGuide.Dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfn2AC8 Lib "APIGuide.Dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn2828 Lib "APIGuide.Dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfn2898 Lib "APIGuide.Dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfn2860 Lib "APIGuide.Dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfn28D0 Lib "APIGuide.Dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfn2F98 Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn17C0 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn17F8 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn1C58 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfn2390 Lib "GDI" Alias "BitBlt" () As Integer
Declare Sub extsub1210 Lib "User" Alias "closewindow" ()
Declare Function extfn1788 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn3078 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn23C8 Lib "GDI" Alias "CreateSolidBrush" (ByVal p1&) As Integer
Declare Sub extsub2278 Lib "GDI" Alias "DeleteObject" (ByVal p1%)
Declare Sub extsub1328 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn2B00 Lib "APIGuide.Dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn1EF8 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfn15C8 Lib "User" Alias "ExitWindow" () As Integer
Declare Function extfn1130 Lib "311.Dll" Alias "AOLGetList" () As Integer
Declare Function extfn24A8 Lib "VBWFind.Dll" Alias "FindChild" () As Integer
Declare Function extfn2518 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn24E0 Lib "VBWFind.Dll" Alias "findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1520 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn1558 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn1590 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn2320 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn1980 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn1B40 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1440 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn1A98 Lib "User" Alias "getcurrenttime" () As Long
Declare Function extfn1B08 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfn1E18 Lib "User" Alias "GetDC" (ByVal p1%) As Integer
Declare Function extfn1DE0 Lib "User" Alias "GetDesktopWindow" () As Integer
Declare Function extfn22B0 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn1168 Lib "User" Alias "getfocus" () As Integer
Declare Function extfn2010 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn1408 Lib "User" Alias "GetFreeSystemResources" (ByVal p1%) As Integer
Declare Function extfn1C90 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn1D00 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn1CC8 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1D38 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn16A8 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn1670 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn2FD0 Lib "311.Dll" Alias "AOLGetList" () As Integer
Declare Function extfn1BB0 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn13D0 Lib "User" Alias "getnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1600 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn2198 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn20B8 Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn20F0 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn2128 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn1AD0 Lib "User" Alias "GetScrollPos" () As Integer
Declare Sub extsub1360 Lib "User" Alias "GetScrollRange" ()
Declare Function extfn1D70 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn19F0 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn2048 Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn1B78 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn1DA8 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn1C20 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn1FD8 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn14E8 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1E88 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn1F68 Lib "Kernel" Alias "GetWindowDirectory" () As Integer
Declare Sub extsub12B8 Lib "User" Alias "GetWindowRect" (ByVal p1%, p2 As Any)
Declare Function extfn11A0 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1BE8 Lib "User" Alias "GetWindowTextLength" () As Integer
Declare Function extfn1910 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1FA0 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn1830 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn1A60 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfn1F30 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn2470 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Function extfn2E80 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Sub extsub2B70 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Function extfn2F28 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfn2F60 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfn2CC0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2CF8 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn2C88 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2D30 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn2C50 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2D68 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn2C18 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Sub extsub2B38 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Function extfn2BE0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn2E10 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfn2DD8 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfn2DA0 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfn2E48 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfn2EB8 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfn2EF0 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Sub extsub2BA8 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Declare Sub extsub2240 Lib "GDI" Alias "Rectangle" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%)
Declare Function extfn1E50 Lib "User" Alias "ReleaseDC" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2400 Lib "GDI" Alias "SelectObject" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn16E0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn1750 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn1718 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn19B8 Lib "User" Alias "setactivewindow" (ByVal p1%) As Integer
Declare Sub extsub2208 Lib "GDI" Alias "setbkcolor" ()
Declare Function extfn1478 Lib "User" Alias "SetClassWord" () As Integer
Declare Sub extsub1398 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn14B0 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn1638 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn1A28 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn2358 Lib "GDI" Alias "SetTextColor" () As Long
Declare Sub extsub12F0 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn1948 Lib "User" Alias "setwindowtext" () As Integer
Declare Sub extsub1280 Lib "User" Alias "ShowOwnedPopups" ()
Declare Function extfn11D8 Lib "User" Alias "ShowWindow" () As Integer
Declare Function extfn2438 Lib "MMSystem" Alias "SndPlaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn1EC0 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn22E8 Lib "GDI" Alias "TextOut" () As Integer
Declare Sub extsub1248 Lib "User" Alias "UpdateWindow" ()
Declare Function extfn3008 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfn3040 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn1868 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn18D8 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn18A0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn2080 Lib "Kernel" Alias "WritePrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer
Declare Function extfn2160 Lib "Kernel" Alias "WriteProfileString" () As Integer
Const c0148 = 16 ' &H10%
Const c014A = 12 ' &HC%
Const c0FD4 = 150 ' &H96%

Sub subC8 (p0EFE As ListBox)
Dim l0F00 As Integer
Dim l0F02 As String
Dim l0F04 As String
Dim l0F06 As Variant
p0EFE.Clear
For l0F00% = 0 To 25
l0F02$ = String$(256, " ")
l0F06 = extfn2F98(l0F00%, l0F02$) & l0F04$
If  Len(Trim$(l0F02$)) <= 1 Then Exit For
l0F02$ = Left$(Trim$(l0F02$), Len(Trim(l0F02$)) - 1)
p0EFE.AddItem (l0F02$)
Next l0F00%
End Sub


Function fn100 () As Variant
fn100 = extfn1520("AOL Frame25", "America  Online")
DoEvents
End Function

Sub sub138 (p0F0E)
Dim l0F10 As Integer
Dim l0F12 As Integer
Dim l0F1A As Variant
Dim l0F1E() As Integer
Dim l0F24 As Variant
l0F10% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
l0F12% = fn1018("Keyword", "*")
l0F24 = extfn2518(l0F12%, "_AOL_EDIT")
l0F1A = fn5D0(l0F1E(), l0F12%, "*", "_AOL_Edit", Null)
subF00 l0F1E(1), p0F0E
l0F1A = extfn1750(l0F1E(1), &H102, 13, 0)
End Sub

Sub sub170 (p0F28 As Integer)
Dim l0F2A As String
Dim l0F2C As Variant
Dim l0F30 As Variant
If  l0F2A$ = "OFF" Then
Exit Sub
End If
subF38 (.1#)
l0F2C = extfn1750(p0F28, gc0842, 0, 0&)
l0F30 = extfn1750(p0F28, gc0844, 0, 0&)
End Sub

Function fn1A8 () As Variant
Dim l0F3A As Variant
Dim l0F3E As String
Dim l0F40 As Integer
Dim l0F42 As Integer
Dim l0F44 As String
Dim l0F46 As String
Let l0F3E$ = l0F3A
Let l0F40% = Len(l0F3E$)
Do While l0F42% <= l0F40%
Let l0F42% = l0F42% + 1
Let l0F44$ = Mid$(l0F3E$, l0F42%, 1)
Let l0F46$ = l0F44$ & l0F46$
Loop
fn1A8 = l0F46$
End Function

Sub sub1E0 (p0F4A As Form)
Dim l0F4C As Variant
Dim l0F50 As Variant
    l0F4C = (Screen.Width - p0F4A.Width) / 2
    l0F50 = (Screen.Height - p0F4A.Height) / 2
    p0F4A.Move l0F4C, l0F50
End Sub


Function fn218 () As Integer
End Function

Sub sub250 (p0F5E As Integer)
Dim l0F60 As Variant
l0F60 = extfn1750(extfn1600(p0F5E), &H111, (extfn1910(p0F5E, (-12))), ByVal CLng(p0F5E))
DoEvents
End Sub

Sub sub288 (p0F64 As Integer)
Dim l0F66 As Variant
l0F66 = extfn1750(p0F64, &H203, 0, 0&)
End Sub

Sub sub2C0 (p0F6A As String)
Dim l0F6C As Integer
Dim l0F6E As Variant
Dim l0F72 As Integer
Dim l0F74 As Variant
l0F6C% = extfn1520("AOL Frame25", 0&)
l0F6E = extfn24E0(l0F6C%, "New Mail")
l0F72% = extfn2518(l0F6E, "_AOL_TREE")
l0F74 = extfn1750(l0F72%, gc1094, 0, 0)
p0F6A = l0F74
End Sub

Sub sub2F8 ()
Dim l0F78 As Integer
Dim l0F7A As Variant
Dim l0F7E As Integer
Dim l0F80 As Variant
AppActivate "America  Online"
SendKeys "^r"
l0F78% = extfn1520("AOL Frame25", 0&)
l0F7A = extfn24E0(l0F78%, "New Mail")
l0F7E% = extfn2518(l0F7A, "_AOL_TREE")
l0F80 = extfn1750(l0F7E%, gc1094, 0, 0)
If  l0F80 > 1 Then
MsgBox "You have " & l0F80 & " messages in your Mailbox."
End If
If  l0F80 = 1 Then
MsgBox "You have 1 message in your Mailbox."
End If
If  l0F80 = 0 Then
MsgBox "You have no messages in your mailbox"
End If
End Sub

Function fn330 (p0F88 As String) As Variant
Dim l0F8A As String
Dim l0F8C As Integer
Dim l0F8E As Integer
Dim l0F90 As String
Dim l0F92 As String
Dim l0F94 As String
Dim l0F96 As Integer
Let l0F8A$ = p0F88
Let l0F8C% = Len(l0F8A$)
Do While l0F8E% <= l0F8C%
DoEvents
Let l0F8E% = l0F8E% + 1
Let l0F90$ = Mid$(l0F8A$, l0F8E%, 1)
Let l0F92$ = Mid$(l0F8A$, l0F8E%, 2)
If  l0F92$ = "ae" Then Let l0F92$ = "æ": Let l0F94$ = l0F94$ + l0F92$: Let l0F96% = 2: GoTo L1998
If  l0F92$ = "AE" Then Let l0F92$ = "Æ": Let l0F94$ = l0F94$ + l0F92$: Let l0F96% = 2: GoTo L1998
If  l0F92$ = "oe" Then Let l0F92$ = "œ": Let l0F94$ = l0F94$ + l0F92$: Let l0F96% = 2: GoTo L1998
If  l0F92$ = "OE" Then Let l0F92$ = "Œ": Let l0F94$ = l0F94$ + l0F92$: Let l0F96% = 2: GoTo L1998
If  l0F96% > 0 Then GoTo L1998
If  l0F90$ = "A" Then Let l0F90$ = "/\"
If  l0F90$ = "a" Then Let l0F90$ = "å"
If  l0F90$ = "B" Then Let l0F90$ = "ß"
If  l0F90$ = "C" Then Let l0F90$ = "Ç"
If  l0F90$ = "c" Then Let l0F90$ = "¢"
If  l0F90$ = "D" Then Let l0F90$ = "Ð"
If  l0F90$ = "d" Then Let l0F90$ = "ð"
If  l0F90$ = "E" Then Let l0F90$ = "Ê"
If  l0F90$ = "e" Then Let l0F90$ = "è"
If  l0F90$ = "f" Then Let l0F90$ = "ƒ"
If  l0F90$ = "H" Then Let l0F90$ = "|-|"
If  l0F90$ = "I" Then Let l0F90$ = "‡"
If  l0F90$ = "i" Then Let l0F90$ = "î"
If  l0F90$ = "k" Then Let l0F90$ = "|‹"
If  l0F90$ = "L" Then Let l0F90$ = "£"
If  l0F90$ = "M" Then Let l0F90$ = "|V|"
If  l0F90$ = "m" Then Let l0F90$ = "^^"
If  l0F90$ = "N" Then Let l0F90$ = "/\/"
If  l0F90$ = "n" Then Let l0F90$ = "ñ"
If  l0F90$ = "O" Then Let l0F90$ = "Ø"
If  l0F90$ = "o" Then Let l0F90$ = "º"
If  l0F90$ = "P" Then Let l0F90$ = "¶"
If  l0F90$ = "p" Then Let l0F90$ = "Þ"
If  l0F90$ = "r" Then Let l0F90$ = "®"
If  l0F90$ = "S" Then Let l0F90$ = "§"
If  l0F90$ = "s" Then Let l0F90$ = "$"
If  l0F90$ = "t" Then Let l0F90$ = "†"
If  l0F90$ = "U" Then Let l0F90$ = "Ú"
If  l0F90$ = "u" Then Let l0F90$ = "µ"
If  l0F90$ = "V" Then Let l0F90$ = "\/"
If  l0F90$ = "W" Then Let l0F90$ = "VV"
If  l0F90$ = "w" Then Let l0F90$ = "vv"
If  l0F90$ = "X" Then Let l0F90$ = "X"
If  l0F90$ = "x" Then Let l0F90$ = "×"
If  l0F90$ = "Y" Then Let l0F90$ = "¥"
If  l0F90$ = "y" Then Let l0F90$ = "ý"
If  l0F90$ = "!" Then Let l0F90$ = "¡"
If  l0F90$ = "?" Then Let l0F90$ = "¿"
If  l0F90$ = "." Then Let l0F90$ = "…"
If  l0F90$ = "," Then Let l0F90$ = "‚"
If  l0F90$ = "1" Then Let l0F90$ = "¹"
If  l0F90$ = "%" Then Let l0F90$ = "‰"
If  l0F90$ = "2" Then Let l0F90$ = "²"
If  l0F90$ = "3" Then Let l0F90$ = "³"
If  l0F90$ = "_" Then Let l0F90$ = "¯"
If  l0F90$ = "-" Then Let l0F90$ = "—"
If  l0F90$ = " " Then Let l0F90$ = " "
Let l0F94$ = l0F94$ + l0F90$

L1998:
If  l0F96% > 0 Then Let l0F96% = l0F96% - 1
DoEvents
Loop
fn330 = l0F94$
End Function

Sub sub368 (p0F98 As String, p0F9A As String, p0F9C As String, p0F9E As String, p0FA0 As String, p0FA2 As Integer)
Dim l0FA8 As Integer
Dim l0FAA As Integer
Dim l0FAC As Integer
Dim l0FAE As Integer
Dim l0FB0 As Integer
Dim l0FB2 As Integer
Dim l0FB4 As Integer
Dim l0FB6 As Integer
Dim l0FB8 As Integer
Dim l0FBA As Integer
Dim l0FBC As Integer
Dim l0FBE As Integer
Dim l0FC0 As Integer
Dim l0FC2 As Integer
Dim l0FC4 As Integer
Dim l0FC6 As Integer
Dim l0FC8 As Integer
Dim l0FCA As Integer
Dim l0FCC As Integer
Call subD08("&Compose Mail", "&Mail")
l0FA8% = fn1018("Compose Mail", "AOL Child")
l0FAA% = extfn14E8(l0FA8%, 5)
l0FAC% = extfn14E8(l0FAA%, 2)
l0FAE% = extfn14E8(l0FAC%, 2)
l0FB0% = extfn14E8(l0FAE%, 2)
l0FB2% = extfn14E8(l0FB0%, 2)
l0FB4% = extfn14E8(l0FB2%, 2)
l0FB6% = extfn14E8(l0FB4%, 2)
l0FB8% = extfn14E8(l0FB6%, 2)
l0FBA% = extfn14E8(l0FB8%, 2)
l0FBC% = extfn14E8(l0FBA%, 2)
l0FBE% = extfn14E8(l0FBC%, 2)
l0FC0% = extfn14E8(l0FBE%, 2)
l0FC2% = extfn14E8(l0FC0%, 2)
l0FC4% = extfn14E8(l0FC2%, 2)
l0FC6% = extfn14E8(l0FC4%, 2)
l0FC8% = extfn14E8(l0FC6%, 2)
l0FCA% = extfn14E8(l0FC8%, 2)
subF00 l0FBC%, CStr(p0F98$)
subF00 l0FC0%, CStr(p0F9A$)
subF00 l0FC4%, CStr(p0F9C$)
subF00 l0FCA%, CStr(p0F9E$)
If  p0FA0$ <> "" Then
    sub250 (l0FB4%): DoEvents
    l0FCC% = fn1018("Attach File", "*")
    subF00 extfn2518(l0FCC%, "Edit"), CStr(p0FA0$)
    sub250 (extfn24E0(l0FCC%, "OK")): DoEvents
End If
If  p0FA2% = 1 Then
    sub250 (l0FAC%)
End If
End Sub

Sub sub3A0 ()
End Sub

Sub sub3D8 (p0FD0 As Form, p0FD2 As Integer)
' Const c0FD4 = 150 ' &H96%
Dim l0FD8 As T00B2
Dim l0FDA As Variant
Dim l0FDE As Integer
Dim l0FE0 As Variant
Dim l0FE4 As Variant
Dim l0FE8 As Variant
Dim l0FEC As Variant
Dim l0FF0 As Integer
Dim l0FF2 As Variant
Dim l0FF6 As Integer
Dim l0FF8 As Variant
    extsub12B8 p0FD0.hWnd, l0FD8
    l0FDA = (l0FD8.M17C9 - l0FD8.M17BA)
    l0FDE = l0FD8.M17D2 - l0FD8.M17C2
    l0FF2 = extfn1E18(0)
    l0FF6 = extfn23C8(p0FD0.P?1)
    l0FF8 = extfn2400(l0FF2, l0FF6)
    For l0FE0 = 1 To c0FD4
        l0FEC = l0FDA * (l0FE0 / c0FD4)
        l0FF0 = l0FDE * (l0FE0 / c0FD4)
        If  p0FD2 Then
            l0FE4 = l0FD8.M17BA + (l0FDA - l0FEC) / 2
            l0FE8 = l0FD8.M17C2 + (l0FDE - l0FF0) / 2
        Else
            l0FE4 = l0FD8.M17BA
            l0FE8 = l0FD8.M17C2
        End If
        extsub2240 l0FF2, l0FE4, l0FE8, l0FE4 + l0FEC, l0FE8 + l0FF0
    Next l0FE0
    If  extfn1E50(0, l0FF2) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    extsub2278 (l0FF6)
    p0FD0.Show
End Sub

Function fn410 () As Integer
Dim l0FFE As Variant
Dim l1002 As Variant
Dim l1006 As Variant
Dim l100A As Variant
Dim l100E As Variant
l0FFE = extfn1520("AOL Frame25", 0&)
l1002 = extfn24E0(l0FFE, "List Rooms")
l1006 = extfn24E0(l0FFE, "Member Directory")
l100A = extfn24E0(l0FFE, "AOL Live!")
l100E = extfn24E0(l0FFE, "The Plaza")
If  extfn1600(l1002) = extfn1600(l1006) And extfn1600(l1002) = extfn1600(l100A) And extfn1600(l1002) = extfn1600(l100E) Then
    If  extfn2518(extfn1600(l1002), "_AOL_View") <> 0 And extfn2518(extfn1600(l1002), "_AOL_Edit") <> 0 And extfn2518(extfn1600(l1002), "_AOL_ListBox") <> 0 Then
    fn410 = extfn1600(l1002)
    Exit Function
    End If
End If
fn410 = 0
End Function

Function fn448 () As Integer
Dim l1014 As Variant
Dim l1018 As Variant
l1014 = extfn1520("AOL Frame25", 0&)
l1018 = extfn2518(l1014, "_AOL_ICON")
If  extfn1600(l1018) = extfn1600(l1018) And extfn1600(l1018) Then
    If  extfn2518(extfn1600(l1018), "_AOL_View") <> 0 And extfn2518(extfn1600(l1018), "_AOL_Edit") <> 0 And extfn2518(extfn1600(l1018), "_AOL_ListBox") <> 0 Then
    fn448 = extfn1600(l1018)
    subF38 (.001#)
    End If
End If
End Function

Function fn480 () As Integer
Dim l1028 As Integer
Dim l1032 As Integer
DoEvents
l1028 = 0
l1032% = extfn14E8(fn100(), 5)
DoEvents
Do
If  extfn24E0(l1032%, "Send") <> 0 Then l1028 = l1028 + 1
DoEvents
If  extfn24E0(l1032%, "Send Now") <> 0 Then l1028 = l1028 + 1
DoEvents
If  extfn24E0(l1032%, "Address" & Chr(13) & "Book") <> 0 Then l1028 = l1028 + 1
If  l1028 = 3 Then GoTo L25F4
DoEvents
l1032% = extfn14E8(l1032%, 2)
DoEvents
l1028 = 0
Loop While l1032% <> 0
Exit Function
L25F4:
DoEvents
fn480 = l1032%
Exit Function
End Function

Function fn4B8 () As Integer
End Function

Function fn4F0 () As Integer
Dim l1042 As Integer
Dim l104C As Integer
l1042 = 0
DoEvents
l104C% = extfn14E8(fn100(), 5)
DoEvents
Do
If  extfn24E0(l104C%, "Read") <> 0 Then l1042 = l1042 + 1
DoEvents
If  extfn24E0(l104C%, "Ignore") <> 0 Then l1042 = l1042 + 1
DoEvents
If  extfn24E0(l104C%, "Keep As New") <> 0 Then l1042 = l1042 + 1
DoEvents
If  extfn24E0(l104C%, "Delete") <> 0 Then l1042 = l1042 + 1
DoEvents
If  extfn2518(l104C%, "_AOL_Tree") <> 0 Then l1042 = l1042 + 1
If  l1042 = 5 Then GoTo L2810
DoEvents
l104C% = extfn14E8(l104C%, 2)
DoEvents
l1042 = 0
Loop While l104C% <> 0
Exit Function
L2810:
DoEvents
fn4F0 = l104C%
Exit Function
End Function

Function fn528 () As Integer
Dim l105A As Integer
Dim l1064 As Integer
l105A = 0
DoEvents
l1064% = extfn14E8(fn100(), 5)
DoEvents
Do
If  extfn24E0(l1064%, "Reply") <> 0 Then l105A = l105A + 1
DoEvents
If  extfn24E0(l1064%, "Forward") <> 0 Then l105A = l105A + 1
DoEvents
If  extfn24E0(l1064%, "Reply to All") <> 0 Then l105A = l105A + 1
DoEvents
If  l105A = 3 Then GoTo L2994
DoEvents
l1064% = extfn14E8(l1064%, 2)
DoEvents
l105A = 0
Loop While l1064% <> 0
Exit Function
L2994:
DoEvents
fn528 = l1064%
Exit Function
End Function

Function fn560 () As Variant
Dim l106A As Integer
Dim l106C As Integer
Dim l1070 As String * 255
Dim l1072 As String
Dim l1074 As Variant
l106A% = extfn1520("AOL Frame25", 0&)
l106C% = extfn24E0(l106A%, "Welcome,")
l1074 = extfn11A0(l106C%, l1070, 255)
l1070 = LTrim(RTrim(Trim(l1070)))
l1070 = Right$(l1070, Len(l1070) - InStr(l1070, ","))
l1072 = Left$(l1070, 10)
l1072 = Right$(l1070, Len(l1070))
If  InStr(l1072, "!") <> 0 Then l1072 = Left$(l1072, InStr(l1072, "!") - 1)
gv10A6$ = l1072$
fn560 = gv10A6$
End Function

Function fn598 () As Variant
Dim l107C As Integer
Dim l107E As Integer
Dim l1080 As Variant
l107E = extfn2518(fn100(), "AOL Child")

L2B76:
l107C = extfn24E0(l107E, "Send")
    If  l107C <> 0 Then Let l1080 = l1080 + 1
l107C = extfn24E0(l107E, "To:")
    If  l107C <> 0 Then Let l1080 = l1080 + 1
l107C = extfn24E0(l107E, "Available?")
    If  l107C <> 0 Then Let l1080 = l1080 + 1
If  l1080 = 3 Then
  fn598 = l107E
  Exit Function
End If
Let l1080 = 0
l107E = extfn13D0(l107E, 2)
If  l107E = extfn14E8(l107E, gc0AFC) Then
   fn598 = 0
   Exit Function
End If
End Function

Function fn5D0 (p1086() As Integer, ByVal pv108C As Integer, ByVal p108E As String, ByVal p1090 As String, p1092 As Variant) As Integer
Dim l1096 As Integer
Dim l1098 As String
Dim l109A As String
Dim l109C As Variant
Dim l10A0 As Integer
Static m10A2 As Integer
Static m10A4 As Integer
   If  m10A2 = 0 Then
      m10A4 = 0
      ReDim p1086(0 To 0)
      If  pv108C = 0 Then pv108C = extfn1DE0()
   End If
   m10A2 = m10A2 + 1
   l1096 = extfn14E8(pv108C, 5)
   Do Until l1096 = 0
      DoEvents
      l10A0 = fn5D0(p1086(), l1096, p108E, p1090, p1092)
      l1098 = Space(255)
      l10A0 = extfn11A0(l1096, l1098, 255)
      l1098 = Left(l1098, l10A0)
      l109A = Space(255)
      l10A0 = extfn1B40(l1096, l109A, 255)
      l109A = Left(l109A, l10A0)
      If  extfn1600(l1096) <> 0 Then
         l10A0 = extfn1910(l1096, (-12))
         l109C = CLng("&H" & Hex(l10A0))
      Else
         l109C = Null
      End If
      If  l1098 Like p108E And l109A Like p1090 Then
         If  IsNull(p1092) Then
            m10A4 = m10A4 + 1
            ReDim Preserve p1086(0 To m10A4)
            p1086(m10A4) = l1096
         ElseIf Not IsNull(l109C) Then
            If  l109C = CLng(p1092) Then
               m10A4 = m10A4 + 1
               ReDim Preserve p1086(0 To m10A4)
               p1086(m10A4) = l1096
            End If
         End If
      End If
      l1096 = extfn14E8(l1096, 2)
   Loop
   m10A2 = m10A2 - 1
   fn5D0 = m10A4
End Function

Function fn608 (p10AC As TextBox) As Variant
Dim l10AE As Variant
Dim l10B2 As Variant
Dim l10B6 As Variant
Dim l10BA As String
Randomize
For l10AE = 0 To Int(Rnd * 10) & Int(Rnd * 10)
    l10B2 = 13
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = "Hello, "
      If  l10B6 = 2 Then l10BA$ = "Good evening, "
      If  l10B6 = 3 Then l10BA$ = "Hello, And welcome to America Online. "
      If  l10B6 = 4 Then l10BA$ = "Welcome to America Online. "
      If  l10B6 = 5 Then l10BA$ = "Excuse me, "
      If  l10B6 = 6 Then l10BA$ = "Dear User, "
      If  l10B6 = 7 Then l10BA$ = "What's Up! "
      If  l10B6 = 8 Then l10BA$ = "Pardon me, "
      If  l10B6 = 9 Then l10BA$ = "Hi, "
      If  l10B6 = 10 Then l10BA$ = "Hello current member, "
      If  l10B6 = 11 Then l10BA$ = "My name is Jack Longfellow, "
      If  l10B6 = 12 Then l10BA$ = "Welcome new AOL user ,"
      If  l10B6 = 13 Then l10BA$ = "Greetings AOL user, "
    l10B2 = 13
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & "I am with America Online Billing "
      If  l10B6 = 2 Then l10BA$ = l10BA$ & "I am with OTC (Online Techincal Consultants) "
      If  l10B6 = 3 Then l10BA$ = l10BA$ & "I am with WWWC (World Wide Web Consultants) "
      If  l10B6 = 4 Then l10BA$ = l10BA$ & "I am with AOL Techincal Staff "
      If  l10B6 = 5 Then l10BA$ = l10BA$ & "I am with AOL System Security "
      If  l10B6 = 6 Then l10BA$ = l10BA$ & "I am with AOL Resource Department "
      If  l10B6 = 7 Then l10BA$ = l10BA$ & "I am with the AOL Community Action Team "
      If  l10B6 = 8 Then l10BA$ = l10BA$ & "I am with the America Online Billing Department, "
      If  l10B6 = 9 Then l10BA$ = l10BA$ & " I am with America Online "
      If  l10B6 = 10 Then l10BA$ = l10BA$ & "I am with AOL's Staff "
      If  l10B6 = 11 Then l10BA$ = l10BA$ & "I am an undercover Guide "
      If  l10B6 = 12 Then l10BA$ = l10BA$ & "I am with the FCS (FEDERAL CREDIT STAFF) "
      If  l10B6 = 13 Then l10BA$ = l10BA$ & "I am with the Guide Team here on AOL, "
      l10B2 = 13
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & "and due to Technical Failures "
      If  l10B6 = 2 Then l10BA$ = l10BA$ & "and due to Billing errors "
      If  l10B6 = 3 Then l10BA$ = l10BA$ & "and due to Line Noise "
      If  l10B6 = 4 Then l10BA$ = l10BA$ & "and due to A virus in our database "
      If  l10B6 = 5 Then l10BA$ = l10BA$ & "and due to Massive data flow in our database "
      If  l10B6 = 6 Then l10BA$ = l10BA$ & "and due to data corruption "
      If  l10B6 = 7 Then l10BA$ = l10BA$ & "and due to hackers by-passing out systems "
      If  l10B6 = 8 Then l10BA$ = l10BA$ & "and when you signed on an interuption occurred in our program "
      If  l10B6 = 9 Then l10BA$ = l10BA$ & "there was mishap here a few days ago resulting in lost information "
      If  l10B6 = 10 Then l10BA$ = l10BA$ & ""
      If  l10B6 = 11 Then l10BA$ = l10BA$ & ""
      If  l10B6 = 12 Then l10BA$ = l10BA$ & "  "
      If  l10B6 = 13 Then l10BA$ = l10BA$ & ""
    l10B2 = 5
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & "we seem to have lost your password. "
      If  l10B6 = 2 Then l10BA$ = l10BA$ & "we seem to have lost your account information. "
      If  l10B6 = 3 Then l10BA$ = l10BA$ & "we seem to have failed to recieve your logon password. "
      If  l10B6 = 4 Then l10BA$ = l10BA$ & "we seem to have failed to recieve your account information. "
      If  l10B6 = 5 Then l10BA$ = l10BA$ & "we have lost your Credit Card information. "
    l10B2 = 4
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & "To correct this situation please click respond and enter "
      If  l10B6 = 2 Then l10BA$ = l10BA$ & "Please click respond and enter your "
      If  l10B6 = 3 Then l10BA$ = l10BA$ & "Please help us by entering your "
      If  l10B6 = 4 Then l10BA$ = l10BA$ & "Please respond with your "
    l10B2 = 3
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & "Your password information."
      If  l10B6 = 2 Then l10BA$ = l10BA$ & "Your current account password."
      If  l10B6 = 3 Then l10BA$ = l10BA$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    l10B2 = 4
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & " Please respond within 2 minutes too keep this account active."
      If  l10B6 = 2 Then l10BA$ = l10BA$ & " It is very important that you respond immediately."
      If  l10B6 = 3 Then l10BA$ = l10BA$ & " Please respond as soon as possible "
      If  l10B6 = 4 Then l10BA$ = l10BA$ & " Incooperation may lead to termination of your account."
    l10B2 = 4
    l10B6 = Int(Rnd * l10B2 + 1)
      If  l10B6 = 1 Then l10BA$ = l10BA$ & " Thank you for using America Online."
      If  l10B6 = 2 Then l10BA$ = l10BA$ & " Thank you for your time."
      If  l10B6 = 3 Then l10BA$ = l10BA$ & " Thank you for your cooperation."
      If  l10B6 = 4 Then l10BA$ = l10BA$ & " Thank you for your help and enjoy the service!"
p10AC = l10BA$
Next
End Function

Function fn640 (p10C0 As Integer, p10C2 As String) As Variant
Dim l10C4 As Integer
Dim l10C6 As Variant
Dim l10CA As Variant
Dim l10CE As Integer
Dim l10D0 As String
Dim l10D2 As Integer
l10C4 = 0
l10CA = (l10C6)
If  l10C6 = 0 Then MsgBox "Are You An Asshole": Exit Function
DoEvents
l10CE% = extfn2518(p10C0%, p10C2$)
l10C4 = l10C4 + 1
Do
DoEvents
If  l10C4 = l10C6 Then
fn640 = l10CE%
Exit Function
End If
DoEvents
Do
DoEvents
l10CE% = extfn14E8(l10CE%, gc0AFE)
l10D0$ = String(255, 0)
l10D2% = extfn1B40(l10CE%, l10D0$, 254)
If  l10D0$ = p10C2$ Then l10C4 = l10C4 + 1
If  l10C4 = l10C6 Then Exit Do
Loop Until l10D0$ = p10C2$
Loop Until l10C4 = l10C6
fn640 = l10CE%
End Function

Function fn678 () As String
Dim l10D6 As Long
    l10D6 = extfn1FA0()
    If  l10D6 And gc01A4 Then
        fn678 = "486"
        ElseIf l10D6 And gc01A2 Then
            fn678 = "386"
            ElseIf l10D6 And gc01A0 Then
                fn678 = "286"
                Else
                    fn678 = "Other"
    End If
End Function

Function fn6B0 () As String
    fn6B0 = Format$(extfn1408(gc019A)) + "%"
End Function

Function fn6E8 () As String
    fn6E8 = Format$(extfn1408(gv0198)) + "%"
End Function

Function fn720 () As String
    fn720 = Format$(extfn1408(gc019C)) + "%"
End Function

Sub sub758 ()
Dim l10DE As Integer
Dim l10E0 As Integer
Dim l10E2 As Integer
Dim l10E4 As Integer
Dim l10E6 As Integer
Dim l10E8 As Variant
l10E0% = extfn1520("AOL FRAME25", 0&)
l10E2% = extfn2518(l10E0%, "MDIClient")
l10E4% = extfn2518(l10E2%, "AOL Child")
l10E6% = extfn2518(l10E4%, "_AOL_TREE")
If  l10E6% <> True Then
subCD0 3, 2
subF38 3
End If
Do
gv10AE = extfn1750(l10E6%, gc1094, 0, 0&)
subF38 .1#
l10DE = extfn1750(l10E6%, gc1094, 0, 0&)
Loop Until gv10AE = l10DE
subF38 .5#
l10E8 = extfn16E0(l10E6%, gc07A8, 0, 0)
subF38 .5#
End Sub

Function fn790 () As Integer
Dim l10EE As Integer
Dim l10F0 As Integer
Dim l10F2 As Variant
fn790% = 0
l10EE% = extfn2518(extfn1520("AOL Frame25", "America  Online"), "MDIClient")
l10F0% = extfn2518(l10EE%, "AOL Child")
l10F2 = 0
While l10F0% <> 0
    If  extfn2518(l10F0%, "_AOL_Static") Then l10F2 = l10F2 + 1
    If  extfn2518(l10F0%, "_AOL_Listbox") Then l10F2 = l10F2 + 1
    If  extfn2518(l10F0%, "_AOL_View") Then l10F2 = l10F2 + 1
    If  extfn2518(l10F0%, "_AOL_Edit") Then l10F2 = l10F2 + 1
    If  extfn2518(l10F0%, "_AOL_Glyph") Then l10F2 = l10F2 + 1
    If  l10F2 > 4 Then fn790% = l10F0%: Exit Function
    l10F0% = extfn14E8(l10F0%, gc0AFE)
Wend
fn790% = 0
End Function

Function fn7C8 () As Variant
Dim l10FA As Variant
Dim l10FE As Variant
Dim l1102 As Variant
Dim l1106 As String
Dim l1108 As Variant
Dim l110C As Variant
Dim l1110 As String
On Error Resume Next
l10FA = extfn1520("aol frame25", 0&)
l10FE = extfn24E0(l10FA, "Welcome,")
If  l10FE = 0 Then fn7C8 = 0: Exit Function
l1102 = extfn16E0(l10FE, gc07A4, 0, 0)
l1106$ = String$(l1102, 0)
l1108 = extfn1718(l10FE, gc07A2, l1102, l1106$)
l110C = InStr(l1106$, ",")
l1110$ = Mid$(l1106$, l110C + 2, (Len(l1106$) - (l110C + 1)))
l1110$ = fnF70(l1110$)
fn7C8 = l1110$
End Function


Function fn800 () As Variant
Dim l1116 As Integer
Dim l1118 As Integer
Dim l111C As String * 255
Dim l111E As String
Dim l1120 As Variant
l1116% = extfn1520("AOL Frame25", 0&)
l1118% = extfn24E0(l1116%, ">Instant Message From: ")
l1120 = extfn11A0(l1118%, l111C, 255)
l111C = LTrim(RTrim(Trim(l111C)))
l111C = Right$(l111C, Len(l111C) - InStr(l111C, ": "))
l111E = Left$(l111C, 24)
l111E = Right$(l111C, Len(l111C) - 1)
If  InStr(l111E, "") <> 0 Then l111E = Left$(l111E, InStr(l111E, ""))
gv10A6$ = l111E$
fn800 = gv10A6$
End Function

Function fn838 () As String
End Function

Function fn870 () As Integer
Dim l1128 As Variant
l1128 = extfn1168()
fn870 = extfn1600(l1128)
End Function

Function fn8A8 () As String
Dim l112E As Long
Dim l1130 As Integer
    l112E = extfn1FD8()
    l1130 = CInt(l112E And &HFFFF&)
    fn8A8 = Format$(l1130 And &HFF) + "." + Format$(CInt(l1130 / 256))
End Function

Function fn8E0 () As String
End Function

Sub sub918 (p1134 As Variant)
Dim l1138 As Variant
l1138 = extfn1750(p1134, &H203, 0, 0&)
l1138 = extfn1750(p1134, gc10C2, &H20, 0&)
End Sub

Function fn950 (p113E As String) As Integer
Dim l1140 As Integer
On Error Resume Next
    l1140 = Len(Dir$(p113E))
    If  Err Or l1140 = 0 Then
        fn950 = False
        Else
            fn950 = True
    End If
End Function

Sub sub988 ()
Dim l1142 As Integer
Dim l1144 As Variant
Dim l1148 As Variant
Dim l114C As Integer
Dim l114E As Integer
Dim l1150 As Integer
Dim l1152 As Variant
Dim l1156 As Variant
l1142% = extfn1520(0&, "America  Online")
Call subD08("Send an Instant Message", "Mem&bers")
Do
For l1144 = 1 To 25
DoEvents
Next l1144
l1148 = extfn24E0(l1142%, "Send Instant Message")
Loop Until l1148
DoEvents
l114C% = extfn2518(l1148, "_AOL_Edit")
l114E% = extfn13D0(l114C%, 2)
l1150% = extfn13D0(l114E%, 2)
subDB0 l114C%, "$im_off"
subDB0 l114E%, "-"
sub170 l1150%
l1152 = extfn24E0(l1148, "Send")
   subFE0
   l1156 = extfn1750(l1148, gc07A8, 0, 0)
End Sub

Sub sub9C0 ()
Dim l115A As Integer
Dim l115C As Variant
Dim l1160 As Variant
Dim l1164 As Integer
Dim l1166 As Integer
Dim l1168 As Integer
Dim l116A As Variant
Dim l116E As Variant
l115A% = extfn1520(0&, "America  Online")
Call subD08("Send an Instant Message", "Mem&bers")
Do
For l115C = 1 To 25
DoEvents
Next l115C
l1160 = extfn24E0(l115A%, "Send Instant Message")
Loop Until l1160
DoEvents
l1164% = extfn2518(l1160, "_AOL_Edit")
l1166% = extfn13D0(l1164%, 2)
l1168% = extfn13D0(l1166%, 2)
subDB0 l1164%, "$im_on"
subDB0 l1166%, "-"
sub170 l1168%
l116A = extfn24E0(l1160, "Send")
   subFE0
   l116E = extfn1750(l1160, gc07A8, 0, 0)
End Sub

Sub sub9F8 ()
Dim l1172 As Integer
Dim l1174 As Integer
Dim l1176 As Integer
l1172% = extfn1520("AOL Frame25", 0&)
l1176% = extfn24E0(l1174%, "Welcome, ")
If  l1176% = 0 Then GoTo L522C
Exit Sub
L522C:
MsgBox "Please open AOL and try again!", , "Error"
End
End Sub

Sub subA30 ()
Dim l1178 As Integer
Dim l117A As Variant
Dim l117E As Integer
Dim l1180 As Integer
Dim l1182 As Variant
l1178% = extfn1520("AOL Frame25", 0&)
If  l1178% = 0 Then Exit Sub
l117A = extfn14B0(l1178%)
l117A = extfn1750(l1178%, &H11, 0, &H0)
Do
l117A = DoEvents()
l117E% = extfn1520("_AOL_Modal", "America Online")
Loop Until l117E% <> 0
l1180% = extfn24E0%(l117E%, "&No")
l1182 = extfn1750(l1180%, &H201, 0, 0&)
l1182 = extfn1750(l1180%, &H202, 0, 0&)
End Sub

Function fnA68 (ByVal pv118A As Variant, ByVal pv118E As Variant, p1192 As Integer) As Variant
Dim l1194 As String
Dim l1196 As Variant
Dim l119A As String
Dim l119C As Integer
Dim l119E As Variant
Dim l11A2 As Variant
Dim l11A6 As Variant
Dim l11AA As Variant
  On Error GoTo L5AEA
  If  Len(pv118A) = 0 Then Error 31100
  If  Len(pv118A) > 255 Then Error 31100
  If  Len(pv118E) = 0 Then Error 31100
  If  p1192% = 0 Then
    l1194$ = Left$(pv118E, 4) + Right$(pv118E, 4)
    If  l1194$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      pv118E = Mid$(pv118E, 5, Len(pv118E) - 8)
      l1196 = 1
      Do
        l1196 = InStr(l1196, pv118E, Chr$(1))
        If  l1196 = 0 Then
          Exit Do
        Else
          l119A$ = Chr$(Asc(Mid$(pv118E, l1196 + 1)) - 1)
          pv118E = Left$(pv118E, l1196 - 1) + l119A$ + Mid$(pv118E, l1196 + 2)
        End If
        l1196 = l1196 + 1
      Loop
      l119C% = False
    Else
      l119C% = True
    End If
  Else
    l119C% = True
  End If
  l119E = 1
  l11A2 = Len(pv118A)
  pv118A = Chr$(Asc(Left$(pv118A, 1)) Xor l11A2) + pv118A
  pv118A = Chr$(Asc(Mid$(pv118A, 1, 1)) Xor Asc(Mid$(pv118A, 2, 1))) + pv118A
  pv118A = pv118A + Chr$(Asc(Right$(pv118A, 1)) Xor l11A2)
  pv118A = pv118A + Chr$(Asc(Right$(pv118A, 2)) Xor Asc(Right$(pv118A, 1)))
  If  l119C% = True Then
    pv118E = Left$(pv118A, 3) + Format$(Asc(Right$(pv118A, 1)), "000") + Format$(Len(pv118A), "000") + pv118E
  End If
  For l11A6 = 1 To Len(pv118E)
    l11AA = Asc(Mid$(pv118E, l11A6, 1)) Xor Asc(Mid$(pv118A, l119E, 1))
    Mid$(pv118E, l11A6, 1) = Chr$(l11AA)
    l119E = l119E + 1
    If  l119E > l11A2 + 4 Then l119E = 1
  Next l11A6
  If  l119C% = True Then
    l1196 = 1
    Do
      l1196 = InStr(l1196, pv118E, Chr$(1))
      If  l1196 > 0 Then
        pv118E = Left$(pv118E, l1196 - 1) + Chr$(1) + Chr$(2) + Mid$(pv118E, l1196 + 1)
        l1196 = l1196 + 1
      End If
    Loop While l1196 > 0
    Do
      l1196 = InStr(pv118E, Chr$(0))
      If  l1196 > 0 Then pv118E = Left$(pv118E, l1196 - 1) + Chr$(1) + Chr$(1) + Mid$(pv118E, l1196 + 1)
    Loop While l1196 > 0
    Do
      l1196 = InStr(pv118E, Chr$(10))
      If  l1196 > 0 Then pv118E = Left$(pv118E, l1196 - 1) + Chr$(1) + Chr$(11) + Mid$(pv118E, l1196 + 1)
    Loop While l1196 > 0
    Do
      l1196 = InStr(pv118E, Chr$(13))
      If  l1196 > 0 Then pv118E = Left$(pv118E, l1196 - 1) + Chr$(1) + Chr$(14) + Mid$(pv118E, l1196 + 1)
    Loop While l1196 > 0
    Do
      l1196 = InStr(pv118E, Chr$(26))
      If  l1196 > 0 Then pv118E = Left$(pv118E, l1196 - 1) + Chr$(1) + Chr$(27) + Mid$(pv118E, l1196 + 1)
    Loop While l1196 > 0
    pv118E = Chr$(1) + "KT" + Chr$(1) + pv118E + Chr$(1) + "KT" + Chr$(1)
  Else
    If  Left$(pv118E, 9) <> Left$(pv118A, 3) + Format$(Asc(Right$(pv118A, 1)), "000") + Format$(Len(pv118A), "000") Then
      Error 31100
    Else
      pv118E = Mid$(pv118E, 10)
    End If
  End If
  fnA68 = pv118E
  Exit Function

L5AEA:
  Exit Function
End Function

Sub subAA0 ()
Dim l11AE As Integer
Dim l11B0 As Variant
Dim l11B4 As Integer
Dim l11B6 As Integer
Dim l11B8 As Variant
l11AE% = extfn1520("AOL Frame25", 0&)
If  l11AE% = 0 Then Exit Sub
l11B0 = extfn14B0(l11AE%)
l11B0 = extfn1750(l11AE%, &H11, 0, &H0)
Do
l11B0 = DoEvents()
l11B4% = extfn1520("_AOL_Modal", "America Online")
Loop Until l11B4% <> 0
l11B6% = extfn24E0%(l11B4%, "&Yes")
l11B8 = extfn1750(l11B6%, &H201, 0, 0&)
l11B8 = extfn1750(l11B6%, &H202, 0, 0&)
End Sub

Sub subAD8 ()
End Sub

Sub subB10 ()
Dim l11BC As Integer
Dim l11BE As Integer
Dim l11C0 As Integer
Dim l11C2 As Variant
l11BC% = extfn1520("AOL FRAME25", 0&)
l11BE% = extfn24E0(l11BC%, "New Mail")
l11C0% = extfn2518(l11BE%, "_AOL_TREE")
l11C2 = extfn1750(l11C0%, gc1094, 0, 0)
If  l11C2 > 1 Then
MsgBox "You have " & l11C2 & " messages in your Mailbox...", 0
End If
If  l11C2 = 1 Then
MsgBox "You have " & l11C2 & " message in your Mailbox...", 0
End If
If  l11C2 < 1 Then
MsgBox "You have no messages in your Mailbox...", 0
End If
End Sub

Sub subB48 (p11C6)
Dim l11C8 As Integer
Dim l11CA As Integer
Dim l11D2 As Variant
Dim l11D6() As Integer
l11C8% = extfn1520("AOL Frame25", 0&)
Call subD08("Keyword...", "&Go To")
l11CA% = fn1018("Keyword", "*")
l11D2 = fn5D0(l11D6(), l11CA%, "*", "_AOL_Edit", Null)
subF00 l11D6(1), p11C6
l11D2 = extfn1750(l11D6(1), &H102, 13, 0)
End Sub

Sub subB80 ()
Dim l11E2 As Integer
Dim l11E6 As Integer
Dim l11EC As Variant
Dim l11F4 As Integer
Dim l11F6 As Integer
Dim l11FC As Variant
Dim l1200 As Integer
Dim l1202 As Integer
Dim l1204 As Integer
Dim l1206 As Integer
Dim l1208 As Variant
Dim l120C As Variant
Dim l1210 As Variant
Dim l1214 As Variant
l1200% = extfn1520(0&, "America  Online")
Call subD08("Send an Instant Message", "Mem&bers")
Do
For l11EC = 1 To 25
DoEvents
Next l11EC
l11E2 = extfn24E0(l1200%, "Send Instant Message")
Loop Until l11E2
DoEvents
l1202% = extfn2518(l11E2, "_AOL_Edit")
l1204% = extfn13D0(l1202%, 2)
l1206% = extfn13D0(l1204%, 2)
subDB0 l1202%, "$im_on"
subDB0 l1204%, "" + Chr(83) + Chr(117) + Chr(112) + Chr(101) + Chr(114) + Chr(110) + Chr(97) + Chr(117) + Chr(116) + " rules"
sub170 l1206%
l11E6 = extfn24E0(l11E2, "Send")
Call subE58(l11E6)
Do
    DoEvents
    l11F4% = extfn1980()
    l11F6% = extfn24E0(l11F4%, "OK")
Loop Until l11F6%
Call subE58(l11F6%)
l11FC = extfn16E0(l11E2, c0148, 0, 0)
l1208 = extfn1520("#32770", "America Online")
Do
DoEvents
Loop Until l1208 <> 0
    l120C = extfn24E0(l1208, "OK")
    l1210 = extfn1750(l120C, gc0842, 0, 0&)
    l1214 = extfn1750(l120C, gc0844, 0, 0&)
End Sub

Sub subBB8 (ByVal p1218 As String)
Dim l121A As Integer
Debug.Print "Xsound  " & p1218
l121A% = extfn2438(p1218, 1)
End Sub

Function fnBF0 (p1220 As String) As Variant
Dim l1222 As String
Dim l1224 As Integer
Dim l1226 As Integer
Dim l1228 As String
Dim l122A As String
Let l1222$ = p1220
Let l1224% = Len(l1222$)
Do While l1226% <= l1224%
Let l1226% = l1226% + 1
Let l1228$ = Mid$(l1222$, l1226%, 1)
If  l1228$ = "A" Then Let l1228$ = "a"
If  l1228$ = "E" Then Let l1228$ = "e"
If  l1228$ = "I" Then Let l1228$ = "i"
If  l1228$ = "O" Then Let l1228$ = "o"
If  l1228$ = "U" Then Let l1228$ = "u"
If  l1228$ = "b" Then Let l1228$ = "B"
If  l1228$ = "c" Then Let l1228$ = "C"
If  l1228$ = "d" Then Let l1228$ = "D"
If  l1228$ = "z" Then Let l1228$ = "Z"
If  l1228$ = "f" Then Let l1228$ = "F"
If  l1228$ = "g" Then Let l1228$ = "G"
If  l1228$ = "h" Then Let l1228$ = "H"
If  l1228$ = "y" Then Let l1228$ = "Y"
If  l1228$ = "j" Then Let l1228$ = "J"
If  l1228$ = "k" Then Let l1228$ = "K"
If  l1228$ = "l" Then Let l1228$ = "L"
If  l1228$ = "m" Then Let l1228$ = "M"
If  l1228$ = "n" Then Let l1228$ = "N"
If  l1228$ = "x" Then Let l1228$ = "X"
If  l1228$ = "p" Then Let l1228$ = "P"
If  l1228$ = "q" Then Let l1228$ = "Q"
If  l1228$ = "r" Then Let l1228$ = "R"
If  l1228$ = "s" Then Let l1228$ = "S"
If  l1228$ = "t" Then Let l1228$ = "T"
If  l1228$ = "w" Then Let l1228$ = "W"
If  l1228$ = "v" Then Let l1228$ = "V"
If  l1228$ = " " Then Let l1228$ = " "
Let l122A$ = l122A$ + l1228$
Loop
fnBF0 = l122A$
End Function

Function fnC28 (p122E As Variant, p1232 As Variant, p1236 As String) As String
Dim l1238 As String
    l1238 = String(255, Chr(0))
    fnC28 = Left(l1238, extfn20B8(p122E, ByVal p1232, "", l1238, Len(l1238), p1236))
End Function

Sub subC60 (p123A)
Seek #1, 6887
Put #1, , p123A
Close #1
End Sub

Sub subC98 (p123E As TextBox, p1242 As TextBox)
Static m1246 As String * 40000
Dim lAE86 As Long
Dim lAE88 As Long
Dim lAE94 As Variant
Dim lAE98 As String
Screen.MousePointer = 11
If  UCase$(Trim$(p123E)) = "NEW USER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo L6C52
If  Len(p123E) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
lAE94 = "New User" + String$(Len(p123E) - 7, " ")
Let lAE98$ = (p1242 & "\idb\main.idx")
Open lAE98$ For Binary As #1
lAE86& = 1
lAE88& = LOF(1)
While lAE86& < lAE88&
    m1246 = String$(40000&, Chr$(0))
    Get #1, lAE86&, m1246
    While InStr(UCase$(m1246), UCase$(p123E)) <> 0
        Mid$(m1246, InStr(UCase$(m1246), UCase$(p123E))) = lAE94
    Wend
    Put #1, lAE86&, m1246
    lAE86& = lAE86& + 40000&
Wend
Seek #1, Len(p123E)
lAE86& = Len(p123E)
While lAE86& < lAE88&
m1246 = String$(40000&, Chr$(0))
    Get #1, lAE86&, m1246
    While InStr(UCase$(m1246), UCase$(p123E)) <> 0
        Mid$(m1246, InStr(UCase$(m1246), UCase$(p123E))) = lAE94
        Wend
    Put #1, lAE86&, m1246
    lAE86& = lAE86& + 40000&
Wend
Close #1
Screen.MousePointer = 0
L6C52:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub

Sub subCD0 (pAE9A As Variant, pAE9E As Variant)
Dim lAEA2 As Variant
Dim lAEA6 As Variant
Dim lAEAA As Variant
Dim lAEAE As Variant
Dim lAEB2 As Integer
lAEB2 = extfn1520("AOL Frame25", 0&)
lAEAE = extfn1C90(lAEB2)
lAEAA = extfn1D70(lAEAE, pAE9A)
lAEA6 = extfn1CC8(lAEAA, pAE9E)
lAEA2 = extfn1750(lAEB2, gc082A, lAEA6, 0)
End Sub

Sub subD08 (ByVal pAEB4 As String, ByVal pAEB6 As String)
Dim lAEB8 As Variant
Dim lAEBC As Integer
Dim lAEBE As Variant
Dim lAEC2 As String
Dim lAEC4 As Integer
Dim lAEC6 As Variant
Dim lAECA As Variant
Dim lAECE As Variant
Dim lAED2 As Integer
Dim lAED4 As Variant
lAEB8 = -1
lAEBC% = extfn1520("AOL Frame20", 0&)
If  lAEBC% = 0 Then lAEBC% = extfn1520("AOL Frame25", 0&)
If  lAEBC% = 0 Then lAEBC% = extfn1520("AOL Frame30", 0&)
If  lAEBC% = 0 Then lAEBC% = extfn1520("AOL Frame31", 0&)
If  lAEBC% = 0 Then lAEBC% = extfn1520("AOL Frame95", 0&)
lAEBE = extfn1C90(lAEBC%)
Do
    DoEvents
    lAEB8 = lAEB8 + 1
    lAEC2$ = String$(255, 0)
    lAEC4% = extfn16A8(lAEBE, lAEB8, lAEC2$, Len(pAEB6) + 1, gc0B36)
    lAEC6 = fnFA8(lAEC2$)
    If  lAEC6 = pAEB6 Then Exit Do
Loop
lAECA = extfn1D70(lAEBE, lAEB8)
lAECE = -1
Do
    DoEvents
    lAECE = lAECE + 1
    lAEC2$ = String(255, 0)
    lAEC4% = extfn16A8(lAECA, lAECE, lAEC2$, Len(pAEB4) + 1, gc0B36)
    lAEC6 = fnFA8(lAEC2$)
    If  lAEC6 = pAEB4 Then Exit Do
Loop
DoEvents
lAED2% = extfn1CC8(lAECA, lAECE)
lAED4 = extfn1750(lAEBC%, gc082A, lAED2%, 0&)
End Sub

Sub subD40 (pAED8, pAEDA)
Dim lAEDC As String
Dim lAEDE As Variant
Dim lAEE2 As Integer
Dim lAEE4 As Integer
Dim lAEE6 As Integer
Dim lAEE8 As Variant
If  lAEDC$ = "OFF" Then
Exit Sub
End If
Call subD08("Send an Instant Message", "Mem&bers")
Do Until lAEDE <> 0
DoEvents
lAEDE = fn598()
DoEvents
Loop
subF38 (.05#)
lAEE2% = extfn2518(lAEDE, "_AOL_Edit")
lAEE4% = extfn13D0(lAEE2%, 2)
lAEE6% = extfn13D0(lAEE4%, 2)
subDB0 lAEE2%, pAED8
subDB0 lAEE4%, pAEDA
sub170 lAEE6%
lAEE8 = 69
Do Until lAEE8 = 0
DoEvents
lAEE8 = fn598()
DoEvents
Loop
End Sub

Sub subD78 (pAEEC As Variant)
Dim lAEF0 As Integer
Dim lAEF2 As Integer
Dim lAEF4 As Integer
Dim lAEF6 As Integer
Dim lAEF8 As Integer
Dim lAEFA As Variant
lAEF0% = extfn1520("AOL FRAME25", 0&)
lAEF2% = extfn2518(lAEF0%, "_AOL_Listbox")
If  lAEF2% = 0 Then
Exit Sub
End If
lAEF4% = extfn1600(lAEF2%)
lAEF6% = extfn2518(lAEF4%, "_AOL_EDIT")
lAEF8% = extfn24E0(lAEF4%, "Send")
lAEFA = extfn1718(lAEF6%, gc10BE, 0, pAEEC)
lAEFA = extfn1750(lAEF6%, gc081A, 13, 0)
subF38 9.99988867182683E-321#
End Sub


Sub subDB0 (pAEFE As Integer, ByVal pAF00 As String)
Dim lAF02 As Variant
lAF02 = extfn1718(pAEFE, gc10BE, 0, pAF00)
End Sub


Sub subDE8 ()
Dim lAF06 As Integer
Dim lAF08 As Variant
lAF06% = extfn1520("AOL Frame25", 0&)
lAF08 = extfn24E0(lAF06%, "Welcome,")
If  lAF08 = 0 Then
MsgBox "You must be signed on!", , "Error"
Exit Sub
End If
End Sub

Sub subE20 (pAF0E As Form)
extsub12F0 pAF0E.hWnd, -1, 0, 0, 0, 0, &H50
End Sub

Sub subE58 (pAF10 As Variant)
Dim lAF14 As Variant
Dim lAF18 As Integer
Dim lAF1A As Variant
Dim lAF1E As Variant
    lAF14 = extfn14B0(pAF10)
    lAF14 = extfn19B8(pAF10)
   subF38 (.3#)
    lAF18% = extfn16E0(pAF10, lAF1A, 0, 0&)
    lAF18% = extfn16E0(pAF10, lAF1E, 0, 0&)
End Sub

Sub subE90 (pAF22 As Variant, pAF26)
Dim lAF2A As Variant
    lAF2A = extfn16E0(pAF22, c014A, 0, pAF26)
End Sub

Sub subEC8 (pAF2E As String)
Dim lAF30 As Integer
Dim lAF32 As String
Dim lAF34 As Integer
Dim lAF36 As Integer
Dim lAF38 As Integer
lAF30 = 0
lAF34 = 0
lAF36 = Len(pAF2E)
subF38 (1)
Do
  lAF32 = UCase$(Left$(pAF2E, 1))
  Randomize
  lAF38 = Int(3 * Rnd) + 1
  If  lAF32 = "A" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- A is for Asshole -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- A is for the Anal sex he gets -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- A is for the Ass he never gets -•-vVv-•-"
     End If
  ElseIf lAF32 = "B" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- B is for Bitchass -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- B is for how Bobbit cut off his tiny dick -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- B is for the Boobs he never sees -•-vVv-•-"
     End If
  ElseIf lAF32 = "C" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- C is for the Candlestick up his ass -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- C is for his mom's Cunt..the only one he ever saw -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- C is for the Cat's who are the only pussy he'll ever see -•-vVv-•-"
     End If
  ElseIf lAF32 = "D" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- D is for Dickhead -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- D is for the Dicks he sucks daily -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- D is for the Deodorant he fucking needs -•-vVv-•-"
     End If
  ElseIf lAF32 = "E" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- E is for his Elementary degree -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- E is for Every guy he'd blown -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- E is for the Exercise he does with his mom -•-vVv-•-"
     End If
  ElseIf lAF32 = "F" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- F is for Fuckface -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- F is for Faggot -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- F is for how he gets Fucked by guys -•-vVv-•-"
     End If
  ElseIf lAF32 = "G" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- G is for how Gay he is -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- G is for how many Guys he blew -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- G is for how he humped his Grandfather -•-vVv-•-"
     End If
  ElseIf lAF32 = "H" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- H is for his Hoe mother -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- H is for how he Humps his mom & dad -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- H is for his Homo family -•-vVv-•-"
     End If
  ElseIf lAF32 = "I" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- I is for his Ingrown dick -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- I is for the Intimate love he has with Steve Case -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- I is for Idiot -•-vVv-•-"
     End If
  ElseIf lAF32 = "J" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- J is for Jackass -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- J is for the gay Jokes he makes -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- J is for a Job he doesn't have -•-vVv-•-"
     End If
  ElseIf lAF32 = "K" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- K is for Kike -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- K is for how he tried to Kill his mom -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- K is for the way he spells (and sucks) Kock -•-vVv-•-"
     End If
  ElseIf lAF32 = "L" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- L is for the Love he gets from Steve Case -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- L is for his two Lesbian moms -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- L is for how he Loves to suck dicks  -•-vVv-•-"
     End If
  ElseIf lAF32 = "M" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- M is for the Macintosh he's using -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- M is for his gay, slutty ass Mother -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- M is for the Many dicks he sucked -•-vVv-•-"
     End If
  ElseIf lAF32 = "N" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- N is for Nests birds make in his hair -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- N is for how he Never gets any -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- N is for his New detachable penis -•-vVv-•-"
     End If
  ElseIf lAF32 = "O" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- O is for his moans when he has a dick up the ass -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- O is for how he fucks Only men -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- O is for Only hope ----> his mom -•-vVv-•-"
     End If
  ElseIf lAF32 = "P" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- P is for the pussy he never gets -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- P is for his Pathetic life -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- P is for Prick sucking faggot -•-vVv-•-"
     End If
  ElseIf lAF32 = "Q" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- Q is for how Quick he is at running away -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- Q is for how Quiet he gets when he gets beat up -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- Q is for Queer -•-vVv-•-"
     End If
  ElseIf lAF32 = "R" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- R is for the Rough sex he has with himself -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- R is for how he Remembers the dicks he sucked -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- R is for the Really hard cocks up his ass -•-vVv-•-"
     End If
  ElseIf lAF32 = "S" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- S is for Steve Case, his ex-husband -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- S is for the way he Sucks dicks -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- S is for how he fucking Stinks -•-vVv-•-"
     End If
  ElseIf lAF32 = "T" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- T is for the Tits he has -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- T is for the Tits his dad has -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- T is for the Tits he'll never see -•-vVv-•-"
     End If
  ElseIf lAF32 = "U" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- U is for Ugly as hell -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- U is for the cocks Up his ass -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- U is for his Ugly mother -•-vVv-•-"
     End If
  ElseIf lAF32 = "V" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- V is for the Vitamins he needs -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- V is for the Vagina he will never see -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- V is for porno gay Videos he watches -•-vVv-•-"
     End If
  ElseIf lAF32 = "W" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- W is for a fucking Wuss -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- W is for how Wacked this fucker is -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- W is for his wishes to get laid -•-vVv-•-"
     End If
  ElseIf lAF32 = "X" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- X is for the Xerox copies of his teachers he jerkes off to -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- X is for the gay XXX movies he watches -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- X is for the Xerox copies of his dads he jerkes off to -•-vVv-•-"
     End If
  ElseIf lAF32 = "Y" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- Y is for his Yuppie parents -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- Y is for how he says Yes to guys -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- Y is for his Yellow fucking teeth -•-vVv-•-"
     End If
  ElseIf lAF32 = "Z" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- Z is for the Zoo where his mom lives -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- Z is for what his mom looks like -> a fucking Zebra -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- Z is for Zillion times he sucked dicks -•-vVv-•-"
     End If
  ElseIf lAF32 = "0" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 0 is for how many times he got laid -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 0 is for what he is -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 0 is for how much money he has -•-vVv-•-"
     End If
  ElseIf lAF32 = "1" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 1 is for the inches of his ingrown dick -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 1 is for how many hours he cooks Minute Rice -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 1 is for the amount of brain cells he has -•-vVv-•-"
     End If
  ElseIf lAF32 = "2" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 2 is for how many dicks his mom has -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 2 is for how many clits his dad has -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 2 is for how many guys he fucks daily -•-vVv-•-"
     End If
  ElseIf lAF32 = "3" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 3 is for how many cents his mom charges a night -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 3 is for how many times he asks every guy out -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 3 is for his 3 dads -•-vVv-•-"
     End If
  ElseIf lAF32 = "4" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 4 is for how many times he gets humped by guys daily -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 4 is for how many cocks he sucks daily -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 4 is for his age -•-vVv-•-"
     End If
  ElseIf lAF32 = "5" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 5 is for how many pubic hairs he has -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 5 is for how many times he licks a dick before swallowing -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 5 is for how many hours he spends jerking off -•-vVv-•-"
     End If
  ElseIf lAF32 = "6" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 6 is for his life savings -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 6 is for how old he'll be in a few years -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 6 is for how many times he fucked Steve Case -•-vVv-•-"
     End If
  ElseIf lAF32 = "7" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 7 is for how many teachers he asked out -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 7 is for how many Macintoshes he bought -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 7 is for how many times he fucked TOSAdvisor -•-vVv-•-"
     End If
  ElseIf lAF32 = "8" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 8 is for how many times he asked his mom out -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 8 is for how many times his dad said no -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 8 is for how many times his dad said no -•-vVv-•-"
     End If
  ElseIf lAF32 = "9" Then
     If  lAF38 = 1 Then
        sub30B0 "-•-vVv-•- 9 is for the numbers he can count up to -•-vVv-•-"
     ElseIf lAF38 = 2 Then
        sub30B0 "-•-vVv-•- 9 is for how old his mom is -•-vVv-•-"
     ElseIf lAF38 = 3 Then
        sub30B0 "-•-vVv-•- 9 is for how many hours a day he watches Nickolodeon -•-vVv-•-"
     End If
  End If
  lAF34 = lAF34 + 1
  lAF30 = lAF30 + 1
  If  lAF30 = 4 Then
     lAF30 = 0
     subF38 (1.8#)
  End If
  pAF2E = Right$(pAF2E, Len(pAF2E) - 1)
Loop Until lAF34 = lAF36
End Sub

Sub subF00 (pAF3C As Integer, ByVal pAF3E As String)
Dim lAF40 As Variant
lAF40 = extfn1718(pAF3C, &HC, 0, pAF3E)
End Sub

Sub subF38 (pAF44 As Variant)
Dim lAF48 As Variant
lAF48 = Timer
Do While Timer - lAF48 < pAF44
DoEvents
Loop
End Sub

Function fnF70 (pAF4E As Variant) As String
Dim lAF52 As Variant
Dim lAF56 As String
For lAF52 = 1 To Len(pAF4E)
    If  (Mid$(pAF4E, lAF52, 1) <> Chr$(0)) Then
        lAF56$ = lAF56$ + Mid$(pAF4E, lAF52, 1)
    Else
        GoTo LA6AC
    End If
Next
LA6AC:
fnF70 = lAF56$
End Function

Function fnFA8 (pAF5A As String) As String
Dim lAF5C As Variant
Dim lAF60 As String
For lAF5C = 1 To Len(pAF5A$)
    If  (Mid$(pAF5A$, lAF5C, 1) <> Chr$(0)) Then
        lAF60$ = lAF60$ + Mid$(pAF5A$, lAF5C, 1)
    Else
        GoTo LA732
    End If
Next
LA732:
fnFA8 = lAF60$
End Function

Sub subFE0 ()
Dim lAF62 As Variant
Dim lAF66 As String
Dim lAF68 As Variant
Dim lAF6C As Variant
Dim lAF70 As Variant
Do
DoEvents
lAF62 = extfn1520("#32770", "America Online")
If  lAF66$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until lAF62 <> 0
    lAF68 = extfn24E0(lAF62, "OK")
    lAF6C = extfn1750(lAF68, gc0842, 0, 0&)
    lAF70 = extfn1750(lAF68, gc0844, 0, 0&)
End Sub

Function fn1018 (ByVal pAF76, ByVal pAF78) As Integer
Dim lAF7A As Variant
Dim lAF7E() As Integer
Do
DoEvents
lAF7A = fn5D0(lAF7E(), 0, pAF76, pAF78, Null)
Loop While lAF7A = 0
fn1018 = lAF7E(1)
End Function

Sub sub1050 ()
Dim lAF84 As Integer
Dim lAF86 As Integer
Dim lAF88 As Variant
lAF84 = extfn1168()
lAF86 = extfn1168()
Do Until lAF86 <> lAF84
lAF88 = DoEvents()
lAF86 = extfn1168()
Loop
End Sub

Sub sub1088 (pAF8C As Variant)
Dim fn10C0 As String
sub30B0 (pAF8C)
End Sub

Function fn10C0 () As String
Dim lAF96 As String * 255
Dim lAF98 As Integer
lAF98 = extfn11A0(pAF92, lAF96, 254)
 ?+lAF96?


 ?+fnF70(<concat?>)?

fn10C0 = <missed arg>
End Function

Sub sub10F8 (pAF9A As Variant, pAF9E As Variant, pAFA2 As Variant, pAFA6 As String)
Dim lAFA8 As Integer
    lAFA8 = extfn2080(pAF9A, pAF9E, pAFA2, pAFA6)
End Sub
