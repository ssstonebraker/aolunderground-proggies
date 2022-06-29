' Module1
Option Explicit
Declare Function extfn1B08 Lib "Apiguide.dll" Alias "agGetStringFromlpstr" () As String
Declare Function extfn1948 Lib "311.dll" Alias "AOLGetCombo" () As Integer
Declare Function extfn1910 Lib "311.dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn15C8 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn2010 Lib "GDI" Alias "BitBlt" () As Integer
Declare Function extfn1600 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn1590 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfnF70 Lib "User" Alias "CreateWindow" () As Integer
Declare Function extfn14E8 Lib "User" Alias "DestroyMenu" () As Integer
Declare Sub extsub1558 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn1050 Lib "Dwspydll.dll" Alias "dwGetStringFromLPSTR" () As String
Declare Function extfnEC8 Lib "User" Alias "EnableWindow" () As Integer
Declare Function extfn1248 Lib "User" Alias "EnumChildWindows" () As Integer
Declare Function extfn13D0 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn1AD0 Lib "vbwfind.dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1A98 Lib "vbwfind.dll" Alias "findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1A60 Lib "User" Alias "findwindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn12B8 Lib "User" Alias "FlashWindow" () As Integer
Declare Function extfn1408 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn11D8 Lib "User" Alias "GetClassName" () As Integer
Declare Sub extsub1210 Lib "User" Alias "GetCursorPos" ()
Declare Function extfn1FA0 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn1F68 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn1D70 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn16E0 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn1DE0 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn1E18 Lib "User" Alias "GetFreeSystemResources" () As Integer
Declare Function extfn1EC0 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn1750 Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn18A0 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1718 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn11A0 Lib "Kernel" Alias "GetModuleFileName" () As Integer
Declare Function extfn1440 Lib "Kernel" Alias "GetModuleUsage" () As Integer
Declare Function extfn18D8 Lib "User" Alias "GetNextWindow" () As Integer
Declare Function extfn1B78 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn2080 Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn20B8 Lib "Kernel" Alias "GetProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%) As Integer
Declare Function extfn1EF8 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1C90 Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn1F30 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn1D38 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn1BB0 Lib "User" Alias "GetTickCount" () As Long
Declare Function extfn1788 Lib "User" Alias "GetTopWindow" () As Integer
Declare Function extfn1D00 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn1B40 Lib "User" Alias "getwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1168 Lib "User" Alias "GetWindowLong" () As Long
Declare Sub extsub1280 Lib "User" Alias "GetWindowRect" ()
Declare Function extfn1C58 Lib "Kernel" Alias "GetWindowsDirectory" () As Integer
Declare Function extfn1C20 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn17C0 Lib "User" Alias "GetWindowTextLength" (ByVal p1%) As Integer
Declare Function extfn1130 Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn1CC8 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn1DA8 Lib "Kernel" Alias "GlobalCompact" () As Long
Declare Sub extsubFA8 Lib "User" Alias "HideCaret" ()
Declare Function extfn1520 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn10C0 Lib "Kernel" Alias "IsBadReadPtr" () As Integer
Declare Function extfn1088 Lib "Kernel" Alias "IsBadReadPtr" () As Integer
Declare Function extfn12F0 Lib "MMSYSTEM.DLL" Alias "mciGetErrorString" () As Integer
Declare Function extfn1328 Lib "MMSYSTEM.DLL" Alias "mciSendString" () As Long
Declare Function extfnF00 Lib "User" Alias "MessageBox" () As Integer
Declare Function extfn1830 Lib "User" Alias "PostMessage" () As Integer
Declare Sub extsubE90 Lib "User" Alias "ReleaseCapture" ()
Declare Sub extsub1FD8 Lib "User" Alias "ReleaseDC" ()
Declare Function extfn1638 Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn1670 Lib "User" Alias "SendDlgItemMessage" () As Long
Declare Function extfn1A28 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn1980 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Integer
Declare Function extfn19B8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Integer
Declare Function extfnFE0 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn19F0 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Sub extsubF38 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn1BE8 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn1868 Lib "User" Alias "SetFocusB" () As Integer
Declare Function extfn14B0 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn1478 Lib "User" Alias "SetParent" () As Integer
Declare Sub extsub2048 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn1398 Lib "User" Alias "ShowCursor" () As Integer
Declare Function extfn17F8 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn1E50 Lib "MMSystem" Alias "sndPlaySound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn20F0 Lib "User" Alias "SystemParametersInfo" () As Integer
Declare Function extfn16A8 Lib "vbstr.dll" Alias "VBENumChildWindows" () As Integer
Declare Function extfn1360 Lib "vbstr.dll" Alias "VBmciSendString" () As String
Declare Function extfn1E88 Lib "MMSystem" Alias "waveOutGetNumDevs" () As Integer
Declare Function extfn10F8 Lib "User" Alias "WindowFromPoint" () As Integer
Declare Function extfn1018 Lib "Kernel" Alias "WritePrivateProfileString" () As Integer

Sub subC8 (p0E16 As ListBox)
Dim l0E18 As String
l0E18$ = InputBox("Please enter what you wanna add")
If  l0E18$ = "" Then Exit Sub
Call sub170(p0E16, l0E18$)
End Sub

Sub sub100 (p0E1C As ComboBox, p0E1E As String)
Dim l0E20 As Variant
For l0E20 = 0 To p0E1C.ListCount - 1
    If  UCase$(p0E1C.List(l0E20)) = UCase$(p0E1E$) Then Exit Sub
Next
p0E1C.AddItem p0E1E$
End Sub

Sub sub138 (p0E26 As ComboBox)
Dim l0E28 As Integer
Dim l0E2A As Integer
Dim l0E2C As Variant
Dim l0E30 As Variant
Dim l0E34 As Integer
Dim l0E36 As String
Dim l0E38 As Variant
Dim l0E3C As String
l0E28% = fn528()
l0E2A% = extfn1AD0(l0E28%, "_AOL_ListBox")
l0E2C = extfn1980(l0E2A%, gc0BC0, 0, 0)
l0E30 = extfn1BE8(l0E28%)
For l0E34% = 0 To l0E2C - 1
    l0E36$ = String$(11, " ")
    l0E38 = extfn1910(l0E34%, l0E36$)
    l0E36$ = Trim$(l0E36$)
    l0E3C$ = fnDB0()
    If  Trim$(UCase$(l0E36$)) = Trim$(UCase(l0E3C$)) Then GoTo LC18
    Call sub100(p0E26, l0E36$)

LC18:
Next
End Sub

Sub sub170 (p0E40 As ListBox, p0E42 As String)
Dim l0E44 As Variant
For l0E44 = 0 To p0E40.ListCount - 1
    If  UCase$(p0E40.List(l0E44)) Like UCase$("*" + p0E42$ + "*") Then Exit Sub
Next
If  p0E42$ = "" Then Exit Sub
p0E40.AddItem p0E42$
End Sub


Sub sub1A8 (p0E48 As String, p0E4C As ListBox)
Dim l0E4E As Variant
Dim l0E52 As String
Dim l0E54 As Variant
If  p0E4C.ListCount = 0 Then
p0E4C.AddItem p0E48
Exit Sub
End If
Do Until l0E4E = (p0E4C.ListCount)
Let l0E52$ = p0E4C.List(l0E4E)
If  Trim(LCase(l0E52$)) = Trim(LCase(p0E48)) Then Let l0E54 = "NO"
Let l0E4E = l0E4E + 1
Loop
If  l0E54 = "NO" Then Exit Sub
p0E4C.AddItem p0E48
End Sub

Sub sub1E0 (p0E5A As ComboBox)
Dim l0E5C As Variant
Dim l0E60 As Integer
Dim l0E62 As Integer
Dim l0E64 As Variant
Dim l0E68 As Variant
Dim l0E6C As Integer
Dim l0E6E As String
Dim l0E70 As Variant
Dim l0E74 As Variant
l0E5C = fnC60()
If  l0E5C = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l0E60% = fn528()
If  l0E60% = 0 Then
    MsgBox "You must be in a chat room to use this feature.", 22, "No Chat Room!"
    Exit Sub
End If
l0E62% = extfn1AD0(l0E60%, "_AOL_Listbox")
l0E64 = fn6E8()
l0E68 = extfn1A28(l0E62%, gc0BC0, 0, 0)
For l0E6C% = 0 To l0E68 - 1
    l0E6E$ = String$(255, 0)
    l0E70 = extfn1910(l0E6C%, l0E6E$)
    l0E6E$ = Left$(l0E6E$, InStr(l0E6E$, Chr$(0)) - 1)
    If  UCase$(l0E6E$) = UCase$(l0E64) Then GoTo L1060
    For l0E74 = 0 To p0E5A.ListCount - 1
        If  UCase$(l0E6E$) = UCase$(p0E5A.List(l0E74)) Then
            GoTo L1060
        End If
    Next l0E74
    p0E5A.AddItem l0E6E$
L1060:
Next l0E6C%
End Sub

Sub sub218 (p0E7A As ListBox)
Dim l0E7C As Variant
Dim l0E80 As Integer
Dim l0E82 As Integer
Dim l0E84 As Variant
Dim l0E88 As Variant
Dim l0E8C As Integer
Dim l0E8E As String
Dim l0E90 As Variant
Dim l0E94 As Variant
l0E7C = fnC60()
If  l0E7C = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l0E80% = fn528()
If  l0E80% = 0 Then
    MsgBox "You must be in a chat room to use this feature.", 22, "No Chat Room!"
    Exit Sub
End If
l0E82% = extfn1AD0(l0E80%, "_AOL_Listbox")
l0E84 = fn6E8()
l0E88 = extfn1A28(l0E82%, gc0BC0, 0, 0)
For l0E8C% = 0 To l0E88 - 1
    l0E8E$ = String$(255, 0)
    l0E90 = extfn1910(l0E8C%, l0E8E$)
    l0E8E$ = Left$(l0E8E$, InStr(l0E8E$, Chr$(0)) - 1)
    If  UCase$(l0E8E$) = UCase$(l0E84) Then GoTo L12E2
    For l0E94 = 0 To p0E7A.ListCount - 1
        If  UCase$(l0E8E$) = UCase$(p0E7A.List(l0E94)) Then
            GoTo L12E2
        End If
    Next l0E94
    p0E7A.AddItem l0E8E$
L12E2:
Next l0E8C%
End Sub

Function fn250 (p0E9C As String) As Variant
Dim l0E9E As Integer
Dim l0EA0 As Integer
Dim l0EA2 As Integer
Dim l0EA4 As Integer
Dim l0EA6 As Integer
Dim l0EA8 As Integer
Dim l0EAA As Integer
Dim l0EAC As Integer
Dim l0EAE As Integer
Dim l0EB0 As Integer
Dim l0EB2 As Variant
Dim l0EB6 As String
Dim l0EB8 As Variant
Dim l0EBC As Variant
Dim l0EC0 As String
l0E9E% = extfn1A60("AOL Frame25", 0&)
subAD8 "AOL Frame25", "&Mail", "&Compose Mail"
Do
DoEvents
l0EA0% = extfn1A98(l0E9E%, "Compose Mail")
Loop Until l0EA0% <> 0
l0EA2% = extfn1AD0(l0EA0%, "_AOL_Edit")
Call subC28(l0EA2%, p0E9C$ + ", +")
l0EA4% = extfn1A98(l0EA0%, "Subject:")
l0EA6% = extfn1B40(l0EA4%, gc0990)
Call subC28(l0EA6%, "Excuse me but...")
l0EA8% = extfn1B40(l0EA6%, gc0990)
l0EA8% = extfn1B40(l0EA8%, gc0990)
l0EAA% = extfn1B40(l0EA8%, gc0990)
Call subC28(l0EAA%, "Yer account is gonna die >=o)")
l0EAC% = extfn1AD0(l0EA0%, "_AOL_Icon")
sub368 l0EAC%
Do
DoEvents
l0EAE% = extfn1A98(l0E9E%, "Error")
Loop Until l0EAE% <> 0
l0EB0% = extfn1AD0(l0EAE%, "_AOL_View")
l0EB2 = extfn1980(l0EB0%, gc073E, 0, 0)
l0EB6$ = Space(l0EB2)
If  UCase$(l0EB6$) Like UCase$("*" & p0E9C$ & "*") Then
    l0EB8 = InStr(l0EB6$, p0E9C$)
    l0EBC = InStr(l0EB6$, Chr$(13))
    l0EC0$ = Mid$(l0EB6$, l0EB8, l0EBC - 1)
    If  UCase$(l0EC0$) Like "*KNOWN*" Then fn250 = False
Else :
    fn250 = True
End If
l0EB2 = extfn1980(l0EAE%, gc0006, 0, 0)
l0EB2 = extfn1980(l0EA0%, gc0006, 0, 0)
End Function

Function fn288 () As Variant
Dim l0EC6 As Integer
Dim l0EC8 As Integer
Dim l0ECA As Integer
l0EC6% = extfn1A60("AOL Frame25", 0&)
l0EC8% = extfn1AD0(l0EC6%, "MDIClient")
l0ECA% = extfn1AD0(l0EC8%, "AOL Child")
fn288 = l0ECA%
End Function

Function fn2C0 () As Variant
Dim l0ED0 As Variant
Dim l0ED4 As Integer
Dim l0ED6 As Integer
Dim l0ED8 As Integer
Dim l0EDA As Integer
l0ED0 = 1
l0ED4% = extfn1A60("AOL Frame25", 0&)
l0ED6% = extfn1AD0(l0ED4%, "AOL Toolbar")
l0ED8% = extfn1AD0(l0ED6%, "_AOL_Icon")
Do
l0EDA% = DoEvents()
l0ED8% = extfn1B40(l0ED8%, gc0990)
l0ED0 = l0ED0 + 1
Loop Until l0ED8% = 0
Select Case l0ED0
Case 19
fn2C0 = 30
Case 21
fn2C0 = 25
End Select
End Function

Sub sub2F8 (p0EDE As Form)
On Error Resume Next
p0EDE.Top = (Screen.Height - p0EDE.Height) / 2
p0EDE.Left = (Screen.Width - p0EDE.Width) / 2
End Sub

Function fn330 () As Integer
Dim l0EE6 As Integer
Dim l0EE8 As Integer
l0EE6% = fn528()
l0EE8% = extfn1AD0(l0EE6%, "_AOL_View")
fn330 = l0EE8%
End Function

Sub sub368 (p0EEA As Integer)
Dim l0EEC As Variant
l0EEC = extfn1980(p0EEA%, gc07B8, 0, 0)
l0EEC = extfn1980(p0EEA%, gc07BA, 0, 0)
End Sub

Function fn3A0 (p0EF2 As String) As String
Dim l0EF4 As String
Dim l0EF6 As Variant
Dim l0EFA As String
Dim l0EFC As String
l0EF4$ = ""
For l0EF6 = 1 To Len(p0EF2$)
    l0EFA$ = ""
    l0EFA$ = Mid$(p0EF2$, l0EF6, 1)
    l0EFC$ = ""
    If  l0EFA$ = "a" Then l0EFC$ = "@"
    If  l0EFA$ = "b" Then l0EFC$ = "þ"
    If  l0EFA$ = "c" Then l0EFC$ = "©"
    If  l0EFA$ = "d" Then l0EFC$ = "d"
    If  l0EFA$ = "e" Then l0EFC$ = "ë"
    If  l0EFA$ = "f" Then l0EFC$ = "ƒ"
    If  l0EFA$ = "g" Then l0EFC$ = "g"
    If  l0EFA$ = "h" Then l0EFC$ = "h"
    If  l0EFA$ = "i" Then l0EFC$ = "ï"
    If  l0EFA$ = "j" Then l0EFC$ = "j"
    If  l0EFA$ = "k" Then l0EFC$ = "k"
    If  l0EFA$ = "l" Then l0EFC$ = "l"
    If  l0EFA$ = "m" Then l0EFC$ = "m"
    If  l0EFA$ = "n" Then l0EFC$ = "ñ"
    If  l0EFA$ = "o" Then l0EFC$ = "ø"
    If  l0EFA$ = "p" Then l0EFC$ = "p"
    If  l0EFA$ = "q" Then l0EFC$ = "q"
    If  l0EFA$ = "r" Then l0EFC$ = "®"
    If  l0EFA$ = "s" Then l0EFC$ = "$"
    If  l0EFA$ = "t" Then l0EFC$ = "†"
    If  l0EFA$ = "u" Then l0EFC$ = "ü"
    If  l0EFA$ = "v" Then l0EFC$ = "v"
    If  l0EFA$ = "w" Then l0EFC$ = "vv"
    If  l0EFA$ = "x" Then l0EFC$ = "×"
    If  l0EFA$ = "y" Then l0EFC$ = "ý"
    If  l0EFA$ = "z" Then l0EFC$ = "z"
    If  l0EFA$ = " " Then l0EFC$ = " "
    If  l0EFA$ = "!" Then l0EFC$ = "¡"
    If  l0EFA$ = "1" Then l0EFC$ = "1"
    If  l0EFA$ = "2" Then l0EFC$ = "2"
    If  l0EFA$ = "3" Then l0EFC$ = "3"
    If  l0EFA$ = "4" Then l0EFC$ = "4"
    If  l0EFA$ = "5" Then l0EFC$ = "5"
    If  l0EFA$ = "6" Then l0EFC$ = "6"
    If  l0EFA$ = "7" Then l0EFC$ = "7"
    If  l0EFA$ = "8" Then l0EFC$ = "8"
    If  l0EFA$ = "9" Then l0EFC$ = "9"
    If  l0EFA$ = "0" Then l0EFC$ = "0"
    If  l0EFA$ = "A" Then l0EFC$ = "Å"
    If  l0EFA$ = "B" Then l0EFC$ = "ß"
    If  l0EFA$ = "C" Then l0EFC$ = "Ç"
    If  l0EFA$ = "D" Then l0EFC$ = "Ð"
    If  l0EFA$ = "E" Then l0EFC$ = "Ë"
    If  l0EFA$ = "F" Then l0EFC$ = "ƒ"
    If  l0EFA$ = "G" Then l0EFC$ = "G"
    If  l0EFA$ = "H" Then l0EFC$ = "/-/"
    If  l0EFA$ = "I" Then l0EFC$ = "‡"
    If  l0EFA$ = "J" Then l0EFC$ = "J"
    If  l0EFA$ = "K" Then l0EFC$ = "]<"
    If  l0EFA$ = "L" Then l0EFC$ = "£"
    If  l0EFA$ = "M" Then l0EFC$ = "/\/\"
    If  l0EFA$ = "N" Then l0EFC$ = "/\/"
    If  l0EFA$ = "O" Then l0EFC$ = "Õ"
    If  l0EFA$ = "P" Then l0EFC$ = "¶"
    If  l0EFA$ = "Q" Then l0EFC$ = "`Q"
    If  l0EFA$ = "R" Then l0EFC$ = "•R•"
    If  l0EFA$ = "S" Then l0EFC$ = "§"
    If  l0EFA$ = "T" Then l0EFC$ = "†"
    If  l0EFA$ = "U" Then l0EFC$ = "Ü"
    If  l0EFA$ = "V" Then l0EFC$ = "\/"
    If  l0EFA$ = "W" Then l0EFC$ = "\\'"
    If  l0EFA$ = "X" Then l0EFC$ = "><"
    If  l0EFA$ = "Y" Then l0EFC$ = "¥"
    If  l0EFA$ = "Z" Then l0EFC$ = "Z"
    If  l0EFA$ = "~" Then l0EFC$ = "~"
    If  l0EFA$ = "`" Then l0EFC$ = "`"
    If  l0EFA$ = "!" Then l0EFC$ = "¡"
    If  l0EFA$ = "@" Then l0EFC$ = "ä"
    If  l0EFA$ = "#" Then l0EFC$ = "#"
    If  l0EFA$ = "$" Then l0EFC$ = "$"
    If  l0EFA$ = "%" Then l0EFC$ = "%"
    If  l0EFA$ = "^" Then l0EFC$ = "^"
    If  l0EFA$ = "&" Then l0EFC$ = "&"
    If  l0EFA$ = "*" Then l0EFC$ = "™"
    If  l0EFA$ = "(" Then l0EFC$ = "("
    If  l0EFA$ = ")" Then l0EFC$ = ")"
    If  l0EFA$ = "-" Then l0EFC$ = "-"
    If  l0EFA$ = "_" Then l0EFC$ = "_"
    If  l0EFA$ = "+" Then l0EFC$ = "+"
    If  l0EFA$ = "=" Then l0EFC$ = "="
    If  l0EFA$ = "[" Then l0EFC$ = "["
    If  l0EFA$ = "]" Then l0EFC$ = "]"
    If  l0EFA$ = "{" Then l0EFC$ = "{"
    If  l0EFA$ = "}" Then l0EFC$ = "}"
    If  l0EFA$ = ":" Then l0EFC$ = ":"
    If  l0EFA$ = ";" Then l0EFC$ = ";"
    If  l0EFA$ = "'" Then l0EFC$ = "'"
    If  l0EFA$ = "," Then l0EFC$ = ","
    If  l0EFA$ = "." Then l0EFC$ = "."
    If  l0EFA$ = "<" Then l0EFC$ = "<"
    If  l0EFA$ = ">" Then l0EFC$ = ">"
    If  l0EFA$ = "?" Then l0EFC$ = "¿"
    If  Len(l0EFC$) = 0 Then l0EFC$ = l0EFA$
    l0EF4$ = l0EF4$ & l0EFC$
Next
fn3A0 = l0EF4$
End Function

Function fn3D8 (p0F00 As String) As String
Dim l0F02 As Variant
Dim l0F06 As String
Dim l0F08 As Variant
Dim l0F0C As Variant
Dim l0F10 As Variant
Dim l0F14 As String
        l0F02 = Len(p0F00$)
        l0F06$ = p0F00$
        p0F00$ = ""
        l0F08 = 1
        l0F0C = True
        For l0F10 = 1 To l0F02
            l0F14$ = Mid$(l0F06$, l0F08, 1)
            If  l0F0C = True Then
                    If  UCase$(l0F14$) = "I" Then
                        p0F00$ = p0F00$ & "i"
                    ElseIf UCase$(l0F14$) = "O" Then
                        p0F00$ = p0F00$ & "0"
                    ElseIf UCase$(l0F14$) = "E" Then
                        p0F00$ = p0F00$ & "3"
                    ElseIf UCase$(l0F14$) = "A" Then
                        p0F00$ = p0F00$ & "4"
                    Else :
                        p0F00$ = p0F00$ & UCase$(l0F14$)
                    End If
                l0F0C = False
                ElseIf l0F0C = False Then
                If  UCase$(l0F14$) = "L" Then
                        p0F00$ = p0F00$ & "L"
                    ElseIf UCase$(l0F14$) = "O" Then
                        p0F00$ = p0F00$ & "0"
                    ElseIf UCase$(l0F14$) = "E" Then
                        p0F00$ = p0F00$ & "3"
                    ElseIf UCase$(l0F14$) = "A" Then
                        p0F00$ = p0F00$ & "4"
                    Else :
                        p0F00$ = p0F00$ & LCase$(l0F14$)
                    End If
                l0F0C = True
            End If
            l0F08 = l0F08 + 1
        Next
fn3D8 = p0F00$
End Function

Sub sub410 ()
Dim l0F16 As Integer
Dim l0F18 As Integer
Dim l0F1A As Integer
Dim l0F1C As Integer
Dim l0F1E As Integer
Dim l0F20 As Variant
Dim l0F24 As Integer
Dim l0F26 As Integer
Dim l0F28 As Variant
Dim l0F2C As Variant
l0F16% = extfn1A60("AOL Frame25", 0&)
l0F18% = extfn1AD0(l0F16%, "MDIClient")
l0F1A% = extfn1A98(l0F18%, "New Mail")
If  l0F1A% <> 0 Then GoTo L2D52
l0F1C% = extfn1AD0(l0F16%, "AOL Toolbar")
l0F1E% = extfn1AD0(l0F1C%, "_AOL_Icon")
l0F20 = extfn1980(l0F1E%, gc07B8, 0, 0)
l0F20 = extfn1980(l0F1E%, gc07BA, 0, 0)
Do
l0F24% = DoEvents()
l0F1A% = extfn1A98(l0F18%, "New Mail")
Loop Until l0F1A% <> 0
l0F26% = extfn1AD0(l0F1A%, "_AOL_Tree")
Do
l0F28 = extfn1980(l0F26%, gc0BC0, 0, 0)
subD08 (2)
l0F2C = extfn1980(l0F26%, gc0BC0, 0, 0)
Loop Until l0F28 = l0F2C
L2D52:
l0F26% = extfn1AD0(l0F1A%, "_AOL_Tree")
l0F2C = extfn1980(l0F26%, gc0BC0, 0, 0)
MsgBox "You have " & l0F2C & " pieces of mail!", 64, "Mail Counter"
End Sub

Sub sub448 (p0F30 As Integer)
Dim l0F32 As Variant
l0F32 = extfn1980(p0F30%, gc0790, 13, 0)
End Sub

Sub sub480 ()
Dim l0F36 As Variant
Dim l0F3A As Integer
l0F36 = MsgBox("You sure you wanna exit CoP KiLLa by Re/X\/X\i?", 68)
If  l0F36 = 6 Then
    l0F3A% = extfn1A60("AOL Frame25", 0&)
    End
End If
End Sub

Function fn4B8 (p0F40 As String, p0F42 As String) As Variant
Dim l0F44 As Variant
Dim l0F48 As String
Dim l0F4A As String
If  InStr(p0F40$, ".") Then
    l0F44 = InStr(p0F40$, ".")
    l0F48$ = Mid$(p0F40$, 1, l0F44 - 1)
    l0F4A$ = Mid$(p0F40$, l0F44 + 2)
    If  l0F4A$ <> p0F42$ Then l0F4A$ = p0F42$
    p0F40$ = l0F48$ + "." + l0F4A$
ElseIf InStr(p0F40$, ".") = False Then
    p0F40$ = p0F40$ + "." + p0F42$
End If
End Function

Function fn4F0 (p0F4E As String) As Integer
Dim l0F50 As Integer
On Error Resume Next
l0F50 = Len(Dir$(p0F4E))
If  Err Or l0F50 = 0 Then
    fn4F0 = False
    fn4F0 = 0
Else
    fn4F0 = True
    fn4F0 = 1
End If
End Function

Function fn528 () As Integer
Dim l0F54 As Integer
Dim l0F56 As Integer
Dim l0F58 As Integer
Dim l0F5A As Integer
Dim l0F5C As Integer
Dim l0F5E As Integer
l0F54% = extfn1A60("AOL Frame25", 0&)
l0F56% = extfn1AD0(l0F54%, "_AOL_ListBox")
l0F58% = extfn1B78(l0F56%)
l0F5A% = extfn1AD0(l0F58%, "_AOL_View")
l0F5C% = extfn1AD0(l0F58%, "_AOL_Edit")
If  l0F5C% <> 0 Then
l0F5E% = extfn1B78(l0F5A%)
fn528 = l0F5E%
End If
End Function

Function fn560 () As Variant
Dim l0F64 As Integer
Dim l0F66 As Integer
Dim l0F68 As Integer
Dim l0F6E As Variant
Dim l0F72 As Variant
Dim l0F76 As Variant
Dim l0F7A As Variant
Dim l0F7E As Variant
l0F7E = 0
l0F64% = extfn1A60("AOL Frame25", 0&)
l0F66% = extfn1AD0(l0F64%, "MDIClient")
l0F68% = extfn1AD0(l0F66%, "AOL Child")
l0F6E = extfn1AD0(l0F68%, "_AOL_Listbox")
If  l0F6E = 0 Then
    GoTo L337C
End If
l0F72 = extfn1AD0(l0F68%, "_AOL_View")
If  l0F72 = 0 Then
    GoTo L337C
End If
l0F76 = extfn1AD0(l0F68%, "_AOL_Edit")
If  l0F76 = 0 Then
    GoTo L337C
End If
l0F7A = extfn1AD0(l0F68%, "_AOL_Icon")
If  l0F7A = 0 Then
    GoTo L337C
End If
fn560 = l0F68%
Exit Function
L337C:
l0F7E = l0F7E + 1
If  l0F7E = 50 Then
    fn560 = 0
    Exit Function
End If
l0F68% = extfn1B40(l0F68%, gc0990)
l0F6E = extfn1AD0(l0F68%, "_AOL_Listbox")
If  l0F6E = 0 Then
    GoTo L337C
End If
l0F72 = extfn1AD0(l0F68%, "_AOL_View")
If  l0F72 = 0 Then
    GoTo L337C
End If
l0F76 = extfn1AD0(l0F68%, "_AOL_Edit")
If  l0F76 = 0 Then
    GoTo L337C
End If
l0F7A = extfn1AD0(l0F68%, "_AOL_Icon")
If  l0F7A = 0 Then
    GoTo L337C
End If
fn560 = l0F68%
End Function

Function fn598 () As Variant
Dim l0F86 As Integer
Dim l0F88 As Integer
Dim l0F8A As Integer
Dim l0F8C As Integer
Dim l0F8E As Integer
l0F86% = extfn1A60("AOL Frame25", 0&)
l0F88% = extfn1AD0(l0F86%, "MDIClient")
l0F8A% = extfn1AD0(l0F88%, "AOL Child")
l0F8C% = extfn1A98(l0F8A%, "Reply")
l0F8E% = extfn1A98(l0F8A%, "Send Now")
If  l0F8C% = 0 And l0F8E% = 0 Then
    Do: DoEvents
        l0F8A% = extfn1B40(l0F8A%, gc0990)
        l0F8C% = extfn1A98(l0F8A%, "Reply")
        l0F8E% = extfn1A98(l0F8A%, "Send Now")
    Loop Until l0F8C% <> 0 Or l0F8E% <> 0 Or l0F8A% = 0
    If  l0F8A% = 0 Then
        fn598 = 0
        Exit Function
    End If
End If
fn598 = l0F8A%
End Function

Function fn5D0 (ByVal p0F92 As String) As String
On Error Resume Next
fn5D0 = Trim(Left$(p0F92, InStr(p0F92, Chr$(0)) - 1))
End Function

Function fn608 () As String
Dim l0F96 As Variant
Dim l0F9A As Variant
Dim l0F9E As String
Randomize
    l0F96 = 7
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = "Hello, "
      If  l0F9A = 2 Then l0F9E$ = "Good evening, "
      If  l0F9A = 3 Then l0F9E$ = "Hello, And welcome to America Online. "
      If  l0F9A = 4 Then l0F9E$ = "Welcome to America Online. "
      If  l0F9A = 5 Then l0F9E$ = "Excuse me, "
      If  l0F9A = 6 Then l0F9E$ = "Dear User, "
      If  l0F9A = 7 Then l0F9E$ = "What's Up! "
    l0F96 = 7
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & "I am with America Online Billing "
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & "I am with OTC (Online Techincal Consultants) "
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & "I am with WWWC (World Wide Web Consultants) "
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & "I am with AOL Techincal Staff "
      If  l0F9A = 5 Then l0F9E$ = l0F9E$ & "I am with AOL System Security "
      If  l0F9A = 6 Then l0F9E$ = l0F9E$ & "I am with AOL Resource Department "
      If  l0F9A = 7 Then l0F9E$ = l0F9E$ & "I am with the AOL Community Action Team "
    l0F96 = 7
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & "and due to Technical Failures "
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & "and due to Billing errors "
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & "and due to Line Noise "
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & "and due to A virus in our database "
      If  l0F9A = 5 Then l0F9E$ = l0F9E$ & "and due to Massive data flow in our database "
      If  l0F9A = 6 Then l0F9E$ = l0F9E$ & "and due to data corruption "
      If  l0F9A = 7 Then l0F9E$ = l0F9E$ & "and due to hackers by-passing out systems "
    l0F96 = 5
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & "we seem to have lost your password. "
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & "we seem to have lost your account information. "
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & "we seem to have failed to recieve your logon password. "
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & "we seem to have failed to recieve your account information. "
      If  l0F9A = 5 Then l0F9E$ = l0F9E$ & "we have lost your Credit Card information. "
    l0F96 = 4
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & "To correct this situation please click respond and enter "
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & "Please click respond and enter your "
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & "Please help us by entering your "
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & "Please respond with your "
    l0F96 = 3
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & "Your password information."
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & "Your current account password."
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    l0F96 = 4
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & " Please respond within 2 minutes too keep this account active."
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & " It is very important that you respond immediately."
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & " Please respond as soon as possible "
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & " Incooperation may lead to termination of your account."
    l0F96 = 4
    l0F9A = Int(Rnd * l0F96 + 1)
      If  l0F9A = 1 Then l0F9E$ = l0F9E$ & " Thank you for using America Online."
      If  l0F9A = 2 Then l0F9E$ = l0F9E$ & " Thank you for your time."
      If  l0F9A = 3 Then l0F9E$ = l0F9E$ & " Thank you for your cooperation."
      If  l0F9A = 4 Then l0F9E$ = l0F9E$ & " Thank you for your help and enjoy the service!"
fn608 = l0F9E$
End Function

Function fn640 () As String
Dim l0FA2 As Integer
Dim l0FA4 As Variant
Dim l0FA8 As String
Dim l0FAA As Variant
Dim l0FAE As String
l0FA2% = fn330()
If  l0FA2% = 0 Then Exit Function
l0FA4 = extfn1980(l0FA2%, gc073E, 0, 0)
l0FA8$ = Space(l0FA4)
l0FA8$ = Trim$(l0FA8$)
For l0FAA = Len(l0FA8$) To 1 Step -1
    l0FAE$ = Mid$(l0FA8$, l0FAA)
    l0FAE$ = Mid$(l0FAE$, 1, 1)
    If  l0FAE$ = Chr$(13) Then
        l0FA8$ = Mid$(l0FA8$, l0FAA, Len(l0FA8$))
        Exit For
    End If
Next
fn640 = l0FA8$
End Function

Function fn678 (p0FB2 As String, p0FB4 As String, p0FB6 As String) As String
Dim l0FB8 As String
   l0FB8 = String(255, Chr(0))
   fn678 = Left(l0FB8, extfn2080(p0FB2$, ByVal p0FB4$, "", l0FB8, Len(l0FB8), p0FB6$))
End Function

Function fn6B0 () As String
Dim l0FBC As Integer
Dim l0FBE As Variant
Dim l0FC2 As String
On Error Resume Next
l0FBC% = fn528()
l0FBE = extfn17C0(l0FBC%)
l0FC2$ = Space(l0FBE + 1)
l0FBE = extfn1C20(l0FBC%, l0FC2$, l0FBE + 1)
l0FC2$ = fn5D0(l0FC2$)
fn6B0 = l0FC2$
End Function

Function fn6E8 () As Variant
Dim l0FC8 As Integer
Dim l0FCA As Integer
Dim l0FCC As Integer
Dim l0FD0 As String * 255
Dim l0FD2 As String
Dim l0FD4 As Variant
l0FCA% = extfn1A60("AOL Frame25", 0&)
l0FC8% = extfn1AD0(l0FCA%, "MDIClient")
l0FCC% = extfn1A98(l0FC8%, "Welcome, ")
l0FD4 = extfn1C20(l0FCC%, l0FD0, 255)
l0FD0 = LTrim(RTrim(Trim(l0FD0)))
l0FD0 = Right$(l0FD0, Len(l0FD0) - InStr(l0FD0, ", "))
l0FD2 = Left$(l0FD0, 10)
l0FD2 = Right$(l0FD0, Len(l0FD0) - 1)
If  InStr(l0FD2, "!") <> 0 Then l0FD2 = Left$(l0FD2, InStr(l0FD2, "!") - 1)
fn6E8 = l0FD2$
End Function

Function fn720 () As Variant
Dim l0FDC As Integer
Dim l0FDE As Integer
Dim l0FE0 As Variant
Dim l0FE4 As Integer
Dim l0FE6 As Integer
Dim l0FE8 As Integer
Dim l0FEA As Integer
Dim l0FEC As String
l0FDE% = extfn1A60("AOL Frame25", 0&)
l0FDC = extfn1A98(l0FDE%, "Welcome, ")
If  l0FDC = 0 Then l0FDC = extfn1A98(l0FDE%, "Welcome to NEA Online,")
l0FE0 = fnE58(l0FDC)
If  InStr(l0FE0, "Welcome,") Or InStr(l0FE0, "Welcome to NEA Online,") Then
    l0FE6% = l0FE4%
    l0FE8% = InStr(l0FE0, ", ")
    l0FEA% = InStr(l0FE0, "!")
    On Error Resume Next
    l0FEC$ = Trim$(Mid$(l0FE0, l0FE8% + 2, l0FEA% - (l0FE8% + 2)))
fn720 = l0FEC$
End If
End Function


Function fn758 (p0FF2 As Variant, p0FF6 As Variant) As Variant
Dim l0FFA As String
Dim l0FFE As Integer
    l0FFA = String$(255, 0)
    l0FFE = extfn20B8(p0FF2, p0FF6, "", l0FFA, Len(l0FFA))
    If  l0FFE = 0 Then
        fn758 = "unknown"
    Else
        fn758 = Left(l0FFA, l0FFE)
    End If
End Function

Sub sub790 (ByVal p1000 As String)
Dim l1002 As Integer
Dim l1004 As Integer
Dim l1006 As Integer
Dim l1008 As Integer
Dim l100A As Integer
Dim l100C As Integer
Dim l100E As Integer
Dim l1010 As String
Dim l1012 As Integer
Dim l1014 As Integer
l1002% = extfn1A60("AOL Frame25", 0&)
l1004% = extfn1EC0(l1002%)
l1006% = extfn1750(l1004%)
For l1008% = 0 To l1006% - 1
    l100A% = extfn1EF8(l1004%, l1008%)
    l100C% = extfn1750(l100A%)
    For l100E% = 0 To l100C% - 1
        l1010$ = Space$(50)
        l1012% = extfn1718(l100A%, l100E%, l1010$, 50, gc00BA)
        l1010$ = Trim$(l1010$): l1010$ = Left(l1010$, Len(l1010$) - 1)
        If  l1010$ = "" Then l1010$ = " -"
        If  InStr(l1010$, p1000$) Then
            l1014% = extfn18A0(l100A%, l100E%)
            l1012% = extfn1980(l1002%, gc07A0, l1014%, 0)
        End If
    Next l100E%
Next l1008%
End Sub

Sub sub7C8 (p1016 As String)
Dim l1018 As Integer
Dim l101A As Integer
Dim l101C As Integer
Dim l101E As Integer
Dim l1020 As Integer
Dim l1022 As Integer
Dim l1024 As Integer
Dim l1026 As String
Dim l1028 As Integer
Dim l102A As Integer
l1018% = extfn1A60("AOL Frame25", 0&)
l101A% = extfn1EC0(l1018%)
l101C% = extfn1750(l101A%)
For l101E% = 0 To l101C% - 1
    l1020% = extfn1EF8(l101A%, l101E%)
    l1022% = extfn1750(l1020%)
    For l1024% = 0 To l1022% - 1
        l1026$ = Space$(50)
        l1028% = extfn1718(l1020%, l1024%, l1026$, 50, gc00BA)
        l1026$ = Trim$(l1026$): l1026$ = Left(l1026$, Len(l1026$) - 1)
        If  l1026$ = "" Then l1026$ = " -"
        If  InStr(l1026$, p1016$) Then
            l102A% = extfn18A0(l1020%, l1024%)
            l1028% = extfn1980(l1018%, gc07A0, l102A%, 0)
        End If
    Next l1024%
Next l101E%
End Sub

Sub sub800 (ByVal p102C As String)
Dim l102E As Integer
Dim l1030 As Integer
Dim l1032 As Integer
Dim l1034 As Integer
Dim l1036 As Integer
Dim l1038 As Variant
Call sub790("Send an Instant Message")
l102E% = extfn1A60("AOL FRAME25", 0&)
Do
DoEvents
l1030% = extfn1A98(l102E%, "Send Instant Message")
Loop Until l1030% <> 0
l1032% = extfn1AD0(l1030%, "_AOL_EDIT")
Call subC28(l1032%, "$IM_OFF")
l1034% = extfn1A98(l1030%, "Send")
l1036% = extfn1B40(l1034%, gc0992)
Call subC28(l1036%, p102C$)
sub368 (l1034%)
subE20
l1038 = extfn1980(l1030%, gc0006, 0, 0&)
End Sub

Sub sub838 (ByVal p103C As String)
Dim l103E As Integer
Dim l1040 As Integer
Dim l1042 As Integer
Dim l1044 As Integer
Dim l1046 As Integer
Dim l1048 As Variant
Call sub790("Send an Instant Message")
l103E% = extfn1A60("AOL FRAME25", 0&)
Do
DoEvents
l1040% = extfn1A98(l103E%, "Send Instant Message")
Loop Until l1040% <> 0
l1042% = extfn1AD0(l1040%, "_AOL_EDIT")
Call subC28(l1042%, "$IM_ON")
l1044% = extfn1A98(l1040%, "Send")
l1046% = extfn1B40(l1044%, gc0992)
Call subC28(l1046%, p103C$)
sub368 (l1044%)
subE20
l1048 = extfn1980(l1040%, gc0006, 0, 0&)
End Sub

Sub sub870 (ByVal p104C As String)
Dim l104E As Integer
Dim l1050 As Integer
Dim l1052 As Integer
l104E% = extfn1A60("AOL Frame25", 0&)
sub790 ("Keyword...")
Do
DoEvents
l1050% = extfn1A98(l104E%, "Keyword")
DoEvents
Loop Until l1050% <> 0
DoEvents
l1052% = extfn1AD0(l1050%, "_AOL_Edit")
Call subC28(l1052%, p104C$)
sub448 (l1052%)
End Sub

Sub sub8A8 ()
Dim l1054 As Integer
Dim l1056 As Variant
Dim l105A As Integer
Dim l1060 As Integer
Dim l1062 As Variant
Dim l1066 As Variant
Dim l106A As Variant
l1066 = True
l1054% = extfn1A60("AOL Frame25", 0&)
If  l1054% = 0 Then Exit Sub
l1056 = extfn1980(l1054%, gc0634, 0, &H0)
If  fn2C0() = 25 Then
    Do
    DoEvents
        l105A% = extfn1A60("_AOL_Modal", 0&)
    Loop Until l105A% <> 0
    l1060% = extfn1A98(l105A%, "&No")
    l1062 = extfn1980(l1060%, &H201, 0, 0&)
    l1062 = extfn1980(l1060%, &H202, 0, 0&)
    DoEvents
    l1066 = False
    Exit Sub
End If
If  fn2C0() = 30 Then
    Do
    DoEvents
        l106A = extfn1A98(l1054%, "Exit?")
    Loop Until l106A <> 0
    l1056 = extfn1A28(l106A, gc0006, 0, 0)
    DoEvents
    l1066 = False
End If
End Sub

Function fn8E0 (p1070 As String, p1072 As String) As String
Dim l1074 As String
Dim l1076 As Variant
Dim l107A As String
On Error GoTo L55F6
Open p1072$ For Input As #1
Do While Not EOF(1)
    Input #1, l1074$
    If  InStr(UCase$(l1074$), UCase$(p1070$)) Then
        l1076 = InStr(UCase$(l1074$), UCase$(p1070$))
        l107A$ = Mid$(l1074$, l1076 + Len(p1070$))
        fn8E0 = l107A$
    End If
Loop
L55F6:
Close #1
Resume L560C
L560C:
End Function

Sub sub918 ()
Dim l107C As Integer
l107C% = extfn1A60("AOL Frame25", 0&)
subAD8 "AOL Frame25", "&Go To", "Lobby"
Do
DoEvents
Loop Until fn528() <> 0
End Sub

Sub sub950 (p107E As Variant)
Dim l1082 As Variant
Dim l108A As Variant
Dim l108E As Variant
Dim l1092 As Variant
Dim l1096 As Variant
Dim l109A As Variant
Dim l109E As Variant
Dim l10A2 As Variant
Dim l10A6 As Variant
l1082 = extfn1A60("AOL FRAME25", 0&)
If  fn2C0() = 25 Then Call sub790("Set Preferences")
If  fn2C0() = 30 Then Call sub790("Preferences")
Do
DoEvents
l108A = extfn1A98(l1082, "Preferences")
Loop Until l108A <> 0
l108E = extfn1A98(l108A, "Mail")
l1092 = extfn1B40(l108E, gc0990)
l1096 = extfn1A28(l1092, gc07B8, 0, 0)
l1096 = extfn1A28(l1092, gc07BA, 0, 0)
Do
DoEvents
l109A = extfn1A60("_AOL_MODAL", "Mail Preferences")
Loop Until l109A <> 0
l109E = extfn1AD0(l109A, "_AOL_BUTTON")
l10A2 = extfn1B40(l109E, gc0990)
l10A6 = extfn1A98(l109A, "OK")
If  UCase$(p107E) = "MM" Then
    l1096 = extfn1980(l109E, gc0790, 43, 0)
    l1096 = extfn1980(l10A2, gc0790, 43, 0)
End If
If  UCase$(p107E) = "MB" Then
    l1096 = extfn1980(l109E, gc0790, 43, 0)
    l1096 = extfn1980(l10A2, gc0790, 45, 0)
End If
l1096 = extfn1A28(l10A6, gc078C, gc063E, 0)
l1096 = extfn1A28(l10A6, gc078E, gc063E, 0)
l1096 = extfn1A28(l108A, gc0006, 0, 0)
End Sub

Sub sub988 (p10AA As String, p10AC As String, p10AE As String)
Dim l10B0 As Integer
Dim l10B2 As String
Dim l10B4 As Integer
Dim l10B6 As String
Dim l10B8 As Integer
Dim l10BA As Integer
Dim l10BC As Integer
Dim l10BE As Integer
Dim l10C2 As Integer
Dim l10C4 As Integer
On Error GoTo L5C94
l10B0 = -1
l10C2% = extfn1A60(p10AA, 0&)
l10C4 = extfn1EC0(l10C2%)
Do
    DoEvents
    l10B0 = l10B0 + 1
    l10B2$ = String$(255, 0)
    l10B4% = extfn1718(l10C4, l10B0, l10B2$, Len(p10AC) + 1, &H400)
    l10B6 = fnD40(l10B2$)
    If  l10B6 = p10AC Then Exit Do
Loop
l10B8 = extfn1EF8(l10C4, l10B0)
l10BA = -1
Do
    DoEvents
    l10BA = l10BA + 1
    l10B2$ = String(255, 0)
    l10B4% = extfn1718(l10B8, l10BA, l10B2$, Len(p10AE) + 1, &H400)
    l10B6 = fnD40(l10B2$)
    If  l10B6 = p10AE Then Exit Do
Loop
DoEvents
l10BC% = extfn18A0(l10B8, l10BA)
l10BE = extfn1980(l10C2, &H111, l10BC%, 0&)
L5C94:
Resume Next
End Sub

Function fn9C0 (p10C8 As Integer, p10CA As Integer, p10CC As Integer) As Integer
Dim l10CE As Variant
l10CE = Int(p10C8 / p10CA * p10CC)
fn9C0 = l10CE
End Function

Sub sub9F8 (p10D2 As Variant)
Dim l10D6 As Variant
l10D6 = extfn1E50(ByVal CStr(p10D2), 1)
End Sub

Sub subA30 (p10DC As ListBox)
If  p10DC = "" Then Exit Sub
p10DC.RemoveItem p10DC.ListIndex
End Sub

Function fnA68 (p10E0 As String, p10E2 As String, p10E4 As String) As String
Dim l10E6 As String
Dim l10E8 As String
Dim l10EA As String
L5D5C:
If  InStr(UCase$(p10E0$), UCase$(p10E2$)) Then
        l10E6$ = Left$(p10E0$, InStr(UCase$(p10E0$), UCase$(p10E2$)) - 1)
        l10E8$ = Mid$(p10E0$, InStr(UCase$(p10E0$), UCase$(p10E2$)) + Len(p10E2$), Len(p10E0$))
        l10EA$ = l10E6$ & p10E4$ & l10E8$
    Else :
        l10EA$ = p10E0$
End If
fnA68 = l10EA$
End Function

Sub subAA0 (p10EC As String, p10EE As String, p10F0 As String)
Dim l10F2 As Variant
Dim l10F6 As String
Dim l10F8 As Variant
Dim l10FC As Variant
Dim l1100 As Variant
Dim l1104 As Variant
Dim l1108 As String
Dim l110A As Single
DoEvents
On Error GoTo L6258
l10F2 = Len(p10EC$)
If  l10F2 = 1 Then p10EC$ = p10EC$ + "         "
If  l10F2 = 2 Then p10EC$ = p10EC$ + "        "
If  l10F2 = 3 Then p10EC$ = p10EC$ + "       "
If  l10F2 = 4 Then p10EC$ = p10EC$ + "      "
If  l10F2 = 5 Then p10EC$ = p10EC$ + "     "
If  l10F2 = 6 Then p10EC$ = p10EC$ + "    "
If  l10F2 = 7 Then p10EC$ = p10EC$ + "   "
If  l10F2 = 8 Then p10EC$ = p10EC$ + "  "
If  l10F2 = 9 Then p10EC$ = p10EC$ + " "
l10F2 = Len(p10F0$)
If  l10F2 = 1 Then p10F0$ = p10F0$ + "         "
If  l10F2 = 2 Then p10F0$ = p10F0$ + "        "
If  l10F2 = 3 Then p10F0$ = p10F0$ + "       "
If  l10F2 = 4 Then p10F0$ = p10F0$ + "      "
If  l10F2 = 5 Then p10F0$ = p10F0$ + "     "
If  l10F2 = 6 Then p10F0$ = p10F0$ + "    "
If  l10F2 = 7 Then p10F0$ = p10F0$ + "   "
If  l10F2 = 8 Then p10F0$ = p10F0$ + "  "
If  l10F2 = 9 Then p10F0$ = p10F0$ + " "
l10F6$ = p10EC$
Open p10EE$ For Binary Access Read Write As #2
l10F8 = LOF(2)
l10FC = l10F8
l1100 = 1
While l10FC >= 0
If  l10FC > 32000 Then
        l1104 = 32000
    ElseIf l10FC = 0 Then
        l1104 = 1
    Else
        l1104 = l10FC
    End If
    l1108$ = String$(l1104, " ")
    Get #2, l1100, l1108$
    l110A! = InStr(1, l1108$, l10F6$, 1)
    If  l110A! Then Mid$(l1108$, l110A!) = p10F0$
    Put #2, l1100, l1108$
    l1100 = l1100 + l1104
    l10FC = l10F8 - l1100
    Wend
Close
Exit Sub
L6258:
Resume L6264
L6264:
End Sub

Sub subAD8 (ByVal p110C As String, ByVal p110E As String, ByVal p1110 As String)
Dim l1112 As Integer
Dim l1114 As String
Dim l1116 As Integer
Dim l1118 As String
Dim l111A As Integer
Dim l111C As Integer
Dim l111E As Integer
Dim l1120 As Integer
Dim l1124 As Integer
Dim l1126 As Integer
On Error GoTo L6438
l1112 = -1
l1124% = extfn1A60(p110C, 0&)
l1126 = extfn1EC0(l1124%)
Do
    DoEvents
    l1112 = l1112 + 1
    l1114$ = String$(255, 0)
    l1116% = extfn1718(l1126, l1112, l1114$, Len(p110E) + 1, &H400)
    l1118 = fnD40(l1114$)
    If  l1118 = p110E Then Exit Do
Loop
l111A = extfn1EF8(l1126, l1112)
l111C = -1
Do
    DoEvents
    l111C = l111C + 1
    l1114$ = String(255, 0)
    l1116% = extfn1718(l111A, l111C, l1114$, Len(p1110) + 1, &H400)
    l1118 = fnD40(l1114$)
    If  l1118 = p1110 Then Exit Do
Loop
DoEvents
l111E% = extfn18A0(l111A, l111C)
l1120 = extfn1980(l1124, &H111, l111E%, 0&)
L6438:
Resume Next
End Sub

Function fnB10 (p112A As Variant) As String
Dim l112E As String
Dim l1130 As Variant
Dim l1134 As String
Dim l1136 As Variant
Dim l113A As Variant
Dim l113E As String
Dim l1140 As String
Dim l1142 As String
On Error GoTo L65B4
l112E$ = p112A
Do: DoEvents
    l1130 = InStr(l112E$, " ")
    l1134$ = l112E$
    If  l1130 <> 0 Then l1134$ = Left$(l112E$, l1130 - 1)
    If  l1130 <> 0 Then l112E$ = Mid$(l112E$, l1130 + 1, Len(l112E$))
    For l1136 = 1 To Len(l1134$)
        Randomize
        l113A = Int(Len(l1134$) * Rnd + 1)
        l113E$ = l113E$ + Mid$(l1134$, l113A, 1)
        l1140$ = Left$(l1134$, l113A - 1)
        l1142$ = Mid$(l1134$, l113A + 1, Len(l1134$))
        l1134$ = l1140$ + l1142$
    Next l1136
    l113E$ = l113E$ & " "
Loop Until l1130 = 0
L65B4:
fnB10 = Trim$(l113E$)
Exit Function
End Function

Sub subB48 (p1144 As String, p1146 As String)
Dim l114A As Integer
Dim l114C As Variant
Dim l1150 As Variant
Dim l115C As Variant
Dim l1160 As Variant
Dim l1164 As Variant
Dim l116C As Integer
On Error Resume Next
l115C = extfn1A60("AOL Frame25", 0&)
Call sub790("Send an Instant Message")
l114A% = extfn1A60("AOL FRAME25", 0&)
L6686:
Call sub790("Send an Instant Message")
Do
DoEvents
l116C = extfn1A98(l114A%, "Send Instant Message")
Loop Until l116C <> 0
l114C = extfn1AD0(l116C, "_AOL_EDIT")
If  l114C = 0 Then GoTo L6686
If  fn2C0() = 30 Then l1150 = extfn1AD0(l116C, "RICHCNTL")
If  fn2C0() = 25 Then l1150 = extfn1B40(l114C, gc0990)
l1164 = extfn1A98(l116C, "Send")
l1160 = extfn19B8(l114C, gc073A, 0, p1144)
l1160 = extfn19B8(l1150, gc073A, 0, p1146)
l115C = extfn1A28(l1164, gc078C, gc063E, 0)
l115C = extfn1A28(l1164, gc078E, gc063E, 0)
End Sub

Sub subB80 (p116E As Variant, p1172 As Variant)
Dim l1176 As Integer
Dim l1178 As Integer
Dim l117A As Integer
Dim l117C As Integer
Dim l117E As Integer
Dim l1180 As Integer
Dim l1182 As Variant
Dim l1186 As Integer
Dim l1188 As Integer
Dim l118A As Variant
Dim l118E As Integer
Dim l1190 As Integer
Dim l1192 As Variant
Dim l1196 As Variant
l1176% = extfn1A60("AOL Frame25", 0&)
l1178% = extfn1AD0(l1176%, "MDIClient")
sub790 "Send an Instant Message"
Do
    l117A% = DoEvents()
    l117C% = extfn1A98(l1178%, "Send Instant Message")
    l117E% = extfn1A98(l117C%, "To:")
Loop Until l117C% <> 0 And l117E% <> 0
l1180% = extfn1AD0(l117C%, "_AOL_Edit")
l1182 = extfn19B8(l1180%, gc073A, 0&, p116E)
l1186% = extfn1AD0(l117C%, "RICHCNTL")
If  l1186% <> 0 Then
    l1182 = extfn19B8(l1186%, gc073A, 0&, p1172)
    l1188% = extfn1AD0(l117C%, "_AOL_Icon")
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
    l1188% = extfn1B40(l1188%, gc0990)
End If
If  l1186% = 0 Then
    l1180% = extfn1B40(l1180%, gc0990)
    l1182 = extfn19B8(l1180%, gc073A, 0&, p1172)
    l1188% = extfn1A98(l117C%, "Send")
End If
l118A = extfn1980(l1188%, gc07B8, 0, 0)
l118A = extfn1980(l1188%, gc07BA, 0, 0)
Do
    l117A% = DoEvents()
    l118E% = extfn1A60("#32770", "America Online")
    l117C% = extfn1A98(l1178%, "Send Instant Message")
Loop Until l118E% <> 0 Or l117C% = 0
If  l118E% <> 0 Then
    l1190% = extfn1A98(l118E%, "OK")
    l1192 = extfn1980(l1190%, gc07B8, 0, 0)
    l1192 = extfn1980(l1190%, gc07BA, 0, 0)
    l118A = extfn1980(l117C%, &H10, 0, 0)
    l1196 = 1
End If
End Sub

Sub subBB8 (p119A As Variant, p119E As Variant, p11A2 As Variant)
Dim l11A6 As Variant
Dim l11AA As Integer
Dim l11AC As Integer
Dim l11AE As Integer
Dim l11B0 As Integer
Dim l11B2 As Integer
Dim l11B4 As Integer
Dim l11B6 As Integer
l11A6 = fnC60()
If  l11A6 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l11AA% = extfn1A60("AOL Frame25", 0&)
l11AC% = extfn1AD0(l11AA%, "MDIClient")
sub790 "Compose Mail"
Do
    l11AE% = DoEvents()
    l11B0% = extfn1A98(l11AC%, "Compose Mail")
Loop Until l11B0% <> 0
l11B2% = extfn1AD0(l11B0%, "_AOL_Edit")
l11A6 = extfn19B8(l11B2%, gc073A, 0&, p119A)
l11B2% = extfn1B40(l11B2%, gc0990)
l11B2% = extfn1B40(l11B2%, gc0990)
l11B2% = extfn1B40(l11B2%, gc0990)
l11B2% = extfn1B40(l11B2%, gc0990)
l11A6 = extfn19B8(l11B2%, gc073A, 0&, p119E)
l11B4% = extfn1AD0(l11B0%, "RICHCNTL")
If  l11B4% <> 0 Then
    l11A6 = extfn19B8(l11B4%, gc073A, 0&, p11A2)
End If
If  l11B4% = 0 Then
    l11B2% = extfn1B40(l11B2%, gc0990)
    l11B2% = extfn1B40(l11B2%, gc0990)
    l11B2% = extfn1B40(l11B2%, gc0990)
    l11A6 = extfn19B8(l11B2%, gc073A, 0&, p11A2)
End If
l11B6% = extfn1AD0(l11B0%, "_AOL_Icon")
sub368 (l11B6%)
End Sub

Sub subBF0 (p11B8 As Variant)
Dim l11BC As Integer
Dim l11BE As Integer
Dim l11C0 As Variant
Dim l11C4 As Variant
l11BC% = fn528()
l11BE% = extfn1AD0(l11BC%, "_AOL_Edit")
l11C0 = extfn19B8(l11BE%, gc073A, 0&, p11B8)
l11C4 = extfn1980(l11BE%, gc0790, 13, 0)
End Sub

Sub subC28 (p11C8 As Integer, ByVal p11CA As String)
Dim l11CC As Variant
l11CC = extfn19B8(p11C8%, gc073A, 0, p11CA$)
End Sub

Function fnC60 () As Variant
Dim l11D4 As Integer
Dim l11D6 As Integer
Dim l11D8 As Integer
Dim l11DA As Integer
l11D4% = extfn1A60("AOL Frame25", 0&)
If  l11D4% = 0 Then
    fnC60 = 0
    Exit Function
End If
l11D6% = extfn1AD0(l11D4%, "MDIClient")
l11D8% = extfn1A98(l11D6%, "Welcome, ")
If  l11D8% = 0 Then
    fnC60 = 0
    Exit Function
End If
l11DA% = extfn1AD0(l11D8%, "_AOL_Combobox")
If  l11DA% <> 0 Then
    fnC60 = 0
    Exit Function
End If
fnC60 = 2
End Function

Sub subC98 ()
Dim l11DC As Variant
Dim l11E0 As Integer
Dim l11E2 As Integer
Dim l11E4 As Integer
Dim l11E6 As Integer
Dim l11E8 As Integer
Dim l11EA As Integer
Dim l11EC As Variant
Dim l11F0 As Integer
l11DC = fnC60()
If  l11DC = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l11E0% = extfn1A60("AOL Frame25", 0&)
l11E2% = extfn1AD0(l11E0%, "MDIClient")
Do
l11E4% = DoEvents()
l11E6% = extfn1A98(l11E2%, "Exit?")
l11E8% = extfn1A60("_AOL_Modal", "America Online")
Loop Until l11E6% <> 0 Or l11E8% <> 0
If  l11E6% <> 0 Then
    l11EA% = extfn1AD0(l11E6%, "_AOL_Icon")
    l11EA% = extfn1B40(l11EA%, gc098E)
    l11EA% = extfn1B40(l11EA%, gc0992)
    l11EA% = extfn1B40(l11EA%, gc0992)
    l11EC = extfn1980(l11EA%, gc07B8, 0, 0)
    l11EC = extfn1980(l11EA%, gc07BA, 0, 0)
End If
If  l11E8% <> 0 Then
    l11F0% = extfn1A98(l11E8%, "&Yes")
    l11EC = extfn1980(l11F0%, gc07B8, 0, 0)
    l11EC = extfn1980(l11F0%, gc07BA, 0, 0)
End If
End Sub

Sub subCD0 (p11F4 As Form)
extsub2048 p11F4.hWnd, gc00BE, -1, -1, -1, -1, gc00C2 Or gc00C4
End Sub

Sub subD08 (p11F6 As Variant)
Dim l11FA As Variant
l11FA = Timer
Do While Timer - l11FA < p11F6
DoEvents
Loop
End Sub

Function fnD40 (p1200 As String) As String
Dim l1202 As Variant
Dim l1206 As String
For l1202 = 1 To Len(p1200$)
    If  (Mid$(p1200$, l1202, 1) <> Chr$(0)) Then
        l1206$ = l1206$ + Mid$(p1200$, l1202, 1)
    Else
        GoTo L788E
    End If
Next
L788E:
fnD40 = l1206$
End Function

Function fnD78 (p120A As Variant) As String
Dim l120E As Variant
Dim l1212 As Variant
Dim l1216 As Variant
Dim l121A As Variant
On Error Resume Next
l120E = p120A
L78B6:
    l1212 = InStr(l120E, " ")
    If  l1212 <> 0 Then
        l1216 = Left$(l120E, l1212 - 1)
        l121A = Right$(l120E, Len(l120E) - (l1212))
        l120E = l1216 + l121A
        GoTo L78B6
    End If
L793C:
fnD78 = l120E
End Function

Function fnDB0 () As String
Dim l1220 As Integer
Dim l1222 As Integer
Dim l1224 As Variant
Dim l1228 As String
Dim l122A As Variant
Dim l122E As Variant
Dim l1232 As String
Dim l1234 As String
    On Error GoTo L7AA4
    l1220% = extfn1A60("AOL Frame25", 0&)
    l1222% = extfn1A98(l1220%, "Welcome, ")
    l1224 = extfn17C0(l1222%)
    l1228$ = Space(l1224)
    l122A = extfn1C20(l1222%, l1228$, l1224)
    l122E = InStr(l1228$, ", ")
    If  l122E = False Then Exit Function
    l1232$ = Mid$(l1228$, l122E + 2)
    l1234$ = Mid$(l1232$, 1, Len(l1232$) - 1)
    l1234$ = Trim(l1234$)
    fnDB0 = l1234$
L7AA4:
Resume Next
End Function

Sub subDE8 ()
Dim l1236 As Integer
Dim l1238 As Variant
    Do
        DoEvents
        l1236% = extfn1A60("#32770", 0&)
    Loop Until l1236% <> 0
    l1238 = extfn1980(l1236%, gc0006, 0, 0)
End Sub

Sub subE20 ()
Dim l123C As Variant
Dim l1240 As String
Dim l1242 As Variant
Dim l1246 As Variant
Dim l124A As Variant
Do
DoEvents
l123C = extfn1A60("#32770", "America Online")
If  l1240$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until l123C <> 0
    l1242 = extfn1A98(l123C, "OK")
    l1246 = extfn1980(l1242, gc07B8, 0, 0&)
    l124A = extfn1980(l1242, gc07BA, 0, 0&)
End Sub

Function fnE58 (p1250 As Integer) As String
Dim l1252 As String
Dim l1254 As Variant
Dim l1258 As Variant
    l1254 = extfn1A28(p1250, &HE, 0, 0&)
    l1252$ = String$(l1254, " ")
    l1258 = extfn1A28(p1250, &HD, (l1254 + 1), ByVal l1252$)
    l1252$ = Trim$(l1252$)
    fnE58 = l1252$
End Function
