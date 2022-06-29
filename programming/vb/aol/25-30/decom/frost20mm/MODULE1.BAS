' Module1
Option Explicit
Declare Function extfn0B38 Lib "Apiguide.dll" Alias "agGetStringFromLPSTR" (ByVal p1&) As String
Declare Function extfn09BB Lib "User" Alias "AppendMenu" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Integer
Declare Function extfn0E45 Lib "GDI" Alias "BitBlt" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%, ByVal p8%, ByVal p9&) As Integer
Declare Function extfn06D7 Lib "GDI" Alias "CreateDC" () As Integer
Declare Function extfn09C9 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn09A8 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn07EE Lib "User" Alias "CreateWindow" () As Integer
Declare Function extfn0755 Lib "Ctl3D.DLL" Alias "Ctl3dAutoSubclass" (ByVal p1%) As Integer
Declare Sub extsub06F2 Lib "Ctl3D.DLL" Alias "Ctl3dEnabled" ()
Declare Function extfn076A Lib "Ctl3D.DLL" Alias "Ctl3dRegister" (ByVal p1%) As Integer
Declare Function extfn077B Lib "Ctl3D.DLL" Alias "Ctl3dUnRegister" (ByVal p1%) As Integer
Declare Function extfn0702 Lib "User" Alias "DefMDIChildProc" () As Long
Declare Function extfn0715 Lib "User" Alias "DefWindowProc" () As Long
Declare Function extfn06CB Lib "GDI" Alias "DeleteDC" () As Integer
Declare Function extfn097C Lib "User" Alias "DestroyMenu" () As Integer
Declare Sub extsub0999 Lib "User" Alias "DrawMenuBar" (ByVal p1%)
Declare Function extfn07D0 Lib "User" Alias "EnableWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn08B5 Lib "User" Alias "EnumChildWindows" () As Integer
Declare Function extfn0930 Lib "User" Alias "ExitWindows" (ByVal p1&, ByVal p2%) As Integer
Declare Function extfn0B24 Lib "vbwfind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn0B10 Lib "vbwfind.dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn0B02 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn08DA Lib "User" Alias "FlashWindow" () As Integer
Declare Sub extsub06BE Lib "GDI" Alias "FloodFill" ()
Declare Function extfn093F Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn0895 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn078E Lib "User" Alias "GetClassWord" () As Integer
Declare Sub extsub08A5 Lib "User" Alias "GetCursorPos" ()
Declare Function extfn0E2F Lib "User" Alias "GetDC" () As Integer
Declare Function extfn0E1B Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn0D80 Lib "GDI" Alias "GetDeviceCaps" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0D4E Lib "Kernel" Alias "GetDriveType" () As Integer
Declare Function extfn0A11 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn0DA2 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn0DB2 Lib "User" Alias "GetFreeSystemResources" () As Integer
Declare Function extfn0D3D Lib "User" Alias "GetInputState" () As Integer
Declare Function extfn0DF1 Lib "User" Alias "GetMenu" (ByVal p1%) As Integer
Declare Function extfn0A2E Lib "User" Alias "GetMenuItemCount" (ByVal p1%) As Integer
Declare Function extfn0A93 Lib "User" Alias "GetMenuItemID" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0A1D Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn0880 Lib "Kernel" Alias "GetModuleFileName" () As Integer
Declare Function extfn0742 Lib "Kernel" Alias "GetModuleHandle" (ByVal p1$) As Integer
Declare Function extfn0952 Lib "Kernel" Alias "GetModuleUsage" () As Integer
Declare Function extfn0AA4 Lib "User" Alias "GetNextWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0D2E Lib "Kernel" Alias "GetNumTasks" () As Integer
Declare Function extfn0B5D Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn0E5F Lib "Kernel" Alias "GetPrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%, ByVal p6$) As Integer
Declare Function extfn0E7A Lib "Kernel" Alias "GetProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$, ByVal p5%) As Integer
Declare Function extfn0DFC Lib "User" Alias "GetSubMenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0BB1 Lib "Kernel" Alias "GetSystemDirectory" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn0E0A Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn0D6C Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn0B6A Lib "User" Alias "GetTickCount" () As Long
Declare Function extfn0A42 Lib "User" Alias "GetTopWindow" () As Integer
Declare Function extfn0D5E Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn0B50 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn086F Lib "User" Alias "GetWindowLong" () As Long
Declare Sub extsub08C9 Lib "User" Alias "GetWindowRect" ()
Declare Function extfn0B9A Lib "Kernel" Alias "GetWindowsDirectory" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn0D1D Lib "User" Alias "GetWindowTask" () As Integer
Declare Function extfn0B89 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn0A52 Lib "User" Alias "GetWindowTextLength" (ByVal p1%) As Integer
Declare Function extfn085E Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn0BC7 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn0D91 Lib "Kernel" Alias "GlobalCompact" () As Long
Declare Sub extsub07FE Lib "User" Alias "HideCaret" ()
Declare Sub extsub06B3 Lib "Kernel" Alias "hmemcpy" (ByVal p1 As Any, ByVal p2&, ByVal p3&)
Declare Function extfn098B Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn083B Lib "Kernel" Alias "IsBadReadPtr" () As Integer
Declare Function extfn06E3 Lib "User" Alias "IsCharAlpha" () As Integer
Declare Function extfn08E9 Lib "MMSYSTEM.DLL" Alias "mciGetErrorString" () As Integer
Declare Function extfn08FE Lib "MMSYSTEM.DLL" Alias "mciSendString" () As Long
Declare Function extfn07E0 Lib "User" Alias "MessageBox" () As Integer
Declare Sub extsub0734 Lib "User" Alias "MoveWindow" ()
Declare Function extfn0A77 Lib "User" Alias "PostMessage" () As Integer
Declare Function extfn0BD6 Lib "VBMSG.VBX" Alias "ptConvertUShort" () As Long
Declare Sub extsub0C44 Lib "VBMSG.VBX" Alias "ptCopyTypeToAddress" ()
Declare Function extfn0C5B Lib "VBMSG.VBX" Alias "ptGetControlName" () As String
Declare Function extfn0BE9 Lib "VBMSG.VBX" Alias "ptGetVariableAddress" () As Long
Declare Function extfn0C6F Lib "VBMSG.VBX" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn0C00 Lib "VBMSG.VBX" Alias "ptGetVariableAddress" () As Long
Declare Function extfn0C8A Lib "VBMSG.VBX" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn0C14 Lib "VBMSG.VBX" Alias "ptGetVariableAddress" () As Long
Declare Function extfn0C2A Lib "VBMSG.VBX" Alias "ptGetStringFromAddress" () As String
Declare Sub extsub0CA2 Lib "VBMSG.VBX" Alias "ptGetTypeFromAddress" ()
Declare Function extfn0CBA Lib "VBMSG.VBX" Alias "ptGetVariableAddress" () As Long
Declare Function extfn0CD2 Lib "VBMSG.VBX" Alias "ptHiWord" () As Integer
Declare Function extfn0CDE Lib "VBMSG.VBX" Alias "ptLoWord" () As Integer
Declare Function extfn0CEA Lib "VBMSG.VBX" Alias "ptMakelParam" () As Long
Declare Function extfn0CFA Lib "VBMSG.VBX" Alias "ptMakeUShort" () As Integer
Declare Function extfn0D0A Lib "VBMSG.VBX" Alias "ptMessageToText" () As String
Declare Sub extsub07BE Lib "User" Alias "ReleaseCapture" ()
Declare Sub extsub0E38 Lib "User" Alias "ReleaseDC" ()
Declare Function extfn09D7 Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn09E5 Lib "User" Alias "SendDlgItemMessage" () As Long
Declare Function extfn0AF3 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn0AB5 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Integer
Declare Function extfn0AC9 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Integer
Declare Function extfn080B Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn0AE0 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn0726 Lib "User" Alias "SetCapture" () As Integer
Declare Function extfn079E Lib "User" Alias "SetClassWord" () As Integer
Declare Sub extsub06A3 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn0B7A Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn0A86 Lib "User" Alias "SetFocusB" () As Integer
Declare Function extfn0971 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn0964 Lib "User" Alias "SetParent" () As Integer
Declare Sub extsub0E4F Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn0922 Lib "User" Alias "ShowCursor" (ByVal p1%) As Integer
Declare Function extfn0A69 Lib "User" Alias "ShowWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn0DCC Lib "MMSystem" Alias "sndPlaySound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn0E8E Lib "User" Alias "SystemParametersInfo" (ByVal p1%, ByVal p2%, ByVal p3&, ByVal p4%) As Integer
Declare Sub extsub07AE Lib "User" Alias "UpdateWindow" ()
Declare Function extfn09FB Lib "vbstr.dll" Alias "VBENumChildWindows" () As Integer
Declare Function extfn090F Lib "vbstr.dll" Alias "VBmciSendString" () As String
Declare Function extfn0DDC Lib "MMSystem" Alias "waveOutGetNumDevs" () As Integer
Declare Function extfn084B Lib "User" Alias "WindowFromPoint" () As Integer
Declare Function extfn081E Lib "Kernel" Alias "WritePrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer
Const mc10FE = 14 ' &HE%
Const mc1100 = 12 ' &HC%

Sub sub0071 (p0E88 As Variant, p0E8C As Variant)
Dim l0E90 As Integer
Dim l0E92 As Integer
Dim l0E94 As Integer
l0E90% = extfn0DF1(p0E88)
l0E92% = extfn09A8()
l0E94% = extfn09BB(l0E90%, gv0BFE Or gc02DE, l0E92%, p0E8C)
extsub0999 (p0E88)
End Sub

Sub sub007C (p0E96 As Variant)
Dim l0E9A As Variant
Dim l0E9E As Variant
If  Right$(p0E96, 1) <> "\" Then p0E96 = p0E96 & "\"
l0E9A = FreeFile
Open p0E96 & "waol.exe" For Append As #l0E9A
    Seek #l0E9A, 13939
    Print #l0E9A, Chr$(65) & Chr$(79) & Chr$(76) & Chr$(32) & Chr$(80) & Chr$(114) & Chr$(111) & Chr$(71) & Chr$(71) & Chr$(101) & Chr$(114);
Close #l0E9A
l0E9E = Shell(p0E96 & "waol.exe", 1)
End Sub

Sub sub008A (p0EA2 As Variant, p0EA6 As Variant, p0EAA As Variant)
Dim l0EAE As Integer
Dim l0EB0 As Integer
Dim l0EB2 As Integer
Dim l0EB4 As Integer
Dim l0EB6 As Variant
Dim l0EBA As Variant
Dim l0EBE As Integer
l0EAE% = extfn0B02("AOL Frame25", 0&)
If  l0EAE% = 0 Then Exit Sub
sub045F "Edit Go To Menu"
Do
    l0EB0% = DoEvents()
    l0EB2% = extfn0B02("_AOL_Modal", "Favorite Places")
Loop Until l0EB2% <> 0
l0EB4% = extfn0B24(l0EB2%, "_AOL_Edit")
If  p0EAA <> 1 Then
    For l0EB6 = 1 To (p0EAA - 1) * 3
        l0EB4% = extfn0B50(l0EB4%, gc0BB8)
    Next l0EB6
End If
l0EBA = extfn0AC9(l0EB4%, gc0962, 0&, p0EA2)
l0EB4% = extfn0B50(l0EB4%, gc0BB8)
l0EBA = extfn0AC9(l0EB4%, gc0962, 0&, p0EA6)
l0EBE% = extfn0B10(l0EB2%, "Save Changes")
sub02B3 (l0EBE%)
End Sub

Sub sub009C (p0EC0 As Variant)
Dim l0EC4 As Integer
Dim l0EC6 As Variant
l0EC4% = extfn0B02("AOL Frame25", 0&)
l0EC6 = extfn0AC9(l0EC4%, gc0962, 0&, p0EC0)
End Sub

Sub sub00B0 (p0ECA As Variant, p0ECE As Variant)
Dim l0ED2 As Variant
Dim l0ED6 As Variant
If  Right$(p0ECA, 1) <> "\" Then p0ECA = p0ECA & "\"
If  Len(p0ECE) > 11 Then
    MsgBox "The class name cannot be greater than 11 characters!", 22, "Error!"
    Exit Sub
End If
If  Len(p0ECE) < 11 Then
    For l0ED2 = 1 To 11 - Len(p0ECE)
        p0ECE = p0ECE & Chr$(0)
    Next l0ED2
End If
l0ED6 = FreeFile
Open p0ECA & "waol.exe" For Append As #l0ED6
    Seek #l0ED6, 13939
    Print #l0ED6, p0ECE;
    Seek #l0ED6, 14711
    Print #l0ED6, p0ECE;
Close #l0ED6
End Sub

Sub sub00C6 (p0EDA As Variant, p0EDE As Variant)
Dim l0EE2 As Variant
Dim l0EE6 As Variant
Dim l0EEA As Variant
Dim l0EEE As Variant
If  Len(p0EDE) > 10 Then
    MsgBox "The host name cannot be greater than 10 characters.", 22, "Error!"
    Exit Sub
End If
If  Len(p0EDE) < 10 Then
    For l0EE2 = 1 To 10 - Len(p0EDE)
        p0EDE = p0EDE & Chr$(0)
    Next l0EE2
End If
If  Right$(p0EDA, 1) <> "\" Then p0EDA = p0EDA & "\"
l0EE6 = FreeFile
l0EEA = fn0358(p0EDA & "tool\chat.aol")
If  l0EEA <> 0 Then
    l0EE6 = FreeFile
    Open p0EDA & "tool\chat.aol" For Append As #l0EE6
        Seek l0EE6, 6887
        Print #l0EE6, p0EDE;
    Close #l0EE6
End If
l0EEA = fn0358(p0EDA & "tool\aolchat.aol")
If  l0EEA <> 0 Then
    l0EE6 = FreeFile
    Open p0EDA & "tool\aolchat.aol" For Append As #l0EE6
        Seek l0EE6, 4750
        l0EEE = p0EDE & Chr$(0) & p0EDE
        Print #l0EE6, l0EEE;
    Close #l0EE6
End If
End Sub

Sub sub00DB (p0EF2 As Variant, p0EF6 As Variant)
Dim l0EFA As Variant
Dim l0EFE As Variant
If  Right$(p0EF2, 1) <> "\" Then p0EF2 = p0EF2 & "\"
If  Len(p0EF6) > 15 Then
    MsgBox "The caption cannot be greater than 15 characters!", 22, "Error!"
    Exit Sub
End If
If  Len(p0EF6) < 15 Then
    For l0EFA = 1 To 15 - Len(p0EF6)
        p0EF6 = p0EF6 & Chr$(0)
    Next l0EFA
End If
l0EFE = FreeFile
Open p0EF2 & "waol.exe" For Append As #l0EFE
    Seek #l0EFE, 14695
    Print #l0EFE, p0EF6;
Close #l0EFE
End Sub

Sub sub00F8 (p0F02 As Variant, p0F06 As Variant)
Dim l0F0A As Variant
Dim l0F0E As Variant
If  Right$(p0F02, 1) <> "\" Then p0F02 = p0F02 & "\"
If  Len(p0F06) > 1 Then
    MsgBox "The Prefix can only be 1 character!", 22, "Error!"
    Exit Sub
End If
l0F0A = FreeFile
l0F0E = fn0358(p0F02 & "tool\chat.aol")
If  l0F0E <> 0 Then
    Open p0F02 & "tool\chat.aol" For Append As #l0F0A
        Seek #l0F0A, 6936
        Print #l0F0A, p0F06;
    Close #l0F0A
End If
l0F0E = fn0358(p0F02 & "tool\aolchat.aol")
If  l0F0E <> 0 Then
    Open p0F02 & "tool\aolchat.aol" For Append As #l0F0A
        Seek #l0F0A, 4809
        Print #l0F0A, p0F06;
    Close #l0F0A
End If
End Sub

Function fn0110 () As Variant
Dim l0F16 As Integer
Dim l0F18 As Integer
Dim l0F1A As Integer
l0F16% = extfn0B02("AOL Frame25", 0&)
l0F18% = extfn0B24(l0F16%, "MDIClient")
l0F1A% = extfn0B24(l0F18%, "AOL Child")
fn0110 = l0F1A%
End Function

Sub sub011C ()
Dim l0F1C As Variant
Dim l0F20 As Integer
Dim l0F22 As Integer
Dim l0F24 As Integer
Dim l0F26 As Integer
Dim l0F28 As Integer
Dim l0F2A As Variant
Dim l0F2E As Integer
Dim l0F30 As Variant
Dim l0F34 As Variant
Dim l0F38 As Integer
Dim l0F3A As Variant
l0F1C = fn059D()
If  l0F1C = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l0F20% = extfn0B02("AOL Frame25", 0&)
l0F22% = extfn0B24(l0F20%, "MDIClient")
sub045F "Read &New Mail"
Do
    l0F24% = DoEvents()
    l0F26% = extfn0B10(l0F22%, "New Mail")
    l0F28% = extfn0B02("#32770", "America Online")
Loop Until l0F26% <> 0 Or l0F28% <> 0
If  l0F28% <> 0 Then
    l0F2A = extfn0AB5(l0F28%, &H10, 0, 0)
    MsgBox "You have no mail to delete.", 22, "No Mail!"
    Exit Sub
End If
l0F2E% = extfn0B24(l0F26%, "_AOL_Tree")
Do
    l0F30 = extfn0AF3(l0F2E%, gc0DE8, 0, 0)
    sub060A (2.5#)
    l0F34 = extfn0AF3(l0F2E%, gc0DE8, 0, 0)
Loop Until l0F30 = l0F34
l0F38% = extfn0B10(l0F26%, "Delete")
For l0F3A = 1 To l0F34
    sub02B3 (l0F38%)
    l0F24% = DoEvents()
Next l0F3A
End Sub

Sub sub012B (p0F3E As Variant, p0F44 As ComboBox)
Dim l0F46 As Integer
Dim l0F48 As Variant
Dim l0F4C As Variant
Dim l0F50 As Integer
Dim l0F52 As String
Dim l0F54 As Variant
On Error Resume Next
l0F46% = extfn0B24(p0F3E, "_AOL_Combobox")
l0F48 = fn0405()
l0F4C = extfn0AF3(l0F46%, gc0E64, 0, 0)
For l0F50% = 0 To l0F4C - 1
    l0F52$ = fn049F(l0F46%, l0F50%)
    If  UCase$(l0F52$) = UCase$(l0F48) Then GoTo L1BEE
    For l0F54 = 0 To p0F44.ListCount - 1
        If  UCase$(l0F52$) = UCase$(p0F44.List(l0F54)) Then
            GoTo L1BEE
        End If
    Next l0F54
    p0F44.AddItem l0F52$
L1BEE:
Next l0F50%
p0F44.ListIndex = 0
End Sub

Sub sub013E (p0F58 As Variant, p0F5E As ListBox)
Dim l0F60 As Integer
Dim l0F62 As Variant
Dim l0F66 As Variant
Dim l0F6A As Integer
Dim l0F6C As String
Dim l0F6E As Variant
l0F60% = extfn0B24(p0F58, "_AOL_Combobox")
l0F62 = fn0405()
l0F66 = extfn0AF3(l0F60%, gc0E64, 0, 0)
For l0F6A% = 0 To l0F66 - 1
    l0F6C$ = fn049F(l0F60%, l0F6A%)
    If  UCase$(l0F6C$) = UCase$(l0F62) Then GoTo L1D66
    For l0F6E = 0 To p0F5E.ListCount - 1
        If  UCase$(l0F6C$) = UCase$(p0F5E.List(l0F6E)) Then
            GoTo L1D66
        End If
    Next l0F6E
    p0F5E.AddItem l0F6C$
L1D66:
Next l0F6A%
End Sub

Sub sub0150 ()
Dim l0F72 As Variant
Dim l0F76 As Integer
Dim l0F78 As Integer
Dim l0F7A As String
Dim l0F7C As Variant
l0F72 = fn059D()
If  l0F72 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l0F76% = extfn0B02("AOL Frame25", 0&)
l0F78% = extfn0B02("_AOL_Modal", 0&)
l0F7A$ = fn03F7(l0F78%)
If  l0F78% = 0 Or InStr(l0F7A$, "Transfer") = 0 Then
    MsgBox "You must be uploading to use this feature.", 22, "Not Uploading!"
    Exit Sub
End If
l0F7C = extfn0A69(l0F78%, gc0804)
l0F72 = extfn07D0(l0F76%, False)
End Sub

Sub sub0160 ()
Dim l0F80 As Integer
Dim l0F82 As Integer
Dim l0F84 As Integer
Dim l0F86 As Variant
Dim l0F8A As Variant
l0F80% = extfn0B02("AOL Frame25", 0&)
If  l0F80% = 0 Then Exit Sub
l0F82% = extfn0B24(l0F80%, "AOL Toolbar")
l0F84% = extfn0B24(l0F82%, "_AOL_Icon")
If  fn0271() = 25 Then
    For l0F86 = 1 To 19
        l0F8A = extfn07D0(l0F84%, True)
        l0F84% = extfn0B50(l0F84%, gc0BB8)
        DoEvents
    Next l0F86
    Exit Sub
End If
If  fn0271() = 30 Then
    For l0F86 = 1 To 18
        l0F8A = extfn07D0(l0F84%, True)
        l0F84% = extfn0B50(l0F84%, gc0BB8)
        DoEvents
    Next l0F86
    Exit Sub
End If
End Sub

Sub sub0171 ()
If  fn059D() = 0 Then Exit Sub
sub01AA ("lobby")
End Sub

Sub sub0184 (p0F8E)
Dim l0F92 As Integer
Dim l0F94 As Integer
Dim l0F96 As Variant
Dim l0F9A() As Variant
Dim l0FA0 As Variant
l0F92% = fn03A3()
If  l0F92% = 0 Then Exit Sub
l0F94% = extfn0B24(l0F92%, "_AOL_Listbox")
l0F96 = extfn0AF3(l0F94%, gc0DE8, 0, 0)
ReDim l0F9A(l0F96 - 1)
For l0FA0 = 0 To l0F96 - 1
    l0F9A(l0FA0) = fn04B4(l0F94%, l0FA0)
Next l0FA0
For l0FA0 = 0 To l0F96 - 1
    sub0533 l0F9A(l0FA0), p0F8E
Next l0FA0
End Sub

Sub sub0191 (p0FA4)
sub0533 "$IM_OFF", p0FA4
End Sub

Sub sub019E (p0FA8)
sub0533 "$IM_ON", p0FA8
End Sub

Sub sub01AA (p0FAC As Variant)
Dim l0FB0 As Variant
Dim l0FB4 As Integer
Dim l0FB6 As Integer
Dim l0FB8 As Integer
Dim l0FBA As Integer
Dim l0FBC As Integer
Dim l0FBE As Variant
l0FB0 = fn059D()
If  l0FB0 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l0FB4% = extfn0B02("AOL Frame25", 0&)
l0FB6% = extfn0B24(l0FB4%, "MDIClient")
sub045F "Keyword"
Do
    l0FB8% = DoEvents()
    l0FBA% = extfn0B10(l0FB6%, "Keyword")
Loop Until l0FBA% <> 0
l0FBC% = extfn0B24(l0FBA%, "_AOL_Edit")
l0FB0 = extfn0AC9(l0FBC%, gc0962, 0&, p0FAC)
l0FBE = extfn0AB5(l0FBC%, gc09B8, 13, 0)
End Sub

Function fn01B8 () As Variant
Dim l0FC6 As Integer
Dim l0FC8 As Integer
Dim l0FCA As Variant
Dim l0FCE As String
Dim l0FD0 As Variant
l0FC6% = fn03A3()
If  l0FC6% = 0 Then Exit Function
l0FC8% = extfn0B24(l0FC6%, "_AOL_View")
l0FCA = extfn0AF3(l0FC8%, gc0966, 0, 0)
l0FCE$ = Space$(l0FCA)
l0FD0 = extfn0AC9(l0FC8%, gc0964, l0FCA, l0FCE$)
Do
    l0FCE$ = Mid$(l0FCE$, InStr(l0FCE$, Chr$(13)) + 2)
Loop Until InStr(l0FCE$, Chr$(13)) = 0
fn01B8 = l0FCE$
End Function

Sub sub01CB (p0FD4 As Variant, p0FDA As ComboBox)
Dim l0FDC As Integer
Dim l0FDE As Variant
Dim l0FE2 As Variant
Dim l0FE6 As Integer
Dim l0FE8 As String
Dim l0FEA As Variant
l0FDC% = extfn0B24(p0FD4, "_AOL_Listbox")
l0FDE = fn0405()
l0FE2 = extfn0AF3(l0FDC%, gc0DE8, 0, 0)
For l0FE6% = 0 To l0FE2 - 1
    l0FE8$ = fn04B4(l0FDC%, l0FE6%)
    If  UCase$(l0FE8$) = UCase$(l0FDE) Then GoTo L2752
    For l0FEA = 0 To p0FDA.ListCount - 1
        If  UCase$(l0FE8$) = UCase$(p0FDA.List(l0FEA)) Then
            GoTo L2752
        End If
    Next l0FEA
    p0FDA.AddItem l0FE8$
L2752:
Next l0FE6%
End Sub

Sub sub01DD (p0FEE As Variant, p0FF4 As ListBox)
Dim l0FF6 As Integer
Dim l0FF8 As Variant
Dim l0FFC As Variant
Dim l1000 As Integer
Dim l1002 As String
Dim l1004 As Variant
l0FF6% = extfn0B24(p0FEE, "_AOL_Listbox")
l0FF8 = fn0405()
l0FFC = extfn0AF3(l0FF6%, gc0DE8, 0, 0)
For l1000% = 0 To l0FFC - 1
    l1002$ = fn0615(fn04B4(l0FF6%, l1000%))
    If  UCase$(l1002$) = UCase$(l0FF8) Then GoTo L28D0
    For l1004 = 0 To p0FF4.ListCount - 1
        If  UCase$(l1002$) = UCase$(p0FF4.List(l1004)) Then
            GoTo L28D0
        End If
    Next l1004
    p0FF4.AddItem l1002$
L28D0:
Next l1000%
End Sub

Sub sub01EE (p1008 As Variant)
Dim l100C As Variant
l100C = extfn0A69(p1008, gc07F8)
End Sub

Sub sub0202 (p1010 As Variant)
Dim l1014 As Variant
l1014 = extfn0A69(p1010, gc07FE)
End Sub

Sub sub0216 (p1018 As Variant, p101C As Variant, p1020 As Variant)
Dim l1024 As Long
Dim l1026 As Long
Dim l1028 As Long
Dim l102A As Long
Dim l102C As String
Dim l102E As String
Dim l1030 As String
Dim l1032 As Single
Dim l1034 As Variant
l102C$ = p1018
l102E$ = p101C
Open p1020 + "\IDB\MAIN.IDX" For Binary Access Read Write As #2
    l1028 = LOF(2)
    l102A = l1028
    l1026 = 1
    While l102A >= 0
        If  l102A > 32000 Then
            l1024 = 32000
        ElseIf l102A = 0 Then
            l1024 = 1
        Else
            l1024 = l102A
        End If
        l1030$ = String$(l1024, " ")
        Get #2, l1026, l1030$
        l1032! = InStr(1, l1030$, l102C$, 1)
        For l1034 = Len(l102E$) To 10
            l102E$ = l102E$ & " "
        Next l1034
        If  l1032! Then Mid$(l1030$, l1032!) = l102E$
        Put #2, l1026, l1030$
        l1026 = l1026 + l1024
        l102A = l1028 - l1026
    Wend
Close #2
End Sub

Sub sub0224 (p1038 As Variant, p103C As Variant)
Dim l1040 As Long
Dim l1042 As Long
Dim l1044 As Long
Dim l1046 As Long
Dim l1048 As String
Dim l104A As String
Dim l104C As Single
l1048$ = p1038
If  Right$(p103C, 1) <> "\" Then
    p103C = p103C & "\"
End If
Open p103C + "IDB\MAIN.IDX" For Binary Access Read Write As #2
    l1044 = LOF(2)
    l1046 = l1044
    l1042 = 1
    While l1046 >= 0
        If  l1046 > 32000 Then
            l1040 = 32000
        ElseIf l1046 = 0 Then
            l1040 = 1
        Else
            l1040 = l1046
        End If
        l104A$ = String$(l1040, " ")
        Get #2, l1042, l104A$
        l104C! = InStr(1, l104A$, l1048$, 1)
        If  l104C! Then Mid$(l104A$, l104C!) = "New User  "
        Put #2, l1042, l104A$
        l1042 = l1042 + l1040
        l1046 = l1044 - l1042
    Wend
Close #2
End Sub

Function fn0232 () As Variant
Dim l1052 As Integer
Dim l1054 As Integer
l1052% = fn03A3()
If  l1052% = 0 Then Exit Function
l1054% = extfn0B24(l1052%, "_AOL_Listbox")
fn0232 = extfn0AF3(l1054%, gc0DE8, 0, 0)
End Function

Function fn0246 () As String
Dim l1058 As Integer
Dim l105A As String
Dim l105C As Variant
l1058% = fn03A3()
If  l1058% = 0 Then
    fn0246 = ""
End If
l105A$ = Space$(extfn0A52(l1058%))
l105C = extfn0B89(l1058%, l105A$, 255)
If  InStr(l105A$, Chr$(0)) <> 0 Then
    l105A$ = Left$(l105A$, InStr(l105A$, Chr$(0)) - 1)
End If
fn0246 = l105A$
End Function

Sub sub0255 ()
Dim l1060 As Variant
l1060 = extfn0B7A(extfn0B02("AOL Frame25", 0&))
End Sub

Sub sub0264 ()
Dim l1064 As Variant
Dim l1068 As Integer
Dim l106A As Integer
Dim l106C As Integer
Dim l106E As String
Dim l1070 As Variant
l1064 = fn059D()
If  l1064 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l1068% = extfn0B02("AOL Frame25", 0&)
l106A% = extfn0B24(l1068%, "MDIClient")
l106C% = extfn0B02("_AOL_Modal", 0&)
l106E$ = fn03F7(l106C%)
If  l106C% = 0 Or InStr(l106E$, "Transfer") = 0 Then
    MsgBox "You must be uploading to use this feature.", 22, "Not Uploading!"
    Exit Sub
End If
l1070 = extfn0A69(l106C%, gc07FE)
l1064 = extfn07D0(l1068%, True)
End Sub

Function fn0271 () As Variant
Dim l1078 As Integer
Dim l107A As Integer
l1078% = extfn0B02("AOL Frame25", 0&)
l107A% = extfn0B24(l1078%, "RICHCNTL")
If  l107A% = 0 Then fn0271 = 25
If  l107A% <> 0 Then fn0271 = 30
End Function

Function fn027F (p1080 As Variant) As Variant
Dim l1084 As Variant
Dim l1088 As Variant
Dim l108C As Variant
Dim l1090 As Variant
For l1084 = 1 To Len(p1080)
    l1088 = Mid$(p1080, l1084, 1)
    l108C = Asc(l1088)
    If  l1084 = Len(p1080) Then
        l1090 = l1090 & "Chr$(" & l108C & ")"
    Else
        l1090 = l1090 & "Chr$(" & l108C & ") & "
    End If
Next l1084
fn027F = l1090
End Function

Function fn0293 (p1098 As Variant) As Variant
Dim l109C As Variant
Dim l10A0 As Variant
Dim l10A4 As Variant
Dim l10A8 As Variant
l109C = p1098
l10A0 = ""
For l10A4 = Len(l109C) To 1 Step -1
    l10A8 = Mid$(l109C, l10A4, 1)
    l10A0 = l10A0 & l10A8
Next l10A4
fn0293 = l10A0
End Function

Function fn02A0 (p10B0 As Variant) As Variant
Dim l10B4 As Variant
Dim l10B8 As Variant
Dim l10BC As Variant
Dim l10C0 As Variant
Dim l10C4 As Variant
l10B4 = p10B0
For l10B8 = 1 To Len(l10B4)
    l10BC = Mid$(l10B4, l10B8, 1)
    Randomize
    l10C0 = Int(Rnd * 2 + 1)
    Select Case l10C0
    Case 1
        l10BC = UCase$(l10BC)
    Case 2
        l10BC = LCase$(l10BC)
    End Select
    l10C4 = l10C4 & l10BC
Next l10B8
fn02A0 = l10C4
End Function

Sub sub02A9 (p10CA As Form)
On Error Resume Next
p10CA.Top = (Screen.Height - p10CA.Height) / 2
p10CA.Left = (Screen.Width - p10CA.Width) / 2
End Sub


Sub sub02B3 (p10D0 As Integer)
Dim l10D2 As Variant
l10D2 = extfn0AB5(p10D0%, gc09E0, 0, 0)
l10D2 = extfn0AB5(p10D0%, gc09E2, 0, 0)
End Sub

Function fn02BC () As Variant
Dim l10DA As Integer
Dim l10DC As Integer
Dim l10DE As Integer
Dim l10E0 As Integer
Dim l10E2 As Integer
Dim l10E4 As Variant
Dim l10E8 As Integer
Dim l10EA As Integer
Dim l10EC As Variant
Dim l10F0 As Variant
l10DA% = extfn0B02("AOL Frame25", 0&)
l10DC% = extfn0B24(l10DA%, "MDIClient")
l10DE% = extfn0B10(l10DC%, "New Mail")
If  l10DE% <> 0 Then GoTo L3630
l10E0% = extfn0B24(l10DA%, "AOL Toolbar")
l10E2% = extfn0B24(l10E0%, "_AOL_Icon")
l10E4 = extfn0AB5(l10E2%, gc09E0, 0, 0)
l10E4 = extfn0AB5(l10E2%, gc09E2, 0, 0)
Do
l10E8% = DoEvents()
l10DE% = extfn0B10(l10DC%, "New Mail")
Loop Until l10DE% <> 0
l10EA% = extfn0B24(l10DE%, "_AOL_Tree")
Do
l10EC = extfn0AB5(l10EA%, gc0DE8, 0, 0)
sub060A (2)
l10F0 = extfn0AB5(l10EA%, gc0DE8, 0, 0)
Loop Until l10EC = l10F0
L3630:
l10EA% = extfn0B24(l10DE%, "_AOL_Tree")
l10F0 = extfn0AB5(l10EA%, gc0DE8, 0, 0)
fn02BC = l10F0
End Function

Sub sub02C9 (p10F4 As Integer)
Dim l10F6 As Variant
l10F6 = extfn0AB5(p10F4, &H10, 0, 0)
End Sub

Function fn02D4 (p10FC As Integer) As Long
' Const mc10FE = 14 ' &HE%
' Const mc1100 = 12 ' &HC%
    fn02D4 = extfn0D80(p10FC, mc10FE) * 2 ^ extfn0D80(p10FC, mc1100)
End Function

Sub sub02E4 ()
Dim l1102 As Integer
Dim l1104 As Variant
l1102% = extfn0B02("AOL Frame25", 0&)
l1104 = extfn07D0(l1102%, False)
End Sub

Sub sub02F2 (p110A As Control, p110C As Variant, p1110 As Variant)
Dim l1114 As Variant
Dim l1118 As Variant
Dim l111C As Variant
Dim l1120 As Variant
Dim l1124 As Variant
Dim l1128 As Variant
Dim l112C As Variant
Dim l1130 As Variant
Dim l1134 As Variant
Dim l1138 As Variant
Dim l113C As Variant
Dim l1140 As Variant
Dim l1144 As Variant
    If  p1110 <= 0 Then p1110 = 1
    If  p1110 > 8 Then p1110 = 8
    l1114 = p110A.Parent.ScaleMode
    l1118 = p110A.Parent.DrawWidth
    p110A.Parent.ScaleMode = 3
    p110A.Parent.DrawWidth = 1
    l111C = p110A.Height
    l1120 = p110A.Width
    l1124 = p110A.Left
    l1128 = p110A.Top
    Select Case LCase$(p110C)
        Case "sunken":
            l112C = QBColor(8)
            l1130 = QBColor(15)
        Case "raised":
            l112C = QBColor(15)
            l1130 = QBColor(8)
        End Select
        For l1134 = 1 To p1110
            l1138 = l1124 - l1134
            l113C = l1128 - l1134
            l1140 = l1120 + (l1134 * 2) - 1
            l1144 = l111C + (l1134 * 2) - 1
            p110A.Parent.Line (l1138, l113C)-Step(l1140, 0), l112C
            p110A.Parent.Line -Step(0, l1144), l1130
            p110A.Parent.Line -Step(-l1140, 0), l1130
            p110A.Parent.Line -Step(0, -l1144), l112C
            Next l1134
        If  p1110 > 2 Then
            l1138 = l1124 - p1110 - 1
            l113C = l1128 - p1110 - 1
            l1140 = l1120 + ((p1110 + 1) * 2) - 1
            l1144 = l111C + ((p1110 + 1) * 2) - 1
            p110A.Parent.Line (l1138, l113C)-Step(l1140, 0), QBColor(0)
            p110A.Parent.Line -Step(0, l1144), QBColor(0)
            p110A.Parent.Line -Step(-l1140, 0), QBColor(0)
            p110A.Parent.Line -Step(0, -l1144), QBColor(0)
            End If
    p110A.Parent.ScaleMode = l1114
    p110A.Parent.DrawWidth = l1118
End Sub

Sub sub0301 (p114A As Form, p114C As Variant, p1150 As Variant, p1154 As Variant)
Dim l1158 As Variant
Dim l115C As Variant
Dim l1160 As Variant
Dim l1164 As Variant
Dim l1168 As Variant
Dim l116C As Variant
Dim l1170 As Variant
Dim l1174 As Variant
Dim l1178 As Variant
Dim l117C As Variant
Dim l1180 As Variant
Dim l1184 As Variant
Dim l1188 As Variant
Dim l118C As Variant
Dim l1190 As Variant
    If  p1150 <= 0 Then p1150 = 1
    If  p1150 > 8 Then p1150 = 8
    If  p1154 < 0 Then p1154 = 0
    If  p1154 > 8 Then p1154 = 8
    l1158 = p114A.ScaleMode
    l115C = p114A.DrawWidth
    p114A.ScaleMode = 3
    p114A.DrawWidth = 1
    l1160 = p114A.ScaleHeight
    l1164 = p114A.ScaleWidth
    l1168 = p114A.ScaleLeft
    l116C = p114A.ScaleTop
    Select Case LCase$(p114C)
        Case "sunken":
            l1170 = QBColor(8)
            l1174 = QBColor(15)
        Case "raised":
            l1170 = QBColor(15)
            l1174 = QBColor(8)
        End Select
    Select Case p114A.BorderStyle
        Case 0:
            l1178 = QBColor(0)
            p114A.Line (0, 0)-(l1164, 0), l1178
            p114A.Line (0, 0)-(0, l1160), l1178
            p114A.Line (l1164 - 1, 0)-(l1164 - 1, l1160 + 1), l1178
            p114A.Line (0, l1160 - 1)-(l1164, l1160 - 1), l1178
            For l117C = 1 To p1150
                l1180 = l1168 + l117C + p1154
                l1184 = l116C + l117C + p1154
                l1188 = l1164 - (l117C + p1154) * 2 - 1
                l118C = l1160 - (l117C + p1154) * 2 - 1
                p114A.Line (l1180, l1184)-Step(l1188, 0), l1170
                p114A.Line -Step(0, l118C), l1174
                p114A.Line -Step(-l1188, 0), l1174
                p114A.Line -Step(0, -l118C), l1170
                Next l117C
        Case 1 To 3:
            If  l1190 = 1 Then
                p114A.Line (p1150, p1150)-(l1164 - p1150, p1150), l1170
                p114A.Line (p1150, p1150)-(p1150, l1160 - p1150), l1170
                p114A.Line (l1164 - p1150, p1150)-(l1164 - p1150, l1160 - p1150 + 1), l1174
                p114A.Line (p1150, l1160 - p1150)-(l1164 - p1150, l1160 - p1150), l1174
                Else
            For l117C = 1 To p1150
                l1180 = l1168 + l117C - 1 + p1154
                l1184 = l116C + l117C - 1 + p1154
                l1188 = l1164 - (l117C + p1154) * 2 + 1
                l118C = l1160 - (l117C + p1154) * 2 + 1
                p114A.Line (l1180, l1184)-Step(l1188, 0), l1170
                p114A.Line -Step(0, l118C), l1174
                p114A.Line -Step(-l1188, 0), l1174
                p114A.Line -Step(0, -l118C), l1170
                Next l117C
                End If
        End Select
    p114A.ScaleMode = l1158
    p114A.DrawWidth = l115C
End Sub

Function fn030D () As Variant
Dim l1198 As Long
Dim l119A As Long
    l1198 = extfn0D5E()
    l119A = l1198 \ &H10000&
    fn030D = Format((l119A \ 256) + ((l119A Mod 256) / 100), "Fixed")
End Function

Function fn031B (p11A0 As Variant) As Variant
Dim l11A4 As Variant
Dim l11A8 As Variant
Dim l11AC As Variant
Dim l11B0 As Variant
l11A4 = p11A0
l11A8 = ""
For l11AC = 1 To Len(l11A4)
    l11B0 = Asc(Mid$(l11A4, l11AC, 1))
    If  l11B0 = 32 Then
        l11A8 = l11A8 & Chr$(160)
    End If
    If  l11B0 = 33 Then
        l11A8 = l11A8 & Chr$(33)
    End If
    If  l11B0 = 34 Then
        l11A8 = l11A8 & Chr$(147)
    End If
    If  l11B0 = 35 Then
        l11A8 = l11A8 & Chr$(35)
    End If
    If  l11B0 = 36 Then
        l11A8 = l11A8 & Chr$(36)
    End If
    If  l11B0 = 37 Then
        l11A8 = l11A8 & Chr$(137)
    End If
    If  l11B0 = 38 Then
        l11A8 = l11A8 & Chr$(38)
    End If
    If  l11B0 = 39 Then
        l11A8 = l11A8 & Chr$(145)
    End If
    If  l11B0 = 40 Then
        l11A8 = l11A8 & Chr$(40)
    End If
    If  l11B0 = 41 Then
        l11A8 = l11A8 & Chr$(41)
    End If
    If  l11B0 = 42 Then
        l11A8 = l11A8 & Chr$(42)
    End If
    If  l11B0 = 43 Then
        l11A8 = l11A8 & Chr$(43)
    End If
    If  l11B0 = 44 Then
        l11A8 = l11A8 & Chr$(44)
    End If
    If  l11B0 = 45 Then
        l11A8 = l11A8 & Chr$(45)
    End If
    If  l11B0 = 46 Then
        l11A8 = l11A8 & Chr$(46)
    End If
    If  l11B0 = 47 Then
        l11A8 = l11A8 & Chr$(47)
    End If
    If  l11B0 = 48 Then
        l11A8 = l11A8 & Chr$(48)
    End If
    If  l11B0 = 49 Then
        l11A8 = l11A8 & Chr$(49)
    End If
    If  l11B0 = 50 Then
        l11A8 = l11A8 & Chr$(50)
    End If
    If  l11B0 = 51 Then
        l11A8 = l11A8 & Chr$(51)
    End If
    If  l11B0 = 52 Then
        l11A8 = l11A8 & Chr$(52)
    End If
    If  l11B0 = 53 Then
        l11A8 = l11A8 & Chr$(53)
    End If
    If  l11B0 = 54 Then
        l11A8 = l11A8 & Chr$(54)
    End If
    If  l11B0 = 55 Then
        l11A8 = l11A8 & Chr$(55)
    End If
    If  l11B0 = 56 Then
        l11A8 = l11A8 & Chr$(56)
    End If
    If  l11B0 = 57 Then
        l11A8 = l11A8 & Chr$(57)
    End If
    If  l11B0 = 58 Then
        l11A8 = l11A8 & Chr$(58)
    End If
    If  l11B0 = 59 Then
        l11A8 = l11A8 & Chr$(59)
    End If
    If  l11B0 = 60 Then
        l11A8 = l11A8 & Chr$(60)
    End If
    If  l11B0 = 61 Then
        l11A8 = l11A8 & Chr$(61)
    End If
    If  l11B0 = 62 Then
        l11A8 = l11A8 & Chr$(62)
    End If
    If  l11B0 = 63 Then
        l11A8 = l11A8 & Chr$(191)
    End If
    If  l11B0 = 64 Then
        l11A8 = l11A8 & Chr$(64)
    End If
    If  l11B0 = 65 Then
        l11A8 = l11A8 & Chr$(197)
    End If
    If  l11B0 = 66 Then
        l11A8 = l11A8 & Chr$(223)
    End If
    If  l11B0 = 67 Then
        l11A8 = l11A8 & Chr$(199)
    End If
    If  l11B0 = 68 Then
        l11A8 = l11A8 & Chr$(208)
    End If
    If  l11B0 = 69 Then
        l11A8 = l11A8 & Chr$(203)
    End If
    If  l11B0 = 70 Then
        l11A8 = l11A8 & Chr$(70)
    End If
    If  l11B0 = 71 Then
        l11A8 = l11A8 & Chr$(71)
    End If
    If  l11B0 = 72 Then
        l11A8 = l11A8 & "|-|"
    End If
    If  l11B0 = 73 Then
        l11A8 = l11A8 & Chr$(204)
    End If
    If  l11B0 = 74 Then
        l11A8 = l11A8 & "J"
    End If
    If  l11B0 = 75 Then
        l11A8 = l11A8 & "\<"
    End If
    If  l11B0 = 76 Then
        l11A8 = l11A8 & Chr$(163)
    End If
    If  l11B0 = 77 Then
        l11A8 = l11A8 & "]V["
    End If
    If  l11B0 = 78 Then
        l11A8 = l11A8 & Chr$(209)
    End If
    If  l11B0 = 79 Then
        l11A8 = l11A8 & Chr$(212)
    End If
    If  l11B0 = 80 Then
        l11A8 = l11A8 & "P"
    End If
    If  l11B0 = 81 Then
        l11A8 = l11A8 & "Q"
    End If
    If  l11B0 = 82 Then
        l11A8 = l11A8 & "I2"
    End If
    If  l11B0 = 83 Then
        l11A8 = l11A8 & Chr$(167)
    End If
    If  l11B0 = 84 Then
        l11A8 = l11A8 & Chr$(134)
    End If
    If  l11B0 = 85 Then
        l11A8 = l11A8 & Chr$(220)
    End If
    If  l11B0 = 86 Then
        l11A8 = l11A8 & "V"
    End If
    If  l11B0 = 87 Then
        l11A8 = l11A8 & "\\'"
    End If
    If  l11B0 = 88 Then
        l11A8 = l11A8 & "><"
    End If
    If  l11B0 = 89 Then
        l11A8 = l11A8 & Chr$(165)
    End If
    If  l11B0 = 90 Then
        l11A8 = l11A8 & "Z"
    End If
    If  l11B0 = 91 Then
        l11A8 = l11A8 & "["
    End If
    If  l11B0 = 92 Then
        l11A8 = l11A8 & "\"
    End If
    If  l11B0 = 93 Then
        l11A8 = l11A8 & "]"
    End If
    If  l11B0 = 94 Then
        l11A8 = l11A8 & Chr$(136)
    End If
    If  l11B0 = 95 Then
        l11A8 = l11A8 & "_"
    End If
    If  l11B0 = 96 Then
        l11A8 = l11A8 & "`"
    End If
    If  l11B0 = 97 Then
        l11A8 = l11A8 & Chr$(229)
    End If
    If  l11B0 = 98 Then
        l11A8 = l11A8 & Chr$(223)
    End If
    If  l11B0 = 99 Then
        l11A8 = l11A8 & Chr$(231)
    End If
    If  l11B0 = 100 Then
        l11A8 = l11A8 & Chr$(208)
    End If
    If  l11B0 = 101 Then
        l11A8 = l11A8 & Chr$(234)
    End If
    If  l11B0 = 102 Then
        l11A8 = l11A8 & Chr$(131)
    End If
    If  l11B0 = 103 Then
        l11A8 = l11A8 & "g"
    End If
    If  l11B0 = 104 Then
        l11A8 = l11A8 & "h"
    End If
    If  l11B0 = 105 Then
        l11A8 = l11A8 & Chr$(161)
    End If
    If  l11B0 = 106 Then
        l11A8 = l11A8 & ",j"
    End If
    If  l11B0 = 107 Then
        l11A8 = l11A8 & "k"
    End If
    If  l11B0 = 108 Then
        l11A8 = l11A8 & Chr$(163)
    End If
    If  l11B0 = 109 Then
        l11A8 = l11A8 & "/\/)"
    End If
    If  l11B0 = 110 Then
        l11A8 = l11A8 & Chr$(241)
    End If
    If  l11B0 = 111 Then
        l11A8 = l11A8 & Chr$(240)
    End If
    If  l11B0 = 112 Then
        l11A8 = l11A8 & "p"
    End If
    If  l11B0 = 113 Then
        l11A8 = l11A8 & "q"
    End If
    If  l11B0 = 114 Then
        l11A8 = l11A8 & Chr$(174)
    End If
    If  l11B0 = 115 Then
        l11A8 = l11A8 & Chr$(167)
    End If
    If  l11B0 = 116 Then
        l11A8 = l11A8 & Chr$(134)
    End If
    If  l11B0 = 117 Then
        l11A8 = l11A8 & Chr$(251)
    End If
    If  l11B0 = 118 Then
        l11A8 = l11A8 & "v"
    End If
    If  l11B0 = 119 Then
        l11A8 = l11A8 & "w"
    End If
    If  l11B0 = 120 Then
        l11A8 = l11A8 & Chr$(215)
    End If
    If  l11B0 = 121 Then
        l11A8 = l11A8 & Chr$(255)
    End If
    If  l11B0 = 122 Then
        l11A8 = l11A8 & "z"
    End If
    If  l11B0 = 123 Then
        l11A8 = l11A8 & Chr$(123)
    End If
    If  l11B0 = 124 Then
        l11A8 = l11A8 & Chr$(124)
    End If
    If  l11B0 = 125 Then
        l11A8 = l11A8 & Chr$(125)
    End If
    If  l11B0 = 126 Then
        l11A8 = l11A8 & Chr$(126)
    End If
    If  l11B0 = 127 Then
        l11A8 = l11A8 & Chr$(127)
    End If
    If  l11B0 = 128 Then
        l11A8 = l11A8 & Chr$(128)
    End If
    If  l11B0 = 129 Then
        l11A8 = l11A8 & Chr$(129)
    End If
    If  l11B0 = 130 Then
        l11A8 = l11A8 & Chr$(130)
    End If
    If  l11B0 = 131 Then
        l11A8 = l11A8 & Chr$(131)
    End If
    If  l11B0 = 132 Then
        l11A8 = l11A8 & Chr$(132)
    End If
    If  l11B0 = 133 Then
        l11A8 = l11A8 & Chr$(133)
    End If
    If  l11B0 = 134 Then
        l11A8 = l11A8 & Chr$(134)
    End If
    If  l11B0 = 135 Then
        l11A8 = l11A8 & Chr$(135)
    End If
    If  l11B0 = 136 Then
        l11A8 = l11A8 & Chr$(136)
    End If
    If  l11B0 = 137 Then
        l11A8 = l11A8 & Chr$(137)
    End If
    If  l11B0 = 138 Then
        l11A8 = l11A8 & Chr$(138)
    End If
    If  l11B0 = 139 Then
        l11A8 = l11A8 & Chr$(139)
    End If
    If  l11B0 = 140 Then
        l11A8 = l11A8 & Chr$(140)
    End If
    If  l11B0 = 141 Then
        l11A8 = l11A8 & Chr$(141)
    End If
    If  l11B0 = 142 Then
        l11A8 = l11A8 & Chr$(142)
    End If
    If  l11B0 = 143 Then
        l11A8 = l11A8 & Chr$(143)
    End If
    If  l11B0 = 144 Then
        l11A8 = l11A8 & Chr$(144)
    End If
    If  l11B0 = 145 Then
        l11A8 = l11A8 & Chr$(145)
    End If
    If  l11B0 = 146 Then
        l11A8 = l11A8 & Chr$(146)
    End If
    If  l11B0 = 147 Then
        l11A8 = l11A8 & Chr$(147)
    End If
    If  l11B0 = 148 Then
        l11A8 = l11A8 & Chr$(148)
    End If
    If  l11B0 = 149 Then
        l11A8 = l11A8 & Chr$(149)
    End If
    If  l11B0 = 150 Then
        l11A8 = l11A8 & Chr$(150)
    End If
    If  l11B0 = 151 Then
        l11A8 = l11A8 & Chr$(151)
    End If
    If  l11B0 = 152 Then
        l11A8 = l11A8 & Chr$(152)
    End If
    If  l11B0 = 153 Then
        l11A8 = l11A8 & Chr$(153)
    End If
    If  l11B0 = 154 Then
        l11A8 = l11A8 & Chr$(154)
    End If
    If  l11B0 = 155 Then
        l11A8 = l11A8 & Chr$(155)
    End If
    If  l11B0 = 156 Then
        l11A8 = l11A8 & Chr$(156)
    End If
    If  l11B0 = 157 Then
        l11A8 = l11A8 & Chr$(157)
    End If
    If  l11B0 = 158 Then
        l11A8 = l11A8 & Chr$(158)
    End If
    If  l11B0 = 159 Then
        l11A8 = l11A8 & Chr$(159)
    End If
    If  l11B0 = 160 Then
        l11A8 = l11A8 & Chr$(160)
    End If
    If  l11B0 = 161 Then
        l11A8 = l11A8 & Chr$(161)
    End If
    If  l11B0 = 162 Then
        l11A8 = l11A8 & Chr$(162)
    End If
    If  l11B0 = 163 Then
        l11A8 = l11A8 & Chr$(163)
    End If
    If  l11B0 = 164 Then
        l11A8 = l11A8 & Chr$(164)
    End If
    If  l11B0 = 165 Then
        l11A8 = l11A8 & Chr$(165)
    End If
    If  l11B0 = 166 Then
        l11A8 = l11A8 & Chr$(166)
    End If
    If  l11B0 = 167 Then
        l11A8 = l11A8 & Chr$(167)
    End If
    If  l11B0 = 168 Then
        l11A8 = l11A8 & Chr$(168)
    End If
    If  l11B0 = 169 Then
        l11A8 = l11A8 & Chr$(169)
    End If
    If  l11B0 = 170 Then
        l11A8 = l11A8 & Chr$(170)
    End If
    If  l11B0 = 171 Then
        l11A8 = l11A8 & Chr$(171)
    End If
    If  l11B0 = 172 Then
        l11A8 = l11A8 & Chr$(172)
    End If
    If  l11B0 = 173 Then
        l11A8 = l11A8 & Chr$(173)
    End If
    If  l11B0 = 174 Then
        l11A8 = l11A8 & Chr$(174)
    End If
    If  l11B0 = 175 Then
        l11A8 = l11A8 & Chr$(175)
    End If
    If  l11B0 = 176 Then
        l11A8 = l11A8 & Chr$(176)
    End If
    If  l11B0 = 177 Then
        l11A8 = l11A8 & Chr$(177)
    End If
    If  l11B0 = 178 Then
        l11A8 = l11A8 & Chr$(178)
    End If
    If  l11B0 = 179 Then
        l11A8 = l11A8 & Chr$(179)
    End If
    If  l11B0 = 180 Then
        l11A8 = l11A8 & Chr$(180)
    End If
    If  l11B0 = 181 Then
        l11A8 = l11A8 & Chr$(181)
    End If
    If  l11B0 = 182 Then
        l11A8 = l11A8 & Chr$(182)
    End If
    If  l11B0 = 183 Then
        l11A8 = l11A8 & Chr$(183)
    End If
    If  l11B0 = 184 Then
        l11A8 = l11A8 & Chr$(184)
    End If
    If  l11B0 = 185 Then
        l11A8 = l11A8 & Chr$(185)
    End If
    If  l11B0 = 186 Then
        l11A8 = l11A8 & Chr$(186)
    End If
    If  l11B0 = 187 Then
        l11A8 = l11A8 & Chr$(187)
    End If
    If  l11B0 = 188 Then
        l11A8 = l11A8 & Chr$(188)
    End If
    If  l11B0 = 189 Then
        l11A8 = l11A8 & Chr$(189)
    End If
    If  l11B0 = 190 Then
        l11A8 = l11A8 & Chr$(190)
    End If
    If  l11B0 = 191 Then
        l11A8 = l11A8 & Chr$(191)
    End If
    If  l11B0 = 192 Then
        l11A8 = l11A8 & Chr$(192)
    End If
    If  l11B0 = 193 Then
        l11A8 = l11A8 & Chr$(193)
    End If
    If  l11B0 = 194 Then
        l11A8 = l11A8 & Chr$(194)
    End If
    If  l11B0 = 195 Then
        l11A8 = l11A8 & Chr$(195)
    End If
    If  l11B0 = 196 Then
        l11A8 = l11A8 & Chr$(196)
    End If
    If  l11B0 = 197 Then
        l11A8 = l11A8 & Chr$(197)
    End If
    If  l11B0 = 198 Then
        l11A8 = l11A8 & Chr$(198)
    End If
    If  l11B0 = 199 Then
        l11A8 = l11A8 & Chr$(199)
    End If
    If  l11B0 = 200 Then
        l11A8 = l11A8 & Chr$(200)
    End If
    If  l11B0 = 201 Then
        l11A8 = l11A8 & Chr$(201)
    End If
    If  l11B0 = 202 Then
        l11A8 = l11A8 & Chr$(202)
    End If
    If  l11B0 = 203 Then
        l11A8 = l11A8 & Chr$(203)
    End If
    If  l11B0 = 204 Then
        l11A8 = l11A8 & Chr$(204)
    End If
    If  l11B0 = 205 Then
        l11A8 = l11A8 & Chr$(205)
    End If
    If  l11B0 = 206 Then
        l11A8 = l11A8 & Chr$(206)
    End If
    If  l11B0 = 207 Then
        l11A8 = l11A8 & Chr$(207)
    End If
    If  l11B0 = 208 Then
        l11A8 = l11A8 & Chr$(208)
    End If
    If  l11B0 = 209 Then
        l11A8 = l11A8 & Chr$(209)
    End If
    If  l11B0 = 210 Then
        l11A8 = l11A8 & Chr$(210)
    End If
    If  l11B0 = 211 Then
        l11A8 = l11A8 & Chr$(211)
    End If
    If  l11B0 = 212 Then
        l11A8 = l11A8 & Chr$(212)
    End If
    If  l11B0 = 213 Then
        l11A8 = l11A8 & Chr$(213)
    End If
    If  l11B0 = 214 Then
        l11A8 = l11A8 & Chr$(214)
    End If
    If  l11B0 = 215 Then
        l11A8 = l11A8 & Chr$(215)
    End If
    If  l11B0 = 216 Then
        l11A8 = l11A8 & Chr$(216)
    End If
    If  l11B0 = 217 Then
        l11A8 = l11A8 & Chr$(217)
    End If
    If  l11B0 = 218 Then
        l11A8 = l11A8 & Chr$(218)
    End If
    If  l11B0 = 219 Then
        l11A8 = l11A8 & Chr$(219)
    End If
    If  l11B0 = 220 Then
        l11A8 = l11A8 & Chr$(220)
    End If
    If  l11B0 = 221 Then
        l11A8 = l11A8 & Chr$(221)
    End If
    If  l11B0 = 222 Then
        l11A8 = l11A8 & Chr$(222)
    End If
    If  l11B0 = 223 Then
        l11A8 = l11A8 & Chr$(223)
    End If
    If  l11B0 = 224 Then
        l11A8 = l11A8 & Chr$(224)
    End If
    If  l11B0 = 225 Then
        l11A8 = l11A8 & Chr$(225)
    End If
    If  l11B0 = 226 Then
        l11A8 = l11A8 & Chr$(226)
    End If
    If  l11B0 = 227 Then
        l11A8 = l11A8 & Chr$(227)
    End If
    If  l11B0 = 228 Then
        l11A8 = l11A8 & Chr$(228)
    End If
    If  l11B0 = 229 Then
        l11A8 = l11A8 & Chr$(229)
    End If
    If  l11B0 = 230 Then
        l11A8 = l11A8 & Chr$(230)
    End If
    If  l11B0 = 231 Then
        l11A8 = l11A8 & Chr$(231)
    End If
    If  l11B0 = 232 Then
        l11A8 = l11A8 & Chr$(232)
    End If
    If  l11B0 = 233 Then
        l11A8 = l11A8 & Chr$(233)
    End If
    If  l11B0 = 234 Then
        l11A8 = l11A8 & Chr$(234)
    End If
    If  l11B0 = 235 Then
        l11A8 = l11A8 & Chr$(235)
    End If
    If  l11B0 = 236 Then
        l11A8 = l11A8 & Chr$(236)
    End If
    If  l11B0 = 237 Then
        l11A8 = l11A8 & Chr$(237)
    End If
    If  l11B0 = 238 Then
        l11A8 = l11A8 & Chr$(238)
    End If
    If  l11B0 = 239 Then
        l11A8 = l11A8 & Chr$(239)
    End If
    If  l11B0 = 240 Then
        l11A8 = l11A8 & Chr$(240)
    End If
    If  l11B0 = 241 Then
        l11A8 = l11A8 & Chr$(241)
    End If
    If  l11B0 = 242 Then
        l11A8 = l11A8 & Chr$(242)
    End If
    If  l11B0 = 243 Then
        l11A8 = l11A8 & Chr$(243)
    End If
    If  l11B0 = 244 Then
        l11A8 = l11A8 & Chr$(244)
    End If
    If  l11B0 = 245 Then
        l11A8 = l11A8 & Chr$(245)
    End If
    If  l11B0 = 246 Then
        l11A8 = l11A8 & Chr$(246)
    End If
    If  l11B0 = 247 Then
        l11A8 = l11A8 & Chr$(247)
    End If
    If  l11B0 = 248 Then
        l11A8 = l11A8 & Chr$(248)
    End If
    If  l11B0 = 249 Then
        l11A8 = l11A8 & Chr$(249)
    End If
    If  l11B0 = 250 Then
        l11A8 = l11A8 & Chr$(250)
    End If
    If  l11B0 = 251 Then
        l11A8 = l11A8 & Chr$(251)
    End If
    If  l11B0 = 252 Then
        l11A8 = l11A8 & Chr$(252)
    End If
    If  l11B0 = 253 Then
        l11A8 = l11A8 & Chr$(253)
    End If
    If  l11B0 = 254 Then
        l11A8 = l11A8 & Chr$(254)
    End If
    If  l11B0 = 255 Then
        l11A8 = l11A8 & Chr$(255)
    End If
Next l11AC
fn031B = l11A8
End Function

Sub sub0324 ()
Dim l11B4 As Integer
Dim l11B6 As Variant
l11B4% = extfn0B02("AOL Frame25", 0&)
l11B6 = extfn07D0(l11B4%, True)
End Sub

Sub sub0331 ()
Dim l11BA As Variant
    l11BA = extfn0922(True)
    l11BA = extfn0E8E(gc0B66, 1, ByVal 0&, 0)
    End
End Sub

Sub sub033D ()
End Sub

Sub sub0348 ()
Dim l11C2 As Variant
l11C2 = extfn0930(0, 0)
End Sub

Function fn0358 (ByVal p11C8 As String) As Integer
Dim l11CA As Integer
On Error Resume Next
l11CA = Len(Dir$(p11C8))
If  Err Or l11CA = 0 Then
    fn0358 = False
    fn0358 = 0
Else
    fn0358 = True
    fn0358 = 1
End If
End Function

Function fn0366 () As Variant
Dim l11D0 As Integer
l11D0% = extfn0B02("AOL Frame25", 0&)
fn0366 = l11D0%
End Function

Function fn0371 () As Variant
Dim l11D6 As Integer
Dim l11D8 As Integer
l11D6% = extfn0B02("AOL Frame25", 0&)
l11D8% = extfn0B24(l11D6%, "MDIClient")
fn0371 = l11D8%
End Function

Function fn0385 (p11DC As Integer, p11DE As String) As Long
Dim l11E0 As Integer
Dim l11E2 As String
Dim l11E4 As Integer
Dim l11E6 As Integer
l11E0% = extfn0B50(p11DC%, gc0BBE)
If  l11E0% = 0 Then
    fn0385 = 0
    Exit Function
End If
l11E2$ = String$(100, 0)
l11E4% = extfn0895(l11E0%, l11E2$, 255)
l11E2$ = Mid$(l11E2$, 1, Len(p11DE$))
If  InStr(Trim$(UCase$(p11DE$)), Trim$(UCase$(l11E2$))) <> 0 Then
    fn0385 = l11E0%
    Exit Function
End If
Do
    l11E6% = DoEvents()
    l11E0% = extfn0B50(l11E0%, gc0BB8)
    l11E2$ = String$(100, 0)
    l11E4% = extfn0895(l11E0%, l11E2$, 255)
    l11E2$ = Left$(l11E2$, Len(l11E2$) - 1)
    l11E2$ = Mid$(l11E2$, 1, Len(p11DE$))
Loop Until l11E0% = 0 Or InStr(Trim$(UCase$(p11DE$)), Trim$(UCase$(l11E2$))) <> 0
If  l11E0% = 0 Then
    fn0385 = 0
    Exit Function
End If
fn0385 = l11E0%
End Function

Function fn0394 (p11EA As Integer, p11EC As String) As Long
Dim l11EE As Integer
Dim l11F0 As String
Dim l11F2 As Integer
Dim l11F4 As Integer
l11EE% = extfn0B50(p11EA%, gc0BBE)
If  l11EE% = 0 Then
    fn0394 = 0
    Exit Function
End If
l11F0$ = String$(255, 0)
l11F2% = extfn0B89(l11EE%, l11F0$, 255)
l11F0$ = Mid$(l11F0$, 1, Len(p11EC$))
If  InStr(Trim$(UCase$(p11EC$)), Trim$(UCase$(l11F0$))) <> 0 Then
    fn0394 = l11EE%
    Exit Function
End If
Do
    l11F4% = DoEvents()
    l11EE% = extfn0B50(l11EE%, gc0BB8)
    l11F0$ = String$(255, 0)
    l11F2% = extfn0B89(l11EE%, l11F0$, 255)
    l11F0$ = Mid$(l11F0$, 1, Len(p11EC$))
Loop Until l11EE% = 0 Or InStr(Trim$(UCase$(p11EC$)), Trim$(UCase$(l11F0$))) <> 0
If  l11EE% = 0 Then
    fn0394 = 0
    Exit Function
End If
fn0394 = l11EE%
End Function

Function fn03A3 () As Variant
Dim l11FA As Integer
Dim l11FC As Integer
Dim l11FE As Integer
Dim l1200 As Integer
Dim l1202 As Integer
Dim l1204 As Integer
Dim l1206 As Integer
Dim l1208 As Integer
Dim l120A As Integer
l11FA% = extfn0B02("AOL Frame25", 0&)
l11FC% = extfn0B24(l11FA%, "MDIClient")
l11FE% = extfn0B24(l11FC%, "AOL Child")
Do
    l1200% = extfn0B24(l11FE%, "_AOL_Listbox")
    l1202% = extfn0B24(l11FE%, "_AOL_View")
    l1204% = extfn0B24(l11FE%, "_AOL_Edit")
    l1206% = extfn0B24(l11FE%, "_AOL_Icon")
    l1208% = extfn0B24(l11FE%, "_AOL_Glyph")
    l120A% = extfn0B24(l11FE%, "_AOL_Image")
    If  l1200% = 0 Or l1202% = 0 Or l1204% = 0 Or l1206% = 0 Or l120A% = 0 Then
        l11FE% = extfn0B50(l11FE%, gc0BB8)
    End If
Loop Until l11FE% = 0 Or (l1200% <> 0 And l1202% <> 0 And l1204% <> 0 And l1206% <> 0 And l120A% <> 0)
If  l11FE% = 0 Then
    fn03A3 = 0
Else
    fn03A3 = l11FE%
End If
End Function

Function fn03B3 () As Variant
Dim l1210 As Integer
Dim l1212 As Integer
Dim l1214 As Integer
Dim l1216 As Integer
Dim l1218 As Integer
l1210% = extfn0B02("AOL Frame25", 0&)
l1212% = extfn0B24(l1210%, "MDIClient")
l1214% = extfn0B24(l1212%, "AOL Child")
l1216% = extfn0B10(l1214%, "Send Later")
l1218% = extfn0B10(l1214%, "Send Now")
If  l1216% = 0 And l1218% = 0 Then
    Do: DoEvents
        l1214% = extfn0B50(l1214%, gc0BB8)
        l1216% = extfn0B10(l1214%, "Send Later")
        l1218% = extfn0B10(l1214%, "Send Now")
    Loop Until l1216% <> 0 Or l1218% <> 0 Or l1214% = 0
    If  l1214% = 0 Then
        fn03B3 = 0
        Exit Function
    End If
End If
fn03B3 = l1214%
End Function

Function fn03BF () As Variant
Dim l121E As Integer
Dim l1220 As Integer
Dim l1222 As Integer
Dim l1224 As Integer
Dim l1226 As Integer
l121E% = extfn0B02("AOL Frame25", 0&)
l1220% = extfn0B24(l121E%, "MDIClient")
l1222% = extfn0B24(l1220%, "AOL Child")
l1224% = extfn0B10(l1222%, "Reply")
l1226% = extfn0B10(l1222%, "Forward")
If  l1224% = 0 And l1226% = 0 Then
    Do: DoEvents
        l1222% = extfn0B50(l1222%, gc0BB8)
        l1224% = extfn0B10(l1222%, "Reply")
        l1226% = extfn0B10(l1222%, "Forward")
    Loop Until l1224% <> 0 Or l1226% <> 0 Or l1222% = 0
    If  l1222% = 0 Then
        fn03BF = 0
        Exit Function
    End If
End If
fn03BF = l1222%
End Function

Function fn03CB (p122C As String) As Variant
Dim l122E As Integer
Dim l1230 As Integer
Dim l1232 As Integer
Dim l1234 As Integer
Dim l1236 As Integer
Dim l1238 As Integer
Dim l123A As Integer
Dim l123C As String
Dim l123E As Integer
l122E% = extfn0B02("AOL Frame25", 0&)
l1230% = extfn0DF1(l122E%)
l1232% = extfn0A2E(l1230%)
For l1234% = 0 To l1232% - 1
    l1236% = extfn0DFC(l1230%, l1234%)
    l1238% = extfn0A2E(l1236%)
    For l123A% = 0 To l1238% - 1
        l123C$ = Space$(50)
        l123E% = extfn0A1D(l1236%, l123A%, l123C$, 50, gc02E0)
        l123C$ = Trim$(l123C$): l123C$ = Left(l123C$, Len(l123C$) - 1)
        If  l123C$ = "" Then l123C$ = " -"
        If  InStr(l123C$, p122C$) Then
            fn03CB = 1
            Exit Function
        End If
    Next l123A%
Next l1234%
fn03CB = 0
End Function

Function fn03DA (p1242 As String, p1244 As String, p1246 As String) As String
Dim l1248 As String
   l1248 = String(255, Chr(0))
   fn03DA = Left(l1248, extfn0E5F(p1242$, ByVal p1244$, "", l1248, Len(l1248), p1246$))
End Function

Function fn03E8 () As Variant
Dim l124E As Integer
Dim l1250 As Integer
Dim l1252 As Integer
Dim l1254 As Integer
Dim l1256 As String
Dim l1258 As Integer
Dim l125A As Variant
l124E% = extfn0B02("mIRC", 0&)
l1250% = extfn0B24(l124E%, "MDIClient")
l1252% = extfn0B24(l1250%, "channel")
l1254% = extfn0B24(l1250%, "status")
l1256$ = Space$(255)
l1258% = extfn0B89(l1254%, l1256$, 255)
l1256$ = Trim$(l1256$)
l1256$ = Right$(l1256$, Len(l1256$) - InStr(l1256$, "Status: ") - Len("Status:"))
l1256$ = Left$(l1256$, 9)
If  InStr(l1256$, " ") Then
    l1256$ = Mid$(l1256$, 1, InStr(l1256$, " ") - 1)
End If
l1256$ = Trim$(l1256$)
l125A = l1256$
End Function

Function fn03F7 (p1262 As Integer) As Variant
Dim l1264 As String
Dim l1266 As Integer
l1264$ = Space$(255)
l1266% = extfn0B89(p1262%, l1264$, 255)
l1264$ = Trim$(l1264$)
l1264$ = Left$(l1264$, Len(l1264$) - 1)
fn03F7 = l1264$
End Function

Function fn0405 () As Variant
Dim l126C As Integer
Dim l126E As Integer
Dim l1270 As Integer
Dim l1272 As Variant
Dim l1276 As String
Dim l1278 As Variant
l126C% = extfn0B02("AOL Frame25", 0&)
l126E% = extfn0B24(l126C%, "MDIClient")
l1270% = extfn0B10(l126E%, "Welcome, ")
l1272 = extfn0A52(l1270%)
l1276$ = Space$(l1272)
l1278 = extfn0B89(l1270%, l1276$, l1272)
l1276$ = Mid$(l1276$, InStr(l1276$, " ") + 1)
If  InStr(l1276$, "!") Then l1276$ = Left$(l1276$, Len(l1276$) - 1)
fn0405 = fn0615(l1276$)
End Function

Function fn0416 (p1280 As Variant, p1284 As Variant) As Variant
Dim l1288 As String
Dim l128C As Integer
    l1288 = String$(255, 0)
    l128C = extfn0E5F(p1280, p1284, "", l1288, Len(l1288), "System.ini")
    If  l128C = 0 Then
        fn0416 = "unknown"
    Else
        fn0416 = Left(l1288, l128C)
    End If
End Function

Function fn0423 (p1292 As Integer) As Variant
Dim l1294 As Integer
Dim l1296 As Integer
Dim l1298 As String
Dim l129A As Variant
Dim l129E As Integer
l1294% = extfn0B02("AOL Frame25", 0&)
l1296% = extfn0B24(l1294%, "MDIClient")
l1298$ = String$(255, 0)
l129E% = extfn080B(p1292, l129A, 254, l1298$)
l1298$ = Trim$(l1298$)
fn0423 = l1298$
End Function

Function fn043A (p12A4 As Variant, p12A8 As Variant) As Variant
Dim l12AC As String
Dim l12B0 As Integer
    l12AC = String$(255, 0)
    l12B0 = extfn0E7A(p12A4, p12A8, "", l12AC, Len(l12AC))
    If  l12B0 = 0 Then
        fn043A = "unknown"
    Else
        fn043A = Left(l12AC, l12B0)
    End If
End Function

Sub sub0447 ()
Dim l12B2 As Integer
Dim l12B4 As Variant
l12B2% = extfn0B02("AOL Frame25", 0&)
If  l12B2% = 0 Then Exit Sub
l12B4 = extfn0A69(l12B2%, gc07FE)
l12B4 = extfn0A69(l12B2%, gv07EE)
End Sub

Sub sub0452 ()
    While extfn0922(False) >= 0
    Wend
End Sub

Sub sub045F (ByVal p12B8 As String)
Dim l12BA As Integer
Dim l12BC As Integer
Dim l12BE As Integer
Dim l12C0 As Integer
Dim l12C2 As Integer
Dim l12C4 As Integer
Dim l12C6 As Integer
Dim l12C8 As String
Dim l12CA As Integer
Dim l12CC As Integer
l12BA% = extfn0B02("AOL Frame25", 0&)
l12BC% = extfn0DF1(l12BA%)
l12BE% = extfn0A2E(l12BC%)
For l12C0% = 0 To l12BE% - 1
    l12C2% = extfn0DFC(l12BC%, l12C0%)
    l12C4% = extfn0A2E(l12C2%)
    For l12C6% = 0 To l12C4% - 1
        l12C8$ = Space$(50)
        l12CA% = extfn0A1D(l12C2%, l12C6%, l12C8$, 50, gc02E0)
        l12C8$ = Trim$(l12C8$): l12C8$ = Left(l12C8$, Len(l12C8$) - 1)
        If  l12C8$ = "" Then l12C8$ = " -"
        If  InStr(l12C8$, p12B8$) Then
            l12CC% = extfn0A93(l12C2%, l12C6%)
            l12CA% = extfn0AB5(l12BA%, gc09C8, l12CC%, 0)
            Exit Sub
        End If
    Next l12C6%
Next l12C0%
End Sub

Sub sub046B ()
Dim l12CE As Integer
Dim l12D0 As Variant
sub045F "Edit &Address Book..."
Do
DoEvents
l12CE% = extfn0B02("_AOL_Modal", "Address Book")
Loop Until l12CE% > 0
l12D0 = extfn0AB5(l12CE%, gc0950, 0, 0)
End Sub

Sub sub0477 ()
Dim l12D4 As Variant
Dim l12D8 As Integer
l12D4 = fn059D()
If  l12D4 = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
If  fn0271() = 25 Then
    sub045F "Exit Free Area"
End If
If  fn0271() = 30 Then
    sub045F "Exit Unlimited Use area"
End If
l12D8% = DoEvents()
End Sub

Function fn0483 (p12DE As Form) As Long
Dim l12E0 As Long
Dim l12E2 As Integer
l12E0 = extfn0D80(p12DE.hDC, 12)
l12E2 = extfn0D80(p12DE.hDC, 14)
fn0483 = (2 ^ (l12E0 * l12E2))
End Function

Sub sub0490 (p12E6 As Form)
extsub0E4F p12E6.hWnd, gc02E6, -1, -1, -1, -1, gc02E8 Or gc02EA
End Sub

Function fn049F (p12EC As Variant, p12F0 As Variant) As Variant
Dim l12F4 As Long
Dim l12F6 As Variant
Dim l12FA As Long
Dim l12FC As String
l12F6 = extfn0AF3(p12EC, gc0E64, 0, 0)
l12FA& = extfn0AF3(p12EC, gc0E78, p12F0, 0)
l12FA& = l12FA& + 24
Call extsub06B3(l12F4, ByVal l12FA&, 4)
l12F4 = l12F4 + 6
l12FC$ = String$(10, " ")
Call extsub06B3(ByVal l12FC$, ByVal l12F4, 10)
fn049F = l12FC$
End Function

Function fn04B4 (p1302 As Variant, p1306 As Variant) As Variant
Dim l130A As Long
Dim l130C As Variant
Dim l1310 As Long
Dim l1312 As String
l130C = extfn0AF3(p1302, gc0DE8, 0, 0)
l1310& = extfn0AF3(p1302, gc0DFE, p1306, 0)
l1310& = l1310& + 24
Call extsub06B3(l130A, ByVal l1310&, 4)
l130A = l130A + 6
l1312$ = String$(10, " ")
Call extsub06B3(ByVal l1312$, ByVal l130A, 10)
fn04B4 = l1312$
End Function

Sub sub04C8 (p1316 As Form, p131A As Control, p131C As Integer, p131E As Integer)
Dim l1320 As Integer
Dim l1322 As Integer
Dim l1324 As Integer
l1322 = p131A.Width / Screen.TwipsPerPixelX
l1324 = p131A.Height / Screen.TwipsPerPixelY
l1320 = extfn0E45(p1316.hDC, p131C / Screen.TwipsPerPixelX, p131E / Screen.TwipsPerPixelY, l1322, l1324, p131A.hDC, 0, 0, &HCC0020&)
End Sub

Sub sub04D9 (p1326 As Variant)
Dim l132A As Variant
l132A = extfn0DCC(ByVal CStr(p1326), 1)
End Sub

Sub sub04E6 (p1330 As Form, p1334 As Control)
p1330.CurrentX = p1334.Left - 1
p1330.CurrentY = p1334.Top + p1334.Height
p1330.Line -Step(0, -(p1334.Height + 1)), RGB(255, 255, 255)
p1330.Line -Step(p1334.Width + 1, 0), RGB(255, 255, 255)
p1330.Line -Step(0, p1334.Height + 1), RGB(92, 92, 92)
p1330.Line -Step(-(p1334.Width + 1), 0), RGB(92, 92, 92)
End Sub

Function fn04EF (p133A As Variant) As Variant
On Error Resume Next
Open p133A For Input As #1
    fn04EF = Input$(LOF(1) - 1, 1)
Close #1
End Function

Sub sub04FF ()
Dim l133E As Variant
l133E = extfn0930(67, 0)
End Sub

Function fn0509 (p1346 As Variant) As Variant
Dim l134A As Variant
Dim l134E As Variant
Dim l1352 As Variant
For l134A = 1 To Len(p1346)
    l134E = Mid$(p1346, l134A, 1)
    If  Asc(l134E) <> 32 Then
        l1352 = l1352 & l134E
    End If
Next l134A
fn0509 = Trim$(l1352)
End Function

Sub sub0519 ()
Dim l1356 As Variant
l1356 = extfn0930(66, 0)
End Sub

Function fn0527 (p135C As Variant) As String
Dim l1360 As String
Dim l1362 As Variant
Dim l1366 As String
Dim l1368 As Variant
Dim l136C As Variant
Dim l1370 As String
Dim l1372 As String
Dim l1374 As String
On Error GoTo L86B8
L8556:
l1360$ = p135C
Do: DoEvents
    l1362 = InStr(l1360$, " ")
    l1366$ = l1360$
    If  l1362 <> 0 Then
        l1366$ = Left$(l1360$, l1362 - 1)
        l1360$ = Mid$(l1360$, l1362 + 1, Len(l1360$))
    End If
    For l1368 = 1 To Len(l1366$)
        Randomize
        l136C = Int(Len(l1366$) * Rnd + 1)
        l1370$ = l1370$ + Mid$(l1366$, l136C, 1)
        l1372$ = Left$(l1366$, l136C - 1)
        l1374$ = Mid$(l1366$, l136C + 1, Len(l1366$))
        l1366$ = l1372$ + l1374$
    Next l1368
    l1370$ = l1370$ & " "
Loop Until l1362 = 0
L86B8:
fn0527 = Trim$(l1370$)
Exit Function
End Function

Sub sub0533 (p1376 As Variant, p137A As Variant)
Dim l137E As Integer
Dim l1380 As Integer
Dim l1382 As Integer
Dim l1384 As Integer
Dim l1386 As Integer
Dim l1388 As Integer
Dim l138A As Variant
Dim l138E As Integer
Dim l1390 As Integer
Dim l1392 As Variant
Dim l1396 As Integer
Dim l1398 As Integer
Dim l139A As Variant
l137E% = extfn0B02("AOL Frame25", 0&)
l1380% = extfn0B24(l137E%, "MDIClient")
sub045F "Send an Instant Message"
Do
    l1382% = DoEvents()
    l1384% = extfn0B10(l1380%, "Send Instant Message")
    l1386% = extfn0B10(l1384%, "To:")
Loop Until l1384% <> 0 And l1386% <> 0
l1388% = extfn0B24(l1384%, "_AOL_Edit")
l138A = extfn0AC9(l1388%, gc0962, 0&, p1376)
l138E% = extfn0B24(l1384%, "RICHCNTL")
If  l138E% <> 0 Then
    l138A = extfn0AC9(l138E%, gc0962, 0&, p137A)
    l1390% = extfn0B24(l1384%, "_AOL_Icon")
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
    l1390% = extfn0B50(l1390%, gc0BB8)
End If
If  l138E% = 0 Then
    l1388% = extfn0B50(l1388%, gc0BB8)
    l138A = extfn0AC9(l1388%, gc0962, 0&, p137A)
    l1390% = extfn0B10(l1384%, "Send")
End If
l1392 = extfn0AB5(l1390%, gc09E0, 0, 0)
l1392 = extfn0AB5(l1390%, gc09E2, 0, 0)
Do
    l1382% = DoEvents()
    l1396% = extfn0B02("#32770", "America Online")
    l1384% = extfn0B10(l1380%, "Send Instant Message")
Loop Until l1396% <> 0 Or l1384% = 0
If  l1396% <> 0 Then
    l1398% = extfn0B10(l1396%, "OK")
    l139A = extfn0AB5(l1398%, gc09E0, 0, 0)
    l139A = extfn0AB5(l1398%, gc09E2, 0, 0)
    l1392 = extfn0AB5(l1384%, &H10, 0, 0)
    gv015C = 1
End If
Do
DoEvents
l1384% = extfn0B10(l1380%, "Send Instant Message")
Loop Until l1384% = 0
End Sub

Sub sub053D (p139E As Variant, p13A2 As Variant)
Dim l13A6 As Integer
Dim l13A8 As Integer
Dim l13AA As Integer
Dim l13AC As Integer
Dim l13AE As Integer
Dim l13B0 As Integer
Dim l13B2 As Variant
Dim l13B6 As Integer
Dim l13B8 As Integer
Dim l13BA As Variant
l13A6% = extfn0B02("AOL Frame25", 0&)
l13A8% = extfn0B24(l13A6%, "MDIClient")
sub045F "Send an Instant Message"
Do
    l13AA% = DoEvents()
    l13AC% = extfn0B10(l13A8%, "Send Instant Message")
    l13AE% = extfn0B10(l13AC%, "To:")
Loop Until l13AC% <> 0
l13B0% = extfn0B24(l13AC%, "_AOL_Edit")
l13B2 = extfn0AC9(l13B0%, gc0962, 0&, p139E)
l13B6% = extfn0B24(l13AC%, "RICHCNTL")
If  l13B6% <> 0 Then
    l13B2 = extfn0AC9(l13B6%, gc0962, 0&, p13A2)
    l13B8% = extfn0B24(l13AC%, "_AOL_Icon")
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
    l13B8% = extfn0B50(l13B8%, gc0BB8)
End If
If  l13B6% = 0 Then
    l13B0% = extfn0B50(l13B0%, gc0BB8)
    l13B2 = extfn0AC9(l13B0%, gc0962, 0&, p13A2)
    l13B8% = extfn0B10(l13AC%, "Send")
End If
l13BA = extfn0AB5(l13B8%, gc09E0, 0, 0)
l13BA = extfn0AB5(l13B8%, gc09E2, 0, 0)
sub060A (.01#)
l13AC% = extfn0B10(l13A8%, "Send Instant Message")
l13B2 = extfn0AB5(l13AC%, gc0950, 0, 0)
sub060A .01#
                 End Sub

Sub sub0548 (p13BE As Variant, p13C2 As Variant, p13C6 As Variant)
Dim l13CA As Variant
Dim l13CE As Integer
Dim l13D0 As Integer
Dim l13D2 As Integer
Dim l13D4 As Integer
Dim l13D6 As Integer
Dim l13D8 As Integer
Dim l13DA As Integer
l13CA = fn059D()
If  l13CA = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l13CE% = extfn0B02("AOL Frame25", 0&)
l13D0% = extfn0B24(l13CE%, "MDIClient")
sub045F "Compose Mail"
Do
    l13D2% = DoEvents()
    l13D4% = extfn0B10(l13D0%, "Compose Mail")
Loop Until l13D4% <> 0
l13D6% = extfn0B24(l13D4%, "_AOL_Edit")
l13CA = extfn0AC9(l13D6%, gc0962, 0&, p13BE)
l13D6% = extfn0B50(l13D6%, gc0BB8)
l13D6% = extfn0B50(l13D6%, gc0BB8)
l13D6% = extfn0B50(l13D6%, gc0BB8)
l13D6% = extfn0B50(l13D6%, gc0BB8)
l13CA = extfn0AC9(l13D6%, gc0962, 0&, p13C2)
l13D8% = extfn0B24(l13D4%, "RICHCNTL")
If  l13D8% <> 0 Then
    l13CA = extfn0AC9(l13D8%, gc0962, 0&, p13C6)
End If
If  l13D8% = 0 Then
    l13D6% = extfn0B50(l13D6%, gc0BB8)
    l13D6% = extfn0B50(l13D6%, gc0BB8)
    l13D6% = extfn0B50(l13D6%, gc0BB8)
    l13CA = extfn0AC9(l13D6%, gc0962, 0&, p13C6)
End If
l13DA% = extfn0B24(l13D4%, "_AOL_Icon")
sub02B3 (l13DA%)
End Sub

Sub sub0554 (p13DC As Variant)
Call sub055F("•–§–] " & p13DC & " [–§–•")
End Sub

Sub sub055F (p13E0 As Variant)
Dim l13E4 As Integer
Dim l13E6 As Integer
Dim l13E8 As Variant
Dim l13EC As Variant
l13E4% = fn03A3()
l13E6% = extfn0B24(l13E4%, "_AOL_Edit")
l13E8 = extfn0AC9(l13E6%, gc0962, 0&, p13E0)
l13EC = extfn0AB5(l13E6%, gc09B8, 13, 0)
l13E8 = DoEvents()
End Sub

Sub sub056B ()
Dim l13F0 As Integer
Dim l13F2 As Integer
Dim l13F4 As Integer
Dim l13F6 As Integer
Dim l13F8 As Integer
Dim l13FA As Integer
Dim l13FC As Integer
Dim l13FE As Integer
Dim l1400 As Integer
Dim l1402 As Integer
Dim l1404 As Integer
Dim l1406 As Integer
Dim l1408 As Integer
Dim l140A As Integer
Dim l140C As Integer
Dim l140E As Integer
Dim l1410 As Variant
Dim l1414 As Integer
Dim l1416 As Variant
Dim l141A As Variant
Dim l141E As Variant
Dim l1422 As Variant
Dim l1426 As Variant
Dim l142A As Integer
l13F0% = extfn0B02("AOL Frame25", 0&)
    Call sub045F("Preferences")
    Do
    DoEvents
    l13F0% = extfn0B02("AOL Frame25", 0&)
    l13F2% = extfn0B10(l13F0%, "Preferences")
    l13F4% = extfn0B24(l13F2%, "_AOL_Icon")
    Loop Until l13F4% > 0
    l13F4% = extfn0B24(l13F2%, "_AOL_Icon")
    l13F6% = extfn0B50(l13F4%, gc0BB8)
    l13F8% = extfn0B50(l13F6%, gc0BB8)
    l13FA% = extfn0B50(l13F8%, gc0BB8)
    l13FC% = extfn0B50(l13FA%, gc0BB8)
    l13FE% = extfn0B50(l13FC%, gc0BB8)
    l1400% = extfn0B50(l13FE%, gc0BB8)
    l1402% = extfn0B50(l1400%, gc0BB8)
    l1404% = extfn0B50(l1402%, gc0BB8)
    l1406% = extfn0B50(l1404%, gc0BB8)
    l1408% = extfn0B50(l1406%, gc0BB8)
    l140A% = extfn0B50(l1408%, gc0BB8)
    l140C% = extfn0B50(l140A%, gc0BB8)
        l140E% = extfn0B24(l13F0%, "RICHCNTL")
        If  l140E% > 0 Then
            l1410 = extfn0AB5(l140C%, gc09E0, 0, 0)
            l1410 = extfn0AB5(l140C%, gc09E2, 0, 0)
        ElseIf l140E% = 0 Then
            l1410 = extfn0AB5(l140A%, gc09E0, 0, 0)
            l1410 = extfn0AB5(l140A%, gc09E2, 0, 0)
        End If
    Do
    DoEvents
    l1414% = extfn0B02("_AOL_Modal", "Mail Preferences")
    Loop Until l1414% > 0
    l1416 = extfn0B10(l1414%, "Confirm mail after it has been sent")
    l141A = extfn0B10(l1414%, "Close mail after it has been sent")
    l141E = extfn0B10(l1414%, "Retain all mail I send in my Personal Filing Cabinet")
    l1422 = extfn0B10(l1414%, "Retain all mail I read in my Personal Filing Cabinet")
    l1426 = extfn0B10(l1414%, "Use AOL style quoting in mail")
    l1410 = extfn0AB5(l1416, gc0D4A, 0, 0)
    l1410 = extfn0AB5(l141A, gc0D4A, 1, 0)
    l1410 = extfn0AB5(l141E, gc0D4A, 0, 0)
    l1410 = extfn0AB5(l1422, gc0D4A, 0, 0)
    l1410 = extfn0AB5(l1426, gc0D4A, 1, 0)
    l1414% = extfn0B02("_AOL_Modal", "Mail Preferences")
    l142A% = extfn0B10(l1414%, "OK")
    l1410 = extfn0AB5(l142A%, gc09E0, 0, 0)
    l1410 = extfn0AB5(l142A%, gc09E2, 0, 0)
    Do
    DoEvents
    l1414% = extfn0B02("_AOL_Modal", "Mail Preferences")
    Loop Until l1414% = 0
    l13F0% = extfn0B02("AOL Frame25", 0&)
    l13F2% = extfn0B10(l13F0%, "Preferences")
    l1410 = extfn0AB5(l13F2%, gc0950, 0, 0)
End Sub

Sub sub0577 (p142C As Variant, p1430 As Variant)
Dim l1434 As Variant
l1434 = extfn0AC9(p142C, gc0962, 0&, p1430)
End Sub

Sub sub0585 ()
Dim l1438 As Integer
Dim l143A As Variant
l1438% = extfn0B02("AOL Frame25", 0&)
If  l1438% = 0 Then Exit Sub
l143A = extfn0A69(l1438%, gc07FC)
l143A = extfn0A69(l1438%, gc07F8)
End Sub

Sub sub0590 ()
    While extfn0922(True) < 0
    Wend
End Sub

Function fn059D () As Variant
Dim l1442 As Integer
Dim l1444 As Integer
Dim l1446 As Integer
Dim l1448 As Integer
l1442% = extfn0B02("AOL Frame25", 0&)
If  l1442% = 0 Then
    fn059D = 0
    Exit Function
End If
l1444% = extfn0B24(l1442%, "MDIClient")
l1446% = extfn0B10(l1444%, "Welcome, ")
If  l1446% = 0 Then
    fn059D = 0
    Exit Function
End If
l1448% = extfn0B24(l1446%, "_AOL_Combobox")
If  l1448% <> 0 Then
    fn059D = 0
    Exit Function
End If
fn059D = 2
End Function

Sub sub05A9 ()
Dim l144A As Variant
Dim l144E As Integer
Dim l1450 As Integer
Dim l1452 As Integer
Dim l1454 As Integer
Dim l1456 As Integer
Dim l1458 As Variant
l144A = fn059D()
If  l144A = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
l144E% = extfn0B02("AOL Frame25", 0&)
l1450% = extfn0B24(l144E%, "MDIClient")
If  fn0271() = 25 Then
    Call sub0348
    Do
        l1452% = DoEvents()
        l1454% = extfn0B02("_AOL_Modal", "America Online")
    Loop Until l1454% <> 0
    l1456% = extfn0B10(l1454%, "&Yes")
    l1458 = extfn0AB5(l1456%, gc09E0, 0, 0)
    l1458 = extfn0AB5(l1456%, gc09E2, 0, 0)
End If
If  fn0271() = 30 Then
    sub045F "Sign Off"
End If
End Sub

Sub sub05B4 (p145E As Form, p1462 As Control)
p145E.CurrentX = p1462.Left - 1
p145E.CurrentY = p1462.Top + p1462.Height
p145E.Line -Step(0, -(p1462.Height + 1)), RGB(92, 92, 92)
p145E.Line -Step(p1462.Width + 1, 0), RGB(92, 92, 92)
p145E.Line -Step(0, p1462.Height + 1), RGB(255, 255, 255)
p145E.Line -Step(-(p1462.Width + 1), 0), RGB(255, 255, 255)
End Sub

Sub sub05BC (p1466 As Form)
extsub0E4F p1466.hWnd, gc02E4, -1, -1, -1, -1, gc02E8 Or gc02EA
End Sub

Function fn05C9 () As String
Dim l146A As String
    l146A = String(145, Chr(0))
    fn05C9 = Left(l146A, extfn0BB1(l146A, Len(l146A)))
End Function

Sub sub05DC (p146E As Form, p1472 As Control, p1474 As Long)
Dim l1476
Dim l1478
Dim l147A As Variant
For l1476 = 0 To p146E.ScaleWidth Step p1472.Width - l147A
   For l1478 = p1474 To p146E.ScaleHeight Step p1472.Height
      sub04C8 p146E, p1472, l1476, l1478
   Next l1478
Next l1476
If  p146E.AutoRedraw Then p146E.Refresh
End Sub

Sub sub05EE (p1480 As Form, p1484 As Control, p1486 As Long, p1488 As Long, p148A As Long, p148C As Long)
Dim l148E
Dim l1490
Dim l1492 As Long
Dim l1494 As Long
Dim l1496 As Long
Dim l1498 As Long
Dim l149A As Variant
Dim l149E As Variant
l1496 = p1484.Width - l149A
l1498 = p1484.Height - l149E
For l148E = p1486 To p1486 + p148A Step l1496
   For l1490 = p1488 To p1488 + p148C Step l1498
      If  l1490 + l1498 <= p1488 + p148C Then
         l1494 = l1498
      Else
         l1494 = p1488 + p148C - (l1490 - l149E)
      End If
      If  l148E + l1496 <= p1486 + p148A Then
         l1492 = l1496
      Else
         l1492 = p1486 + p148A - (l148E - l149A)
      End If
      sub04C8 p1480, p1484, l148E, l1490
   Next l1490
Next l148E
If  p1480.AutoRedraw Then p1480.Refresh
End Sub

Sub sub05FB (p14A4 As Form, p14A8 As Control)
Dim l14AA
For l14AA = 0 To p14A4.ScaleWidth Step p14A8.Width
   sub04C8 p14A4, p14A8, l14AA, 0
Next l14AA
If  p14A4.AutoRedraw Then p14A4.Refresh
End Sub

Sub sub060A (p14AE As Variant)
Dim l14B2 As Variant
Dim l14B6 As Integer
Dim l14B8 As Variant
l14B2 = extfn0B6A()
Do
    l14B6% = DoEvents()
    l14B8 = extfn0B6A()
Loop Until l14B8 - l14B2 >= p14AE * 1000
End Sub

Function fn0615 (p14C0 As Variant) As Variant
If  InStr(p14C0, Chr$(0)) = 0 Then
    fn0615 = p14C0
    Exit Function
End If
fn0615 = Trim$(Left$(p14C0, InStr(p14C0, Chr$(0)) - 1))
End Function

Sub sub0621 ()
Dim l14C4 As Integer
Dim l14C6 As Integer
l14C4 = extfn0742(App.EXEName)
l14C6 = extfn077B(l14C4)
End Sub

Sub sub062E ()
Dim l14CC As Integer
Dim l14CE As Integer
l14CC = extfn0742(App.EXEName)
l14CE = extfn076A(l14CC)
l14CE = extfn0755(l14CC)
End Sub

Sub sub063A (p14D2 As Control)
Dim l14D4 As Variant
l14D4 = extfn0AF3(p14D2.hWnd, gc0CE8, 0, 0)
End Sub

Sub sub0642 ()
Dim l14D8 As Integer
Dim l14DA As Integer
Dim l14DC As Integer
Dim l14DE As Integer
Dim l14E0 As Variant
l14D8% = extfn0A11()
l14DA% = DoEvents()
l14DC% = extfn0A11()
Do Until l14DE% <> l14DC%
    l14E0 = DoEvents()
    l14DE% = extfn0A11()
Loop
End Sub

Sub sub0653 ()
Dim l14E4 As Integer
Dim l14E6 As Integer
Dim l14E8 As Variant
Do
    l14E4% = DoEvents()
    l14E6% = extfn0B02("#32770", 0&)
Loop Until l14E6% <> 0
l14E8 = extfn0AB5(l14E6%, &H10, 0, 0)
End Sub

Function fn0660 () As String
Dim l14EE As String
l14EE = String(145, Chr(0))
fn0660 = Left(l14EE, extfn0B9A(l14EE, Len(l14EE)))
End Function

Function fn0674 () As Variant
Dim l14F4 As Long
Dim l14F6 As Long
l14F4 = extfn0D5E()
l14F6 = l14F4 And &HFFFF&
fn0674 = Format((l14F6 Mod 256) + ((l14F6 \ 256) / 100), "Fixed")
End Function

Sub sub0686 (p14F8 As Variant, p14FC As Variant)
Dim l1500 As Variant
If  p14F8 = "" Then Exit Sub
If  p14FC = "" Then Exit Sub
l1500 = FreeFile
Open p14FC For Output As #l1500
    Print #l1500, p14F8
Close #l1500
End Sub

Sub sub0695 (p1504 As Variant, p1508 As Variant, p150C As Variant, p1510 As Variant)
Dim l1514 As Variant
l1514 = extfn081E(p1504, p1508, p150C, p1510)
End Sub
