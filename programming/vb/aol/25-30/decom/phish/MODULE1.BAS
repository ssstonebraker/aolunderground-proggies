' Module1
Option Explicit
Declare Function extfn33C0 Lib "GDI" Alias "AddFontResource" () As Integer
Declare Sub extsub23C8 Lib "User" Alias "AdjustWindowRect" ()
Declare Sub extsub2400 Lib "User" Alias "AdjustWindowRectEx" ()
Declare Function extfnD08 Lib "User" Alias "AnyPopup" () As Integer
Declare Function extfn138 Lib "baka.dll" Alias "AOLGetList" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1C90 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn3CF0 Lib "User" Alias "ArrangeIconicWindows" () As Integer
Declare Function extfnBF0 Lib "User" Alias "BeginDeferWindowPos" () As Variant
Declare Function extfn1E88 Lib "User" Alias "BeginPaint" () As Integer
Declare Sub extsubD40 Lib "User" Alias "BringWindowToTop" ()
Declare Function extfn3A88 Lib "User" Alias "BuildCommDCB" () As Integer
Declare Function extfn10F8 Lib "User" Alias "CallMsgFilter" () As Integer
Declare Function extfn1248 Lib "User" Alias "ChangeClipboardChain" () As Integer
Declare Function extfn1948 Lib "User" Alias "ChangeMenu" () As Integer
Declare Sub extsubF00 Lib "User" Alias "CheckDlgButton" ()
Declare Function extfn1B40 Lib "User" Alias "CheckMenuItem" () As Integer
Declare Sub extsubF38 Lib "User" Alias "CheckRadioButton" ()
Declare Function extfn2828 Lib "User" Alias "ChildWindowFromPoint" () As Integer
Declare Function extfn3BA0 Lib "User" Alias "ClearCommBreak" () As Integer
Declare Sub extsub2780 Lib "User" Alias "ClientToScreen" ()
Declare Sub extsub2588 Lib "User" Alias "ClipCursor" ()
Declare Function extfn1168 Lib "User" Alias "CloseClipboard" () As Integer
Declare Function extfn3A50 Lib "User" Alias "CloseComm" () As Integer
Declare Sub extsub35B8 Lib "Sound" Alias "CloseSound" ()
Declare Sub extsubB80 Lib "User" Alias "CloseWindow" ()
Declare Function extfn2B38 Lib "User" Alias "CopyRect" () As Integer
Declare Function extfn1328 Lib "User" Alias "CountClipboardFormats" () As Integer
Declare Function extfn3820 Lib "Sound" Alias "CountVoiceNotes" () As Integer
Declare Sub extsub25C0 Lib "User" Alias "CreateCaret" ()
Declare Function extfn3270 Lib "User" Alias "CreateCursor" () As Integer
Declare Function extfn3318 Lib "User" Alias "CreateIcon" () As Integer
Declare Function extfn1A98 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn1AD0 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn988 Lib "User" Alias "CreateWindow" () As Integer
Declare Function extfn9C0 Lib "User" Alias "CreateWindowEx" () As Integer
Declare Function extfn10C0 Lib "User" Alias "DefDlgProc" () As Long
Declare Function extfnC28 Lib "User" Alias "DeferWindowPos" () As Variant
Declare Function extfn3C48 Lib "User" Alias "DefFrameProc" () As Long
Declare Function extfn3C80 Lib "User" Alias "DefMDIChildProc" () As Long
Declare Function extfn800 Lib "User" Alias "DefWindowProc" () As Long
Declare Function extfn1D38 Lib "User" Alias "DeleteMenu" () As Integer
Declare Sub extsub2668 Lib "User" Alias "DestroyCaret" ()
Declare Function extfn32A8 Lib "User" Alias "DestroyCursor" () As Integer
Declare Function extfn3350 Lib "User" Alias "DestroyIcon" () As Integer
Declare Function extfn1B08 Lib "User" Alias "DestroyMenu" () As Integer
Declare Function extfnA68 Lib "User" Alias "DestroyWindow" () As Integer
Declare Function extfn288 Lib "User" Alias "DispatchMessage" () As Long
Declare Function extfn34A0 Lib "User" Alias "DlgDirList" () As Integer
Declare Function extfn3510 Lib "User" Alias "DlgDirListComboBox" () As Integer
Declare Function extfn34D8 Lib "User" Alias "DlgDirSelect" () As Integer
Declare Function extfn3548 Lib "User" Alias "DlgDirSelectComboBox" () As Integer
Declare Sub extsub29E8 Lib "User" Alias "DrawFocusRect" ()
Declare Sub extsub1A28 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn13D0 Lib "User" Alias "EmptyClipboard" () As Integer
Declare Function extfn1600 Lib "User" Alias "EnableHardwareInput" () As Integer
Declare Function extfn1B78 Lib "User" Alias "EnableMenuItem" () As Integer
Declare Function extfn1750 Lib "User" Alias "EnableWindow" () As Integer
Declare Sub extsubC60 Lib "User" Alias "EndDeferWindowPos" ()
Declare Sub extsubDB0 Lib "User" Alias "EndDialog" ()
Declare Sub extsub1EC0 Lib "User" Alias "EndPaint" ()
Declare Function extfn1360 Lib "User" Alias "EnumClipboardFormats" () As Integer
Declare Function extfn2C88 Lib "User" Alias "EqualRect" () As Integer
Declare Function extfn3C10 Lib "User" Alias "EscapeCommFunction" () As Integer
Declare Function extfn1F68 Lib "User" Alias "ExcludeUpdateRgn" () As Integer
Declare Function extfn560 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn2A20 Lib "User" Alias "FillRect" () As Integer
Declare Function extfn28D0 Lib "GDI" Alias "FillRgn" () As Integer
Declare Function extfnC8 Lib "phish.dll" Alias "FindChildByClass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn100 Lib "phish.dll" Alias "FindChildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3078 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfnAD8 Lib "User" Alias "FlashWindow" () As Integer
Declare Function extfn3BD8 Lib "User" Alias "FlushComm" () As Integer
Declare Function extfn2A58 Lib "User" Alias "FrameRect" () As Integer
Declare Function extfn2908 Lib "GDI" Alias "FrameRgn" () As Integer
Declare Function extfn14E8 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn1558 Lib "User" Alias "GetAsyncKeyState" () As Integer
Declare Function extfn1670 Lib "User" Alias "GetCapture" () As Integer
Declare Function extfn25F8 Lib "User" Alias "GetCaretBlinkTime" () As Integer
Declare Sub extsub2748 Lib "User" Alias "GetCaretPos" ()
Declare Function extfn2F28 Lib "User" Alias "GetClassLong" () As Long
Declare Function extfn30B0 Lib "User" Alias "GetClassName" () As Integer
Declare Function extfn2EB8 Lib "User" Alias "GetClassWord" () As Integer
Declare Sub extsub2358 Lib "User" Alias "GetClientRect" ()
Declare Function extfn12B8 Lib "User" Alias "GetClipboardData" () As Integer
Declare Function extfn1398 Lib "User" Alias "GetClipboardFormatName" () As Integer
Declare Function extfn11A0 Lib "User" Alias "GetClipboardOwner" () As Integer
Declare Function extfn1210 Lib "User" Alias "GetClipboardViewer" () As Integer
Declare Function extfn3B30 Lib "User" Alias "GetCommEventMask" () As Integer
Declare Function extfn3970 Lib "User" Alias "GetCommState" () As Integer
Declare Function extfn2D68 Lib "User" Alias "GetCurrentTime" () As Long
Declare Sub extsub2550 Lib "User" Alias "GetCursorPos" ()
Declare Function extfn2F98 Lib "User" Alias "GetDesktopHwnd" () As Integer
Declare Function extfn2FD0 Lib "User" Alias "GetDesktopWindow" () As Integer
Declare Function extfn1088 Lib "User" Alias "GetDialogBaseUnits" () As Long
Declare Function extfn1050 Lib "User" Alias "GetDlgCtrlID" () As Integer
Declare Function extfnDE8 Lib "User" Alias "GetDlgItem" () As Integer
Declare Function extfnE58 Lib "User" Alias "GetDlgItemInt" () As Integer
Declare Function extfnEC8 Lib "User" Alias "GetDlgItemText" () As Integer
Declare Function extfn8A8 Lib "User" Alias "GetDoubleClickTime" () As Integer
Declare Function extfn14B0 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn1638 Lib "User" Alias "GetInputState" () As Integer
Declare Sub extsub1590 Lib "User" Alias "GetKeyboardState" ()
Declare Function extfn1520 Lib "User" Alias "GetKeyState" () As Integer
Declare Function extfn3190 Lib "User" Alias "GetLastActivePopup" () As Integer
Declare Function extfn18D8 Lib "User" Alias "GetMenu" () As Integer
Declare Function extfn1DA8 Lib "User" Alias "GetMenuCheckMarkDimensions" () As Long
Declare Function extfn1C20 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn1BE8 Lib "User" Alias "GetMenuItemID" () As Integer
Declare Function extfn19F0 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn19B8 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn218 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn5D0 Lib "User" Alias "GetMessagePos" () As Long
Declare Function extfn608 Lib "User" Alias "GetMessageTime" () As Long
Declare Function extfnFE0 Lib "User" Alias "GetNextDlgGroupItem" () As Integer
Declare Function extfn1018 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn3120 Lib "User" Alias "GetNextWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn3008 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1440 Lib "User" Alias "GetPriorityClipboardFormat" () As Integer
Declare Function extfn2240 Lib "User" Alias "GetProp" () As Integer
Declare Function extfn2128 Lib "User" Alias "GetScrollPos" () As Integer
Declare Sub extsub2198 Lib "User" Alias "GetScrollRange" ()
Declare Function extfn1BB0 Lib "User" Alias "GetSubMenu" () As Integer
Declare Function extfn2860 Lib "User" Alias "GetSysColor" () As Long
Declare Function extfn640 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn1A60 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn1830 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn3858 Lib "Sound" Alias "GetThresholdEvent" () As Integer
Declare Function extfn3890 Lib "Sound" Alias "GetThresholdStatus" () As Integer
Declare Function extfn2DA0 Lib "User" Alias "GetTickCount" () As Long
Declare Function extfn30E8 Lib "User" Alias "GetTopWindow" (ByVal p1%) As Integer
Declare Function extfn1EF8 Lib "User" Alias "GetUpdateRect" () As Integer
Declare Function extfn1F30 Lib "User" Alias "GetUpdateRgn" () As Integer
Declare Function extfn31C8 Lib "User" Alias "GetWindow" () As Integer
Declare Function extfn2E48 Lib "User" Alias "GetWindowLong" () As Long
Declare Sub extsub2390 Lib "User" Alias "GetWindowRect" ()
Declare Function extfn3158 Lib "User" Alias "GetWindowTask" () As Integer
Declare Function extfn22E8 Lib "User" Alias "GetWindowText" () As Integer
Declare Function extfn2320 Lib "User" Alias "GetWindowTextLength" () As Integer
Declare Function extfn2DD8 Lib "User" Alias "GetWindowWord" () As Integer
Declare Sub extsub26A0 Lib "User" Alias "HideCaret" ()
Declare Function extfn1980 Lib "User" Alias "HiliteMenuItem" () As Integer
Declare Sub extsub2B70 Lib "User" Alias "InflateRect" ()
Declare Function extfn870 Lib "User" Alias "InSendMessage" () As Integer
Declare Function extfn1C58 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn2BA8 Lib "User" Alias "IntersectRect" () As Integer
Declare Sub extsub1FA0 Lib "User" Alias "InvalidateRect" ()
Declare Sub extsub2010 Lib "User" Alias "InvalidateRgn" ()
Declare Sub extsub2A90 Lib "User" Alias "InvertRect" ()
Declare Function extfn2940 Lib "GDI" Alias "InvertRgn" () As Integer
Declare Function extfnA30 Lib "User" Alias "IsChild" () As Integer
Declare Function extfn1408 Lib "User" Alias "IsClipboardFormatAvailable" () As Integer
Declare Function extfn3430 Lib "User" Alias "IsDialogMessage" () As Integer
Declare Function extfnF70 Lib "User" Alias "IsDlgButtonChecked" () As Integer
Declare Function extfnCD0 Lib "User" Alias "IsIconic" () As Integer
Declare Function extfn2C50 Lib "User" Alias "IsRectEmpty" () As Integer
Declare Function extfn9F8 Lib "User" Alias "IsWindow" () As Integer
Declare Function extfn1788 Lib "User" Alias "IsWindowEnabled" () As Integer
Declare Function extfnC98 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfnD78 Lib "User" Alias "IsZoomed" () As Integer
Declare Function extfn1718 Lib "User" Alias "KillTimer" () As Integer
Declare Function extfn448 Lib "Kernel" Alias "_lclose" () As Integer
Declare Function extfn480 Lib "Kernel" Alias "_lcreat" () As Integer
Declare Function extfn4B8 Lib "Kernel" Alias "_llseek" () As Long
Declare Function extfn17C0 Lib "User" Alias "LoadAccelerators" () As Integer
Declare Function extfn3200 Lib "User" Alias "LoadBitmap" () As Integer
Declare Function extfn3238 Lib "User" Alias "LoadCursor" () As Integer
Declare Function extfn32E0 Lib "User" Alias "LoadIcon" () As Integer
Declare Function extfn1868 Lib "User" Alias "LoadMenu" () As Integer
Declare Function extfn18A0 Lib "User" Alias "LoadMenuIndirect" () As Integer
Declare Function extfn3388 Lib "User" Alias "LoadString" () As Integer
Declare Function extfn410 Lib "Kernel" Alias "_lopen" () As Integer
Declare Function extfn4F0 Lib "Kernel" Alias "_lread" () As Integer
Declare Function extfn3A0 Lib "Kernel" Alias "lstrcat" () As Long
Declare Function extfn2F8 Lib "User" Alias "lstrcmp" () As Integer
Declare Function extfn330 Lib "User" Alias "lstrcmpi" () As Integer
Declare Function extfn368 Lib "Kernel" Alias "lstrcpy" () As Long
Declare Function extfn3D8 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn528 Lib "Kernel" Alias "_lwrite" () As Integer
Declare Sub extsub3468 Lib "User" Alias "MapDialogRect" ()
Declare Sub extsub2470 Lib "User" Alias "MessageBeep" ()
Declare Function extfn2438 Lib "User" Alias "MessageBox" () As Integer
Declare Function extfn1CC8 Lib "User" Alias "ModifyMenu" () As Integer
Declare Sub extsubBB8 Lib "User" Alias "MoveWindow" ()
Declare Sub extsub2C18 Lib "User" Alias "OffsetRect" ()
Declare Function extfn1130 Lib "User" Alias "OpenClipboard" () As Integer
Declare Function extfn3900 Lib "User" Alias "OpenComm" () As Integer
Declare Function extfnB48 Lib "User" Alias "OpenIcon" () As Integer
Declare Function extfn3580 Lib "Sound" Alias "OpenSound" () As Integer
Declare Function extfn2978 Lib "GDI" Alias "PaintRgn" () As Integer
Declare Function extfn2C0 Lib "User" Alias "PeekMessage" () As Integer
Declare Function extfn758 Lib "User" Alias "PostAppMessage" () As Integer
Declare Function extfn720 Lib "User" Alias "PostMessage" () As Integer
Declare Sub extsub838 Lib "User" Alias "PostQuitMessage" ()
Declare Sub extsub3E78 Lib "User" Alias "ProfClear" ()
Declare Sub extsub3EE8 Lib "User" Alias "ProfFinish" ()
Declare Sub extsub3EB0 Lib "User" Alias "ProfFlush" ()
Declare Function extfn3D60 Lib "User" Alias "ProfInsChk" () As Integer
Declare Sub extsub3DD0 Lib "User" Alias "ProfSampRate" ()
Declare Sub extsub3D98 Lib "User" Alias "ProfSetup" ()
Declare Sub extsub3E08 Lib "User" Alias "ProfStart" ()
Declare Sub extsub3E40 Lib "User" Alias "ProfStop" ()
Declare Function extfn2CC0 Lib "User" Alias "PtInRect" () As Integer
Declare Function extfn29B0 Lib "GDI" Alias "PtInRegion" () As Integer
Declare Function extfn39A8 Lib "User" Alias "ReadComm" () As Integer
Declare Function extfn2D30 Lib "User" Alias "RectInRegion" () As Integer
Declare Function extfn2CF8 Lib "User" Alias "RectVisible" () As Integer
Declare Function extfn12F0 Lib "User" Alias "RegisterClipboardFormat" () As Integer
Declare Function extfn1E0 Lib "User" Alias "RegisterWindowMessage" () As Integer
Declare Sub extsub16E0 Lib "User" Alias "ReleaseCapture" ()
Declare Function extfn33F8 Lib "GDI" Alias "RemoveFontResource" () As Integer
Declare Function extfn1D00 Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn2278 Lib "User" Alias "RemoveProp" () As Integer
Declare Sub extsub790 Lib "User" Alias "ReplyMessage" ()
Declare Sub extsub27B8 Lib "User" Alias "ScreenToClient" ()
Declare Function extfn20B8 Lib "User" Alias "ScrollDC" () As Integer
Declare Sub extsub2080 Lib "User" Alias "ScrollWindow" ()
Declare Function extfnFA8 Lib "User" Alias "SendDlgItemMessage" () As Long
Declare Function extfn6E8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4&) As Long
Declare Function extfn6B0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn1E50 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn16A8 Lib "User" Alias "SetCapture" () As Integer
Declare Sub extsub2630 Lib "User" Alias "SetCaretBlinkTime" ()
Declare Sub extsub2710 Lib "User" Alias "SetCaretPos" ()
Declare Function extfn2F60 Lib "User" Alias "SetClassLong" () As Long
Declare Function extfn2EF0 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfn1280 Lib "User" Alias "SetClipboardData" () As Integer
Declare Function extfn11D8 Lib "User" Alias "SetClipboardViewer" () As Integer
Declare Function extfn3B68 Lib "User" Alias "SetCommBreak" () As Integer
Declare Function extfn3AF8 Lib "User" Alias "SetCommEventMask" () As Long
Declare Function extfn3938 Lib "User" Alias "SetCommState" () As Integer
Declare Function extfn2518 Lib "User" Alias "SetCursor" () As Integer
Declare Sub extsub24E0 Lib "User" Alias "SetCursorPos" ()
Declare Sub extsubE20 Lib "User" Alias "SetDlgItemInt" ()
Declare Sub extsubE90 Lib "User" Alias "SetDlgItemText" ()
Declare Sub extsub8E0 Lib "User" Alias "SetDoubleClickTime" ()
Declare Function extfn1478 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Sub extsub15C8 Lib "User" Alias "SetKeyboardState" ()
Declare Function extfn1910 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn1D70 Lib "User" Alias "SetMenuItemBitmaps" () As Integer
Declare Function extfn950 Lib "User" Alias "SetMessageQueue" () As Integer
Declare Function extfn3040 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn2208 Lib "User" Alias "SetProp" () As Integer
Declare Sub extsub2AC8 Lib "User" Alias "SetRect" ()
Declare Sub extsub2B00 Lib "User" Alias "SetRectEmpty" ()
Declare Function extfn20F0 Lib "User" Alias "SetScrollPos" () As Integer
Declare Sub extsub2160 Lib "User" Alias "SetScrollRange" ()
Declare Function extfn36D0 Lib "Sound" Alias "SetSoundNoise" () As Integer
Declare Sub extsub2898 Lib "User" Alias "SetSysColors" ()
Declare Function extfn678 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn3660 Lib "Sound" Alias "SetVoiceAccent" () As Integer
Declare Function extfn3698 Lib "Sound" Alias "SetVoiceEnvelope" () As Integer
Declare Function extfn3628 Lib "Sound" Alias "SetVoiceNote" () As Integer
Declare Function extfn35F0 Lib "Sound" Alias "SetVoiceQueueSize" () As Integer
Declare Function extfn3708 Lib "Sound" Alias "SetVoiceSound" () As Integer
Declare Function extfn38C8 Lib "Sound" Alias "SetVoiceThreshold" () As Integer
Declare Function extfn2E80 Lib "User" Alias "SetWindowLong" () As Long
Declare Sub extsub22B0 Lib "User" Alias "SetWindowText" ()
Declare Function extfn2E10 Lib "User" Alias "SetWindowWord" () As Integer
Declare Sub extsub26D8 Lib "User" Alias "ShowCaret" ()
Declare Function extfn24A8 Lib "User" Alias "ShowCursor" () As Integer
Declare Sub extsubB10 Lib "User" Alias "ShowOwnedPopups" ()
Declare Sub extsub21D0 Lib "User" Alias "ShowScrollBar" ()
Declare Function extfnAA0 Lib "User" Alias "ShowWindow" () As Integer
Declare Function extfn170 Lib "mmsystem.dll" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn3740 Lib "Sound" Alias "StartSound" () As Integer
Declare Function extfn3778 Lib "Sound" Alias "StopSound" () As Integer
Declare Function extfn598 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn37E8 Lib "Sound" Alias "SyncAllVoices" () As Integer
Declare Function extfn1DE0 Lib "User" Alias "TrackPopupMenu" () As Integer
Declare Function extfn17F8 Lib "User" Alias "TranslateAccelerator" () As Integer
Declare Function extfn3CB8 Lib "User" Alias "TranslateMDISysAccel" () As Integer
Declare Function extfn250 Lib "User" Alias "TranslateMessage" () As Integer
Declare Function extfn3AC0 Lib "User" Alias "TransmitCommChar" () As Integer
Declare Function extfn39E0 Lib "User" Alias "UngetCommChar" () As Integer
Declare Function extfn2BE0 Lib "User" Alias "UnionRect" () As Integer
Declare Function extfn918 Lib "User" Alias "UnregisterClass" () As Integer
Declare Sub extsub1E18 Lib "User" Alias "UpdateWindow" ()
Declare Sub extsub1FD8 Lib "User" Alias "ValidateRect" ()
Declare Sub extsub2048 Lib "User" Alias "ValidateRgn" ()
Declare Sub extsub7C8 Lib "User" Alias "WaitMessage" ()
Declare Function extfn37B0 Lib "Sound" Alias "WaitSoundState" () As Integer
Declare Function extfn27F0 Lib "User" Alias "WindowFromPoint" () As Integer
Declare Function extfn3D28 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn3A18 Lib "User" Alias "WriteComm" () As Integer
Declare Function extfn1A8 Lib "User" Alias "wvsprintf" () As Integer

Function fn3F20 (ByVal p0884 As String) As Integer
Dim l0886 As Variant
Dim l088A As Variant
  l0886 = fn3F90()
  l088A = extfnC8(l0886, p0884)
  fn3F20 = l088A
End Function

Function fn3F58 (ByVal p0890 As String) As Integer
Dim l0892 As Variant
Dim l0896 As Variant
  l0892 = fn3F90()
  l0896 = extfn100(l0892, p0890)
  fn3F58 = l0896
End Function

Function fn3F90 () As Integer
Dim l089C As Variant
  l089C = extfn3078("AOL Frame25", "America  Online")
  fn3F90 = l089C
End Function

Sub sub3FC8 (ByVal p08A0 As String)
  sub4000
  DoEvents
  SendKeys p08A0
End Sub

Sub sub4000 ()
Dim l08A2 As Variant
Dim l08A6 As Variant
  l08A2 = fn3F90()
  l08A6 = extfn1478(l08A2)
End Sub

Function fn4038 (ByVal p08AC) As Integer
Dim l08AE As Variant
  While l08AE = 0
    DoEvents
    l08AE = fn3F58(p08AC)
  Wend
  fn4038 = l08AE
End Function

Function fn4070 (p08B4) As Integer
Dim l08B6 As Variant
Dim l08BA As Variant
Dim l08BE As Variant
  l08B6 = 1
  While l08B6 <> 0
    DoEvents
    l08B6 = fn3F58(p08B4)
  Wend
  l08BA = fn3F20("MDI Client")
  l08BE = extfn30E8(l08BA)
  fn4070 = l08BE
End Function

Function fn40A8 (p08C4 As Integer) As Integer
Dim l08C6 As Variant
Dim l08CA As Variant
  While l08C6 = 0
    DoEvents
    l08C6 = extfn3078("#32770", "America Online")
  Wend
  If  p08C4 = 1 Then
    l08CA = extfn1478(l08C6)
    SendKeys " "
  End If
End Function
