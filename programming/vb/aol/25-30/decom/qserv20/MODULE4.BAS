' Module4
Option Explicit
Declare Function extfn2D10 Lib "Kernel" Alias "AccessResource" () As Integer
Declare Function extfn2DF0 Lib "Kernel" Alias "AddAtom" () As Integer
Declare Function extfn89D0 Lib "GDI" Alias "AddFontResource" () As Integer
Declare Sub extsub79D8 Lib "User" Alias "AdjustWindowRect" ()
Declare Sub extsub7A10 Lib "User" Alias "AdjustWindowRectEx" ()
Declare Function extfn2AE0 Lib "Kernel" Alias "AllocDStoCSAlias" () As Integer
Declare Function extfn2CA0 Lib "Kernel" Alias "AllocResource" () As Integer
Declare Function extfn2B18 Lib "Kernel" Alias "AllocSelector" () As Integer
Declare Sub extsub55F8 Lib "GDI" Alias "AnimatePalette" ()
Declare Function extfn33D8 Lib "User" Alias "AnsiLower" () As Long
Declare Function extfn3410 Lib "User" Alias "AnsiLowerBuff" () As Integer
Declare Function extfn3448 Lib "User" Alias "AnsiNext" () As Long
Declare Function extfn3480 Lib "User" Alias "AnsiPrev" () As Long
Declare Function extfn3288 Lib "Keyboard" Alias "AnsiToOem" () As Integer
Declare Sub extsub32F8 Lib "Keyboard" Alias "AnsiToOemBuff" ()
Declare Function extfn3368 Lib "User" Alias "AnsiUpper" () As String
Declare Function extfn33A0 Lib "User" Alias "AnsiUpperBuff" () As Integer
Declare Function extfn6318 Lib "User" Alias "AnyPopup" () As Integer
Declare Function extfn72A0 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn3C60 Lib "GDI" Alias "Arc" () As Integer
Declare Function extfn9300 Lib "User" Alias "ArrangeIconicWindows" () As Integer
Declare Function extfn6200 Lib "User" Alias "BeginDeferWindowPos" () As Variant
Declare Function extfn7498 Lib "User" Alias "BeginPaint" () As Integer
Declare Function extfn3D40 Lib "GDI" Alias "BitBlt" () As Integer
Declare Sub extsub6350 Lib "User" Alias "BringWindowToTop" ()
Declare Function extfn9098 Lib "User" Alias "BuildCommDCB" () As Integer
Declare Function extfn6708 Lib "User" Alias "CallMsgFilter" () As Integer
Declare Function extfn31A8 Lib "Kernel" Alias "Catch" () As Integer
Declare Function extfn6858 Lib "User" Alias "ChangeClipboardChain" () As Integer
Declare Function extfn6F58 Lib "User" Alias "ChangeMenu" () As Integer
Declare Function extfn2B50 Lib "Kernel" Alias "ChangeSelector" () As Integer
Declare Sub extsub6510 Lib "User" Alias "CheckDlgButton" ()
Declare Function extfn7150 Lib "User" Alias "CheckMenuItem" () As Integer
Declare Sub extsub6548 Lib "User" Alias "CheckRadioButton" ()
Declare Function extfn7E38 Lib "User" Alias "ChildWindowFromPoint" () As Integer
Declare Function extfn3C98 Lib "GDI" Alias "Chord" () As Integer
Declare Function extfn91B0 Lib "User" Alias "ClearCommBreak" () As Integer
Declare Sub extsub7D90 Lib "User" Alias "ClientToScreen" ()
Declare Sub extsub7B98 Lib "User" Alias "ClipCursor" ()
Declare Function extfn6778 Lib "User" Alias "CloseClipboard" () As Integer
Declare Function extfn9060 Lib "User" Alias "CloseComm" () As Integer
Declare Function extfn5320 Lib "GDI" Alias "CloseMetaFile" () As Integer
Declare Sub extsub8BC8 Lib "Sound" Alias "CloseSound" ()
Declare Sub extsub6158 Lib "User" Alias "CloseWindow" ()
Declare Function extfn4E18 Lib "GDI" Alias "CombineRgn" () As Integer
Declare Function extfn5080 Lib "GDI" Alias "CopyMetaFile" () As Integer
Declare Function extfn8148 Lib "User" Alias "CopyRect" () As Integer
Declare Function extfn6938 Lib "User" Alias "CountClipboardFormats" () As Integer
Declare Function extfn8E30 Lib "Sound" Alias "CountVoiceNotes" () As Integer
Declare Function extfn4130 Lib "GDI" Alias "CreateBitmap" () As Integer
Declare Function extfn4168 Lib "GDI" Alias "CreateBitmapIndirect" () As Integer
Declare Function extfn40F8 Lib "GDI" Alias "CreateBrushIndirect" () As Integer
Declare Sub extsub7BD0 Lib "User" Alias "CreateCaret" ()
Declare Function extfn41A0 Lib "GDI" Alias "CreateCompatibleBitmap" () As Integer
Declare Function extfn38E0 Lib "GDI" Alias "CreateCompatibleDC" () As Integer
Declare Function extfn8880 Lib "User" Alias "CreateCursor" () As Integer
Declare Function extfn3870 Lib "GDI" Alias "CreateDC" () As Integer
Declare Function extfn5470 Lib "GDI" Alias "CreateDIBitmap" () As Integer
Declare Function extfn54A8 Lib "GDI" Alias "CreateDIBPatternBrush" () As Integer
Declare Function extfn41D8 Lib "GDI" Alias "CreateDiscardableBitmap" () As Integer
Declare Function extfn4478 Lib "GDI" Alias "CreateEllipticRgn" () As Integer
Declare Function extfn4440 Lib "GDI" Alias "CreateEllipticRgnIndirect" () As Integer
Declare Function extfn42F0 Lib "GDI" Alias "CreateFont" () As Integer
Declare Function extfn4328 Lib "GDI" Alias "CreateFontIndirect" () As Integer
Declare Function extfn4018 Lib "GDI" Alias "CreateHatchBrush" () As Integer
Declare Function extfn38A8 Lib "GDI" Alias "CreateIC" () As Integer
Declare Function extfn8928 Lib "User" Alias "CreateIcon" () As Integer
Declare Function extfn70A8 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn52E8 Lib "GDI" Alias "CreateMetaFile" () As Integer
Declare Function extfn5518 Lib "GDI" Alias "CreatePalette" () As Integer
Declare Function extfn40C0 Lib "GDI" Alias "CreatePatternBrush" () As Integer
Declare Function extfn3F70 Lib "GDI" Alias "CreatePen" () As Integer
Declare Function extfn3FA8 Lib "GDI" Alias "CreatePenIndirect" () As Integer
Declare Function extfn44B0 Lib "GDI" Alias "CreatePolygonRgn" () As Integer
Declare Function extfn44E8 Lib "GDI" Alias "CreatePolyPolygonRgn" () As Integer
Declare Function extfn70E0 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn4398 Lib "GDI" Alias "CreateRectRgn" () As Integer
Declare Function extfn4408 Lib "GDI" Alias "CreateRectRgnIndirect" () As Integer
Declare Function extfn4520 Lib "GDI" Alias "CreateRoundRectRgn" () As Integer
Declare Function extfn3FE0 Lib "GDI" Alias "CreateSolidBrush" () As Integer
Declare Function extfn5F60 Lib "User" Alias "CreateWindow" () As Integer
Declare Function extfn5F98 Lib "User" Alias "CreateWindowEx" () As Integer
Declare Sub extsub2060 Lib "Kernel" Alias "DebugBreak" ()
Declare Function extfn66D0 Lib "User" Alias "DefDlgProc" () As Long
Declare Function extfn6238 Lib "User" Alias "DeferWindowPos" () As Variant
Declare Function extfn9258 Lib "User" Alias "DefFrameProc" () As Long
Declare Function extfn9290 Lib "User" Alias "DefMDIChildProc" () As Long
Declare Function extfn5DD8 Lib "User" Alias "DefWindowProc" () As Long
Declare Function extfn2E28 Lib "Kernel" Alias "DeleteAtom" () As Integer
Declare Function extfn3918 Lib "GDI" Alias "DeleteDC" () As Integer
Declare Function extfn7348 Lib "User" Alias "DeleteMenu" () As Integer
Declare Function extfn5048 Lib "GDI" Alias "DeleteMetaFile" () As Integer
Declare Function extfn4590 Lib "GDI" Alias "DeleteObject" () As Integer
Declare Sub extsub7C78 Lib "User" Alias "DestroyCaret" ()
Declare Function extfn88B8 Lib "User" Alias "DestroyCursor" () As Integer
Declare Function extfn8960 Lib "User" Alias "DestroyIcon" () As Integer
Declare Function extfn7118 Lib "User" Alias "DestroyMenu" () As Integer
Declare Function extfn6040 Lib "User" Alias "DestroyWindow" () As Integer
Declare Function extfn5898 Lib "User" Alias "DispatchMessage" () As Long
Declare Function extfn8AB0 Lib "User" Alias "DlgDirList" () As Integer
Declare Function extfn8B20 Lib "User" Alias "DlgDirListComboBox" () As Integer
Declare Function extfn8AE8 Lib "User" Alias "DlgDirSelect" () As Integer
Declare Function extfn8B58 Lib "User" Alias "DlgDirSelectComboBox" () As Integer
Declare Function extfn5278 Lib "GDI" Alias "DPtoLP" () As Integer
Declare Sub extsub7FF8 Lib "User" Alias "DrawFocusRect" ()
Declare Function extfn3790 Lib "User" Alias "DrawIcon" () As Integer
Declare Sub extsub7038 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn3758 Lib "User" Alias "DrawText" () As Integer
Declare Function extfn3C28 Lib "GDI" Alias "Ellipse" () As Integer
Declare Function extfn69E0 Lib "User" Alias "EmptyClipboard" () As Integer
Declare Function extfn6C10 Lib "User" Alias "EnableHardwareInput" () As Integer
Declare Function extfn7188 Lib "User" Alias "EnableMenuItem" () As Integer
Declare Function extfn6D60 Lib "User" Alias "EnableWindow" () As Integer
Declare Sub extsub6270 Lib "User" Alias "EndDeferWindowPos" ()
Declare Sub extsub63C0 Lib "User" Alias "EndDialog" ()
Declare Sub extsub74D0 Lib "User" Alias "EndPaint" ()
Declare Function extfn6970 Lib "User" Alias "EnumClipboardFormats" () As Integer
Declare Function extfn8298 Lib "User" Alias "EqualRect" () As Integer
Declare Function extfn4E50 Lib "GDI" Alias "EqualRgn" () As Integer
Declare Function extfn5128 Lib "GDI" Alias "Escape" () As Integer
Declare Function extfn9220 Lib "User" Alias "EscapeCommFunction" () As Integer
Declare Function extfn4DA8 Lib "GDI" Alias "ExcludeClipRect" () As Integer
Declare Function extfn7578 Lib "User" Alias "ExcludeUpdateRgn" () As Integer
Declare Function extfn5B70 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn3F00 Lib "GDI" Alias "ExtFloodFill" () As Integer
Declare Function extfn3AD8 Lib "GDI" Alias "ExtTextOut" () As Integer
Declare Sub extsub3218 Lib "Kernel" Alias "FatalExit" ()
Declare Function extfn8030 Lib "User" Alias "FillRect" () As Integer
Declare Function extfn7EE0 Lib "GDI" Alias "FillRgn" () As Integer
Declare Function extfn2E60 Lib "Kernel" Alias "FindAtom" () As Integer
Declare Function extfn1D88 Lib "VBWFind.dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1DC0 Lib "VBWFind.dll" Alias "FindCHildByTitle" () As Integer
Declare Function extfn2BC0 Lib "Kernel" Alias "FindResource" () As Integer
Declare Function extfn8688 Lib "User" Alias "findwindow" (ByVal p1$, ByVal p2&) As Integer
Declare Function extfn60B0 Lib "User" Alias "FlashWindow" () As Integer
Declare Function extfn3EC8 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn91E8 Lib "User" Alias "FlushComm" () As Integer
Declare Function extfn8068 Lib "User" Alias "FrameRect" () As Integer
Declare Function extfn7F18 Lib "GDI" Alias "FrameRgn" () As Integer
Declare Sub extsub1FB8 Lib "Kernel" Alias "FreeLibrary" ()
Declare Sub extsub1F80 Lib "Kernel" Alias "FreeModule" ()
Declare Function extfn2C30 Lib "Kernel" Alias "FreeResource" () As Integer
Declare Function extfn2B88 Lib "Kernel" Alias "FreeSelector" () As Integer
Declare Function extfn6AF8 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn4830 Lib "GDI" Alias "GetAspectRatioFilter" () As Long
Declare Function extfn6B68 Lib "User" Alias "GetAsyncKeyState" () As Integer
Declare Function extfn2FB0 Lib "Kernel" Alias "GetAtomHandle" () As Integer
Declare Function extfn2E98 Lib "Kernel" Alias "GetAtomName" () As Integer
Declare Function extfn4248 Lib "GDI" Alias "GetBitmapBits" () As Long
Declare Function extfn42B8 Lib "GDI" Alias "GetBitmapDimension" () As Long
Declare Function extfn4670 Lib "GDI" Alias "GetBkColor" () As Long
Declare Function extfn46E0 Lib "GDI" Alias "GetBkMode" () As Integer
Declare Function extfn4088 Lib "GDI" Alias "GetBrushOrg" () As Long
Declare Function extfn6C80 Lib "User" Alias "GetCapture" () As Integer
Declare Function extfn7C08 Lib "User" Alias "GetCaretBlinkTime" () As Integer
Declare Sub extsub7D58 Lib "User" Alias "GetCaretPos" ()
Declare Function extfn3E20 Lib "GDI" Alias "GetCharWidth" () As Integer
Declare Function extfn8538 Lib "User" Alias "GetClassLong" () As Long
Declare Function extfn86C0 Lib "User" Alias "GetClassName" () As Integer
Declare Function extfn84C8 Lib "User" Alias "GetClassWord" () As Integer
Declare Sub extsub7968 Lib "User" Alias "GetClientRect" ()
Declare Function extfn68C8 Lib "User" Alias "GetClipboardData" () As Integer
Declare Function extfn69A8 Lib "User" Alias "GetClipboardFormatName" () As Integer
Declare Function extfn67B0 Lib "User" Alias "GetClipboardOwner" () As Integer
Declare Function extfn6820 Lib "User" Alias "GetClipboardViewer" () As Integer
Declare Function extfn4D00 Lib "GDI" Alias "GetClipBox" () As Integer
Declare Function extfn9140 Lib "User" Alias "GetCommEventMask" () As Integer
Declare Function extfn8F80 Lib "User" Alias "GetCommState" () As Integer
Declare Function extfn2140 Lib "Kernel" Alias "GetCurrentPDB" () As Integer
Declare Function extfn39F8 Lib "GDI" Alias "GetCurrentPosition" () As Long
Declare Function extfn2D80 Lib "Kernel" Alias "GetCurrentTask" () As Integer
Declare Function extfn8378 Lib "User" Alias "GetCurrentTime" () As Long
Declare Sub extsub7B60 Lib "User" Alias "GetCursorPos" ()
Declare Function extfn3800 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn3A68 Lib "GDI" Alias "GetDCOrg" () As Long
Declare Function extfn85A8 Lib "User" Alias "GetDesktopHwnd" () As Integer
Declare Function extfn85E0 Lib "User" Alias "GetDesktopWindow" () As Integer
Declare Function extfn51D0 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn6698 Lib "User" Alias "GetDialogBaseUnits" () As Long
Declare Function extfn5400 Lib "GDI" Alias "GetDIBits" () As Integer
Declare Function extfn6660 Lib "User" Alias "GetDlgCtrlID" () As Integer
Declare Function extfn63F8 Lib "User" Alias "GetDlgItem" () As Integer
Declare Function extfn6468 Lib "User" Alias "GetDlgItemInt" () As Integer
Declare Function extfn64D8 Lib "User" Alias "GetDlgItemText" () As Integer
Declare Function extfn2AA8 Lib "Kernel" Alias "GetDOSEnvironment" () As Long
Declare Function extfn5E80 Lib "User" Alias "GetDoubleClickTime" () As Integer
Declare Function extfn2290 Lib "Kernel" Alias "GetDriveType" () As Integer
Declare Function extfn5240 Lib "GDI" Alias "GetEnvironment" () As Integer
Declare Function extfn6AC0 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn1FF0 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn6C48 Lib "User" Alias "GetInputState" () As Integer
Declare Function extfn1ED8 Lib "Kernel" Alias "GetInstanceData" () As Integer
Declare Function extfn3598 Lib "Keyboard" Alias "GetKBCodePage" () As Variant
Declare Sub extsub6BA0 Lib "User" Alias "GetKeyboardState" ()
Declare Function extfn3528 Lib "Keyboard" Alias "GetKeyboardType" () As Integer
Declare Function extfn35D0 Lib "Keyboard" Alias "GetKeyNameText" () As Integer
Declare Function extfn6B30 Lib "User" Alias "GetKeyState" () As Integer
Declare Function extfn87A0 Lib "User" Alias "GetLastActivePopup" () As Integer
Declare Function extfn4A28 Lib "GDI" Alias "GetMapMode" () As Integer
Declare Function extfn6EE8 Lib "User" Alias "GetMenu" () As Integer
Declare Function extfn73B8 Lib "User" Alias "GetMenuCheckMarkDimensions" () As Long
Declare Function extfn7230 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn71F8 Lib "User" Alias "GetMenuItemID" () As Integer
Declare Function extfn7000 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn6FC8 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn5828 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn5BE0 Lib "User" Alias "GetMessagePos" () As Long
Declare Function extfn5C18 Lib "User" Alias "GetMessageTime" () As Long
Declare Function extfn5010 Lib "GDI" Alias "GetMetaFile" () As Integer
Declare Function extfn5358 Lib "GDI" Alias "GetMetaFileBits" () As Integer
Declare Function extfn1EA0 Lib "Kernel" Alias "GetModuleFileName" () As Integer
Declare Function extfn1E30 Lib "Kernel" Alias "GetModuleHandle" () As Integer
Declare Function extfn1E68 Lib "Kernel" Alias "GetModuleUsage" () As Integer
Declare Function extfn4868 Lib "GDI" Alias "GetNearestColor" () As Long
Declare Function extfn56A0 Lib "GDI" Alias "GetNearestPaletteIndex" () As Integer
Declare Function extfn65F0 Lib "User" Alias "GetNextDlgGroupItem" () As Integer
Declare Function extfn6628 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn8730 Lib "User" Alias "GetNextWindow" () As Integer
Declare Function extfn1DF8 Lib "Kernel" Alias "GetNumTasks" () As Integer
Declare Function extfn4558 Lib "GDI" Alias "GetObject" () As Integer
Declare Function extfn5668 Lib "GDI" Alias "GetPaletteEntries" () As Integer
Declare Function extfn8618 Lib "User" Alias "GetParent" () As Integer
Declare Function extfn3E90 Lib "GDI" Alias "GetPixel" () As Long
Declare Function extfn49B8 Lib "GDI" Alias "GetPolyFillMode" () As Integer
Declare Function extfn6A50 Lib "User" Alias "GetPriorityClipboardFormat" () As Integer
Declare Function extfn3090 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn30C8 Lib "Kernel" Alias "GetPrivateProfileString" () As Integer
Declare Function extfn2FE8 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn3020 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn7850 Lib "User" Alias "GetProp" () As Integer
Declare Function extfn4EC0 Lib "GDI" Alias "GetRgnBox" () As Integer
Declare Function extfn48D8 Lib "GDI" Alias "GetROP2" () As Integer
Declare Function extfn7738 Lib "User" Alias "GetScrollPos" () As Integer
Declare Sub extsub77A8 Lib "User" Alias "GetScrollRange" ()
Declare Function extfn3F38 Lib "GDI" Alias "GetStockObject" () As Integer
Declare Function extfn4948 Lib "GDI" Alias "GetStretchBltMode" () As Integer
Declare Function extfn71C0 Lib "User" Alias "GetSubMenu" () As Integer
Declare Function extfn7E70 Lib "User" Alias "GetSysColor" () As Long
Declare Function extfn5C50 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn3170 Lib "Kernel" Alias "GetSystemDirectory" () As Integer
Declare Function extfn7070 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn6E40 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn5710 Lib "GDI" Alias "GetSystemPaletteEntries" () As Integer
Declare Function extfn5748 Lib "GDI" Alias "GetSystemPaletteUse" () As Integer
Declare Function extfn4F68 Lib "GDI" Alias "GetTabbedTextExtent" () As Long
Declare Function extfn21E8 Lib "Kernel" Alias "GetTempDrive" () As Integer
Declare Function extfn2220 Lib "Kernel" Alias "GetTempFileName" () As Integer
Declare Function extfn47C0 Lib "GDI" Alias "GetTextAlign" () As Integer
Declare Function extfn4FD8 Lib "GDI" Alias "GetTextCharacterExtra" () As Integer
Declare Function extfn4750 Lib "GDI" Alias "GetTextColor" () As Long
Declare Function extfn4F30 Lib "GDI" Alias "GetTextExtent" () As Long
Declare Function extfn5160 Lib "GDI" Alias "GetTextFace" () As Integer
Declare Function extfn5198 Lib "GDI" Alias "GetTextMetrics" () As Integer
Declare Function extfn8E68 Lib "Sound" Alias "GetThresholdEvent" () As Integer
Declare Function extfn8EA0 Lib "Sound" Alias "GetThresholdStatus" () As Integer
Declare Function extfn83B0 Lib "User" Alias "GetTickCount" () As Long
Declare Function extfn86F8 Lib "User" Alias "GetTopWindow" () As Integer
Declare Function extfn7508 Lib "User" Alias "GetUpdateRect" () As Integer
Declare Function extfn7540 Lib "User" Alias "GetUpdateRgn" () As Integer
Declare Function extfn2178 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn4BE8 Lib "GDI" Alias "GetViewportExt" () As Long
Declare Function extfn4B78 Lib "GDI" Alias "GetViewportOrg" () As Long
Declare Function extfn87D8 Lib "User" Alias "getwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn37C8 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn4B08 Lib "GDI" Alias "GetWindowExt" () As Long
Declare Function extfn8458 Lib "User" Alias "GetWindowLong" () As Long
Declare Function extfn4A98 Lib "GDI" Alias "GetWindowOrg" () As Long
Declare Sub extsub79A0 Lib "User" Alias "GetWindowRect" ()
Declare Function extfn3138 Lib "Kernel" Alias "GetWindowsDirectory" () As Integer
Declare Function extfn8768 Lib "User" Alias "GetWindowTask" () As Integer
Declare Function extfn78F8 Lib "User" Alias "GetWindowText" () As Integer
Declare Function extfn7930 Lib "User" Alias "GetWindowTextLength" () As Integer
Declare Function extfn83E8 Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn3720 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn2ED0 Lib "Kernel" Alias "GlobalAddAtom" () As Integer
Declare Function extfn22C8 Lib "Kernel" Alias "GlobalAlloc" () As Integer
Declare Function extfn2300 Lib "Kernel" Alias "GlobalCompact" () As Long
Declare Function extfn2F08 Lib "Kernel" Alias "GlobalDeleteAtom" () As Integer
Declare Function extfn2F40 Lib "Kernel" Alias "GlobalFindAtom" () As Integer
Declare Sub extsub2648 Lib "Kernel" Alias "GlobalFix" ()
Declare Function extfn24C0 Lib "Kernel" Alias "GlobalFlags" () As Integer
Declare Function extfn2338 Lib "Kernel" Alias "GlobalFree" () As Integer
Declare Function extfn2F78 Lib "Kernel" Alias "GlobalGetAtomName" () As Integer
Declare Function extfn2370 Lib "Kernel" Alias "GlobalHandle" () As Long
Declare Function extfn23A8 Lib "Kernel" Alias "GlobalLock" () As Long
Declare Function extfn2568 Lib "Kernel" Alias "GlobalLRUNewest" () As Integer
Declare Function extfn25A0 Lib "Kernel" Alias "GlobalLRUOldest" () As Integer
Declare Function extfn25D8 Lib "Kernel" Alias "GlobalPageLock" () As Integer
Declare Function extfn2610 Lib "Kernel" Alias "GlobalPageUnlock" () As Integer
Declare Function extfn23E0 Lib "Kernel" Alias "GlobalReAlloc" () As Integer
Declare Function extfn2418 Lib "Kernel" Alias "GlobalSize" () As Long
Declare Function extfn2680 Lib "Kernel" Alias "GlobalUnfix" () As Integer
Declare Function extfn2450 Lib "Kernel" Alias "GlobalUnlock" () As Integer
Declare Function extfn2530 Lib "Kernel" Alias "GlobalUnWire" () As Integer
Declare Function extfn24F8 Lib "Kernel" Alias "GlobalWire" () As Long
Declare Sub extsub7CB0 Lib "User" Alias "HideCaret" ()
Declare Function extfn6F90 Lib "User" Alias "HiliteMenuItem" () As Integer
Declare Sub extsub8180 Lib "User" Alias "InflateRect" ()
Declare Function extfn2DB8 Lib "Kernel" Alias "InitAtomTable" () As Integer
Declare Function extfn5E48 Lib "User" Alias "InSendMessage" () As Integer
Declare Function extfn7268 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn4D38 Lib "GDI" Alias "IntersectClipRect" () As Integer
Declare Function extfn81B8 Lib "User" Alias "IntersectRect" () As Integer
Declare Sub extsub75B0 Lib "User" Alias "InvalidateRect" ()
Declare Sub extsub7620 Lib "User" Alias "InvalidateRgn" ()
Declare Sub extsub80A0 Lib "User" Alias "InvertRect" ()
Declare Function extfn7F50 Lib "GDI" Alias "InvertRgn" () As Integer
Declare Function extfn3640 Lib "User" Alias "IsCharAlpha" () As Integer
Declare Function extfn3678 Lib "User" Alias "IsCharAlphaNumeric" () As Integer
Declare Function extfn36E8 Lib "User" Alias "IsCharLower" () As Integer
Declare Function extfn36B0 Lib "User" Alias "IsCharUpper" () As Integer
Declare Function extfn6008 Lib "User" Alias "IsChild" () As Integer
Declare Function extfn6A18 Lib "User" Alias "IsClipboardFormatAvailable" () As Integer
Declare Function extfn8A40 Lib "User" Alias "IsDialogMessage" () As Integer
Declare Function extfn6580 Lib "User" Alias "IsDlgButtonChecked" () As Integer
Declare Function extfn62E0 Lib "User" Alias "IsIconic" () As Integer
Declare Function extfn8260 Lib "User" Alias "IsRectEmpty" () As Integer
Declare Function extfn5FD0 Lib "User" Alias "IsWindow" () As Integer
Declare Function extfn6D98 Lib "User" Alias "IsWindowEnabled" () As Integer
Declare Function extfn62A8 Lib "User" Alias "IsWindowVisible" () As Integer
Declare Function extfn6388 Lib "User" Alias "IsZoomed" () As Integer
Declare Function extfn6D28 Lib "User" Alias "KillTimer" () As Integer
Declare Function extfn5A58 Lib "Kernel" Alias "_lclose" () As Integer
Declare Function extfn5A90 Lib "Kernel" Alias "_lcreat" () As Integer
Declare Sub extsub2A00 Lib "Kernel" Alias "LimitEmsPages" ()
Declare Function extfn3A30 Lib "GDI" Alias "LineTo" () As Integer
Declare Function extfn5AC8 Lib "Kernel" Alias "_llseek" () As Long
Declare Function extfn6DD0 Lib "User" Alias "LoadAccelerators" () As Integer
Declare Function extfn8810 Lib "User" Alias "LoadBitmap" () As Integer
Declare Function extfn8848 Lib "User" Alias "LoadCursor" () As Integer
Declare Function extfn88F0 Lib "User" Alias "LoadIcon" () As Integer
Declare Function extfn1F10 Lib "Kernel" Alias "LoadLibrary" () As Integer
Declare Function extfn6E78 Lib "User" Alias "LoadMenu" () As Integer
Declare Function extfn6EB0 Lib "User" Alias "LoadMenuIndirect" () As Integer
Declare Function extfn1F48 Lib "Kernel" Alias "LoadModule" () As Integer
Declare Function extfn2BF8 Lib "Kernel" Alias "LoadResource" () As Integer
Declare Function extfn8998 Lib "User" Alias "LoadString" () As Integer
Declare Function extfn2728 Lib "Kernel" Alias "LocalAlloc" () As Integer
Declare Function extfn2760 Lib "Kernel" Alias "LocalCompact" () As Integer
Declare Function extfn2920 Lib "Kernel" Alias "LocalFlags" () As Integer
Declare Function extfn2798 Lib "Kernel" Alias "LocalFree" () As Integer
Declare Function extfn27D0 Lib "Kernel" Alias "LocalHandle" () As Integer
Declare Function extfn2808 Lib "Kernel" Alias "LocalInit" () As Integer
Declare Function extfn2840 Lib "Kernel" Alias "LocalLock" () As Integer
Declare Function extfn2878 Lib "Kernel" Alias "LocalReAlloc" () As Integer
Declare Function extfn2958 Lib "Kernel" Alias "LocalShrink" () As Integer
Declare Function extfn28B0 Lib "Kernel" Alias "LocalSize" () As Integer
Declare Function extfn28E8 Lib "Kernel" Alias "LocalUnlock" () As Integer
Declare Function extfn2C68 Lib "Kernel" Alias "LockResource" () As Long
Declare Function extfn26B8 Lib "Kernel" Alias "LockSegment" () As Integer
Declare Function extfn5A20 Lib "Kernel" Alias "_lopen" () As Integer
Declare Function extfn52B0 Lib "GDI" Alias "LPtoDP" () As Integer
Declare Function extfn5B00 Lib "Kernel" Alias "_lread" () As Integer
Declare Function extfn59B0 Lib "Kernel" Alias "lstrcat" () As Long
Declare Function extfn5908 Lib "User" Alias "lstrcmp" () As Integer
Declare Function extfn5940 Lib "User" Alias "lstrcmpi" () As Integer
Declare Function extfn5978 Lib "Kernel" Alias "lstrcpy" () As Long
Declare Function extfn59E8 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn5B38 Lib "Kernel" Alias "_lwrite" () As Integer
Declare Sub extsub8A78 Lib "User" Alias "MapDialogRect" ()
Declare Function extfn3560 Lib "Keyboard" Alias "MapVirtualKey" () As Integer
Declare Sub extsub7A80 Lib "User" Alias "MessageBeep" ()
Declare Function extfn7A48 Lib "User" Alias "MessageBox" () As Integer
Declare Function extfn72D8 Lib "User" Alias "ModifyMenu" () As Integer
Declare Function extfn39C0 Lib "GDI" Alias "MoveTo" () As Long
Declare Sub extsub6190 Lib "User" Alias "MoveWindow" ()
Declare Function extfn3AA0 Lib "GDI" Alias "MulDiv" () As Integer
Declare Function extfn34B8 Lib "Keyboard" Alias "OemKeyScan" () As Long
Declare Function extfn32C0 Lib "Keyboard" Alias "OemToAnsi" () As Integer
Declare Sub extsub3330 Lib "Keyboard" Alias "OemToAnsiBuff" ()
Declare Function extfn4D70 Lib "GDI" Alias "OffsetClipRgn" () As Integer
Declare Sub extsub8228 Lib "User" Alias "OffsetRect" ()
Declare Function extfn4E88 Lib "GDI" Alias "OffsetRgn" () As Integer
Declare Function extfn4C20 Lib "GDI" Alias "OffsetViewportOrg" () As Long
Declare Function extfn4C90 Lib "GDI" Alias "OffsetWindowOrg" () As Long
Declare Function extfn6740 Lib "User" Alias "OpenClipboard" () As Integer
Declare Function extfn8F10 Lib "User" Alias "OpenComm" () As Integer
Declare Function extfn21B0 Lib "Kernel" Alias "OpenFile" () As Integer
Declare Function extfn6120 Lib "User" Alias "OpenIcon" () As Integer
Declare Function extfn8B90 Lib "Sound" Alias "OpenSound" () As Integer
Declare Sub extsub2098 Lib "Kernel" Alias "OutputDebugString" ()
Declare Function extfn7F88 Lib "GDI" Alias "PaintRgn" () As Integer
Declare Function extfn3D08 Lib "GDI" Alias "PatBlt" () As Integer
Declare Function extfn58D0 Lib "User" Alias "PeekMessage" () As Integer
Declare Function extfn3CD0 Lib "GDI" Alias "Pie" () As Integer
Declare Function extfn50B8 Lib "GDI" Alias "PlayMetaFile" () As Integer
Declare Sub extsub50F0 Lib "GDI" Alias "PlayMetaFileRecord" ()
Declare Function extfn3B48 Lib "GDI" Alias "Polygon" () As Integer
Declare Function extfn3B10 Lib "GDI" Alias "Polyline" () As Integer
Declare Function extfn3B80 Lib "GDI" Alias "PolyPolygon" () As Integer
Declare Function extfn5D30 Lib "User" Alias "PostAppMessage" () As Integer
Declare Function extfn5CF8 Lib "User" Alias "PostMessage" () As Integer
Declare Sub extsub5E10 Lib "User" Alias "PostQuitMessage" ()
Declare Sub extsub9488 Lib "User" Alias "ProfClear" ()
Declare Sub extsub94F8 Lib "User" Alias "ProfFinish" ()
Declare Sub extsub94C0 Lib "User" Alias "ProfFlush" ()
Declare Function extfn9370 Lib "User" Alias "ProfInsChk" () As Integer
Declare Sub extsub93E0 Lib "User" Alias "ProfSampRate" ()
Declare Sub extsub93A8 Lib "User" Alias "ProfSetup" ()
Declare Sub extsub9418 Lib "User" Alias "ProfStart" ()
Declare Sub extsub9450 Lib "User" Alias "ProfStop" ()
Declare Function extfn82D0 Lib "User" Alias "PtInRect" () As Integer
Declare Function extfn7FC0 Lib "GDI" Alias "PtInRegion" () As Integer
Declare Function extfn4DE0 Lib "GDI" Alias "PtVisible" () As Integer
Declare Function extfn8FB8 Lib "User" Alias "ReadComm" () As Integer
Declare Function extfn5588 Lib "User" Alias "RealizePalette" () As Integer
Declare Function extfn3BB8 Lib "GDI" Alias "Rectangle" () As Integer
Declare Function extfn8340 Lib "User" Alias "RectInRegion" () As Integer
Declare Function extfn8308 Lib "User" Alias "RectVisible" () As Integer
Declare Function extfn6900 Lib "User" Alias "RegisterClipboardFormat" () As Integer
Declare Function extfn57F0 Lib "User" Alias "RegisterWindowMessage" () As Integer
Declare Sub extsub6CF0 Lib "User" Alias "ReleaseCapture" ()
Declare Function extfn3838 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn8A08 Lib "GDI" Alias "RemoveFontResource" () As Integer
Declare Function extfn7310 Lib "User" Alias "RemoveMenu" () As Integer
Declare Function extfn7888 Lib "User" Alias "RemoveProp" () As Integer
Declare Sub extsub5D68 Lib "User" Alias "ReplyMessage" ()
Declare Function extfn56D8 Lib "GDI" Alias "ResizePalette" () As Integer
Declare Function extfn3988 Lib "GDI" Alias "RestoreDC" () As Integer
Declare Function extfn3BF0 Lib "GDI" Alias "RoundRect" () As Integer
Declare Function extfn3950 Lib "GDI" Alias "SaveDC" () As Integer
Declare Function extfn4C58 Lib "GDI" Alias "ScaleViewportExt" () As Long
Declare Function extfn4CC8 Lib "GDI" Alias "ScaleWindowExt" () As Long
Declare Sub extsub7DC8 Lib "User" Alias "ScreenToClient" ()
Declare Function extfn76C8 Lib "User" Alias "ScrollDC" () As Integer
Declare Sub extsub7690 Lib "User" Alias "ScrollWindow" ()
Declare Function extfn4360 Lib "GDI" Alias "SelectClipRgn" () As Integer
Declare Function extfn45C8 Lib "GDI" Alias "SelectObject" () As Integer
Declare Function extfn5550 Lib "User" Alias "SelectPalette" () As Integer
Declare Function extfn65B8 Lib "User" Alias "SendDlgItemMessage" () As Long
Declare Function extfn5CC0 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn7460 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn4210 Lib "GDI" Alias "SetBitmapBits" () As Long
Declare Function extfn4280 Lib "GDI" Alias "SetBitmapDimension" () As Long
Declare Function extfn4638 Lib "GDI" Alias "SetBkColor" () As Long
Declare Function extfn46A8 Lib "GDI" Alias "SetBkMode" () As Integer
Declare Function extfn4050 Lib "GDI" Alias "SetBrushOrg" () As Long
Declare Function extfn6CB8 Lib "User" Alias "SetCapture" () As Integer
Declare Sub extsub7C40 Lib "User" Alias "SetCaretBlinkTime" ()
Declare Sub extsub7D20 Lib "User" Alias "SetCaretPos" ()
Declare Function extfn8570 Lib "User" Alias "SetClassLong" () As Long
Declare Function extfn8500 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfn6890 Lib "User" Alias "SetClipboardData" () As Integer
Declare Function extfn67E8 Lib "User" Alias "SetClipboardViewer" () As Integer
Declare Function extfn9178 Lib "User" Alias "SetCommBreak" () As Integer
Declare Function extfn9108 Lib "User" Alias "SetCommEventMask" () As Long
Declare Function extfn8F48 Lib "User" Alias "SetCommState" () As Integer
Declare Function extfn7B28 Lib "User" Alias "SetCursor" () As Integer
Declare Sub extsub7AF0 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn53C8 Lib "GDI" Alias "SetDIBits" () As Integer
Declare Function extfn5438 Lib "GDI" Alias "SetDIBitsToDevice" () As Integer
Declare Sub extsub6430 Lib "User" Alias "SetDlgItemInt" ()
Declare Sub extsub64A0 Lib "User" Alias "SetDlgItemText" ()
Declare Sub extsub5EB8 Lib "User" Alias "SetDoubleClickTime" ()
Declare Function extfn5208 Lib "GDI" Alias "SetEnvironment" () As Integer
Declare Function extfn2A38 Lib "Kernel" Alias "SetErrorMode" () As Integer
Declare Function extfn6A88 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn2258 Lib "Kernel" Alias "SetHandleCount" () As Integer
Declare Sub extsub6BD8 Lib "User" Alias "SetKeyboardState" ()
Declare Function extfn49F0 Lib "GDI" Alias "SetMapMode" () As Integer
Declare Function extfn47F8 Lib "GDI" Alias "SetMapperFlags" () As Long
Declare Function extfn6F20 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn7380 Lib "User" Alias "SetMenuItemBitmaps" () As Integer
Declare Function extfn5F28 Lib "User" Alias "SetMessageQueue" () As Integer
Declare Function extfn5390 Lib "GDI" Alias "SetMetaFileBits" () As Integer
Declare Function extfn5630 Lib "GDI" Alias "SetPaletteEntries" () As Integer
Declare Function extfn8650 Lib "User" Alias "setparent" () As Integer
Declare Function extfn3E58 Lib "GDI" Alias "SetPixel" () As Long
Declare Function extfn4980 Lib "GDI" Alias "SetPolyFillMode" () As Integer
Declare Function extfn7818 Lib "User" Alias "SetProp" () As Integer
Declare Sub extsub80D8 Lib "User" Alias "SetRect" ()
Declare Sub extsub8110 Lib "User" Alias "SetRectEmpty" ()
Declare Sub extsub43D0 Lib "GDI" Alias "SetRectRgn" ()
Declare Function extfn48A0 Lib "GDI" Alias "SetROP2" () As Integer
Declare Function extfn7700 Lib "User" Alias "SetScrollPos" () As Integer
Declare Sub extsub7770 Lib "User" Alias "SetScrollRange" ()
Declare Function extfn8CE0 Lib "Sound" Alias "SetSoundNoise" () As Integer
Declare Function extfn4910 Lib "GDI" Alias "SetStretchBltMode" () As Integer
Declare Function extfn2990 Lib "Kernel" Alias "SetSwapAreaSize" () As Long
Declare Sub extsub7EA8 Lib "User" Alias "SetSysColors" ()
Declare Function extfn5C88 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn5780 Lib "GDI" Alias "SetSystemPaletteUse" () As Integer
Declare Function extfn4788 Lib "GDI" Alias "SetTextAlign" () As Integer
Declare Function extfn4FA0 Lib "GDI" Alias "SetTextCharacterExtra" () As Integer
Declare Function extfn4718 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn4EF8 Lib "GDI" Alias "SetTextJustification" () As Integer
Declare Function extfn4BB0 Lib "GDI" Alias "SetViewportExt" () As Long
Declare Function extfn4B40 Lib "GDI" Alias "SetViewportOrg" () As Long
Declare Function extfn8C70 Lib "Sound" Alias "SetVoiceAccent" () As Integer
Declare Function extfn8CA8 Lib "Sound" Alias "SetVoiceEnvelope" () As Integer
Declare Function extfn8C38 Lib "Sound" Alias "SetVoiceNote" () As Integer
Declare Function extfn8C00 Lib "Sound" Alias "SetVoiceQueueSize" () As Integer
Declare Function extfn8D18 Lib "Sound" Alias "SetVoiceSound" () As Integer
Declare Function extfn8ED8 Lib "Sound" Alias "SetVoiceThreshold" () As Integer
Declare Function extfn4AD0 Lib "GDI" Alias "SetWindowExt" () As Long
Declare Function extfn8490 Lib "User" Alias "SetWindowLong" () As Long
Declare Function extfn4A60 Lib "GDI" Alias "SetWindowOrg" () As Long
Declare Sub extsub61C8 Lib "User" Alias "SetWindowPos" ()
Declare Sub extsub78C0 Lib "User" Alias "SetWindowText" ()
Declare Function extfn8420 Lib "User" Alias "SetWindowWord" () As Integer
Declare Sub extsub7CE8 Lib "User" Alias "ShowCaret" ()
Declare Function extfn7AB8 Lib "User" Alias "ShowCursor" () As Integer
Declare Sub extsub60E8 Lib "User" Alias "ShowOwnedPopups" ()
Declare Sub extsub77E0 Lib "User" Alias "ShowScrollBar" ()
Declare Function extfn6078 Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2CD8 Lib "Kernel" Alias "SizeofResource" () As Integer
Declare Function extfn8D50 Lib "Sound" Alias "StartSound" () As Integer
Declare Function extfn8D88 Lib "Sound" Alias "StopSound" () As Integer
Declare Function extfn3D78 Lib "GDI" Alias "StretchBlt" () As Integer
Declare Function extfn54E0 Lib "GDI" Alias "StretchDIBits" () As Double
Declare Function extfn5BA8 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Sub extsub3250 Lib "Kernel" Alias "SwapRecording" ()
Declare Sub extsub20D0 Lib "Kernel" Alias "SwitchStackBack" ()
Declare Sub extsub2108 Lib "Kernel" Alias "SwitchStackTo" ()
Declare Function extfn8DF8 Lib "Sound" Alias "SyncAllVoices" () As Integer
Declare Function extfn3DE8 Lib "User" Alias "TabbedTextOut" () As Long
Declare Function extfn3DB0 Lib "GDI" Alias "TextOut" () As Integer
Declare Sub extsub31E0 Lib "Kernel" Alias "Throw" ()
Declare Function extfn3608 Lib "Keyboard" Alias "ToAscii" () As Integer
Declare Function extfn73F0 Lib "User" Alias "TrackPopupMenu" () As Integer
Declare Function extfn6E08 Lib "User" Alias "TranslateAccelerator" () As Integer
Declare Function extfn92C8 Lib "User" Alias "TranslateMDISysAccel" () As Integer
Declare Function extfn5860 Lib "User" Alias "TranslateMessage" () As Integer
Declare Function extfn90D0 Lib "User" Alias "TransmitCommChar" () As Integer
Declare Function extfn8FF0 Lib "User" Alias "UngetCommChar" () As Integer
Declare Function extfn81F0 Lib "User" Alias "UnionRect" () As Integer
Declare Function extfn2488 Lib "Kernel" Alias "GlobalUnlock" () As Integer
Declare Function extfn26F0 Lib "Kernel" Alias "UnlockSegment" () As Integer
Declare Function extfn4600 Lib "GDI" Alias "UnrealizeObject" () As Integer
Declare Function extfn5EF0 Lib "User" Alias "UnregisterClass" () As Integer
Declare Function extfn55C0 Lib "GDI" Alias "UpdateColors" () As Integer
Declare Sub extsub7428 Lib "User" Alias "UpdateWindow" ()
Declare Sub extsub2A70 Lib "Kernel" Alias "ValidateCodeSegments" ()
Declare Function extfn29C8 Lib "Kernel" Alias "ValidateFreeSpaces" () As Long
Declare Sub extsub75E8 Lib "User" Alias "ValidateRect" ()
Declare Sub extsub7658 Lib "User" Alias "ValidateRgn" ()
Declare Function extfn34F0 Lib "Keyboard" Alias "VkKeyScan" () As Integer
Declare Sub extsub5DA0 Lib "User" Alias "WaitMessage" ()
Declare Function extfn8DC0 Lib "Sound" Alias "WaitSoundState" () As Integer
Declare Function extfn7E00 Lib "User" Alias "WindowFromPoint" () As Integer
Declare Function extfn2028 Lib "Kernel" Alias "WinExec" () As Integer
Declare Function extfn9338 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn9028 Lib "User" Alias "WriteComm" () As Integer
Declare Function extfn3100 Lib "Kernel" Alias "WritePrivateProfileString" () As Integer
Declare Function extfn3058 Lib "Kernel" Alias "WriteProfileString" () As Integer
Declare Function extfn57B8 Lib "User" Alias "wvsprintf" () As Integer
Declare Sub extsub2D48 Lib "Kernel" Alias "Yield" ()
