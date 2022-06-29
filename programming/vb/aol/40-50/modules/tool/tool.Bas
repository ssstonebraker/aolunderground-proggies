Attribute VB_Name = "tool"
'tool.bas (32 bit module)
'
'  hey, well i worked hard on this bas and all these subs
'and functions i wrote myself but i may have got help!!!


'site: http://www.8op.com/tool
'email: tool9@marijuana.com
'aim : ett0


' the greets are too long to say but well
' thanxs for the help!








'don't mess with anything below here it is
'the constants and declarations.....
Option Explicit

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetHostName Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function GetHostByName Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function getmodulefilename Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, LPRect As RECT) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Const GW_CHILD = 5
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_GETTEXT = &HD
Public Const WM_GetTextLength = &HE
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CLEAR = &H303
Public Const WM_DRAWITEM = &H2B
Public Const WM_PAINT = &HF
Public Const WM_ERASEBKGND = &H14
Public Const WM_NCPAINT = &H85
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const COLOR_RED = &HFF&
Public Const COLOR_GREEN = &HFF00&
Public Const COLOR_BLUE = &HFF0000
Public Const COLOR_YELLOW = &HFFFF&
Public Const COLOR_WHITE = &HFFFFFE
Public Const COLOR_BLACK = &H0&
Public Const COLOR_PEACH = &HC0C0FF
Public Const COLOR_PURPLE = &HFF00FF
Public Const COLOR_GREY = &HC0C0C0
Public Const COLOR_PINK = &HFF80FF
Public Const COLOR_TURQUOISE = &HC0C000
Public Const COLOR_LIGHTBLUE = &HFF8080
Public Const COLOR_ORANGE = &H80FF&
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_POPUP = &H10&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const ENTER_KEY = 13
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
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
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4
Public Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12
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

Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type CPINFO
        MaxCharSize As Long
        DefaultChar(MAX_DEFAULTCHAR - 1) As Byte
        LeadByte(MAX_LEADBYTES - 1) As Byte
End Type

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type COLORRGB
    red As Long
    GREEN As Long
    BLUE As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Private mvarFilePath As String





Public Function getlinefromstring(strString As String, lngLine As Long) As String
    Dim strLine As String, lngCount As Long, lngspot1 As Long, lngspot2 As Long, Index As Long
    Let lngCount& = getstringlinecount(strString$)
    If lngLine& > lngCount& Then Exit Function
    If lngLine& = 1& And lngCount& = 1& Then Let getlinefromstring$ = strString$:  Exit Function
    If lngLine& = 1& And lngCount& <> 1& Then
        Let strLine$ = Left$(strString$, InStr(strString$, Chr$(13&)) - 1&)
        Let strLine$ = replacestring(strLine$, Chr$(13&), "")
        Let strLine$ = replacestring(strLine$, Chr$(10&), "")
        Let getlinefromstring$ = strLine$
        Exit Function
    Else
        Let lngspot1& = InStr(strString$, Chr$(13&))
        For Index& = 1& To lngLine& - 1&
            Let lngspot2& = lngspot1&
            Let lngspot1& = InStr(lngspot1& + 1&, strString$, Chr$(13&))
        Next Index
        If lngspot1& = 0& Then Let lngspot1& = Len(strString$)
        If (lngspot1& - lngspot2&) + 1& <= Len(strString$) Then
            If lngspot2& = 0& Then Let lngspot2& = lngspot2& + 1&
            Let strLine$ = Mid$(strString$, lngspot2&, (lngspot1& - lngspot2&) + 1&)
        End If
        Let strLine$ = replacestring(strLine$, Chr$(13&), "")
        Let strLine$ = replacestring(strLine$, Chr$(10&), "")
        Let getlinefromstring$ = strLine$
    End If
End Function



Public Sub chatignorebyindex(lngindex As Long)
    Dim lngroom As Long, lnglist As Long, lnginfo As Long, strwho As String
    Dim lngcheck As Long, checkval As Long, lngCount As Long

    If lngindex& > lngCount& - 1& Then
        Exit Sub
    Else
        Let lngroom& = findroom&
        Let lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
        Let strwho$ = getlistitemtext(lnglist&, lngindex&)
        Call SendMessageLong(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
        Do: DoEvents
            Let lnginfo& = FindWindowEx(GetParent(lngroom&), 0&, "aol child", strwho$)
            Let lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
        Loop Until lnginfo& <> 0& And lngcheck& <> 0&
        Do: DoEvents
            Let checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until checkval& <> 0&
        Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub chatunignorebyindex(lngindex As Long)
    Dim lngroom As Long, lnglist As Long, lnginfo As Long, strwho As String
    Dim lngcheck As Long, checkval As Long, lngCount As Long

    If lngindex& > lngCount& - 1& Then
        Exit Sub
    Else
        Let lngroom& = findroom&
        Let lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
        Let strwho$ = getlistitemtext(lnglist&, lngindex&)
        Call SendMessageLong(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
        Do: DoEvents
            Let lnginfo& = FindWindowEx(GetParent(lngroom&), 0&, "aol child", strwho$)
            Let lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
        Loop Until lnginfo& <> 0& And lngcheck& <> 0&
        Do: DoEvents
            Let checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until checkval& <> 1&
        Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub chatignorebyname(strname As String)
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim strwho As String, lngcheck As Long, lnginfo As Long, checkval As Long
    Dim lIndex As Long
    On Error Resume Next
    Let Room& = findroom&
    If Room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For Index& = 0& To SendMessage(rlist, LB_GETCOUNT, 0&, 0&) - 1&
                Let Screenname$ = String$(4&, vbNullChar)
                Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
                Let itmhold& = itmhold& + 24&
                Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes)
                Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
                Let psnHold& = psnHold& + 6&
                Let Screenname$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
                Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
                If Screenname$ <> getuser$ And LCase$(Screenname$) = LCase$(strname$) Then
                    Call SendMessage(rlist&, LB_SETCURSEL, Index&, 0&)
                    Let strwho$ = getlistitemtext(rlist&, Index&)
                    Call SendMessageLong(rlist&, WM_LBUTTONDBLCLK, 0&, 0&)
                    Do: DoEvents
                        Let lnginfo& = FindWindowEx(GetParent(Room&), 0&, "aol child", strwho$)
                        Let lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
                    Loop Until lnginfo& <> 0& And lngcheck& <> 0&
                    Do: DoEvents
                        Let checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
                        DoEvents
                    Loop Until checkval& <> 0&
                    Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
                    DoEvents
                    Exit Sub
                End If
            Next Index&
            Call CloseHandle(mthread)
        End If
    End If
End Sub

Public Sub chatunignorebyname(strname As String)
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim strwho As String, lngcheck As Long, lnginfo As Long, checkval As Long
    On Error Resume Next
    Let Room& = findroom&
    If Room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For Index& = 0& To SendMessage(rlist, LB_GETCOUNT, 0&, 0&) - 1&
                Let Screenname$ = String$(4&, vbNullChar)
                Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
                Let itmhold& = itmhold& + 24&
                Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes)
                Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
                Let psnHold& = psnHold& + 6&
                Let Screenname$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
                Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
                If Screenname$ <> getuser$ And LCase$(Screenname$) = LCase$(strname$) Then
                    Call SendMessage(rlist&, LB_SETCURSEL, Index&, 0&)
                    Let strwho$ = getlistitemtext(rlist&, Index&)
                    Call SendMessageLong(rlist&, WM_LBUTTONDBLCLK, 0&, 0&)
                    Do: DoEvents
                        Let lnginfo& = FindWindowEx(GetParent(Room&), 0&, "aol child", strwho$)
                        Let lngcheck& = FindWindowEx(lnginfo&, 0&, "_AOL_Checkbox", vbNullString)
                    Loop Until lnginfo& <> 0& And lngcheck& <> 0&
                    Do: DoEvents
                        Let checkval& = SendMessage(lngcheck&, BM_GETCHECK, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONDOWN, 0&, 0&)
                        Call SendMessageLong(lngcheck&, WM_LBUTTONUP, 0&, 0&)
                        DoEvents
                    Loop Until checkval& <> 1&
                    Call PostMessage(lnginfo&, WM_CLOSE, 0&, 0&)
                    DoEvents
                    Exit Sub
                End If
            Next Index&
            Call CloseHandle(mthread)
        End If
    End If
End Sub





Public Function getstringlinecount(strString As String) As Long
    Dim enterpos As Long
    If Len(strString$) = 0& Then
        Let getstringlinecount& = 0&
    Else
        Let enterpos& = InStr(strString$, Chr$(13&))
        If enterpos& <> 0& Then
            Let getstringlinecount& = 1&
            Do While enterpos& <> 0&
                Let enterpos& = InStr(enterpos& + 1&, strString$, Chr$(13&))
                If enterpos& <> 0& Then
                    Let getstringlinecount& = getstringlinecount& + 1&
                End If
            Loop: DoEvents
        End If
        Let getstringlinecount& = getstringlinecount& + 1&
    End If
End Function
Public Function deletelinefromstring(strString As String, lngLine As Long) As String
    Dim lnglinecount As Long, Index As Long, strText As String
    Let lnglinecount& = getstringlinecount(strString)
    For Index& = 1& To lnglinecount&
        If Index& <> lngLine& Then
            strText$ = strText$ & getlinefromstring(strString, Index&) & vbCrLf$
        End If
    Next Index&
    deletelinefromstring$ = strText$
End Function





Function GetClass(child)
Dim Buffer, getclas

Buffer = String$(250, 0)
getclas = GetClassName(child, Buffer, 250)

GetClass = Buffer
End Function


Function FindChildByClass(parentw, childhand)
Dim firs, firss, Room

firs = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs
firss = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs = GetWindow(firs, 2)
If UCase(Mid(GetClass(firs), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room = firs
FindChildByClass = Room

End Function



Sub anti45mintimer()
Dim aotimer%, aoicon%
aotimer% = FindWindow("_AOL_Palette", vbNullString)
aoicon% = FindChildByClass(aotimer%, "_AOL_Icon")
ClickIcon (aoicon%)
End Sub





Public Sub IMsOn()
Call imsend("$im_on", ".")

End Sub

Public Sub runaolmenu(lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngaol As Long, lngmenu As Long, lngsubmenu As Long, lngid As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmenu& = GetMenu(lngaol&)
    Let lngsubmenu& = GetSubMenu(lngmenu&, lngmenunumber&)
    Let lngid& = GetMenuItemID(lngsubmenu&, lngsubmenunumber&)
    Call SendMessageLong(lngaol&, WM_COMMAND, ByVal lngid&, 0&)
End Sub
Public Sub runanymenu(lngwindow&, lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngmenu As Long, lngsubmenu As Long, lngid As Long
    Let lngmenu& = GetMenu(lngwindow&)
    Let lngsubmenu& = GetSubMenu(lngmenu&, lngmenunumber&)
    Let lngid& = GetMenuItemID(lngsubmenu&, lngsubmenunumber&)
    Call SendMessageLong(lngwindow&, WM_COMMAND, ByVal lngid&, 0&)
End Sub
Public Sub runanymenubystring(lngwindow As Long, strmenutext As String)
    Dim lngmmenu As Long, lngmmcount As Long, lngindex As Long
    Dim lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strString As String
    Let lngmmenu& = GetMenu(lngwindow&)
    Let lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
        Let lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
        Let lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
            Let lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
            Let strString$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strString$, 100&, 1&)
            If LCase$(strString$) = replacestring(LCase$(strmenutext$), Chr$(0&), "") Then
                Call SendMessageLong(lngwindow&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub






Public Sub windowclose(lnghwnd As Long)
  Call PostMessage(lnghwnd, WM_CLOSE, 0&, 0&)
End Sub
Public Sub windowhide(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_HIDE)
End Sub
Public Sub windowshow(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_SHOW)
End Sub


Public Function replacestring(strString As String, strwhat As String, strwith As String) As String
    Dim lngPos As Long
    Do While InStr(1&, strString$, strwhat$)
        DoEvents
        Let lngPos& = InStr(1&, strString$, strwhat$)
        Let strString$ = Left$(strString$, (lngPos& - 1&)) & Right$(strString$, Len(strString$) - (lngPos& + Len(strwhat$) - 1&))
    Loop
    Let replacestring$ = strString$
End Function
Public Function alive(strname As String) As Boolean
    Dim lngaol As Long, lngmdi As Long, lngerrorwin As Long
    Dim lngerrorview As Long, strerrorstring As String
    Dim lngmailwin As Long, lngnowin As Long, lngbutton As Long
    Call SendMail("*, " & strname$, "you alive?", " ")
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorview& = FindWindowEx(lngerrorwin&, 0&, "_AOL_View", vbNullString)
        Let strerrorstring$ = GetText(lngerrorview&)
    Loop Until lngerrorwin& <> 0& And lngerrorview& <> 0& And strerrorstring$ <> ""
    If InStr(LCase$(replacestring(strerrorstring$, " ", "")), LCase$(replacestring(strname$, " ", ""))) > 0& Then
        Let alive = False
    Else
        Let alive = True
    End If
    Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(lngmailwin&, WM_CLOSE, 0&, 0&)
    Call PostMessage(lngerrorwin&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do: DoEvents
        Let lngnowin& = FindWindow("#32770", "America Online")
        Let lngbutton& = FindWindowEx(lngnowin&, 0&, "Button", "&No")
    Loop Until lngnowin& <> 0& And lngbutton& <> 0
    Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub SignOnGuest(strscreenname As String, strpassword As String)
    Dim lngaol As Long, lngmdi As Long, lngsignonwin As Long, lngcombo As Long
    Dim lngedit As Long, lngsignonbut1 As Long, lngsignonbut2 As Long, lngsignonbut3 As Long
    Dim lngsignonbut As Long, lngedit1 As Long, lngokbut As Long, lngmodal As Long
    Dim lngokbut1 As Long, lngokbut2 As Long, lngerrwin As Long, lngsignonwin2 As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngsignonwin& = findsignonwin&
    If lngsignonwin& = 0& Then
        Call runaolmenubystring("&Sign On Screen")
    End If
    Do: DoEvents
        Let lngsignonwin& = findsignonwin&
        Let lngsignonwin2& = findsignoffwin&
    Loop Until (lngsignonwin& <> 0& Or lngsignonwin2& <> 0&)
    If lngsignonwin& = 0& And lngsignonwin2& <> 0& Then
        Let lngsignonwin& = lngsignonwin2&
    End If
    Do: DoEvents
        Let lngcombo& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Combobox", vbNullString)
        Let lngedit& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngsignonbut1& = FindWindowEx(lngsignonwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsignonbut2& = FindWindowEx(lngsignonwin&, lngsignonbut1&, "_AOL_Icon", vbNullString)
        Let lngsignonbut3& = FindWindowEx(lngsignonwin&, lngsignonbut2&, "_AOL_Icon", vbNullString)
        Let lngsignonbut& = FindWindowEx(lngsignonwin&, lngsignonbut3&, "_AOL_Icon", vbNullString)
    Loop Until lngcombo& <> 0& And lngsignonbut1& <> 0& And lngsignonbut2& <> 0& And lngsignonbut3& <> 0& And lngsignonbut& <> 0& And _
    lngsignonbut1& <> lngsignonbut2& And lngsignonbut2& <> lngsignonbut3& And lngsignonbut3& <> lngsignonbut&
    Call PostMessage(lngcombo&, CB_SETCURSEL, SendMessage(lngcombo&, CB_GETCOUNT, 0&, 0&) - 1&, 0&)
    If lngedit& <> 0& Then
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, "yo b")
    End If
    Call PostMessage(lngsignonbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngsignonbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        Let lngedit1& = FindWindowEx(lngmodal&, 0&, "_AOL_Edit", vbNullString)
        Let lngedit& = FindWindowEx(lngmodal&, lngedit1&, "_AOL_Edit", vbNullString)
        Let lngokbut1& = FindWindowEx(lngmodal&, 0&, "_AOL_Icon", vbNullString)
        Let lngokbut& = FindWindowEx(lngmodal&, lngokbut1&, "_AOL_Icon", vbNullString)
    Loop Until lngmodal& <> 0& And lngedit1& <> 0& And lngedit& <> 0& And lngokbut1& <> 0& And lngokbut& <> 0&: pause 0.2
    pause 0.2
    Call SendMessageByString(lngedit1&, WM_SETTEXT, 0&, strscreenname$)
    pause 0.2
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strpassword$)
    pause 0.2
    Do: DoEvents
    Call clicksignon
        Let lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        Let lngerrwin& = FindWindow("#32770", "America Online")
    Loop Until lngmodal& = 0& Or lngerrwin& <> 0&: pause 0.2
    If lngerrwin& <> 0& Then
        Let lngokbut2& = FindWindowEx(lngerrwin&, 0&, "Button", vbNullString)
        Do: DoEvents
            Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
           Loop Until findsignonwin& <> 0& Or findwelcomewin&
    End If
End Sub

Public Sub waitforok()
    Dim lngokwin As Long, lngokbut As Long, lngokwin2 As Long, lngokbut2 As Long
    Do: DoEvents
        Let lngokwin& = FindWindow("#32770", "America Online")
        Let lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Let lngokwin2& = FindWindow("_AOL_Modal", "America Online")
        Let lngokbut2& = FindWindowEx(lngokwin2&, 0&, "_AOL_Icon", vbNullString)
    Loop Until (lngokwin& <> 0& And lngokbut& <> 0&) Or (lngokwin2& <> 0& And lngokbut2& <> 0&)
    If lngokwin& <> 0& Then
        Do: DoEvents
            pause 0.2
            Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
            Let lngokwin& = FindWindow("#32770", "America Online")
            Let lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Loop Until lngokwin& = 0&
    ElseIf lngokwin2& <> 0& Then
        Do: DoEvents
            pause 0.2
            Call SendMessageLong(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
            Let lngokwin2& = FindWindow("_AOL_Modal", "America Online")
            Let lngokbut2& = FindWindowEx(lngokwin2&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngokwin2& = 0&
    End If
End Sub

Public Sub sendpage(pagerid As String, message As String)
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long, lngsendpage As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngsendwin As Long
    Dim lngsendbut As Long, lngidbox As Long, lngmsgbox As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyE, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyE, 0&)
    Do: DoEvents
        Let lngsendpage& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Send Page")
        Let lngicon1& = FindWindowEx(lngsendpage&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon2& = FindWindowEx(lngsendpage&, lngicon1&, "_AOL_Icon", vbNullString)
        Let lngicon3& = FindWindowEx(lngsendpage&, lngicon2&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngsendpage&, lngicon3&, "_AOL_Icon", vbNullString)
    Loop Until lngsendpage& <> 0& And lngicon1 <> 0& And lngicon2 <> 0& And lngicon3 <> 0& And lngicon <> 0& And lngicon1 <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", replacestring(UCase$(getuser$), " ", "") & "'s Send a Page")
        Let lngidbox& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngmsgbox& = FindWindowEx(lngsendwin&, lngidbox&, "_AOL_Edit", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngsendwin& <> 0& And lngidbox& <> 0& And lngmsgbox& <> 0& And lngicon& <> 0&: pause 0.2
    Call SendMessage(lngsendpage&, WM_CLOSE, 0&, 0&)
    Call SendMessageByString(lngidbox&, WM_SETTEXT, 0&, pagerid$)
    Call SendMessageByString(lngmsgbox&, WM_SETTEXT, 0&, message$)
    Call PostMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
    pause 4
    Call SendMessage(lngsendwin&, WM_CLOSE, 0&, 0&)
End Sub



Public Function TrimSpaces(strText As String) As String
    Let TrimSpaces$ = replacestring(strText$, " ", "")
End Function
Public Function trimenters(strText As String) As String
    Let trimenters$ = replacestring(strText$, Chr$(13&), "")
    Let trimenters$ = replacestring(trimenters$, Chr$(10&), "")
End Function
Public Function trimchar(strText As String, strChar As String) As String
    trimchar$ = replacestring(strText$, strChar$, "")
End Function


Public Sub waitforlisttoload(lnglist As Long)
    Dim lngCount As Long
    Do: DoEvents
        Let lngCount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&): Call pause(2&)
        If lngCount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Do
    Loop
End Sub
Public Sub waitfortextboxtoload(lngtextbox As Long)
    Dim strText As String
    Do: DoEvents
        Let strText$ = GetText(lngtextbox&): Call pause(2&)
        If strText$ = GetText(lngtextbox&) Then Exit Do
    Loop
End Sub


Public Sub AntiIdle()
    Dim lngpalette As Long, lngmodal As Long
    Dim lngbutton As Long, lngbutton2 As Long
    Let lngpalette& = FindWindow("_aol_palette", vbNullString)
    Let lngmodal& = FindWindow("_aol_modal", vbNullString)
    Let lngbutton& = FindWindowEx(lngpalette&, 0&, "_aol_icon", vbNullString)
    Let lngbutton2& = FindWindowEx(lngmodal&, 0&, "_aol_icon", vbNullString)
    Call SendMessageLong(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngbutton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessageLong(lngbutton2&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngbutton2&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function getuser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        getuser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                getuser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    getuser$ = ""
End Function


Public Function findwelcomewin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
        If InStr(GetCaption(lngchild&), "Welcome, ") Then findwelcomewin& = lngchild&: Exit Function
    Loop While lngchild& <> 0&
End Function
Public Sub sendchatlist(TheControl As Control)
    Dim lngindex As Long
    For lngindex& = 0& To TheControl.ListCount
        Call SendChat(TheControl.List(lngindex&))
        pause 0.7
    Next lngindex&
End Sub



Sub SetText(SetThis As Integer, ByVal what As String)
Dim Q
Q = SendMessageByString(SetThis%, WM_SETTEXT, 0, what)
End Sub

Public Function stringreverse(strString As String) As String
    'note: for vb6 users this code can be replaced with
    'strstring$ = strreverse(strstring$)
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strString$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strString$, lngindex&, 1&)
        Let stringreverse$ = strnextchar$ & stringreverse$
    Loop
End Function
Public Function stringdotted(strString As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strString$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strString$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & "o"
        Let stringdotted$ = stringdotted$ & strnextchar$
    Loop
End Function
Public Function stringcustomized(strString As String, streveryotherchar As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strString$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strString$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & streveryotherchar$
        Let stringcustomized$ = stringcustomized$ & strnextchar$
    Loop
End Function
Public Function stringspaced(strString As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strString$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strString$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & " "
        Let stringspaced$ = stringspaced$ & strnextchar$
    Loop
End Function
Public Function stringlinked(strurl As String, strText$) As String
    Let stringlinked$ = "< a href=" & Chr$(34&) & strurl$ & Chr$(34&) & ">" & strText$ & "</a>"
End Function

Public Sub resethistorycombo()
    Dim lngaol As Long, lngtoolbar As Long, lngcombo As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_aol_toolbar", vbNullString)
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
    Call SendMessage(lngcombo&, CB_RESETCONTENT, 0&, 0&)
End Sub
Public Function getlistitemindex(TheList As Long, item As String) As Long
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long, Index As Long
    Dim Screenname As String, itmhold As Long, psnHold As Long
    Dim rbytes As Long, cprocess As Long
    Let rlist& = TheList&
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let Screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let Screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
            If Screenname$ = item$ Then
                Let getlistitemindex& = Index&
                Call CloseHandle(mthread&)
                Exit Function
            End If
        Next Index&
        Let getlistitemindex& = -1&
        Call CloseHandle(mthread&)
    End If
End Function
Public Sub fadepicbox(picture As PictureBox, color1 As Long, color2 As Long)
Dim lngcon As Long, longcon As Long, lnghlfwidth As Long, lngcolorval1 As Long
    Dim lngcolorval2 As Long, lngcolorval3 As Long, lngrgb1 As Long, lngrgb2 As Long
    Dim lngrgb3 As Long, lngyval As Long, strcolor1 As String, strcolor2 As String
    Dim strred1 As String, strgreen1 As String, strblue1 As String, strred2 As String
    Dim strgreen2 As String, strblue2 As String, lngred1 As Long, lnggreen1 As Long
    Dim lngblue1 As Long, lngred2 As Long, lnggreen2 As Long, lngblue2 As Long
    Let picture.AutoRedraw = True
    Let picture.DrawStyle = 6&
    Let picture.DrawMode = 13&
    Let picture.DrawWidth = 2&
    Let lngcon& = 0&
    Let lnghlfwidth& = picture.Width / 2&
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Do: DoEvents
        On Error Resume Next
        Let lngcolorval1& = lngred2& - lngred1&
        Let lngcolorval2& = lnggreen2& - lnggreen1&
        Let lngcolorval3& = lngblue2& - lngblue1&
        Let lngrgb1& = (lngcolorval1& / lnghlfwidth& * lngcon&) + lngred1&
        Let lngrgb2& = (lngcolorval2& / lnghlfwidth& * lngcon&) + lnggreen1&
        Let lngrgb3& = (lngcolorval3& / lnghlfwidth& * lngcon&) + lngblue1&
        picture.Line (lngyval&, 0&)-(lngyval& + 2&, picture.Height), RGB(lngrgb1&, lngrgb2&, lngrgb3&), BF
        Let lngyval& = lngyval& + 10&
        Let lngcon& = lngcon& + 5&
    Loop Until lngcon& > lnghlfwidth&
End Sub

Public Function GetListCount(lnglist As Long) As Long
    Let GetListCount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function








Public Sub imsend(Person As String, message As String)
    If Person = "" Then Exit Sub
    
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    Dim AOLFrame As Long, MDIClient As Long, OpenIM As Long
   AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    pause 0.2
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
  OK& = FindWindow("#32770", "America Online")
  If OK& = True Then
  Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
  ClickIcon Button&
  Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
  Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
  Exit Sub
  End If
    Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
    pause 0.1
    OpenIM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    If OpenIM = False Then pause 0.3
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
  OK& = FindWindow("#32770", "America Online")
  If OK& = True Then
  Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
  ClickIcon Button&
  Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
  Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
  Exit Sub
  End If
Do
OK& = FindWindow("#32770", "America Online")

IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Loop Until OK& <> 0& Or IM& = 0&
If OK& <> 1 Then
Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
ClickIcon Button&
Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)

If OpenIM = True Then pause 0.2
Call PostMessage(OpenIM&, WM_CLOSE, 0&, 0&)
End If


End Sub
Function Is25() As Boolean
Dim hIcon As Long, hTool As Long, hMain As Long

hMain& = FindWindow("AOL Frame25", vbNullString)
hTool& = FindWindowEx(hMain&, 0&, "AOL Toolbar", vbNullString)
hIcon& = FindWindowEx(hTool&, 0&, "_AOL_Icon", vbNullString)
hIcon& = FindWindowEx(hTool&, hIcon&, "_AOL_Icon", vbNullString)

If Len(GetText(hIcon)) = 0 And hIcon <> 0 Then
    Is25 = True
Else
    Is25 = False
End If

End Function


Public Sub addfavoriteitem(strdescription As String, strurl As String)
    'add's item to certian folder!
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim Index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
    Dim lngcheck2 As Long, lngedit As Long, lngedit2 As Long, lngedit3 As Long
    Dim lngfavwin As Long, lngfavwin2 As Long, lngicon2 As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
        Let lngfavwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Favorite Places")
        Let lngicon& = FindWindowEx(lngfavwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngfavwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngfavwin& <> 0& And lngicon& <> 0&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
        lngcheck& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Checkbox", vbNullString)
        lngcheck2& = FindWindowEx(lngfavwin2&, lngcheck&, "_AOL_Checkbox", vbNullString)
        lngedit& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Edit", vbNullString)
        lngedit2& = FindWindowEx(lngfavwin2&, lngedit&, "_AOL_Edit", vbNullString)
        lngedit3& = FindWindowEx(lngfavwin2&, lngedit2&, "_AOL_Edit", vbNullString)
        lngicon& = FindWindowEx(lngfavwin2&, 0, "_AOL_Icon", vbNullString)
        lngicon2& = FindWindowEx(lngfavwin2&, lngicon&, "_AOL_Icon", vbNullString)
        pause 0.2
    Loop Until lngfavwin2& <> 0& And lngcheck& <> 0& And lngcheck2& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngicon& <> 0& And lngicon2& <> 0&
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, strdescription$)
    Call SendMessageByString(lngedit3&, WM_SETTEXT, 0&, strurl$)
    Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
    Loop Until lngfavwin2& = 0&
    Call windowclose(lngfavwin&)
End Sub
Function randomascii(Text As String) As String
Dim arrow1$, arrow2$, X, blah
Select Case Int(Rnd * 4)
Case 0:
    arrow1$ = "« "
    arrow2$ = " »"
Case 1:
    arrow1$ = "÷—"
    arrow2$ = "—÷"
Case 2:
    arrow1$ = "—"
    arrow2$ = "—"
Case 3:
    arrow1$ = "·•"
    arrow2$ = "•·"
End Select
For X = 0 To Int(Rnd * 5) + 1
    Select Case Int(Rnd * 7)
        Case 0:
            arrow1$ = arrow1$ + " "
            arrow2$ = " " + arrow2$
        Case 1:
            arrow1$ = arrow1$ + "÷"
            arrow2$ = "÷" + arrow2$
        Case 2:
            arrow1$ = arrow1$ + "•"
            arrow2$ = "•" + arrow2$
        Case 3:
            arrow1$ = arrow1$ + "×"
            arrow2$ = "×" + arrow2$
        Case 4:
            arrow1$ = arrow1$ + "¤"
            arrow2$ = "¤" + arrow2$
        Case 5:
            arrow1$ = arrow1$ + "["
            arrow2$ = "]" + arrow2$
        Case 6:
            arrow1$ = arrow1$ + "(•"
            arrow2$ = "•)" + arrow2$
    End Select
Next X
blah = arrow1$ + Text$ + arrow2$
End Function





Function labelcredits(lab As Label, name As String, Num As Integer)
'for greets on prog set bg to black
lab.Alignment = 2
lab.Caption = name
lab.BackStyle = 0
lab.AutoSize = True
Dim i, X, newc
For i = 64 To 1 Step -5
        X = 255 - (i * 4 - 1)
         
        newc = RGB(X, X, X)
        lab.ForeColor = newc
       pause (0.0000001)
    Next i
pause (Num)
For i = 1 To 64 Step 5
        X = 255 - (i * 4 - 1)
         
        newc = RGB(X, X, X)
        lab.ForeColor = newc
         pause (0.0000001)
    Next i
lab.ForeColor = &H0&
End Function


Public Sub addfavoritefolder(strfoldername As String)
    'adds favorite's folder!
    
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim Index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
    Dim lngcheck2 As Long, lngedit As Long, lngedit2 As Long, lngedit3 As Long
    Dim lngfavwin As Long, lngfavwin2 As Long, lngicon2 As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
        Let lngfavwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Favorite Places")
        Let lngicon& = FindWindowEx(lngfavwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngfavwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngfavwin& <> 0& And lngicon& <> 0&
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Let lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
        Let lngcheck& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Checkbox", vbNullString)
        Let lngcheck2& = FindWindowEx(lngfavwin2&, lngcheck&, "_AOL_Checkbox", vbNullString)
        Let lngedit& = FindWindowEx(lngfavwin2&, 0&, "_AOL_Edit", vbNullString)
        Let lngedit2& = FindWindowEx(lngfavwin2&, lngedit&, "_AOL_Edit", vbNullString)
        Let lngedit3& = FindWindowEx(lngfavwin2&, lngedit2&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lngfavwin2&, 0, "_AOL_Icon", vbNullString)
        Let lngicon2& = FindWindowEx(lngfavwin2&, lngicon&, "_AOL_Icon", vbNullString)
        pause 0.2
    Loop Until lngfavwin2& <> 0& And lngcheck& <> 0& And lngcheck2& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngicon& <> 0& And lngicon2& <> 0&
    Call PostMessage(lngcheck2&, BM_SETCHECK, True, 0&)
    Call PostMessage(lngcheck2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngcheck2&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngcheck2&, BM_SETCHECK, True, 0&)
    Do: DoEvents
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfoldername$)
        Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(lngicon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon2&, WM_LBUTTONUP, 0&, 0&)
        Let lngfavwin2& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
    Loop Until lngfavwin2& = 0&
    Call windowclose(lngfavwin&)
End Sub
Public Sub enableaoltoolbar()
  Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim Index As Long, lngcombo As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Call enablewindow(lngtoolbar&, True)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Call enablewindow(lngtoolbar&, True)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Call enablewindow(lngicon&, True)
    For Index& = 0& To 20&
        Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
        Call enablewindow(lngicon&, True)
    Next Index&
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
End Sub
Sub chatsend25and40(Text As String)
'this sends to cat 25 and 40
Dim hIcon As Long, hTool As Long, hMain As Long

hMain& = FindWindow("AOL Frame25", vbNullString)
hTool& = FindWindowEx(hMain&, 0&, "AOL Toolbar", vbNullString)
hIcon& = FindWindowEx(hTool&, 0&, "_AOL_Icon", vbNullString)
hIcon& = FindWindowEx(hTool&, hIcon&, "_AOL_Icon", vbNullString)

If Len(GetText(hIcon)) = 0 And hIcon <> 0 Then
ChatSend25 (Text$)
Else
End If
    

Dim hBar As Long
Dim hFrame As Long
hFrame& = FindWindow("AOL Frame25", vbNullString)
hBar& = FindWindowEx(hFrame, 0&, "AOL Toolbar", vbNullString)
hBar& = FindWindowEx(hBar&, 0&, "_AOL_Toolbar", vbNullString)

If hBar <> 0 Then
ChatSend40 (Text$)
Else
    
End If

End Sub
Function MDI() As Long
Dim hMain&, hMDI&
hMain = FindWindow("AOL Frame25", vbNullString)
hMDI = FindWindowEx(hMain, 0, "MDIClient", vbNullString)
MDI = hMDI
End Function

Private Function Room25() As Long
Dim hMain&, hView&, hEdit&, hListbox&, hIcon&

Do
    hMain = FindWindowEx(MDI, hMain, "AOL Child", vbNullString)
    hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hListbox = FindWindowEx(hMain, 0, "_AOL_Listbox", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until (hIcon <> 0 And hListbox <> 0 And hEdit <> 0 And hView <> 0) Or hMain = 0
Room25 = hMain
End Function

Public Sub ChatSend40(strmessage As String)
'chatsend for aol4.0 [ex: chatsend ("hey...")]
If FindChatWnd = False Then Exit Sub
   Dim lngchat As Long, lngrich As Long, strText As String
    Let lngchat& = findroom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strText$ = GetText(lngrich&)
     
    If strText$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "</b></u></i>" + strmessage$ + "")
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strText$)
End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function








Public Sub sncollect(TheList As Control, lnghowmanysns As Long)
    'this is great for a spammer
    'this uses who's chatting method"
    Dim blah$, Window&
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngchild As Long, lnglist1 As Long, lnglist As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngicon4 As Long, lngicon5 As Long
    Dim lngicon6 As Long, lngicon7 As Long, lngicon8 As Long, lngchatting As Long, lngcounter As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    pause 1
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Find a Chat")
        Let lnglist1& = FindWindowEx(lngchild&, 0&, "_AOL_Listbox", vbNullString)
        Let lnglist& = FindWindowEx(lngchild&, lnglist&, "_AOL_Listbox", vbNullString)
        Let lngicon1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon2& = FindWindowEx(lngchild&, lngicon1&, "_AOL_Icon", vbNullString)
        Let lngicon3& = FindWindowEx(lngchild&, lngicon2&, "_AOL_Icon", vbNullString)
        Let lngicon4& = FindWindowEx(lngchild&, lngicon3&, "_AOL_Icon", vbNullString)
        Let lngicon5& = FindWindowEx(lngchild&, lngicon4&, "_AOL_Icon", vbNullString)
        Let lngicon6& = FindWindowEx(lngchild&, lngicon5&, "_AOL_Icon", vbNullString)
        Let lngicon7& = FindWindowEx(lngchild&, lngicon6&, "_AOL_Icon", vbNullString)
        Let lngicon8& = FindWindowEx(lngchild&, lngicon7&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngchild&, lngicon8&, "_AOL_Icon", vbNullString)
        pause 1.5
    Loop Until lngchild& <> 0& And lnglist1& <> 0& And lnglist& <> 0& And lnglist1& <> lnglist& And lngicon& <> 0& And lngicon1& <> 0& And lngicon2& <> 0& And lngicon3& <> 0& And lngicon4& <> 0& And lngicon5& <> 0& And lngicon6& <> 0& And lngicon7& <> 0& And lngicon8& <> 0& And _
    lngicon1& <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon4& And lngicon4& <> lngicon5& And lngicon5& <> lngicon6& And lngicon6& <> lngicon7& And lngicon7& <> lngicon8& And lngicon8& <> lngicon&
    Call waitforlisttoload(lnglist&)
    Do: DoEvents
        Call clicklistindex(lnglist&, lngcounter&)
        Call ClickIcon(lngicon&)
        Let lngcounter& = lngcounter& + 1&
        Do: DoEvents
            Let lngchatting& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Who's Chatting")
            Let lnglist1& = FindWindowEx(lngchatting&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until lngchatting& <> 0& And lnglist1& <> 0&
        Call waitforlisttoload(lnglist1&)
        Call addlisttolistbox(lnglist1&, TheList)
        'blah$ = Form1.List1.ListCount
        'Form1.Label1.Caption = "" + blah$ + " names"
        Call PostMessage(lngchatting&, WM_CLOSE, 0&, 0&)
    Loop Until TheList.ListCount > lnghowmanysns&
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
    Window& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Find a Chat")
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
    MsgBox "auto sn collect complete"
    
End Sub




Public Function GetCaption(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = GetWindowTextLength(lngwindow&)
    Let strbuffer$ = String$(lngtextlen&, 0&)
    Call GetWindowText(lngwindow&, strbuffer$, lngtextlen& + 1&)
    Let GetCaption$ = strbuffer$
End Function

Public Sub clickignore(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngignorebut As Long, lnglist As Long
    Let lngmailbox& = findmailbox&
    Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
    Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
    Let lnglist& = FindWindowEx(lngtabwin&, 0&, "_AOL_List", vbNullString)
    Call PostMessage(lnglist&, LB_SETCURSEL, CLng(Index&), 0&)
    Let lngignorebut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngignorebut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngignorebut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub KillDupes(List As Control)
Dim a, i
For i = 0 To List.ListCount - 1
For a = 0 To List.ListCount - 1
If LCase(List.List(i)) Like LCase(List.List(a)) And i <> a Then
List.RemoveItem (a)
End If
Next a
Next i
End Sub


Public Sub KillWait()
    Dim lngmodal As Long, lngbut As Long
    Call runaolmenu(4&, 10&)
    Do: DoEvents
        lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        lngbut& = FindWindowEx(lngmodal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngmodal& <> 0& And lngbut& <> 0&
    Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function GetText(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GetTextLength, 0&, 0&)
    Let strbuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strbuffer$)
    Let GetText$ = strbuffer$
End Function
Public Function getlistitemtext(lnglist As Long, Index As Long) As String
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim Screenname As String, itmhold As Long, psnHold As Long
    Dim rbytes As Long, cprocess As Long
    Let rlist& = lnglist&
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        Let Screenname$ = String$(4, vbNullChar)
        Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
        Let itmhold& = itmhold& + 24
        Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4, rbytes&)
        Call CopyMemory(psnHold&, ByVal Screenname$, 4)
        Let psnHold& = psnHold& + 6
        Let Screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
        Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1)
        Let getlistitemtext$ = Screenname$
        Call CloseHandle(mthread&)
    End If
End Function




Public Function FindOpenMail() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& = 0& Then
        Let FindOpenMail& = lngchild&
        Exit Function
    End If
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
        Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
        If lngstatic& <> 0& Then
            Let FindOpenMail& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
    Let FindOpenMail& = 0&
End Function

Public Function findlocatewin() As Long
    Dim lngokwin As Long, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngchildcap As String, lnglocatewin As Long, lngstaticwin As Long
    Dim strstaticmsg As String, lngbutton As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngchildcap$ = GetCaption(lngchild&)
    If LCase$(lngchildcap$) Like LCase$("locate *") Then
        Let findlocatewin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngaol& = FindWindow("AOL Frame25", "America  Online")
            Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
            Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
            Let lngchildcap$ = GetCaption(lngchild&)
            If LCase$(lngchildcap$) Like LCase$("locate *") Then
                Let findlocatewin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
        Let findlocatewin& = 0&
    End If
End Function
Public Function locatemember(strname As String) As String
    Dim lngokwin As Long, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngchildcap As String, lnglocatewin As Long, lngstaticwin As Long
    Dim strstaticmsg As String, lngbutton As Long
    Call Keyword("aol://3548:" & strname$)
    Do: DoEvents
        Let lngokwin& = FindWindow("#32770", "America Online")
        If lngokwin& <> 0& Then
            Exit Do
        Else
            Let lngaol& = FindWindow("AOL Frame25", "America  Online")
            Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
            Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
            Let lngchildcap$ = GetCaption(lngchild&)
            If LCase$(lngchildcap$) = LCase$("locate " & strname$) Then
                Let lnglocatewin& = lngchild&
            Else
                Let lngaol& = FindWindow("AOL Frame25", "America  Online")
                Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
                Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
                Do: DoEvents
                    Let lngchildcap$ = GetCaption(lngchild&)
                    If LCase$(lngchildcap$) = LCase$("locate " & strname$) Then
                        Let lnglocatewin& = lngchild&
                        Exit Do
                    End If
                    Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
                    Let lngokwin& = FindWindow("#32770", "America Online")
                Loop Until lngchild& = 0& Or lngokwin& <> 0&
            End If
        End If
    Loop Until lnglocatewin& <> 0& Or lngokwin& <> 0&
    If lnglocatewin& <> 0& Then
        Let lngstaticwin& = FindWindowEx(lnglocatewin&, 0&, "_AOL_Static", vbNullString)
        Let strstaticmsg$ = GetText(lngstaticwin&)
        If LCase$(strstaticmsg$) = LCase$(strname$ & " is online, but not in a chat area.") Then
            Let locatemember$ = "" + strname$ + " is online, but not in a chat room"
        ElseIf LCase$(strstaticmsg$) = LCase$(strname$ & " is online, but in a private room.") Then
            Let locatemember$ = "" + strname$ + " is in a private room"
        ElseIf LCase$(strstaticmsg$) Like LCase$(strname$ & " is in chat room *") Then
            Let locatemember$ = "" + strname$ + " is in " & Right$(strstaticmsg$, Len(strstaticmsg$) - Len(strname$) + 17&)
        End If
        Call SendMessage(lnglocatewin&, WM_CLOSE, 0&, 0&)
    ElseIf lngokwin& <> 0& Then
        Let lngbutton& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Do: DoEvents
            Call PostMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
            Let lngokwin& = FindWindow("#32770", "America Online")
            Let lngbutton& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        Loop Until lngokwin& = 0& And lngbutton& = 0&
        Let locatemember$ = "" + strname$ + " is not online."
        End If
        MsgBox locatemember$
End Function

Public Sub sendmailbcc(strpeople As String, strsubject As String, strmessage As String)
    Call SendMail("(" & getuser$ & "," & strpeople$ & ")", strsubject$, strmessage$)
End Sub
Public Sub SendMail(strperson As String, strsubject As String, strmessage As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long, lngcomposeico As Long
    Dim lngmailwin As Long, lngeditwin As Long, lngsubjectwin As Long, lngrich As Long
    Dim lngsendbut As Long, lngsendbut1 As Long, lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long, lngsendbut9 As Long
    Dim lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If lngmdi& <= 0& Then
        Exit Sub
    Else
        Let lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
        Let lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
        Let lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
        Let lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
            Let lngeditwin& = FindWindowEx(lngmailwin&, 0&, "_aol_edit", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strperson$)
        Do: DoEvents
            Let lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
            Let lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
            Let lngrich& = FindWindowEx(lngmailwin&, 0&, "richcntl", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0& And lngrich& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strsubject$)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
        Do: DoEvents
            Let lngsendbut1& = FindWindowEx(lngmailwin&, 0&, "_aol_icon", vbNullString)
            Let lngsendbut2& = FindWindowEx(lngmailwin&, lngsendbut1&, "_aol_icon", vbNullString)
            Let lngsendbut3& = FindWindowEx(lngmailwin&, lngsendbut2&, "_aol_icon", vbNullString)
            Let lngsendbut4& = FindWindowEx(lngmailwin&, lngsendbut3&, "_aol_icon", vbNullString)
            Let lngsendbut5& = FindWindowEx(lngmailwin&, lngsendbut4&, "_aol_icon", vbNullString)
            Let lngsendbut6& = FindWindowEx(lngmailwin&, lngsendbut5&, "_aol_icon", vbNullString)
            Let lngsendbut7& = FindWindowEx(lngmailwin&, lngsendbut6&, "_aol_icon", vbNullString)
            Let lngsendbut8& = FindWindowEx(lngmailwin&, lngsendbut7&, "_aol_icon", vbNullString)
            Let lngsendbut9& = FindWindowEx(lngmailwin&, lngsendbut8&, "_aol_icon", vbNullString)
            Let lngsendbut10& = FindWindowEx(lngmailwin&, lngsendbut9&, "_aol_icon", vbNullString)
            Let lngsendbut11& = FindWindowEx(lngmailwin&, lngsendbut10&, "_aol_icon", vbNullString)
            Let lngsendbut12& = FindWindowEx(lngmailwin&, lngsendbut11&, "_aol_icon", vbNullString)
            Let lngsendbut13& = FindWindowEx(lngmailwin&, lngsendbut12&, "_aol_icon", vbNullString)
            Let lngsendbut& = FindWindowEx(lngmailwin&, lngsendbut13&, "_aol_icon", vbNullString)
        Loop Until lngsendbut& <> 0& And _
        lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
        lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
        lngsendbut13& <> lngsendbut&
        Do While lngmailwin& <> 0&
            Call SendMessageLong(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
            Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
            pause 0.2
        Loop: DoEvents
    End If
End Sub
Public Sub sendmailattachment(strperson As String, strsubject As String, strmessage As String, strfilepath As String, strfilename As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long, lngcomposeico As Long, lngokbut As Long
    Dim lngmailwin As Long, lngeditwin As Long, lngsubjectwin As Long, lngrich As Long, lngattachbut As Long
    Dim lngsendbut As Long, lngsendbut1 As Long, lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long, lngsendbut9 As Long
    Dim lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long, lngattachwin As Long
    Dim lngbrowsewin As Long, lngedit As Long, lngopenbut As Long, lngcombo As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If lngmdi& <= 0& Then
        Exit Sub
    Else
        Let lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
        Let lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
        Let lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
        Let lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngcomposeico&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
            Let lngeditwin& = FindWindowEx(lngmailwin&, 0&, "_aol_edit", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strperson$)
        Do: DoEvents
            Let lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
            Let lngeditwin& = FindWindowEx(lngmailwin&, lngeditwin&, "_aol_edit", vbNullString)
            Let lngrich& = FindWindowEx(lngmailwin&, 0&, "richcntl", vbNullString)
        Loop Until lngmailwin& <> 0& And lngeditwin& <> 0& And lngrich& <> 0&
        Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strsubject$)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
        Do: DoEvents
            Let lngsendbut1& = FindWindowEx(lngmailwin&, 0&, "_aol_icon", vbNullString)
            Let lngsendbut2& = FindWindowEx(lngmailwin&, lngsendbut1&, "_aol_icon", vbNullString)
            Let lngsendbut3& = FindWindowEx(lngmailwin&, lngsendbut2&, "_aol_icon", vbNullString)
            Let lngsendbut4& = FindWindowEx(lngmailwin&, lngsendbut3&, "_aol_icon", vbNullString)
            Let lngsendbut5& = FindWindowEx(lngmailwin&, lngsendbut4&, "_aol_icon", vbNullString)
            Let lngsendbut6& = FindWindowEx(lngmailwin&, lngsendbut5&, "_aol_icon", vbNullString)
            Let lngsendbut7& = FindWindowEx(lngmailwin&, lngsendbut6&, "_aol_icon", vbNullString)
            Let lngsendbut8& = FindWindowEx(lngmailwin&, lngsendbut7&, "_aol_icon", vbNullString)
            Let lngsendbut9& = FindWindowEx(lngmailwin&, lngsendbut8&, "_aol_icon", vbNullString)
            Let lngsendbut10& = FindWindowEx(lngmailwin&, lngsendbut9&, "_aol_icon", vbNullString)
            Let lngsendbut11& = FindWindowEx(lngmailwin&, lngsendbut10&, "_aol_icon", vbNullString)
            Let lngsendbut12& = FindWindowEx(lngmailwin&, lngsendbut11&, "_aol_icon", vbNullString)
            Let lngsendbut13& = FindWindowEx(lngmailwin&, lngsendbut12&, "_aol_icon", vbNullString)
            Let lngsendbut& = FindWindowEx(lngmailwin&, lngsendbut13&, "_aol_icon", vbNullString)
        Loop Until lngsendbut& <> 0& And _
        lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
        lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
        lngsendbut13& <> lngsendbut&
        Let lngattachbut& = lngsendbut13&
        Call PostMessage(lngattachbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngattachbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngattachwin& = FindWindow("_AOL_Modal", "Attachments")
            Let lngattachbut& = FindWindowEx(lngattachwin&, 0&, "_AOL_Icon", vbNullString)
            Let lngokbut& = FindWindowEx(lngattachwin&, lngattachbut&, "_AOL_Icon", vbNullString)
            Let lngokbut& = FindWindowEx(lngattachwin&, lngokbut&, "_AOL_Icon", vbNullString)
        Loop Until lngattachwin& <> 0& And lngattachbut& <> 0& And lngokbut& <> 0&
        Call PostMessage(lngattachbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngattachbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngbrowsewin& = FindWindow("#32770", "Attach")
            Let lngedit& = FindWindowEx(lngbrowsewin&, 0&, "Edit", vbNullString)
            Let lngopenbut& = FindWindowEx(lngbrowsewin&, 0&, "Button", "&Open")
            Let lngcombo& = FindWindowEx(lngbrowsewin&, 0&, "ComboBox", vbNullString)
        Loop Until lngbrowsewin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfilepath$)
        Call PostMessage(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngcombo& = FindWindowEx(lngbrowsewin&, 0&, "ComboBox", vbNullString)
        Loop Until GetText(lngcombo&) = strfilepath$
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfilename$)
        Do: DoEvents
            Let lngbrowsewin& = FindWindow("#32770", "Attach")
            Let lngopenbut& = FindWindowEx(lngbrowsewin&, 0&, "Button", "&Open")
            Call PostMessage(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
            pause 0.2
        Loop Until lngbrowsewin& = 0&
        Do: DoEvents
            Let lngattachwin& = FindWindow("_AOL_Modal", "Attachments")
            Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
            pause 0.2
        Loop Until lngattachwin& = 0&
        Do While lngmailwin& <> 0&
            Call SendMessageLong(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
            Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "aol child", "write mail")
            pause 0.2
        Loop: DoEvents
    End If
End Sub



Public Function Available(strname As String) As Boolean
    'checks to see if the sn is available
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngrich As Long, lngokwin As Long
    Dim lngicon As Long, lngbut As Long, lngstatic As Long, strmsg As String, lngedit As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & strname$)
    Do: DoEvents
        Let lngimwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Send Instant Message")
        Let lngrich& = FindWindowEx(lngimwin&, 0&, "RICHCNTL", vbNullString)
        Let lngedit& = FindWindowEx(lngimwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngimwin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngimwin& <> 0& And lngicon& <> 0& And lngrich& <> 0& And lngedit& <> 0&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngokwin& = FindWindow("#32770", "America Online")
        Let lngbut& = FindWindowEx(lngokwin&, 0&, "Button", vbNullString)
        Let lngstatic& = FindWindowEx(lngokwin&, 0&, "Static", vbNullString)
        Let lngstatic& = FindWindowEx(lngokwin&, lngstatic&, "Static", vbNullString)
    Loop Until lngokwin& <> 0& And lngbut& <> 0& And lngstatic& <> 0&
    Let strmsg$ = GetText(lngstatic&)
    Do: DoEvents
        Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngokwin& = FindWindow("#32770", "America Online")
        Let lngbut& = FindWindowEx(lngokwin&, 0&, "Button", vbNullString)
    Loop Until lngokwin& = 0& And lngbut& = 0&
    Call SendMessage(lngimwin&, WM_CLOSE, 0&, 0&)
    If LCase$(strmsg$) = LCase$(strname$) & " is online and able to receive instant messages." Then
        Let Available = True
    Else
        Let Available = False
    End If
End Function


Public Sub makeshortcut(strshortcutdir As String, strshortcutname As String, strshortcutpath As String)
    Dim strwinshortcutdir As String, strwinshortcutname As String, strwinshortcutexepath As String, lngretval As Long
    Let strwinshortcutdir$ = strshortcutdir$
    Let strwinshortcutname$ = strshortcutname$
    Let strwinshortcutexepath$ = strshortcutpath$
    Let lngretval& = fCreateShellLink("", strwinshortcutname$, strwinshortcutexepath$, "")
    Name "c:\windows\start menu\programs\" & strwinshortcutname$ & ".lnk" As strwinshortcutdir$ & "\" & strwinshortcutname$ & ".lnk"
End Sub
Public Function Online() As Boolean
'checks to see if you on u can use this like...
' If online=false then msgbox "you must signon to use this"
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long
    Dim lngcomposeico As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
    Let lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
    Let lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
    If lngcomposeico& <> 0& Then
        Let Online = True
    Else
        Let Online = False
    End If
End Function
Public Function master() As Boolean
'checks to see if the account you on is a master
    Dim lngaol As Long, lngmdi As Long, lngprntwindow As Long
    Dim lngprnticon As Long, lngmodal As Long, lngmodalstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do Until lngprntwindow& <> 0& And lngprnticon& <> 0&
        Let lngprntwindow& = FindWindowEx(lngmdi&, 0&, "AOL Child", " Parental Controls")
        Let lngprnticon& = FindWindowEx(lngprntwindow&, 0&, "_AOL_Icon", vbNullString)
    Loop: DoEvents
    Do: DoEvents
        Call PostMessage(lngprnticon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngprnticon&, WM_LBUTTONUP, 0&, 0&)
        Let lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        Let lngmodalstatic& = FindWindowEx(lngmodal&, 0&, "_AOL_Static", vbNullString)
    Loop Until lngmodal& <> 0 And lngmodalstatic& <> 0& And GetText(lngmodalstatic&) <> ""
    If replacestring(replacestring(GetText(lngmodalstatic&), Chr$(10&), ""), Chr$(13&), "") = "Set Parental Controls" Then
        Let master = True
    Else
        Let master = False
    End If
    Call SendMessageLong(lngmodal&, WM_CLOSE, 0&, 0&): DoEvents
    Call SendMessageLong(lngprntwindow&, WM_CLOSE, 0&, 0&): DoEvents
End Function
Public Sub SignOff()
    Call runaolmenu(3&, 1&)
    Call runaolmenubystring("&Sign Off")
End Sub
Public Sub sendimresponse(strmessage As String)
    Dim lngimwin As Long, lngrich As Long, lngicon As Long
    Dim lngicon1 As Long, lngicon2 As Long, lngicon3 As Long, lngicon4 As Long
    Dim lngicon5 As Long, lngicon6 As Long, lngicon7 As Long, lngicon8 As Long
    Let lngimwin& = findim&
    If lngimwin& = 0& Then Exit Sub
    Let lngrich& = FindWindowEx(lngimwin&, 0&, "RICHCNTL", vbNullString)
    Let lngrich& = FindWindowEx(lngimwin&, lngrich&, "RICHCNTL", vbNullString)
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Let lngicon1& = FindWindowEx(lngimwin&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon2& = FindWindowEx(lngimwin&, lngicon1&, "_AOL_Icon", vbNullString)
    Let lngicon3& = FindWindowEx(lngimwin&, lngicon2&, "_AOL_Icon", vbNullString)
    Let lngicon4& = FindWindowEx(lngimwin&, lngicon3&, "_AOL_Icon", vbNullString)
    Let lngicon5& = FindWindowEx(lngimwin&, lngicon4&, "_AOL_Icon", vbNullString)
    Let lngicon6& = FindWindowEx(lngimwin&, lngicon5&, "_AOL_Icon", vbNullString)
    Let lngicon7& = FindWindowEx(lngimwin&, lngicon6&, "_AOL_Icon", vbNullString)
    Let lngicon8& = FindWindowEx(lngimwin&, lngicon7&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngimwin&, lngicon8&, "_AOL_Icon", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function getimtext() As String
    Dim lngrich As Long
    Let lngrich& = FindWindowEx(findim&, 0&, "RICHCNTL", vbNullString)
    Let getimtext$ = GetText(lngrich&)
End Function
Public Function getimlastmsg() As String
    Dim strimmsg As String, lngtab As Long, lngnewtab As Long
    Let strimmsg$ = getimtext$
    Let lngtab& = InStr(strimmsg$, Chr$(9&))
    Do: DoEvents
        Let lngtab& = lngnewtab&
        Let lngnewtab& = InStr(lngtab& + 1&, strimmsg$, Chr$(9&))
    Loop Until lngnewtab& <= 0&
    Let strimmsg$ = Right$(strimmsg$, Len(strimmsg$) - lngtab& - 1&)
    Let getimlastmsg$ = Left$(strimmsg$, Len(strimmsg$) - 1&)
End Function
Public Function getimsn() As String
    Dim strimcap As String
    Let strimcap$ = GetCaption(findim&)
    If InStr(strimcap$, ":") = 0& Then
        Let getimsn$ = ""
        Exit Function
    Else
        Let getimsn$ = Right$(strimcap$, Len(strimcap$) - InStr(strimcap$, ":") - 1&)
    End If
End Function
Public Sub runaolmenubystring(strmenutext As String)
    Dim lngaol As Long, lngmmenu As Long, lngmmcount As Long
    Dim lngindex As Long, lngsubmenu As Long, lngsmcount As Long
    Dim lngindex2 As Long, lngsmid As Long, strString As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmmenu& = GetMenu(lngaol&)
    Let lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
        Let lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
        Let lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
            Let lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
            Let strString$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strString$, 100&, 1&)
            If InStr(LCase$(strString$), LCase$(strmenutext$)) Then
                Call SendMessageLong(lngaol&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub



Public Sub removelistitem(TheList As ListBox, strentry As String)
    Dim lngindex As Long
    If TheList.ListCount = 0& Then Exit Sub
    For lngindex& = 0& To TheList.ListCount - 1&
        Let TheList.ListIndex = lngindex&
        If LCase$(TheList.List(lngindex&)) = LCase$(strentry$) Then
            Call TheList.RemoveItem(lngindex&)
            Exit Sub
            If Err Then Exit Sub
        End If
    Next lngindex&
End Sub
Public Sub Loadlistbox(strnameandpath As String, TheList As ListBox)
  
    Dim strlinetext As String
    On Error Resume Next
    Open strnameandpath$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheList.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Sub SaveListBox(strnameandpath As String, TheList As ListBox)
    Dim Index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For Index& = 0& To TheList.ListCount - 1&
        Print #1&, TheList.List(Index&)
    Next Index&
    Close #1&
End Sub
Public Sub loadcombobox(strnameandpath As String, thecombo As ComboBox)
    'this won't work with all saved combos
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    Open strnameandpath$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thecombo.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Function findim() As Long
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        findim& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                findim& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    findim& = child&
End Function

Public Sub sendim(strperson As String, strmessage As String)
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngsendbut As Long
    Dim lngrich As Long, lngok As Long, lngokbut As Long, lngokwin2 As Long
    Dim lngokwin As Long, lngokbut2 As Long, lngokwin3 As Long, lngsendbut1 As Long
    Dim lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long
    Call Keyword("aol://9293:" & strperson$)
    Do: DoEvents
        Let lngimwin& = FindWindowEx(lngmdi&, 0&, "aol child", "send instant message")
        Let lngrich& = FindWindowEx(lngimwin&, 0&, "richcntl", vbNullString)
    Loop Until lngimwin& <> 0& And lngrich& <> 0&
    pause 0.15
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Do: DoEvents: Loop Until GetText(lngrich&) = strmessage$
    Do: DoEvents
        Let lngsendbut& = FindWindowEx(lngimwin&, 0&, "_aol_icon", vbNullString)
        Let lngsendbut1& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut2& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut3& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut4& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut5& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut6& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut7& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngimwin&, lngsendbut&, "_aol_icon", vbNullString)
    Loop Until lngsendbut& <> 0& And lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And lngsendbut7& <> lngsendbut&
    Let lngimwin& = FindWindowEx(lngmdi&, 0&, "aol child", "send instant message")
    Let lngokwin& = FindWindow("#32770", "america online")
    Let lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "ok")
    Let lngokwin2& = FindWindowEx(lngaol&, 0&, "#32770", vbNullString)
    Let lngokbut2& = FindWindowEx(lngokwin2&, 0&, "button", vbNullString)
    Do: DoEvents
        Let lngokwin2& = FindWindowEx(lngaol&, 0&, "#32770", vbNullString)
        Let lngokbut2& = FindWindowEx(lngokwin2&, 0&, "button", vbNullString)
        Call SendMessageLong(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        If LCase$(strperson$) = "$im_off" Then
            Do: DoEvents: Loop Until lngokwin2& <> 0&
        End If
        If LCase$(strperson$) = "$im_on" Then
            Do: DoEvents: Loop Until lngokwin2& <> 0&
        End If
    Loop Until lngsendbut& = 0& Or lngokwin2& <> 0& Or lngokwin& <> 0&
    If lngokwin2& <> 0& Then
        Do: DoEvents
            Call SendMessageLong(lngokwin2&, WM_CLOSE, 0&, 0&)
            Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
            pause 0.7
        Loop Until lngokwin2& = 0&
    End If
    Do: DoEvents
        Let lngokwin3& = FindWindowEx(lngaol&, lngokwin2&, "#32770", vbNullString)
        Let lngokbut2& = FindWindowEx(lngokwin3&, 0&, "button", vbNullString)
        Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessageLong(lngokwin&, WM_CLOSE, 0&, 0&)
        Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
    Loop Until lngokwin& = 0&
    If lngimwin& <> 0& Then Call SendMessageLong(lngimwin&, WM_CLOSE, 0&, 0&)
    
        Call PostMessage(findim, WM_CLOSE, 0&, 0&)
End Sub


Public Sub aolbrowserback()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aolbrowserforward()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aolbrowserstop()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aolbrowserrefresh()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aolbroswerhome()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngmdi&, 0, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub savecombobox(strnameandpath As String, thecombo As ComboBox)
    Dim Index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For Index& = 0& To thecombo.ListCount - 1&
        Print #1&, thecombo.List(Index&)
    Next Index&
    Close #1&
End Sub


Public Sub loadlistboxwithcommondialog(commondialog32 As Control, TheList As ListBox)
    'this won't work with all saved lists
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    With commondialog32
        .DialogTitle = "load a listbox"
        .CancelError = True
        .FilterIndex = 0&
        .ShowOpen
    End With
    Open commondialog32.filename For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            TheList.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Function gethexfromrgb(rgbvalue As Long) As String
    Dim hexstate As String, hexlen As Long
    Let hexstate$ = Hex(rgbvalue&)
    Let hexlen& = Len(hexstate$)
    Select Case hexlen&
        Case 1&
            Let gethexfromrgb$ = "00000" & hexstate$
            Exit Function
        Case 2&
            Let gethexfromrgb$ = "0000" & hexstate$
            Exit Function
        Case 3&
            Let gethexfromrgb$ = "000" & hexstate$
            Exit Function
        Case 4&
            Let gethexfromrgb$ = "00" & hexstate$
            Exit Function
        Case 5&
            Let gethexfromrgb$ = "0" & hexstate$
            Exit Function
        Case 6&
            Let gethexfromrgb$ = "" & hexstate$
            Exit Function
        Case Else
            Exit Function
    End Select
End Function

Public Sub formtotray(frmForm As Form)
     'puts icon in system tray!
    
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .uId = vbNull
        Let .hwnd = frmForm.hwnd
        Let .ucallbackMessage = WM_MOUSEMOVE
        Let .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        Let .hIcon = frmForm.Icon
        Let .szTip = frmForm.Caption
    End With
    Call Shell_NotifyIcon(NIM_ADD, systray)
    frmForm.Hide
End Sub
Public Sub formfromtray(frmForm As Form)
    'removes icon in system tray!
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .hwnd = frmForm.hwnd
        Let .uId = vbNull
    End With
    Call Shell_NotifyIcon(NIM_DELETE, systray)
    'frmform.Show
End Sub

Public Function FindChatWnd() As Long


Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long, AOLStatic As Long
Dim RICHCNTL As Long, AOLCombobox As Long, AOLIcon As Long, AOLStatic2 As Long
Dim RICHCNTL2 As Long, AOLIcon2 As Long, AOLImage As Long, AOLStatic3 As Long
Dim AOLListbox As Long, AOLStatic4 As Long, AOLIcon3 As Long, i As Long, Counter As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 6&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0& Or AOLStatic3& = 0& Or AOLListbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 6&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And RICHCNTL& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And AOLListbox& And AOLStatic4& And AOLIcon3& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChatWnd& = AOLChild&
    Exit Function
End If
End Function


Public Function findroom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        findroom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                findroom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    findroom& = child&
End Function


Public Function findmailbox() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngtabwin As Long, lngtabwin2 As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngtabwin& = FindWindowEx(lngchild&, 0&, "_AOL_TabControl", vbNullString)
    Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
    If lngtabwin& <> 0& And lngtabwin& <> 0& Then
        Let findmailbox& = lngchild&
        Exit Function
    End If
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
        Let lngtabwin& = FindWindowEx(lngchild&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        If lngtabwin& <> 0& And lngtabwin2& <> 0& Then
            Let findmailbox& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
    Let findmailbox& = 0&
End Function
Public Function findsignonwin() As Long
    Dim lngaol As Long, lngmdi As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let findsignonwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Sign On")
End Function
Public Function findsignoffwin() As Long
    Dim lngaol As Long, lngmdi As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let findsignoffwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Goodbye from America Online!")
End Function

Public Sub openflashmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyD, 0&)
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", "Incoming/Saved Mail") <> 0&
    Let lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", "Incoming/Saved Mail"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.X&, cursorpos.Y&)
End Sub
Public Function findflashmailbox() As Long
    Dim lngaol As Long, lngmdi As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let findflashmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
End Function
Public Sub openoldmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyO, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyO, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox") <> 0&
    Let lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
End Sub
Public Sub opensentmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyS, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyS, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox") <> 0&
    Let lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
End Sub
Public Sub opennewmailbox()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyR, 0&)
    Do: DoEvents: Loop Until FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox") <> 0&
    Let lngtree& = FindWindowEx(FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s Online Mailbox"), 0&, "_AOL_Tree", vbNullString)
    Call waitforlisttoload(lngtree&)
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
End Sub
Public Function getstringlineindex(strString As String, strlinetext As String) As Long
    Dim Index As Long
    For Index& = 0& To getstringlinecount(strString$)
        If getlinefromstring(strString$, Index&) = strlinetext$ Then
            Let getstringlineindex& = Index&
            Exit Function
        End If
    Next Index&
End Function
Public Function findforwardwin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long, strfwdcap As String
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    Let strfwdcap$ = Left$(GetCaption(lngchild&), 4&)
    If strfwdcap$ = "Fwd:" And lngstatic& <> 0& Then
        Let findforwardwin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            Let strfwdcap$ = Left$(GetCaption(lngchild&), 4&)
            If strfwdcap$ = "Fwd:" And lngstatic& <> 0& Then
                Let findforwardwin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
        Let findforwardwin& = lngchild&
    End If
End Function
Public Function findreplywin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long, strfwdcap As String
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    Let strfwdcap$ = Left$(GetCaption(lngchild&), 3&)
    If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
        Let findreplywin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            Let strfwdcap$ = Left$(GetCaption(lngchild&), 3&)
            If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
                Let findreplywin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
        Let findreplywin& = lngchild&
    End If
End Function

Public Function getflashmailcount() As Long
    Dim lngaol As Long, lngmdi As Long, lngmailwin As Long, lnglist As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let lnglist& = FindWindowEx(lngmailwin&, 0&, "_AOL_Tree", vbNullString)
    Let getflashmailcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function getnewmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        Let getnewmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function
Public Function getsentmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        Let getsentmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function
Public Function getoldmailcount() As Long
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        Let getoldmailcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
    End If
End Function

Public Sub IMsOFF()
Call imsend("$im_off", ".")

End Sub


Public Sub Keyword(Address As String)

Dim AOLFrame As Long, AOLToolbar As Long, AOLToolbar2 As Long
Dim AOLCombobox As Long, edit As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLCombobox& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
edit& = FindWindowEx(AOLCombobox&, 0&, "Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, Address$)
Call SendMessageLong(edit&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(edit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Function clicksignon()
Dim parhand1 As Long, parhand2 As Long, ourparent As Long, ourhandle As Long, hand1 As Long, hand2 As Long, hand3 As Long
parhand1& = FindWindow("AOL Frame25", "aol - Microsoft Internet Explorer")
parhand2& = FindWindowEx(parhand1&, 0, "MDIClient", vbNullString)
ourparent& = FindWindowEx(parhand2&, 0, "AOL Child", vbNullString)
hand1& = FindWindowEx(ourparent&, 0, "_AOL_Icon", vbNullString)
hand2& = FindWindowEx(ourparent&, hand1&, "_AOL_Icon", vbNullString)
hand3& = FindWindowEx(ourparent&, hand2&, "_AOL_Icon", vbNullString)
ourhandle& = FindWindowEx(ourparent&, hand3&, "_AOL_Icon", vbNullString)
clicksignon = ourhandle&
ClickIcon (clicksignon)
End Function
Public Sub addflashmailtolist(TheList As Control)
    
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    Dim strString As String, Index As Long, lnglen As Long, TheTab As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let strString$ = String$(255&, 0&)
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        For Index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&: DoEvents
            Let lnglen& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
            Let strString$ = String(lnglen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strString$)
            Let TheTab& = InStr(strString$, Chr$(9&))
            Let TheTab& = InStr(TheTab& + 1&, strString$, Chr$(9&))
            Let strString$ = Right$(strString$, Len(strString$) - TheTab&)
            Let strString$ = replacestring(strString$, Chr$(0&), "")
            TheList.AddItem strString$
        Next Index&
    End If
End Sub

Public Sub addnewmailtolist(TheList As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, Index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
            For Index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next Index&
        End If
    End If
End Sub

Public Sub addoldmailtolist(TheList As Control)
    
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, Index As Long
       Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
            For Index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next Index&
        End If
    End If
End Sub

Public Sub addsentmailtolist(TheList As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, Index As Long
    Dim lngtextlength As Long, strmailtext As String, lngtabkey As Long, Count As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Then
            Exit Sub
        Else
            For Index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                TheList.AddItem strmailtext$
                DoEvents
            Next Index&
        End If
    End If
End Sub

Public Function FindSendWin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& <> 0& Then
        Let FindSendWin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            If lngstatic& <> 0& Then
                Let FindSendWin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
    End If
    Let FindSendWin& = 0&
End Function


Public Sub pause(Length As Long)
    Dim current As Long
    Let current& = Timer
    Do Until (Timer - current&) >= Length&
        DoEvents
    Loop
End Sub

Public Sub clicklistindex(lnglist As Long, lngindex As Long)
    Call SendMessage(lnglist&, LB_SETCURSEL, CLng(lngindex&), 0&)
End Sub
Public Sub deletenewmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, lngbutton As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
            Let lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub
Public Sub deleteoldmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, lngbutton As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
            Let lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub
Public Sub deletesentmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, lngbutton As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
            Let lngbutton& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Let lngbutton& = FindWindowEx(lngmailbox&, lngbutton&, "_AOL_Icon", vbNullString)
            Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub
Public Sub deleteallnewmail()
    Dim Index As Long
    For Index& = 0& To getnewmailcount&
        Call deletenewmailitem(Index&)
    Next Index&
End Sub

Public Sub deletealloldmail()
    Dim Index As Long
    For Index& = 0& To getoldmailcount& - 1&
        Call deleteoldmailitem(Index&)
    Next Index&
End Sub

Public Sub deleteallsentmail()
    Dim Index As Long
    For Index& = 0& To getsentmailcount& - 1&
        Call deletesentmailitem(Index&)
    Next Index&
End Sub

Public Sub deleteallflashmail()
    Dim Index As Long
    For Index& = 0& To getflashmailcount& - 1&
        Call deleteflashmailitem(Index&)
    Next Index&
End Sub

Public Sub deleteflashmailitem(Index As Long)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngdelbut As Long, lngokwin As Long, lngokbut As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& < Index& Then
        Exit Sub
    Else
        Let lngdelbut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, Index&, 0&)
        Call PostMessage(lngdelbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngdelbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngokwin& = FindWindow("#32770", vbNullString)
            Let lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "&Yes")
        Loop Until lngokbut& <> 0&
        Do: DoEvents
            Let lngokwin& = FindWindow("#32770", vbNullString)
            Let lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "&Yes")
            Call SendMessageLong(lngokbut&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessageLong(lngokbut&, WM_KEYUP, VK_SPACE, 0&)
        Loop Until lngokbut& = 0&
    End If
End Sub
Public Sub formontop(frmForm As Form, OnTop As Boolean)
    If OnTop = True Then Call SetWindowPos(frmForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
    If OnTop = False Then Call SetWindowPos(frmForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub
Public Sub sendtext(lngwindow As Long, strText As String)
    Call SendMessageByString(lngwindow&, WM_SETTEXT, 0&, strText$)
End Sub
Public Sub windowmaximize(lnghwnd As Long)
    Call ShowWindow(lnghwnd, SW_MAXIMIZE)
End Sub
Public Sub IMIgnore(strwho As String, ignore As Boolean)
    Dim strstate As String
    If ignore = True Then
        Let strstate$ = "off"
    Else
        Let strstate$ = "on"
    End If
    Call sendim("$im_" & strstate$ & " " & strwho$, " ")
End Sub

Public Function doeswindowexist(lngwindow As Long) As Boolean
    If IsWindow(lngwindow&) Then
        Let doeswindowexist = True
    Else
        Let doeswindowexist = False
    End If
End Function

Public Sub ClickIcon(lngico As Long)
    Call SendMessageLong(lngico&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngico&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub clickbutton(lngbut As Long)
    Call SendMessageLong(lngbut&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(lngbut&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub AddRoomToListbox(TheControl As Control)
 Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let Room& = findroom&
    If Room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(Room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let Screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
            Let psnHold& = psnHold& + 6
            Let Screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Screenname$ = TrimSpaces(Screenname$)
          '  Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
            Call TheControl.AddItem(Screenname$ & "")
            If Screenname = getuser$ Then
            TheControl.RemoveItem (getuser)
            End If
        Next Index&
        Call CloseHandle(mthread&)
    End If
End Sub



Public Sub addlisttolistbox(rlist As Long, TheControl As Control)
    
    
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim sthread As Long, mthread As Long, strText As String
    If rlist& = 0& Then Exit Sub
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let Screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let Screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Screenname$ = TrimSpaces(Screenname$)
            'Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
            TheControl.AddItem (Screenname$)
        Next Index&
        Call CloseHandle(mthread&)
    End If
End Sub

Public Function AddListToString(rlist As Long, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim sthread As Long, mthread As Long
    If rlist& = 0& Then Exit Function
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let Screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let Screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
            Let AddListToString$ = Screenname$ & strseperator$
        Next Index&
        Call CloseHandle(mthread&)
    End If
End Function
Function converthr(ByVal minutes As Integer) As String
Dim hour As Integer, hour2 As Integer, Part1 As String, Part2 As String
If Not minutes < 60 Then
Do
hour = (hour + 1)
minutes = (minutes - 60)
Loop Until minutes < 60
End If
Part1 = hour
Part2 = minutes
If hour > 59 Then
Do
hour2 = (hour2 + 1)
hour = (hour - 60)
Loop Until hour < 60
Part1 = hour2 & ":" & hour
End If
If minutes < 10 Then Part2 = "0" & minutes
converthr = Part1 + ":" + Part2
End Function
Public Function Findimwnd() As Long
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long, Counter As Long
Dim AOLView As Long, AOLStatic As Long, RICHCNTL As Long, i As Long
Dim AOLIcon As Long, RICHCNTL2 As Long, AOLIcon2 As Long, AOLStatic2 As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLView& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLView& = 0& Or AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLStatic2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 5&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    If AOLView& And AOLStatic& And RICHCNTL& And AOLIcon& And RICHCNTL2& And AOLIcon2& And AOLStatic2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    Findimwnd& = AOLChild&
    Exit Function
End If
End Function
Public Sub setprofile(strname As String, strlocation As String, strbday As String, male As Boolean, female As Boolean, nosex As Boolean, strlovestatus As String, strhobbies As String, strcomputer As String, stroccupation As String, strquote As String)

    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, mousecur As POINTAPI
    Dim lngmenu As Long, lngmdi As Long, lngprofwin As Long, lngname As Long
    Dim lngmsex As Long, lngfsex As Long, lngnosex As Long, lngupbut As Long
    Dim lnglocation As Long, lngbirthday As Long, lngmstatus As Long, lnghobbies As Long
    Dim computer As Long, lngother As Long, lngquote As Long, lngcomputer As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyY, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyY, 0&)
    Call SetCursorPos(mousecur.X, mousecur.Y)
    Do: DoEvents
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Edit Your Online Profile")
        Let lngname& = FindWindowEx(lngprofwin&, 0&, "_AOL_Edit", vbNullString)
        Let lnglocation& = FindWindowEx(lngprofwin&, lngname&, "_AOL_Edit", vbNullString)
        Let lngbirthday& = FindWindowEx(lngprofwin&, lnglocation&, "_AOL_Edit", vbNullString)
        Let lngmstatus& = FindWindowEx(lngprofwin&, lngbirthday&, "_AOL_Edit", vbNullString)
        Let lnghobbies& = FindWindowEx(lngprofwin&, lngmstatus&, "_AOL_Edit", vbNullString)
        Let lngcomputer& = FindWindowEx(lngprofwin&, lnghobbies&, "_AOL_Edit", vbNullString)
        Let lngother& = FindWindowEx(lngprofwin&, lngcomputer&, "_AOL_Edit", vbNullString)
        Let lngquote& = FindWindowEx(lngprofwin&, lngother&, "_AOL_Edit", vbNullString)
        Let lngmsex& = FindWindowEx(lngprofwin&, 0&, "_AOL_Checkbox", vbNullString)
        Let lngfsex& = FindWindowEx(lngprofwin&, lngmsex&, "_AOL_Checkbox", vbNullString)
        Let lngnosex& = FindWindowEx(lngprofwin&, lngfsex&, "_AOL_Checkbox", vbNullString)
        Let lngupbut& = FindWindowEx(lngprofwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngupbut& = FindWindowEx(lngprofwin&, lngupbut&, "_AOL_Icon", vbNullString)
    Loop Until lngprofwin& <> 0& And lngquote& <> 0& And lngnosex& <> 0& And lngupbut& <> 0& And lnglocation <> 0&
     pause 3
    Call SendMessageByString(lngname&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lngname&, WM_SETTEXT, 0&, strname$): DoEvents
    Call SendMessageByString(lnglocation&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lnglocation&, WM_SETTEXT, 0&, strlocation$): DoEvents
    Call SendMessageByString(lngbirthday&, WM_SETTEXT, 0&, strbday$): DoEvents
    Call SendMessageByString(lngbirthday&, WM_SETTEXT, 0&, strbday$): DoEvents
    Call SendMessageByString(lngmstatus&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lngmstatus&, WM_SETTEXT, 0&, strlovestatus$): DoEvents
    Call SendMessageByString(lnghobbies&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lnghobbies&, WM_SETTEXT, 0&, strhobbies$): DoEvents
    Call SendMessageByString(lngcomputer&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lngcomputer&, WM_SETTEXT, 0&, strcomputer$): DoEvents
    Call SendMessageByString(lngquote&, WM_SETTEXT, 0&, ""): DoEvents
    Call SendMessageByString(lngquote&, WM_SETTEXT, 0&, strquote$): DoEvents
    If male = True Then
        Call PostMessage(lngmsex&, BM_SETCHECK, True, 0&)
    ElseIf female = True Then
        Call PostMessage(lngfsex&, BM_SETCHECK, True, 0&)
    ElseIf nosex = True Then
        Call PostMessage(lngnosex&, BM_SETCHECK, True, 0&)
    End If
    'Call PostMessage(lngupbut&, WM_CHAR, ENTER_KEY, 0&)
    Do
    Call PostMessage(lngupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngupbut&, WM_LBUTTONUP, 0&, 0&)
    Loop Until lngupbut& <> 0
    Call waitforok
End Sub
Public Function addlisttoclipboard(rlist As Long, strseperator As String) As String
    
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim sthread As Long, mthread As Long, strText As String
    If rlist& = 0& Then Exit Function
    Clipboard.Clear
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let Screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let Screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1&)
            Let strText$ = strText$ & Screenname$ & strseperator$
        Next Index&
        Call CloseHandle(mthread&)
    End If
    Clipboard.SetText strText$
End Function

Public Function addroomtostring(AddUser As Boolean, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, Screenname As String
        Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, strString As String
    Let Room& = findroom&
    If Room& = 0& Then
        Exit Function
    Else
        Let rlist& = FindWindowEx(Room&, 0&, "_aol_listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For Index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
                Let Screenname$ = String$(4, vbNullChar)
                Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
                Let itmhold& = itmhold& + 24
                Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4, rbytes&)
                Call CopyMemory(psnHold&, ByVal Screenname$, 4)
                Let psnHold& = psnHold& + 6
                Let Screenname$ = String$(16, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
                Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1)
                If Screenname$ <> getuser$ Or AddUser = True Then
                    Let strString$ = strString$ & strseperator$ & Screenname$
                End If
            Next Index&
            Call CloseHandle(mthread&)
        End If
        Let strString$ = Left$(strString$, Len(strString$) - 1&)
        Let addroomtostring$ = strString$
    End If
End Function

Public Function getprofile(strperson As String) As String
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim mousecur As POINTAPI, lngprofwin As Long, lngedit As Long, lngbutton As Long
    Dim lngview As Long, lngnoprofwin As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do Until IsWindowVisible(lngmenu&) = 1&
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop: DoEvents
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.X, mousecur.Y)
    Do: DoEvents
        Let lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Get a Member's Profile")
        Let lngedit& = FindWindowEx(lngprofwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lngprofwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngprofwin& <> 0& And lngedit& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strperson$)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngprofwin&, WM_CLOSE, 0&, 0&)
    Do: DoEvents
        Let lngprofwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Member Profile")
        Let lngview& = FindWindowEx(lngprofwin&, 0&, "_AOL_View", vbNullString)
        Let lngnoprofwin& = FindWindow("#32770", "America Online")
    Loop Until lngprofwin& <> 0& And lngview& <> 0& Or lngnoprofwin& <> 0&
    If lngnoprofwin& = 0& Then
        Let getprofile$ = GetText(lngview&)
        Call PostMessage(lngprofwin&, WM_CLOSE, 0&, 0&)
    ElseIf lngnoprofwin& <> 0& Then
        Let lngbutton& = FindWindowEx(lngnoprofwin&, 0&, "Button", "OK")
        Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(lngbutton&, WM_CLOSE, 0&, 0&)
        Let getprofile$ = "[no profile]"
    End If
End Function

Public Sub closeallims()
    Do: DoEvents
        If findim& <> 0& Then
            Call PostMessage(findim&, WM_CLOSE, 0&, 0&): DoEvents
        End If
    Loop Until findim& = 0&
End Sub
Public Sub closeallchildwindows(closebuddylist As Boolean)
    Dim child As Long, lngCount As Long
    Do: DoEvents
        child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
        If child& <> 0& Then
            If GetCaption(child&) = "Buddy List Window" And closebuddylist = False Then
                lngCount& = lngCount& + 1&
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), child&, "AOL Child", vbNullString)
                Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
            End If
            Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
        End If
    Loop Until child& = 0& Or lngCount& = 5&
End Sub
Public Sub setmailprefrences()
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim lngprefwin As Long, lngbut As Long, lnggenstatic As Long, lngmailstatic As Long
    Dim lngfontstatic As Long, lngmarketstatic As Long, lngmodal As Long, chkconfirm As Long
    Dim chkclose As Long, chkspellchk As Long, lngokbut As Long, cursorpos As POINTAPI
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(cursorpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyP, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyP, 0&)
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
    Do: DoEvents
        Let lngmodal& = FindWindow("_AOL_Modal", "Mail Preferences")
        Let lngokbut& = FindWindowEx(lngmodal&, 0&, "_AOL_icon", vbNullString)
        pause 1
    Loop Until lngmodal& <> 0& And lngokbut& <> 0&
    Let chkconfirm& = FindWindowEx(lngmodal&, 0&, "_AOL_Checkbox", vbNullString)
    Let chkclose& = FindWindowEx(lngmodal&, chkconfirm&, "_AOL_Checkbox", vbNullString)
    Let chkspellchk& = FindWindowEx(lngmodal&, chkclose&, "_AOL_Checkbox", vbNullString)
    Let chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
    Let chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
    Let chkspellchk& = FindWindowEx(lngmodal&, chkspellchk&, "_AOL_Checkbox", vbNullString)
    Call SendMessage(chkconfirm&, BM_SETCHECK, False, 0&)
    Call SendMessage(chkclose&, BM_SETCHECK, True, 0&)
    Call SendMessage(chkspellchk&, BM_SETCHECK, False, 0&)
    Call SendMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        lngokbut& = FindWindowEx(lngmodal&, 0&, "_AOL_icon", vbNullString)
    Loop Until lngokbut& = 0&
    Call PostMessage(lngprefwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub scrollmultitextbox(strText As String)
    Dim lngcounter As Long
    If Mid(strText$, Len(strText$), 1&) <> Chr$(10&) Then strText$ = strText$ & Chr$(13&) & Chr$(10&)
    Do While InStr(strText$, Chr$(13&)) <> 0&
        Let lngcounter& = lngcounter& + 1&
        Call ChatSend40(Mid$(strText$, 1&, InStr(strText$, Chr$(13&)) - 1&))
        If lngcounter = 3& Then
            Call pause(0.6)
            Let lngcounter& = 0&
        End If
        Let strText$ = Mid$(strText$, InStr(strText$, Chr$(13&) & Chr$(10&)) + 2&)
    Loop
End Sub

Public Sub FormDrag(frmForm As Form)
    'put this in mousedown for borderless forms
    Call ReleaseCapture
    Call SendMessage(frmForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub formcenter(frmForm As Form)
    Let frmForm.Top = (Screen.Height * 0.85) / 2& - frmForm.Height / 2&
    Let frmForm.Left = Screen.Width / 2& - frmForm.Width / 2&
End Sub



Public Sub formbevellines(frmForm As Form, lngside As Long, lngwidth As Long, Color As Long)
    'used by forminnerbevel, and formouterbevel
    Dim X1 As Long, Y1 As Long, x2 As Long, Y2 As Long, rightx As Long, bottomy As Long
    Dim x1b As Long, x2b As Long, y1b As Long, y2b As Long, Index As Long
    Let rightx = frmForm.ScaleWidth - 1&
    Let bottomy = frmForm.ScaleHeight - 1&
    Select Case lngside
        Case 0&
            Let X1& = 0&
            Let x1b& = 1&
            Let x2& = 0&
            Let x2b& = 1&
            Let Y1& = 0&
            Let y1b& = 1&
            Let Y2& = bottomy& + 1&
            Let y2b = -1&
        Case 1&
            Let X1& = rightx&
            Let x1b& = -1&
            Let x2& = X1&
            Let x2b& = x1b&
            Let Y1& = 0&
            Let y1b& = 1&
            Let Y2& = bottomy& + 1&
            Let y2b& = -1&
        Case 2&
            Let X1& = 0&
            Let x1b& = 1&
            Let x2& = rightx&
            Let x2b& = -1&
            Let Y1& = 0&
            Let y1b& = 1&
            Let Y2& = 0&
            Let y2b& = 1&
        Case 3&
            Let X1& = 1&
            Let x1b& = 1&
            Let x2& = rightx& + 1&
            Let x2b& = -1&
            Let Y1& = bottomy&
            Let y1b& = -1&
            Let Y2& = Y1&
            Let y2b& = y1b&
    End Select
    For Index& = 1& To lngwidth&
        frmForm.Line (X1&, Y1&)-(x2&, Y2&), Color&
        Let X1& = X1& + x1b&
        Let x2& = x2& + x2b&
        Let Y1& = Y1& + y1b&
        Let Y2& = Y2& + y2b&
    Next Index&
End Sub
Public Sub formouterbevel(frmForm As Form, lngbevelwidth As Long)
    'adds border to form
    Let frmForm.ScaleMode = 3&
    Call formbevellines(frmForm, 0&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 1&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 2&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 3&, lngbevelwidth&, QBColor(8&))
End Sub
Public Sub forminnerbevel(frmForm As Form, lngbevelwidth As Long)
    'adds border to form
    Let frmForm.ScaleMode = 3&
    Call formbevellines(frmForm, 0&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 1&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmForm, 2&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmForm, 3&, lngbevelwidth&, QBColor(15&))
End Sub


Public Function getchatname() As String
    Let getchatname$ = GetCaption(findroom&)
End Function

Public Sub addroomtoclipboard(AddUser As Boolean)
    Dim cprocess As Long, itmhold As Long, Screenname As String
    Dim psnHold As Long, rbytes As Long, Index As Long, Room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, strString As String
    Let Room& = findroom&
    If Room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(Room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
            Let Screenname$ = String$(4, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, Screenname$, 4, rbytes&)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4)
            Let psnHold& = psnHold& + 6
            Let Screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, Screenname$, Len(Screenname$), rbytes&)
            Let Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1)
            If Screenname$ <> getuser$ Or AddUser = True Then
                Let strString$ = strString$ & "," & Screenname$
            End If
        Next Index&
        Call CloseHandle(mthread&)
    End If
    If Left$(strString$, 1&) = "," Then
        Let strString$ = Right$(strString$, Len(strString$) - 1&)
    End If
    If Right$(strString$, 1&) = "," Then
        Let strString$ = Left$(strString$, Len(strString$) - 1&)
    End If
    Call Clipboard.Clear
    Call Clipboard.SetText(strString$)
End Sub

Public Function GetChatText() As String
    Dim ChatWin As Long
    Let ChatWin& = FindWindowEx(findroom&, 0&, "RICHCNTL", vbNullString)
    Let GetChatText$ = GetText(ChatWin&)
End Function
Public Function getchatline() As String
    Dim ChatWin As Long, strText As String, lngenter As Long, lngenter2 As Long
    Let ChatWin& = findroom&
    Let ChatWin& = FindWindowEx(ChatWin&, 0&, "RICHCNTL", vbNullString)
    Let strText$ = GetText(ChatWin&)
    Let lngenter& = InStr(strText$, Chr$(13&))
    Do: DoEvents
        If lngenter& <> 0& Then Let lngenter2& = lngenter&
        Let lngenter& = InStr(lngenter& + 1&, strText$, Chr$(13&))
    Loop Until lngenter& = 0&
    Let getchatline$ = Right$(strText$, Len(strText$) - lngenter2&)
End Function
Public Function getchatlinemsg() As String
    Dim strLine As String, lngtab As Long
    Let strLine$ = getchatline$
    Let lngtab& = InStr(strLine$, Chr$(9&))
    Let getchatlinemsg$ = Right$(strLine$, Len(strLine$) - lngtab&)
End Function
Public Function getchatlinesn() As String
    Dim strLine As String, lngtab As Long
    Let strLine$ = getchatline$
    Let lngtab& = InStr(strLine$, Chr$(9&))
    Let getchatlinesn$ = Left$(strLine$, lngtab& - 3&)
End Function







Public Function findwindows(lpclassname1 As String, lpclassname2 As String, lpclassname3 As String, lpclassname4 As String, lpclassname5 As String, lpclassname6 As String, lpclassname7 As String, lpclassname8 As String, lpclassname9 As String, lpclassname10 As String) As Long
    'from frenzy3.bas by izekial(me)
    'the variables lpclassname etc. should
    'be set to the class name of the window
    'you are trying to find, when you run
    'out of windows you want to find, put
    '"" for the left over variables.
    Dim lngwin As Long, lngwin2 As Long, lngwin3 As Long
    Dim lngwin4 As Long, lngwin5 As Long, lngwin6 As Long
    Dim lngwin7 As Long, lngwin8 As Long, lngwin9 As Long, lngwin10 As Long
    Let lngwin& = FindWindow(lpclassname1$, vbNullString)
    If lpclassname2$ = "" Then
        Let findwindows& = lngwin&
        Exit Function
    Else
        Let lngwin2& = FindWindowEx(lngwin&, 0&, lpclassname2$, vbNullString)
        If lpclassname3$ = "" Then
            Let findwindows& = lngwin2&
            Exit Function
        Else
            Let lngwin3& = FindWindowEx(lngwin2&, 0&, lpclassname3$, vbNullString)
            If lpclassname4$ = "" Then
                Let findwindows& = lngwin3&
                Exit Function
            Else
                Let lngwin4& = FindWindowEx(lngwin3&, 0&, lpclassname4$, vbNullString)
                If lpclassname5$ = "" Then
                    Let findwindows& = lngwin4&
                    Exit Function
                Else
                    Let lngwin5& = FindWindowEx(lngwin4&, 0&, lpclassname5$, vbNullString)
                    If lpclassname6$ = "" Then
                        Let findwindows& = lngwin5&
                        Exit Function
                    Else
                        Let lngwin6& = FindWindowEx(lngwin5&, 0&, lpclassname6$, vbNullString)
                        If lpclassname7$ = "" Then
                            Let findwindows& = lngwin6&
                            Exit Function
                        Else
                            Let lngwin7& = FindWindowEx(lngwin6&, 0&, lpclassname7$, vbNullString)
                            If lpclassname8$ = "" Then
                                Let findwindows& = lngwin7&
                                Exit Function
                            Else
                                Let lngwin8& = FindWindowEx(lngwin7&, 0&, lpclassname8$, vbNullString)
                                If lpclassname9$ = "" Then
                                    Let findwindows& = lngwin8&
                                    Exit Function
                                Else
                                    lngwin9& = FindWindowEx(lngwin8&, 0&, lpclassname9$, vbNullString)
                                    If lpclassname10$ = "" Then
                                        Let findwindows& = lngwin9&
                                        Exit Function
                                    Else
                                        Let lngwin10& = FindWindowEx(lngwin9&, 0&, lpclassname10$, vbNullString)
                                        Let findwindows& = lngwin10&
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
Public Function findwindowsex(lngnumberover As Long, lngwindowhandle As Long, lngparenthandle As Long, strclassname As String) As Long
    
    Do: DoEvents
        Let findwindowsex& = FindWindowEx(lngparenthandle&, lngwindowhandle&, strclassname$, vbNullString)
        Let lngnumberover& = lngnumberover& - 1&
    Loop Until lngnumberover& = 0&
End Function
Public Sub privateroom(strroom As String)
    Call Keyword("aol://2719:2-2-" & strroom$)
End Sub
Public Sub publicroom(strroom As String)
    Call Keyword("aol://2719:21-2-" & strroom$)
End Sub
Public Function FileExists(strfilepathandname As String) As Boolean
    'checks to see if the file is found!
    Dim lngcheck As Long
    If InStr(strfilepathandname$, ".") = 0& Then
        Let FileExists = False
    End If
    Let lngcheck& = Len(Dir$(strfilepathandname$))
    If lngcheck& = 0& Then
        Let FileExists = False
        Exit Function
    Else
        Let FileExists = True
    End If
End Function
Public Sub filecopy(File$, destination$)
    If Not FileExists(File$) Then Exit Sub
    If InStr(File$, ".") = 0& Then Exit Sub
    If InStr(destination$, "\") = 0& Then Exit Sub
    Call filecopy(File$, destination$)
End Sub
Public Sub openflashmailitem(lngindex As Long)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    lngaol& = FindWindow("AOL Frame25", vbNullString)
    lngmdi& = FindWindowEx(lngaol&, 0, "MDIClient", vbNullString)
    lngmailbox& = FindWindowEx(lngmdi&, 0, "AOL Child", "Incoming/Saved Mail")
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        If lngindex& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, lngindex&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, lngindex&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub openoldmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngCount As Long, lngindex As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub opennewmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngCount As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub opensentmailitem(Index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngCount As Long, lngindex As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, Index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub WAVStop()
    Call sndPlaySound("", SND_FLAG)
End Sub
Public Sub WAVLoop(File As String)
    Dim lngflags As Long
    If FileExists(File$) = False Then
        Exit Sub
    Else
        Let lngflags& = SND_ASYNC Or SND_LOOP
        Call sndPlaySound(File$, lngflags&)
    End If
End Sub
Public Sub wavloopstop()
    Dim lngflags As Long
    Let lngflags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound("", lngflags&)
End Sub
Public Sub clickaolmenuicon(lngiconhandle As Long, lngmenuentry As Long)
    Dim mousecur As POINTAPI, lngmenu As Long, Index As Long
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngiconhandle&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngiconhandle&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    For Index& = 1& To lngmenuentry&
        Call PostMessage(lngmenu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(lngmenu&, WM_KEYUP, VK_DOWN, 0&)
    Next Index&
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.X, mousecur.Y)
End Sub


Public Function listfindstring(List As ListBox, findstring As String) As Long
    Dim Index As Long
    If List.ListCount = 0 Then Exit Function
    For Index& = 0 To List.ListCount - 1
        Let List.ListIndex = Index&
        If UCase(List.Text) = UCase(findstring$) Then
            Let listfindstring& = Index&
            Exit Function
            If Err Then Exit Function
        End If
    Next Index&
End Function
Public Sub filedelete(File$)
    If FileExists(File$) = False Then
        Exit Sub
    Else
        If InStr(File$, ".") = 0& Then
            Exit Sub
        Else
            Call Kill(File$)
            DoEvents
        End If
    End If
End Sub

Public Function filegetattributes(FilePath As String) As Integer
    
    Dim strcheck As String
    Let strcheck$ = Dir(FilePath$)
    If strcheck$ = "" Then
        MsgBox "file not found", vbCritical, "tool.bas": Exit Function
    Else
        Let filegetattributes = GetAttr(FilePath$)
    End If
End Function
Public Function ReadINI(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let ReadINI$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function
Public Sub writeini(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub





Public Sub memberroom(strroom As String)
    Call Keyword("aol://2719:61-2-" & strroom$)
End Sub
Public Sub midiplay(strmidi As String)
    Dim strfile As String
    Let strfile$ = Dir(strmidi$)
    If strfile$ = "" Then
        Exit Sub
    Else
        Call mciSendString("play " & strmidi$, 0&, 0, 0)
    End If
End Sub
Public Sub midistop(strmidi As String)
    Dim strfile As String
    strfile$ = Dir(strmidi$)
    If strfile$ = "" Then
        Exit Sub
    Else
        Call mciSendString("stop " & strmidi$, 0&, 0, 0)
    End If
End Sub



Public Sub sendmassim(strwho As Control, strmessage As String)
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call sendim(strwho.List(lngindex&), strmessage$)
        pause 0.7
    Next lngindex&
End Sub


Public Function getflashmailsubject(Index As Long) As String
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) < Index& - 1& Then
        Exit Function
    Else
        If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) = 0& Or Index& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lnglist&, LB_GETTEXTLEN, Index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lnglist&, LB_GETTEXT, Index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getflashmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Function getnewmailsubject(Index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getnewmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Function getoldmailsubject(Index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getoldmailsubject$ = strmailtext$
        End If
    End If
End Function

Public Function getsentmailsubject(Index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or Index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or Index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, Index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, Index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getsentmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Sub forwardallflashmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim Index As Long
    Call openflashmailbox
    For Index& = 0& To getflashmailcount& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardallflashmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim Index As Long
    Call openflashmailbox
    Let progbar.Max = getflashmailcount& - 1&
    For Index& = 0& To getflashmailcount& - 1&
        Let progbar.Value = Index& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardallnewmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim Index As Long
    Call opennewmailbox
    For Index& = 0& To getnewmailcount& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardallnewmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim Index As Long
    Call opennewmailbox
    Let progbar.Max = getnewmailcount& - 1&
    For Index& = 0& To getnewmailcount& - 1&
        Let progbar.Value = Index& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardalloldmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim Index As Long
    Call openoldmailbox
    For Index& = 0& To getoldmailcount& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardallsentmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim Index As Long
    Call opensentmailbox
    For Index& = 0& To getsentmailcount& - 1&
        Call forwardsentmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardalloldmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim Index As Long
    Call openoldmailbox
    Let progbar.Max = getoldmailcount& - 1&
    For Index& = 0& To getoldmailcount& - 1&
        Let progbar.Value = Index& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardallsentmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim Index As Long
    Call opensentmailbox
    Let progbar.Max = getsentmailcount& - 1&
    For Index& = 0& To getsentmailcount& - 1&
        Let progbar.Value = Index& - 1&
        Call forwardsentmailitem(True, strwho$, strmessage$, Index&, deletefwdtext): DoEvents
    Next Index&
End Sub
Public Sub forwardflashmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, Index As Long, deletefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call openflashmailbox
    End If
    Call openflashmailitem(Index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(Index&))
        Let lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
        Let lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
        Let lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
        Let lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
        Let lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
        Let lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(Index&))
        Let lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
        Let lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
        Let lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
        Let lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
        Let lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
        Let lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
        Let lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
        Let lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
        Let lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
        Let lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
        Let lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
        Let lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
        Let lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
        Let lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
        Let lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
        Let lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
        Let lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
        Let lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        pause 0.3
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deletefwdtext = True Then
        Let strnewsubject$ = GetText(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(Index&))
        Let lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(Index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(Index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(Index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardnewmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, Index As Long, deletefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call opennewmailbox
    End If
    Call opennewmailitem(Index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(Index&))
        Let lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
        Let lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
        Let lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
        Let lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
        Let lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
        Let lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(Index&))
        Let lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
        Let lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
        Let lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
        Let lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
        Let lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
        Let lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
        Let lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
        Let lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
        Let lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
        Let lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
        Let lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
        Let lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
        Let lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
        Let lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
        Let lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
        Let lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
        Let lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
        Let lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        pause 0.2
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deletefwdtext = True Then
        Let strnewsubject$ = GetText(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(Index&))
        Let lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(Index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(Index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(Index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardoldmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, Index As Long, deletefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call openoldmailbox
    End If
    Call openoldmailitem(Index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(Index&))
        Let lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
        Let lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
        Let lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
        Let lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
        Let lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
        Let lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(Index&))
        Let lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
        Let lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
        Let lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
        Let lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
        Let lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
        Let lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
        Let lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
        Let lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
        Let lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
        Let lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
        Let lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
        Let lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
        Let lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
        Let lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
        Let lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
        Let lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
        Let lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
        Let lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        pause 0.2
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deletefwdtext = True Then
        Let strnewsubject$ = GetText(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(Index&))
        Let lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(Index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(Index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(Index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardsentmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, Index As Long, deletefwdtext As Boolean)
    Dim lngsendwin As Long, lngtowin As Long, lngccwin As Long, lngsubjectwin As Long
    Dim lngrichwin As Long, lngfontcombo As Long, lngcombo As Long, lngsendbut As Long
    Dim lngbut As Long, strnewsubject As String, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngerrorwin As Long, lngerrorbut As Long, lngsendbut2 As Long, lngfwdbut As Long
    Dim lngfwdbut1 As Long, lngfwdbut2 As Long, lngfwdbut3 As Long, lngfwdbut4 As Long, lngfwdbut5 As Long
    Dim lngfwdbut6 As Long, lngsendbut1 As Long, lngsendbut3 As Long
    Dim lngsendbut4 As Long, lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long, lngsendbut8 As Long
    Dim lngsendbut9 As Long, lngsendbut10 As Long, lngsendbut11 As Long, lngsendbut12 As Long, lngsendbut13 As Long
    Dim lngsendbut14 As Long, lngsendbut15 As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    If mailboxopen = False Then
        Call opensentmailbox
    End If
    Call opensentmailitem(Index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getsentmailsubject(Index&))
        Let lngfwdbut1& = FindWindowEx(lngchild&, 0&, "_AOL_Icon", vbNullString)
        Let lngfwdbut2& = FindWindowEx(lngchild&, lngfwdbut1&, "_AOL_Icon", vbNullString)
        Let lngfwdbut3& = FindWindowEx(lngchild&, lngfwdbut2&, "_AOL_Icon", vbNullString)
        Let lngfwdbut4& = FindWindowEx(lngchild&, lngfwdbut3&, "_AOL_Icon", vbNullString)
        Let lngfwdbut5& = FindWindowEx(lngchild&, lngfwdbut4&, "_AOL_Icon", vbNullString)
        Let lngfwdbut6& = FindWindowEx(lngchild&, lngfwdbut5&, "_AOL_Icon", vbNullString)
        Let lngfwdbut& = FindWindowEx(lngchild&, lngfwdbut6&, "_AOL_Icon", vbNullString)
    Loop Until lngchild& <> 0& And lngfwdbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngfwdbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngfwdbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(Index&))
        Let lngtowin& = FindWindowEx(lngsendwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngccwin& = FindWindowEx(lngsendwin&, lngtowin&, "_AOL_Edit", vbNullString)
        Let lngsubjectwin& = FindWindowEx(lngsendwin&, lngccwin&, "_AOL_Edit", vbNullString)
        Let lngrichwin& = FindWindowEx(lngsendwin&, 0&, "RICHCNTL", vbNullString)
        Let lngfontcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Fontcombo", vbNullString)
        Let lngcombo& = FindWindowEx(lngsendwin&, 0&, "_AOL_Combobox", vbNullString)
        Let lngsendbut1& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut2& = FindWindowEx(lngsendwin&, lngsendbut1&, "_AOL_Icon", vbNullString)
        Let lngsendbut3& = FindWindowEx(lngsendwin&, lngsendbut2&, "_AOL_Icon", vbNullString)
        Let lngsendbut4& = FindWindowEx(lngsendwin&, lngsendbut3&, "_AOL_Icon", vbNullString)
        Let lngsendbut5& = FindWindowEx(lngsendwin&, lngsendbut4&, "_AOL_Icon", vbNullString)
        Let lngsendbut6& = FindWindowEx(lngsendwin&, lngsendbut5&, "_AOL_Icon", vbNullString)
        Let lngsendbut7& = FindWindowEx(lngsendwin&, lngsendbut6&, "_AOL_Icon", vbNullString)
        Let lngsendbut8& = FindWindowEx(lngsendwin&, lngsendbut7&, "_AOL_Icon", vbNullString)
        Let lngsendbut9& = FindWindowEx(lngsendwin&, lngsendbut8&, "_AOL_Icon", vbNullString)
        Let lngsendbut10& = FindWindowEx(lngsendwin&, lngsendbut9&, "_AOL_Icon", vbNullString)
        Let lngsendbut11& = FindWindowEx(lngsendwin&, lngsendbut10&, "_AOL_Icon", vbNullString)
        Let lngsendbut12& = FindWindowEx(lngsendwin&, lngsendbut11&, "_AOL_Icon", vbNullString)
        Let lngsendbut13& = FindWindowEx(lngsendwin&, lngsendbut12&, "_AOL_Icon", vbNullString)
        Let lngsendbut14& = FindWindowEx(lngsendwin&, lngsendbut13&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut14&, "_AOL_Icon", vbNullString)
        pause 0.2
    Loop Until lngsendwin& <> 0& And lngtowin& <> 0& And lngccwin& <> 0& And lngsubjectwin& <> 0& And lngrichwin& <> 0& And lngsendbut& <> 0& And lngcombo& <> 0& And lngfontcombo& <> 0& And _
    lngsendbut1& <> lngsendbut2& And lngsendbut2& <> lngsendbut3& And lngsendbut3& <> lngsendbut4& And lngsendbut4& <> lngsendbut5& And lngsendbut5& <> lngsendbut6& And lngsendbut6& <> lngsendbut7& And _
    lngsendbut7& <> lngsendbut8& And lngsendbut8& <> lngsendbut9& And lngsendbut9& <> lngsendbut10& And lngsendbut10& <> lngsendbut11& And lngsendbut11& <> lngsendbut12& And lngsendbut12& <> lngsendbut13& And _
    lngsendbut13& <> lngsendbut14& And lngsendbut14& <> lngsendbut15& And lngsendbut15& <> lngsendbut&
    If deletefwdtext = True Then
        Let strnewsubject$ = GetText(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(Index&))
        Let lngsendbut& = FindWindowEx(lngsendwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngsendwin&, lngsendbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lngsendbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(lngsendbut&, WM_LBUTTONUP, 0&, 0&)
        Call pause(0.7)
    Loop Until lngsendwin& = 0& Or lngerrorwin& <> 0&
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(Index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getsentmailsubject(Index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(Index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub

Public Function waitforokorchatroom(Room As String)
    Dim strroomname As String, lngokwin As Long, lngbutton As Long
    Let Room$ = LCase$(replacestring(Room$, " ", ""))
    Do: DoEvents
        Let strroomname$ = GetCaption(findroom&)
        Let strroomname$ = LCase$(replacestring(strroomname$, " ", ""))
        Let lngokwin& = FindWindow("#32770", "america online")
        Let lngbutton& = FindWindowEx(lngokwin&, 0&, "button", "ok")
    Loop Until (lngokwin& <> 0& And lngbutton& <> 0&) Or Room$ = strroomname$
    DoEvents
    If lngokwin& <> 0& Then
        Do: DoEvents
            Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
            Let lngokwin& = FindWindow("#32770", "america online")
            Let lngbutton& = FindWindowEx(lngokwin&, 0&, "button", "ok")
        Loop Until lngokwin& = 0& And lngbutton& = 0&
    End If
End Function
























