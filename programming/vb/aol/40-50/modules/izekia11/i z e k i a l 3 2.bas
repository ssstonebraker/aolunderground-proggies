Attribute VB_Name = "izekial32"
'============================================================'
'          i     z    e    k    i    a    l    3   2         '
'              created in whole by izekial83                 '
'              created using visual basic 6                  '
'                     ~ aim:imizekial ~                      '
'               ~ mail:izekial83@hotmail.com ~               '
'            ~ http://www.come.to/izekial83/ ~               '
'             ~ 6-3-1999 ~ thru ~ 8-10-1999 ~                '
'  ~aol4~mirc5.41~aim2~winamp~photoshop5~aol2.5~windows98    '
'      testers: wolph~pentium~snow~xeno~beav~izekial         '
'============================================================'
'  Sup,  This is my new and maybe even last bas for aol4.
'  I spent a lot of time on this (over 2 months) and hope
'  that you enjoy it. I have personally  tested every sub
'  in  here at least  3 times, some  more. If you find an
'  error,  which is not  likely, ;0), send  the error and
'  sub  name  to  izekial83@hotmail.com.  I would  really
'  appreciate that. I have  provided an overview  section
'  for any help or confusion-subname: aa_overview

Option Explicit

Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lparam As Long) As Boolean
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetHostName Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function GetHostByName Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wflags As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

'Window Messages
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
Public Const WM_GETTEXTLENGTH = &HE
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



'Combo Box Functions
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146

'hWnd Functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

'Show Window Functions
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

'Sound Functions
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8

'Screen Saver Function
Public Const SPI_SCREENSAVERRUNNING = 97

'Get Window Word Functions
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

'Virtual Key Statements
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

'Phader Color Presets
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

'Processor Types
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

'Menu Functions
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_POPUP = &H10&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&

'Key Presets
Public Const ENTER_KEY = 13

'Button Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

'List Box Functions
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

'Notify Icon Functions
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

'Edit Window Messages
Public Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1

'Dev Mode Const's
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

'Windows Version Functions
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Byte Functions
Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12

'winsck functions
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

'types
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
        hWnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type COLORRGB
    red As Long
    green As Long
    blue As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

'Enums
Public Enum aolshortcutkeys
    ctrl1
    ctrl2
    ctrl3
    ctrl4
    ctrl5
    ctrl6
    ctrl7
    ctrl8
    ctrl9
    ctrl0
End Enum

Public chatline As String
Public Sub aa_overview()
    'this is a  small  overview, all  other questions should be sent
    'to izekial83@hotmail.com
    
    'prefixs : the  prefix str & lng  as used in this bas a  lot str
    '          stands for string, lng stands for long(a #)
    
    'colors  : as you may have seen there are many fade subs in here
    '          you  can  find  preset  colors  up  above ,  such  as
    '          color_red, color_blue, color_green, color_white etc..
    
    'dog bar : if you are wondering what all these subs that say with
    '          dogbar  at the end of them are,  dog bar is a progress
    '          bar  i use that is  really neat. it  lets you keep the
    '          progress of something,and its faded. get it at my site
    
    'autofade: for the autophader, you need spyworks 5 or 6
    
    'sub     : a  sub will  perform a  routine. it returns nothing ex.
    '          call sendchat("hey, this is a sub")
    
    'function: a  function will  perform a routine  and then return a
    '          value. ex. lngroomhandle& = findroom&
    
    'more information and help is provided in my tutorial,  available
    'at my site... http://www.come.to/izekial83/
End Sub
Public Sub aimsendblueblackchat(strtext As String)
    Dim lnglen As Long, strstring As String, strstring2 As String, index As Long, strtext2 As String
    Let lnglen& = Len(strtext$)
    For index& = 1& To lnglen& Step 2&
        Let strstring$ = Mid$(strtext$, index&, 1&)
        Let strstring2$ = Mid$(strtext$, index& + 1&, 1&)
        Let strtext2$ = strtext2$ & "<font color=" & Chr$(34&) & "#5ba5f9" & Chr$(34&) & ">" & strstring$ & "<font color=" & Chr$(34&) & "#000000" & Chr$(34&) & ">" & strstring2$
    Next index&
    Call aimsendchat(strtext2$)
End Sub
Public Sub aimsendbluegreenchat(strtext As String)
    Dim lnglen As Long, strstring As String, strstring2 As String, index As Long, strtext2 As String
    Let lnglen& = Len(strtext$)
    For index& = 1& To lnglen& Step 2&
        Let strstring$ = Mid$(strtext$, index&, 1&)
        Let strstring2$ = Mid$(strtext$, index& + 1&, 1&)
        Let strtext2$ = strtext2$ & "<font color=" & Chr$(34&) & "#5ba5f9" & Chr$(34&) & ">" & strstring$ & "<font color=" & Chr$(34&) & "#30e230" & Chr$(34&) & ">" & strstring2$
    Next index&
    Call aimsendchat(strtext2$)
End Sub
Public Sub aimsendpurplegreenchat(strtext As String)
    Dim lnglen As Long, strstring As String, strstring2 As String, index As Long, strtext2 As String
    Let lnglen& = Len(strtext$)
    For index& = 1& To lnglen& Step 2&
        Let strstring$ = Mid$(strtext$, index&, 1&)
        Let strstring2$ = Mid$(strtext$, index& + 1&, 1&)
        Let strtext2$ = strtext2$ & "<font color=" & Chr$(34&) & "#cf30e2" & Chr$(34&) & ">" & strstring$ & "<font color=" & Chr$(34&) & "#80e230" & Chr$(34&) & ">" & strstring2$
    Next index&
    Call aimsendchat(strtext2$)
End Sub
Public Sub aimsendmustardblackchat(strtext As String)
    Dim lnglen As Long, strstring As String, strstring2 As String, index As Long, strtext2 As String
    Let lnglen& = Len(strtext$)
    For index& = 1& To lnglen& Step 2&
        Let strstring$ = Mid$(strtext$, index&, 1&)
        Let strstring2$ = Mid$(strtext$, index& + 1&, 1&)
        Let strtext2$ = strtext2$ & "<font color=" & Chr$(34&) & "#88981a" & Chr$(34&) & ">" & strstring$ & "<font color=" & Chr$(34&) & "#000000" & Chr$(34&) & ">" & strstring2$
    Next index&
    Call aimsendchat(strtext2$)
End Sub
Public Sub aimsignon(stryoursn As String, stryourpassword As String)
    Dim isopened As Boolean, lngcombo As Long, lngaimwin As Long, lngedit As Long, lngstatic As Long
    On Error Resume Next
    If aimisonline Then Exit Sub
    If fileexists("c:\program files\aim95\aim.exe") = True Then
        Call Shell("c:\program files\aim95\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\program files\netscape\communicator\program\aim\aim.exe") = True Then
        Call Shell("c:\program files\netscape\communicator\program\aim\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\aim\aim.exe") Then
        Call Shell("c:\aim\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\aim95\aim.exe") Then
        Call Shell("c:\aim95\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\aim2\aim.exe") Then
        Call Shell("c:\aim2\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\program files\aim2\aim.exe") Then
        Call Shell("c:\program files\aim2\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\online\aim\aim.exe") Then
        Call Shell("c:\online\aim\aim.exe", vbNormalFocus)
            ElseIf fileexists("c:\program files\aim98\aim.exe") Then
        Call Shell("c:\program files\aim98\aim.exe", vbNormalFocus)
    Else
        Exit Sub
    End If
    'i just realized, selectcase would have been a little bit easier there, damn.
    Do: DoEvents
        Let lngaimwin& = FindWindow("#32770", "sign on")
    Loop Until lngaimwin& <> 0&
    Let lngcombo& = FindWindowEx(lngaimwin&, 0&, "combobox", vbNullString)
    Let lngedit& = FindWindowEx(lngcombo&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, stryoursn$): DoEvents
    Let lngedit& = FindWindowEx(lngaimwin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, stryourpassword$)
    Call PostMessage(lngedit&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngedit&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Function getlinefromstring(strstring As String, lngline As Long) As String
    Dim strline As String, lngcount As Long, lngspot1 As Long, lngspot2 As Long, index As Long
    Let lngcount& = getstringlinecount(strstring$)
    If lngline& > lngcount& Then Exit Function
    If lngline& = 1& And lngcount& = 1& Then Let getlinefromstring$ = strstring$:  Exit Function
    If lngline& = 1& And lngcount& <> 1& Then
        Let strline$ = Left$(strstring$, InStr(strstring$, Chr$(13&)) - 1&)
        Let strline$ = replacestring(strline$, Chr$(13&), "")
        Let strline$ = replacestring(strline$, Chr$(10&), "")
        Let getlinefromstring$ = strline$
        Exit Function
    Else
        Let lngspot1& = InStr(strstring$, Chr$(13&))
        For index& = 1& To lngline& - 1&
            Let lngspot2& = lngspot1&
            Let lngspot1& = InStr(lngspot1& + 1&, strstring$, Chr$(13&))
        Next index
        If lngspot1& = 0& Then Let lngspot1& = Len(strstring$)
        If (lngspot1& - lngspot2&) + 1& <= Len(strstring$) Then
            If lngspot2& = 0& Then Let lngspot2& = lngspot2& + 1&
            Let strline$ = Mid$(strstring$, lngspot2&, (lngspot1& - lngspot2&) + 1&)
        End If
        Let strline$ = replacestring(strline$, Chr$(13&), "")
        Let strline$ = replacestring(strline$, Chr$(10&), "")
        Let getlinefromstring$ = strline$
    End If
End Function
Public Sub aimaddbuddylisttolist(thelist As ListBox)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngcount As Long
    Dim index As Long, strbuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, lnglbtext As String
    thelist.Clear
    If lngtree = 0& Then Exit Sub
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngtree& = FindWindowEx(lnggroupwin&, 0&, "_oscar_tree", vbNullString)
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then If InStr(strbuffer$, "/") <> 0& Then thelist.AddItem strbuffer$
    Next index&
End Sub
Public Sub addtreetocontrol(lngtree As Long, thelist As Control)
    Dim index As Long, strbuffer As String, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    thelist.Clear
    If lngtree& = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        thelist.AddItem strbuffer$
    Next index&
End Sub
Public Sub addtreetocontrolwithdogbar(lngtree As Long, thelist As Control, progbar As Control)
    Dim index As Long, strbuffer As String, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    thelist.Clear
    If lngtree& = 0& Then Exit Sub
    Let progbar.Max = SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let progbar.Value = index& - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        thelist.AddItem strbuffer$
    Next index&
End Sub
Public Function addtreetostring(lngtree As Long, strseperator As String) As String
    Dim strstring As String, index As Long, strbuffer As String, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    If lngtree& = 0& Then Exit Function
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        Let strstring$ = strstring$ & strbuffer$ & strseperator$
    Next index&
    Let addtreetostring$ = Left$(strstring$, Len(strstring$) - 1&)
End Function
Public Function addtreetoclipboard(lngtree As Long, strseperator As String) As String
    Dim strstring As String, index As Long, strbuffer As String, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    If lngtree& = 0& Then Exit Function
    Clipboard.Clear
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        Let strstring$ = strstring$ & strbuffer$ & strseperator$
    Next index&
    Clipboard.SetText Left$(strstring$, Len(strstring$) - 1&)
End Function
Public Sub chatignorebyindex(lngindex As Long)
    Dim lngroom As Long, lnglist As Long, lnginfo As Long, strwho As String
    Dim lngcheck As Long, checkval As Long, lngcount As Long
    Let lngcount& = roomcount&
    If lngindex& > lngcount& - 1& Then
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
    Dim lngcheck As Long, checkval As Long, lngcount As Long
    Let lngcount& = roomcount&
    If lngindex& > lngcount& - 1& Then
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
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim strwho As String, lngcheck As Long, lnginfo As Long, checkval As Long
    Dim lIndex As Long
    On Error Resume Next
    Let room& = findroom&
    If room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For index& = 0& To SendMessage(rlist, LB_GETCOUNT, 0&, 0&) - 1&
                Let screenname$ = String$(4&, vbNullChar)
                Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24&
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes)
                Call CopyMemory(psnHold&, ByVal screenname$, 4&)
                Let psnHold& = psnHold& + 6&
                Let screenname$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
                If screenname$ <> getuser$ And LCase$(screenname$) = LCase$(strname$) Then
                    Call SendMessage(rlist&, LB_SETCURSEL, index&, 0&)
                    Let strwho$ = getlistitemtext(rlist&, index&)
                    Call SendMessageLong(rlist&, WM_LBUTTONDBLCLK, 0&, 0&)
                    Do: DoEvents
                        Let lnginfo& = FindWindowEx(GetParent(room&), 0&, "aol child", strwho$)
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
            Next index&
            Call CloseHandle(mthread)
        End If
    End If
End Sub

Public Sub chatunignorebyname(strname As String)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim strwho As String, lngcheck As Long, lnginfo As Long, checkval As Long
    On Error Resume Next
    Let room& = findroom&
    If room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For index& = 0& To SendMessage(rlist, LB_GETCOUNT, 0&, 0&) - 1&
                Let screenname$ = String$(4&, vbNullChar)
                Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24&
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes)
                Call CopyMemory(psnHold&, ByVal screenname$, 4&)
                Let psnHold& = psnHold& + 6&
                Let screenname$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
                If screenname$ <> getuser$ And LCase$(screenname$) = LCase$(strname$) Then
                    Call SendMessage(rlist&, LB_SETCURSEL, index&, 0&)
                    Let strwho$ = getlistitemtext(rlist&, index&)
                    Call SendMessageLong(rlist&, WM_LBUTTONDBLCLK, 0&, 0&)
                    Do: DoEvents
                        Let lnginfo& = FindWindowEx(GetParent(room&), 0&, "aol child", strwho$)
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
            Next index&
            Call CloseHandle(mthread)
        End If
    End If
End Sub
Public Function aimcountonlinebuddys() As Long
    Dim lngbuddywin As Long, lnggroupwin As Long, lngtree As Long
    Dim index As Long, strbuffer As String, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String, templong As Long
    templong& = 0&
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngtree& = FindWindowEx(lnggroupwin&, 0&, "_oscar_tree", vbNullString)
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then If InStr(strbuffer$, "/") <> 0& Then templong& = templong& + 1&
    Next index&
    aimcountonlinebuddys& = templong&
End Function
Public Sub aimaddbuddylisttocombo(thecombo As ComboBox)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngcount As Long
    Dim index As Long, strbuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, lnglbtext As String
    thecombo.Clear
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngtree& = FindWindowEx(lnggroupwin&, 0&, "_oscar_tree", vbNullString)
    If lngtree = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then If InStr(strbuffer$, "/") <> 0& Then thecombo.AddItem strbuffer$
    Next index&
End Sub
Public Function roomcount() As Long
    Dim lngroom As Long, lnglist As Long, lngcount As Long
    Let lngroom& = findroom&
    If lngroom& = 0& Then
        Exit Function
    Else
        Let lnglist& = FindWindowEx(lngroom&, 0&, "_AOL_Listbox", vbNullString)
        Let lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
        Let roomcount& = lngcount&
    End If
End Function
Public Sub aimaddbuddygroupstolist(thelist As ListBox)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngcount As Long
    Dim index As Long, strbuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, lnglbtext As String
    thelist.Clear
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngtree& = FindWindowEx(lnggroupwin&, 0&, "_oscar_tree", vbNullString)
    If lngtree = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then If InStr(strbuffer$, "/") > 0& And InStr(strbuffer$, "%") <= 0& Then thelist.AddItem strbuffer$
    Next index&
End Sub
Public Sub aimaddbuddygroupstocombo(thecombo As ComboBox)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngcount As Long
    Dim index As Long, strbuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, lnglbtext As String
    thecombo.Clear
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngtree& = FindWindowEx(lnggroupwin&, 0&, "_oscar_tree", vbNullString)
    If lngtree = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then If InStr(strbuffer$, "/") > 0& And InStr(strbuffer$, "%") <= 0& Then thecombo.AddItem strbuffer$
    Next index&
End Sub
Public Sub aimaddchatroomtolist(thelist As ListBox)
    Dim lngchat As Long, lngcount As Long, index As Long
    Dim strbuffer As String, lngtree As Long, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    thelist.Clear
    Let lngchat& = FindWindow("aim_chatwnd", vbNullString)
    Let lngtree& = FindWindowEx(lngchat&, 0, "_oscar_tree", vbNullString)
    If lngtree = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then thelist.AddItem (strbuffer$)
    Next index&
End Sub
Public Sub aimaddchatroomtocombo(thecombo As ComboBox)
    Dim lngchat As Long, lngcount As Long, index As Long
    Dim strbuffer As String, lngtree As Long, lngitemdata As Long
    Dim lngtextlen As Long, lnglbtext As String
    thecombo.Clear
    Let lngchat& = FindWindow("aim_chatwnd", vbNullString)
    Let lngtree& = FindWindowEx(lngchat&, 0, "_oscar_tree", vbNullString)
    If lngtree = 0& Then Exit Sub
    For index& = 0& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let lnglbtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then thecombo.AddItem (strbuffer$)
    Next index&
End Sub
Public Function getstringlinecount(strstring As String) As Long
    Dim enterpos As Long
    If Len(strstring$) = 0& Then
        Let getstringlinecount& = 0&
    Else
        Let enterpos& = InStr(strstring$, Chr$(13&))
        If enterpos& <> 0& Then
            Let getstringlinecount& = 1&
            Do While enterpos& <> 0&
                Let enterpos& = InStr(enterpos& + 1&, strstring$, Chr$(13&))
                If enterpos& <> 0& Then
                    Let getstringlinecount& = getstringlinecount& + 1&
                End If
            Loop: DoEvents
        End If
        Let getstringlinecount& = getstringlinecount& + 1&
    End If
End Function
Public Function deletelinefromstring(strstring As String, lngline As Long) As String
    Dim lnglinecount As Long, index As Long, strtext As String
    Let lnglinecount& = getstringlinecount(strstring)
    For index& = 1& To lnglinecount&
        If index& <> lngline& Then
            strtext$ = strtext$ & getlinefromstring(strstring, index&) & vbCrLf$
        End If
    Next index&
    deletelinefromstring$ = strtext$
End Function
Public Sub aimadshow()
    Dim lngbuddywin As Long, atewin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let atewin& = FindWindowEx(lngbuddywin&, 0&, "wndate32class", vbNullString)
    Call ShowWindow(atewin&, SW_SHOW)
End Sub
Public Sub aimadhide()
    Dim lngbuddywin As Long, atewin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let atewin& = FindWindowEx(lngbuddywin&, 0&, "wndate32class", vbNullString)
    Call ShowWindow(atewin&, SW_HIDE)
End Sub
Public Function aimcountopenims() As Long
    Dim lngimwin As Long
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& <> 0& Then
        Let aimcountopenims& = 1&
        Do: DoEvents
            Let lngimwin& = FindWindowEx(0&, lngimwin&, "aim_imessage", vbNullString)
            If lngimwin& <> 0& Then Let aimcountopenims& = aimcountopenims& + 1&
        Loop Until lngimwin& = 0&
    Else
        Exit Function
    End If
End Function
Public Sub aimsendim(strsn As String, strmessage As String)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngbutton As Long, lngimwin As Long
    Dim lngcombowin As Long, lngeditwin As Long, lngatewin&, lngicon As Long
    Dim lngerrorwin As Long, tempstring As String, tempstring2 As String
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& <> 0 Then
        Let tempstring$ = getcaption(lngimwin&)
        Let tempstring$ = LCase$(replacestring(tempstring$, " ", ""))
        Let tempstring2$ = LCase$(replacestring(strsn$, " ", "") & "-instantmessage")
        If tempstring$ Like tempstring2$ Then
            Let lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
            Let lngatewin& = FindWindowEx(lngimwin&, lngatewin&, "wndate32class", vbNullString)
            Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
            DoEvents
            Let lngicon& = FindWindowEx(lngimwin&, 0&, "_oscar_iconbtn", vbNullString)
            Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
        End If
    End If
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngbutton& = FindWindowEx(lnggroupwin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    Loop Until lngimwin& <> 0&
    Let lngcombowin& = FindWindowEx(lngimwin&, 0&, "_oscar_persistantcombo", vbNullString)
    Let lngeditwin& = FindWindowEx(lngcombowin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strsn$)
    DoEvents
    Let lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
    Let lngatewin& = FindWindowEx(lngimwin&, lngatewin&, "wndate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
    DoEvents
    Let lngicon& = FindWindowEx(lngimwin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function aol25findwelcomewin() As Long
    Dim aol As Long, mdi As Long, child As Long, childtxt As String
    Let aol& = FindWindow("AOL Frame25", "America  Online")
    Let mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Let child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Let childtxt$ = getcaption(child&)
    If InStr(childtxt$, "Welcome, ") <> 0& Then
        Let aol25findwelcomewin& = child&
        Exit Function
    Else
        Do: DoEvents
            Let child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            Let childtxt$ = getcaption(child&)
            If InStr(childtxt$, "Welcome, ") <> 0& Then
                Let aol25findwelcomewin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
End Function
Public Sub aol25keyword(strkeyword As String)
    Dim aol As Long, mdi As Long, child As Long
    Dim edit As Long, Icon As Long
    Call runaolmenubystring("Keyword...")
    Let aol& = FindWindow("AOL Frame25", vbNullString)
    Let mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Let child& = FindWindowEx(mdi&, 0&, "AOL Child", "Keyword")
    Do: DoEvents: Let child& = FindWindowEx(mdi&, 0&, "AOL Child", "Keyword"): Loop Until child& <> 0&
    Do: DoEvents: Let edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString): Loop Until edit& <> 0&
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, strkeyword$)
    Let Icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    Do: DoEvents
        Let child& = FindWindowEx(mdi&, 0&, "AOL Child", "Keyword")
        Call SendMessageLong(Icon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(Icon&, WM_LBUTTONUP, 0&, 0&)
    Loop Until child& = 0&
End Sub
Public Function aol25findroom() As Long
    Dim aol As Long, mdi As Long, child As Long
    Dim edit As Long, Icon As Long, list As Long
    Let aol& = FindWindow("AOL Frame25", vbNullString)
    Let mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Let child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Let edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
    Let list& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    Let Icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    If edit& <> 0& And list& <> 0& And Icon& <> 0& Then
        Let aol25findroom& = child&
    Else
        Do: DoEvents
            Let child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            Let edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
            Let list& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            Let Icon& = FindWindowEx(child&, 0&, "_AOL_Child", vbNullString)
            If edit& <> 0& And list& <> 0& And Icon& <> 0& Then
                Let aol25findroom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
End Function
Public Function aol25getchattext() As String
    Dim room As Long, edit As Long, chattext As String
    Let room& = findroom&
    Let edit& = FindWindowEx(room&, 0&, "_AOL_View", vbNullString)
    Let chattext$ = gettext(edit&)
    If chattext$ = "" Then Exit Function
    Let chattext$ = striphtml$(chattext$)
    Let aol25getchattext$ = chattext$
End Function
Public Function aol25getchatlinemsg() As String
    Dim lastline As String, thetab As Long
    Let lastline$ = aol25getchatline$
    Let thetab& = InStr(lastline$, Chr$(9&))
    Let aol25getchatlinemsg$ = Right$(lastline$, Len(lastline$) - thetab&)
End Function
Public Function aol25getchatlinesn() As String
    Dim lastline As String, thetab As Long
    Let lastline$ = aol25getchatline$
    Let thetab& = InStr(lastline$, Chr$(9&))
    Let aol25getchatlinesn$ = Left$(lastline$, thetab& - 3)
End Function
Public Function aol25getchatline() As String
    Dim Chat As Long, edit As Long, chattext As String, enter As Long, enter2 As Long
    Let chattext$ = aol25getchattext$
    Let enter& = InStr(chattext$, Chr$(13&))
    Do: DoEvents
        If enter& <> 0& Then enter2& = enter&
        Let enter& = InStr(enter& + 1&, chattext$, Chr$(13&))
    Loop Until enter& = 0&
    Let aol25getchatline$ = Right(chattext$, Len(chattext$) - enter2&)
End Function
Public Sub aol25privateroom(strroom As String)
    Call aol25keyword("aol://2719:2-2-" & strroom$)
End Sub
Public Sub aol25memberroom(strroom As String)
    Call aol25keyword("aol://2719:61-2-" & strroom$)
End Sub
Public Sub aol25sendim(strname As String, strmessage As String)
    Dim aol As Long, imwin As Long, Button As Long, edit As Long
    Dim okwin As Long, Button2 As Long
    Call aol25keyword("aol://9293:" & strname$)
    Do: DoEvents: Let imwin& = FindWindowEx(FindWindowEx(FindWindow("AOL Frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", "Send Instant Message"): Loop Until imwin& <> 0&
    Let edit& = FindWindowEx(imwin&, 0, "_AOL_Edit", vbNullString)
    Let edit& = FindWindowEx(imwin&, edit&, "_AOL_Edit", vbNullString)
    Do: DoEvents: Loop Until edit& <> 0&
    Do: DoEvents: Call SendMessageByString(edit&, WM_SETTEXT, 0&, strmessage$): Loop Until gettext(edit&) = strmessage$
    Let Button& = FindWindowEx(imwin&, 0&, "_AOL_Button", vbNullString)
    Call SendMessageLong(Button&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(Button&, WM_LBUTTONUP, 0&, 0&)
    Let aol& = FindWindow("AOL Frame25", "America  Online")
    Let okwin& = FindWindowEx(aol&, 0&, "#32770", vbNullString)
    Let Button2& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Do: DoEvents: Loop Until okwin& <> 0& Or imwin& = 0&
    If okwin& <> 0& Then
        Let Button2& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Call SendMessageLong(Button2&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(Button2&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessageLong(imwin&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Function striphtml(thestring As String, Optional returns As Boolean) As String
    'from juggalo32.bas
    Dim roomtext As String, takeout As String, ReplaceWith As String
    Dim whereat As Long, lefttext As String, righttext As String
    Dim takeout2 As String, whereat1 As Long, whereat2 As Long
    Dim takeout3 As String, takeout4 As String
    roomtext$ = thestring$
    If returns = True Then
        takeout$ = "<br>"
        takeout2$ = "<Br>"
        takeout3$ = "<bR>"
        takeout4$ = "<BR>"
        ReplaceWith$ = Chr(13) & Chr(10)
        whereat& = 0&
        Do: DoEvents
            whereat& = InStr(whereat& + 1, roomtext$, takeout$)
            If whereat& = 0& Then
                whereat& = InStr(whereat& + 1, roomtext$, takeout2$)
                If whereat& = 0& Then
                    whereat& = InStr(whereat& + 1, roomtext$, takeout3$)
                    If whereat& = 0& Then
                        whereat& = InStr(whereat& + 1, roomtext$, takeout4$)
                        If whereat& = 0& Then
                            Exit Do
                        End If
                    End If
                End If
            End If
            lefttext$ = Left(roomtext$, whereat& - 1)
            righttext$ = Mid(roomtext$, whereat& + 4, Len(roomtext$))
            roomtext$ = lefttext$ & ReplaceWith$ & righttext$
        Loop
    End If
    takeout$ = "<"
    takeout2$ = ">"
    whereat& = 0&
    whereat1& = 0&
    whereat2& = 0&
    Do: DoEvents
        whereat1& = InStr(whereat1& + 1, roomtext$, takeout$)
        If whereat1& = 0& Then Exit Do
        whereat2& = InStr(whereat2& + 1, roomtext$, takeout2$)
        whereat& = whereat2& - whereat1&
        lefttext$ = Left(roomtext$, whereat1& - 1)
        righttext$ = Mid(roomtext$, whereat2& + 1, Len(roomtext$) - whereat& + 1)
        roomtext$ = lefttext$ & righttext$
        whereat& = 0&
        whereat1& = 0&
        whereat2& = 0&
    Loop
    striphtml$ = Left(roomtext$, Len(roomtext$) - 2)
End Function
Public Sub aol25sendmail(strname As String, strsubject As String, strmessage As String)
    Dim aol As Long, Toolbar As Long, Icon As Long, mdi As Long, child As Long
    Dim edit As Long, Edit2 As Long, edit3 As Long
    Let aol& = FindWindow("AOL Frame25", "America  Online")
    Let Toolbar& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Let Icon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    Let Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
    Call SendMessageLong(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(Icon&, WM_LBUTTONUP, 0&, 0&)
    Let mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        Let child& = FindWindowEx(mdi&, 0&, "AOL Child", "Compose Mail")
        Let edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
        Let Edit2& = FindWindowEx(child&, edit&, "_AOL_Edit", vbNullString)
        Let Edit2& = FindWindowEx(child&, Edit2&, "_AOL_Edit", vbNullString)
        Let edit3& = FindWindowEx(child&, Edit2&, "_AOL_Edit", vbNullString)
        Let Icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    Loop Until child& <> 0& And edit& <> 0& And Edit2& <> 0& And edit3& <> 0& And Icon& <> 0&
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, strname$)
    Call SendMessageByString(Edit2&, WM_SETTEXT, 0&, strsubject$)
    Call SendMessageByString(edit3&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageLong(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(Icon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents: Let child& = FindWindowEx(mdi&, 0&, "AOL Child", "Compose Mail"): Loop Until child& = 0&
End Sub
Public Sub aimchatstamps()
    Dim lngchatwin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Call runanymenubystring(lngchatwin&, "timestamp")
End Sub
Public Sub aimsendmassim(thepeople As ListBox, strmessage As String)
    Dim index As Long
    If thepeople.ListCount = 0& Then Exit Sub
    For index& = 0& To thepeople.ListCount - 1&
        Call aimsendim(thepeople.list(index&), strmessage$)
        Call pause(3)
    Next index&
End Sub
Public Sub aimsendchat(strmessage As String)
    Dim lngchatwin As Long, lngatewin As Long, lngicon As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngatewin& = FindWindowEx(lngchatwin&, 0&, "wndate32class", vbNullString)
    Let lngatewin& = FindWindowEx(lngchatwin&, lngatewin&, "wndate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0, strmessage$)
    Let lngicon& = FindWindowEx(lngchatwin&, 0&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function aimroomcount() As Long
    Dim lngchatwin As Long, lngtree As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngtree& = FindWindowEx(lngchatwin&, 0, "_oscar_tree", vbNullString)
    Let aimroomcount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function aimcountopenrooms() As Long
    Dim lngchatwin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then aimcountopenrooms& = 1&
    Do: DoEvents
        Let lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
        If lngchatwin& <> 0& Then Let aimcountopenrooms& = aimcountopenrooms& + 1&
    Loop Until lngchatwin& = 0&
End Function
Public Sub aimsendchattoallrooms(message As String)
    Dim lngchatwin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then
        Do: DoEvents
            lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
            If lngchatwin& <> 0& Then Call aimsendchat(message$)
        Loop Until lngchatwin& = 0&
    End If
End Sub
Public Sub aimcloseim(strwho As String)
    Dim lngimwin As Long, tempstring As String, tempstring2 As String
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& <> 0& Then
        Let tempstring$ = getcaption(lngimwin&)
        Let tempstring$ = LCase(replacestring(tempstring$, " ", ""))
        Let tempstring2$ = LCase(replacestring(strwho$, " ", "") & "-instantmessage")
        If tempstring$ Like tempstring2$ Then
            Call windowclose(lngimwin&)
        Else
            Do: DoEvents
                Let lngimwin& = FindWindowEx(0&, lngimwin&, "aim_imessage", vbNullString)
                Let tempstring$ = getcaption(lngimwin&)
                Let tempstring$ = replacestring(tempstring$, " ", "")
                Let tempstring2$ = replacestring(strwho$, " ", "") & "-instantmessage"
                If tempstring$ Like tempstring2$ Then
                    Call windowclose(lngimwin&)
                End If
            Loop Until lngimwin& = 0 Or tempstring$ Like tempstring2$
        End If
    End If
End Sub
Public Sub aimopennewuserwizard()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "&new user wizard")
End Sub
Public Sub aimsavebuddylist()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "&save buddy list...")
End Sub
Public Sub aimloadbuddylist()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "&load buddy list...")
End Sub
Public Sub aimclosechat(strchatname As String)
    Dim lngchatwin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then
        If LCase$(getcaption(lngchatwin&)) Like LCase$(strchatname$) Then
            Call windowclose(lngchatwin&)
        Else
            Do: DoEvents
                Let lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
                If LCase$(getcaption(lngchatwin&)) Like LCase$(strchatname$) Then
                    Call windowclose(lngchatwin&)
                End If
            Loop Until LCase$(getcaption(lngchatwin&)) Like LCase$(strchatname$) Or lngchatwin& = 0&
        End If
    End If
End Sub
Public Sub aimchatignorebyname(strwho$)
    Dim lngchatwin As Long, lngicon&, tempstring As String, tempstring2 As String
    Dim index As Long, strbuffer As String, lngtree As Long
    Dim lngitemdata As Long, lngtextlen As Long, strtext As String
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngtree& = FindWindowEx(lngchatwin&, 0&, "_oscar_tree", vbNullString)
    If lngtree& = 0& Then Exit Sub
    For index& = 1& To SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessageLong(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strbuffer$ = String(lngtextlen&, 0&)
        Let strtext$ = SendMessageByString(lngtree&, LB_GETTEXT, index&, ByVal strbuffer$)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Let lngitemdata& = SendMessage(lngtree&, LB_GETITEMDATA, index&, 0&)
        If Not index& Then
            Let tempstring$ = replacestring(strbuffer$, " ", "")
            Let tempstring2$ = replacestring(strwho$, " ", "")
            If LCase$(tempstring$) = LCase$(tempstring2$) Then
                Let lngicon& = FindWindowEx(lngchatwin&, 0&, "_oscar_iconbtn", vbNullString)
                Let lngicon& = FindWindowEx(lngchatwin&, lngicon&, "_oscar_iconbtn", vbNullString)
                Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
            End If
        End If
    Next index&
End Sub
Public Sub aimchatignorebyindex(lngindex As Long)
    Dim lngchat As Long, lngtree As Long, lngicon As Long
    Let lngchat& = FindWindow("aim_chatwnd", vbNullString)
    Let lngtree& = FindWindowEx(lngchat&, 0&, "_oscar_tree", vbNullString)
    If lngchat& = 0& Then Exit Sub
    If lngindex& > SendMessageLong(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
    Call SendMessage(lngtree&, LB_SETCURSEL, lngindex&, 0&)
    Let lngicon& = FindWindowEx(lngchat&, 0&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lngchat&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aimopenim()
    Dim lngbuddy As Long, lnggroup As Long, lngicon As Long, lngimwin As Long
    Let lngbuddy& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroup& = FindWindowEx(lngbuddy&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngicon& = FindWindowEx(lnggroup&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    Loop Until lngimwin& <> 0&
End Sub
Public Function aimfindroom(strname As String) As Long
    Dim lngchatwin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    If lngchatwin& <> 0& Then
        If LCase$(getcaption(lngchatwin&)) Like LCase$(strname$) Then
            Let aimfindroom = lngchatwin&
        Else
            Do: DoEvents
                Let lngchatwin& = FindWindowEx(0&, lngchatwin&, "aim_chatwnd", vbNullString)
                If LCase$(getcaption(lngchatwin&)) Like LCase$(strname$) Then
                    Let aimfindroom = lngchatwin&
                End If
            Loop Until LCase$(getcaption(lngchatwin&)) Like LCase$(strname$) Or lngchatwin& = 0&
        End If
    End If
End Function
Public Function aimisonline() As Boolean
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    If lngbuddywin& <> 0& Then Let aimisonline = True: Exit Function
    Let aimisonline = False
End Function
Public Sub aimclearchat()
    Dim lngchatwin As Long, lngatewin As Long
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngatewin& = FindWindowEx(lngchatwin&, 0&, "wndate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(lngatewin&, WM_CLEAR, 0&, 0&)
End Sub
Public Sub aimsendchatlink(strurl As String, strmessage As String)
    Call aimsendchat("< a href=" & Chr$(34&) & strurl$ & Chr$(34&) & ">" & strmessage$ & "</a>")
End Sub
Public Sub aimcreateprofile(yourinfo As String)
    Dim lngbuddywin As Long, lngprofwin As Long, lngbutton As Long
    Dim lngprofchild As Long, lngatewin As Long, lngatewin2 As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "my &profile")
    Do While lngprofwin = 0&
        Let lngprofwin& = FindWindow("#32770", vbNullString)
    Loop: DoEvents
    Let lngbutton& = FindWindowEx(lngprofwin&, 0&, "button", vbNullString)
    Let lngbutton& = FindWindowEx(lngprofwin&, lngbutton&, "button", vbNullString)
    Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
    Let lngprofchild& = FindWindowEx(lngprofwin&, 0, "#32770", vbNullString)
    Let lngatewin& = FindWindowEx(lngprofchild&, 0, "wndate32class", vbNullString)
    Do: DoEvents
        Let lngatewin2& = FindWindowEx(lngatewin&, 0, "ate32class", vbNullString)
    Loop Until lngatewin2& <> 0&
    Call SendMessage(lngbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(lngbutton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub aimopenbuddywizard()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "find a buddy &wizard")
End Sub
Public Function aimisbuddyavailable() As Boolean
    Dim lngbuddywin As Long, lnglocatewin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call runanymenubystring(lngbuddywin&, "get member inf&o")
    Do: DoEvents
        Let lnglocatewin& = FindWindow("_oscar_locate", vbNullString)
    Loop Until lnglocatewin& <> 0&
End Function
Public Function aimusersn() As String
    Dim lngbuddywin As Long, thecap As String
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let thecap$ = getcaption(lngbuddywin&)
    If Not Right(thecap$, 20&) = "'s Buddy List Window" Then
        Let aimusersn$ = ""
        Exit Function
    Else
        Let aimusersn$ = Left$(thecap$, Len(thecap$) - 20&)
    End If
End Function
Public Sub aimsendchatinvite(strwho As String, strmessage As String, strchatname As String)
    Dim lngbuddywin As Long, lnggroupwin As Long, lngicon As Long, lnginvitewin As Long
    Dim lngeditwin As Long, lngchatwin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngicon& = FindWindowEx(lnggroupwin&, 0&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lnggroupwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do While lnginvitewin = 0& And lngeditwin& = 0& And lngicon& = 0&
        Let lnginvitewin& = FindWindow("aim_chatinvitesendwnd", vbNullString)
        Let lngeditwin& = FindWindowEx(lnginvitewin&, 0&, "edit", vbNullString)
        Let lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
        Let lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
        Let lngicon& = FindWindowEx(lnginvitewin&, 0&, "_oscar_iconbtn", vbNullString)
        Let lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
        Let lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
        DoEvents
    Loop
    Let lngeditwin& = FindWindowEx(lnginvitewin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strwho$)
    DoEvents
    Let lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strmessage$)
    DoEvents
    Let lngeditwin& = FindWindowEx(lnginvitewin&, lngeditwin&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, strchatname$)
    DoEvents
    Let lngicon& = FindWindowEx(lnginvitewin&, 0&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lnginvitewin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Loop Until lngchatwin& <> 0&
    If FindWindow("#32770", "Buddy Chat Invitation Error") Then Call windowclose(FindWindow("#32770", "Buddy Chat Invitation Error"))
End Sub
Public Sub aimenterroom(strroom As String)
    Dim lngchatwin As Long, strstring As String
    Call aimsendchatinvite(aimusersn$, aimusersn$, strroom$)
    Do: DoEvents
        Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Loop Until lngchatwin& <> 0
    Do: DoEvents
        Let strstring$ = getcaption(lngchatwin&)
    Loop Until Right$(strstring$, 5&) = Right$(strroom$, 5&)
End Sub
Public Sub aimopenchatinvite()
    Dim lngbuddywin As Long, lnggroupwin As Long, lngicon As Long, lnginvitewin As Long
    Dim lngeditwin As Long, lngchatwin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lnggroupwin& = FindWindowEx(lngbuddywin&, 0&, "_oscar_tabgroup", vbNullString)
    Let lngicon& = FindWindowEx(lnggroupwin&, 0&, "_oscar_iconbtn", vbNullString)
    Let lngicon& = FindWindowEx(lnggroupwin&, lngicon&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do While lnginvitewin = 0&
        Let lnginvitewin& = FindWindow("aim_chatinvitesendwnd", vbNullString)
    Loop: DoEvents
End Sub
Public Sub aimsignoff()
    Call runanymenubystring(FindWindow("_oscar_buddylistwin", vbNullString), "sign o&ff")
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
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
    Let lngmmenu& = GetMenu(lngwindow&)
    Let lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
        Let lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
        Let lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
            Let lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
            Let strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If LCase$(strstring$) = replacestring(LCase$(strmenutext$), Chr$(0&), "") Then
                Call SendMessageLong(lngwindow&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub aimexit()
    Call runanymenubystring(FindWindow("_oscar_buddylistwin", vbNullString), "e&xit")
End Sub
Public Function aimsnfromim() As String
    Dim lngimwin As Long, strimcap As String, lngtextlen As Long, strsn As String
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    If lngimwin& = 0& Then
        Exit Function
    Else
        Let strimcap$ = getcaption(lngimwin&)
        If InStr(strimcap$, "- instant message") <> 0& Then
            Let lngtextlen& = GetWindowTextLength(lngimwin&) - 19&
            Let strsn$ = Left$(strimcap$, InStr(strimcap$, "" & strimcap$ & "") + lngtextlen&)
            Let aimsnfromim$ = strsn$
        Else
            Let aimsnfromim$ = ""
        End If
    End If
End Function
Public Function aimgetwhatusersaid() As String
    'this will retrieve the message in the edit box of an im
    Dim lngimwin As Long, lngatewin As Long, strimtext As String
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    Let lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
    Let lngatewin& = FindWindowEx(lngimwin&, lngatewin&, "wndate32class", vbNullString)
    Let strimtext$ = striphtml(gettext(lngatewin&))
    Let aimgetwhatusersaid$ = strimtext$
End Function
Public Function aimgetlastimmsg() As String
    'this will retrieve the message in the sent box of an im
    Dim lngimwin As Long, lngatewin As Long, strimtext As String, lnglong As Long
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    Let lngatewin& = FindWindowEx(lngimwin&, 0&, "wndate32class", vbNullString)
    Let strimtext$ = gettext(lngatewin&)
    Let lnglong& = InStrRev(strimtext$, "<br>")
    Let strimtext$ = Right$(strimtext$, Len(strimtext$) - lnglong& - 3&)
    Let strimtext$ = striphtml(strimtext$)
    Let aimgetlastimmsg$ = strimtext$
End Function
Public Sub aimhide()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call ShowWindow(lngbuddywin&, SW_HIDE)
End Sub
Public Sub aimshow()
    Dim lngbuddywin As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Call ShowWindow(lngbuddywin&, SW_SHOW)
End Sub
Public Function aimchatline() As String
    Dim lngchatwin As Long, lngatewin As Long
    Dim lnglong As Long, strstring As String
    Let lngchatwin& = FindWindow("aim_chatwnd", vbNullString)
    Let lngatewin& = FindWindowEx(lngchatwin&, 0&, "wndate32class", vbNullString)
    Let strstring$ = gettext(lngatewin&)
    Let lnglong& = InStrRev(strstring$, "<br>")
    Let aimchatline$ = Right$(strstring$, Len(strstring$) - lnglong& - 3&)
End Function
Public Function aimchatlinemsg() As String
    Dim strtext As String, lnglong As Long, lngatewin As Long
    Let strtext$ = aimchatline$
    Let lnglong& = InStrRev(strtext$, "<br>")
    Let strtext$ = Right$(strtext$, Len(strtext$) - lnglong& - 3&)
    Let strtext$ = Right$(strtext$, Len(strtext$) - InStrRev(strtext$, Chr(34) & "#000000" & Chr(34) & ">") - 9&)
    Let strtext$ = striphtml(strtext$)
    Let aimchatlinemsg$ = strtext$
End Function
Public Function aimchatlinesn() As String
    Dim strtext As String, lnglong As Long
    Let strtext$ = aimchatline$
    If strtext$ = "" Then
        Exit Function
    Else
        Let lnglong& = InStrRev(strtext$, "<br>")
        Let strtext$ = Right$(strtext$, Len(strtext$) - lnglong& - 3&)
        Let lnglong& = InStr(strtext$, "#ff0000" & Chr(34) & ">")
        Let strtext$ = Mid$(strtext$, lnglong& + Len("#ff0000" & Chr(34) & ">"), Len(strtext$) - lnglong&)
        Let strtext$ = striphtml(strtext$)
        Let aimchatlinesn$ = Left$(strtext$, InStr(strtext$, ":") - 1&)
    End If
End Function
Public Function getlastlinefromstring(thestring As String) As String
    Let getlastlinefromstring$ = getlinefromstring(thestring$, getstringlinecount(thestring$) - 1&)
End Function
Public Sub aimkeyword(TheKeyWord$)
    Dim lngbuddywin As Long, lngeditwin As Long, lngicon As Long
    Let lngbuddywin& = FindWindow("_oscar_buddylistwin", vbNullString)
    Let lngeditwin& = FindWindowEx(lngbuddywin&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0, TheKeyWord$)
    Let lngicon& = FindWindowEx(lngbuddywin&, 0&, "_oscar_iconbtn", vbNullString)
    Call SendMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub aimclearim()
    Dim lngimwin As Long, lngatewin As Long
    Let lngimwin& = FindWindow("aim_imessage", vbNullString)
    Let lngatewin& = FindWindowEx(lngimwin&, 0&, "ate32class", vbNullString)
    Call SendMessageByString(lngatewin&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(lngatewin&, WM_CLEAR, 0&, 0&)
End Sub
Public Sub windowclose(lnghwnd As Long)
    Call SendMessageLong(lnghwnd&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub windowhide(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_HIDE)
End Sub
Public Sub windowshow(lnghwnd As Long)
    Call ShowWindow(lnghwnd&, SW_SHOW)
End Sub
Public Sub buddylist(openit As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long
    If openit = True Then
        Call keyword("buddy view")
        Exit Sub
    Else
        Let lngaol& = FindWindow("AOL Frame25", "America  Online")
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
        If lngbuddywin& <> 0& Then
            Call SendMessage(lngbuddywin&, WM_CLOSE, 0&, 0&)
        End If
    End If
End Sub
Public Function controltostring(thelist As Control, strseperator As String) As String
    Dim index As Long, strtext As String
    For index& = 0& To thelist.ListCount - 1&
        Let strtext$ = strtext$ & thelist.list(index&) & strseperator$
    Next index&
    Let strtext$ = Left$(strtext$, Len(strtext$) - Len(strseperator$))
    Let controltostring$ = strtext$
End Function
Public Sub ghost(ghost As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long
    Dim lngsetupbut As Long, lngsetupwin As Long, lngpprefbut1 As Long
    Dim lngpprefbut2 As Long, lngpprefbut3 As Long, lngpprefbut4 As Long
    Dim lngpprefbut As Long, lngpprefswin As Long, lngblock1 As Long, buddyopen As Boolean
    Dim lngblock2 As Long, lngblock3 As Long, lngblock4 As Long, lngblock5 As Long
    Dim lngblock6 As Long, lngblock As Long, lngsavebut1 As Long, lngsavebut2 As Long
    Dim lngsavebut3 As Long, lngsavebut As Long, lngokwin As Long, lngokbut As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
    If lngbuddywin& <> 0& Then
        Let buddyopen = True
    Else
        Call keyword("buddy view")
        Let buddyopen = False
    End If
    Do: DoEvents
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_AOL_Icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_AOL_Icon", vbNullString)
        Call PostMessage(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", getuser$ & "'s Buddy Lists")
        Let lngpprefbut1& = FindWindowEx(lngsetupwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngpprefbut2& = FindWindowEx(lngsetupwin&, lngpprefbut1&, "_AOL_Icon", vbNullString)
        Let lngpprefbut3& = FindWindowEx(lngsetupwin&, lngpprefbut2&, "_AOL_Icon", vbNullString)
        Let lngpprefbut4& = FindWindowEx(lngsetupwin&, lngpprefbut3&, "_AOL_Icon", vbNullString)
        Let lngpprefbut& = FindWindowEx(lngsetupwin&, lngpprefbut4&, "_AOL_Icon", vbNullString)
    Loop Until lngsetupwin& <> 0& And lngpprefbut1& <> 0& And lngpprefbut2& <> 0& And lngpprefbut3& <> 0& And lngpprefbut4& <> 0& And _
    lngpprefbut1& <> lngpprefbut2& And lngpprefbut2& <> lngpprefbut3& And lngpprefbut3& <> lngpprefbut4& And lngpprefbut4& <> lngpprefbut&: pause 0.2
    Do: DoEvents
        Call PostMessage(lngpprefbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngpprefbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngpprefswin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Privacy Preferences")
        Let lngblock1& = FindWindowEx(lngpprefswin&, 0&, "_AOL_Checkbox", vbNullString)
        Let lngblock2& = FindWindowEx(lngpprefswin&, lngblock1&, "_AOL_Checkbox", vbNullString)
        Let lngblock3& = FindWindowEx(lngpprefswin&, lngblock2&, "_AOL_Checkbox", vbNullString)
        Let lngblock4& = FindWindowEx(lngpprefswin&, lngblock3&, "_AOL_Checkbox", vbNullString)
        Let lngblock5& = FindWindowEx(lngpprefswin&, lngblock4&, "_AOL_Checkbox", vbNullString)
        Let lngblock6& = FindWindowEx(lngpprefswin&, lngblock5&, "_AOL_Checkbox", vbNullString)
        Let lngblock& = FindWindowEx(lngpprefswin&, lngblock6&, "_AOL_Checkbox", vbNullString)
        Let lngsavebut1& = FindWindowEx(lngpprefswin&, 0&, "_AOL_Icon", vbNullString)
        Let lngsavebut2& = FindWindowEx(lngpprefswin&, lngsavebut1&, "_AOL_Icon", vbNullString)
        Let lngsavebut3& = FindWindowEx(lngpprefswin&, lngsavebut2&, "_AOL_Icon", vbNullString)
        Let lngsavebut& = FindWindowEx(lngpprefswin&, lngsavebut3&, "_AOL_Icon", vbNullString)
    Loop Until lngblock1& <> 0& And lngblock2& <> 0& And lngblock3& <> 0& And lngblock4& <> 0& And lngblock5& <> 0& And lngblock6& <> 0& And lngblock& <> 0& And lngsavebut& <> 0& And lngsavebut1& <> 0& And lngsavebut2& <> 0& And lngsavebut3& <> 0& And _
    lngblock1& <> lngblock2& And lngblock2& <> lngblock3& And lngblock3& <> lngblock4& And lngblock4& <> lngblock5 And lngblock5& <> lngblock6& And lngblock6& <> lngblock& And lngsavebut1& <> lngsavebut2& And lngsavebut2& <> lngsavebut3& And lngsavebut3& <> lngsavebut&
    If ghost = True Then
        Call PostMessage(lngblock5&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngblock5&, WM_LBUTTONUP, 0&, 0&)
    ElseIf ghost = False Then
        Call PostMessage(lngblock3&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngblock3&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(lngblock&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngblock&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Call PostMessage(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
        Let lngokwin& = FindWindow("#32770", "America Online")
        Let lngokbut& = FindWindowEx(lngokwin&, 0&, "Button", "OK")
        pause 0.2
    Loop Until lngokwin& <> 0& And lngokbut& <> 0&
    Do: DoEvents
        Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngokwin& = FindWindow("#32770", "America Online")
    Loop Until lngokwin& = 0&: pause 0.2
    Call PostMessage(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Sub

Public Function replacestring(strstring As String, strwhat As String, strwith As String) As String
    Dim lngpos As Long
    Do While InStr(1&, strstring$, strwhat$)
        DoEvents
        Let lngpos& = InStr(1&, strstring$, strwhat$)
        Let strstring$ = Left$(strstring$, (lngpos& - 1&)) & Right$(strstring$, Len(strstring$) - (lngpos& + Len(strwhat$) - 1&))
    Loop
    Let replacestring$ = strstring$
End Function
Public Function isalive(strname As String) As Boolean
    Dim lngaol As Long, lngmdi As Long, lngerrorwin As Long
    Dim lngerrorview As Long, strerrorstring As String
    Dim lngmailwin As Long, lngnowin As Long, lngbutton As Long
    Call sendmail("*, " & strname$, "you alive?", " ")
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorview& = FindWindowEx(lngerrorwin&, 0&, "_AOL_View", vbNullString)
        Let strerrorstring$ = gettext(lngerrorview&)
    Loop Until lngerrorwin& <> 0& And lngerrorview& <> 0& And strerrorstring$ <> ""
    If InStr(LCase$(replacestring(strerrorstring$, " ", "")), LCase$(replacestring(strname$, " ", ""))) > 0& Then
        Let isalive = False
    Else
        Let isalive = True
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
Public Sub signonguest(strscreenname As String, strpassword As String)
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
    Call SendMessageByString(lngedit1&, WM_SETTEXT, 0&, strscreenname$)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strpassword$)
    Do: DoEvents
        Call PostMessage(lngokbut1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngokbut1&, WM_LBUTTONUP, 0&, 0&)
        Let lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        Let lngerrwin& = FindWindow("#32770", "America Online")
    Loop Until lngmodal& = 0& Or lngerrwin& <> 0&: pause 0.2
    If lngerrwin& <> 0& Then
        Let lngokbut2& = FindWindowEx(lngerrwin&, 0&, "Button", vbNullString)
        Do: DoEvents
            Call PostMessage(lngokbut2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut2&, WM_LBUTTONUP, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
        Loop Until findsignonwin& <> 0& Or findwelcomewin&
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
Public Sub buddyinvite(strpeople As String, strmessage As String, strroom As String, followtochat As Boolean)
    Dim lngaol As Long, lngmdi As Long, lnginvwin As Long, lngpeepsbox As Long
    Dim lngmsgbox As Long, lngroombox As Long, lngicon As Long
    Call keyword("buddy chat")
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lnginvwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy Chat")
        Let lngpeepsbox& = FindWindowEx(lnginvwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngmsgbox& = FindWindowEx(lnginvwin&, lngpeepsbox&, "_AOL_Edit", vbNullString)
        Let lngroombox& = FindWindowEx(lnginvwin&, lngmsgbox&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lnginvwin& <> 0& And lngpeepsbox& <> 0& And lngmsgbox& <> 0& And lngroombox& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngpeepsbox&, WM_SETTEXT, 0&, strpeople$)
    Call SendMessageByString(lngmsgbox&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageByString(lngroombox&, WM_SETTEXT, 0&, strroom$)
    Do: DoEvents
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Let lnginvwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Invitation from: " & getuser$)
        Let lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lnginvwin& <> 0& And lngicon& <> 0&: pause 0.2
    If followtochat = True Then
        Do: DoEvents
            Let lngicon& = FindWindowEx(lnginvwin&, 0&, "_AOL_Icon", vbNullString)
            Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Loop Until lngicon& = 0&
        Exit Sub
    End If
    Call PostMessage(lnginvwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addaddressbookperson(strfirstname As String, strlastname As String, stremailaddy As String, strnotes As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim lngaddywin As Long, lngnewwin As Long, lngmenu As Long, lngtabwin As Long
    Dim lngedit As Long, lngedit2 As Long, lngedit3 As Long, lngedit4 As Long
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
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyA, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyA, 0&)
    Do: DoEvents
        Let lngaddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Address Book")
        Let lngicon& = FindWindowEx(lngaddywin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngaddywin& <> 0& And lngicon& <> 0&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngnewwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "New Person")
        Let lngtabwin& = FindWindowEx(lngnewwin&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngedit& = FindWindowEx(lngtabwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngedit2& = FindWindowEx(lngtabwin&, lngedit&, "_AOL_Edit", vbNullString)
        Let lngedit3& = FindWindowEx(lngtabwin&, lngedit2&, "_AOL_Edit", vbNullString)
        Let lngedit4& = FindWindowEx(lngtabwin&, lngedit3&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lngnewwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngnewwin& <> 0& And lngtabwin& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngedit4& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strfirstname$): DoEvents
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, strlastname$): DoEvents
    Call SendMessageByString(lngedit3&, WM_SETTEXT, 0&, stremailaddy$): DoEvents
    Call SendMessageByString(lngedit4&, WM_SETTEXT, 0&, strnotes$): DoEvents
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngnewwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "New Person")
    Loop Until lngnewwin& = 0&
    Call windowclose(lngaddywin&)
End Sub
Public Sub addaddressbookgroup(strgroupname As String, straddresss As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim lngaddywin As Long, lngnewwin As Long, lngmenu As Long, lngtabwin As Long
    Dim lngedit As Long, lngedit2 As Long, lngedit3 As Long, lngedit4 As Long
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
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyA, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyA, 0&)
    Do: DoEvents
        Let lngaddywin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Address Book")
        Let lngicon& = FindWindowEx(lngaddywin&, 0&, "_AOL_Icon", vbNullString)
        Let lngicon& = FindWindowEx(lngaddywin&, lngicon&, "_AOL_Icon", vbNullString)
    Loop Until lngaddywin& <> 0& And lngicon& <> 0&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngnewwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "New Group")
        Let lngedit& = FindWindowEx(lngnewwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngedit2& = FindWindowEx(lngnewwin&, lngedit&, "_AOL_Edit", vbNullString)
        Let lngicon& = FindWindowEx(lngnewwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngnewwin& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngicon& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strgroupname$): DoEvents
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, straddresss$): DoEvents
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngnewwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "New Person")
    Loop Until lngnewwin& = 0&
    Call windowclose(lngaddywin&)
End Sub
Public Sub addaolshortcutkey(theshortcutkey As aolshortcutkeys, strtitle As String, straddy As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim lngaddywin As Long, lngnewwin As Long, lngmenu As Long, lngtabwin As Long
    Dim lngedit As Long, lngedit2 As Long, lngedit3 As Long, lngedit4 As Long, lngshortcutwin As Long
    Dim lngedit5 As Long, lngedit6 As Long, lngedit7 As Long, lngedit8 As Long
    Dim lngedit9 As Long, lngedit10 As Long, lngedit11 As Long, lngedit12 As Long
    Dim lngedit13 As Long, lngedit14 As Long, lngedit15 As Long, lngedit16 As Long
    Dim lngedit17 As Long, lngedit18 As Long, lngedit19 As Long, lngedit20 As Long
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
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyDown, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyRight, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyRight, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyReturn, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyReturn, 0&)
    Do: DoEvents
        Let lngshortcutwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
        Let lngicon& = FindWindowEx(lngshortcutwin&, 0&, "_AOL_Icon", vbNullString)
        Let lngedit& = FindWindowEx(lngshortcutwin&, 0&, "_AOL_Edit", vbNullString)
        Let lngedit2& = FindWindowEx(lngshortcutwin&, lngedit&, "_AOL_Edit", vbNullString)
        Let lngedit3& = FindWindowEx(lngshortcutwin&, lngedit2&, "_AOL_Edit", vbNullString)
        Let lngedit4& = FindWindowEx(lngshortcutwin&, lngedit3&, "_AOL_Edit", vbNullString)
        Let lngedit5& = FindWindowEx(lngshortcutwin&, lngedit4&, "_AOL_Edit", vbNullString)
        Let lngedit6& = FindWindowEx(lngshortcutwin&, lngedit5&, "_AOL_Edit", vbNullString)
        Let lngedit7& = FindWindowEx(lngshortcutwin&, lngedit6&, "_AOL_Edit", vbNullString)
        Let lngedit8& = FindWindowEx(lngshortcutwin&, lngedit7&, "_AOL_Edit", vbNullString)
        Let lngedit9& = FindWindowEx(lngshortcutwin&, lngedit8&, "_AOL_Edit", vbNullString)
        Let lngedit10& = FindWindowEx(lngshortcutwin&, lngedit9&, "_AOL_Edit", vbNullString)
        Let lngedit11& = FindWindowEx(lngshortcutwin&, lngedit10&, "_AOL_Edit", vbNullString)
        Let lngedit12& = FindWindowEx(lngshortcutwin&, lngedit11&, "_AOL_Edit", vbNullString)
        Let lngedit13& = FindWindowEx(lngshortcutwin&, lngedit12&, "_AOL_Edit", vbNullString)
        Let lngedit14& = FindWindowEx(lngshortcutwin&, lngedit13&, "_AOL_Edit", vbNullString)
        Let lngedit15& = FindWindowEx(lngshortcutwin&, lngedit14&, "_AOL_Edit", vbNullString)
        Let lngedit16& = FindWindowEx(lngshortcutwin&, lngedit15&, "_AOL_Edit", vbNullString)
        Let lngedit17& = FindWindowEx(lngshortcutwin&, lngedit16&, "_AOL_Edit", vbNullString)
        Let lngedit18& = FindWindowEx(lngshortcutwin&, lngedit17&, "_AOL_Edit", vbNullString)
        Let lngedit19& = FindWindowEx(lngshortcutwin&, lngedit18&, "_AOL_Edit", vbNullString)
        Let lngedit20& = FindWindowEx(lngshortcutwin&, lngedit19&, "_AOL_Edit", vbNullString)
    Loop Until lngshortcutwin& <> 0& And lngicon& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngedit3& <> 0& And lngedit4& <> 0& And lngedit5& <> 0& And lngedit6& <> 0& And lngedit7& <> 0& And lngedit8& <> 0& And lngedit9& <> 0& And lngedit10& <> 0& _
    And lngedit11& <> 0& And lngedit12& <> 0& And lngedit13& <> 0& And lngedit14& <> 0& And lngedit15& <> 0& And lngedit16& <> 0& And lngedit17& <> 0& And lngedit18& <> 0& And lngedit19& <> 0& And lngedit20& <> 0& And _
    lngedit& <> lngedit2& And lngedit2& <> lngedit3& And lngedit4& <> lngedit5& And lngedit5& <> lngedit6& And lngedit6& <> lngedit7& And lngedit7& <> lngedit8& And _
    lngedit8& <> lngedit9& And lngedit9& <> lngedit10& And lngedit10& <> lngedit11& And lngedit11& <> lngedit12& And lngedit12& <> lngedit13& And lngedit13& <> lngedit14& And lngedit14& <> lngedit15& And lngedit15& <> lngedit16& And _
    lngedit16& <> lngedit17& And lngedit17& <> lngedit18& And lngedit18& <> lngedit19&
    If theshortcutkey = ctrl0 Then
        Let lngedit& = lngedit19&
        Let lngedit2& = lngedit20&
        GoTo harschniff 'harschniff = whatever ;0) [its a word i made up, lol]
    End If
    If theshortcutkey = ctrl1 Then
        Let lngedit& = lngedit&
        Let lngedit2& = lngedit2&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl2 Then
        Let lngedit& = lngedit3&
        Let lngedit2& = lngedit4&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl3 Then
        Let lngedit& = lngedit5&
        Let lngedit2& = lngedit6&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl4 Then
        Let lngedit& = lngedit7&
        Let lngedit2& = lngedit8&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl5 Then
        Let lngedit& = lngedit9&
        Let lngedit2& = lngedit10&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl6 Then
        Let lngedit& = lngedit11&
        Let lngedit2& = lngedit12&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl7 Then
        Let lngedit& = lngedit13&
        Let lngedit2& = lngedit14&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl8 Then
        Let lngedit& = lngedit15&
        Let lngedit2& = lngedit16&
        GoTo harschniff
    End If
    If theshortcutkey = ctrl9 Then
        Let lngedit& = lngedit17&
        Let lngedit2& = lngedit18&
        GoTo harschniff
    End If
harschniff:
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strtitle$)
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, straddy$)
    Do: DoEvents
        Let lngshortcutwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
        Call SendMessageLong(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngicon&, WM_LBUTTONUP, 0&, 0&)
        pause 0.2
    Loop Until lngshortcutwin& = 0&
End Sub
Public Sub formsizetowindow(frmform As Object, lngwindow As Long)
    Dim hwndrect As RECT, lngrect As Long
    Let lngrect& = GetWindowRect(lngwindow&, hwndrect)
    With frmform
        .Top = (hwndrect.Top * Screen.TwipsPerPixelY) - 1700&
        .Left = hwndrect.Left * Screen.TwipsPerPixelX
        .Height = (hwndrect.Bottom - hwndrect.Top) * Screen.TwipsPerPixelY
        .Width = (hwndrect.Right - hwndrect.Left) * Screen.TwipsPerPixelX
    End With
End Sub
Public Sub addpicturetoaolmdi(thepic As PictureBox)
    Dim lngaol As Long, lngmdi As Long, therect As RECT, lngrect As Long, lngchild As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Call formsizetowindow(thepic, lngmdi&)
    lngrect& = GetWindowRect(lngmdi&, therect)
    'Call StretchBlt(thepic.hdc, therect.Top, therect.Left, thepic.Height, thepic.Width, thepic.Picture, thepic.Top, thepic.Left, thepic.Picture.Width, thepic.Picture.Height, 0&)
    Call SetParent(thepic.hWnd, lngmdi&)
    Call BringWindowToTop(thepic.hWnd)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Do: DoEvents
        Call BringWindowToTop(lngchild&)
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
    Loop Until lngchild& = 0&
End Sub
Public Function trimspaces(strtext As String) As String
    Let trimspaces$ = replacestring(strtext$, " ", "")
End Function
Public Function trimenters(strtext As String) As String
    Let trimenters$ = replacestring(strtext$, Chr$(13&), "")
    Let trimenters$ = replacestring(trimenters$, Chr$(10&), "")
End Function
Public Function trimchar(strtext As String, strchar As String) As String
    trimchar$ = replacestring(strtext$, strchar$, "")
End Function
Public Function convstringtolong(strstring As String) As Long
    If IsNumeric(strstring$) Then Let convstringtolong& = CLng(strstring$)
End Function
Public Function convstringtointeger(strstring As String) As Integer
    If IsNumeric(strstring$) Then Let convstringtointeger% = CInt(strstring$)
End Function
Public Function convstringtoboolean(strstring As String) As Boolean
    If LCase$(strstring$) <> "true" And LCase$(strstring$) <> "false" Then Exit Function
    Let convstringtoboolean = CBool(strstring$)
End Function
Public Function convlongtostring(lnglong As Long) As String
    Let convlongtostring$ = CStr(lnglong&)
End Function
Public Function convlongtointeger(lnglong As Long) As Integer
    Let convlongtointeger% = CInt(lnglong&)
End Function
Public Function convlongtoboolean(lnglong As Long) As Boolean
    If lnglong& > 1& Then Exit Function
    Let convlongtoboolean = CBool(lnglong&)
End Function
Public Function convinttostring(lnginteger As Integer) As String
    Let convinttostring$ = CStr(lnginteger%)
End Function
Public Function convinttolong(intinteger As Integer) As Long
    Let convinttolong& = CInt(intinteger%)
End Function
Public Function convinttoboolean(intinteger As Integer) As Boolean
    If intinteger% > 1& Then Exit Function
    Let convinttoboolean = CBool(intinteger%)
End Function
Public Sub waitforlisttoload(lnglist As Long)
    Dim lngcount As Long
    Do: DoEvents
        Let lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&): Call pause(2&)
        If lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Do
    Loop
End Sub
Public Sub waitfortextboxtoload(lngtextbox As Long)
    Dim strtext As String
    Do: DoEvents
        Let strtext$ = gettext(lngtextbox&): Call pause(2&)
        If strtext$ = gettext(lngtextbox&) Then Exit Do
    Loop
End Sub
Public Sub clearchatwindow()
    Dim lngroom As Long
    Let lngroom& = findroom&
    Let lngroom& = FindWindowEx(lngroom&, 0&, "richcntl", vbNullString)
    Call SendMessage(lngroom&, WM_CLEAR, 0&, 0&)
    Call SendMessageByString(lngroom&, WM_SETTEXT, 0&, "")
End Sub
Public Sub scrollprofile(strscreenname As String)
    Dim index As Long, strtext As String
    Let strtext$ = getprofile(strscreenname$)
    Call sendchat2("scrolling " & strscreenname$ & "'s profile")
    For index& = 0& To getstringlinecount(strtext$)
        Call sendchat(getlinefromstring(strtext$, index&))
        pause 0.6
    Next index&
End Sub
Public Sub antiidle()
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
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
        If InStr(getcaption(lngchild&), "Welcome, ") Then Exit Do
    Loop While lngchild& <> 0&
    If lngchild& = 0& Then
        Let getuser$ = ""
    ElseIf lngchild& <> 0& Then
        Let getuser$ = Mid$(getcaption$(lngchild&), 10&, InStr(getcaption$(lngchild&), "!") - 10)
    End If
End Function
Public Function findwelcomewin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
        If InStr(getcaption(lngchild&), "Welcome, ") Then findwelcomewin& = lngchild&: Exit Function
    Loop While lngchild& <> 0&
End Function
Public Sub sendchatlist(thecontrol As Control)
    Dim lngindex As Long
    For lngindex& = 0& To thecontrol.ListCount
        Call sendchat(thecontrol.list(lngindex&))
        pause 0.7
    Next lngindex&
End Sub
Public Sub getserverstatus(strserver As String)
    Call sendchat("/" & strserver$ & " send status")
End Sub
Public Sub getserveritembyindex(strserver As String, lngindexnumber As Long)
    Call sendchat("/" & strserver$ & " send " & lngindexnumber&)
End Sub
Public Sub getserverlists(strserver As String)
    Call sendchat("/" & strserver$ & " send list")
End Sub
Public Sub getserveritembyname(strserver As String, stritemname As String)
    Call sendchat("/" & strserver$ & " find " & stritemname$)
End Sub
Public Sub filesearch(strfilename As String)
    'doesn't work for most recent aol4
    Dim lngaol As Long, lngmdi As Long, lngsearchwin As Long
    Dim lngicon As Long, lngeditwin As Long, lngsearchwin2 As Long
    Call keyword("file search")
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Do: DoEvents
        Let lngsearchwin& = FindWindowEx(lngmdi&, 0&, "aol child", "filesearch")
        Let lngicon& = FindWindowEx(lngsearchwin&, 0&, "_aol_icon", vbNullString)
        Let lngicon& = FindWindowEx(lngsearchwin&, lngicon&, "_aol_icon", vbNullString)
    Loop Until lngsearchwin& <> 0& And lngicon& <> 0&
    Call SendMessageLong(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsearchwin2& = FindWindowEx(lngmdi&, 0&, vbNullString, "software search")
    Loop Until lngsearchwin2& <> 0&
    Call SendMessageLong(lngsearchwin&, WM_CLOSE, 0&, 0&)
    Let lngeditwin& = FindWindowEx(lngsearchwin2&, 0&, "_aol_edit", vbNullString)
    Call SendMessageByString(lngeditwin&, WM_SETTEXT, 0&, strfilename$)
    Call SendMessageByNum(lngeditwin&, WM_CHAR, 0&, 13&)
End Sub
Public Function stringreverse(strstring As String) As String
    'note: for vb6 users this code can be replaced with
    'strstring$ = strreverse(strstring$)
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strstring$, lngindex&, 1&)
        Let stringreverse$ = strnextchar$ & stringreverse$
    Loop
End Function
Public Function stringdotted(strstring As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strstring$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & "o"
        Let stringdotted$ = stringdotted$ & strnextchar$
    Loop
End Function
Public Function stringcustomized(strstring As String, streveryotherchar As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strstring$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & streveryotherchar$
        Let stringcustomized$ = stringcustomized$ & strnextchar$
    Loop
End Function
Public Function stringspaced(strstring As String) As String
    Dim lnglen As Long, lngindex As Long, strnextchar As String
    Let lnglen& = Len(strstring$)
    Do While lngindex& <= lnglen&
        Let lngindex& = lngindex& + 1&
        Let strnextchar$ = Mid$(strstring$, lngindex&, 1&)
        Let strnextchar$ = strnextchar$ & " "
        Let stringspaced$ = stringspaced$ & strnextchar$
    Loop
End Function
Public Function stringlinked(strurl As String, strtext$) As String
    Let stringlinked$ = "< a href=" & Chr$(34&) & strurl$ & Chr$(34&) & ">" & strtext$ & "</a>"
End Function
Public Function stringboldfirstletter(startstring As String, Optional boldcolor As String, Optional textcolor As String) As String
    'idea from beav
    Dim strcount As Long, thelen As Long, strletter As String, strletter2 As String, strstring As String
    Dim strbefore As String, strafter As String, strcheck As String
    If startstring$ = "" Then
        Exit Function
    Else
        Let startstring$ = "<b>" & Left$(startstring$, 1&) & "</b>" & Right$(startstring$, Len(startstring$) - 1&)
        Let thelen& = Len(startstring$)
        If thelen& = 0& Then
            Exit Function
        Else
            Let strbefore$ = ""
            Let strafter$ = ""
            Let strstring$ = strbefore$ & Left$(startstring$, 1&) & strafter$
            Let startstring$ = Right$(startstring$, Len(startstring$) - 1&)
            Let thelen& = Len(startstring$)
            For strcount& = 1& To thelen&
                Let strletter$ = Mid(startstring$, strcount&, 1&)
                If strcheck$ = "1" Then
                    Let strcheck$ = "0"
                    Let strstring$ = strstring$
                    Let strletter$ = ""
                End If
                If strletter$ = " " Then
                    Let strletter2$ = Mid(startstring$, strcount& + 1&, 1&)
                    Let strletter2$ = "<b>" & strletter2$ & "</b>"
                    Let strletter$ = strbefore$ & strletter$ & strletter2$ & strafter$
                    Let strcheck$ = "1"
                End If
                Let strstring$ = strstring$ & strletter$
            Next strcount&
            Let stringboldfirstletter$ = strstring$
        End If
    End If
End Function
Public Sub resethistorycombo()
    Dim lngaol As Long, lngtoolbar As Long, lngcombo As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_aol_toolbar", vbNullString)
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
    Call SendMessage(lngcombo&, CB_RESETCONTENT, 0&, 0&)
End Sub
Public Function getlistitemindex(thelist As Long, item As String) As Long
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long, index As Long
    Dim screenname As String, itmhold As Long, psnHold As Long
    Dim rbytes As Long, cprocess As Long
    Let rlist& = thelist&
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            If screenname$ = item$ Then
                Let getlistitemindex& = index&
                Call CloseHandle(mthread&)
                Exit Function
            End If
        Next index&
        Let getlistitemindex& = -1&
        Call CloseHandle(mthread&)
    End If
End Function
Public Function getlistcount(lnglist As Long) As Long
    Let getlistcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function gettreecount(lngtree As Long) As Long
    Let gettreecount& = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function gettreeitemindex(lngtree As Long, stritem As String) As Long
    Dim index As Long, lngtextlen As Long, strstring As String
    For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
        Let strstring$ = String$(lngtextlen& + 1&, 0&)
        Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strstring$)
        If strstring$ = stritem$ Then
            Let gettreeitemindex& = index&
            Exit Function
        End If
    Next index&
End Function
Public Function gettreeitemtext(lngtree As Long, index As Long) As String
    Dim lngtextlen As Long
    Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
    Let gettreeitemtext$ = String$(lngtextlen& + 1&, 0&)
    Call SendMessageByString(lngtree&, LB_GETTEXT, index&, gettreeitemtext$)
End Function
Public Sub addorremovebuddy(strname As String, strgroup As String, booladdbuddy As Boolean)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let strrealgroup$ = strgroup$
    Let strgroup$ = replacestring(strgroup$, " ", "")
    Let strname$ = replacestring(strname$, " ", "")
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until lngsetupwin& <> 0& And lnggroups& <> 0& And lngeditbut& <> 0&
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&)
        Let strstring$ = getlistitemtext(lnggroups&, lngindex&)
        Let strstring$ = replacestring(Left$(LCase$(strstring$), InStr(strstring, Chr$(9&)) - 1&), " ", "")
        If strstring$ = strgroup$ Then
            Exit For
        End If
    Next lngindex&
    If lngindex& = -1& Then
        Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
        Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    Else
        Let lngtab& = InStr(strstring$, Chr$(9&))
        If lngtab& <> 0& Then Let strstring$ = Left$(strstring$, lngtab& - 1&)
        Call SendMessageLong(lnggroups&, LB_SETCURSEL, CLng(lngindex&), 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lnggroupwin& = FindWindowEx(lngmdi&, 0&, "aol child", "edit list " & strrealgroup$)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, 0&, "_aol_edit", vbNullString)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, lngaddbox&, "_aol_edit", vbNullString)
            Let lngblist& = FindWindowEx(lnggroupwin&, 0&, "_aol_listbox", vbNullString)
            Let lngaddbut& = FindWindowEx(lnggroupwin&, 0&, "_aol_icon", vbNullString)
            Let lngremovebut& = FindWindowEx(lnggroupwin&, lngaddbut&, "_aol_icon", vbNullString)
            Let lngsavebut& = FindWindowEx(lnggroupwin&, lngremovebut&, "_aol_icon", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, 0&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
        Loop Until lnggroupwin& <> 0& And lngaddbox& <> 0& And lngblist <> 0& And lngaddbut <> 0& And lngremovebut& <> 0& And lngsavebut& <> 0& And lngstatic& <> 0&
        Call waitforlisttoload(lngblist&)
        If booladdbuddy = True Then
            If getlistitemindex(lngblist&, LCase$(replacestring(strname$, " ", ""))) <> -1& Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
                Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
                Exit Sub
            End If
            Call SendMessageByString(lngaddbox&, WM_SETTEXT, 0&, strname$)
            Call SendMessageLong(lngaddbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngaddbut&, WM_LBUTTONUP, 0&, 0&)
            Do: DoEvents
                Let lngindex& = getlistitemindex(lngblist&, LCase$(replacestring(strname$, " ", "")))
                Let strstring$ = gettext(lngstatic&) ' meaning buddy is already there !
            Loop Until lngindex& <> 0& Or strstring$ <> " "
            If strstring$ <> " " Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
            ElseIf strstring$ = " " Then
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            End If
        ElseIf booladdbuddy = False Then
            Let lngbindex& = getlistitemindex(lngblist&, LCase$(replacestring(strname$, " ", "")))
            If lngbindex& = -1& Then
                Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
                Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
                Exit Sub
            End If
            Call SendMessageLong(lngblist&, LB_SETCURSEL, CLng(lngbindex&), 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
        End If
        Do Until lngokwin& <> 0& And lngokbut& <> 0&
            Let lngokwin& = FindWindow("#32770", "america online")
            Let lngokbut& = FindWindowEx(lngokwin&, 0&, "button", "ok")
        Loop: DoEvents
        Do Until lngokwin& = 0&
            Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
            Let lngokwin& = FindWindow("#32770", "america online")
        Loop: DoEvents
        Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub usernamereplace(strsn As String, strnewsn As String, straolpath As String)
    Dim lngsnlen As Long, lngnumhold As Long, strtext As String, lngthesn As Long
    Dim lngthesn2 As Long, strtoreplace As String, lnglenofile2 As Long
    Call windowclose(findsignonwin&)
    Let lngsnlen& = Len(strsn$)
    Select Case lngsnlen&
        Case 3&
            Let strsn$ = strsn$ & "       "
        Case 4&
            Let strsn$ = strsn$ & "      "
        Case 5&
            Let strsn$ = strsn$ & "     "
        Case 6&
            Let strsn$ = strsn$ & "    "
        Case 7&
            Let strsn$ = strsn$ & "   "
        Case 8&
            Let strsn$ = strsn$ & "  "
        Case 9&
            Let strsn$ = strsn$ & " "
        Case 10&
            Let strsn$ = strsn$
    End Select
    Let lngsnlen& = Len(strnewsn$)
    Select Case lngsnlen&
        Case 3&
            Let strnewsn$ = strnewsn$ & "       "
        Case 4&
            Let strnewsn$ = strnewsn$ & "      "
        Case 5&
            Let strnewsn$ = strnewsn$ & "     "
        Case 6&
            Let strnewsn$ = strnewsn$ & "    "
        Case 7&
            Let strnewsn$ = strnewsn$ & "   "
        Case 8&
            Let strnewsn$ = strnewsn$ & "  "
        Case 9&
            Let strnewsn$ = strnewsn$ & " "
        Case 10&
            Let strnewsn$ = strnewsn$
    End Select
    Let lngnumhold& = 1&
    If Right$(straolpath$, 1&) = "\" Then straolpath$ = Right$(straolpath$, 1&)
    Do Until 2& > 3&: DoEvents
        Let strtext$ = ""
        On Error Resume Next
        Open straolpath$ & "\idb\main.idx" For Binary As #1&
            If Err Then
                Exit Sub
            Else
                Let strtext$ = String$(32000&, 0&)
                Get #1, lngnumhold&, strtext$
            End If
        Close #1&
        Open straolpath$ & "\idb\main.idx" For Binary As #2&
            Let lngthesn& = InStr(1&, strtext$, strsn$, 1&)
            If lngthesn& Then
                Mid(strtext$, lngthesn&) = strnewsn$
                Let strtoreplace$ = strnewsn$
                Put #2&, (lngnumhold& + lngthesn&) - 1&, strtoreplace$
getthatsn:
                DoEvents
                Let lngthesn2& = InStr(1&, strtext$, strsn$, 1&)
                If lngthesn2& Then
                    Mid(strtext$, lngthesn2&) = strnewsn$
                    Put #2&, (lngnumhold& + lngthesn2&) - 1&, strtoreplace$
                    GoTo getthatsn
                End If
            End If
            Let lngnumhold& = lngnumhold& + 32000&
            Let lnglenofile2& = LOF(2&)
        Close #2&
        If lngnumhold& > lnglenofile2& Then Call runaolmenubystring("&Sign On Screen"): Exit Sub
    Loop
End Sub
Public Sub addfavoriteitem(strdescription As String, strurl As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
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
Public Sub addfavoritefolder(strfoldername As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long, lngmenu As Long, lngcheck As Long
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
    'this will enable the aol toolbar and items when signed off
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Call EnableWindow(lngicon&, True)
    For index& = 0& To 20&
        Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
        Call EnableWindow(lngicon&, True)
    Next index&
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
End Sub
Public Sub aaolonlinenow()
    'this will enable the aol toolbar and items when signed off
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim index As Long, lngcombo As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Call EnableWindow(lngicon&, True)
    For index& = 0& To 20&
        Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
        Call EnableWindow(lngicon&, True)
    Next index&
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
End Sub
Public Sub aolautosncollect(thelist As Control, lnghowmanysns As Long)
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
    pause 0.2
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
        pause 1
    Loop Until lngchild& <> 0& And lnglist1& <> 0& And lnglist& <> 0& And lnglist1& <> lnglist& And lngicon& <> 0& And lngicon1& <> 0& And lngicon2& <> 0& And lngicon3& <> 0& And lngicon4& <> 0& And lngicon5& <> 0& And lngicon6& <> 0& And lngicon7& <> 0& And lngicon8& <> 0& And _
    lngicon1& <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon4& And lngicon4& <> lngicon5& And lngicon5& <> lngicon6& And lngicon6& <> lngicon7& And lngicon7& <> lngicon8& And lngicon8& <> lngicon&
    Call waitforlisttoload(lnglist&)
    Do: DoEvents
        Call clicklistindex(lnglist&, lngcounter&)
        Call clickicon(lngicon&)
        Let lngcounter& = lngcounter& + 1&
        Do: DoEvents
            Let lngchatting& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Who's Chatting")
            Let lnglist1& = FindWindowEx(lngchatting&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until lngchatting& <> 0& And lnglist1& <> 0&
        Call waitforlisttoload(lnglist1&)
        Call addlisttocontrol(lnglist1&, thelist)
        Call PostMessage(lngchatting&, WM_CLOSE, 0&, 0&)
    Loop Until thelist.ListCount > lnghowmanysns&
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
End Sub
Public Sub ctrlaltdel(enabled As Boolean)
    Dim lnggogo As Long, pOld As Boolean
    Let lnggogo& = SystemParametersInfo(SPI_SCREENSAVERRUNNING, enabled, pOld, 0&)
End Sub
Public Sub formrattle(frmform As Form)
    Dim index As Long, firstleft As Long
    Let firstleft& = frmform.Left
    frmform.Show
    Let frmform.DrawMode = 2&
    For index& = 0& To 15&
        Let frmform.Left = firstleft& + 25&
        Let frmform.Left = frmform.Left - 50&
    Next index&
End Sub
Public Sub autophader(spyworks As Control, thewp, color1 As Long, color2 As Long, makewavy As Boolean)
    'call this in the SubClass1_WndMessageX of spyworks
    'for 'thewp' just put 'wp'
    Dim lngchat As Long, lngrich As Long, strmessage As String, strnewmessage As String
    Let lngchat& = findroom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let spyworks.hwndparam = lngrich&
    If thewp = 13& Then
        Let spyworks.hwndparam = 0&
        Let strmessage$ = gettext(lngrich&)
        Let strnewmessage$ = fadetext2byrgb(color1&, color2&, strmessage$, makewavy)
        Call SendMessageLong(lngrich&, EM_SETSEL, 0&, Len(strmessage$))
        Call SendMessageByString(lngrich&, EM_REPLACESEL, 0&, strnewmessage$)
        Call SendMessageLong(lngrich&, WM_CHAR, 13&, 0&)
        Let spyworks.hwndparam = 0&
    End If
End Sub
Public Sub addbuddiestocontrol(thelist As Control)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long, lngsetupwin2 As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long, lngindex2 As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lngsetupwin2& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy list")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until (lngsetupwin& <> 0& Or lngsetupwin2& <> 0&) And lnggroups& <> 0& And lngeditbut& <> 0&
    If lngsetupwin& = 0& And lngsetupwin2& <> 0& Then
        Let lngsetupwin& = lngsetupwin2&
    End If
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&) - 1&
        Call SendMessageLong(lnggroups&, LB_SETCURSEL, CLng(lngindex&), 0&)
        Let strrealgroup$ = Left$(getlistitemtext(lnggroups&, lngindex&), InStr(getlistitemtext(lnggroups&, lngindex&), Chr$(9&)) - 1&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lnggroupwin& = FindWindowEx(lngmdi&, 0&, "aol child", "edit list " & strrealgroup$)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, 0&, "_aol_edit", vbNullString)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, lngaddbox&, "_aol_edit", vbNullString)
            Let lngblist& = FindWindowEx(lnggroupwin&, 0&, "_aol_listbox", vbNullString)
            Let lngaddbut& = FindWindowEx(lnggroupwin&, 0&, "_aol_icon", vbNullString)
            Let lngremovebut& = FindWindowEx(lnggroupwin&, lngaddbut&, "_aol_icon", vbNullString)
            Let lngsavebut& = FindWindowEx(lnggroupwin&, lngremovebut&, "_aol_icon", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, 0&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
        Loop Until lnggroupwin& <> 0& And lngaddbox& <> 0& And lngblist <> 0& And lngaddbut <> 0& And lngremovebut& <> 0& And lngsavebut& <> 0& And lngstatic& <> 0&
        Call waitforlisttoload(lngblist&)
        For lngindex2& = 0& To SendMessage(lngblist&, LB_GETCOUNT, 0&, 0&) - 1&
            thelist.AddItem (getlistitemtext(lngblist&, lngindex2&))
        Next lngindex2&
        Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
    Next lngindex&
    Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub alphabetizebuddies(thelist As Control)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long, lngsetupwin2 As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long, lngindex2 As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lngsetupwin2& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy list")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until (lngsetupwin& <> 0& Or lngsetupwin2& <> 0&) And lnggroups& <> 0& And lngeditbut& <> 0&
    If lngsetupwin& = 0& And lngsetupwin2& <> 0& Then
        Let lngsetupwin& = lngsetupwin2&
    End If
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&) - 1&
        Call SendMessageLong(lnggroups&, LB_SETCURSEL, CLng(lngindex&), 0&)
        Let strrealgroup$ = Left$(getlistitemtext(lnggroups&, lngindex&), InStr(getlistitemtext(lnggroups&, lngindex&), Chr$(9&)) - 1&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lnggroupwin& = FindWindowEx(lngmdi&, 0&, "aol child", "edit list " & strrealgroup$)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, 0&, "_aol_edit", vbNullString)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, lngaddbox&, "_aol_edit", vbNullString)
            Let lngblist& = FindWindowEx(lnggroupwin&, 0&, "_aol_listbox", vbNullString)
            Let lngaddbut& = FindWindowEx(lnggroupwin&, 0&, "_aol_icon", vbNullString)
            Let lngremovebut& = FindWindowEx(lnggroupwin&, lngaddbut&, "_aol_icon", vbNullString)
            Let lngsavebut& = FindWindowEx(lnggroupwin&, lngremovebut&, "_aol_icon", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, 0&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
        Loop Until lnggroupwin& <> 0& And lngaddbox& <> 0& And lngblist <> 0& And lngaddbut <> 0& And lngremovebut& <> 0& And lngsavebut& <> 0& And lngstatic& <> 0&
        Call waitforlisttoload(lngblist&)
        Call SendMessageLong(lngaddbox&, WM_CLEAR, 0&, 0&)
        For lngindex2& = 0& To SendMessage(lngblist&, LB_GETCOUNT, 0&, 0&) - 1&
            Call SendMessageLong(lngblist&, LB_SETCURSEL, CLng(lngindex2&), 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngremovebut&, WM_LBUTTONUP, 0&, 0&)
            Do: DoEvents
            Loop Until gettext(lngaddbox&) <> ""
            thelist.AddItem (gettext(lngaddbox&))
            Call SendMessageLong(lngaddbox&, WM_CLEAR, 0&, 0&)
        Next lngindex2&
        thelist.Sorted = True
        For lngindex2& = 0& To thelist.ListCount - 1&
            Call SendMessageByString(lngaddbox&, WM_SETTEXT, 0&, thelist.list(lngindex2&))
            Call SendMessageLong(lngaddbut&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessageLong(lngaddbut&, WM_LBUTTONUP, 0&, 0&)
        Next lngindex2&
        Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
    Next lngindex&
    Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addbuddygroupstocontrol(thelist As Control)
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long, lngsetupwin2 As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long, lngindex2 As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lngsetupwin2& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy list")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until (lngsetupwin& <> 0& Or lngsetupwin2& <> 0&) And lnggroups& <> 0& And lngeditbut& <> 0&
    If lngsetupwin& = 0& And lngsetupwin2& <> 0& Then
        Let lngsetupwin& = lngsetupwin2&
    End If
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&) - 1&
        thelist.AddItem Left$(getlistitemtext(lnggroups&, lngindex&), InStr(getlistitemtext(lnggroups&, lngindex&), Chr$(9&)) - 1&)
    Next lngindex&
    Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
    Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function addbuddygroupstostring(strseperator As String) As String
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long, lngsetupwin2 As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long, lngindex2 As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lngsetupwin2& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy list")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until (lngsetupwin& <> 0& Or lngsetupwin2& <> 0&) And lnggroups& <> 0& And lngeditbut& <> 0&
    If lngsetupwin& = 0& And lngsetupwin2& <> 0& Then
        Let lngsetupwin& = lngsetupwin2&
    End If
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&) - 1&
        Let addbuddygroupstostring$ = addbuddygroupstostring$ & Left$(getlistitemtext(lnggroups&, lngindex&), InStr(getlistitemtext(lnggroups&, lngindex&), Chr$(9&)) - 1&) & strseperator$
    Next lngindex&
    Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
    Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Function
Public Function addbuddiestostring(strseperator As String) As String
    Dim lngaol As Long, lngmdi As Long, lngbuddywin As Long, lngsetupbut As Long, lnggroupwin As Long, lngsetupwin2 As Long
    Dim lngsetupwin As Long, lnggroups As Long, lngeditbut As Long, strstring As String, lngaddbut As Long, lngindex2 As Long
    Dim lngindex As Long, lngaddbox As Long, lngblist As Long, lngremovebut As Long, lnglong As Long, lngtab As Long
    Dim lngsavebut As Long, lngstatic As Long, lngokwin As Long, lngokbut As Long, lngbindex As Long, strrealgroup As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
    If lngbuddywin& <= 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        Let lngbuddywin& = FindWindowEx(lngmdi&, 0&, "aol child", "buddy list window")
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, 0&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
        Let lngsetupbut& = FindWindowEx(lngbuddywin&, lngsetupbut&, "_aol_icon", vbNullString)
    Loop Until lngbuddywin& <> 0& And lngsetupbut& <> 0&
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsetupbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngsetupwin& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy lists")
        Let lngsetupwin2& = FindWindowEx(lngmdi&, 0&, "aol child", getuser$ & "'s buddy list")
        Let lnggroups& = FindWindowEx(lngsetupwin&, 0&, "_aol_listbox", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, 0&, "_aol_icon", vbNullString)
        Let lngeditbut& = FindWindowEx(lngsetupwin&, lngeditbut&, "_aol_icon", vbNullString)
    Loop Until (lngsetupwin& <> 0& Or lngsetupwin2& <> 0&) And lnggroups& <> 0& And lngeditbut& <> 0&
    If lngsetupwin& = 0& And lngsetupwin2& <> 0& Then
        Let lngsetupwin& = lngsetupwin2&
    End If
    For lngindex& = 0& To SendMessage(lnggroups&, LB_GETCOUNT, 0&, 0&) - 1&
        Call SendMessageLong(lnggroups&, LB_SETCURSEL, CLng(lngindex&), 0&)
        Let strrealgroup$ = Left$(getlistitemtext(lnggroups&, lngindex&), InStr(getlistitemtext(lnggroups&, lngindex&), Chr$(9&)) - 1&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(lngeditbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lnggroupwin& = FindWindowEx(lngmdi&, 0&, "aol child", "edit list " & strrealgroup$)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, 0&, "_aol_edit", vbNullString)
            Let lngaddbox& = FindWindowEx(lnggroupwin&, lngaddbox&, "_aol_edit", vbNullString)
            Let lngblist& = FindWindowEx(lnggroupwin&, 0&, "_aol_listbox", vbNullString)
            Let lngaddbut& = FindWindowEx(lnggroupwin&, 0&, "_aol_icon", vbNullString)
            Let lngremovebut& = FindWindowEx(lnggroupwin&, lngaddbut&, "_aol_icon", vbNullString)
            Let lngsavebut& = FindWindowEx(lnggroupwin&, lngremovebut&, "_aol_icon", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, 0&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
            Let lngstatic& = FindWindowEx(lnggroupwin&, lngstatic&, "_aol_static", vbNullString)
        Loop Until lnggroupwin& <> 0& And lngaddbox& <> 0& And lngblist <> 0& And lngaddbut <> 0& And lngremovebut& <> 0& And lngsavebut& <> 0& And lngstatic& <> 0&
        Call waitforlisttoload(lngblist&)
        For lngindex2& = 0& To SendMessage(lngblist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let addbuddiestostring$ = addbuddiestostring$ & (getlistitemtext(lngblist&, lngindex2&)) & strseperator$
        Next lngindex2&
        Call SendMessageLong(lnggroupwin&, WM_CLOSE, 0&, 0&)
    Next lngindex&
    Call SendMessageLong(lngsetupwin&, WM_CLOSE, 0&, 0&)
End Function
Public Function getcaption(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = GetWindowTextLength(lngwindow&)
    Let strbuffer$ = String$(lngtextlen&, 0&)
    Call GetWindowText(lngwindow&, strbuffer$, lngtextlen& + 1&)
    Let getcaption$ = strbuffer$
End Function
Public Function getaolversion() As Long
    Dim lngaol As Long, lngtoolwin As Long, lngtoolico As Long
    Dim lngcombo As Long, lngcomboedit As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngtoolwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolico& = FindWindowEx(lngtoolwin&, 0&, "_aol_toolbar", vbNullString)
    Let lngcombo& = FindWindowEx(lngtoolico&, 0&, "_aol_combobox", vbNullString)
    Let lngcomboedit& = FindWindowEx(lngcombo&, 0&, "edit", vbNullString)
    If lngcomboedit& = 0& Then
        Let getaolversion& = 3&
    ElseIf lngcomboedit& > 0& Then
        Let getaolversion& = 4&
    End If
End Function
Public Sub clickignore(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngignorebut As Long, lnglist As Long
    Let lngmailbox& = findmailbox2&
    Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
    Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
    Let lnglist& = FindWindowEx(lngtabwin&, 0&, "_AOL_List", vbNullString)
    Call PostMessage(lnglist&, LB_SETCURSEL, CLng(index&), 0&)
    Let lngignorebut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Let lngignorebut& = FindWindowEx(lngmailbox&, lngignorebut&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngignorebut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngignorebut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub keyword(strkeyword As String)
    Dim lngaol As Long, lngtoolwin As Long, lngtoolico As Long
    Dim lngcombo As Long, lngcomboedit As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngtoolwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolico& = FindWindowEx(lngtoolwin&, 0&, "_aol_toolbar", vbNullString)
    Let lngcombo& = FindWindowEx(lngtoolico&, 0&, "_aol_combobox", vbNullString)
    Let lngcomboedit& = FindWindowEx(lngcombo&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngcomboedit&, WM_SETTEXT, 0&, strkeyword$)
    Call SendMessageLong(lngcomboedit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(lngcomboedit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub killmailad(lngmailbox As Long)
    'ex: call killmailad(findnewmailbox&)
    Dim lngadv As Long
    Let lngadv& = FindWindowEx(lngmailbox&, 0&, "_AOL_Image", vbNullString)
    Call SendMessage(lngadv&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub killwait()
    Dim lngmodal As Long, lngbut As Long
    Call runaolmenu(4&, 10&)
    Do: DoEvents
        lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        lngbut& = FindWindowEx(lngmodal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngmodal& <> 0& And lngbut& <> 0&
    Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function gettext(lngwindow As Long) As String
    Dim strbuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strbuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_GETTEXT, lngtextlen& + 1&, strbuffer$)
    Let gettext$ = strbuffer$
End Function
Public Function getlistitemtext(lnglist As Long, index As Long) As String
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim screenname As String, itmhold As Long, psnHold As Long
    Dim rbytes As Long, cprocess As Long
    Let rlist& = lnglist&
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        Let screenname$ = String$(4, vbNullChar)
        Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        Let itmhold& = itmhold& + 24
        Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
        Call CopyMemory(psnHold&, ByVal screenname$, 4)
        Let psnHold& = psnHold& + 6
        Let screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
        Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        Let getlistitemtext$ = screenname$
        Call CloseHandle(mthread&)
    End If
End Function

Public Function findopenmail() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& = 0& Then
        Let findopenmail& = lngchild&
        Exit Function
    End If
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
        Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
        If lngstatic& <> 0& Then
            Let findopenmail& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
    Let findopenmail& = 0&
End Function
Public Function findlocatewin() As Long
    Dim lngokwin As Long, lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngchildcap As String, lnglocatewin As Long, lngstaticwin As Long
    Dim strstaticmsg As String, lngbutton As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngchildcap$ = getcaption(lngchild&)
    If LCase$(lngchildcap$) Like LCase$("locate *") Then
        Let findlocatewin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngaol& = FindWindow("AOL Frame25", "America  Online")
            Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
            Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
            Let lngchildcap$ = getcaption(lngchild&)
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
    Call keyword("aol://3548:" & strname$)
    Do: DoEvents
        Let lngokwin& = FindWindow("#32770", "America Online")
        If lngokwin& <> 0& Then
            Exit Do
        Else
            Let lngaol& = FindWindow("AOL Frame25", "America  Online")
            Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
            Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
            Let lngchildcap$ = getcaption(lngchild&)
            If LCase$(lngchildcap$) = LCase$("locate " & strname$) Then
                Let lnglocatewin& = lngchild&
            Else
                Let lngaol& = FindWindow("AOL Frame25", "America  Online")
                Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
                Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
                Do: DoEvents
                    Let lngchildcap$ = getcaption(lngchild&)
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
        Let strstaticmsg$ = gettext(lngstaticwin&)
        If LCase$(strstaticmsg$) = LCase$(strname$ & " is online, but not in a chat area.") Then
            Let locatemember$ = "online, but not in a chat room"
        ElseIf LCase$(strstaticmsg$) = LCase$(strname$ & " is online, but in a private room.") Then
            Let locatemember$ = "online, and in a private room"
        ElseIf LCase$(strstaticmsg$) Like LCase$(strname$ & " is in chat room *") Then
            Let locatemember$ = "online, and in " & Right$(strstaticmsg$, Len(strstaticmsg$) - Len(strname$) + 17&)
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
        Let locatemember$ = "not online."
    End If
End Function
Public Sub sendim(strperson As String, strmessage As String)
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngsendbut As Long
    Dim lngrich As Long, lngok As Long, lngokbut As Long, lngokwin2 As Long
    Dim lngokwin As Long, lngokbut2 As Long, lngokwin3 As Long, lngsendbut1 As Long
    Dim lngsendbut2 As Long, lngsendbut3 As Long, lngsendbut4 As Long
    Dim lngsendbut5 As Long, lngsendbut6 As Long, lngsendbut7 As Long
    Call keyword("aol://9293:" & strperson$)
    Do: DoEvents
        Let lngimwin& = FindWindowEx(lngmdi&, 0&, "aol child", "send instant message")
        Let lngrich& = FindWindowEx(lngimwin&, 0&, "richcntl", vbNullString)
    Loop Until lngimwin& <> 0& And lngrich& <> 0&
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Do: DoEvents: Loop Until gettext(lngrich&) = strmessage$
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
End Sub
Public Sub sendmailbcc(strpeople As String, strsubject As String, strmessage As String)
    Call sendmail("(" & getuser$ & "," & strpeople$ & ")", strsubject$, strmessage$)
End Sub
Public Sub sendmail(strperson As String, strsubject As String, strmessage As String)
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
        Loop Until gettext(lngcombo&) = strfilepath$
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
Public Sub cleardocuments()
    Call SHAddToRecentDocs(0&, 0&)
End Sub
Public Sub addfontstocontrol(thelist As Control)
    Dim lngindex As Long, lngsw2 As Long, lngsw As Long
    thelist.Clear
    For lngindex& = 1& To Screen.FontCount
        DoEvents
        thelist.AddItem (Screen.Fonts(lngindex& - 1&))
    Next lngindex&
End Sub
Public Sub addfontstocontrolwithdogbar(thelist As Control, progbar As Object)
    Dim lngindex As Long, lngsw2 As Long, lngsw As Long
    thelist.Clear
    Let progbar.Max = Screen.FontCount
    For lngindex& = 1& To Screen.FontCount
        DoEvents
        thelist.AddItem (Screen.Fonts(lngindex& - 1&))
        Let progbar.Value = lngindex& - 1&
    Next lngindex&
End Sub
Public Function getfontcount() As Long
    Let getfontcount& = Screen.FontCount
End Function
Public Sub loadmeonstartup()
    Call writeini("windows", "load", App.Path & "\" & App.EXEName, "c:\window\win.ini")
End Sub
Public Function isavailable(strname As String) As Boolean
    Dim lngaol As Long, lngmdi As Long, lngimwin As Long, lngrich As Long, lngokwin As Long
    Dim lngicon As Long, lngbut As Long, lngstatic As Long, strmsg As String, lngedit As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & strname$)
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
    Let strmsg$ = gettext(lngstatic&)
    Do: DoEvents
        Call PostMessage(lngbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngbut&, WM_LBUTTONUP, 0&, 0&)
        Let lngokwin& = FindWindow("#32770", "America Online")
        Let lngbut& = FindWindowEx(lngokwin&, 0&, "Button", vbNullString)
    Loop Until lngokwin& = 0& And lngbut& = 0&
    Call SendMessage(lngimwin&, WM_CLOSE, 0&, 0&)
    If LCase$(strmsg$) = LCase$(strname$) & " is online and able to receive instant messages." Then
        Let isavailable = True
    Else
        Let isavailable = False
    End If
End Function
Public Sub decompileprotect(strexepathandname As String)
    Dim strfile As String
    On Error Resume Next
    If Not InStr(strexepathandname$, "\") Then
        MsgBox "executable file not found", vbOKOnly, "izekial32.bas": Exit Sub
    Else
        Let strfile$ = FreeFile
        Open strexepathandname$ For Binary As #strfile$
            Seek #strfile$, 25&
            Put #strfile$, , "."
        Close #strfile$
        If Err Then
            MsgBox "not a visual basic made file!", vbOKOnly, "error in file": Exit Sub
        Else
            MsgBox "the file below has been decompile protected" & vbCrLf$ & strexepathandname$, vbOKOnly, "izekial32.bas"
        End If
    End If
End Sub
Public Function getwindowsversion() As String
    'taken from frenzy3.bas by izekial(me)
    Dim strstring As String, dl As Long
    Dim myver As OSVERSIONINFO, mysys As SYSTEM_INFO
    #If Win32 Then
        Let myver.dwOSVersionInfoSize = 148
        Let dl& = GetVersionEx&(myver)
        If myver.dwPlatformId = VER_PLATFORM_WIN32s Then
            Let strstring$ = "windows95 "
        ElseIf myver.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            Let strstring$ = "windowsNT "
        ElseIf myver.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            Let strstring$ = "windows98 "
        End If
        #If Win16 Then
            Let strstring$ = "windows3.x"
            Exit Function
        #End If
    #End If
    Let strstring$ = strstring$ & myver.dwMajorVersion & "." & myver.dwMinorVersion & " build " & myver.dwBuildNumber
    Let getwindowsversion$ = strstring$
End Function
Public Sub makeshortcut(strshortcutdir As String, strshortcutname As String, strshortcutpath As String)
    Dim strwinshortcutdir As String, strwinshortcutname As String, strwinshortcutexepath As String, lngretval As Long
    Let strwinshortcutdir$ = strshortcutdir$
    Let strwinshortcutname$ = strshortcutname$
    Let strwinshortcutexepath$ = strshortcutpath$
    Let lngretval& = fCreateShellLink("", strwinshortcutname$, strwinshortcutexepath$, "")
    Name "c:\windows\start menu\programs\" & strwinshortcutname$ & ".lnk" As strwinshortcutdir$ & "\" & strwinshortcutname$ & ".lnk"
End Sub
Public Function isonline() As Boolean
    Dim lngaol As Long, lngmdi As Long, lngtoolbarwin As Long, lngtoolbarico As Long
    Dim lngcomposeico As Long
    Let lngaol& = FindWindow("aol frame25", "america  online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbarwin& = FindWindowEx(lngaol&, 0&, "aol toolbar", vbNullString)
    Let lngtoolbarico& = FindWindowEx(lngtoolbarwin&, 0&, "_aol_toolbar", vbNullString)
    Let lngcomposeico& = FindWindowEx(lngtoolbarico&, 0&, "_aol_icon", vbNullString)
    Let lngcomposeico& = FindWindowEx(lngtoolbarico&, lngcomposeico&, "_aol_icon", vbNullString)
    If lngcomposeico& <> 0& Then
        Let isonline = True
    Else
        Let isonline = False
    End If
End Function
Public Function ismaster() As Boolean
    Dim lngaol As Long, lngmdi As Long, lngprntwindow As Long
    Dim lngprnticon As Long, lngmodal As Long, lngmodalstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Call keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do Until lngprntwindow& <> 0& And lngprnticon& <> 0&
        Let lngprntwindow& = FindWindowEx(lngmdi&, 0&, "AOL Child", " Parental Controls")
        Let lngprnticon& = FindWindowEx(lngprntwindow&, 0&, "_AOL_Icon", vbNullString)
    Loop: DoEvents
    Do: DoEvents
        Call PostMessage(lngprnticon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngprnticon&, WM_LBUTTONUP, 0&, 0&)
        Let lngmodal& = FindWindow("_AOL_Modal", vbNullString)
        Let lngmodalstatic& = FindWindowEx(lngmodal&, 0&, "_AOL_Static", vbNullString)
    Loop Until lngmodal& <> 0 And lngmodalstatic& <> 0& And gettext(lngmodalstatic&) <> ""
    If replacestring(replacestring(gettext(lngmodalstatic&), Chr$(10&), ""), Chr$(13&), "") = "Set Parental Controls" Then
        Let ismaster = True
    Else
        Let ismaster = False
    End If
    Call SendMessageLong(lngmodal&, WM_CLOSE, 0&, 0&): DoEvents
    Call SendMessageLong(lngprntwindow&, WM_CLOSE, 0&, 0&): DoEvents
End Function
Public Sub signoff()
    Call runaolmenu(3&, 1&)
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
    Let getimtext$ = gettext(lngrich&)
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
    Let strimcap$ = getcaption(findim&)
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
    Dim lngindex2 As Long, lngsmid As Long, strstring As String
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmmenu& = GetMenu(lngaol&)
    Let lngmmcount& = GetMenuItemCount(lngmmenu&)
    For lngindex& = 0& To lngmmcount& - 1&
        Let lngsubmenu& = GetSubMenu(lngmmenu&, lngindex&)
        Let lngsmcount& = GetMenuItemCount(lngsubmenu&)
        For lngindex2& = 0& To lngsmcount& - 1&
            Let lngsmid& = GetMenuItemID(lngsubmenu&, lngindex2&)
            Let strstring$ = String$(100, " ")
            Call GetMenuString(lngsubmenu&, lngsmid&, strstring$, 100&, 1&)
            If InStr(LCase$(strstring$), LCase$(strmenutext$)) Then
                Call SendMessageLong(lngaol&, WM_COMMAND, lngsmid&, 0&)
                Exit Sub
            End If
        Next lngindex2&
    Next lngindex&
End Sub
Public Sub loadsnlist(thelist As Control, commondlg32 As Control)
    Dim strsnhold As String, strthecomma As String, strsn As String, index As Long
    With commondlg32
        .dialogtitle = "load a sn list"
        .cancelerror = True
        .Filter = "text file (*.txt)|*.txt"
        .filterindex = 0&
        .showopen
    End With
    Open commondlg32.filename For Input As #1
        Let strsnhold$ = Input(LOF(1&), 1&)
    Close #1
    For index& = 1& To Len(strsnhold$)
        Let strthecomma$ = Mid$(strthecomma$, index&, 1)
        If strthecomma$ = "," Then
            Call thelist.AddItem(strsn$)
            Let strsn$ = ""
        Else
             Let strsn$ = strsn$ & strthecomma$
        End If
    Next index&
End Sub

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
    Loop Until lngprofwin& <> 0& And lngquote& <> 0& And lngnosex& <> 0& And lngupbut& <> 0&
    Call SendMessageByString(lngname&, WM_SETTEXT, 0&, strname$): DoEvents
    Call SendMessageByString(lnglocation&, WM_SETTEXT, 0&, strlocation$): DoEvents
    Call SendMessageByString(lngbirthday&, WM_SETTEXT, 0&, strbday$): DoEvents
    Call SendMessageByString(lngmstatus&, WM_SETTEXT, 0&, strlovestatus$): DoEvents
    Call SendMessageByString(lnghobbies&, WM_SETTEXT, 0&, strhobbies$): DoEvents
    Call SendMessageByString(lngcomputer&, WM_SETTEXT, 0&, strcomputer$): DoEvents
    Call SendMessageByString(lngquote&, WM_SETTEXT, 0&, strquote$): DoEvents
    If male = True Then
        Call PostMessage(lngmsex&, BM_SETCHECK, True, 0&)
    ElseIf female = True Then
        Call PostMessage(lngfsex&, BM_SETCHECK, True, 0&)
    ElseIf nosex = True Then
        Call PostMessage(lngnosex&, BM_SETCHECK, True, 0&)
    End If
    Call PostMessage(lngupbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngupbut&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
End Sub
Public Sub savesnlist(thelist As Control, commondlg32 As Control)
    Dim strsnhold As String, index As Long
    With commondlg32
        .cancelerror = True
        .dialogtitle = "save a sn list"
        .Filter = "text files (*.txt)|*.txt"
        .filterindex = 0&
        .ShowSave
    End With
    For index& = 0& To thelist.ListCount - 1
        If index& = 0& Then
            Let strsnhold$ = thelist.list(index&)
        Else
            Let strsnhold$ = strsnhold$ & "," & thelist.list(index&)
        End If
    Next index&
    Open commondlg32.filename For Output As #1
        Print #1, strsnhold$
    Close #1
End Sub
Public Function findbuddylist() As Long
    Dim lngaol As Long, lngmdi As Long
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let findbuddylist& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Buddy List Window")
End Function
Public Sub removelistitem(thelist As ListBox, strentry As String)
    Dim lngindex As Long
    If thelist.ListCount = 0& Then Exit Sub
    For lngindex& = 0& To thelist.ListCount - 1&
        Let thelist.ListIndex = lngindex&
        If LCase$(thelist.list(lngindex&)) = LCase$(strentry$) Then
            Call thelist.RemoveItem(lngindex&)
            Exit Sub
            If Err Then Exit Sub
        End If
    Next lngindex&
End Sub
Public Sub loadlistbox(strnameandpath As String, thelist As ListBox)
    'this won't work with all saved lists
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    Open strnameandpath$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thelist.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Sub savelistbox(strnameandpath As String, thelist As ListBox)
    Dim index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For index& = 0& To thelist.ListCount - 1&
        Print #1&, thelist.list(index&)
    Next index&
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
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, strchildcap As String
    Let lngaol& = FindWindow("AOL Frame25", "America  Online")
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let strchildcap$ = getcaption(lngchild&)
    If LCase$(strchildcap$) Like ">instant message from: *" Or LCase$(strchildcap$) Like "  instant message from: *" Or LCase$(strchildcap$) Like "  instant message to: *" Then
        Let findim& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
            Let strchildcap$ = getcaption(lngchild&)
            If LCase$(strchildcap$) Like ">instant message from: *" Or LCase$(strchildcap$) Like "  instant message from: *" Or LCase$(strchildcap$) Like "  instant message to: *" Then
                Let findim& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
    End If
    Let findim& = 0&
End Function
Public Sub aolgoback()
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
Public Sub aolgoforward()
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
Public Sub aolgostop()
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
Public Sub aolgorefresh()
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
Public Sub aolgohome()
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
    Dim index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For index& = 0& To thecombo.ListCount - 1&
        Print #1&, thecombo.list(index&)
    Next index&
    Close #1&
End Sub
Public Sub savecomboboxwithcommondialog(commondialog32 As Control, thecombo As ComboBox)
    Dim index As Long
    On Error Resume Next
    With commondialog32
        .dialogtitle = "load a listbox"
        .cancelerror = True
        .filterindex = 0&
        .showopen
    End With
    Open commondialog32.filename For Output As #1&
    For index& = 0& To thecombo.ListCount - 1&
        Print #1&, thecombo.list(index&)
    Next index&
    Close #1&
End Sub
Public Sub savelistboxwithcommondialog(commondialog32 As Control, thelist As ListBox)
    Dim index As Long
    On Error Resume Next
    With commondialog32
        .dialogtitle = "save a listbox"
        .cancelerror = True
        .filterindex = 0&
        .showopen
    End With
    Open commondialog32.filename For Output As #1&
        For index& = 0& To thelist.ListCount - 1&
            Print #1&, thelist.list(index&)
        Next index&
    Close #1&
End Sub
Public Sub loadcomboboxwithcommondialog(commondialog32 As Control, thecombo As ComboBox)
    'this won't work with all saved combos
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    With commondialog32
        .dialogtitle = "load a listbox"
        .cancelerror = True
        .filterindex = 0&
        .showopen
    End With
    Open commondialog32.filename For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thecombo.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Sub loadlistboxwithcommondialog(commondialog32 As Control, thelist As ListBox)
    'this won't work with all saved lists
    'it depends on how they were saved
    Dim strlinetext As String
    On Error Resume Next
    With commondialog32
        .dialogtitle = "load a listbox"
        .cancelerror = True
        .filterindex = 0&
        .showopen
    End With
    Open commondialog32.filename For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thelist.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Sub sendchat(strmessage As String)
    Dim lngchat As Long, lngrich As Long, strtext As String
    Let lngchat& = findroom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strtext$ = gettext(lngrich&)
    If strtext$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until gettext$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strtext$)
End Sub
Public Sub sendchat2(strmessage As String)
    'for my server [mentality server]
    Dim lngchat As Long, lngrich As Long, strtext As String
    Let lngchat& = findroom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strtext$ = gettext(lngrich&)
    If strtext$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "÷•.• " & strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until gettext$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strtext$)
End Sub
Public Sub fadeobjecttop2btm(frmname As Object, color1 As Long, color2 As Long)
    'IDEA from monkefade3
    'call this in the form_paint of your form
    'this will fade it from top to bottom
    'preset colors are in the dec's of this bas
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
    Let lngblue1& = Int(color1& / 65536)
    Let lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
    Let lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
    Let lngblue2& = Int(color2& / 65536)
    Let lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
    Let lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
    Let frmname.ScaleMode = vbPixels
    Let frmname.ScaleHeight = 256&
    Let frmname.DrawWidth = 2&
    Let frmname.DrawMode = vbCopyPen
    Let frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (0&, index&)-(Screen.Width, index& - 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub
Public Sub fadeobjectside2side2colors(theobject As Object, color1 As Long, color2 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 2&, 3&) As Integer
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
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    ReDim lngcolors2(1& To 2&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleWidth / 100&
    Let lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 2&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(2&, 1&), lngcolors2(2&, 2&), lngcolors2(2&, 3&))
    For index& = 1& To (2& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (2& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (2& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (2& - 1&))
        For index2& = 1& To (100& / (2& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeobjectside2side3colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 3&, 3&) As Integer
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strcolor3$ = gethexfromrgb(color3&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let strred3$ = "&h" & Right$(strcolor3$, 2&)
    Let strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
    Let strblue3$ = "&h" & Left$(strcolor3$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Let lngred3& = Val(strred3$)
    Let lnggreen3& = Val(strgreen3$)
    Let lngblue3& = Val(strblue3$)
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    Let lngcolors(3&, 1&) = lngred3&
    Let lngcolors(3&, 2&) = lnggreen3&
    Let lngcolors(3&, 3&) = lngblue3&
    ReDim lngcolors2(1& To 3&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleWidth / 100&
    Let lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 3&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(3&, 1&), lngcolors2(3&, 2&), lngcolors2(3&, 3&))
    For index& = 1& To (3& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (3& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (3& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (3& - 1&))
        For index2& = 1& To (100& / (3& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeobjectside2side4colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long, color4 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    Dim strcolor4 As String, strred4 As String, strblue4 As String, strgreen4 As String
    Dim lngred4 As Long, lngblue4 As Long, lnggreen4 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 4&, 3&) As Integer
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strcolor3$ = gethexfromrgb(color3&)
    Let strcolor4$ = gethexfromrgb(color4&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let strred3$ = "&h" & Right$(strcolor3$, 2&)
    Let strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
    Let strblue3$ = "&h" & Left$(strcolor3$, 2&)
    Let strred4$ = "&h" & Right$(strcolor4$, 2&)
    Let strgreen4$ = "&h" & Mid$(strcolor4$, 3&, 2&)
    Let strblue4$ = "&h" & Left$(strcolor4$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Let lngred3& = Val(strred3$)
    Let lnggreen3& = Val(strgreen3$)
    Let lngblue3& = Val(strblue3$)
    Let lngred4& = Val(strred4$)
    Let lnggreen4& = Val(strgreen4$)
    Let lngblue4& = Val(strblue4$)
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    Let lngcolors(3&, 1&) = lngred3&
    Let lngcolors(3&, 2&) = lnggreen3&
    Let lngcolors(3&, 3&) = lngblue3&
    Let lngcolors(4&, 1&) = lngred4&
    Let lngcolors(4&, 2&) = lnggreen4&
    Let lngcolors(4&, 3&) = lngblue4&
    ReDim lngcolors2(1& To 4&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleWidth / 100&
    Let lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 4&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(4&, 1&), lngcolors2(4&, 2&), lngcolors2(4&, 3&))
    For index& = 1& To (4& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (4& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (4& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (4& - 1&))
        For index2& = 1& To (100& / (4& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeobjecttop2bottom2colors(theobject As Object, color1 As Long, color2 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 2&, 3&) As Integer
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
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    ReDim lngcolors2(1& To 2&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleHeight / 100&
    Let lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 2&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(2&, 1&), lngcolors2(2&, 2&), lngcolors2(2&, 3&))
    For index& = 1& To (2& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (2& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (2& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (2& - 1&))
        For index2& = 1 To (100& / (2& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeobjecttop2bottom3colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 3&, 3&) As Integer
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strcolor3$ = gethexfromrgb(color3&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let strred3$ = "&h" & Right$(strcolor3$, 2&)
    Let strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
    Let strblue3$ = "&h" & Left$(strcolor3$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Let lngred3& = Val(strred3$)
    Let lnggreen3& = Val(strgreen3$)
    Let lngblue3& = Val(strblue3$)
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    Let lngcolors(3&, 1&) = lngred3&
    Let lngcolors(3&, 2&) = lnggreen3&
    Let lngcolors(3&, 3&) = lngblue3&
    ReDim lngcolors2(1& To 3&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleHeight / 100&
    Let lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 3&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(3&, 1&), lngcolors2(3&, 2&), lngcolors2(3&, 3&))
    For index& = 1& To (3& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (3& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (3& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (3& - 1&))
        For index2& = 1& To (100& / (3& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Public Sub fadeobjecttop2bottom4colors(theobject As Object, color1 As Long, color2 As Long, color3 As Long, color4 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    Dim strcolor4 As String, strred4 As String, strblue4 As String, strgreen4 As String
    Dim lngred4 As Long, lngblue4 As Long, lnggreen4 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 4&, 3&) As Integer
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strcolor3$ = gethexfromrgb(color3&)
    Let strcolor4$ = gethexfromrgb(color4&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let strred3$ = "&h" & Right$(strcolor3$, 2&)
    Let strgreen3$ = "&h" & Mid$(strcolor3$, 3&, 2&)
    Let strblue3$ = "&h" & Left$(strcolor3$, 2&)
    Let strred4$ = "&h" & Right$(strcolor4$, 2&)
    Let strgreen4$ = "&h" & Mid$(strcolor4$, 3&, 2&)
    Let strblue4$ = "&h" & Left$(strcolor4$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Let lngred3& = Val(strred3$)
    Let lnggreen3& = Val(strgreen3$)
    Let lngblue3& = Val(strblue3$)
    Let lngred4& = Val(strred4$)
    Let lnggreen4& = Val(strgreen4$)
    Let lngblue4& = Val(strblue4$)
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    Let lngcolors(3&, 1&) = lngred3&
    Let lngcolors(3&, 2&) = lnggreen3&
    Let lngcolors(3&, 3&) = lngblue3&
    Let lngcolors(4&, 1&) = lngred4&
    Let lngcolors(4&, 2&) = lnggreen4&
    Let lngcolors(4&, 3&) = lngblue4&
    ReDim lngcolors2(1& To 4&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleHeight / 100&
    Let lngswidth& = theobject.ScaleWidth
    theobject.Cls
    For index& = 1& To 4&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(4&, 1&), lngcolors2(4&, 2&), lngcolors2(4&, 3&))
    For index& = 1& To (4& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (4& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (4& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (4& - 1&))
        For index2& = 1& To (100& / (4& - 1&))
            theobject.Line (0&, lnghold)-(lngswidth&, lnghold + obj100th), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub
Public Sub fadeformdiagonallyslow(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it diagonally
    'preset colors are in the dec's of this bas
    
    'works with most sized forms
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
    Let lngblue1& = Int(color1& / 65536)
    Let lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
    Let lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
    Let lngblue2& = Int(color2& / 65536)
    Let lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
    Let lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
    Let frmname.ScaleMode = vbPixels
    Let frmname.ScaleWidth = 256&
    Let frmname.DrawWidth = 2&
    Let frmname.DrawMode = vbCopyPen
    Let frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(Screen.Height, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(index& + 1&, Screen.Width), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub
Public Sub fadeformdiagonallyfast(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it from top to bottom
    'preset colors are in the dec's of this bas
    
    'works with most size forms
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
    Let lngblue1& = Int(color1& / 65536)
    Let lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
    Let lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
    Let lngblue2& = Int(color2& / 65536)
    Let lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
    Let lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
    Let frmname.ScaleMode = vbPixels
    Let frmname.ScaleWidth = 256&
    Let frmname.DrawWidth = 2&
    Let frmname.DrawMode = vbCopyPen
    Let frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(Screen.Height, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(Screen.Width, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
End Sub
Public Sub fadeformdiagonallywithborder(frmname As Form, color1 As Long, color2 As Long)
    'call this in the form_paint of your form
    'this will fade it from top to bottom
    'preset colors are in the dec's of this bas
    Dim index As Long, lngblue1 As Long, lnggreen1 As Long, lngred1 As Long
    Dim lngblue2 As Long, lnggreen2 As Long, lngred2 As Long
    Let lngblue1& = Int(color1& / 65536)
    Let lnggreen1& = Int((color1& - (65536 * lngblue1&)) / 256)
    Let lngred1& = color1& - (65536 * lngblue1& + 256& * lnggreen1&)
    Let lngblue2& = Int(color2& / 65536)
    Let lnggreen2& = Int((color2& - (65536 * lngblue2&)) / 256)
    Let lngred2& = color2& - (65536 * lngblue2& + 256& * lnggreen2)
    Let frmname.ScaleMode = vbPixels
    Let frmname.ScaleWidth = 256&
    Let frmname.DrawWidth = 2&
    Let frmname.DrawMode = vbCopyPen
    Let frmname.DrawStyle = vbInsideSolid
    For index& = 0& To 255&
        frmname.Line (index&, 0&)-(5, index& + 1&), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
        frmname.Line (0&, index&)-(Screen.Width, 5), RGB(((lngred2& - lngred1&) / 255& * index&) + lngred1&, ((lnggreen2& - lnggreen2&) / 255& * index&) + lnggreen1&, ((lngblue2& - lngblue1&) / 255& * index&) + lngblue1&), B
    Next index&
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
Public Function fadetext2bybars(Red1 As Long, Green1 As Long, Blue1 As Long, red2 As Long, green2 As Long, blue2 As Long, strtext As String, wavyonoff As Boolean) As String
    'IDEA from monkefade3
    'this will fade text by 2 colors
    'using the value of scroll bars
    Dim lnglen As Long, strcurstring As String, strlastchar As String
    Dim lngrgbcolor As Long, lnghexcolor As String, wherewavy As Long
    Dim wavystring As String, stroutput As String, index As Long
    Let lnglen& = Len(strtext$)
    For index& = 1& To lnglen&
        Let strcurstring$ = Left$(strtext$, index&)
        Let strlastchar$ = Right$(strcurstring$, 1&)
        Let lngrgbcolor& = RGB(((blue2& - Blue1&) / lnglen& * index&) + Blue1&, ((green2& - Green1&) / lnglen& * index&) + Green1&, ((red2& - Red1&) / lnglen& * index&) + Red1&)
        Let lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
            Let wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                    Let wherewavy& = 1&
                Case 1&
                    Let wavystring$ = "<sup>"
                Case 2&
                    Let wavystring$ = "</sup>"
                Case 3&
                    Let wavystring$ = "<sub>"
                Case 4&
                    Let wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
            Let wavystring$ = ""
        End If
        Let stroutput$ = stroutput$ & "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$
    Next index&
    Let fadetext2bybars$ = stroutput$
End Function
Public Function fadetext2byrgb(color1 As Long, color2 As Long, strtext As String, wavyonoff As Boolean) As String
    'IDEA from monkefade3
    'this will fade text, using 2 colors
    'you must supply the sub with the two colors
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long
    If strtext$ = "" Then
        Exit Function
    Else
        Let strcolor1$ = gethexfromrgb(color1&)
        Let strcolor2$ = gethexfromrgb(color2&)
        Let strred1$ = "&H" & Right$(strcolor1$, 2&)
        Let strgreen1$ = "&H" & Mid$(strcolor1$, 3&, 2&)
        Let strblue1$ = "&H" & Left$(strcolor1$, 2&)
        Let strred2$ = "&H" & Right$(strcolor2$, 2&)
        Let strgreen2$ = "&H" & Mid$(strcolor2$, 3&, 2&)
        Let strblue2$ = "&H" & Left$(strcolor2$, 2&)
        Let lngred1& = Val(strred1$)
        Let lnggreen1& = Val(strgreen1$)
        Let lngblue1& = Val(strblue1$)
        Let lngred2& = Val(strred2$)
        Let lnggreen2& = Val(strgreen2$)
        Let lngblue2& = Val(strblue2$)
        Let fadetext2byrgb$ = fadetext2bybars(lngred1&, lnggreen1&, lngblue1&, lngred2&, lnggreen2&, lngblue2&, strtext$, wavyonoff)
    End If
End Function
Public Function fadetext3bybars(Red1 As Long, Green1 As Long, Blue1 As Long, red2 As Long, green2 As Long, blue2 As Long, red3 As Long, green3 As Long, blue3 As Long, strtext As String, wavyonoff As Boolean)
    'IDEA from monkefade3
    'this will fade text by 3 colors
    'using the value of scroll bars
    Dim lnglen As Long, strfirsthalf As String, strsecondhalf As String, index As Long
    Dim strcurstring As String, strlastchar As String, lngrgbcolor As Long, lnghexcolor As String
    Dim stroutput2 As String, stroutput1 As String, wherewavy As Long, wavystring As String
    Let lnglen& = Len(strtext$)
    Let strfirsthalf$ = Left$(strtext$, Int(lnglen&) / 2&)
    Let strsecondhalf$ = Right$(strtext$, lnglen& - Int(lnglen&) / 2&)
    For index& = 1& To Len(strfirsthalf$)
        Let strcurstring$ = Left$(strfirsthalf$, index&)
        Let strlastchar$ = Right$(strcurstring$, 1&)
        Let lngrgbcolor& = RGB(((blue2& - Blue1&) / Len(strfirsthalf$) * index&) + Blue1&, ((green2& - Green1&) / Len(strfirsthalf$) * index&) + Green1&, ((red2& - Red1&) / Len(strfirsthalf$) * index&) + Red1&)
        Let lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
            Let wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                    Let wherewavy& = 1&
                Case 1&
                    Let wavystring$ = "<sup>"
                Case 2&
                    Let wavystring$ = "</sup>"
                Case 3&
                    Let wavystring$ = "<sub>"
                Case 4&
                    Let wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
            Let wavystring$ = ""
        End If
        Let stroutput1$ = stroutput1$ & "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$
    Next index&
    For index& = 1 To Len(strsecondhalf$)
        Let strcurstring$ = Left$(strsecondhalf$, index&)
        Let strlastchar$ = Right$(strcurstring$, 1&)
        Let lngrgbcolor& = RGB(((blue2& - Blue1&) / Len(strsecondhalf$) * index&) + Blue1&, ((green2& - Green1&) / Len(strsecondhalf$) * index&) + Green1&, ((red2& - Red1&) / Len(strsecondhalf$) * index&) + Red1&)
        Let lnghexcolor$ = gethexfromrgb(lngrgbcolor&)
        If wavyonoff = True Then
            Let wherewavy& = wherewavy& + 1&
            Select Case wherewavy
                Case Is > 4&
                    Let wherewavy& = 1&
                Case 1&
                    Let wavystring$ = "<sup>"
                Case 2&
                    Let wavystring$ = "</sup>"
                Case 3&
                    Let wavystring$ = "<sub>"
                Case 4&
                    Let wavystring$ = "</sub>"
                Case Else
                    Exit Function
            End Select
        Else
            Let wavystring$ = ""
        End If
        Let stroutput2$ = "<Font Color=#" & lnghexcolor$ & ">" & wavystring$ & strlastchar$ & stroutput2$
    Next index&
    Let fadetext3bybars = stroutput1$ & stroutput2$
End Function
Public Function fadetext3byrgb(color1 As Long, color2 As Long, color3 As Long, strtext As String, wavyonoff As Boolean) As String
    'this will fade text, using 3 colors
    'you must supply the sub with the three colors
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    If strtext$ = "" Then
        Exit Function
    Else
        Let strcolor1$ = gethexfromrgb(color1&)
        Let strcolor2$ = gethexfromrgb(color2&)
        Let strcolor3$ = gethexfromrgb(color3&)
        Let strred1$ = "&H" & Right$(strcolor1$, 2&)
        Let strgreen1$ = "&H" & Mid$(strcolor1$, 3&, 2&)
        Let strblue1$ = "&H" & Left$(strcolor1$, 2&)
        Let strred2$ = "&H" & Right$(strcolor2$, 2&)
        Let strgreen2$ = "&H" & Mid$(strcolor2$, 3&, 2&)
        Let strblue2$ = "&H" & Left$(strcolor2$, 2&)
        Let strred3$ = "&H" & Right$(strcolor3$, 2&)
        Let strgreen3$ = "&H" & Mid$(strcolor3$, 3&, 2&)
        Let strblue3$ = "&H" & Left$(strcolor3$, 2&)
        Let lngred1& = Val(strred1$)
        Let lnggreen1& = Val(strgreen1$)
        Let lngblue1& = Val(strblue1$)
        Let lngred2& = Val(strred2$)
        Let lnggreen2& = Val(strgreen2$)
        Let lngblue2& = Val(strblue2$)
        Let fadetext3byrgb$ = fadetext3bybars(lngred1&, lnggreen1&, lngblue1&, lngred2&, lnggreen2&, lngblue2&, lngred3&, lnggreen3&, lngblue3&, strtext$, wavyonoff)
    End If
End Function
Public Sub fadepicbox(picbox As Object, color1 As Long, color2 As Long)
    'this goes in the _paint event of
    'the picture box
    Dim lngcon As Long, longcon As Long, lnghlfwidth As Long, lngcolorval1 As Long
    Dim lngcolorval2 As Long, lngcolorval3 As Long, lngrgb1 As Long, lngrgb2 As Long
    Dim lngrgb3 As Long, lngyval As Long, strcolor1 As String, strcolor2 As String
    Dim strred1 As String, strgreen1 As String, strblue1 As String, strred2 As String
    Dim strgreen2 As String, strblue2 As String, lngred1 As Long, lnggreen1 As Long
    Dim lngblue1 As Long, lngred2 As Long, lnggreen2 As Long, lngblue2 As Long
    Let picbox.AutoRedraw = True
    Let picbox.DrawStyle = 6&
    Let picbox.DrawMode = 13&
    Let picbox.DrawWidth = 2&
    Let lngcon& = 0&
    Let lnghlfwidth& = picbox.Width / 2&
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
        picbox.Line (lngyval&, 0&)-(lngyval& + 2&, picbox.Height), RGB(lngrgb1&, lngrgb2&, lngrgb3&), BF
        Let lngyval& = lngyval& + 10&
        Let lngcon& = lngcon& + 5&
    Loop Until lngcon& > lnghlfwidth&
End Sub
Public Sub fadeobjectcircle(theobject As Object, blue As Boolean, red As Boolean, green As Boolean)
    'IDEA from fireball
    Dim lngwidth As Long, lngheight As Long, lngblueval As Long, lngobjwidth As Long, lngredval As Long, lnggreenval As Long
    Let theobject.FillStyle = 0&
    Let lngwidth& = theobject.Width
    Let lngheight& = theobject.Height
    If blue = True Then
        Do Until lngblueval& = 255&: DoEvents
            Let lngblueval& = lngblueval& + 1&
            Let lngwidth& = lngwidth& - theobject.Width / 255&
            Let theobject.FillColor = RGB(0&, 0&, lngblueval&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(0&, 0&, lngblueval&)
            End If
        Loop
    ElseIf red = True Then
        Do Until lngredval& = 255&: DoEvents
            Let lngredval& = lngredval& + 1&
            Let lngwidth& = lngwidth& - theobject.Width / 255&
            Let theobject.FillColor = RGB(lngredval&, 0&, 0&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(lngredval&, 0&, 0&)
            End If
        Loop
    ElseIf green = True Then
        Do Until lnggreenval& = 255&: DoEvents
            Let lnggreenval& = lnggreenval& + 1&
            Let lngwidth& = lngwidth& - theobject.Width / 255&
            Let theobject.FillColor = RGB(0&, lnggreenval&, 0&)
            If lngwidth& < 0& Then
                Exit Sub
            Else
                theobject.Circle (theobject.Width / 2&, theobject.Height / 2&), lngwidth&, RGB(0&, lnggreenval&, 0&)
            End If
        Loop
    End If
End Sub
Public Function get3colorbarsvalue(redbar As Control, greenbar As Control, bluebar As Control) As Long
    get3colorbarsvalue& = RGB(redbar.Value, greenbar.Value, bluebar.Value)
End Function
Public Sub formtotray(frmform As Form)
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .uId = vbNull
        Let .hWnd = frmform.hWnd
        Let .ucallbackMessage = WM_MOUSEMOVE
        Let .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        Let .hIcon = frmform.Icon
        Let .szTip = frmform.Caption
    End With
    Call Shell_NotifyIcon(NIM_ADD, systray)
    frmform.Hide
End Sub
Public Sub formfromtray(frmform As Form)
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .hWnd = frmform.hWnd
        Let .uId = vbNull
    End With
    Call Shell_NotifyIcon(NIM_DELETE, systray)
    frmform.Show
End Sub
Public Function findroom() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngrich As Long
    Dim lnglist As Long, lngsendbut As Long, lngstatic As Long
    Let lngaol& = FindWindow("aol frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", vbNullString)
    Let lngrich& = FindWindowEx(lngchild&, 0&, "richcntl", vbNullString)
    Let lnglist& = FindWindowEx(lngchild&, 0&, "_aol_listbox", vbNullString)
    Let lngsendbut& = FindWindowEx(lngchild&, 0&, "_aol_icon", vbNullString)
    Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
    Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
    Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
    Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_aol_static", vbNullString)
    If lngrich& <> 0& And lnglist& <> 0& And lngsendbut& <> 0& And lngstatic& <> 0& Then
        Let findroom& = lngchild&
        Exit Function
    End If
    Do: DoEvents
        Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "aol child", vbNullString)
        Let lngrich& = FindWindowEx(lngchild&, 0&, "richcntl", vbNullString)
        Let lnglist& = FindWindowEx(lngchild&, 0&, "_aol_listbox", vbNullString)
        Let lngsendbut& = FindWindowEx(lngchild&, 0&, "_aol_icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngsendbut& = FindWindowEx(lngchild&, lngsendbut&, "_aol_icon", vbNullString)
        Let lngstatic& = FindWindowEx(lngchild&, 0&, "_aol_static", vbNullString)
        If lngrich& <> 0& And lnglist& <> 0& And lngsendbut& <> 0& And lngstatic& <> 0& Then
            Let findroom& = lngchild&
            Exit Function
        End If
    Loop Until lngchild& = 0&
    Let findroom& = lngchild&
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
Public Function findmailbox2() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long
    Dim lngtabwin As Long, lngtabwin2 As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let findmailbox2& = FindWindowEx(lngmdi&, 0&, "AOL Child", getuser$ & "'s Online Mailbox")
End Function
Public Sub clickdownloadlater(lngmailwin As Long)
    Dim lngdlbutton As Long
    Let lngdlbutton& = FindWindowEx(lngmailwin&, 0&, "_AOL_Icon", vbNullString)
    Let lngdlbutton& = FindWindowEx(lngmailwin&, lngdlbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(lngdlbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngdlbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
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
Public Function getstringlineindex(strstring As String, strlinetext As String) As Long
    Dim index As Long
    For index& = 0& To getstringlinecount(strstring$)
        If getlinefromstring(strstring$, index&) = strlinetext$ Then
            Let getstringlineindex& = index&
            Exit Function
        End If
    Next index&
End Function
Public Function findforwardwin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long, strfwdcap As String
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    Let strfwdcap$ = Left$(getcaption(lngchild&), 4&)
    If strfwdcap$ = "Fwd:" And lngstatic& <> 0& Then
        Let findforwardwin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            Let strfwdcap$ = Left$(getcaption(lngchild&), 4&)
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
    Let strfwdcap$ = Left$(getcaption(lngchild&), 3&)
    If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
        Let findreplywin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            Let strfwdcap$ = Left$(getcaption(lngchild&), 3&)
            If strfwdcap$ = "Re:" And lngstatic& <> 0& Then
                Let findreplywin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
        Let findreplywin& = lngchild&
    End If
End Function
Public Sub startdownloadlater()
    Dim lngaol As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim cursorpos As POINTAPI, lngmdi As Long, lngtree As Long
    Dim lngdlmanager As Long, lngicon1 As Long
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
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyD, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyD, 0&)
    Do: DoEvents
        lngdlmanager& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Download Manager")
        lngicon1& = FindWindowEx(lngdlmanager&, 0&, "_AOL_Icon", vbNullString)
        lngicon& = FindWindowEx(lngdlmanager&, lngicon1&, "_AOL_Icon", vbNullString)
    Loop Until lngdlmanager& <> 0& And lngicon& <> 0& And lngicon1 <> 0& And lngicon1& <> lngicon&
    Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(lngdlmanager&, WM_CLOSE, 0&, 0&)
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
End Sub

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
Public Sub addflashmailtocontrol(thelist As Control)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    Dim strstring As String, index As Long, lnglen As Long, thetab As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let strstring$ = String$(255&, 0&)
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&: DoEvents
            Let lnglen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
            Let strstring$ = String(lnglen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strstring$)
            Let thetab& = InStr(strstring$, Chr$(9&))
            Let thetab& = InStr(thetab& + 1&, strstring$, Chr$(9&))
            Let strstring$ = Right$(strstring$, Len(strstring$) - thetab&)
            Let strstring$ = replacestring(strstring$, Chr$(0&), "")
            thelist.AddItem strstring$
        Next index&
    End If
End Sub
Public Sub addflashmailtocontrolwithdogbar(thelist As Control, progbar As Control)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lngtree As Long
    Dim strstring As String, index As Long, lnglen As Long, thetab As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let strstring$ = String$(255&, 0&)
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtree& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
        Let progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
        For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&: DoEvents
            Let progbar.Value = index& - 1&
            Let lnglen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
            Let strstring$ = String(lnglen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strstring$)
            Let thetab& = InStr(strstring$, Chr$(9&))
            Let thetab& = InStr(thetab& + 1&, strstring$, Chr$(9&))
            Let strstring$ = Right$(strstring$, Len(strstring$) - thetab&)
            Let strstring$ = replacestring(strstring$, Chr$(0&), "")
            thelist.AddItem strstring$
        Next index&
    End If
End Sub
Public Sub addnewmailtocontrol(thelist As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, index As Long
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
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Sub addnewmailtocontrolwithdogbar(thelist As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, index As Long
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
            Let progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let progbar.Value = index& - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Sub addoldmailtocontrol(thelist As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
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
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Sub addoldmailtocontrolwithdogbar(thelist As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
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
            Let progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let progbar.Value = index& - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Sub addsentmailtocontrol(thelist As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
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
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Sub addsentmailtocontrolwithdogbar(thelist As Control, progbar As Control)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, index As Long
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
            Let progbar.Max = SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1&
                Let progbar.Value = index& - 1&
                Let lngtextlength& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
                Let strmailtext$ = String$(lngtextlength& + 1&, 0&)
                Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
                Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
                Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
                Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
                thelist.AddItem strmailtext$
                DoEvents
            Next index&
        End If
    End If
End Sub
Public Function findsendwin() As Long
    Dim lngaol As Long, lngmdi As Long, lngchild As Long, lngstatic As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngchild& = FindWindowEx(lngmdi&, 0&, "AOL Child", vbNullString)
    Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
    If lngstatic& <> 0& Then
        Let findsendwin& = lngchild&
        Exit Function
    Else
        Do: DoEvents
            Let lngchild& = FindWindowEx(lngmdi&, lngchild&, "AOL Child", vbNullString)
            Let lngstatic& = FindWindowEx(lngchild&, 0&, "_AOL_Static", "Send Now")
            If lngstatic& <> 0& Then
                Let findsendwin& = lngchild&
                Exit Function
            End If
        Loop Until lngchild& = 0&
    End If
    Let findsendwin& = 0&
End Function

Public Sub printtext(Text As String)
    Dim lngoldcursor As Long
    Let lngoldcursor& = Screen.MousePointer
    Let Screen.MousePointer = 11&
    Printer.Print (Text$)
    Printer.NewPage
    Printer.EndDoc
    Let Screen.MousePointer = lngoldcursor&
End Sub
Public Sub pause(length As Long)
    Dim current As Long
    Let current& = Timer
    Do Until (Timer - current&) >= length&
        DoEvents
    Loop
End Sub
Public Sub clickicondouble(lnghwnd As Long)
    Call SendMessage(lnghwnd&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub
Public Sub clicklistindex(lnglist As Long, lngindex As Long)
    Call SendMessage(lnglist&, LB_SETCURSEL, CLng(lngindex&), 0&)
End Sub
Public Sub deletenewmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long, lngbutton As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
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
Public Sub deleteoldmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long, lngtree As Long, lngbutton As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
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
Public Sub deletesentmailitem(index As Long)
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
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Sub
        Else
            Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
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
    Dim index As Long
    For index& = 0& To getnewmailcount&
        Call deletenewmailitem(index&)
    Next index&
End Sub
Public Sub deleteallnewmailwithdogbar(progbar As Control)
    Dim index As Long
    Let progbar.Max = getnewmailcount& - 1&
    For index& = 0& To getnewmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call deletenewmailitem(index&)
    Next index&
End Sub
Public Sub deletealloldmail()
    Dim index As Long
    For index& = 0& To getoldmailcount& - 1&
        Call deleteoldmailitem(index&)
    Next index&
End Sub
Public Sub deletealloldmailwithdogbar(progbar As Control)
    Dim index As Long
    Let progbar.Max = getoldmailcount& - 1&
    For index& = 0& To getoldmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call deleteoldmailitem(index&)
    Next index&
End Sub
Public Sub deleteallsentmail()
    Dim index As Long
    For index& = 0& To getsentmailcount& - 1&
        Call deletesentmailitem(index&)
    Next index&
End Sub
Public Sub deleteallsentmailwithdogbar(progbar As Control)
    Dim index As Long
    Let progbar.Max = getsentmailcount& - 1&
    For index& = 0& To getsentmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call deletesentmailitem(index&)
    Next index&
End Sub
Public Sub deleteallflashmail()
    Dim index As Long
    For index& = 0& To getflashmailcount& - 1&
        Call deleteflashmailitem(index&)
    Next index&
End Sub
Public Sub deleteallflashmailwithdogbar(progbar As Control)
    Dim index As Long
    Let progbar.Max = getflashmailcount& - 1&
    For index& = 0& To getflashmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call deleteflashmailitem(index&)
    Next index&
End Sub
Public Sub deleteflashmailitem(index As Long)
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngdelbut As Long, lngokwin As Long, lngokbut As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDICLIENT", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& < index& Then
        Exit Sub
    Else
        Let lngdelbut& = FindWindowEx(lngmailbox&, 0&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Let lngdelbut& = FindWindowEx(lngmailbox&, lngdelbut&, "_AOL_Icon", vbNullString)
        Call SendMessage(lnglist&, LB_SETCURSEL, index&, 0&)
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
Public Sub formontop(frmform As Form, ontop As Boolean)
    If ontop = True Then Call SetWindowPos(frmform.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
    If ontop = False Then Call SetWindowPos(frmform.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Public Sub sendtext(lngwindow As Long, strtext As String)
    Call SendMessageByString(lngwindow&, WM_SETTEXT, 0&, strtext$)
End Sub
Public Sub windowmaximize(lnghwnd As Long)
    Call ShowWindow(lnghwnd, SW_MAXIMIZE)
End Sub
Public Sub imignore(strwho As String, ignore As Boolean)
    Dim strstate As String
    If ignore = True Then
        Let strstate$ = "off"
    Else
        Let strstate$ = "on"
    End If
    Call sendim("$im_" & strstate$ & " " & strwho$, " ")
End Sub
Public Sub apispywithtextboxs(winhdl As TextBox, winclass As TextBox, wintxt As TextBox, winstyle As TextBox, winidnum As TextBox, winphandle As TextBox, winptext As TextBox, winpclass As TextBox, winmodule As TextBox)
    'taken from frenzy32, by izekial(me)
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
    Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
    Dim sparentclassname As String * 100, wid As Long, lwindowstyle As Long
    Dim hInstance As Long, sparentwindowtext As String * 100
    Dim smodulefilename As String * 100, r As Long
    Static hwndlast As Long
        Call GetCursorPos(pt32)
        Let ptx = pt32.X
        Let pty = pt32.Y
        Let hwndover = WindowFromPointXY(ptx, pty)
        If hwndover <> hwndlast Then
            Let hwndlast = hwndover
            Let winhdl.Text = "window handle: " & hwndover
            Let r = GetWindowText(hwndover, swindowtext, 100)
            Let wintxt.Text = "window text: " & Left(swindowtext, r)
            Let r = GetClassName(hwndover, sclassname, 100)
            Let winclass.Text = "window class name: " & Left(sclassname, r)
            Let lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
            Let winstyle.Text = "window style: " & lwindowstyle
            Let hWndParent = GetParent(hwndover)
                If hWndParent <> 0 Then
                    Let wid = GetWindowWord(hwndover, GWW_ID)
                    Let winidnum.Text = "window id number: " & wid
                    Let winphandle.Text = "parent window handle: " & hWndParent
                    Let r = GetWindowText(hWndParent, sparentwindowtext, 100)
                    Let winptext.Text = "parent window text: " & Left(sparentwindowtext, r)
                    Let r = GetClassName(hWndParent, sparentclassname, 100)
                    Let winpclass.Text = "parent window class name: " & Left(sparentclassname, r)
                Else
                    Let winidnum.Text = "window id number: n/a"
                    Let winphandle.Text = "parent window handle: n/a"
                    Let winptext.Text = "parent window text : n/a"
                    Let winpclass.Text = "parent window class name: n/a"
                End If
                    Let hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                    Let r = GetModuleFileName(hInstance, smodulefilename, 100)
            Let winmodule.Text = "module: " & Left(smodulefilename, r)
        End If
End Sub

Public Sub apispywithlabels(winhdl As Label, winclass As Label, wintxt As Label, winstyle As Label, winidnum As Label, winphandle As Label, winptext As Label, winpclass As Label, winmodule As Label)
    'taken from frenzy32, by izekial(me)
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
    Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
    Dim sparentclassname As String * 100, wid As Long, lwindowstyle As Long
    Dim hInstance As Long, sparentwindowtext As String * 100
    Dim smodulefilename As String * 100, r As Long
    Static hwndlast As Long
        Call GetCursorPos(pt32)
        Let ptx = pt32.X
        Let pty = pt32.Y
        Let hwndover = WindowFromPointXY(ptx, pty)
        If hwndover <> hwndlast Then
            Let hwndlast = hwndover
            Let winhdl.Caption = "window handle: " & hwndover
            Let r = GetWindowText(hwndover, swindowtext, 100)
            If Left(swindowtext, r) = "" Then
                Let swindowtext = "n/a"
                Let r = Len(swindowtext)
            End If
            Let wintxt.Caption = "window text: " & Left(swindowtext, r)
            Let r = GetClassName(hwndover, sclassname, 100)
            Let winclass.Caption = "window class: " & Left(sclassname, r)
            Let lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
            Let winstyle.Caption = "window style: " & lwindowstyle
            Let hWndParent = GetParent(hwndover)
                If hWndParent <> 0 Then
                    Let wid = GetWindowWord(hwndover, GWW_ID)
                    Let winidnum.Caption = "window id number: " & wid
                    Let winphandle.Caption = "parent win handle: " & hWndParent
                    Let r = GetWindowText(hWndParent, sparentwindowtext, 100)
                    Let winptext.Caption = "parent win text: " & Left(sparentwindowtext, r)
                    Let r = GetClassName(hWndParent, sparentclassname, 100)
                    Let winpclass.Caption = "parent win class: " & Left(sparentclassname, r)
                Else
                    Let winidnum.Caption = "window id number: n/a"
                    Let winphandle.Caption = "parent win handle: n/a"
                    Let winptext.Caption = "parent win text : n/a"
                    Let winpclass.Caption = "parent win class: n/a"
                End If
                    Let hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                    Let r = GetModuleFileName(hInstance, smodulefilename, 100)
            Let winmodule.Caption = "module: " & Left(smodulefilename, r)
        End If
End Sub
Public Function doeswindowexist(lngwindow As Long) As Boolean
    If IsWindow(lngwindow&) Then
        Let doeswindowexist = True
    Else
        Let doeswindowexist = False
    End If
End Function
Public Function getdiskinfo(thedrive As String) As String
    'from frenzy3.bas by izekial(me)
    Dim dl&, S$, spaceloc%, freebytes&, totalbytes&
    Dim sectorspercluster&, bytespersector&, numberoffreeclustors&
    Dim bytesfree&, bytestotal&, percentfree&, totalnumberofclustors&
    Dim tmp1$, tmp2$, tmp3$, tmp4$, tmp5$, tmp6$, tmp7$
    If Right(thedrive$, 1) = "\" Then
        Let thedrive$ = Left(thedrive$, Len(thedrive$) - 1)
    End If
    If InStr(thedrive$, ":") Then
        Let thedrive$ = thedrive$
    Else
        Let thedrive$ = thedrive$ & ":"
    End If
    Let dl& = GetDiskFreeSpace(thedrive$, sectorspercluster, bytespersector, numberoffreeclustors, totalnumberofclustors)
    Let tmp1$ = "sectors per cluster :" & Format(sectorspercluster, "#,0")
    Let tmp2$ = "bytes per sector : " & Format(bytespersector, "#,0")
    Let tmp3$ = "number of free clusters : " & Format(numberoffreeclustors, "#,0")
    Let tmp4$ = "total number of clustors : " & Format(totalnumberofclustors, "#,0")
    Let tmp5$ = totalnumberofclustors * sectorspercluster * bytespersector
    Let tmp5$ = "total bytes : " & Format(tmp5$, "#,0")
    Let tmp6$ = numberoffreeclustors * sectorspercluster * bytespersector
    Let tmp6$ = "total free bytes : " & Format(tmp6$, "#,0")
    Let getdiskinfo$ = tmp1$ & vbCrLf & tmp2$ & vbCrLf & tmp3$ & vbCrLf & tmp4$ & vbCrLf & tmp5$ & vbCrLf & tmp6$
End Function
Public Function getcomputerprocessor() As String
    'from frenzy3.bas by izekial(me)
    Dim sstr As String, myver As OSVERSIONINFO, mysys As SYSTEM_INFO
    Call GetSystemInfo(mysys)
    Select Case mysys.dwProcessorType
        Case PROCESSOR_INTEL_386
            Let sstr$ = "intel 386 dx"
        Case PROCESSOR_INTEL_486
            Let sstr$ = "intel 486 dx"
        Case PROCESSOR_INTEL_PENTIUM
            Let sstr$ = "intel pentium pro"
        Case PROCESSOR_MIPS_R4000
            Let sstr$ = "mips r-4000"
        Case PROCESSOR_ALPHA_21064
            Let sstr$ = "alpha 21064"
        Case Else
            Let sstr$ = "unknown processor"
    End Select
    If mysys.dwNumberOrfProcessors > 1 Then
        Let sstr$ = "multiple " & sstr$ & " processors"
    Else
        Let sstr$ = sstr$ & " processor"
    End If
    Let getcomputerprocessor$ = sstr$
End Function
Public Sub clickicon(lngico As Long)
    Call SendMessageLong(lngico&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngico&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub clickbutton(lngbut As Long)
    Call SendMessageLong(lngbut&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(lngbut&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub addroomtocontrol(thecontrol As Control, adduser As Boolean)
    'edited,  out of dos32.bas
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let room& = findroom&
    If room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            If screenname$ <> getuser$ Or adduser = True Then
                Call thecontrol.AddItem(screenname$)
            End If
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Sub addroomtocontrolwithext(thecontrol As Control, adduser As Boolean)
    'edited,  out of dos32.bas
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let room& = findroom&
    If room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            If screenname$ <> getuser$ Or adduser = True Then
                Call thecontrol.AddItem(screenname$ & "@aol.com")
            End If
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Sub addroomtocontrolwithdogbar(thecontrol As Control, progbar As Control, adduser As Boolean)
    'edited,  out of dos32.bas
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let room& = findroom&
    If room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            Let progbar.Max = SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
                Let progbar.Value = index& - 1&
                Let screenname$ = String$(4&, vbNullChar)
                Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24&
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
                Call CopyMemory(psnHold&, ByVal screenname$, 4&)
                Let psnHold& = psnHold& + 6
                Let screenname$ = String$(16&, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
                If screenname$ <> getuser$ Or adduser = True Then
                    Call thecontrol.AddItem(screenname$)
                End If
            Next index&
            Call CloseHandle(mthread&)
        End If
    End If
End Sub
Public Sub addlisttocontrol(rlist As Long, thecontrol As Control)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim sthread As Long, mthread As Long, strtext As String
    If rlist& = 0& Then Exit Sub
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            thecontrol.AddItem (screenname$)
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Sub addlisttocontrolwithdogbar(rlist As Long, progbar As Control, thecontrol As Control)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim sthread As Long, mthread As Long, strtext As String
    If rlist& = 0& Then Exit Sub
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        Let progbar.Max = SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let progbar.Value = index& - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            thecontrol.AddItem (screenname$)
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Function addlisttostring(rlist As Long, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim sthread As Long, mthread As Long
    If rlist& = 0& Then Exit Function
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            Let addlisttostring$ = screenname$ & strseperator$
        Next index&
        Call CloseHandle(mthread&)
    End If
End Function
Public Function addlisttoclipboard(rlist As Long, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim sthread As Long, mthread As Long, strtext As String
    If rlist& = 0& Then Exit Function
    Clipboard.Clear
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            Let strtext$ = strtext$ & screenname$ & strseperator$
        Next index&
        Call CloseHandle(mthread&)
    End If
    Clipboard.SetText strtext$
End Function
Public Function addroomtostringwithext(adduser As Boolean, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, strstring As String
    Let room& = findroom&
    If room& = 0& Then
        Exit Function
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
                Let screenname$ = String$(4, vbNullChar)
                Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
                Call CopyMemory(psnHold&, ByVal screenname$, 4)
                Let psnHold& = psnHold& + 6
                Let screenname$ = String$(16, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
                If screenname$ <> getuser$ Or adduser = True Then
                    Let strstring$ = strstring$ & "@aol.com" & strseperator$ & screenname$
                End If
            Next index&
            Call CloseHandle(mthread&)
        End If
        Let strstring$ = Left$(strstring$, Len(strstring$) - 1&)
        Let addroomtostringwithext$ = strstring$
    End If
End Function
Public Function addroomtostring(adduser As Boolean, strseperator As String) As String
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, strstring As String
    Let room& = findroom&
    If room& = 0& Then
        Exit Function
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
                Let screenname$ = String$(4, vbNullChar)
                Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
                Call CopyMemory(psnHold&, ByVal screenname$, 4)
                Let psnHold& = psnHold& + 6
                Let screenname$ = String$(16, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
                If screenname$ <> getuser$ Or adduser = True Then
                    Let strstring$ = strstring$ & strseperator$ & screenname$
                End If
            Next index&
            Call CloseHandle(mthread&)
        End If
        Let strstring$ = Left$(strstring$, Len(strstring$) - 1&)
        Let addroomtostring$ = strstring$
    End If
End Function
Public Sub openprofile(strperson As String)
    Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long, lngmenu As Long
    Dim mousecur As POINTAPI, lngprofwin As Long, lngedit As Long
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
End Sub
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
        Let getprofile$ = gettext(lngview&)
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
    Dim child As Long, lngcount As Long
    Do: DoEvents
        child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
        If child& <> 0& Then
            If getcaption(child&) = "Buddy List Window" And closebuddylist = False Then
                lngcount& = lngcount& + 1&
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), 0&, "AOL Child", vbNullString)
                child& = FindWindowEx(FindWindowEx(FindWindow("aol frame25", vbNullString), 0&, "MDIClient", vbNullString), child&, "AOL Child", vbNullString)
                Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
            End If
            Call PostMessage(child&, WM_CLOSE, 0&, 0&): DoEvents
        End If
    Loop Until child& = 0& Or lngcount& = 5&
End Sub
Public Sub setmailprefs()
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
Public Sub scrollmacro(strtext As String)
    Dim lngcounter As Long
    If Mid(strtext$, Len(strtext$), 1&) <> Chr$(10&) Then strtext$ = strtext$ & Chr$(13&) & Chr$(10&)
    Do While InStr(strtext$, Chr$(13&)) <> 0&
        Let lngcounter& = lngcounter& + 1&
        Call sendchat(Mid$(strtext$, 1&, InStr(strtext$, Chr$(13&)) - 1&))
        If lngcounter = 3& Then
            Call pause(0.6)
            Let lngcounter& = 0&
        End If
        Let strtext$ = Mid$(strtext$, InStr(strtext$, Chr$(13&) & Chr$(10&)) + 2&)
    Loop
End Sub

Public Sub formdrag(frmform As Form)
    Call ReleaseCapture
    Call SendMessage(frmform.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub formcenter(frmform As Form)
    Let frmform.Top = (Screen.Height * 0.85) / 2& - frmform.Height / 2&
    Let frmform.Left = Screen.Width / 2& - frmform.Width / 2&
End Sub
Public Sub buildasciichart(thelist As ListBox)
    Dim theIndex As Long
    Let thelist.Columns = 1&
    For theIndex& = 33& To 255&
        thelist.AddItem Chr$(theIndex&)
    Next theIndex&
End Sub
Public Sub buddyremovebyname(strname As String)
    Dim lngaol As Long, lngmdi As Long, lngbuddyedit As Long
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, rlist As Long
    Dim sthread As Long, mthread As Long, lIndex As Long
    On Error Resume Next
    Call keyword("buddy list")
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies")
    Do: DoEvents: lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies"): Loop Until lngbuddyedit& <> 0&
    Let rlist& = FindWindowEx(lngbuddyedit&, 0&, "_AOL_Listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0 To SendMessage(rlist, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24&
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4&, rbytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1&)
            If LCase$(screenname$) = LCase$(strname$) Then
                Let lIndex& = index&
                Call buddyremovebyindex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mthread)
    End If
End Sub
Public Sub buddyremovebyindex(lngindex As Long)
    Dim lngaol As Long, lngmdi As Long, lnglist As Long, lngbuddyedit As Long
    Call keyword("buddy list")
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngbuddyedit& = FindWindowEx(lngmdi&, 0&, vbNullString, "Edit List Buddies")
    Let lnglist& = FindWindowEx(lngbuddyedit&, 0&, "_AOL_Listbox", vbNullString)
    If lngindex& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
    Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
    Call SendMessage(lnglist&, LB_SETCURSEL, lngindex&, 0&)
    Call PostMessage(lnglist&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub
Public Sub formunloadright(frmform As Form)
    Do: DoEvents
        Let frmform.Left = frmform.Left + 250&
    Loop Until frmform.Left > Screen.Width
    Unload frmform
End Sub
Public Sub formbevellines(frmform As Form, lngside As Long, lngwidth As Long, color As Long)
    'used by forminnerbevel, and formouterbevel
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, rightx As Long, bottomy As Long
    Dim x1b As Long, x2b As Long, y1b As Long, y2b As Long, index As Long
    Let rightx = frmform.ScaleWidth - 1&
    Let bottomy = frmform.ScaleHeight - 1&
    Select Case lngside
        Case 0&
            Let x1& = 0&
            Let x1b& = 1&
            Let x2& = 0&
            Let x2b& = 1&
            Let y1& = 0&
            Let y1b& = 1&
            Let y2& = bottomy& + 1&
            Let y2b = -1&
        Case 1&
            Let x1& = rightx&
            Let x1b& = -1&
            Let x2& = x1&
            Let x2b& = x1b&
            Let y1& = 0&
            Let y1b& = 1&
            Let y2& = bottomy& + 1&
            Let y2b& = -1&
        Case 2&
            Let x1& = 0&
            Let x1b& = 1&
            Let x2& = rightx&
            Let x2b& = -1&
            Let y1& = 0&
            Let y1b& = 1&
            Let y2& = 0&
            Let y2b& = 1&
        Case 3&
            Let x1& = 1&
            Let x1b& = 1&
            Let x2& = rightx& + 1&
            Let x2b& = -1&
            Let y1& = bottomy&
            Let y1b& = -1&
            Let y2& = y1&
            Let y2b& = y1b&
    End Select
    For index& = 1& To lngwidth&
        frmform.Line (x1&, y1&)-(x2&, y2&), color&
        Let x1& = x1& + x1b&
        Let x2& = x2& + x2b&
        Let y1& = y1& + y1b&
        Let y2& = y2& + y2b&
    Next index&
End Sub
Public Sub formouterbevel(frmform As Form, lngbevelwidth As Long)
    Let frmform.ScaleMode = 3&
    Call formbevellines(frmform, 0&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmform, 1&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmform, 2&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmform, 3&, lngbevelwidth&, QBColor(8&))
End Sub
Public Sub forminnerbevel(frmform As Form, lngbevelwidth As Long)
    Let frmform.ScaleMode = 3&
    Call formbevellines(frmform, 0&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmform, 1&, lngbevelwidth&, QBColor(15&))
    Call formbevellines(frmform, 2&, lngbevelwidth&, QBColor(8&))
    Call formbevellines(frmform, 3&, lngbevelwidth&, QBColor(15&))
End Sub

Public Sub formleaveright(frmform As Form)
    Do: DoEvents
        frmform.Left = frmform.Left + 250&
    Loop Until frmform.Left > Screen.Width
End Sub
Public Function getchatname() As String
    Let getchatname$ = getcaption(findroom&)
End Function

Public Sub addroomtoclipboard(adduser As Boolean)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, strstring As String
    Let room& = findroom&
    If room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
            Let screenname$ = String$(4, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            Let psnHold& = psnHold& + 6
            Let screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ Or adduser = True Then
                Let strstring$ = strstring$ & "," & screenname$
            End If
        Next index&
        Call CloseHandle(mthread&)
    End If
    If Left$(strstring$, 1&) = "," Then
        Let strstring$ = Right$(strstring$, Len(strstring$) - 1&)
    End If
    If Right$(strstring$, 1&) = "," Then
        Let strstring$ = Left$(strstring$, Len(strstring$) - 1&)
    End If
    Call Clipboard.Clear
    Call Clipboard.SetText(strstring$)
End Sub
Public Sub win98changeresolution(screenwidth As Single, screenheight As Single)
    Dim issafe As Boolean, index As Long, lngchange As Long
    Do: DoEvents
        Let issafe = EnumDisplaySettings(0&, index&, DevM)
        Let index& = index& + 1&
    Loop Until issafe = False
    Let DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    Let DevM.dmPelsWidth = screenwidth
    Let DevM.dmPelsHeight = screenheight
    Let lngchange& = ChangeDisplaySettings(DevM, 0)
End Sub
Public Function getchattext() As String
    Dim ChatWin As Long
    Let ChatWin& = FindWindowEx(findroom&, 0&, "RICHCNTL", vbNullString)
    Let getchattext$ = gettext(ChatWin&)
End Function
Public Function getchatline() As String
    Dim ChatWin As Long, strtext As String, lngenter As Long, lngenter2 As Long
    Let ChatWin& = findroom&
    Let ChatWin& = FindWindowEx(ChatWin&, 0&, "RICHCNTL", vbNullString)
    Let strtext$ = gettext(ChatWin&)
    Let lngenter& = InStr(strtext$, Chr$(13&))
    Do: DoEvents
        If lngenter& <> 0& Then Let lngenter2& = lngenter&
        Let lngenter& = InStr(lngenter& + 1&, strtext$, Chr$(13&))
    Loop Until lngenter& = 0&
    Let getchatline$ = Right$(strtext$, Len(strtext$) - lngenter2&)
End Function
Public Function getchatlinemsg() As String
    Dim strline As String, lngtab As Long
    Let strline$ = getchatline$
    Let lngtab& = InStr(strline$, Chr$(9&))
    Let getchatlinemsg$ = Right$(strline$, Len(strline$) - lngtab&)
End Function
Public Function getchatlinesn() As String
    Dim strline As String, lngtab As Long
    Let strline$ = getchatline$
    Let lngtab& = InStr(strline$, Chr$(9&))
    Let getchatlinesn$ = Left$(strline$, lngtab& - 3&)
End Function
Public Sub chatcommandenterroom()
    'for all chat commands i recommend using chatscan3
    'and just put in the chat scan, before calling this...
    'Let chatline$ What_Said$
    Dim lnglen As Long
    If LCase$(getchatlinemsg$) Like LCase$(".enterroom*") Then
        Let lnglen& = Len(getchatlinemsg$)
        Call privateroom(Right$(getchatlinemsg$, (lnglen& - 10&)))
    End If
End Sub
Public Sub chatcommandimstatus()
    If LCase$(getchatlinemsg$) Like LCase$(".ims off*") Then
        Call sendim("$im_off", " ")
    End If
    If LCase$(getchatlinemsg$) Like LCase$(".ims on*") Then
        Call sendim("$im_on", " ")
    End If
End Sub
Public Sub chatcommandignorebysn()
    Dim strstring As String, lnglen As Long
    If LCase$(getchatlinemsg$) Like LCase$(".ignore*") Then
        Let lnglen& = Len(getchatlinemsg$)
        Let strstring$ = Right$(getchatlinemsg$, lnglen - 8&)
        Call chatignorebyname(strstring$)
    End If
End Sub
Public Sub chatcommandkeyword()
    Dim strstring As String, lnglen As Long
    If LCase$(getchatlinemsg$) Like LCase$(".keyword*") Then
        Let lnglen& = Len(getchatlinemsg$)
        Let strstring$ = Right(getchatlinemsg$, lnglen - 8&)
        Call keyword(strstring$)
    End If
End Sub
Public Sub chatcommandsendim()
    Dim lnglen&, strstring$, lngcolon&, strwho$, strmessage$
    If LCase$(getchatlinemsg$) Like ".sendim*" Then
        Let lnglen& = Len(getchatlinemsg$) - 8&
        Let strstring$ = Right$(getchatlinemsg$, lnglen&)
        Let lngcolon& = InStr(strstring$, ":")
        Let strwho$ = Left$(strstring$, lngcolon& - 1&)
        Let strmessage$ = Mid$(strstring$, lngcolon& + 1&, Len(strstring$) - lngcolon& + 1&)
    End If
    Call sendim(strwho$, strmessage$)
End Sub
Public Sub chatcommandsendmail()
    Dim lnglen&, strstring$, lngcolon&, lngcolon2&, strwho$, strmessage$, strsubject$
    If LCase$(getchatlinemsg) Like ".sendmail*" Then
        Let lnglen& = Len(getchatlinemsg$) - 10&
        Let strstring$ = Right$(getchatlinemsg$, lnglen&)
        Let lngcolon& = InStr(strstring$, ":")
        Let lngcolon2& = InStr(lngcolon& + 1&, strstring$, ":")
        Let strwho$ = Left$(strstring$, lngcolon& - 1&)
        Let strsubject$ = Mid$(strstring$, lngcolon& + 1&, lngcolon2& - lngcolon& - 1)
        Let strmessage$ = Mid$(strstring$, Len(strwho$) + 1& + Len(strsubject$) + 2&, Len(strstring$) - Len(strwho$) + Len(strsubject$) - 1&)
    End If
    Call sendmail(strwho$, strsubject$, strmessage$)
End Sub
Public Sub removenetzerobanner(strdirectory As String)
    If Right$(strdirectory$, 1&) <> "\" Then
        Let strdirectory$ = strdirectory$ & "\"
    End If
    If Dir(strdirectory$) = "" Then Exit Sub
    Let strdirectory$ = strdirectory$ & "bin\"
    Call Kill(strdirectory$ & "jdbcodbc.dll")
    Call Kill(strdirectory$ & "jpeg.dll")
    Call Kill(strdirectory$ & "jre.exe")
    Call Kill(strdirectory$ & "jrew.exe")
    Call Kill(strdirectory$ & "math.dll")
    DoEvents
    Call Kill(strdirectory$ & "mmedia.dll")
    Call Kill(strdirectory$ & "net.dll")
    Call Kill(strdirectory$ & "rmiregistry.exe")
    Call Kill(strdirectory$ & "symcjit.dll")
    Call Kill(strdirectory$ & "sysresource.dll")
End Sub
Public Sub chatcommandimignorer()
    Dim strstring As String
    If LCase$(getchatlinemsg$) Like LCase$(".ignore*") Then
        Let strstring$ = Left$(getchatlinemsg$, 7&)
        Call sendim("$im_off " & strstring$, " ")
    End If
End Sub
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
    'from frenzy3.bas by izekial(me)
    'lngnumberover is the number of times to
    'repeat the findwindowex command. lngwindowhandle
    'is the handle of the window that has the same
    'handle as another window. lngparenthandle is the
    'handle of the parent window. and strclassname is the
    'class name of the window.
    Do: DoEvents
        Let findwindowsex& = FindWindowEx(lngparenthandle&, lngwindowhandle&, strclassname$, vbNullString)
        Let lngnumberover& = lngnumberover& - 1&
    Loop Until lngnumberover& = 0&
End Function
Public Sub privateroom(strroom As String)
    Call keyword("aol://2719:2-2-" & strroom$)
End Sub
Public Sub publicroom(strroom As String)
    Call keyword("aol://2719:21-2-" & strroom$)
End Sub
Public Function fileexists(strfilepathandname As String) As Boolean
    Dim lngcheck As Long
    If InStr(strfilepathandname$, ".") = 0& Then
        Let fileexists = False
    End If
    Let lngcheck& = Len(Dir$(strfilepathandname$))
    If lngcheck& = 0& Then
        Let fileexists = False
        Exit Function
    Else
        Let fileexists = True
    End If
End Function
Public Sub filecopy(file$, destination$)
    If Not fileexists(file$) Then Exit Sub
    If InStr(file$, ".") = 0& Then Exit Sub
    If InStr(destination$, "\") = 0& Then Exit Sub
    Call filecopy(file$, destination$)
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
Public Sub openoldmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long, lngindex As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub opennewmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub opensentmailitem(index As Long)
    Dim lngmailbox As Long, lngtabwin As Long, lngtabwin2 As Long
    Dim lngtree As Long, lngcount As Long, lngindex As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Sub
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtabwin2& = FindWindowEx(lngtabwin&, lngtabwin2&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin2&, 0&, "_AOL_Tree", vbNullString)
        If index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call SendMessage(lngtree&, LB_SETCURSEL, index&, 0&)
        Call PostMessage(lngtree&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(lngtree&, WM_KEYUP, VK_RETURN, 0&)
    End If
End Sub
Public Sub wavstop()
    Call sndPlaySound("", SND_FLAG)
End Sub
Public Sub wavloop(file As String)
    Dim lngflags As Long
    If fileexists(file$) = False Then
        Exit Sub
    Else
        Let lngflags& = SND_ASYNC Or SND_LOOP
        Call sndPlaySound(file$, lngflags&)
    End If
End Sub
Public Sub wavloopstop()
    Dim lngflags As Long
    Let lngflags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound("", lngflags&)
End Sub
Public Sub clickaolmenuicon(lngiconhandle As Long, lngmenuentry As Long)
    Dim mousecur As POINTAPI, lngmenu As Long, index As Long
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(lngiconhandle&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngiconhandle&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Let lngmenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(lngmenu&) = 1&
    For index& = 1& To lngmenuentry&
        Call PostMessage(lngmenu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(lngmenu&, WM_KEYUP, VK_DOWN, 0&)
    Next index&
    Call PostMessage(lngmenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(mousecur.X, mousecur.Y)
End Sub
Public Sub scanforpws(filepath$, filename$)
    Dim filelen As Long, fileinfo As String
    If filename$ = "" Then Exit Sub
    If Right(filepath$, 1) = "\" Then
        Let filename$ = filepath$ & filename$
    Else
        Let filename$ = filepath$ & "\" & filename$
    End If
    If Not fileexists(filename$) Then MsgBox "file not found", 16, "error": Exit Sub
    Let filelen& = Len(filename$)
    Open filename$ For Binary As #1&
        If Err Then MsgBox "an unexpected error occured while opening file!", 16, "error": Exit Sub
        Let fileinfo$ = String(32000&, 0&)
        Get #1&, 1&, fileinfo$
    Close #1&
    Open filename$ For Binary As #2&
        If Err Then MsgBox "an unexpected error occured while opening file!", 16, "error": Exit Sub
        If InStr(1&, LCase$(fileinfo$), "main.idx" & Chr$(0&), 1&) Then
            MsgBox "this file has references to the main.idx file that contains your password, if it is not a program that whould need to use that file, it is most likely a password stealer"
        End If
    Close #2&
End Sub
Public Sub scanfordeltree(filepath$, filename$)
    Dim filelen As Long, fileinfo As String
    If filename$ = "" Then Exit Sub
    If Right(filepath$, 1) = "\" Then
        Let filename$ = filepath$ & filename$
    Else
        Let filename$ = filepath$ & "\" & filename$
    End If
    If Not fileexists(filename$) Then MsgBox "file not found", 16&, "error": Exit Sub
    Let filelen& = Len(filename$)
    Open filename$ For Binary As #1&
        If Err Then MsgBox "an unexpected error occured while opening file!", 16, "error": Exit Sub
        Let fileinfo$ = String(32000&, 0&)
        Get #1&, 1&, fileinfo$
    Close #1&
    Open filename$ For Binary As #2&
        If Err Then MsgBox "an unexpected error occured while opening file!", 16&, "error": Exit Sub
        If InStr(1&, LCase$(fileinfo$), "deltree" & Chr$(0&), 1&) Then
            MsgBox "this file has a reference to the term deltree, which is a vb command used to delete a directory, i would serious consider deleting this file because it may be infected with a deltree"
        End If
        If InStr(1&, LCase$(fileinfo$), "kill" & Chr$(0&), 1&) Then
            MsgBox "this file has a reference to the vb command kill, which is used to delete files. this could just be the word kill but it may also be infected with a deltree"
        End If
    Close #2&
End Sub
Public Sub win98shutdown()
    Static ewx_shutdown
    Dim getmsg As Long
    Let getmsg& = MsgBox("are you sure you want to exit windows?", vbYesNo Or vbQuestion)
    If getmsg& = vbNo Then
        Exit Sub
    Else
        Call ExitWindowsEx(ewx_shutdown, 0&)
    End If
End Sub
Public Function listfindstring(list As ListBox, findstring As String) As Long
    Dim index As Long
    If list.ListCount = 0 Then Exit Function
    For index& = 0 To list.ListCount - 1
        Let list.ListIndex = index&
        If UCase(list.Text) = UCase(findstring$) Then
            Let listfindstring& = index&
            Exit Function
            If Err Then Exit Function
        End If
    Next index&
End Function
Public Sub filedelete(file$)
    If fileexists(file$) = False Then
        Exit Sub
    Else
        If InStr(file$, ".") = 0& Then
            Exit Sub
        Else
            Call Kill(file$)
            DoEvents
        End If
    End If
End Sub
Public Sub filerename(fileandpath As String, newname As String)
    If InStr(fileandpath$, ".") = 0& Then
        Exit Sub
    Else
        Name fileandpath$ As Left$(fileandpath$, InStrRev(fileandpath$, "\")) & newname$
        DoEvents
    End If
End Sub
Public Function filegetattributes(filepath As String) As Integer
    Dim strcheck As String
    Let strcheck$ = Dir(filepath$)
    If strcheck$ = "" Then
        MsgBox "file not found", vbCritical, "izekial32.bas": Exit Function
    Else
        Let filegetattributes = GetAttr(filepath$)
    End If
End Function
Public Function readini(strsection As String, strkey As String, strfullpath As String) As String
   Dim strbuffer As String
   Let strbuffer$ = String$(750, Chr$(0&))
   Let readini$ = Left$(strbuffer$, GetPrivateProfileString(strsection$, ByVal LCase$(strkey$), "", strbuffer, Len(strbuffer), strfullpath$))
End Function
Public Sub writeini(strsection As String, strkey As String, strkeyvalue As String, strfullpath As String)
    Call WritePrivateProfileString(strsection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub
Public Sub win98run(strfile As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngrunform As Long, lngcombo As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
    Let lngbasebar& = FindWindow("BaseBar", vbNullString)
    Let lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
    Let lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyR, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyR, 0&)
    Do: DoEvents
        Let lngrunform& = FindWindow("#32770", vbNullString)
        Let lngcombo& = FindWindowEx(lngrunform&, 0&, "ComboBox", vbNullString)
        Let lngcombo& = FindWindowEx(lngcombo&, 0&, "Edit", vbNullString)
        Let lngokbut& = FindWindowEx(lngrunform&, 0&, "Button", vbNullString)
    Loop Until lngrunform& <> 0& And lngcombo& <> 0&
    Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub win98help()
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngrunform As Long, lngcombo As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
    Let lngbasebar& = FindWindow("BaseBar", vbNullString)
    Let lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
    Let lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyH, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyH, 0&)
End Sub
Public Sub win98controlpanel()
    Call Shell("RUNDLL32.EXE SHELL32.DLL,Control_RunDLL ,", vbNormalFocus)
End Sub
Public Sub win98changewallpaper(strfile As String)
    Dim lngsysparam As Long
    lngsysparam& = SystemParametersInfo(ByVal 20&, vbNullString, ByVal strfile$, &H1)
    If lngsysparam& = 0& Then Exit Sub
End Sub
Public Sub win98runstartmenu(lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
    Let lngbasebar& = FindWindow("BaseBar", vbNullString)
    Let lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
    Let lngmenusite& = GetMenu(lngmenusite&)
    Call runanymenu(lngmenusite&, lngmenunumber&, lngsubmenunumber&)
End Sub
Public Sub win98runstartmenubystring(strmenutext As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
    Let lngbasebar& = FindWindow("BaseBar", vbNullString)
    Let lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
    Let lngmenusite& = GetMenu(lngmenusite&)
    Call runanymenubystring(lngmenusite&, strmenutext$)
End Sub
Public Sub win98findfiles(strtrigger As String)
    Dim lngshell As Long, lngstartbut As Long, lngbasebar As Long
    Dim lngmenusite As Long, lngtoolwin As Long, lngokbut As Long
    Dim lngfindwin As Long, lngedit As Long, lngshelldef As Long
    Dim mousecur As POINTAPI, lngsyslist As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0&, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
    Let lngbasebar& = FindWindow("BaseBar", vbNullString)
    Let lngmenusite& = FindWindowEx(lngbasebar&, 0&, "MenuSite", vbNullString)
    Let lngtoolwin& = FindWindowEx(lngmenusite&, 0&, "ToolbarWindow32", vbNullString)
    Call GetCursorPos(mousecur)
    Call SetCursorPos(Screen.Height, Screen.Width)
    pause 0.2
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyRight, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyRight, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngtoolwin&, WM_KEYUP, vbKeyF, 0&)
    Do: DoEvents
        Let lngfindwin& = FindWindow("#32770", "Find: All Files")
        Let lngedit& = FindWindowEx(lngfindwin&, 0&, "#32770", vbNullString)
        Let lngedit& = FindWindowEx(lngedit&, 0&, "ComboBox", vbNullString)
        Let lngedit& = FindWindowEx(lngedit&, 0&, "Edit", vbNullString)
        Let lngokbut& = FindWindowEx(lngfindwin&, 0&, "Button", "F&ind Now")
    Loop Until lngfindwin& <> 0& And lngedit& <> 0& And lngokbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strtrigger$)
    Call PostMessage(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngokbut&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        lngshelldef& = FindWindowEx(lngfindwin&, 0&, "SHELLDLL_DefView", vbNullString)
        lngsyslist& = FindWindowEx(lngshelldef&, 0&, "SysListView32", vbNullString)
    Loop Until lngshelldef& <> 0& And lngsyslist& <> 0&
    Call waitforlisttoload(lngsyslist&)
    Call SetCursorPos(mousecur.X, mousecur.Y)
End Sub
Public Sub win98clickstart()
    Dim lngshell As Long, lngstartbut As Long
    Let lngshell& = FindWindow("Shell_TrayWnd", vbNullString)
    Let lngstartbut& = FindWindowEx(lngshell&, 0, "Button", vbNullString)
    Call PostMessage(lngstartbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(lngstartbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub formstepdown(theform As Form, lngsteps As Long)
    'thanks nkillaz for the idea
    Dim lngbackcolor As Long, lngindex As Long, lngx As Long, lngy As Long
    On Error Resume Next
    Let lngbackcolor& = theform.BackColor
    Let theform.BackColor = RGB(0&, 0&, 0&)
    For lngindex& = 0& To theform.Count - 1&
        Let theform.Controls(lngindex&).Visible = False
    Next lngindex&
    theform.Show
    Let lngx& = ((Screen.Width - theform.Width) - theform.Left) / lngsteps&
    Let lngy& = ((Screen.Height - theform.Height) - theform.Top) / lngsteps&
    Do: DoEvents
        theform.Move theform.Left + lngx&, theform.Top + lngy&
    Loop Until (theform.Left >= (Screen.Width - theform.Width)) Or (theform.Top >= (Screen.Height - theform.Height))
    Let theform.Left = Screen.Width - theform.Width
    Let theform.Top = Screen.Height - theform.Height
    Let theform.BackColor = lngbackcolor&
    For lngindex& = 0& To theform.Count - 1&
        Let theform.Controls(lngindex&).Visible = True
    Next lngindex&
End Sub
Public Sub createtacsoftini()
    Dim inipath As String, thestring As String, index As Long
    Let inipath$ = App.Path & "\info.ini"
    Let thestring$ = "[Planner]" & vbCrLf$
    For index& = 1& To 10&
        Let thestring$ = thestring$ & "Date" & index& & "Month=" & vbCrLf$
        Let thestring$ = thestring$ & "Date" & index& & "Day=" & vbCrLf$
        Let thestring$ = thestring$ & "Date" & index& & "Year=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
        Let thestring$ = thestring$ & "Time" & index& & "Hour=" & vbCrLf$
        Let thestring$ = thestring$ & "Time" & index& & "Min=" & vbCrLf$
        Let thestring$ = thestring$ & "Time" & index& & "Sec=" & vbCrLf$
        Let thestring$ = thestring$ & "Time" & index& & "State=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
        Let thestring$ = thestring$ & "Purpose" & index& & "=" & vbCrLf$
    Next index&
    Let thestring$ = thestring$ & vbCrLf$ & "[People]"
    For index& = 1& To 10&
        Let thestring$ = thestring$ & "Person" & index& & "Name=" & vbCrLf$
        Let thestring$ = thestring$ & "Person" & index& & "Address=" & vbCrLf$
        Let thestring$ = thestring$ & "Person" & index& & "Misc=" & vbCrLf$
        Let thestring$ = thestring$ & "Person" & index& & "Age=" & vbCrLf$
        Let thestring$ = thestring$ & "Person" & index& & "Number=" & vbCrLf$
        Let thestring$ = thestring$ & "Person" & index& & "Mail=" & vbCrLf$
    Next index&
    For index& = 1& To 10&
        Let thestring$ = thestring$ & "Password" & index& & "Is" & vbCrLf$
        Let thestring$ = thestring$ & "Password" & index& & "Reason" & vbCrLf$
    Next index&
    Let thestring$ = thestring$ & vbCrLf$ & "[Options]" & vbCrLf$
    Let thestring$ = thestring$ & "MakeShortut="
    Let thestring$ = thestring$ & "ShowSplashScreen="
    Open inipath$ For Output As #1&
        Print #1&, thestring$
    Close #1&
End Sub
Public Sub memberroom(strroom As String)
    Call keyword("aol://2719:61-2-" & strroom$)
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
Public Function sendchatspiral(strstring As String)
    Dim lngindex As Long, lnglen As Long
    lnglen& = Len(strstring$)
    Do: DoEvents
        lngindex& = lngindex& + 1&
        Call sendchat(Left$(strstring$, lngindex&))
        pause 0.6
    Loop Until lngindex& = lnglen&
End Function
Public Sub removeroomfromcontrol(thecontrol As Control)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let room& = findroom&
    If room& = 0& Then Exit Sub
    Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
    Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
            Let screenname$ = String$(4, vbNullChar)
            Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            Let itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            Let psnHold& = psnHold& + 6
            Let screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            Call removelistitem(thecontrol, screenname$)
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Sub removeroomfromcontrolwithdogbar(thecontrol As Control, progbar As Control)
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Let room& = findroom&
    If room& = 0& Then
        Exit Sub
    Else
        Let rlist& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
        Let sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
        Let mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
        If mthread& Then
            Let progbar.Max = SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
            For index& = 0 To SendMessage(rlist&, LB_GETCOUNT, 0, 0) - 1
                Let progbar.Value = index& - 1&
                Let screenname$ = String$(4, vbNullChar)
                Let itmhold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
                Let itmhold& = itmhold& + 24
                Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
                Call CopyMemory(psnHold&, ByVal screenname$, 4)
                Let psnHold& = psnHold& + 6
                Let screenname$ = String$(16, vbNullChar)
                Call ReadProcessMemory(mthread&, psnHold&, screenname$, Len(screenname$), rbytes&)
                Let screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
                Call removelistitem(thecontrol, screenname$)
            Next index&
            Call CloseHandle(mthread&)
        End If
    End If
End Sub
Public Sub mircchangecaption(strcaption As String)
    Dim lngmirc As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Call SendMessageByString(lngmirc&, WM_SETTEXT, 0&, strcaption$)
End Sub
Public Function mircchatclear()
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lngedit As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lngedit& = FindWindowEx(lngchannel&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, "/clear")
    Call SendMessageByNum(lngedit&, WM_CHAR, 13&, 0&)
End Function
Public Sub mircsendchat(strstring As String)
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lngedit As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lngedit& = FindWindowEx(lngchannel&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strstring$)
    Call SendMessageByNum(lngedit&, WM_CHAR, 13&, 0&)
End Sub
Public Function mircchatline() As String
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lngedit As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let mircchatline$ = getlastlinefromstring(gettext(lngchannel&))
End Function
Public Sub mircenterroom(strroom As String)
    Call mircsendchat("/j #" & strroom$)
End Sub
Public Sub mircsendmessage(strwho As String, strmessage As String)
    Call mircsendchat("/msg " & strwho$ & strmessage$)
End Sub
Public Function mircgetroomcount() As Long
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lnglist As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lnglist& = FindWindowEx(lngchannel&, 0&, "listbox", vbNullString)
    Let mircgetroomcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function
Public Sub mircmsgclear()
    Dim lngmirc As Long, lngmdi As Long, lngquery As Long, lngedit As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngquery& = FindWindowEx(lngmdi&, 0&, "query", vbNullString)
    Let lngedit& = FindWindowEx(lngquery&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, "/clear")
    Call SendMessageByNum(lngedit&, WM_CHAR, 13&, 0&)
End Sub
Public Sub mircping(strwho As String)
    Call mircsendchat("/ping " & strwho$)
End Sub
Public Sub mircsendchatnotice(strtext As String)
    Call mircsendchat("/notice " & strtext$)
End Sub
Public Sub mircsendwhois(strwho As String)
    Call mircsendchat("/whois " & strwho$)
End Sub
Public Sub sendmassim(strwho As Control, strmessage As String)
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call sendim(strwho.list(lngindex&), strmessage$)
        pause 0.7
    Next lngindex&
End Sub
Public Sub sendmassimwithdogbar(strwho As Control, strmessage As String, progbar As Control)
    Dim lngindex As Long
    Let progbar.Max = strwho.ListCount - 1&
    For lngindex& = 0& To strwho.ListCount - 1&
        Let progbar.Value = lngindex& - 1&
        Call sendim(strwho.list(lngindex&), strmessage$)
        pause 0.7
    Next lngindex&
End Sub
Public Sub changehostname(straoldirectory As String, strnewname As String)
    If Right$(straoldirectory$, 1&) = "\" Then
        straoldirectory$ = straoldirectory$ & "tool\chat.aol"
    Else
        straoldirectory$ = straoldirectory$ & "\tool\chat.aol"
    End If
    Open straoldirectory$ For Binary As #1&
        Seek #1&, 6887
        Put #1&, , strnewname$
    Close #1&
End Sub
Public Sub sendmassmailslow(strwho As Control, strsubject As String, strmessage As String)
    'this is a   s  l  o  w   way to do it
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call sendmail(strwho.list(lngindex&), strsubject$, strmessage$)
    Next lngindex&
End Sub
Public Sub sendmassmailfast(strwho As Control, strsubject As String, strmessage As String)
    'this is a fast way to do it
    Dim lngindex As Long, strpeople As String
    For lngindex& = 0& To strwho.ListCount - 1&
        Let strpeople$ = strpeople$ & strwho.list(lngindex&) & ","
    Next lngindex&
    If Left$(strpeople$, 1&) = "," Then
        Let strpeople$ = Right$(strpeople$, Len(strpeople$) - 1&)
    End If
    If Right$(strpeople$, 1&) = "," Then
        Let strpeople$ = Left$(strpeople$, Len(strpeople$) - 1&)
    End If
    Call sendmail(strpeople$, strsubject$, strmessage$)
End Sub
Public Function getflashmailsubject(index As Long) As String
    Dim lngaol As Long, lngmdi As Long, lngmailbox As Long, lnglist As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Let lngmailbox& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    Let lnglist& = FindWindowEx(lngmailbox&, 0&, "_AOL_Tree", vbNullString)
    If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) < index& - 1& Then
        Exit Function
    Else
        If SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lnglist&, LB_GETTEXTLEN, index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lnglist&, LB_GETTEXT, index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getflashmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Function getnewmailsubject(index As Long) As String
    Dim lngmailbox As Long, lngtabwin As Long, lngtree As Long
    Dim lngtextlen As Long, strmailtext As String, lngtabkey As Long
    Let lngmailbox& = findmailbox&
    If lngmailbox& = 0& Then
        Exit Function
    Else
        Let lngtabwin& = FindWindowEx(lngmailbox&, 0&, "_AOL_TabControl", vbNullString)
        Let lngtabwin& = FindWindowEx(lngtabwin&, 0&, "_AOL_TabPage", vbNullString)
        Let lngtree& = FindWindowEx(lngtabwin&, 0&, "_AOL_Tree", vbNullString)
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getnewmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Function getoldmailsubject(index As Long) As String
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
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getoldmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Sub switchscreenname(index As Long, strpassword As String)
    Dim lngaol As Long, lngmdi As Long, lngswitchwin As Long
    Dim lnglist As Long, lngicon As Long, lngedit As Long
    Call runaolmenu(3&, 0&)
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        Let lngswitchwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Switch Screen Names")
        Let lnglist& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Listbox", vbNullString)
        Let lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until lngswitchwin& <> 0& And lnglist& <> 0& And lngicon& <> 0&
    If index& > SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) - 1& Then
        Exit Sub
    Else
        Call SendMessageLong(lnglist&, LB_SETCURSEL, CLng(index&), 0&)
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngswitchwin& = FindWindow("_AOL_Modal", "Switch Screen Name")
            Let lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngswitchwin& <> 0& And lngicon& <> 0&
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngswitchwin& = FindWindow("_AOL_Modal", "Switch Screen Name")
            Let lngedit& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Edit", vbNullString)
            Let lngicon& = FindWindowEx(lngswitchwin&, 0&, "_AOL_Icon", vbNullString)
        Loop Until lngswitchwin& <> 0& And lngedit& <> 0& And lngicon& <> 0&
        Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strpassword$)
        Call PostMessage(lngicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngicon&, WM_LBUTTONUP, 0&, 0&)
    End If
End Sub
Public Function getsentmailsubject(index As Long) As String
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
        If SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) = 0& Or index& > SendMessage(lngtree&, LB_GETCOUNT, 0&, 0&) - 1& Or index& < 0& Then
            Exit Function
        Else
            Let lngtextlen& = SendMessage(lngtree&, LB_GETTEXTLEN, index&, 0&)
            Let strmailtext$ = String(lngtextlen& + 1&, 0&)
            Call SendMessageByString(lngtree&, LB_GETTEXT, index&, strmailtext$)
            Let lngtabkey& = InStr(strmailtext$, Chr$(9&))
            Let lngtabkey& = InStr(lngtabkey& + 1&, strmailtext$, Chr$(9&))
            Let strmailtext$ = Right$(strmailtext$, Len(strmailtext$) - lngtabkey&)
            Let strmailtext$ = replacestring(strmailtext$, Chr$(0&), "")
            Let getsentmailsubject$ = strmailtext$
        End If
    End If
End Function
Public Sub forwardallflashmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim index As Long
    Call openflashmailbox
    For index& = 0& To getflashmailcount& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardallflashmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call openflashmailbox
    Let progbar.Max = getflashmailcount& - 1&
    For index& = 0& To getflashmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call forwardflashmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardallnewmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim index As Long
    Call opennewmailbox
    For index& = 0& To getnewmailcount& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardallnewmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call opennewmailbox
    Let progbar.Max = getnewmailcount& - 1&
    For index& = 0& To getnewmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call forwardnewmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardalloldmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim index As Long
    Call openoldmailbox
    For index& = 0& To getoldmailcount& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardallsentmail(strwho As String, strmessage As String, deletefwdtext As Boolean)
    Dim index As Long
    Call opensentmailbox
    For index& = 0& To getsentmailcount& - 1&
        Call forwardsentmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardalloldmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call openoldmailbox
    Let progbar.Max = getoldmailcount& - 1&
    For index& = 0& To getoldmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call forwardoldmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardallsentmailwithdogbar(strwho As String, strmessage As String, deletefwdtext As Boolean, progbar As Control)
    Dim index As Long
    Call opensentmailbox
    Let progbar.Max = getsentmailcount& - 1&
    For index& = 0& To getsentmailcount& - 1&
        Let progbar.Value = index& - 1&
        Call forwardsentmailitem(True, strwho$, strmessage$, index&, deletefwdtext): DoEvents
    Next index&
End Sub
Public Sub forwardflashmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deletefwdtext As Boolean)
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
    Call openflashmailitem(index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(index&))
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
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
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
        Let strnewsubject$ = gettext(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
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
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getflashmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getflashmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardnewmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deletefwdtext As Boolean)
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
    Call opennewmailitem(index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(index&))
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
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
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
        Let strnewsubject$ = gettext(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
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
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getnewmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getnewmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardoldmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deletefwdtext As Boolean)
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
    Call openoldmailitem(index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(index&))
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
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
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
        Let strnewsubject$ = gettext(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
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
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getoldmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getoldmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub forwardsentmailitem(mailboxopen As Boolean, strtowho As String, strmessage As String, index As Long, deletefwdtext As Boolean)
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
    Call opensentmailitem(index&)
    Do: DoEvents
        pause 1&
        Let lngchild& = FindWindowEx(lngmdi&, 0&, "aol child", getsentmailsubject(index&))
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
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(index&))
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
        Let strnewsubject$ = gettext(lngsubjectwin&)
        Let strnewsubject$ = Right$(strnewsubject$, Len(strnewsubject$) - 5&)
        Call SendMessageByString(lngsubjectwin&, WM_SETTEXT, 0&, strnewsubject$)
    End If
    Call SendMessageByString(lngtowin&, WM_SETTEXT, 0&, strtowho$): DoEvents
    Call SendMessageByString(lngrichwin&, WM_SETTEXT, 0&, strmessage$): DoEvents
    Do: DoEvents
        Let lngaol& = FindWindow("AOL Frame25", vbNullString)
        Let lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(index&))
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
    Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(index&))
    If lngsendwin& = 0& Then
        Call PostMessage(FindWindowEx(lngmdi&, 0&, "aol child", getsentmailsubject(index&)), WM_CLOSE, 0&, 0&)
    Else
        Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
        Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
        Call PostMessage(lngerrorbut&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(lngerrorbut&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            Let lngerrorwin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Error")
            Let lngerrorbut& = FindWindowEx(lngerrorwin&, 0&, "_AOL_Button", "ok")
            Let lngsendwin& = FindWindowEx(lngmdi&, 0&, "aol child", "Fwd: " & getsentmailsubject(index&))
        Loop Until lngerrorbut& = 0& Or lngsendwin& = 0&
    End If
End Sub
Public Sub chateater()
    'this will clear the chat for everybody
    Dim index As Long
    For index& = 0& To 3&
        Call sendchat(" <p=" & String$(1800&, " "))
        pause 0.2
    Next index&
End Sub
Public Sub roomrunner(strroomname As String)
    'beta testers have asked if this is the code
    'from my prog: ghetto star room buster, the
    'answer is no
    Dim lngroom As Long, lngnumber As Long, lngokwin As Long, lngokbut As Long
    Let lngroom& = findroom&
    Let lngnumber& = 0&
    If lngroom& Then
        windowclose (lngroom&)
        Do: DoEvents
            If lngnumber& = 0& Then lngnumber& = ""
            Call keyword("aol://2719:2-2-" & strroomname$ & lngnumber&)
            Do: DoEvents
                Let lngokwin& = FindWindow("#32770", "america online")
                Let lngokbut& = FindWindowEx(lngokwin&, 0&, vbNullString, "ok")
                Let lngroom& = findroom&
            Loop Until lngokbut& Or lngroom&
            If lngokwin& <> 0& Then
                Do
                    Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
                    Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
                Loop Until lngokwin& = 0&
            End If
            If lngnumber& = "" Then
                Let lngnumber& = 1&
            Else
                Let lngnumber& = lngnumber& + 1&
            End If
        Loop Until lngroom& <> 0&
    End If
End Sub
Public Function waitforokorchatroom(room As String)
    Dim strroomname As String, lngokwin As Long, lngbutton As Long
    Let room$ = LCase$(replacestring(room$, " ", ""))
    Do: DoEvents
        Let strroomname$ = getcaption(findroom&)
        Let strroomname$ = LCase$(replacestring(strroomname$, " ", ""))
        Let lngokwin& = FindWindow("#32770", "america online")
        Let lngbutton& = FindWindowEx(lngokwin&, 0&, "button", "ok")
    Loop Until (lngokwin& <> 0& And lngbutton& <> 0&) Or room$ = strroomname$
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
Public Function servercreatelist(thelist As Control, strseperator As String) As String
    Dim index As Long, strmailstring As String
    For index& = 0& To thelist.ListCount - 1&
        Let strmailstring$ = strmailstring$ & "~" & index& & "~ " & thelist.list(index&) & vbCrLf$
    Next index&
    Let servercreatelist$ = strmailstring$
End Function
Public Function serverlookfor(thelist As Control, strentry As String) As String
    Dim index As Long
    If thelist.ListCount = 0& Then Exit Function
    For index& = 0& To thelist.ListCount - 1&
        If InStr(LCase$(thelist.list(index&)), strentry$) > -1& Then
            Let serverlookfor$ = serverlookfor$ & vbCrLf$ & thelist.list(index&)
            If Err Then Exit Function
        End If
    Next index&
End Function
Public Sub serversendfind(strwho As String, strrequest As String, thelist As ListBox)
    Dim endstring As String, findlist As String
    Let findlist$ = serverlookfor(thelist, strrequest$)
    If findlist$ = "" Then
        If strrequest$ = "my penis" Then strrequest$ = "your penis"
        Call sendchat("~•~ sorry " & strwho$ & ", but " & strrequest & " wasn't found")
        Exit Sub
    End If
    Call sendmail(strwho$, "~•~ find results for " & strrequest$ & " ~•~", "please remember that you requested this mail and that " & getuser$ & ", does nottake responsibility for what you do with the contents of this mail, this is for entertainment purposes only." & vbCrLf$ & stringlinked("http://www.nkillaz.com/izekial/", "izekial32.bas, click here") & vbCrLf$ & findlist$)
    Call sendchat("~•~ " & strwho$ & ", the results for " & strrequest$ & " were sent")
End Sub
Public Sub roombuster(strroomname As String)
    'this is not the code i use in ghetto star room buster
    Dim lngroom As Long, lngokwin As Long, lngokbut As Long
    Let lngroom& = findroom&
    If replacestring$(LCase$(gettext$(lngroom&)), " ", "") = replacestring$(LCase$(strroomname$), " ", "") Then Exit Sub
    If lngroom& Then
        Call windowclose(lngroom&)
        Do: DoEvents
            Call keyword("aol://2719:2-2-" & strroomname$)
            Do: DoEvents
                Let lngokwin& = FindWindow("#32770", "america online")
                Let lngokbut& = FindWindowEx(lngokwin&, 0&, vbNullString, "ok")
                Let lngroom& = findroom&
            Loop Until lngokbut& Or lngroom&
            If lngokwin& <> 0& Then
                Do: DoEvents
                    Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
                    Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
                Loop Until lngokwin& = 0&
            End If
        Loop Until lngroom& <> 0&
    End If
End Sub
Public Function mircmsgsend(strwhat$)
    Dim lngmirc As Long, lngmdi As Long, lngquery As Long, lngedit As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngquery& = FindWindowEx(lngmdi&, 0&, "query", vbNullString)
    Let lngedit& = FindWindowEx(lngquery&, 0&, "edit", vbNullString)
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strwhat$)
    Call SendMessageByNum(lngedit&, WM_CHAR, 13&, 0&)
End Function
Public Sub mircchatroomtocontrol(thecontrol As Control)
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lnglist As Long
    Dim lngcount As Long, lngindex As Long, lnglen As Long, strbuffer As String
    Dim lngitemdata As Long, strtext As String
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lnglist& = FindWindowEx(lngchannel&, 0&, "listbox", vbNullString)
    Let lngcount& = SendMessageLong(lnglist&, LB_GETCOUNT, 0&, 0&)
    For lngindex& = 1& To lngcount&
        Let lnglen& = SendMessageLong(lnglist&, LB_GETTEXTLEN, lngindex& - 1, 0&)
        Let strbuffer$ = String$(lnglen&, 0&)
        Let strtext$ = SendMessageByString(lnglist&, LB_GETTEXT, lngindex& - 1, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lnglist&, LB_GETITEMDATA, lngindex& - 1, 0&)
        If Not lngindex& - 1 Then
            Call thecontrol.AddItem(strbuffer$)
        End If
    Next lngindex&
End Sub
Public Function mircchatroomtostring(strseperator As String) As String
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lnglist As Long
    Dim lngcount As Long, lngindex As Long, lnglen As Long, strbuffer As String
    Dim lngitemdata As Long, strtext As String
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lnglist& = FindWindowEx(lngchannel&, 0&, "listbox", vbNullString)
    Let lngcount& = SendMessageLong(lnglist&, LB_GETCOUNT, 0&, 0&)
    For lngindex& = 1& To lngcount&
        Let lnglen& = SendMessageLong(lnglist&, LB_GETTEXTLEN, lngindex& - 1, 0&)
        Let strbuffer$ = String$(lnglen&, 0&)
        Let strtext$ = SendMessageByString(lnglist&, LB_GETTEXT, lngindex& - 1, ByVal strbuffer$)
        Let lngitemdata& = SendMessage(lnglist&, LB_GETITEMDATA, lngindex& - 1, 0&)
        If Not lngindex& - 1 Then
            mircchatroomtostring$ = mircchatroomtostring$ & strbuffer$ & strseperator$
        End If
    Next lngindex&
End Function
Public Function winampwin() As Long
    Let winampwin& = FindWindow("Winamp v1.x", vbNullString)
End Function
Public Function winampplaylistwin() As Long
    Let winampplaylistwin& = FindWindow("Winamp PE", vbNullString)
End Function
Public Function winampequalizerwin() As Long
    Let winampequalizerwin& = FindWindow("Winamp EQ", vbNullString)
End Function
Public Sub winampopensong(strsong As String)
    Dim lngopenwin As Long, lngedit As Long, lngopenbut As Long
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyL, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyL, 0&)
    Do: DoEvents
        Let lngopenwin& = FindWindow("#32770", "Open file(s)")
        Let lngedit& = FindWindowEx(lngopenwin&, 0&, "Edit", vbNullString)
        Let lngopenbut& = FindWindowEx(lngopenwin&, 0&, "Button", "&Open")
        pause 1
    Loop Until lngopenwin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strsong$)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub winampnextsong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyB, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyB, 0&)
End Sub
Public Sub winampprevioussong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyZ, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyZ, 0&)
End Sub
Public Sub winampaddsongtoplaylist(strsong As String)
    Dim lngopenwin As Long, lngedit As Long, lngopenbut As Long
    Call PostMessage(winampplaylistwin&, WM_KEYDOWN, vbKeyL, 0&)
    Call PostMessage(winampplaylistwin&, WM_KEYUP, vbKeyL, 0&)
    Do: DoEvents
        Let lngopenwin& = FindWindow("#32770", "Open file(s)")
        Let lngedit& = FindWindowEx(lngopenwin&, 0&, "Edit", vbNullString)
        Let lngopenbut& = FindWindowEx(lngopenwin&, 0&, "Button", "&Open")
        pause 1
    Loop Until lngopenwin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strsong$)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub winampjumptosongbyname(strname As String)
    Dim lngjumpwin As Long, lnglist As Long, strtext As String
    Let lngjumpwin& = FindWindow("#32770", "Jump to file")
    Let lnglist& = FindWindowEx(lngjumpwin&, 0&, "ListBox", vbNullString)
    Call SendMessageLong(lnglist&, LB_SETCURSEL, gettreeitemindex(lnglist&, strname$), 0&)
    Call PostMessage(lnglist&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lnglist&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub winampjumptosongbyindex(index As Long)
    Dim lngjumpwin As Long, lnglist As Long, strtext As String
    Let lngjumpwin& = FindWindow("#32770", "Jump to file")
    Let lnglist& = FindWindowEx(lngjumpwin&, 0&, "ListBox", vbNullString)
    Call SendMessageLong(lnglist&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(lnglist&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(lnglist&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub winampstartplay()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyX, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyX, 0&)
End Sub
Public Sub winamppausesong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyC, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyC, 0&)
End Sub
Public Sub winampstopsong()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyV, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyV, 0&)
End Sub
Public Sub winampopenskins()
    Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyControl And vbKeyL, 0&)
    Call PostMessage(winampwin&, WM_KEYUP, vbKeyControl And vbKeyL, 0&)
End Sub
Public Sub winampturnupvolume(lngdecibals As Long)
    Dim index As Long
    For index& = 0& To lngdecibals&
        Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyUp, 0&)
        Call PostMessage(winampwin&, WM_KEYUP, vbKeyUp, 0&)
    Next index&
End Sub
Public Sub winampturndownvolume(lngdecibals As Long)
    Dim index As Long
    For index& = 0& To lngdecibals&
        Call PostMessage(winampwin&, WM_KEYDOWN, vbKeyDown, 0&)
        Call PostMessage(winampwin&, WM_KEYUP, vbKeyDown, 0&)
    Next index&
End Sub
Public Function photoshopwin() As Long
    Let photoshopwin& = FindWindow("Photoshop", vbNullString)
End Function
Public Sub photoshopopenimage(strimage As String)
    'for some reason it just stops after running the menu
    Dim lngopenwin As Long, lngedit As Long, lngopenbut As Long
    Call runanymenu(photoshopwin&, 0&, 1&)
    Do: DoEvents
        lngopenwin& = FindWindow("#32770", "Open")
        lngedit& = FindWindowEx(lngopenwin&, 0&, "Edit", vbNullString)
        lngopenbut& = FindWindowEx(lngopenwin&, 0&, "Button", "&Open")
    Loop Until lngopenwin& <> 0& And lngedit& <> 0& And lngopenbut& <> 0&
    pause 2&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strimage$)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngopenbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub photoshopnewimage(lngwidth As Long, lngheight As Long, frmname As Form)
    'for some reason it just stops after running the menu
    Dim lngfloat As Long, lngedit As Long, lngedit2 As Long, lngokbut As Long, lngfparent As Long
    Call runanymenu(photoshopwin&, 0&, 0&)
    Do: DoEvents
        lngfparent& = photoshopwin&
        lngfloat& = FindWindowEx(lngfparent&, 0&, "PSFloatC", "New")
        lngedit& = FindWindowEx(lngfloat&, 0&, "Edit", vbNullString)
        lngedit2& = FindWindowEx(lngfloat&, lngedit&, "Edit", vbNullString)
        lngedit2& = FindWindowEx(lngfloat&, lngedit2&, "Edit", vbNullString)
        lngokbut& = FindWindowEx(lngfloat&, 0&, "Button", "OK")
    Loop Until lngfloat& <> 0& And lngedit& <> 0& And lngedit2& <> 0& And lngokbut& <> 0&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, CStr(lngwidth&))
    Call SendMessageByString(lngedit2&, WM_SETTEXT, 0&, CStr(lngheight&))
    Call SendMessageLong(lngokbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngokbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub photoshopsaveimage(strname As String)
    'for some reason it just stops after running the menu
    Dim lngsavewin As Long, lngedit As Long, lngsavebut As Long
    Call runanymenu(photoshopwin&, 0&, 6&)
    Do: DoEvents
        lngsavewin& = FindWindow("#32770", "Open")
        lngedit& = FindWindowEx(lngsavewin&, 0&, "Edit", vbNullString)
        lngsavebut& = FindWindowEx(lngsavewin&, 0&, "Button", "&Save")
    Loop Until lngsavewin& <> 0& And lngedit& <> 0& And lngsavebut& <> 0&
    pause 2&
    Call SendMessageByString(lngedit&, WM_SETTEXT, 0&, strname$)
    Call SendMessageLong(lngsavebut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngsavebut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub photoshopprintimage()
    'for some reason it just stops after running the menu
    Dim lngprintwin As Long, lngprintbut As Long
    Call runanymenu(photoshopwin&, 0&, 20&)
    Do: DoEvents
        Let lngprintwin& = FindWindow("#32770", "Print")
        Let lngprintbut& = FindWindowEx(lngprintwin&, 0&, "Button", "OK")
    Loop Until lngprintwin& <> 0& And lngprintbut& <> 0&
    Call SendMessageLong(lngprintbut&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngprintbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub photoshopcloseimage()
    Call runanymenu(photoshopwin&, 0&, 4&)
End Sub















