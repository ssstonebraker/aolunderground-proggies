Attribute VB_Name = "sonicUpload"
Option Explicit

' Object varibles
Dim m_sLineString As String * 1056
Dim m_lngRet As Long
Dim m_sRetString As String


' Constants

Public Const BIF_RETURNONLYFSDIRS = &H1

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const cdlOpen = 1
Public Const cdlSave = 2
Public Const cdlColor = 3
Public Const cdlPrint = 4
Public Const cdlOFNReadOnly = 1             'Checks Read-Only check box for Open and Save As dialog boxes.
Public Const cdlOFNOverwritePrompt = 2      'Causes the Save As dialog box to generate a message box if the selected file already exists.
Public Const cdlOFNHideReadOnly = 4         'Hides the Read-Only check box.
Public Const cdlOFNNoChangeDir = 8          'Sets the current directory to what it was when the dialog box was invoked.
Public Const cdlOFNHelpButton = 10          'Causes the dialog box to display the Help button.
Public Const cdlOFNNoValidate = 100         'Allows invalid characters in the returned filename.
Public Const cdlOFNAllowMultiselect = 200   'Allows the File Name list box to have multiple selections.
Public Const cdlOFNExtensionDifferent = 400 'The extension of the returned filename is different from the extension set by the DefaultExt property.
Public Const cdlOFNPathMustExist = 800      'User can enter only valid path names.
Public Const cdlOFNFileMustExist = 1000     'User can enter only names of existing files.
Public Const cdlOFNCreatePrompt = 2000      'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
Public Const cdlOFNShareAware = 4000        'Sharing violation errors will be ignored.
Public Const cdlOFNNoReadOnlyReturn = 8000  'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
Public Const cdlOFNExplorer = 8000          'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
Public Const cdlOFNNoDereferenceLinks = 100000
Public Const cdlOFNLongNames = 200000

Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA

Public Const errUnableToAddIcon = 1    'Icon can not be added to system tray
Public Const errUnableToModifyIcon = 2 'System tray icon can not be modified
Public Const errUnableToDeleteIcon = 3 'System tray icon can not be deleted
Public Const errUnableToLoadIcon = 4   'Icon could not be loaded (occurs while using icon property)

Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Global Const GW_CHILD = 5
Global Const GW_HWNDNEXT = 2

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_RESETCONTENT = &H14B
Public Const CB_INSERTSTRING = &H14A
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETCURSEL = &H147
Public Const CB_GETITEMDATA = &H150
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLBTEXT = &H148

Public Const MAX_TOOLTIP As Integer = 64
Public Const MAX_PATH = 32

Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_LONG_NAME_CHARS = 64

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

Public Const RGN_OR = 2

Public Const IP_ADDR_ADDED = (11000 + 23)
Public Const IP_ADDR_DELETED = (11000 + 19)
Public Const IP_BAD_DESTINATION = (11000 + 18)
Public Const IP_BAD_OPTION = (11000 + 7)
Public Const IP_BAD_REQ = (11000 + 11)
Public Const IP_BAD_ROUTE = (11000 + 12)
Public Const IP_BUF_TOO_SMALL = (11000 + 1)
Public Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Public Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Public Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Public Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Public Const IP_GENERAL_FAILURE = (11000 + 50)
Public Const IP_HW_ERROR = (11000 + 8)
Public Const IP_MTU_CHANGE = (11000 + 21)
Public Const IP_NO_RESOURCES = (11000 + 6)
Public Const IP_OPTION_TOO_BIG = (11000 + 17)
Public Const IP_PACKET_TOO_BIG = (11000 + 9)
Public Const IP_PARAM_PROBLEM = (11000 + 15)
Public Const IP_PENDING = (11000 + 255)
Public Const IP_REQ_TIMED_OUT = (11000 + 10)
Public Const IP_SOURCE_QUENCH = (11000 + 16)
Public Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Public Const IP_STATUS_BASE = 11000
Public Const IP_SUCCESS = 0
Public Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Public Const IP_UNLOAD = (11000 + 22)
Public Const PING_TIMEOUT = 200

Public Const MAX_IP_STATUS = 11000 + 50
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Const MIN_SOCKETS_REQD = 1
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR = -1

' Parameter for SystemParametersInfo()
Public Const SPI_GETBEEP = 1
Public Const SPI_SETBEEP = 2
Public Const SPI_GETMOUSE = 3
Public Const SPI_SETMOUSE = 4
Public Const SPI_GETBORDER = 5
Public Const SPI_SETBORDER = 6
Public Const SPI_GETKEYBOARDSPEED = 10
Public Const SPI_SETKEYBOARDSPEED = 11
Public Const SPI_LANGDRIVER = 12
Public Const SPI_ICONHORIZONTALSPACING = 13
Public Const SPI_GETSCREENSAVETIMEOUT = 14
Public Const SPI_SETSCREENSAVETIMEOUT = 15
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const SPI_SETSCREENSAVEACTIVE = 17
Public Const SPI_GETGRIDGRANULARITY = 18
Public Const SPI_SETGRIDGRANULARITY = 19
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDESKPATTERN = 21
Public Const SPI_GETKEYBOARDDELAY = 22
Public Const SPI_SETKEYBOARDDELAY = 23
Public Const SPI_ICONVERTICALSPACING = 24
Public Const SPI_GETICONTITLEWRAP = 25
Public Const SPI_SETICONTITLEWRAP = 26
Public Const SPI_GETMENUDROPALIGNMENT = 27
Public Const SPI_SETMENUDROPALIGNMENT = 28
Public Const SPI_SETDOUBLECLKWIDTH = 29
Public Const SPI_SETDOUBLECLKHEIGHT = 30
Public Const SPI_GETICONTITLELOGFONT = 31
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETMOUSEBUTTONSWAP = 33
Public Const SPI_SETICONTITLELOGFONT = 34
Public Const SPI_GETFASTTASKSWITCH = 35
Public Const SPI_SETFASTTASKSWITCH = 36
Public Const SPI_SETDRAGFULLWINDOWS = 37
Public Const SPI_GETDRAGFULLWINDOWS = 38
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Const SPI_SETNONCLIENTMETRICS = 42
Public Const SPI_GETMINIMIZEDMETRICS = 43
Public Const SPI_SETMINIMIZEDMETRICS = 44
Public Const SPI_GETICONMETRICS = 45
Public Const SPI_SETICONMETRICS = 46
Public Const SPI_SETWORKAREA = 47
Public Const SPI_GETWORKAREA = 48
Public Const SPI_SETPENWINDOWS = 49
Public Const SPI_GETFILTERKEYS = 50
Public Const SPI_SETFILTERKEYS = 51
Public Const SPI_GETTOGGLEKEYS = 52
Public Const SPI_SETTOGGLEKEYS = 53
Public Const SPI_GETMOUSEKEYS = 54
Public Const SPI_SETMOUSEKEYS = 55
Public Const SPI_GETSHOWSOUNDS = 56
Public Const SPI_SETSHOWSOUNDS = 57
Public Const SPI_GETSTICKYKEYS = 58
Public Const SPI_SETSTICKYKEYS = 59
Public Const SPI_GETACCESSTIMEOUT = 60
Public Const SPI_SETACCESSTIMEOUT = 61
Public Const SPI_GETSERIALKEYS = 62
Public Const SPI_SETSERIALKEYS = 63
Public Const SPI_GETSOUNDSENTRY = 64
Public Const SPI_SETSOUNDSENTRY = 65
Public Const SPI_GETHIGHCONTRAST = 66
Public Const SPI_SETHIGHCONTRAST = 67
Public Const SPI_GETKEYBOARDPREF = 68
Public Const SPI_SETKEYBOARDPREF = 69
Public Const SPI_GETSCREENREADER = 70
Public Const SPI_SETSCREENREADER = 71
Public Const SPI_GETANIMATION = 72
Public Const SPI_SETANIMATION = 73
Public Const SPI_GETFONTSMOOTHING = 74
Public Const SPI_SETFONTSMOOTHING = 75
Public Const SPI_SETDRAGWIDTH = 76
Public Const SPI_SETDRAGHEIGHT = 77
Public Const SPI_SETHANDHELD = 78
Public Const SPI_GETLOWPOWERTIMEOUT = 79
Public Const SPI_GETPOWEROFFTIMEOUT = 80
Public Const SPI_SETLOWPOWERTIMEOUT = 81
Public Const SPI_SETPOWEROFFTIMEOUT = 82
Public Const SPI_GETLOWPOWERACTIVE = 83
Public Const SPI_GETPOWEROFFACTIVE = 84
Public Const SPI_SETLOWPOWERACTIVE = 85
Public Const SPI_SETPOWEROFFACTIVE = 86
Public Const SPI_SETCURSORS = 87
Public Const SPI_SETICONS = 88
Public Const SPI_GETDEFAULTINPUTLANG = 89
Public Const SPI_SETDEFAULTINPUTLANG = 90
Public Const SPI_SETLANGTOGGLE = 91
Public Const SPI_GETWINDOWSEXTENSION = 92
Public Const SPI_SETMOUSETRAILS = 93
Public Const SPI_GETMOUSETRAILS = 94
Public Const SPI_SCREENSAVERRUNNING = 97

' SystemParametersInfo flags
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

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

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const sys_Add = 0       'Specifies that an icon is being add
Public Const sys_Modify = 1    'Specifies that an icon is being modified
Public Const sys_Delete = 2    'Specifies that an icon is being deleted

Public Const vbLeftButton = 1     'Left button is pressed
Public Const vbRightButton = 2    'Right button is pressed
Public Const vbMiddleButton = 4   'Middle button is pressed

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE


' Types
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Type SonicUPLOADSTAT
    UL_PERDONE As String
    UL_MINLEFT As String
    UL_FILENAME As String
End Type

Public Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Dim ICMPOPT As ICMP_OPTIONS

Public Type ICMP_ECHO_REPLY
    Address         As Long
    Status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Public Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpstring As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "User" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Public Declare Function GetClassName Lib "User" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function waveOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long


Public Function getUpload() As SonicUPLOADSTAT
    '  how to use:
    '  this is a simple as sitting down
    '  make 3 labels; and 1 timer
    '  set the timer's interval to 200
    '  and use this code
    '
    '  label1.Caption = getUpload.UL_FILENAME
    '  label2.Caption = getUpload.UL_MINLEFT
    '  label3.Caption = getUpload.UL_PERDONE
    '
    '  note: please give credit to me, sonic if you
    '  do use this function, i spent at least an hr
    '  writing a whole new bas for you to use.
    '  this bas makes reading the upload window
    '  easy.
    
    Dim Y As SonicUPLOADSTAT, uWin&, uStatic&, uStatic2&
    Dim uCap$, sCap$, sCap2$
    uWin& = FindWindow("_AOL_Modal", vbNullString)
    If uWin& = 0& Then Exit Function
    uStatic& = FindWindowEx(uWin&, 0&, "_AOL_Static", vbNullString)
    uStatic2& = FindWindowEx(uWin&, uStatic&, "_AOL_Static", vbNullString)
    uCap$ = returnText(uWin&)
    If InStr(uCap$, " - ") = 0 Then Exit Function
    sCap$ = returnText(uStatic&)
    sCap2$ = returnText(uStatic2&)
    uCap$ = Mid(uCap$, InStr(uCap$, "r - ") + 4)
    uCap$ = Left(uCap$, Len(uCap$) - 2)
    sCap$ = Mid(sCap$, InStr(sCap$, "ing") + 3)
    sCap2$ = Mid(Right(sCap2$, Len(sCap2$) - 6), InStr(sCap$, Chr(32)))
    sCap2$ = Left(sCap2, Len(sCap2$) - Len(Mid(sCap2$, InStr(sCap2$, " "))))
    If IsNumeric(uCap$) = False Then uCap$ = "na"
    If IsNumeric(sCap2$) = False Then sCap2$ = "na"
    With Y
        .UL_FILENAME = sCap$
        .UL_MINLEFT = sCap2$
        .UL_PERDONE = uCap$
    End With
    getUpload = Y
End Function

Public Function returnText(winhWnd As Long) As String
    Dim txtHld$, txtLen As Long
    txtLen& = SendMessage(winhWnd&, WM_GETTEXTLENGTH, 0&, 0&) + 1
    txtHld$ = String(txtLen&, 0&)
    Call SendMessageByString(winhWnd&, WM_GETTEXT, txtLen&, txtHld$)
    returnText$ = txtHld
End Function

Public Sub Upchat()
    Dim UpWin As Long
    UpWin& = FindWindow("_AOL_MODAL", vbNullString)
    If UpWin& = 0 Then
        Exit Sub
    End If
    Call ShowWindow(UpWin&, SW_HIDE)
    Call ShowWindow(UpWin&, SW_MINIMIZE)
End Sub
Public Sub UnUpchat()
    Dim UpWin As Long
    UpWin& = FindWindow("_AOL_MODAL", vbNullString)
    If UpWin& = 0 Then
        Exit Sub
    End If
    Call ShowWindow(UpWin&, SW_HIDE)
    Call ShowWindow(UpWin&, SW_RESTORE)
End Sub
