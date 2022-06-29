Attribute VB_Name = "Module1"
'                  ----------------------------------------------
'                  First Release Of The Kronosv1 Bas File!
'                  ----------------------------------------------
'
'                                       "oo  oo"
'                                        oo oo
'                                        oooo
'                                        oo oo
'                                       "oo  oo"v1
'
'                  -----------------------------------------------
'                  AOL 4.0 32 And Aol 5.0 Released August 10, 2001
'                  -----------------------------------------------

'I hope you like this module, and if you use this please give me credit in Your Prog.
'Mail me at Saph And Three@aol.com ,With Your Prog Or Comments Thanx!















































































































































































































































































































































































































































































































































Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVlpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHQueryRecycleBin Lib "Shell32" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHEmptyRecycleBin Lib "Shell32" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwflags As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function lstrcpy Lib "Kernel" Alias "LstrCpy" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function GetHostName Lib "WSOCK32.DLL" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Public Declare Function GetHostByName Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
ucallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Const SPI_SETDESKWALLPAPER = 20
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SENDWININICHANGE = &H2
Const WM_NCCREATE = &H81
Const WM_NCDESTROY = &H82
Const WM_NCCALCSIZE = &H83
Const WM_NCHITTEST = &H84
Const WM_NCPAINT = &H85
Const WM_NCACTIVATE = &H86
Const WM_GETDLGCODE = &H87
Const WM_NCLBUTTONDBLCLK = &HA3
Const WM_NCRBUTTONDOWN = &HA4
Const WM_NCRBUTTONUP = &HA5
Const WM_NCRBUTTONDBLCLK = &HA6
Const WM_NCMBUTTONDOWN = &HA7
Const WM_NCMBUTTONUP = &HA8
Const WM_NCMBUTTONDBLCLK = &HA9
Const WM_KEYFIRST = &H100
Const WM_DEADCHAR = &H103
Const WM_SYSKEYDOWN = &H104
Const WM_SYSKEYUP = &H105
Const WM_SYSCHAR = &H106
Const WM_SYSDEADCHAR = &H107
Const WM_KEYLAST = &H108
Const WM_INITDIALOG = &H110
Const WM_TIMER = &H113
Const WM_HSCROLL = &H114
Const WM_VSCROLL = &H115
Const WM_INITMENU = &H116
Const WM_INITMENUPOPUP = &H117
Const WM_MENUSELECT = &H11F
Const WM_MENUCHAR = &H120
Const WM_ENTERIDLE = &H121
Const WM_CTLCOLORMSGBOX = &H132
Const WM_CTLCOLOREDIT = &H133
Const WM_CTLCOLORLISTBOX = &H134
Const WM_CTLCOLORBTN = &H135
Const WM_CTLCOLORDLG = &H136
Const WM_CTLCOLORSCROLLBAR = &H137
Const WM_CTLCOLORSTATIC = &H138
Const WM_MOUSEFIRST = &H200
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const WM_MOUSELAST = &H209
Const WM_PARENTNOTIFY = &H210
Const WM_ENTERMENULOOP = &H211
Const WM_EXITMENULOOP = &H212
Const WM_DROPFILES = &H233
Const WM_MDIREFRESHMENU = &H234
Const WM_UNDO = &H304
Const WM_RENDERFORMAT = &H305
Const WM_RENDERALLFORMATS = &H306
Const WM_DESTROYCLIPBOARD = &H307
Const WM_DRAWCLIPBOARD = &H308
Const WM_PAINTCLIPBOARD = &H309
Const WM_VSCROLLCLIPBOARD = &H30A
Const WM_SIZECLIPBOARD = &H30B
Const WM_ASKCBFORMATNAME = &H30C
Const WM_CHANGECBCHAIN = &H30D
Const WM_HSCROLLCLIPBOARD = &H30E
Const WM_QUERYNEWPALETTE = &H30F
Const WM_PALETTEISCHANGING = &H310
Const WM_PALETTECHANGED = &H311
Const WM_HOTKEY = &H312
Const WM_PENWINFIRST = &H380
Const WM_PENWINLAST = &H38F
Const WM_NULL = &H0
Const WM_CREATE = &H1
Const WM_SIZE = &H5
Const WM_ACTIVATE = &H6

Dim Buttun, Caption As String

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BROWN = 996633
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000

Type COLORRGB
red As Long
green As Long
blue As Long
End Type



Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const Cb_AddString& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETCOUNT& = &H146
Public Const Cb_GetItemData = &H150
Public Const Cb_GetLbText& = &H148
Public Const CB_RESETCONTENT& = &H14B
Public Const CB_SETCURSEL& = &H14E
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
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
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
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const BM_GETSTATE = &HF2
Public Const BM_SETSTATE = &HF3
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
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

Type SYSTEM_INFO
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

Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
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


Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_586 = 586
Public Const PROCESSOR_INTEL_PENTIUM = "Pentium"
Public Const PROCESSOR_INTEL_786 = 786
Public Const PROCESSOR_INTEL_886 = 886
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000

Dim DevM As DEVMODE

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const conMCIAppTitle = "MCI Control Application"
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304

Private Type DEVMODE
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

Public Const SND_MEMORY = &H4

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Type DISKSPACEINFO
RootPath As String * 3
freebytes As Long
totalbytes As Long
FreePcnt As Single
UsedPcnt As Single
End Type

Global CurrentDisk As DISKSPACEINFO

Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const VK_MENU = &H12
Public Const VK_SHIFT = &H10
Public Const VK_UP = &H26
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const EM_GETLINE = &HC4
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
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
Public Const Plug = "Hound_Dog"
Public GiveClsNam As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
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
Private Const REG_SZ = 1
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const ERROR_SUCCESS = 0&
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const CLR_MENUBAR = &H80000004
Const Number_of_Menu_Selections = 3
Const SC_MOVE = &HF012
Const WA_INACTIVE = 0
Const WA_ACTIVE = 1
Const WA_CLICKACTIVE = 2
Const WM_SETFOCUS = &H7
Const WM_KILLFOCUS = &H8
Const WM_ENABLE = &HA
Const WM_SETREDRAW = &HB
Const WM_PAINT = &HF
Const WM_QUERYENDSESSION = &H11
Const WM_QUIT = &H12
Const WM_QUERYOPEN = &H13
Const WM_ERASEBKGND = &H14
Const WM_SYSCOLORCHANGE = &H15
Const WM_ENDSESSION = &H16
Const WM_SHOWWINDOW = &H18
Const WM_WININICHANGE = &H1A
Const WM_DEVMODECHANGE = &H1B
Const WM_ACTIVATEAPP = &H1C
Const WM_FONTCHANGE = &H1D
Const WM_TIMECHANGE = &H1E
Const WM_CANCELMODE = &H1F
Const WM_SETCURSOR = &H20
Const WM_MOUSEACTIVATE = &H21
Const WM_CHILDACTIVATE = &H22
Const WM_QUEUESYNC = &H23
Const WM_GETMINMAXINFO = &H24
Const WM_PAINTICON = &H26
Const WM_ICONERASEBKGND = &H27
Const WM_NEXTDLGCTL = &H28
Const WM_SPOOLERSTATUS = &H2A
Const WM_DRAWITEM = &H2B
Const WM_MEASUREITEM = &H2C
Const WM_DELETEITEM = &H2D
Const WM_VKEYTOITEM = &H2E
Const WM_CHARTOITEM = &H2F
Const WM_SETFONT = &H30
Const WM_GETFONT = &H31
Const WM_SETHOTKEY = &H32
Const WM_GETHOTKEY = &H33
Const WM_QUERYDRAGICON = &H37
Const WM_COMPAREITEM = &H39
Const WM_COMPACTING = &H41
Const CN_RECEIVE = &H1
Const CN_TRANSMIT = &H2
Const CN_EVENT = &H4
Const WM_WINDOWPOSCHANGING = &H46
Const WM_WINDOWPOSCHANGED = &H47
Const WM_POWER = &H48

Type COPYDATASTRUCT
dwData As Long
cbData As Long
lpData As Long
End Type

Type myLong64
var1 As Long
var2 As Long
End Type

Public Type SHQUERYRBINFO
cbSize         As Long
i64Size        As myLong64
i64NumItems    As myLong64
End Type

Public Const SW_NORMAL = 1
Public Const SWW_SHOWNORMAL = 1
Public Const HWND_DESKTOP = 0
Public Const SHERB_NOCONFIRMATION = &H1
Public Const SHERB_NOPROGRESSUI = &H2
Public Const SHERB_NOSOUND = &H4
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
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


Type FormState
Deleted As Integer
Dirty As Integer
color As Long
End Type
Public Sub WindowsShutDown()
Call ExitWindows(EWX_SHUTDOWN, 0)
End Sub
Public Sub WindowsReStart()
Call ExitWindows(EWX_REBOOT, 0)
End Sub
Public Function WindowsGetUser()
Dim GasStation As String
     Dim Kronos As Long
     GasStation = Space$(255)
     Kronos = Len(Spcs)
     Call GetUserName(Spcs, Lent)

    If Lent > 0 Then
         WindowsGetUser = Left$(Spcs, Lent)
    Else
         WindowsGetUser = vbNullString
    End If

    End Function

Sub ShowWelcome()
Dim aol As Long, MdiFractor As Long, child As Long, Title As Long, X
aol& = FindWindow("AOL Frame25", vbNullString)
MdiFractor& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Title& = FindChildByTitle(mdi&, chil&, "Welcome")
If Title& <> 0 Then
X = ShowWindow(Tit&, SW_SHOW)
End If

End Sub

Public Function TDate()
Dim X
X = Format(Date, "mmmm/dd/yyyy")
TDate = X
End Function

Sub CheckPhish(CmdButton As Control)
TalkToChat ("(¯`·._-=-Function X Sighning Off Aol ")
HoldUp (0.3)
TalkToChat ("(¯`·._-=-Status : Account Verification")
HoldUp (2)
Call AolSignOff
For X = 0 To List1.ListCount - 1
HoldUp (0.3)
Call SignOnAsGuest("" & List1.List(X), List2.List(X))
HoldUp (1)
If AolOnline = True Then
List1.AddItem List3.List(X) + "Phish Valid"
HoldUp (1)
Else
List1.AddItem List3.List(X) + "Phish InValid"
HoldUp (0.15)
End If
Next X
End Sub




End Sub

Sub DimHandle(CmdButton As Control)
Dim Handle As String
AskHandle:
Handle$ = InputBox("Enter the screen name you wish to add:", "Function X - Aim Phisher")
If Trim(Handle$) = "" Then GoTo AskHandle
ListA.AddItem LCase(Handle$)
End Sub
Function FileInput(FileName As String)
Free = FreeFile
Open FileName For Input As Free
    i = FileLen(FileName)
    X = Input(i, Free)
Close Free
    FileInput = X
End Function

Function FileInput2(FileName As String)
Free = FreeFile
Open FileName For Input As Free
    i = FileLen(FileName)
    X = Input(i - 2, Free)
Close Free
    FileInput2 = X
End Function
Function FileLoadList(FileName As String, Lis As ListBox)
On Error Resume Next
Open FileName For Input As #1
Do While Not EOF(1)
 Line Input #1, ln$
Lis.AddItem ln$
Loop
Close #1
End Function

Function FileSaveList(FileName As String, Lis As ListBox)
Free = FreeFile
Open FileName For Output As Free
For X = 0 To Lis.ListCount
Print #1, Lis.List(X)
Next X
Close #1
End Function

Function LoadTimes(FileName)
        Free = FreeFile
    Open FileName For Random As Free
    Close Free
    Open FileName For Input As Free
i = FileLen(FileName)
X = Input(i, Free)
    Close Free

    Open FileName For Output As #1
X = Val(X) + 1
    Print #1, X
    Close #1
        LoadTimes = X
End Function
Sub AolSighnOff()
Dim aol As Long
aol& = FindWindow("AOL Frame25", vbNullString)
End Sub
Sub CenterForm(Frm As Form)
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub
Public Sub CenterFormTop(Frm As Form)
Frm.Left = (Screen.Width - Frm.Width) / 2
Frm.Top = (Screen.Height - Frm.Height) / (Screen.Height)
End Sub
Function CharToChr(Letter As String)
X = Asc(Letter)
CharToChr = X
End Function

Public Function CheckAlive(ScreenName As String) As Boolean

Dim aol As Long, mdi As Long, ErrorWindow As Long
Dim ErrorTextWindow As Long, ErrorString As String
Dim MailWindow As Long, NoWindow As Long, NoButton As Long

Call SendMail("*, " & ScreenName$, "Kronos Toser_v1!", "=)")
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)

Do
DoEvents
ErrorWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorString$ = GetText(ErrorTextWindow&)
Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""

If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
sendchat "•ø• tøs results •dead account•"
CheckAlive = False
Else
CheckAlive = True
sendchat "•ø• tøs results •active account•"
End If

MailWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")

Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
DoEvents

Do
DoEvents
NoWindow& = FindWindow("#32770", "America Online")
NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
Loop Until NoWindow& <> 0& And NoButton& <> 0

Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)

End Function
Public Sub ClickButton(WHwnd As Long)
Call SendMessage(WHwnd, WM_LBUTTONDOWN, 0, 0&)
DoEvents
Call SendMessage(WHwnd, WM_LBUTTONUP, 0, 0&)
DoEvents
End Sub
Sub ClickForward()
AOLWindow = FindWindow("AOL Frame25", vbNullString)
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")

For l = 1 To 8
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()

Next l

clickicon (AOIcon%)

End Sub

Sub clickicon(Icon%)
ClickButton% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
ClickButton% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ClickKeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailBox% = FindChildByTitle(mdi%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 2
AOIcon% = GetWindow(AOIcon%, 2)

Next l

clickicon (AOIcon%)

End Sub

Sub ClickNext()

mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")

For l = 1 To 5
AOIcon% = GetWindow(AOIcon%, 2)

Next l

clickicon (AOIcon%)

End Sub

Sub ClickOK(Caption)
OKW = FindWindow("#32770", Caption)
OKD = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
OKU = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub

Sub ClickRead()

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailBox% = FindChildByTitle(mdi%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")

For l = 1 To 0
AOIcon% = GetWindow(AOIcon%, 2)

Next l

clickicon (AOIcon%)

End Sub
Sub ClickSendAfterError(Recipients)

aol% = FindWindow("AOL Frame25", vbNullString)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOedit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOedit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOedit%, WM_SETTEXT, 0, Recipients)

For GetIcon = 1 To 14
AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

clickicon (AOIcon%)

Do: DoEvents
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOedit% = FindChildByClass(AOMail%, "_AOL_Edit")

Loop Until AOedit% = 0

End Sub

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

Sub CloseWindow(CloseWin)
Closes = SendMessage(CloseWin, WM_CLOSE, 0, 0)
End Sub

Sub ComboDuplicates(cmbBox As Control)

For X = 0 To cmbBox.ListCount - 1
For Y = 0 To cmbBox.ListCount - 1
If LCase(cmbBox.List(X)) Like LCase(cmbBox.List(Y)) And X <> Y Then
cmbBox.RemoveItem (Y)
End If

Next Y
Next X

End Sub

Sub LoginPw(CmdButton As Control, Form As Form)
If Text1.Text = "" Then
Msgretval = MsgBox("You Must Enter Your Handle AnD Please Dont Put Your Sn Like A Fucking Idiot!", 48, "Enter Your Handle!!")
Select Case Msgretval
Case 1
End Select
Else
If Text2.Text = "" Then
Msgretval = MsgBox("You Must Enter A PassWord To Begin!", 48, "Add A The Password That Has Been Given To You!")
Select Case Msgretval
Case 1
End Select
Else
Call Playwav("click.wav")
Pause (0.2)
If Text2 = "levelone" Then
Form.Show
Form.Hide
HoldUp (0.2)
TalkToChat ("(¯`·._-=-Funtion X 2001 Baiter  ")
HoldUp (0.6)
TalkToChat ("(¯`·._-=-By Kronos: Now Loaded ")
HoldUp (0.5)
TalkToChat ("(¯`·._-=-By Aol User :    " & LCase(GetUser))
HoldUp (0.6)
TalkToChat ("(¯`·._-=-Users Handle: " & (Text1.Text))
HoldUp (0.5)
Else
Msgretval = MsgBox("You Must Enter The PassWord That Has Been Given To you!", 48, "Add  The Password That Has Been Given To You!")
Select Case Msgretval
Case 1
End Select
TalkToChat ("(¯`·._-=- PassWord Has Been Denied!  ")
End
End If
End If
End If
End Sub

End Sub

Public Function GetUser2()
'Sub Taken From Hound2000
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


Sub PhishEm(CmdButton As Control)
If AIMOnline = False Then
Msgretval = MsgBox("You Must sigh On To Aim Before Using This Function!", 48, "Come On Man Think!")
Select Case Msgretval
Case 1
End Select
Else
If Text1.Text = "" Then
Msgretval = MsgBox("You Must Generate A Phrase First Before Using This Function ! <- Kronos !", 48, "Generate Phrase !")
Select Case Msgretval
Case 1
End Select
Else
End If
End If

If List1.Text = "" Then
Msgretval = MsgBox("You Must atleast add a name To the list <-Kronos!", 48, "Add A Fucking Name Idiot !")
Select Case Msgretval
Case 1
End Select
Else
End If

For X = 0 To List1.ListCount - 1
Call AimImsend("" & List1.List(X), Text1.Text)
HoldUp (0.2)
Call AimCloseIM
HoldUp (5)
Next X
End Sub

Public Sub AimImOpen()
DoEvents:
    lParIm& = FindWindowEx(FindMain&, 0, "_Oscar_TabGroup", vbNullString)
    lHanIm& = FindWindowEx(lParIm&, 0, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lHanIm&)
End Sub
Public Sub AimCloseIM()
DoEvents:
    lIm& = FindIM&
    If lIm& > 0 Then
    Call Win_Close(lIm&)
    End If
End Sub


Public Sub SendMail(Person As String, subject As String, message As String)

Dim aol As Long, mdi As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
Dim Rich As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, SendButton As Long
Dim Combo As Long, fCombo As Long, ErrorWindow As Long
Dim Button1 As Long, Button2 As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)

Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
DoEvents

Do
DoEvents
OpenSend& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)

For DoIt& = 1 To 13
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&

Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&

Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
DoEvents

Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
DoEvents

Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents

Pause 0.2

Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Public Function GetUser() As String

Dim aol As Long, mdi As Long, welcome As Long
Dim child As Long, UserString As String

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
UserString$ = GetCaption(child&)

If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
GetUser$ = UserString$
Exit Function
Else
Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
UserString$ = GetCaption(child&)

If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
GetUser$ = UserString$
Exit Function
End If

Loop Until child& = 0&
End If

GetUser$ = ""

End Function
Sub IMsOn()
Call InstantMessage("$IM_ON", "_Kronos! ")
End Sub
Sub IMSOff()
Call InstantMessage("$IM_OFF", "_Kronos ")
End Sub
Public Sub InstantMessage(Recipient, message)
 
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Call Keyword("Instant Message")

Do: DoEvents
IMW% = FindChildByTitle(mdi%, "Send Instant Message")
AOedit% = FindChildByClass(IMW%, "_AOL_Edit")
AORich% = FindChildByClass(IMW%, "RICHCNTL")
AOIcon% = FindChildByClass(IMW%, "_AOL_Icon")
Loop Until AOedit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOedit%, WM_SETTEXT, 0, Recipient)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)

clickicon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMW% = FindChildByTitle(mdi%, "Send Instant Message")
OKW% = FindWindow("#32770", "America Online")

If OKW% <> 0 Then Call SendMessage(OKW%, WM_CLOSE, 0, 0): Closer2 = SendMessage(IMW%, WM_CLOSE, 0, 0): Exit Do
If IMW% = 0 Then Exit Do
Loop

End Sub
Sub FormOffTop(Frm As Form)
SetWinOnTop = SetWindowPos(Frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub FormOnTop(Frm As Form)
Call SetWindowPos(Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub Keyword(KeywordAddress)
KWA$ = KeywordAddress
    aol% = FindWindow("AOL Frame25", vbNullString)
Temp% = FindChildByClass(aol%, "AOL Toolbar")
Temp% = FindChildByClass(Temp%, "_AOL_Toolbar")
Temp% = FindChildByClass(Temp%, "_AOL_Combobox")
KWBox% = FindChildByClass(Temp%, "Edit")
Call SendMessageByString(KWBox%, WM_SETTEXT, 1, KeywordAddress)
Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
    
End Sub
Public Sub PrivateRoom(RoomName As String)
Call Keyword("aol://2719:2-2-" & RoomName)
Dim Timer As Integer, FullWindows As Long
For Timer = 1 To 4
FullWindows& = FindWindow("#32770", vbNullString)
    If FullWindows& <> 0 Then
        Call SendMessage(FullWin&, WM_CLOSE, 0, 0&)
    Exit For
    End If
HoldUp (0.5)
Next Timer

End Sub
Function IfProgIsActive() As Boolean
IfProgIsActive = False

                                                                                                                                                        If (App.PrevInstance = True) Then
                                                                                                                                                                    IfProgIsActive = True
    End If

    End Function
Public Function InTheRoom() As Boolean
Dim child As Long
Dim InRoom As Boolean
child& = FindChat
If child& <> 0 Then
InRoom = True
Else
InRoom = False
End If
InTheRoom = InRoom
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

Sub BuddyListGet(BLis As ListBox)
'Sub From Hound2000
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

Sub BuddyListSet(Lis As ListBox)
'Sub From Hound2000
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




Function LastChatLine()

On Error Resume Next

chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$

End Function
Function LastChatLineWithSN()

chattext$ = getchattext

For FindChar = 1 To Len(chattext$)
thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(chattext$, lastlen, Len(thechars$))
LastChatLineWithSN = LastLine

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
Sub ListDuplicates(lbBox As ListBox)

For a = 0 To lbBox.ListCount - 1
Current = lbBox.List(a)

For b = 0 To lbBox.ListCount - 1
Nower = lbBox.List(b)

If b = a Then GoTo DontKill

If Nower = Current Then lbBox.RemoveItem (b)

DontKill:
Next b
Next a

End Sub
Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)

On Error Resume Next

Dim MyString As String, aString As String, bString As String

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
Sub LoadComboBox(cmbBox As ComboBox, Directory As String)

On Error Resume Next

Dim MyString As String

Open Directory$ For Input As #1
While Not EOF(1)
Input #1, MyString$
DoEvents
cmbBox.AddItem MyString$
Wend
Close #1

End Sub
Public Sub Loadlistbox(Directory As String, lbBox As ListBox)

On Error Resume Next

Dim MyString As String

Open Directory$ For Input As #1
While Not EOF(1)
Input #1, MyString$
DoEvents
lbBox.AddItem MyString$
Wend
Close #1

End Sub

Sub LoadNewMail(lbBox As ListBox)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1

DoEvents
sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)

Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)

Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
lbBox.AddItem MyString$

Next AddMails&

End Sub

Sub LoadText(txtLoad As TextBox, Path As String)

On Error Resume Next

Dim TextString As String

Open Path$ For Input As #1
TextString$ = Input(LOF(1), #1)
Close #1

txtLoad.Text = TextString$

End Sub
Public Function LookForMailBox()
Dim aol As Long, mdi As Long, child As Long
Dim TabControl As Long, TabPage As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)

If TabControl& <> 0& And TabPage& <> 0& Then
'FindMailBox& = Child&
Exit Function
Else
Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)

If TabControl& <> 0& And TabPage& <> 0& Then
'FindMailBox& = Child&
Exit Function
End If

Loop Until child& = 0&
End If

'FindMailBox& = 0&

End Function

Public Function MailCountFlash() As Long

Dim aol As Long, mdi As Long, fMail As Long, fList As Long
Dim Count As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
MailCountFlash& = Count&

End Function

Public Function MailCountOld() As Long

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
MailCountOld& = Count&

End Function

Public Function MailCountSent() As Long

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Function
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
MailCountSent& = Count&

End Function

Public Sub MailForward(SendTo As String, message As String, DeleteFwd As Boolean)

Dim aol As Long, mdi As Long, Error As Long
Dim OpenForward As Long, OpenSend As Long, SendButton As Long
Dim DoIt As Long, EditTo As Long, EditCC As Long
Dim EditSubject As Long, Rich As Long, fCombo As Long
Dim Combo As Long, Button1 As Long, Button2 As Long
Dim TempSubject As String

Do
DoEvents
OpenForward& = FindForwardWindow
OpenSend& = FindSendWindow
EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)

For DoIt& = 1 To 13
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&
Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    
If DeleteFwd = True Then
TempSubject$ = GetText(EditSubject&)
TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)

Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)

DoEvents
End If

Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SendTo$)
DoEvents

Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
DoEvents

Do Until OpenSend& = 0& Or Error& <> 0&
DoEvents

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
Error& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
OpenSend& = FindSendWindow
SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)

For DoIt& = 1 To 11
SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
Next DoIt&

Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)

TimeOut (1)

Loop

If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)

End Sub
Public Sub MailOpenEmailFlash(index As Long)

Dim aol As Long, mdi As Long, fMail As Long, fList As Long
Dim fCount As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)

If fCount& < index& Then Exit Sub
Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)
Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)

End Sub
Public Sub MailOpenEmailNew(index As Long)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& < index& Then Exit Sub
Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)

End Sub
Public Sub MailOpenEmailOld(index As Long)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& < index& Then Exit Sub
Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)

End Sub
Public Sub MailOpenEmailSent(index As Long)

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
      
If Count& < index& Then Exit Sub
Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)

End Sub

Public Sub MailOpenFlash()

Dim aol As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
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
Call PostMessage(sMod&, WM_KEYDOWN, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)

End Sub
Public Sub MailOpenNew()

Dim aol As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
Dim WinVis As Long

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
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

Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)

End Sub

Public Sub MailOpenOld()

Dim aol As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
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

For DoThis& = 1 To 4
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&

Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)

End Sub
Public Sub MailOpenSent()

Dim aol As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim CurPos As POINTAPI, WinVis As Long

aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
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

For DoThis& = 1 To 5
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&

Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)

End Sub
Public Sub MailToListFlash(lbBox As ListBox)

Dim aol As Long, mdi As Long, fMail As Long, fList As Long
Dim Count As Long, MyString As String, AddMails As Long
Dim sLength As Long, Spot As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
fMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")

If fMail& = 0& Then Exit Sub
fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
MyString$ = String(255, 0)

For AddMails& = 0 To Count& - 1
DoEvents
sLength& = SendMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)

Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)

Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
MyString$ = ReplaceString(MyString$, Chr(0), "")
lbBox.AddItem MyString$

Next AddMails&

End Sub
Public Sub MailToListNew(lbBox As ListBox)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)

Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)

Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
lbBox.AddItem MyString$

Next AddMails&

End Sub
Public Sub MailToListOld(lbBox As ListBox)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)

Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)

Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
lbBox.AddItem MyString$

Next AddMails&

End Sub
Public Sub MailToListSent(lbBox As ListBox)

Dim MailBox As Long, TabControl As Long, TabPage As Long
Dim mTree As Long, AddMails As Long, sLength As Long
Dim Spot As Long, MyString As String, Count As Long

MailBox& = FindMailBox&

If MailBox& = 0& Then Exit Sub
TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)

If Count& = 0 Then Exit Sub
For AddMails& = 0 To Count& - 1
DoEvents
sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
MyString$ = String(sLength& + 1, 0)
    
Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)

Spot& = InStr(MyString$, Chr(9))
Spot& = InStr(Spot& + 1, MyString$, Chr(9))
MyString$ = Right(MyString$, Len(MyString$) - Spot&)
lbBox.AddItem MyString$

Next AddMails&

End Sub

Sub MailWaitForLoadFlash()

Do
DoEvents
M1% = MailCountFlash
TimeOut (1)
M2% = MailCountFlash
TimeOut (1)
M3% = MailCountFlash
Loop Until M1% = M2% And M2% = M3%

M3% = MailCountFlash

End Sub
Sub MailWaitForLoadNew()

MailOpenNew

Do
Box% = FindChildByTitle(AOLMDI(), GetUser & "'s Online Mailbox")
Loop Until Box% <> 0

Lists = FindChildByClass(Box%, "_AOL_Tree")

Do
DoEvents
M1% = MailCountNew
TimeOut (1)
M2% = MailCountNew
TimeOut (1)
M3% = MailCountNew
Loop Until M1% = M2% And M2% = M3%

M3% = MailCountNew

End Sub
Sub MailWaitForLoadOld()

MailOpenOld

Do
DoEvents
M1% = MailCountOld
HoldUp (1)
M2% = MailCountOld
TimeOut (1)
M3% = MailCountOld
Loop Until M1% = M2% And M2% = M3%

M3% = MailCountOld

End Sub

Sub MailWaitForLoadSent()
MailOpenSent

Do
DoEvents
M1% = MailCountSent
TimeOut (1)
M2% = MailCountSent
TimeOut (1)
M3% = MailCountSent
Loop Until M1% = M2% And M2% = M3%

M3% = MailCountSent

End Sub
Sub MinimizeIMsWhileMMing()
If FindMailBox = 0 Then
Exit Sub
End If

Do
IMW% = FindChildByTitle(AOLMDI, ">Instant Message From:")
IMW2% = FindChildByTitle(AOLMDI, "  Instant Message From:")

If IMW% Then GoTo Greed
If IMW2% Then GoTo Greed2

Greed:
MinimizeWindow (IMW%)
TimeOut (1)

Greed2:
MinimizeWindow (IMW2%)
TimeOut (1)
Flash% = FindChildByTitle(AOLMDI, "Incoming/Saved Mail")
Loop Until FindMailBox = 0 And Flash% = 0

End Sub
Sub MinimizeWindow(hwnd)
Shrink% = ShowWindow(hwnd, SW_MINIMIZE)
End Sub
Public Sub MoveForm(Frm As Form)
ReleaseCapture
Call SendMessage(Frm.hwnd, &HA1, 2, 0&)
End Sub
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
On Error Resume Next

Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False

X = Done / Total * Shape.Width

Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print Percent(Done, Total, 100) & "%"

End Sub
Sub PlaySound(File)
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
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
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub
Function RunMenuByChar(IconNum As Integer, Letter As String)
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
ToolB% = FindChildByClass(tool%, "_AOL_Toolbar")
icona% = FindChildByClass(ToolB%, "_AOL_Icon")

For X = 1 To IconNum
icona% = GetWindow(icona%, GW_HWNDNEXT)
Next X

DoEvents
Chng$ = CharToChr(Letter)
SendLetter = SendMessageByString(icona%, WM_CHAR, Chng$, 0&)

End Function
Public Sub RunMenuByString(SearchString As String)

Dim aol As Long, aMenu As Long, mCount As Long
Dim LookFor As Long, sMenu As Long, sCount As Long
Dim LookSub As Long, sID As Long, sString As String

aol& = FindWindow("AOL Frame25", vbNullString)
aMenu& = GetMenu(aol&)
mCount& = GetMenuItemCount(aMenu&)

For LookFor& = 0& To mCount& - 1
sMenu& = GetSubMenu(aMenu&, LookFor&)
sCount& = GetMenuItemCount(sMenu&)

For LookSub& = 0 To sCount& - 1
sID& = GetMenuItemID(sMenu&, LookSub&)
sString$ = String$(100, " ")

Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)

If InStr(LCase(sString$), LCase(SearchString$)) Then
Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
Exit Sub
End If

Next LookSub&
Next LookFor&

End Sub
Sub RunMenuByString2(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For findstring = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, findstring)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
menuitem% = SubCount%
GoTo MatchString
End If

Next getstring

Next findstring
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, menuitem%, 0)

End Sub
Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
On Error Resume Next
Dim SaveLists As Long

Open Directory$ For Output As #1
For SaveLists& = 0 To ListA.ListCount - 1
Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
Next SaveLists&
Close #1

End Sub
Sub SaveComboBox(cmbBox As ComboBox, Directory As String)

On Error Resume Next

Dim SaveList As Long

Open Directory$ For Output As #1
For SaveList& = 0 To cmbBox.ListCount - 1
Print #1, cmbBox.List(SaveList&)
Next SaveList&
Close #1

End Sub
Public Sub SaveListBox(Directory As String, lbBox As ListBox)

On Error Resume Next

Dim SaveList As Long

Open Directory$ For Output As #1
For SaveList& = 0 To lbBox.ListCount - 1
Print #1, lbBox.List(SaveList&)
Next SaveList&
Close #1

End Sub
Sub SaveText(txtSave As TextBox, Path As String)

On Error Resume Next

Dim TextString As String

TextString$ = txtSave.Text
Open Path$ For Output As #1
Print #1, TextString$
Close #1

End Sub
Public Sub SetMailPrefs()
Dim aol As Long, tool As Long, Toolbar As Long
Dim ToolIcon As Long, DoThis As Long, sMod As Long
Dim mdi As Long, mPrefs As Long, mButton As Long
Dim gStatic As Long, mStatic As Long, fStatic As Long
Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
Dim CurPos As POINTAPI, WinVis As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
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

For DoThis& = 1 To 3
Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
Next DoThis&

Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)

Do
DoEvents
mPrefs& = FindWindowEx(mdi&, 0&, "AOL Child", "Preferences")
gStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "General")
mStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Mail")
fStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Font")
maStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Marketing")
Loop Until mPrefs& <> 0& And gStatic& <> 0& And mStatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&

mButton& = FindWindowEx(mPrefs&, 0&, "_AOL_Icon", vbNullString)
mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)

Do
DoEvents

Call SendMessage(mButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(mButton&, WM_LBUTTONUP, 0&, 0&)

dMod& = FindWindow("_AOL_Modal", "Mail Preferences")

HoldUp (0.6)

Loop Until dMod& <> 0&

ConfirmCheck& = FindWindowEx(dMod&, 0&, "_AOL_Checkbox", vbNullString)
CloseCheck& = FindWindowEx(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
OKButton& = FindWindowEx(dMod&, 0&, "_AOL_icon", vbNullString)

Call SendMessage(ConfirmCheck&, BM_SETCHECK, False, vbNullString)
Call SendMessage(CloseCheck&, BM_SETCHECK, True, vbNullString)
Call SendMessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
Call SendMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)

DoEvents
Call PostMessage(mPrefs&, WM_CLOSE, 0&, 0&)

End Sub
Sub SetText(AOLW, txt)
WindowText% = SendMessageByString(AOLW, WM_SETTEXT, 0, txt)
End Sub

Sub SignOnAsGuest(GName As String, GPass As String)
'Sub From hound2000
Dim aol As Long, mdi As Long, chil As Long, Tit1 As Long, _
Tit2 As Long, Tit As Long, Combo As Long, Edi As Long
Dim Modal As Long, Stat As Long, IconB As Long, _
Edi2 As Long, Edi3 As Long, icona As Long, i As Integer, an As Long
    If AolOnline4 = True Then
       HoldUp (0.3)
        Call AolSighnOff
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
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(Combo&, WM_LBUTTONUP, 0, 0&)
    If Edi& <> 0 Then
For i = 1 To 9
Call SendMessageByNum(Combo&, WM_KEYDOWN, VK_RIGHT, 0)
Call SendMessageByNum(Combo&, WM_KEYUP, VK_RIGHT, 0)
holup (0.2)
Next i
    End If
    DoEvents
   icona& = GetWindow(icona&, GW_HWNDNEXT)
        icona& = GetWindow(icona&, GW_HWNDNEXT)
            icona& = GetWindow(icona&, GW_HWNDNEXT)
            Call ClickButton(icona&)
Call WaitForModal
Pause (0.2)
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
HoldUp (0.5)
     If AolOnline = True Then
            Pause (0.5)
            Exit Sub
     End If
Loop Until an& <> 0
Call SendMessage(an&, WM_CLOSE, 0, 0&)
Pause (0.5)

Call SendMessageByString(Edi&, WM_SETTEXT, 0, GName)
Call SendMessageByString(Edi3&, WM_SETTEXT, 0, GPass)
Call SendMessage(IconB&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(IconB&, WM_LBUTTONUP, 0, 0&)
HoldUp (0.5)

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "Goodbye from America Online!")
Stat& = FindWindowEx(Tit&, 0&, "_AOL_Static", vbNullString)
    If Stat& <> 0 Then
        Call SendMessageByString(Stat&, WM_SETTEXT, 0, "    Hey You Are You Annoyed?!   ")
    End If

End Sub

Sub StopLoop(CmdButton As Control)
Do
DoEvents:
Loop

End Sub

Public Sub WaitForModal()
Dim Modal As Long, Stat As Long, Edi As Long
Do: DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
Edi& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
If Modal& <> 0 And Stat& <> 0 And Edi& <> 0 Then Exit Do
Loop

End Sub

Function UserFromIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient") '
IMW% = FindChildByTitle(mdi%, ">Instant Message From:")

If IMW% Then GoTo Greed
IMW% = FindChildByTitle(mdi%, "  Instant Message From:")

If IMW% Then GoTo Greed
Exit Function

Greed:
IMCap$ = GetCaption(IMW%)
thesn$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
UserFromIM = thesn$

End Function
Sub waitforok()
Do
DoEvents
OKW = FindWindow("#32770", "America Online")
DoEvents
Loop Until OKW <> 0
OKB = FindChildByTitle(OKW, "OK")
OKD = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
OKU = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub
Function WaitForWin(Caption As String) As Integer
Do
DoEvents
MDIW% = FindChildByTitle(AOLMDI, Caption$)
Loop Until MDIW% <> 0
WaitForWin = MDIW%

End Function

Public Sub WindowHide(hwnd As Long)
Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub WindowShow(hwnd As Long)
Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text)
DoEvents
thechar$ = Mid(Text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$

End Function
Sub UnUpchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")

If AOGauge% <> 0 Then Upp% = AOModal%

Call EnableWindow(Upp%, 1)
Call EnableWindow(aol%, 0)

Y = ShowWindow(Upp%, SW_MAXIMIZE)

End Sub
Sub Upchat()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")

If AOGauge% <> 0 Then Upp% = AOModal%

Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)

X = ShowWindow(Upp%, SW_MINIMIZE)

End Sub

Sub TalkToChat(ChatSend)

If InStr(LCase(ChatSend), LCase("<font")) Then ChatSend = ChatSend Else ChatSend = "<font style=ARIAL>" & ChatSend

Dim Ch As String

Ch$ = ChatSend

room& = FindRoom
AORich& = FindChildByClass(room&, "RICHCNTL")
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich2& = GetWindow(AORich&, 2)
AORich3& = GetWindow(AORich2&, 2)
AORich4& = GetWindow(AORich3&, 2)
AORich5& = GetWindow(AORich4&, 2)

Call ClickOK("America Online")

Past = GetText(AORich&)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich&, WM_SETTEXT, 0, ChatSend)
Call SendMessageByNum(AORich&, WM_CHAR, 13, 0)

Past = GetText(AORich2&)
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, ChatSend)
Call SendMessageByNum(AORich2&, WM_CHAR, 13, 0)

Past = GetText(AORich3&)
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, ChatSend)
Call SendMessageByNum(AORich3&, WM_CHAR, 13, 0)

Past = GetText(AORich4&)
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, ChatSend)
Call SendMessageByNum(AORich4&, WM_CHAR, 13, 0)

OKW = FindWindow("#32770", "America Online")
OKC = FindChildByTitle(OKW, "OK")

If OKW Then Call ClickOK("America Online")

HoldUp (0.000000000001)

Call SendMessageByString(AORich&, WM_SETTEXT, 0, Past)
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, Past)
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, Past)
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, Past)
Call SendMessageByString(AORich5&, WM_SETTEXT, 0, Past)

Call ClickOK("America Online")

End Sub
Sub DeleteItem(lbBox As ListBox, item$)

On Error Resume Next

Do
NoFreeze% = DoEvents()
If LCase$(lbBox.List(a)) = LCase$(item$) Then lbBox.RemoveItem (Y)
Y = 1 + Y
Loop Until Y >= lbBox.ListCount

End Sub
Public Function ErrorName(name As Long) As String

Dim aol As Long, mdi As Long, ErrorWindow As Long
Dim ErrorTextWindow As Long, ErrorString As String
Dim NameCount As Long, TempString As String

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
ErrorWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")

If ErrorWindow& = 0& Then Exit Function
ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorString$ = GetText(ErrorTextWindow&)
NameCount& = LineCount(ErrorString$) - 2

If NameCount& < name& Then Exit Function
TempString$ = LineFromString(ErrorString$, name& + 2)
TempString$ = Left(TempString$, InStr(TempString$, "-") - 2)
ErrorName$ = TempString$

End Function
Public Sub Implode(Form As Form)
Dim sStart As Integer, GoNow As Long

For sStart% = 1 To 273
DoEvents
Form.Height = Form.Height - 10
Form.Top = (Screen.Height - Form.Height) \ 2
Next sStart%
    
End
    
End Sub

Public Function ErrorNameCount() As Long

Dim aol As Long, mdi As Long, ErrorWindow As Long
Dim ErrorTextWindow As Long, ErrorString As String
Dim NameCount As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
ErrorWindow& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")

If ErrorWindow& = 0& Then Exit Function
ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorString$ = GetText(ErrorTextWindow&)
NameCount& = LineCount(ErrorString$) - 2
ErrorNameCount& = NameCount&

End Function

Public Sub Explode(Form As Form)

Dim sStart As Integer, GoNow As Long

For sStart% = 1 To 273
DoEvents
Form.Height = Form.Height + 10
Form.Top = (Screen.Height - Form.Height) \ 2
Next sStart%
    
End Sub

Public Function FileAttributes(TheFile As String) As Integer

Dim SafeFile As String

SafeFile$ = Dir(TheFile$)

If SafeFile$ <> "" Then
FileGetAttributes% = GetAttr(TheFile$)
End If

End Function

Public Function FileExists(sFileName As String) As Boolean

If Len(sFileName$) = 0 Then
FileExists = False
Exit Function
End If

If Len(Dir$(sFileName$)) Then
FileExists = True
Else
FileExists = False
End If

End Function
Public Sub FileSetHidden(TheFile As String)

Dim SafeFile As String

SafeFile$ = Dir(TheFile$)

If SafeFile$ <> "" Then
SetAttr TheFile$, vbHidden
End If

End Sub

Public Sub FileSetNormal(TheFile As String)

Dim SafeFile As String

SafeFile$ = Dir(TheFile$)

If SafeFile$ <> "" Then
SetAttr TheFile$, vbNormal
End If

End Sub

Public Sub FileSetReadOnly(TheFile As String)

Dim SafeFile As String

SafeFile$ = Dir(TheFile$)

If SafeFile$ <> "" Then
SetAttr TheFile$, vbReadOnly
End If

End Sub
Function FindChildByClass(ParentW, ChildHand)

FindW% = GetWindow(ParentW, 5)

If UCase(Mid(GetClass(FindW%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Found
FindW% = GetWindow(ParentW, GW_CHILD)
If UCase(Mid(GetClass(FindW%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Found
While FindW%
Firss% = GetWindow(ParentW, 5)
If UCase(Mid(GetClass(Firss%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Found
FindW% = GetWindow(FindW%, 2)
If UCase(Mid(GetClass(FindW%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Found
Wend

FindChildByClass = 0

Found:
room% = FindW%
FindChildByClass = room%

End Function

Public Function FindForwardWindow() As Long

Dim aol As Long, mdi As Long, child As Long
Dim Rich1 As Long, Rich2 As Long, Combo As Long
Dim FontCombo As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)

If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
FindForwardWindow& = child&
Exit Function
Else
Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)

If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
FindForwardWindow& = child&
Exit Function
End If

Loop Until child& = 0&
End If
    
FindForwardWindow& = 0&

End Function

Function FindIt(ParentW, ChildHand)

Find1% = GetWindow(ParentW, 5)

If UCase(Mid(GetClass(Find1%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Results
Find1% = GetWindow(ParentW, GW_CHILD)

If UCase(Mid(GetClass(Find1%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Results

While Find1%
Find2% = GetWindow(ParentW, 5)
If UCase(Mid(GetClass(Find2%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Results
Find1% = GetWindow(Find1%, 2)
If UCase(Mid(GetClass(Find1%), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo Results
Wend

FindIt = 0

Results:
AIMF% = Find1%
FindIt = AIMF%

End Function

Function FindItsTitle(ParentW, ChildHand)

FindT1% = GetWindow(ParentW, 5)

If UCase(GetCaption(FindT1%)) Like UCase(ChildHand) Then GoTo Results
FindT1% = GetWindow(ParentW, GW_CHILD)
While FindT1%

FindT2% = GetWindow(ParentW, 5)

If UCase(GetCaption(FindT2%)) Like UCase(ChildHand) & "*" Then GoTo Results
FindT1% = GetWindow(FindT1%, 2)

If UCase(GetCaption(Num1%)) Like UCase(ChildHand) & "*" Then GoTo Results
Wend

FindItsTitle = 0

Results:
AIMF% = FindT1%
FindItsTitle = AIMF%

End Function

Public Function FindMailBox() As Long

Dim aol As Long, mdi As Long, child As Long
Dim TabControl As Long, TabPage As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)

If TabControl& <> 0& And TabPage& <> 0& Then
FindMailBox& = child&
Exit Function
Else
Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)

If TabControl& <> 0& And TabPage& <> 0& Then
FindMailBox& = child&
Exit Function
End If

Loop Until child& = 0&
End If

FindMailBox& = 0&

End Function
Public Function FindSendWindow() As Long

Dim aol As Long, mdi As Long, child As Long
Dim SendStatic As Long
    
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")

If SendStatic& <> 0& Then
FindSendWindow& = child&
Exit Function
Else

Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")

If SendStatic& <> 0& Then
FindSendWindow& = child&
Exit Function
End If

Loop Until child& = 0&
End If

FindSendWindow& = 0&

End Function
Function GetCaption(hwnd)
hWndLength% = GetWindowTextLength(hwnd)
hWndTitle$ = String$(hWndLength%, 0)
b% = GetWindowText(hwnd, hWndTitle$, (hWndLength% + 1))

GetCaption = hWndTitle$

End Function
Function getchattext()
Box& = FindWindowEx(FindChatRoom, 0&, "RICHCNTL", vbNullString)
getchattext = GetText(Box&)
End Function

Function GetClass(child)

Buffers$ = String$(250, 0)
GetClas% = GetClassName(child, Buffers$, 250)

GetClass = Buffers$

End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function


Sub AOLHide()

aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)

End Sub

Sub AOLMaxamize()

aol% = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(aol%, SW_MAXAMIZE)

End Sub
Function AOLMDI()

aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")

End Function

Sub AOLMinimize()

aol% = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(aol%, SW_MINIMIZE)

End Sub
Public Function AolOnline4() As Boolean
'For Aol 4.0
Dim aol As Long, mdi As Long, chil As Long, mdi2 As Long, Tit As Long, Tit2 As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
chil& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Tit& = FindChildByTitle(mdi&, chil&, "welcome,")
AolOnline = IIf(Tit& <> 0, True, False)
End Function

Sub AOLSetText(win, txt)
SetWhat% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AOLShow()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Sub

Sub AOLTimer()
OurHandle& = FindWindow("_AOL_TimeKeeper", "")
Call EnableWindow(OurHandle&, 0)
End Sub
Function AOLVersion()
'For All Version's Of Aol
aol% = FindWindow("AOL Frame25", vbNullString)
Temp% = FindChildByClass(aol%, "AOL Toolbar")
Temp% = FindChildByClass(Temp%, "_AOL_Toolbar")
Temp% = FindChildByClass(Temp%, "_AOL_Combobox")
KWBox% = FindChildByClass(Temp%, "Edit")
If V60 = 1 Then AOLVersion = 6: Exit Function
If V50 = 1 Then AOLVersion = 5: Exit Function
If KWBox% <> 0& Then AOLVersion = 4: Exit Function
If V30 = 1 Then AOLVersion = 3: Exit Function
If KWBox% = 0 Then AOLVersion = 2: Exit Function
End Function
Function AOLWindow()
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Sub CenterForm(Frm As Form)
Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub

Sub AddRoomEmailExtension(lbBox As ListBox)

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long

room = FindChatRoom()
AOLHandle = FindChildByClass(room, "_AOL_ListBox")
AolThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AolProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AolProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24

Call ReadProcessMemory(AolProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)

ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)

Call ReadProcessMemory(AolProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)

If Person$ = GetUser Then GoTo NextUser

lbBox.AddItem LCase((Person$) & "@aol.com")

NextUser:

Next index
Call CloseHandle(AolProcessThread)

End If

End Sub


Sub AddRoomToListbox(lbBox As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long

room = FindChatRoom()
AOLHandle = FindChildByClass(room, "_AOL_ListBox")
AolThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AolProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AolProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24

Call ReadProcessMemory(AolProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)

ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)

Call ReadProcessMemory(AolProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)

If Person$ = GetUser Then GoTo NextUser

lbBox.AddItem LCase((Person$))

NextUser:

Next index
Call CloseHandle(AolProcessThread)

End If

End Sub

Sub AddRoomToCombobox(cmbBox As ComboBox)

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long

room = FindChatRoom()
AOLHandle = FindChildByClass(room, "_AOL_ListBox")
AolThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AolProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AolProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24

Call ReadProcessMemory(AolProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)

ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)

Call ReadProcessMemory(AolProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)

If Person$ = GetUser Then GoTo NextUser

cmbBox.AddItem LCase((Person$))

NextUser:

Next index
Call CloseHandle(AolProcessThread)

End If

End Sub
Sub PhishPrases(CmdButton As Control)
If VScroll1.Value = 1 Then
Text1.Text = " Hi Im a represenatative for America Online.We are having problems with are main database.It Seems to be some one has uploaded a virus  in our main computer.Our information Has been erased from our computers,we would like you to comply to these terms, and respond with your screen name and password So we can fix these errors,and store information back into our computers as soon as possible. Incompliance with these terms will result to complete termination of your account. Thank You for using america Online!"
Label6.Caption = "Phish Sub's 1"
End If
If VScroll1.Value = 2 Then
Text1.Text = " Hi My Screen Name is TosReport I would please advise you to ignore the message down at the bottom of this instant message and reply with you Master Screen name and Passoword, we are experiencing technical diffilculties with are database.One Our main computers have failed to respond to our services.If you do not respond to this instant message your account will be terminated thank you from america online! "
Label6.Caption = "Phish Master's 1"
End If
If VScroll1.Value = 3 Then
Text1.Text = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
Label6.Caption = "Phish Sub's 2"
End If
If VScroll1.Value = 4 Then
Text1.Text = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
Label6.Caption = "Phish Information 1"
End If
If VScroll1.Value = 5 Then
Text1.Text = "Hello! I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
Label6.Caption = "Phish Sub's 3"
End If
If VScroll1.Value = 6 Then
Text1.Text = "Dear Over Head Account User, We at AOL try our best to avoid hiring corrupt employees but sometimes our hardest is not our best. Unfortunately one of our former employees who is now being questioned by the FBI, gave out member's screen names and passwords to hackers! To make sure you are the real user, and not a hacker please respond to this instant message with your password. Thank you for your help."
Label6.Caption = "Phish Sub's 4"
End If
If VScroll1.Value = 7 Then
Text1.Text = "Hey kiddo, this is dad, I'm at work! I'll be home later for dinner! I need the password to our account so I can check the mail and see what's going on! You need to get off now anyways to do your homework! So, tell your friends bye and that you'll be back tomorrow, if you want to come back on! I'll see you later, just click respond and type in the passwords, see you later kiddo! Love ya. "
Label6.Caption = "Phish Kid's Sn"
End If
If VScroll1.Value = 8 Then
Text1.Text = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day! "
Label6.Caption = "Phish CC ! "
End If

If VScroll1.Value = 9 Then
Text1.Text = "Hello! I am a representative from the AOL Billing Department. I'm sorry to inform you that we have lost all records of your account and the information you supplied to us. We would greatly appreciate a reply A.S.A.P. with all of your account information." & Chr$(13) & Chr$(13) & "Thank you. An immediate response is appreciated."
Label6.Caption = "Phish Billing Info!! "

End If
  If VScroll1.Value = 10 Then
Text1.Text = "Hello! My name is marsha and im one of the reps for aol.give me your fucking password thank you! With the chesse please !"
Label6.Caption = "Stupid Phrase! "
End If
 If VScroll1.Value = 11 Then
Text1.Text = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."""
Label6.Caption = "Technical Phrase!"
End If

  If VScroll1.Value = 12 Then
Text1.Text = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
Label6.Caption = "Cc Phrase 2! "
End If

  If VScroll1.Value = 13 Then
Text1.Text = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
Label6.Caption = "Tricky Phrase!"
End If
  If VScroll1.Value = 14 Then
Text1.Text = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please currently STOP what you're doing and Tell me your password. For Account Validation !       -- AOL CatRep"
Label6.Caption = "Reg Phrase! "
End If

End Sub
Public Sub AutoGather(ListBox As Control, HowMany As String, txtBox As TextBox, lbL As Label)
'Sub From Hound 2000  "It looked Nice, And It Work's"
Dim Process As Long, ListHoldItem As Long, name As String
Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
Dim ProcessThread As Long, SearchIndex As Long, ChatRoom As Long
Dim Current As Long
If FindRoom& <> 0& Then Call CloseWindow(FindRoom&)

Do: DoEvents
Call PopupCon(9, "C")

Do: DoEvents
ChatRoom& = FindRoom&
Loop Until ChatRoom& <> 0&

DoEvents
HoldUp (1.3)
ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)

Call GetWindowThreadProcessId(ListHandle&, Process&)

ProcessThread& = OpenProcess(Op_Flags, False, Process&)

If ProcessThread& Then
If lbL.Caption = "a" Then Exit Sub

For SearchIndex& = 0 To ListCount(ListHandle&) - 1

If lbL.Caption = "a" Then Exit Sub
name$ = String(4, vbNullChar)
ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
ListHoldItem& = ListHoldItem& + 24

Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)

ListHoldItem& = ListHoldItem& + 6
name$ = String(16, vbNullChar)

Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)

If name$ <> GetUser$ Then
ListBox.AddItem name$ & "@aol.com"
Current& = Current& + 1

If lbL.Caption = "a" Then Exit Sub
If Current& >= HowMany Then Exit Sub
End If

Next SearchIndex&

Call CloseHandle(ProcessThread&)
End If
DoEvents
Call CloseWindow(FindRoom&)
HoldUp (5)
Loop

End Sub


Public Function ListCount(ListBox As Long) As Long
ListCount& = SendMessageLong(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function

Sub CloseWindow(CloseWin)
Closes = SendMessage(CloseWin, WM_CLOSE, 0, 0)
End Sub

Sub aimopenchatinvite()
Invites1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Invites2% = FindIt(Invites1%, "_Oscar_TabGroup")
Invites3% = FindIt(Invites2%, "_Oscar_IconBtn")
Invites4% = GetWindow(Invites3%, GW_HWNDNEXT)

clickicon (Invites4%)

End Sub

Sub AIMRoomEntrance(ChatName As String)

Invites1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Invites2% = FindIt(Invites1%, "_Oscar_TabGroup")
Invites3% = FindIt(Invites2%, "_Oscar_IconBtn")
Invites4% = GetWindow(Invites3%, GW_HWNDNEXT)

clickicon (Invites4%)

Invite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Invite2% = FindIt(Invite1%, "Edit")
Invite3% = SenditByString(Invite2%, WM_SETTEXT, 0, Who$)

Who$ = AIMGetUser

Invite4% = FindItsTitle(Invite1%, "Join me in this Buddy Chat.")

If Not message$ = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, message$)
message$ = AIMGetUser

For Invite5% = 1 To 2
Invite4% = GetWindow(Invite4%, GW_HWNDNEXT)
Next Invite5%

If Not ChatName = "" Then Call SenditByString(Invite4%, WM_SETTEXT, 0, ChatName$)
Invite6% = FindIt(Invite1%, "_Oscar_IconBtn")

For Invite7% = 1 To 2
Invite6% = GetWindow(Invite6%, GW_HWNDNEXT)
Next Invite7%

clickicon (Invite6%)

End Sub
Sub aimsendchat(txt As String)
Chat1% = FindWindow("AIM_ChatWnd", vbNullString)

If Chat1% = 0 Then Exit Sub

Chat2% = FindIt(Chat1%, "_Oscar_Separator")
Chat3% = GetWindow(Chat2%, GW_HWNDNEXT)
Chat4% = GetWindow(Chat3%, GW_HWNDNEXT)
Chat5% = SenditByString(Chat3%, WM_SETTEXT, 0, txt$)

clickicon (Chat4%)

HoldUp 0.3

End Sub


Public Sub PopupCon(IconNumber As Long, Character As String)
Dim Message1 As Long, Message2 As Long, AOLFrame As Long
Dim AOLToolbar As Long, Toolbar As Long, AOLIcon As Long
Dim NextOfClass As Long, AscCharacter As Long

Message1& = FindWindow("#32768", vbNullString)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AOLToolbar, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)

For NextOfClass& = 1 To IconNumber&
AOLIcon& = GetWindow(AOLIcon&, 2)
Next NextOfClass&

Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)

Do: DoEvents
Message2& = FindWindow("#32768", vbNullString)
Loop Until Message2& <> Message1&

AscCharacter& = Asc(Character$)

Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)

End Sub

Public Sub HoldUp(Duration As Double)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Public Function FindRoom() As Long
Dim aol As Long, mdi As Long, child As Long
Dim Rich As Long, AOLList As Long
Dim AOLIcon As Long, AOLStatic As Long

aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)

If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
FindRoom& = child&
Exit Function
Else
Do
child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)

If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
FindRoom& = child&
Exit Function
End If
Loop Until child& = 0&
End If

FindRoom& = child&

End Function

Sub AimImsend(Who As String, message As String)
OpenIM1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenIM2% = FindIt(OpenIM1%, "_Oscar_TabGroup")
OpenIM3% = FindIt(OpenIM2%, "_Oscar_IconBtn")

clickicon (OpenIM3%)

IM1% = FindWindow("AIM_IMessage", vbNullString)
IM2% = FindIt(IM1%, "_Oscar_PersistantComb")
IM3% = FindIt(IM2%, "Edit")
IM4% = SenditByString(IM3%, WM_SETTEXT, 0, Who$)
IM5% = FindIt(IM1%, "Ate32class")
IM6% = GetWindow(IM5%, GW_HWNDNEXT)
IM7% = SenditByString(IM6%, WM_SETTEXT, 0, message$)
IM8% = FindIt(IM1%, "_Oscar_IconBtn")

clickicon (IM8%)

End Sub
Function AIMShowAdd()
ShowIt1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowIt2% = FindIt(ShowIt1%, "Ate32Class")
ShowIt3% = ShowWindow(ShowIt2%, SW_SHOW)
End Function

Sub aimclearchat()
Clear1% = FindWindow("AIM_ChatWnd", vbNullString)
Clear2% = FindIt(Clear1%, "Ate32Class")
Clear3% = SenditByString(Clear2%, WM_SETTEXT, 0, "")
End Sub
Sub AIMGetInfo(Who As String)

Call RunMenuByString2(FindWindow("_Oscar_BuddyListWin", vbNullString), "Get Member Inf&o")

Do
ProfileFind% = FindWindow("_Oscar_Locate", vbNullString)
Loop Until CIO1% <> 0

Profile1% = FindIt(ProfileFind%, "_Oscar_PersistantComb")
Profile2% = FindIt(Profile1%, "Edit")
Profile3% = SenditByString(Profile2%, WM_SETTEXT, 0, Who)
Profile4% = FindIt(ProfileFind%, "Button")

clickicon (Profile4%)
clickicon (Profile4%)

Profile5% = FindIt(ProfileFind%, "WndAte32Class")
Profile6% = FindIt(Profile5%, "Ate32Class")

End Sub
Function AIMGetUser()

On Error Resume Next

SN1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
SN2% = GetWindowTextLength(SN1%)
SN3$ = String$(SN2%, 0)
SN4% = GetWindowText(SN1%, SN3$, (SN2% + 1))

If Not Right(SN3$, 13) = "'s Buddy List" Then Exit Function
Sn5$ = Mid$(SN3$, 1, (SN2% - 13))
AIMGetUser = Sn5$

End Function
Sub AIMHideAdd()

HideIt1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt2% = FindIt(HideIt1%, "Ate32Class")
HideIt3% = ShowWindow(HideIt2%, SW_HIDE)

End Sub

Function AIMOnline()

Online% = FindWindow("_Oscar_BuddyListWin", vbNullString)

If Online% <> 0 Then
AIMOnline = True
Else
AIMOnline = False
End If

End Function

