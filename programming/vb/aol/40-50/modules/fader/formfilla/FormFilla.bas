Attribute VB_Name = "FormFilla"
'           ******  *****  ***** **   **
'           *       *   *  *   * * * * *
'           *****   *   *  ***** *  *  *
'           *       *   *  * *   *     *
'           *       *   *  *  *  *     *
'           *       *****  *   * *     *
'
'         *******  *****   *     *       *
'         *          *     *     *      * *
'         *****      *     *     *      ***
'         *          *     *     *      * *
'         *          *     *     *      * *
'         *        *****   ****  ****   * *
'

'                      Form Filla 1.3
'                          By:
'                         SpiDeR
'                     with help from
'                    HacK ,aka, ToXiD
'       Some Sub's Were Taken From Other Bas's
'
'           For The Text Faders u Do :
'   SendChat "<b>" & RedBlackRed("Form Filla 1.3 Kik's")
'   For the Form faders u do
'         call fadeform...(then what they want in ())
'Some Declarations Have nuthin do do with this bas
Global Const META_RECTANGLE = &H41B
Global Const META_SELECTOBJECT = &H12D
Global Const SRCCOPY = &HCC0020
Global Const Pi = 3.14159265359

Global R%       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Public majornum
Public minornum
Public timesload As String
Public introyn As String
Public scinyn As String
Public scexyn As String
Public soundyn As String
Public ik As String
Public ima As String
Public removedpeeps
Public soundsyn
Public chat
Public INTR
Public ICO
Public LB
Public EB
Public UL
'Public AOL
'Public MDi
Public nobut As Integer
Public access As String
Public fState As FormState
Public gFindString As String
Public gFindCase As Integer
Public gFindDirection As Integer
Public gCurPos As Integer
Public gFirstTime As Integer
Public tabclicked As String

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2

Public Const SPI_SCREENSAVERRUNNING = 97
Public Const ThisApp = "MDINote"
Public Const ThisKey = "Recent Files"
Public Const WM_CHAR = &H102
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
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_CREATE = &H1
Public Const WM_MDICREATE = &H220

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

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Const SM_CLEANBOOT = 67

'*****************************************************************************************
Type DEVMODE
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
'*****************************************************************************************
Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wId As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
'*****************************************************************************************
Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'*****************************************************************************************
Type pointapi
   X As Long
   Y As Long
End Type
'*****************************************************************************************
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
'*****************************************************************************************
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
'*****************************************************************************************
Public Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
End Type
'*****************************************************************************************
Public Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
End Type
'*****************************************************************************************
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, HostLen&) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
'*****************************************************************************************
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
'*****************************************************************************************
Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'*****************************************************************************************
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'*****************************************************************************************
Declare Function sndplatsound Lib "mmsystem.dll" (ByVal wavfile As Any, ByVal wFlags As Integer) As Integer '<--All one line
'*****************************************************************************************
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'*****************************************************************************************
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'*****************************************************************************************
'Public Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long)as long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, LPRect As Rect) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, LPRect As Rect) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As pointapi) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Integer, LPRect As Rect, ByVal hBrush As Integer) As Integer
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function movewindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetRect Lib "user32" (LPRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function setparent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function Getmenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function extfn0BD2 Lib "user32" Alias "SendMessageA" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn0144 Lib "user32" Alias "SendMessageA" (ByVal p1%, ByVal p2%, ByVal p3%, p4&) As Long
'*****************************************************************************************
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'*****************************************************************************************
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Long
Declare Function StretchBlt% Lib "gdi32" (ByVal hdc%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Option Explicit
Dim nXCoord(50) As Integer
Dim nYCoord(50) As Integer
Dim nXSpeed(50) As Integer
Dim nYSpeed(50) As Integer



Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
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
  GREEN As Long
  BLUE As Long
End Type
Sub CoolBeginning(e As Form)

'Call CoolBeginning(Me)
'Sub made by ToXiD
'Change A few colors and Positions So They Match Ur Form
e.Visible = True
e.Enabled = True
e.Left = "-30"
e.Top = "1170"
e.Height = "0"
Call IceFade(e)
Pause (0.1)
e.Left = "30"
e.Top = "1250"
e.Height = "100"
Call FireFade(e)
Call IceFade(e)
Pause (0.1)
e.Left = "130"
e.Top = "1300"
e.Height = "150"
Call FadeFormBlue(e)
Pause (0.1)
e.Left = "200"
e.Top = "1350"
e.Height = "350"
Call FadeFormYellow(e)
Pause (0.1)
e.Left = "400"
e.Top = "1400"
e.Height = "400"
Call FadeFormRed(e)
Pause (0.1)
e.Left = "430"
e.Top = "1450"
e.Height = "450"
Call FadeFormGreen(e)
Pause (0.1)
e.Left = "530"
e.Top = "1500"
e.Height = "500"
Call FadeFormPurple(e)
Pause (0.1)
e.Left = "630"
e.Top = "1550"
e.Height = "550"
Call FadeFormWhitePinkRed(e)
Pause (0.1)
e.Left = "830"
e.Top = "1600"
e.Height = "600"
Call FadeForm(e, vbRed, vbBlue)
Pause (0.1)
e.Top = "1650"
e.Left = "1030"
e.Height = "650"
Call FadeForm(e, vbRed, vbGreen)
Pause (0.1)
e.Top = "1700"
e.Left = "1230"
e.Height = "700"
Call FadeForm(e, vbRed, vbYellow)
Pause (0.1)
e.Top = "1750"
e.Left = "1430"
e.Height = "750"
Call FadeForm(e, vbBlue, vbRed)
Pause (0.1)
e.Top = "1800"
e.Left = "1630"
e.Height = "800"
Call FadeForm(e, vbGreen, vbRed)
Pause (0.1)
e.Top = "1850"
e.Left = "1830"
e.Height = "850"
Call FadeForm(e, vbGreen, vbRed)
Pause (0.1)
e.Top = "1900"
e.Left = "2030"
e.Height = "900"
Pause (0.1)
e.Top = "1910"
e.Left = "2230"
e.Height = "950"
Pause (0.1)
e.Top = "1920"
e.Left = "2430"
e.Height = "1000"
Pause (0.1)
e.Top = "1930"
e.Left = "2630"
e.Height = "1020"
Pause (0.1)
e.Top = "1940"
e.Left = "2830"
e.Height = "1040"
Pause (0.1)
e.Top = "1950"
e.Left = "3030"
e.Height = "1060"
Pause (0.1)
e.Height = "1080"
Pause (0.1)
e.Enabled = True


End Sub
Function CoolBeginningAgain(e As Form)

'Call CoolBeginningagain(Me)
'Sub made by ToXiD
Do
e.Visible = True
e.Enabled = True
e.Left = "-30"
e.Top = "1170"
e.Height = "0"
Call IceFade(e)
Pause (0.1)
e.Left = "30"
e.Top = "1250"
e.Height = "100"
Call FireFade(e)
Pause (0.1)
e.Left = "130"
e.Top = "1300"
e.Height = "150"
Call FadeFormBlue(e)
Pause (0.1)
e.Left = "200"
e.Top = "1350"
e.Height = "350"
Call FadeFormYellow(e)
Pause (0.1)
e.Left = "400"
e.Top = "1400"
e.Height = "400"
Call FadeFormRed(e)
Pause (0.1)
e.Left = "430"
e.Top = "1450"
e.Height = "450"
Call FadeFormGreen(e)
Pause (0.1)
e.Left = "530"
e.Top = "1500"
e.Height = "500"
Call FadeFormPurple(e)
Pause (0.1)
e.Left = "630"
e.Top = "1550"
e.Height = "550"
Call FadeFormWhitePinkRed(e)
Pause (0.1)
e.Left = "830"
e.Top = "1600"
e.Height = "600"
Call FadeForm(e, vbRed, vbBlue)
Pause (0.1)
e.Top = "1650"
e.Left = "1030"
e.Height = "650"
Call FadeForm(e, vbRed, vbGreen)
Pause (0.1)
e.Top = "1700"
e.Left = "1230"
e.Height = "700"
Call FadeForm(e, vbRed, vbYellow)
Pause (0.1)
e.Top = "1750"
e.Left = "1430"
e.Height = "750"
Call FadeForm(e, vbBlue, vbRed)
Pause (0.1)
e.Top = "1800"
e.Left = "1630"
e.Height = "800"
Call FadeForm(e, vbGreen, vbRed)
Pause (0.1)
e.Top = "1850"
e.Left = "1830"
e.Height = "850"
Call FadeForm(e, vbYellow, vbRed)
Pause (0.1)
e.Top = "1900"
e.Left = "2030"
e.Height = "900"
Pause (0.1)
e.Top = "1910"
e.Left = "2230"
e.Height = "950"
Pause (0.1)
e.Top = "1920"
e.Left = "2430"
e.Height = "1000"
Pause (0.1)
e.Top = "1930"
e.Left = "2630"
e.Height = "1020"
Pause (0.1)
e.Top = "1940"
e.Left = "2830"
e.Height = "1040"
Pause (0.1)
e.Top = "1950"
e.Left = "3030"
e.Height = "1060"
Pause (0.1)
e.Height = "1080"
Pause (0.1)
e.Enabled = True
Loop

End Function


Sub FadeFormBlink(theform As Form)
'Really Cool
Dim a
Dim i
Dim B
Dim Y

theform.BackColor = &H0&
theform.DrawStyle = 6
theform.DrawMode = 13

theform.DrawWidth = 2
theform.ScaleMode = 3
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, B)-(theform.Width, B + 2), RGB(a + 3, a, a * 3), BF

B = B + 2
Next a

For i = 255 To 0 Step -1
theform.Line (0, 0)-(theform.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i

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
Sub BlueFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub

Sub GreenFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub RedFade(vForm As Object)
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
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub FireFade(vForm As Object)
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
Sub FormDrawSunLine(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbInches
    vForm.DrawWidth = 15
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (45, intLoop)-(Screen.Width, intLoop - 7), RGB(252, 255 - intLoop, 56 - intLoop), B  'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub FadeFormWhitePinkRed(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = 4 'Set Form Modes
    vForm.DrawMode = 13
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
            'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
       
    vForm.DrawStyle = 4 'Set Form Modes
    vForm.DrawMode = 13
    vForm.ScaleMode = 1
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 345 - intLoop), B 'Draw boxes with specified color of loop

 Next intLoop
End Sub
Sub FadeFormRustLine(vForm As Form)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = 4 'Set Form Modes
    vForm.DrawMode = 13
    vForm.ScaleMode = 1
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(135, 255 - intLoop, 145 - intLoop), B  'Draw boxes with specified color of loop

 Next intLoop
End Sub

Sub PlatinumFade(vForm As Object)
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
Sub IceFade(vForm As Object)
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
Sub FadeFormAquaLine(vForm As Form)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 333 - intLoop, 222), B 'Draw boxes with specified color of loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 123 - intLoop, 432), B 'Draw boxes with specified color of loop
                   'This code works best when called in the paint event
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 431 - intLoop, 419), B 'Draw boxes with specified color of loop
                   'This code works best when called in the paint event
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 153 - intLoop, 442), B 'Draw boxes with specified color of loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 134 - intLoop, 134), B 'Draw boxes with specified color of loop
        Next intLoop
        
End Sub
Sub FormDance(m As Form)

'  This makes a form dance across the screen
m.Left = 5
timeout (0.1)
m.Left = 400
timeout (0.1)
m.Left = 700
timeout (0.1)
m.Left = 1000
timeout (0.1)
m.Left = 2000
timeout (0.1)
m.Left = 3000
timeout (0.1)
m.Left = 4000
timeout (0.1)
m.Left = 5000
timeout (0.1)
m.Left = 4000
timeout (0.1)
m.Left = 3000
timeout (0.1)
m.Left = 2000
timeout (0.1)
m.Left = 1000
timeout (0.1)
m.Left = 700
timeout (0.1)
m.Left = 400
timeout (0.1)
m.Left = 5
timeout (0.1)
m.Left = 400
timeout (0.1)
m.Left = 700
timeout (0.1)
m.Left = 1000
timeout (0.1)
m.Left = 2000

End Sub
Sub FadeForm(FormX As Form, Colr1, Colr2)
Dim b1
Dim g1
Dim r1
Dim b2
Dim g2
Dim r2

'by monk-e-god (modified from a sub by MaRZ)
    b1 = GetRGB(Colr1).BLUE
    g1 = GetRGB(Colr1).GREEN
    r1 = GetRGB(Colr1).red
    b2 = GetRGB(Colr2).BLUE
    g2 = GetRGB(Colr2).GREEN
    r2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((r2 - r1) / 255 * intLoop) + r1, ((g2 - g1) / 255 * intLoop) + g1, ((b2 - b1) / 255 * intLoop) + b1), B
    Next intLoop
End Sub
Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.BLUE = Int(CVal / 65536)
  GetRGB.GREEN = Int((CVal - (65536 * GetRGB.BLUE)) / 256)
  GetRGB.red = CVal - (65536 * GetRGB.BLUE + 256 * GetRGB.GREEN)
End Function
Sub FadeFormBlackToWhite(vForm As Form)
vForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
vForm.Line (0, B)-(vForm.Width, B + 1), RGB(a + 1, a, a * 1), BF
B = B + 2
Next a
End Sub
Sub FormFadeSunset(theform As Form)
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, B)-(theform.Width, B + 55), RGB(a + 99, a, a * 0), BF
B = B + 2
Next a
End Sub
Sub FadeFormPink2Red(theform As Form)
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, B)-(theform.Width, B + 55), RGB(a + 225, a, a * 3), BF
B = B + 2
Next a
End Sub
Sub FadeFormHorizon(theform As Form)

theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, B)-(theform.Width, B + 2), RGB(a + 3, a, a * 3), BF
B = B + 2
Next a
End Sub


Function Yellow_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Yellow_LBlue = msg
End Function
    
Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowRedYellow = msg
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowRed = msg
End Function
Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowPurpleYellow = msg
End Function
Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowPurple = msg
End Function
Function YellowPink(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(78, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowPink = msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowGreenYellow = msg
End Function
Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowGreen = msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlueYellow = msg
End Function
Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlue = msg
End Function
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlackYellow = msg
End Function
Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    YellowBlack = msg
End Function
Function WhitePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    WhitePurple = msg
End Function
Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRB = P$
End Function

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Function

Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
SendChat (P$)
End Sub
Function Wavy(thetext As String)

G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    U$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & R$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next w
Wavy = P$

End Function

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



Function RGB2HEX(R, G, B)
    Dim X&
    Dim XX&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For XX& = 1 To 2
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
        Next XX&
    Next X&
    Configuring$ = (Configuring$)
    RGB2HEX = Configuring$
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedYellowRed = msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedYellow = msg
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedPurpleRed = msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedPurple = msg
End Function
Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlueRed = msg
End Function
Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlue = msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlackRed = msg
End Function
Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    RedBlack = msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleYellowPurple = msg
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleYellow = msg
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleRedPurple = msg
End Function
Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleRed = msg
End Function
Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleGreenPurple = msg
End Function
Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleGreen = msg
End Function
Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBluePurple = msg
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlue = msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlackPurple = msg
End Function
Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleBlack = msg
End Function

Function Purple_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Purple_LBlue = msg
End Function

Function LBlue_Yellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Yellow = msg
End Function
Function LBlue_Green(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Green = msg
End Function
Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyYellowGrey = msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyYellow = msg
End Function
Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyRedGrey = msg
End Function
Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyRed = msg
End Function
Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyPurpleGrey = msg
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyPurple = msg
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyGreenGrey = msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyGreen = msg
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlueGrey = msg
End Function
Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlue = msg
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlackGrey = msg
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / a
        f = e * B
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreyBlack = msg
End Function
Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenYellowGreen = msg
End Function
Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenYellow = msg
End Function
Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenRedGreen = msg
End Function
Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenRed = msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenPurpleGreen = msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenPurple = msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlueGreen = msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlue = msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlackGreen = msg
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    GreenBlack = msg
End Function
Function DBlue_Black(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    DBlue_Black = msg
End Function
Public Sub CenterFormTop(frm As Form)
' this function will center your form and also keep
' it on top of the screen
' to use type - CenterFormTop Me ( in form_load )
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueYellowBlue = msg
End Function
Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueYellow = msg
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueRedBlue = msg
End Function


Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BluePurpleBlue = msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueRed = msg
End Function
Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BluePurple = msg
End Function
Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueGreenBlue = msg
End Function
Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueGreen = msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueBlackBlue = msg
End Function


Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlueBlack = msg
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackYellowBlack = msg
End Function
Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackYellow = msg
End Function
Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackRedBlack = msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackRed = msg
End Function
Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackPurpleBlack = msg
End Function
Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackPurple = msg
End Function
Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGreyBlack = msg
End Function
Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / a
        f = e * B
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    BlackGrey = msg
End Function
Function Black_LBlue_Black(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Black_LBlue_Black = msg
End Function

Function Black_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Black_LBlue = msg
End Function



Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
  
End Function

Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    WhitePurpleWhite = msg
End Function

Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Green_LBlue = msg
End Function

Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Yellow_LBlue = msg
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Purple_LBlue_Purple = msg
End Function

Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 450 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    DBlue_Black_DBlue = msg
End Function

Function DGreen_Black(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(0, f - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    DGreen_Black = msg
End Function



Function LBlue_Orange(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Orange = msg
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_Orange_LBlue = msg
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 220 / a
        f = e * B
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LGreen_DGreen = msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LGreen_DGreen_LGreen = msg
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        f = e * B
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_DBlue = msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    LBlue_DBlue_LBlue = msg
End Function

Function PinkOrange(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / a
        f = e * B
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PinkOrange = msg
End Function

Function PinkOrangePink(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 490 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PinkOrangePink = msg
End Function

Function PurpleWhite(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 200 / a
        f = e * B
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleWhite = msg
End Function

Function PurpleWhitePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    PurpleWhitePurple = msg
End Function

Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        f = e * B
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & h & ">" & D
    Next B
    Yellow_LBlue_Yellow = msg
End Function
Sub KeyDown()
'--= Sub by, ToxiD =--
'COPY THIS WHOLE CODE DOWN AND PASTE IT IN FORM_KEYDOWN
'OR IT WILL NOT WORK AND I REPEAT WILL NOT WORK!
Rem "Call KeyDown...That doesnt work 'make msgbox stuff"
Dim msg As String
Dim Title As String
Dim Style

Title = "What key?"
Style = vbOKOnly + vbInformation
'determine what key is pressed
'instead of having the program play a message box, you can add other
'things like sound, and use the sound call soundplay
Select Case KeyCode
Case vbKeyA
msg = "Key A"
Case vbKeyB
msg = "Key B"
Case vbKeyC
msg = "Key C"
Case vbKeyD
msg = "Key D"
Case vbKeyE
msg = "Key E"
Case vbKeyF
msg = "Key F"
Case vbKeyG
msg = "Key G"
Case vbKeyH
msg = "Key H"
Case vbKeyI
msg = "Key I"
Case vbKeyJ
msg = "Key J"
Case vbKeyK
msg = "Key K"
Case vbKeyL
msg = "Key L"
Case vbKeyM
msg = "Key M"
Case vbKeyN
msg = "Key N"
Case vbKeyO
msg = "Key O"
Case vbKeyP
msg = "Key P"
Case vbKeyQ
msg = "Key Q"
Case vbKeyR
msg = "Key R"
Case vbKeyS
msg = "Key S"
Case vbKeyT
msg = "Key T"
Case vbKeyU
msg = "Key U"
Case vbKeyV
msg = "Key V"
Case vbKeyW
msg = "Key W"
Case vbKeyX
msg = "Key X"
Case vbKeyY
msg = "Key Y"
Case vbKeyZ
msg = "Key Z"
Case vbKey1
msg = "Key 1"
Case vbKeyReturn
msg = "Key Enter"
Case vbKeyRight
msg = "Key Right"
Case vbKeyLeft
msg = "Key Left"
Case vbKeyUp
msg = "Key Up"
Case vbKeyDown
msg = "Key Down"
Case vbKeyTab
msg = "Tab"
Case vbKey1
msg = "1"
Case vbKey2
msg = "2"
Case vbKey3
msg = "3"
Case vbKey4
msg = "4"
Case vbKey5
msg = "5"
Case vbKey6
msg = "6"
Case vbKey7
msg = "7"
Case vbKey8
msg = "8"
Case vbKey9
msg = "9"
Case vbKey0
msg = "0"
Case vbKeyEnd
msg = "End"
Case vbKeyHome
msg = "Home"
Case vbKeyInsert
msg = "Insert"
Case vbKeyPageUp
msg = "PageUp"
Case vbKeyPageDown
msg = "PageDown"
Case vbKeyDelete
msg = "Delete"
Case vbKeyAdd
msg = "Add"
Case vbKeySubtract
msg = "Subtract"
Case vbKeyNumlock
msg = "Numlock"
Case vbKeyShift
msg = "Shift"
Case vbKeyEscape
msg = "Escape"
Case vbKeyControl
msg = "CTRL"
Case vbKeyF1
msg = "F1"
Case vbKeyF2
msg = "F2"
Case vbKeyF3
msg = "F3"
Case vbKeyF4
msg = "F4"
Case vbKeyF5
msg = "F5"
Case vbKeyF6
msg = "F6"
Case vbKeyF7
msg = "F7"
Case vbKeyF8
msg = "F8"
Case vbKeyF9
msg = "F9"
Case vbKeyF10
msg = "F10"
Case vbKeyF11
msg = "F11"
Case vbKeyF12
msg = "F12"
Case vbKeyScrollLock
msg = "ScrollLock"
Case vbKeyDelete
msg = "Delete"
End Select
MsgBox msg, Style, Title
End Sub
Function FadeFormNeon(vForm As Form)
Call FadeForm(vForm, vbYellow, vbGreen)

End Function

Sub FormBevelLines(FormFrame As Form, side, wId, Color)

       '     ' This Sub is called by FormInner/Outer Bevel to draw the
       '     ' lines for FormInnerBevel and FormOuterBevel
       Dim X1, Y1, x2, Y2 As Integer
       Dim rightX, bottomY
       Dim dx1, dx2, dy1, dy2 As Integer
       Dim i
       rightX = FormFrame.ScaleWidth - 1
       bottomY = FormFrame.ScaleHeight - 1
       Select Case side
       Case 0 'Left side
       X1 = 0: dx1 = 1
       x2 = 0: dx2 = 1
       Y1 = 0: dy1 = 1
       Y2 = bottomY + 1: dy2 = -1
       Case 1 'Right side
       X1 = rightX: dx1 = -1
       x2 = X1: dx2 = dx1
       Y1 = 0: dy1 = 1
       Y2 = bottomY + 1: dy2 = -1
       Case 2 'Top side
       X1 = 0: dx1 = 1
       x2 = rightX: dx2 = -1
       Y1 = 0: dy1 = 1
       Y2 = 0: dy2 = 1
       Case 3 'Bottom side
       X1 = 1: dx1 = 1
       x2 = rightX + 1: dx2 = -1
       Y1 = bottomY: dy1 = -1
       Y2 = Y1: dy2 = dy1
End Select


For i = 1 To wId

              FormFrame.Line (X1, Y1)-(x2, Y2), Color
                     X1 = X1 + dx1
                     x2 = x2 + dx2
                     Y1 = Y1 + dy1
                     Y2 = Y2 + dy2
              Next i

End Sub

'     'Here are the 2 main routines:

Sub FormOuterBevel(FormFrame As Form, BevelWidth As Integer)

       '     ' This sub draws raised bevels on a Form
       '     '
       '     ' Parameters TypeComments
       '     'FormFrameFormthe Form to bevel
       '     'BevelWidthintegerwidth of bevel in pixels
              FormFrame.ScaleMode = 3 ' Pixels

                            FormBevelLines FormFrame, 0, BevelWidth, QBColor(15) 'White

                                          FormBevelLines FormFrame, 1, BevelWidth, QBColor(8) 'D.Gray

                                                        FormBevelLines FormFrame, 2, BevelWidth, QBColor(15) 'White

                                                                      FormBevelLines FormFrame, 3, BevelWidth, QBColor(8) 'D.Gray
                                                                      End Sub


Sub FormInnerBevel(FormFrame As Form, BevelWidth As Integer)

       '     ' This sub draws recessed bevels on a Form
       '     '
       '     ' Parameters TypeComments
       '     'FormFrameFormthe Form to bevel
       '     'BevelWidthintegerwidth of bevel in pixels
       '     '

              FormFrame.ScaleMode = 3 ' Pixels

                            FormBevelLines FormFrame, 0, BevelWidth, QBColor(8) 'D.Gray
                                          FormBevelLines FormFrame, 1, BevelWidth, QBColor(15) 'White

                                                        FormBevelLines FormFrame, 2, BevelWidth, QBColor(8)

                                                                      FormBevelLines FormFrame, 3, BevelWidth, QBColor(15)
                                                                      End Sub



Sub ExitDown(Form As Form)


              Do
                           Form.Top = Trim(Str(Int(Form.Top) + 300))
        
                            DoEvents
                            
                                    Loop Until Form.Top > 7200

                                          If Form.Top > 7200 Then End
                                          End Sub


Sub ExitLeft(Form As Form)


              Do
                           Form.Left = Trim(Str(Int(Form.Left) - 300))

                            DoEvents
                                    Loop Until Form.Left < -6300

                                          If Form.Left < -6300 Then End
                                          End Sub


Sub ExitRight(Form As Form)


              Do
                           Form.Left = Trim(Str(Int(Form.Left) + 300))

                            DoEvents
                                    Loop Until Form.Left > 9600

                                          If Form.Left > 9600 Then End
                                          
                                          End Sub
Function ExitForm(theform As Form)
Unload theform

End Function



Sub ExitUp(Form As Form)


              Do
                        Form.Top = Trim(Str(Int(Form.Top) - 300))

                            DoEvents
                                    Loop Until Form.Top < -4500

                                          If Form.Top < -4500 Then End
                                          
                                          
                                          End Sub
Function MiniWindow(vForm As Form)
vForm.WindowState = 1

End Function
Function MaxWindow(vForm As Form)
vForm.WindowState = 2

End Function
Function M_RegularWindow(vForm As Form)
vForm.WindowState = 0
End Function
Sub Funnytext(l As label)
'Try this out its cool
'Sub by 'Toxid
'Example: Make a label on a form
'And in for_click Put the code:
'Call funnytext(Label1) or what number
l.FontSize = 33
l.ForeColor = vbRed
Call timeout(0.01)
l.FontSize = 31
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 29
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 27
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 25
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 23
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 21
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 19
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 17
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 15
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 13
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 11
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 9
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 7
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 5
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 3
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 1
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 33
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 31
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 29
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 27
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 25
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 23
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 21
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 19
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 17
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 15
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 13
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 11
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 9
l.ForeColor = vbBlue
Call Pause(0.01)
l.FontSize = 7
l.ForeColor = vbGreen
Call Pause(0.01)
l.FontSize = 5
l.ForeColor = vbYellow
Call Pause(0.01)
l.FontSize = 3
l.ForeColor = vbRed
Call Pause(0.01)
l.FontSize = 1
l.ForeColor = vbBlue
End Sub
' The following are the color HTML codes
' for these colors --------------
'
' Red = FE0000
' Blue=0000FE
' Green=00FE00
' dkBlue=000066
' orange=FE7C00
' White=FEFEFE
' purple=C200C2
' yellow=FEFE00
' DkRed=660000
'
Function AddCustomColorToText(Text As String, Blend As Boolean, Wavey As Boolean, Lagger As Boolean, Bold As Boolean, Italics As Boolean, Strikeout As Boolean, Underline As Boolean, CC1 As Long, CC2 As Long) As String
If Text = "" Then Exit Function
wavestep = 1
MaxColor = &HFE
MinColor = &H0
txtsize = Len(Text)
' set colors...
Dim red, GREEN, BLUE, ERED, EGREEN, EBLUE
Dim REDBL, GREENBL, BLUEBL
Dim SRED As String, SGREEN As String, SBLUE As String
Dim CustomC1Str As String, CustomC2Str As String
CustomC1Str = Hex$(CC1)
CustomC2Str = Hex$(CC2)
If Len(CustomC1Str) < 6 Then
    If Len(CustomC1Str) = 5 Then CustomC1Str = "0" & CustomC1Str
    If Len(CustomC1Str) = 4 Then CustomC1Str = "00" & CustomC1Str
    If Len(CustomC1Str) = 3 Then CustomC1Str = "000" & CustomC1Str
    If Len(CustomC1Str) = 2 Then CustomC1Str = "0000" & CustomC1Str
    If Len(CustomC1Str) = 1 Then CustomC1Str = "00000" & CustomC1Str
End If

If Len(CustomC2Str) < 6 Then
    If Len(CustomC2Str) = 5 Then CustomC2Str = "0" & CustomC2Str
    If Len(CustomC2Str) = 4 Then CustomC2Str = "00" & CustomC2Str
    If Len(CustomC2Str) = 3 Then CustomC2Str = "000" & CustomC2Str
    If Len(CustomC2Str) = 2 Then CustomC2Str = "0000" & CustomC2Str
    If Len(CustomC2Str) = 1 Then CustomC2Str = "00000" & CustomC2Str
End If

red = Right(CustomC1Str, 2)
GREEN = Mid(CustomC1Str, 3, 2)
BLUE = Left(CustomC1Str, 2)
ERED = Right(CustomC2Str, 2)
EGREEN = Mid(CustomC2Str, 3, 2)
EBLUE = Left(CustomC2Str, 2)

red = Val("&H" & red)
GREEN = Val("&H" & GREEN)
BLUE = Val("&H" & BLUE)
ERED = Val("&H" & ERED)
EGREEN = Val("&H" & EGREEN)
EBLUE = Val("&H" & EBLUE)


' If Blend is true then takes you to the
' blending function

FinalTxt$ = ""

If Blend = True Then GoTo BlendTxt

' Blends colors from startcolor to endcolor

REDBL = -(Int(((red - ERED)) / txtsize))
GREENBL = -(Int(((GREEN - EGREEN)) / txtsize))
BLUEBL = -(Int(((BLUE - EBLUE)) / txtsize))

For X = 1 To txtsize

SRED = MakeHexString(red)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(Text, X, 1)
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If Underline = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
red = red + REDBL
GREEN = GREEN + GREENBL
BLUE = BLUE + BLUEBL
If red > 254 Then red = 254
If red < 0 Then red = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

AddCustomColorToText = FinalTxt$

Exit Function


BlendTxt:

If (Len(Text) / 2) <> (Abs(Len(Text) / 2)) Then txtsize = txtsize - 1


REDBL = -(Int(((red - ERED)) / txtsize))
GREENBL = -(Int(((GREEN - EGREEN)) / txtsize))
BLUEBL = -(Int(((BLUE - EBLUE)) / txtsize))

REDBL = (Int(REDBL * 2))
GREENBL = (Int(GREENBL * 2))
BLUEBL = (Int(BLUEBL * 2))

For X = 1 To Int(txtsize / 2)

SRED = MakeHexString(red)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(Text, X, 1)
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If Underline = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1

red = red + REDBL
GREEN = GREEN + GREENBL
BLUE = BLUE + BLUEBL
If red > 254 Then red = 254
If red < 0 Then red = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

For X = (Int(txtsize / 2)) + 1 To txtsize
If Lagger = True Then
FinalTxt$ = FinalTxt$ + "<HTML><HTML><HTML><HTML>"
If Bold = True Then FinalTxt$ = FinalTxt$ + "<B>"
If Italics = True Then FinalTxt$ = FinalTxt$ + "<I>"
If Strikeout = True Then FinalTxt$ = FinalTxt$ + "<S>"
If Underline = True Then FinalTxt$ = FinalTxt$ + "<U>"
End If
SRED = MakeHexString(red)
SGREEN = MakeHexString(GREEN)
SBLUE = MakeHexString(BLUE)
If wavestep = 5 Then wavestep = 1
If Wavey = True And wavestep = 1 Then FinalTxt$ = FinalTxt$ + "<SUP>"
If Wavey = True And wavestep = 2 Then FinalTxt$ = FinalTxt$ + "</SUP>"
If Wavey = True And wavestep = 3 Then FinalTxt$ = FinalTxt$ + "<SUB>"
If Wavey = True And wavestep = 4 Then FinalTxt$ = FinalTxt$ + "</SUB>"
wavestep = wavestep + 1

FinalTxt$ = FinalTxt$ + "<FONT COLOR=""#" & SRED & SGREEN & SBLUE & """>" & Mid(Text, X, 1)

red = red - REDBL
GREEN = GREEN - GREENBL
BLUE = BLUE - BLUEBL
If red > 254 Then red = 254
If red < 0 Then red = 0
If GREEN > 254 Then GREEN = 254
If GREEN < 0 Then GREEN = 0
If BLUE > 254 Then BLUE = 254
If BLUE < 0 Then BLUE = 0

Next X

AddCustomColorToText = FinalTxt$

End Function


Sub HowTo_CircleForm()

'THIS CODE WILL MAKE CIRCLE OR OVEL SHAPED FORMS
'PLACE IN THE FORM LOAD
'SetWindowRgn hwnd, _
'  CreateEllipticRgn(0, 0, 300, 200), True
'
End Sub


Sub HowTO_StarField()
'this is how to make a star field

'put in Private Sub Form_Load()
'    Dim nIndex As Integer
'    ' At form load, the initial coordinates of all the stars needs to be set to off screen.
'    ' The timer event will recognise this and bring the stars back on screen.
'    For nIndex = 0 To 49
'        nXCoord(nIndex) = -1
'        nYCoord(nIndex) = -1
'    Next
'    ' Call the randomize method to tell VB to get ready to think of some random numbers
'    Randomize
'    Timer1.Enabled = True
'End Sub
'

'Private Sub Timer1_Timer()
'    ' The timer event performs three functions here.
'    '       1. Stars that are off screen are remade at the centre of the screen
'    '       2. Stars previously drawn are erase by redrawing them in black
'    '       3. Each star's position is recalculated and the star redrawn.
'    Dim nIndex As Integer
'    For nIndex = 0 To 49
'        'erase the previously drawn star
'        PSet (nXCoord(nIndex), nYCoord(nIndex)), &H0&
'        ' If the star number nIndex is off screen, then bring it back
'        If nXCoord(nIndex) < 0 Or nXCoord(nIndex) > frmMain.ScaleWidth Or nYCoord(nIndex) < 0 Or nYCoord(nIndex) > frmMain.ScaleHeight Then
'            nXCoord(nIndex) = frmMain.ScaleWidth \ 2
'            nYCoord(nIndex) = frmMain.ScaleHeight \ 2
'            ' Decide on some random speeds for the new star
'            nXSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
'            nYSpeed(nIndex) = Int(Rnd(1) * 200) - 100   ' Gives a speed between -100 and 100
'        End If
'        ' Now redraw the star so that it appears to move
'        nXCoord(nIndex) = nXCoord(nIndex) + nXSpeed(nIndex)
'        nYCoord(nIndex) = nYCoord(nIndex) + nYSpeed(nIndex)
'        PSet (nXCoord(nIndex), nYCoord(nIndex)), &HFFFFFF
'    ' Move on to the next star
'    Next
'End Sub
End Sub

Sub Form_Explode(f As Form, Movement As Integer)
'this will explode a form
    Dim myRect As Rect
    Dim formWidth%, formHeight%, i%, X%, Y%, cX%, cY%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
    For i = 1 To Movement
        cX = formWidth * (i / Movement)
        cY = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cX) / 2
        Y = myRect.Top + (formHeight - cY) / 2
        Rectangle TheScreen, X, Y, X + cX, Y + cY
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub Form_Implode(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'The larger the "Movement" value, the slower the "Implosion"
    Dim myRect As Rect
    Dim formWidth%, formHeight%, i%, X%, Y%, cX%, cY%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        cX = formWidth * (i / Movement)
        cY = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cX) / 2
        Y = myRect.Top + (formHeight - cY) / 2
        Rectangle TheScreen, X, Y, X + cX, Y + cY
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub


Sub Form_ScrollDown(frm As Form, startNUM, endNUM)
'This will make the form slowly scroll down
'You can use a timeout to stop it and put it in a
'timer
Dim X
Dim Y
frm.Show
frm.Height = startNUM
X = frm.Height
For Y = X To endNUM
frm.Height = frm.Height + 20
timeout (0.0001)
If frm.Height >= endNUM Then GoTo out:
Next Y
out:
End Sub

Sub Form_ScrollUp(frm As Form, startNUM, endNUM)
'This will make the form slowly scroll up
'You can use a timeout to stop it and put it in a
'timer
Dim X
Dim Y
frm.Show
frm.Height = startNUM
X = frm.Height
For Y = X To endNUM
frm.Height = frm.Height - 20
timeout (0.0001)
'If frm.Height <= endNUM Then GoTo out:
Next Y
out:
End Sub

Sub Form_Suckin(frm As Form)

Do
DoEvents
frm.Height = frm.Height - 50
frm.Width = frm.Width - 50
Loop Until frm.Height < 450 And frm.Width < 1700
End Sub


Sub FormDraw3DBorder(f As Form)
'adds a 3d border to a form

Dim iOldScaleMode As Integer
Dim iOldDrawWidth As Integer
    iOldScaleMode = f.ScaleMode
    iOldDrawWidth = f.DrawWidth
    f.ScaleMode = vbPixels
    f.DrawWidth = 1
    f.Line (0, 0)-(f.ScaleWidth, 0), QBColor(15)
    f.Line (0, 0)-(0, f.ScaleHeight), QBColor(15)
    f.Line (0, f.ScaleHeight - 1)-(f.ScaleWidth - 1, f.ScaleHeight - 1), QBColor(8)
    f.Line (f.ScaleWidth - 1, 0)-(f.ScaleWidth - 1, f.ScaleHeight), QBColor(8)

    f.ScaleMode = iOldScaleMode
    f.DrawWidth = iOldDrawWidth
End Sub

Public Sub Form_MakeTransparent(frm As Form)
'makes a form transparent.....
       Dim rctClient As Rect, rctFrame As Rect
       Dim hClient As Long, hFrame As Long
       '     '// Grab client area and frame area
       GetWindowRect frm.hwnd, rctFrame
       GetClientRect frm.hwnd, rctClient
       '     '// Convert client coordinates to screen coordinates
       Dim lpTL As pointapi, lpBR As pointapi
       lpTL.X = rctFrame.Left
       lpTL.Y = rctFrame.Top
       lpBR.X = rctFrame.Right
       lpBR.Y = rctFrame.Bottom
       ScreenToClient frm.hwnd, lpTL
       ScreenToClient frm.hwnd, lpBR
       rctFrame.Left = lpTL.X
       rctFrame.Top = lpTL.Y
       rctFrame.Right = lpBR.X
       rctFrame.Bottom = lpBR.Y
       rctClient.Left = Abs(rctFrame.Left)
       rctClient.Top = Abs(rctFrame.Top)
       rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
       rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
       rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
       rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
       rctFrame.Top = 0
       rctFrame.Left = 0
       '     '// Convert RECT structures to region handles
       hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
       hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
       '     '// Create the new "Transparent" region
       CombineRgn hFrame, hClient, hFrame, RGN_XOR
       '     '// Now lock the window's area to this created region
       SetWindowRgn frm.hwnd, hFrame, True
End Sub

Public Sub Form_Move(theform As Form)
'WILL HELP YOU MOVE A FORM WITHOUT
'A TITLE BAR, PLACE IN MOUSEDOWN
           Call SendMessage(theform.hwnd, &HA1, 2, 0&)
       '
End Sub

Sub shrink(label As label, startSIZE, endSIZE)
'this function makes the text in a lebel shrink.
'the startSIZE has to be greater that the endSize.
'if the endSIZE is 0, the label dispears

label.Visible = True
label.FontSize = startSIZE
Dim X
Do
X = label.FontSize - 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize < (endSIZE Or 2) Then Exit Do
Loop
If endSIZE = 0 Then
    label.Visible = False
End If
End Sub


Sub Grow(label As label, startSIZE, endSIZE)
'this function makes a label grow
'the font size starts a startSIZE,
'and ends and endSIZE
'MADE SURE THE FONT OF THE LABEL
'CAN GO BIG ENOUGH
label.Visible = True
label.FontSize = startSIZE
Dim X
Do While label.FontSize < endSIZE
    label.FontSize = label.FontSize + 2
        timeout (0.001)
    
Loop
End Sub

Sub bounce(label As label, minSIZE, MAXSIZE, numOFbounces)
'this function makes a label, grow, and shrink
'giving it a "BOUNCING" effect
'minSIZE is the smallest it goes, maxSIZE is the largest the label goes
'numOFbounces is the number of times the label bounces
'to make a label bounce forever, call this
'function in a timer, and have numOFbounces = 1
'MADE SURE THE FONT OF THE LABEL
'CAN GO BIG ENOUGH

label.FontSize = minSIZE
Dim X
Dim Y
Dim num
Start:
If (num >= numOFbounces) Then GoTo out:
Do
X = label.FontSize + 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize >= MAXSIZE Then Exit Do
Loop

Do
X = label.FontSize - 2
label.FontSize = X
    timeout (0.001)
    If label.FontSize < (minSIZE Or 2) Then Exit Do
Loop
num = num + 1
GoTo Start:
out:
End Sub

Sub bounce2(label As label, big_size, small_size, numOFbounces)
'this is another bounce function
'Makesure the labels font will support
'the font sizes first
Dim X
For X = 1 To numOFbounces
Call Grow(label, small_size, big_size)
Call shrink(label, big_size, small_size)
Next X
End Sub

Sub FlipPictureHorizontal(pic1 As PictureBox, pic2 As PictureBox)
'pic1 = the existing pic
'pic2 = the pic to be fliped
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Dim px%
    Dim py%
    Dim retval%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retval% = StretchBlt(pic2.hdc, px%, 0, -px%, py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub FlipPictureVertical(pic1 As PictureBox, pic2 As PictureBox)
'pic1 = the existing pic
'pic2 = the pic to be fliped
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Dim px%
    Dim py%
    Dim retval%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retval% = StretchBlt(pic2.hdc, 0, py%, px%, -py%, pic1.hdc, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub PicRotate45(pic1 As PictureBox, pic2 As PictureBox)
'rotate 45 degrees
'pic1 = the existing pic
'pic2 = the pic to be rotated
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Call bmp_rotate(pic1, pic2, 3.14 / 4)
               End Sub


               Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, ByVal theta!)
                ' bmp_rotate(pic1, pic2, theta)
                ' Rotate the image in a picture box.
                '   pic1 is the picture box with the bitmap to rotate
                '   pic2 is the picture box to receive the rotated bitmap
                '   theta is the angle of rotation
                Dim c1x As Integer, c1y As Integer
                Dim c2x As Integer, c2y As Integer
                Dim a As Single
                Dim p1x As Integer, p1y As Integer
                Dim p2x As Integer, p2y As Integer
                Dim n As Integer, R   As Integer

                c1x = pic1.ScaleWidth \ 2
                c1y = pic1.ScaleHeight \ 2
                c2x = pic2.ScaleWidth \ 2
                c2y = pic2.ScaleHeight \ 2

                If c2x < c2y Then n = c2y Else n = c2x
                n = n - 1
                Dim pic1hdc%
                Dim pic2hdc%
                pic1hdc% = pic1.hdc
                pic2hdc% = pic2.hdc

                For p2x = 0 To n
                  For p2y = 0 To n
                    If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
                    R = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
                    p1x = R * Cos(a + theta!)
                    p1y = R * Sin(a + theta!)
                    
                    Dim c0&
                    Dim c1&
                    Dim c2&
                    Dim c3&
                    Dim xret&
                    Dim T%
                    
                    c0& = GetPixel(pic1hdc%, c1x + p1x, c1y + p1y)
                    c1& = GetPixel(pic1hdc%, c1x - p1x, c1y - p1y)
                    c2& = GetPixel(pic1hdc%, c1x + p1y, c1y - p1x)
                    c3& = GetPixel(pic1hdc%, c1x - p1y, c1y + p1x)
                    If c0& <> -1 Then xret& = SetPixel(pic2hdc%, c2x + p2x, c2y + p2y, c0&)
                    If c1& <> -1 Then xret& = SetPixel(pic2hdc%, c2x - p2x, c2y - p2y, c1&)
                    If c2& <> -1 Then xret& = SetPixel(pic2hdc%, c2x + p2y, c2y - p2x, c2&)
                    If c3& <> -1 Then xret& = SetPixel(pic2hdc%, c2x - p2y, c2y + p2x, c3&)
                  Next
                  T% = DoEvents()
                Next
               End Sub
               Sub Fire(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
Y = Y - PiSS.Height / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(255, red, 0), BF
Loop
End Sub
Sub BLUE(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
Y = Y - PiSS.Height / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, 0, red), BF
Loop
End Sub
Sub CircleFire(PiSS As Object)
Dim X
Dim Y
Dim red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
red = 0
TuRd = PiSS.Width
Do Until red = 255
red = red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(255, red, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(255, red, 0)
Loop
End Sub

Sub CircleRed(PiSS As Object)
Dim X
Dim Y
Dim red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
red = 0
TuRd = PiSS.Width
Do Until red = 255
red = red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(red, 0, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(red, 0, 0)
Loop
End Sub

Sub CircleBlue(PiSS As Object)
Dim X
Dim Y
Dim red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
red = 0
TuRd = PiSS.Width
Do Until red = 255
red = red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(0, 0, red)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(0, 0, red)
Loop
End Sub


Sub CircleGreen(PiSS As Object)
Dim X
Dim Y
Dim red
Dim TuRd
X = PiSS.Width
Y = PiSS.Height
PiSS.FillStyle = 0
red = 0
TuRd = PiSS.Width
Do Until red = 255
red = red + 1
TuRd = TuRd - PiSS.Width / 255 * 1
PiSS.FillColor = RGB(0, red, 0)
If TuRd < 0 Then Exit Do
PiSS.Circle (PiSS.Width / 2, PiSS.Height / 2), TuRd, RGB(0, red, 0)
Loop
End Sub


Sub red(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
Y = Y - PiSS.Height / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(red, 0, 0), BF
Loop
End Sub
Sub GREEN(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
Y = Y - PiSS.Height / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, red, 0), BF
Loop
End Sub
Sub SideRed(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
X = X - PiSS.Width / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(red, 0, 0), BF
Loop
End Sub
Sub SideGreen(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
X = X - PiSS.Width / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, red, 0), BF
Loop
End Sub
Sub SideBlue(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
X = X - PiSS.Width / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(0, 0, red), BF
Loop
End Sub

Sub SideFire(PiSS As Object)
Dim X
Dim Y
Dim red
Dim GREEN
Dim BLUE
X = PiSS.Width
Y = PiSS.Height
red = 255
GREEN = 255
BLUE = 255
Do Until red = 0
X = X - PiSS.Width / 255 * 1
red = red - 1
PiSS.Line (0, 0)-(X, Y), RGB(255, red, 0), BF
Loop
End Sub
Sub sweet(a As Form)
a.Enabled = False
a.Left = "-30"
a.Top = "1170"
a.Height = "0"
Pause (0.1)
a.Left = "30"
a.Top = "1250"
a.Height = "100"
Pause (0.1)
a.Left = "130"
a.Top = "1300"
a.Height = "200"
Pause (0.1)
a.Left = "230"
a.Top = "1350"
a.Height = "300"
Pause (0.1)
a.Left = "330"
a.Top = "1400"
a.Height = "400"
Pause (0.1)
a.Left = "430"
a.Top = "1450"
a.Height = "500"
Pause (0.1)
a.Left = "530"
a.Top = "1500"
a.Height = "600"
Pause (0.1)
a.Left = "630"
a.Top = "1550"
a.Height = "700"
Pause (0.1)
a.Left = "830"
a.Top = "1600"
a.Height = "800"
Pause (0.1)
a.Top = "1650"
a.Left = "1030"
a.Height = "900"
Pause (0.1)
a.Top = "1700"
a.Left = "1230"
a.Height = "1000"
Pause (0.1)
a.Top = "1750"
a.Left = "1430"
a.Height = "1100"
Pause (0.1)
a.Top = "1800"
a.Left = "1630"
a.Height = "1200"
Pause (0.1)
a.Top = "1850"
a.Left = "1830"
a.Height = "1300"
Pause (0.1)
a.Top = "1900"
a.Left = "2030"
a.Height = "1400"
Pause (0.1)
a.Top = "1910"
a.Left = "2230"
a.Height = "1500"
Pause (0.1)
a.Top = "1920"
a.Left = "2430"
a.Height = "1600"
Pause (0.1)
a.Top = "1930"
a.Left = "2630"
a.Height = "1700"
Pause (0.1)
a.Top = "1940"
a.Left = "2830"
a.Height = "1800"
Pause (0.1)
a.Top = "1950"
a.Left = "3030"
a.Height = "1900"
Pause (0.1)
a.Left = "3130"
a.Height = "2000"
Pause (0.1)
a.Left = "3230"
a.Height = "2100"
Pause (0.1)
a.Left = "3330"
a.Height = "2200"
Pause (0.1)
a.Height = "2300"
Pause (0.1)
a.Height = "2360"
a.Enabled = True
End Sub
Sub formWidth(m As Form)

m.Width = 5
Pause (0.1)
m.Width = 400
Pause (0.1)
m.Width = 700
Pause (0.1)
m.Width = 1000
Pause (0.1)
m.Width = 2000
Pause (0.1)
m.Width = 3000
Pause (0.1)
m.Width = 4000
Pause (0.1)
m.Width = 5000
Pause (0.1)
m.Width = 4000
Pause (0.1)
m.Width = 3000
Pause (0.1)
m.Width = 2000
Pause (0.1)
m.Width = 1000
Pause (0.1)
m.Width = 700
Pause (0.1)
m.Width = 400
Pause (0.1)
m.Width = 5
Pause (0.1)
m.Width = 400
Pause (0.1)
m.Width = 700
Pause (0.1)
m.Width = 1000
Pause (0.1)
m.Width = 2190

End Sub
Sub Formhw(m As Form)

m.Height = 5
m.Width = 5
Pause (0.1)
m.Height = 400
m.Width = 400
Pause (0.1)
m.Height = 700
m.Width = 700
Pause (0.1)
m.Height = 1000
m.Width = 1000
Pause (0.1)
m.Height = 2000
m.Width = 2000
Pause (0.1)
m.Height = 3000
m.Width = 3000
Pause (0.1)
m.Height = 4000
m.Width = 4000
Pause (0.1)
m.Height = 5000
m.Width = 5000
Pause (0.1)
m.Height = 4000
m.Width = 4000
Pause (0.1)
m.Height = 3000
m.Width = 3000
Pause (0.1)
m.Height = 2000
m.Width = 2000
Pause (0.1)
m.Height = 1000
m.Width = 1000
Pause (0.1)
m.Height = 700
m.Width = 700
Pause (0.1)
m.Height = 400
m.Width = 400
Pause (0.1)
m.Height = 5
m.Width = 5
Pause (0.1)
m.Height = 400
Pause (0.1)
m.Height = 700
Pause (0.1)
m.Height = 1000
Pause (0.1)
m.Height = 2000

End Sub
Sub formHeight(m As Form)

m.Height = 5
Pause (0.1)
m.Height = 400
Pause (0.1)
m.Height = 700
Pause (0.1)
m.Height = 1000
Pause (0.1)
m.Height = 2000
Pause (0.1)
m.Height = 3000
Pause (0.1)
m.Height = 4000
Pause (0.1)
m.Height = 5000
Pause (0.1)
m.Height = 4000
Pause (0.1)
m.Height = 3000
Pause (0.1)
m.Height = 2000
Pause (0.1)
m.Height = 1000
Pause (0.1)
m.Height = 700
Pause (0.1)
m.Height = 400
Pause (0.1)
m.Height = 5
Pause (0.1)
m.Height = 400
Pause (0.1)
m.Height = 700
Pause (0.1)
m.Height = 1000
Pause (0.1)
m.Height = 2000

End SubSub FormDance(m As Form)

'  This makes a form dance across the screen
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 5000
Pause (0.1)
m.Left = 4000
Pause (0.1)
m.Left = 3000
Pause (0.1)
m.Left = 2000
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 5
Pause (0.1)
m.Left = 400
Pause (0.1)
m.Left = 700
Pause (0.1)
m.Left = 1000
Pause (0.1)
m.Left = 2000

End Sub
Sub coolform(e As Form)
e.Visible = True
e.Enabled = True
e.Left = "-30"
e.Top = "1170"
e.Height = "0"
Pause (0.1)
e.Left = "30"
e.Top = "1250"
e.Height = "100"
Pause (0.1)
e.Left = "130"
e.Top = "1300"
e.Height = "150"
Pause (0.1)
e.Left = "200"
e.Top = "1350"
e.Height = "350"
Pause (0.1)
e.Left = "400"
e.Top = "1400"
e.Height = "400"
Pause (0.1)
e.Left = "430"
e.Top = "1450"
e.Height = "450"
Pause (0.1)
e.Left = "530"
e.Top = "1500"
e.Height = "500"
Pause (0.1)
e.Left = "630"
e.Top = "1550"
e.Height = "550"
Pause (0.1)
e.Left = "830"
e.Top = "1600"
e.Height = "600"
Pause (0.1)
e.Top = "1650"
e.Left = "1030"
e.Height = "650"
Pause (0.1)
e.Top = "1700"
e.Left = "1230"
e.Height = "700"
Pause (0.1)
e.Top = "1750"
e.Left = "1430"
e.Height = "750"
Pause (0.1)
e.Top = "1800"
e.Left = "1630"
e.Height = "800"
Pause (0.1)
e.Top = "1850"
e.Left = "1830"
e.Height = "850"
Pause (0.1)
e.Top = "1900"
e.Left = "2030"
e.Height = "900"
Pause (0.1)
e.Top = "1910"
e.Left = "2230"
e.Height = "950"
Pause (0.1)
e.Top = "1920"
e.Left = "2430"
e.Height = "1000"
Pause (0.1)
e.Top = "1930"
e.Left = "2630"
e.Height = "1020"
Pause (0.1)
e.Top = "1940"
e.Left = "2830"
e.Height = "1040"
Pause (0.1)
e.Top = "1950"
e.Left = "3030"
e.Height = "1060"
Pause (0.1)
e.Height = "1080"
Pause (0.1)
e.Enabled = True
End Sub
Sub CenterForm_Top(f As Form)
f.Left = 3800
f.Top = 0
End Sub
Sub CenterForm(f As Form)
f.Left = 3800
f.Top = 3000
End Sub
Sub bigtext(l As label)
l.FontSize = 33
Call Pause(0.1)
l.FontSize = 31
Call Pause(0.1)
l.FontSize = 29
Call Pause(0.1)
l.FontSize = 27
Call Pause(0.1)
l.FontSize = 25
Call Pause(0.1)
l.FontSize = 23
Call Pause(0.1)
l.FontSize = 21
Call Pause(0.1)
l.FontSize = 19
Call Pause(0.1)
l.FontSize = 17
Call Pause(0.1)
l.FontSize = 15
Call Pause(0.1)
l.FontSize = 13
Call Pause(0.1)
l.FontSize = 11
Call Pause(0.1)
l.FontSize = 9
Call Pause(0.1)
l.FontSize = 7
Call Pause(0.1)
l.FontSize = 5
Call Pause(0.1)
l.FontSize = 3
Call Pause(0.1)
l.FontSize = 1
Call Pause(0.2)
l.FontSize = 33
Call Pause(0.1)
l.FontSize = 31
Call Pause(0.1)
l.FontSize = 29
Call Pause(0.1)
l.FontSize = 27
Call Pause(0.1)
l.FontSize = 25
Call Pause(0.1)
l.FontSize = 23
Call Pause(0.1)
l.FontSize = 21
Call Pause(0.1)
l.FontSize = 19
Call Pause(0.1)
l.FontSize = 17
Call Pause(0.1)
l.FontSize = 15
Call Pause(0.1)
l.FontSize = 13
Call Pause(0.1)
l.FontSize = 11
Call Pause(0.1)
l.FontSize = 9
Call Pause(0.1)
l.FontSize = 7
Call Pause(0.1)
l.FontSize = 5
Call Pause(0.1)
l.FontSize = 3
Call Pause(0.1)
l.FontSize = 1
End Sub

Sub Circle_Suckin()
'Put This Whole code into an object
'so then when u click that object
'it will make the form a circle and suck it in
'then it will be a half form top
SetWindowRgn hwnd, _
CreateEllipticRgn(0, 0, 300, 200), True
Call Form_Suckin(Me)
End Sub
Sub Cartmen_Cap()
'Put This Whole code into an object
'so then when u click that object
'it will make the form look like cartmens hat
'SORTA
SetWindowRgn hwnd, _
CreateEllipticRgn(0, 0, 300, 200), True
Me.Height = 400


End Sub


Sub UltraKilla(kill As String)
'-==Sub made by, Anders, JM!!.

' This is a good macrokill code!
' And Yo put this in a command button or
' it will not work!!!
'I.E. Call UltraKilla("@")
'dont put alot of @@@@@ cause it dont matter

s = String(1699, Asc(kill))
SendChat ("<font color=#00ff00> <P=" & s)
End Sub

Sub HileriousForm()
'-= Sub made by, ToXiD

'READ COMMENTS
'Follow the instructions! And it will work!
'And oh yea! This is funny
'Just Dont temper with the timer intervals!
'All you need is 2 timers!
'In timer 1 set interval to 1 and put this code:
Height = Height + 50
If Height > 5000 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
'In timer 2! Make it enabled = false and
'And interval = 1! Put this code in!!!
Height = Height - 50
If Height < 2000 Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
'And your done!!
End Sub
Function BgChange()
'-= Sub made by, ToXiD
'This will change the background color! from ur
'hand writing! Just read my instructions!:

Rem This one you have to copy the code! Not the
Rem abbrevieation! just were it says Text# put a
Rem number instead of #

If Text#.Text = "red" Then BackColor = vbRed
If Text#.Text = "blue" Then BackColor = vbBlue
If Text#.Text = "yellow" Then BackColor = vbYellow
If Text#.Text = "green" Then BackColor = vbGreen
If Text#.Text = "cyan" Then BackColor = vbCyan
If Text#.Text = "magenta" Then BackColor = vbMagenta
End Function

Private Sub ZzZzZEnd()
'    `·.·¨'¹l|[]|l¹¨'·.·´
'
' ***** *****   *   ****  ****
' * * * *      * *  *     *
' ***** *****  ***  *     ****
' *     *      * *  *     *
' *     *****  * *  ****  ****
'
'    `·.·¨'¹l|[]|l¹¨'·.·´
End Sub
