Attribute VB_Name = "Module1"

Dim WinVersion As Integer, SoundAvailable As Integer
Global Const Wm_Close = &H10
Global Cancel As Integer
Global VisibleFrame As Frame
Global GH
Global NotOnline
Global Const WM_ProGGer = 13 ' &HD%
Global Const RES_INFO = 2
Global BeginTheMassMail
Global StopBustin
Global Mailstosend
Global STOPP
Global Number
Type POINTAPI
    X As Integer
    Y As Integer
End Type

Declare Sub ReleaseCapture Lib "User" ()
Declare Sub closewindow Lib "User" ()
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)

Declare Sub extsubD08 Lib "User" Alias "setwindowpos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)

Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal X As Integer, ByVal Y As Integer)




Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer



Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)


Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)

'Declare Function sendmessage& Lib "User" (ByVal hWnd%, ByVal wmsg%, ByVal wparam%, ByVal lParam As Any)





Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)


Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)


Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function setwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()

Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function getcurrenttime& Lib "User" ()
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetCursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer


Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)

Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wFlags%)



Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)




Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer

Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer


'Subs and Functions for "GDI"
Declare Sub setbkcolor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)

Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long

Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer

'Subs and Functions for "MMSystem"


'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "vbwfind.dll" (ByVal hWnd%, ByVal Title$)



'Subs and Functions for "APIGuide.Dll"
Declare Sub agCopyData Lib "Apiguide.dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "Apiguide.dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "Apiguide.dll" (ByVal L&, lw%, lh%)
Declare Sub agOutp Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "Apiguide.dll" (ByVal portid%, ByVal outval%)
Declare Function agGetControlHwnd% Lib "Apiguide.dll" (hctl As Control)
Declare Function agGetInstance% Lib "Apiguide.dll" ()
Declare Function agGetAddressForObject& Lib "Apiguide.dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "Apiguide.dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "Apiguide.dll" Alias "agGetAddressForObject" (ByVal lpString$)
Declare Function agGetAddressForVBString& Lib "Apiguide.dll" (vbstring$)
Declare Function agGetControlName$ Lib "Apiguide.dll" (ByVal hWnd%)
Declare Function agXPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "Apiguide.dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
Declare Function agYTwipsToPixels% Lib "Apiguide.dll" (ByVal twips&)
Declare Function agDeviceCapabilities& Lib "Apiguide.dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "Apiguide.dll" (ByVal hWnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "Apiguide.dll" (ByVal hWnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "Apiguide.dll" (ByVal portid%)
Declare Function agInpw% Lib "Apiguide.dll" (ByVal portid%)
Declare Function agHugeOffset& Lib "Apiguide.dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "Apiguide.dll" ()
Declare Function agVBSendControlMsg& Lib "Apiguide.dll" (ctl As Control, ByVal Msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "Apiguide.dll" (ctl As Control, ByVal mask&, ByVal value&)
Declare Function dwVBSetControlFlags& Lib "Apiguide.dll" (ctl As Control, ByVal mask&, ByVal value&)
Declare Function EnableWindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function MessageBox Lib "User" (ByVal hWnd As Integer, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Integer) As Integer

Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hwndparent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Declare Sub HideCaret Lib "User" (ByVal hWnd As Integer)
Declare Function SendMessageLong Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal hMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, lpKeyName As Any, lpString As Any, ByVal lplFileName As String) As Integer
Declare Function dwGetStringFromLPSTR Lib "Dwspydll.dll" (ByVal p1&) As String
Declare Function IsBadReadPtr Lib "Kernel" (ByVal lp As String, ByVal cb As Integer) As Integer
Declare Function IsBadReadPro Lib "Kernel" Alias "IsBadReadPtr" (ByVal p1&, ByVal p2%) As Integer
Declare Function WindowFromPoint Lib "User" (ByVal ptScreen As Any) As Integer

Declare Function GetWindowLong Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer

Declare Sub GetCursorPos Lib "User" (lpPoint As POINTAPI)
Declare Function EnumChildWindows Lib "User" () As Integer
Declare Sub GetWindowRect Lib "User" (ByVal p1%, p2 As Any)
Declare Function FlashWindow% Lib "User" (ByVal hWnd%, ByVal bInvert%)
Declare Function mciGetErrorString Lib "MMSYSTEM.DLL" (ByVal p1&, ByVal p2$, ByVal p3%) As Integer
Declare Function MciSendString Lib "MMSYSTEM.DLL" (ByVal p1$, ByVal p2$, ByVal p3%, ByVal p4%) As Long
Declare Function VBmciSendString$ Lib "vbstr.dll" (ByVal lpCmd$)
Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer
Declare Function ExitWindows Lib "User" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer

Declare Function GetModuleUsage Lib "Kernel" (ByVal p1%) As Integer
Declare Function SetParent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function SetMenu Lib "User" (ByVal hWnd As Integer, ByVal hMenu As Integer) As Integer
Declare Function DestroyMenu Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As String) As Integer
Declare Function CreateMenu Lib "User" () As Integer
Declare Function RemoveMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function SendDlgItemMessage Lib "User" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function VBENumChildWindows Lib "vbstr.dll" (ByVal p1%, ByVal p2%) As Integer
Declare Function getfocus Lib "User" () As Integer
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function gettopwindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function PostMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Integer
Declare Function SetFocusB Lib "User" (ByVal p1%) As Integer
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetNextWindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function AOLGetList% Lib "311.dll" (ByVal index%, ByVal buf$)
Declare Function AOLGetCombo% Lib "311.dll" (ByVal index%, ByVal buf$)
Declare Function SendMessageByNum Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Long) As Integer
Declare Function SendMessageByString Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
Declare Function setactivewindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function FindWindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function FindChildByTitle% Lib "vbwfind.dll" (ByVal hWnd%, ByVal Title$)
Declare Function FindChildByClass% Lib "vbwfind.dll" (ByVal hWnd%, ByVal Title$)
Declare Function aggetstringfromlpstr$ Lib "Apiguide.dll" (ByVal lpString&)

Declare Function getwindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetParent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetTickCount Lib "User" () As Long
Declare Function SetFocusAPI Lib "User" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
Declare Function GetWindowText Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetWinFlags Lib "Kernel" () As Long
Global Const WF_CPU286 = &H2&
Global Const WF_CPU386 = &H4&
Global Const WF_CPU486 = &H8&
Global Const WF_STANDARD = &H10&
Global Const WF_ENHANCED = &H20&
Global Const WF_80x87 = &H400&
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetSystemMetrics Lib "User" (ByVal n As Integer) As Integer
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function GlobalCompact Lib "Kernel" (ByVal flag&) As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal flag%) As Long


Global Const GFSR_SYSTEMRESOURCES = &H0
Global Const GFSR_GDIRESOURCES = &H1
Global Const GFSR_USERRESOURCES = &H2


Declare Function waveOutGetNumDevs Lib "MMSystem" () As Integer
Declare Function GetMenu Lib "User" (ByVal hWd%) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu%, ByVal nPos%) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer

Global Const MF_POPUP = &H10
Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800

Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function getdc Lib "User" (ByVal hWnd%) As Integer
Declare Sub ReleaseDC Lib "User" (ByVal hWnd%, ByVal hdc%)
Declare Function BitBlt Lib "GDI" (ByVal destDC%, ByVal X%, ByVal Y%, ByVal w%, ByVal h%, ByVal srchDC%, ByVal srcX%, ByVal srcY%, ByVal rop&) As Integer
Declare Sub SetWindowPos Lib "User" (ByVal h1%, ByVal h2%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%)

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

Type WINDOWPOS
    hWnd As Integer
    hwndInsertAfter As Integer
    X As Integer
    Y As Integer
    cx As Integer
    cy As Integer
    Flags As Integer
End Type
Type EVENTMSG
    message As Integer
    paramL As Integer
    paramH As Integer
    time As Long
End Type

''  OpenFile() Flags
Global Const OF_READ = &H0
Global Const OF_WRITE = &H1
Global Const OF_READWRITE = &H2
Global Const OF_SHARE_COMPAT = &H0
Global Const OF_SHARE_EXCLUSIVE = &H10
Global Const OF_SHARE_DENY_WRITE = &H20
Global Const OF_SHARE_DENY_READ = &H30
Global Const OF_SHARE_DENY_NONE = &H40
Global Const OF_PARSE = &H100
Global Const OF_DELETE = &H200
Global Const OF_VERIFY = &H400
Global Const OF_SEARCH = &H400
Global Const OF_CANCEL = &H800
Global Const OF_CREATE = &H1000
Global Const OF_PROMPT = &H2000
Global Const OF_EXIST = &H4000
Global Const OF_REOPEN = &H8000
Global Const TF_FORCEDRIVE = &H80

''  GetDriveType return values
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4

''  Global Memory Flags
Global Const GMEM_FIXED = &H0
Global Const GMEM_MOVEABLE = &H2
Global Const GMEM_NOCOMPACT = &H10
Global Const GMEM_NODISCARD = &H20
Global Const GMEM_ZEROINIT = &H40
Global Const GMEM_MODIFY = &H80
Global Const GMEM_DISCARDABLE = &H100
Global Const GMEM_NOT_BANKED = &H1000
Global Const GMEM_SHARE = &H2000
Global Const GMEM_DDESHARE = &H2000
Global Const GMEM_NOTIFY = &H4000
Global Const GMEM_LOWER = GMEM_NOT_BANKED
Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

'  Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF

''  Predefined Resource Types
Global Const RT_CURSOR = 1&
Global Const RT_BITMAP = 2&
Global Const RT_ICON = 3&
Global Const RT_MENU = 4&
Global Const RT_DIALOG = 5&
Global Const RT_STRING = 6&
Global Const RT_FONTDIR = 7&
Global Const RT_FONT = 8&
Global Const RT_ACCELERATOR = 9&
Global Const RT_RCDATA = 10&

'' GetWinFlags constants
Global Const WF_PMODE = &H1
Global Const WF_WIN286 = &H10
Global Const WF_WIN386 = &H20
Global Const WF_CPU086 = &H40
Global Const WF_CPU186 = &H80
Global Const WF_LARGEFRAME = &H100
Global Const WF_SMALLFRAME = &H200


'' Parameter error checking constants
Global Const ERR_WARNING = 8
Global Const ERR_PARAM = 4
Global Const ERR_SIZE_MASK = 3
Global Const ERR_BYTE = 1
Global Const ERR_WORD = 2
Global Const ERR_DWORD = 3
Global Const ERR_BAD_VALUE = &H6001
Global Const ERR_BAD_FLAGS = &H6002
Global Const ERR_BAD_INDEX = &H6003
Global Const ERR_BAD_DVALUE = &H7004
Global Const ERR_BAD_DFLAGS = &H7005
Global Const ERR_BAD_DINDEX = &H7006
Global Const ERR_BAD_PTR = &H7007
Global Const ERR_BAD_FUNC_PTR = &H7008
Global Const ERR_BAD_SELECTOR = &H6009
Global Const ERR_BAD_STRING_PTR = &H700A
Global Const ERR_BAD_HANDLE = &H600B

' KERNEL parameter errors
Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025

' USER parameter errors
Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046

' GDI parameter errors
Global Const ERR_BAD_COORDS = &H7060
Global Const ERR_BAD_GDI_OBJECT = &H6061
Global Const ERR_BAD_HDC = &H6062
Global Const ERR_BAD_HPEN = &H6063
Global Const ERR_BAD_HFONT = &H6064
Global Const ERR_BAD_HBRUSH = &H6065
Global Const ERR_BAD_HBITMAP = &H6066
Global Const ERR_BAD_HRGN = &H6067
Global Const ERR_BAD_HPALETTE = &H6068
Global Const ERR_BAD_HMETAFILE = &H6069


'/**** LogError() values */

' KERNEL errors
Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9

' USER errors
Global Const ERR_CREATEDLG = &H40
Global Const ERR_CREATEDLG2 = &H41
Global Const ERR_REGISTERCLASS = &H42
Global Const ERR_DCBUSY = &H43
Global Const ERR_CREATEWND = &H44
Global Const ERR_STRUCEXTRA = &H45
Global Const ERR_LOADSTR = &H46
Global Const ERR_LOADMENU = &H47
Global Const ERR_NESTEDBEGINPAINT = &H48
Global Const ERR_BADINDEX = &H49
Global Const ERR_CREATEMENU = &H4A

' GDI errors
Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83

'' Exit Windows parameters

Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_REBOOTSYSTEM = &H43


''  Stock system bitmaps
Global Const OBM_CLOSE = 32754
Global Const OBM_UPARROW = 32753
Global Const OBM_DNARROW = 32752
Global Const OBM_RGARROW = 32751
Global Const OBM_LFARROW = 32750
Global Const OBM_REDUCE = 32749
Global Const OBM_ZOOM = 32748
Global Const OBM_RESTORE = 32747
Global Const OBM_REDUCED = 32746
Global Const OBM_ZOOMD = 32745
Global Const OBM_RESTORED = 32744
Global Const OBM_UPARROWD = 32743
Global Const OBM_DNARROWD = 32742
Global Const OBM_RGARROWD = 32741
Global Const OBM_LFARROWD = 32740
Global Const OBM_MNARROW = 32739
Global Const OBM_COMBO = 32738
Global Const OBM_UPARROWI = 32737
Global Const OBM_DNARROWI = 32736
Global Const OBM_RGARROWI = 32735
Global Const OBM_LFARROWI = 32734
Global Const OBM_OLD_CLOSE = 32767
Global Const OBM_SIZE = 32766
Global Const OBM_OLD_UPARROW = 32765
Global Const OBM_OLD_DNARROW = 32764
Global Const OBM_OLD_RGARROW = 32763
Global Const OBM_OLD_LFARROW = 32762
Global Const OBM_BTSIZE = 32761
Global Const OBM_CHECK = 32760
Global Const OBM_CHECKBOXES = 32759
Global Const OBM_BTNCORNERS = 32758
Global Const OBM_OLD_REDUCE = 32757
Global Const OBM_OLD_ZOOM = 32756
Global Const OBM_OLD_RESTORE = 32755

'' Stock system Icons
Global Const OCR_NORMAL = 32512
Global Const OCR_IBEAM = 32513
Global Const OCR_WAIT = 32514
Global Const OCR_CROSS = 32515
Global Const OCR_UP = 32516
Global Const OCR_SIZE = 32640
Global Const OCR_ICON = 32641
Global Const OCR_SIZENWSE = 32642
Global Const OCR_SIZENESW = 32643
Global Const OCR_SIZEWE = 32644
Global Const OCR_SIZENS = 32645
Global Const OCR_SIZEALL = 32646
Global Const OCR_ICOCUR = 32647

Global Const OIC_SAMPLE = 32512
Global Const OIC_HAND = 32513
Global Const OIC_QUES = 32514
Global Const OIC_BANG = 32515
Global Const OIC_NOTE = 32516


' --------------------------------------------------------------------------
'   GDI Section
' --------------------------------------------------------------------------

''  Raster-ops (Binary)
Global Const R2_BLACK = 1       '   0
Global Const R2_NOTMERGEPEN = 2 '  DPon
Global Const R2_MASKNOTPEN = 3  '  DPna
Global Const R2_NOTCOPYPEN = 4  '  PN
Global Const R2_MASKPENNOT = 5  '  PDna
Global Const R2_NOT = 6 '  Dn
Global Const R2_XORPEN = 7      '  DPx
Global Const R2_NOTMASKPEN = 8  '  DPan
Global Const R2_MASKPEN = 9     '  DPa
Global Const R2_NOTXORPEN = 10  '  DPxn
Global Const R2_NOP = 11        '  D
Global Const R2_MERGENOTPEN = 12        '  DPno
Global Const R2_COPYPEN = 13    '  P
Global Const R2_MERGEPENNOT = 14        '  PDno
Global Const R2_MERGEPEN = 15   '  DPo
Global Const R2_WHITE = 16      '   1

'' Raster-ops (Ternary)










''  StretchBlt() Modes




''  PolyFill() Modes

''  Text Alignment Options
Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

'' ExtTextOut flags
Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4

'' SetMapperFlags constants
Global Const ASPECT_FILTERING = &H1

''  Metafile Functions
Global Const META_SETBKCOLOR = &H201
Global Const META_SETBKMODE = &H102
Global Const META_SETMAPMODE = &H103
Global Const META_SETROP2 = &H104
Global Const META_SETRELABS = &H105
Global Const META_SETPOLYFILLMODE = &H106
Global Const META_SETSTRETCHBLTMODE = &H107
Global Const META_SETTEXTCHAREXTRA = &H108
Global Const META_SETTEXTCOLOR = &H209
Global Const META_SETTEXTJUSTIFICATION = &H20A
Global Const META_SETWINDOWORG = &H20B
Global Const META_SETWINDOWEXT = &H20C
Global Const META_SETVIEWPORTORG = &H20D
Global Const META_SETVIEWPORTEXT = &H20E
Global Const META_OFFSETWINDOWORG = &H20F
Global Const META_SCALEWINDOWEXT = &H400
Global Const META_OFFSETVIEWPORTORG = &H211
Global Const META_SCALEVIEWPORTEXT = &H412
Global Const META_LINETO = &H213
Global Const META_MOVETO = &H214
Global Const META_EXCLUDECLIPRECT = &H415
Global Const META_INTERSECTCLIPRECT = &H416
Global Const META_ARC = &H817
Global Const META_ELLIPSE = &H418
Global Const META_FLOODFILL = &H419
Global Const META_PIE = &H81A
Global Const META_RECTANGLE = &H41B
Global Const META_ROUNDRECT = &H61C
Global Const META_PATBLT = &H61D
Global Const META_SAVEDC = &H1E
Global Const META_SETPIXEL = &H41F
Global Const META_OFFSETCLIPRGN = &H220
Global Const META_TEXTOUT = &H521
Global Const META_BITBLT = &H922
Global Const META_STRETCHBLT = &HB23
Global Const META_POLYGON = &H324
Global Const META_POLYLINE = &H325
Global Const META_ESCAPE = &H626
Global Const META_RESTOREDC = &H127
Global Const META_FILLREGION = &H228
Global Const META_FRAMEREGION = &H429
Global Const META_INVERTREGION = &H12A
Global Const META_PAINTREGION = &H12B
Global Const META_SELECTCLIPREGION = &H12C
Global Const META_SELECTOBJECT = &H12D
Global Const META_SETTEXTALIGN = &H12E
Global Const META_DRAWTEXT = &H62F
Global Const META_CHORD = &H830
Global Const META_SETMAPPERFLAGS = &H231
Global Const META_EXTTEXTOUT = &HA32
Global Const META_SETDIBTODEV = &HD33
Global Const META_SELECTPALETTE = &H234
Global Const META_REALIZEPALETTE = &H35
Global Const META_ANIMATEPALETTE = &H436
Global Const META_SETPALENTRIES = &H37
Global Const META_POLYPOLYGON = &H538
Global Const META_RESIZEPALETTE = &H139
Global Const META_DIBBITBLT = &H940
Global Const META_DIBSTRETCHBLT = &HB41
Global Const META_DIBCREATEPATTERNBRUSH = &H142
Global Const META_STRETCHDIB = &HF43
Global Const META_DELETEOBJECT = &H1F0
Global Const META_CREATEPALETTE = &HF7
Global Const META_CREATEBRUSH = &HF8
Global Const META_CREATEPATTERNBRUSH = &H1F9
Global Const META_CREATEPENINDIRECT = &H2FA
Global Const META_CREATEFONTINDIRECT = &H2FB
Global Const META_CREATEBRUSHINDIRECT = &H2FC
Global Const META_CREATEBITMAPINDIRECT = &H2FD
Global Const META_CREATEBITMAP = &H6FE
Global Const META_CREATEREGION = &H6FF


'' Escape Constants
Global Const NEWFRAME = 1
Global Const ABORTDOCCONST = 2
Global Const NEXTBAND = 3
Global Const SETCOLORTABLE = 4
Global Const GETCOLORTABLE = 5
Global Const FLUSHOUTPUT = 6
Global Const DRAFTMODE = 7
Global Const QUERYESCSUPPORT = 8
Global Const SETABORTPROCCONST = 9
Global Const STARTDOCCONST = 10
Global Const ENDDOCAPICONST = 11
Global Const GETPHYSPAGESIZE = 12
Global Const GETPRINTINGOFFSET = 13
Global Const GETSCALINGFACTOR = 14
Global Const MFCOMMENT = 15
Global Const GETPENWIDTH = 16
Global Const SETCOPYCOUNT = 17
Global Const SELECTPAPERSOURCE = 18
Global Const DEVICEDATA = 19
Global Const PASSTHROUGH = 19
Global Const GETTECHNOLGY = 20
Global Const GETTECHNOLOGY = 20
Global Const SETENDCAP = 21
Global Const SETLINEJOIN = 22
Global Const SETMITERLIMIT = 23
Global Const BANDINFO = 24
Global Const DRAWPATTERNRECT = 25
Global Const GETVECTORPENSIZE = 26
Global Const GETVECTORBRUSHSIZE = 27
Global Const ENABLEDUPLEX = 28
Global Const GETSETPAPERBINS = 29
Global Const GETSETPRINTORIENT = 30
Global Const ENUMPAPERBINS = 31
Global Const SETDIBSCALING = 32
Global Const EPSPRINTING = 33
Global Const ENUMPAPERMETRICS = 34
Global Const GETSETPAPERMETRICS = 35
Global Const POSTSCRIPT_DATA = 37
Global Const POSTSCRIPT_IGNORE = 38
Global Const GETEXTENDEDTEXTMETRICS = 256
Global Const GETEXTENTTABLE = 257
Global Const GETPAIRKERNTABLE = 258
Global Const GETTRACKKERNTABLE = 259
Global Const EXTTEXTOUTCONST = 512
Global Const ENABLERELATIVEWIDTHS = 768
Global Const ENABLEPAIRKERNING = 769
Global Const SETKERNTRACK = 770
Global Const SETALLJUSTVALUES = 771
Global Const SETCHARSET = 772
Global Const STRETCHBLTCONST = 2048
Global Const BEGIN_PATH = 4096
Global Const CLIP_TO_PATH = 4097
Global Const END_PATH = 4098
Global Const EXT_DEVICE_CAPS = 4099
Global Const RESTORE_CTM = 4100
Global Const SAVE_CTM = 4101
Global Const SET_ARC_DIRECTION = 4102
Global Const SET_BACKGROUND_COLOR = 4103
Global Const SET_POLY_MODE = 4104
Global Const SET_SCREEN_ANGLE = 4105
Global Const SET_SPREAD = 4106
Global Const TRANSFORM_CTM = 4107
Global Const SET_CLIP_BOX = 4108
Global Const SET_BOUNDS = 4109
Global Const SET_MIRROR_MODE = 4110

''  Spooler Error Codes
Global Const SP_NOTREPORTED = &H4000
Global Const SP_ERROR = (-1)
Global Const SP_APPABORT = (-2)
Global Const SP_USERABORT = (-3)
Global Const SP_OUTOFDISK = (-4)
Global Const SP_OUTOFMEMORY = (-5)
Global Const PR_JOBSTATUS = &H0


'' biCompression field constants for DIB
Global Const BI_RGB = 0&
Global Const BI_RLE8 = 1&
Global Const BI_RLE4 = 2&

'' LOGFONT and TEXTMETRIC constants
Global Const OUT_DEFAULT_PRECIS = 0
Global Const OUT_STRING_PRECIS = 1
Global Const OUT_CHARACTER_PRECIS = 2
Global Const OUT_STROKE_PRECIS = 3
Global Const OUT_TT_PRECIS = 4
Global Const OUT_DEVICE_PRECIS = 5
Global Const OUT_RASTER_PRECIS = 6
Global Const OUT_TT_ONLY_PRECIS = 7



Global Const CLIP_DEFAULT_PRECIS = 0
Global Const CLIP_CHARACTER_PRECIS = 1
Global Const CLIP_STROKE_PRECIS = 2
Global Const CLIP_LH_ANGLES = &H10
Global Const CLIP_TT_ALWAYS = &H20
Global Const CLIP_EMBEDDED = &H80


Global Const DEFAULT_QUALITY = 0
Global Const DRAFT_QUALITY = 1
Global Const PROOF_QUALITY = 2
Global Const DEFAULT_PITCH = 0
Global Const FIXED_PITCH = 1
Global Const VARIABLE_PITCH = 2
Global Const TMPF_FIXED_PITCH = 1
Global Const TMPF_VECTOR = 2
Global Const TMPF_DEVICE = 8
Global Const TMPF_TRUETYPE = 4
Global Const ANSI_CHARSET = 0
Global Const DEFAULT_CHARSET = 1
Global Const SYMBOL_CHARSET = 2
Global Const SHIFTJIS_CHARSET = 128
Global Const OEM_CHARSET = 255
Global Const NTM_REGULAR = &H40&
Global Const NTM_BOLD = &H20&
Global Const NTM_ITALIC = &H1&
Global Const LF_FULLFACESIZE = 64
Global Const RASTER_FONTTYPE = 1
Global Const DEVICE_FONTTYPE = 2
Global Const TRUETYPE_FONTTYPE = 4


'  Font Families
'
Global Const FF_DONTCARE = 0    '  Don't care or don't know.
Global Const FF_ROMAN = 16      '  Variable stroke width, serifed.

'  Times Roman, Century Schoolbook, etc.
Global Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

'  Helvetica, Swiss, etc.
Global Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

'  Pica, Elite, Courier, etc.
Global Const FF_SCRIPT = 64     '  Cursive, etc.
Global Const FF_DECORATIVE = 80 '  Old English, etc.

''  Font Weights
Global Const FW_DONTCARE = 0
Global Const FW_THIN = 100
Global Const FW_EXTRALIGHT = 200
Global Const FW_LIGHT = 300
Global Const FW_NORMAL = 400
Global Const FW_MEDIUM = 500
Global Const FW_SEMIBOLD = 600
Global Const FW_BOLD = 700
Global Const FW_EXTRABOLD = 800
Global Const FW_HEAVY = 900
Global Const FW_ULTRALIGHT = FW_EXTRALIGHT
Global Const FW_REGULAR = FW_NORMAL
Global Const FW_DEMIBOLD = FW_SEMIBOLD
Global Const FW_ULTRABOLD = FW_EXTRABOLD
Global Const FW_BLACK = FW_HEAVY

''  Background Modes

Global Const OPAQUE = 2

''  Mapping Modes
Global Const MM_TEXT = 1
Global Const MM_LOMETRIC = 2
Global Const MM_HIMETRIC = 3
Global Const MM_LOENGLISH = 4
Global Const MM_HIENGLISH = 5
Global Const MM_TWIPS = 6
Global Const MM_ISOTROPIC = 7
Global Const MM_ANISOTROPIC = 8

''  Coordinate Modes
Global Const ABSOLUTE = 1
Global Const RELATIVE = 2

''  Stock Logical Objects
Global Const WHITE_BRUSH = 0
Global Const LTGRAY_BRUSH = 1
Global Const GRAY_BRUSH = 2
Global Const DKGRAY_BRUSH = 3
Global Const BLACK_BRUSH = 4
Global Const NULL_BRUSH = 5
Global Const HOLLOW_BRUSH = NULL_BRUSH
Global Const WHITE_PEN = 6
Global Const BLACK_PEN = 7
Global Const NULL_PEN = 8
Global Const OEM_FIXED_FONT = 10
Global Const ANSI_FIXED_FONT = 11
Global Const ANSI_VAR_FONT = 12
Global Const SYSTEM_FONT = 13
Global Const DEVICE_DEFAULT_FONT = 14
Global Const DEFAULT_PALETTE = 15
Global Const SYSTEM_FIXED_FONT = 16

''  Brush Styles
Global Const BS_SOLID = 0
Global Const BS_NULL = 1
Global Const BS_HOLLOW = BS_NULL
Global Const BS_HATCHED = 2
Global Const BS_PATTERN = 3
Global Const BS_INDEXED = 4
Global Const BS_DIBPATTERN = 5

''  Hatch Styles
Global Const HS_HORIZONTAL = 0  '  -----
Global Const HS_VERTICAL = 1    '  |||||
Global Const HS_FDIAGONAL = 2   '  \\\\\
Global Const HS_BDIAGONAL = 3   '  /////
Global Const HS_CROSS = 4       '  +++++
Global Const HS_DIAGCROSS = 5   '  xxxxx

''  Pen Styles
Global Const PS_SOLID = 0
Global Const PS_DASH = 1        '  -------
Global Const PS_DOT = 2 '  .......
Global Const PS_DASHDOT = 3     '  _._._._
Global Const PS_DASHDOTDOT = 4  '  _.._.._
Global Const PS_NULL = 5
Global Const PS_INSIDEFRAME = 6

'' Bounds Rectangle Constants
Global Const DCB_RESET = 1
Global Const DCB_ACCUMULATE = 2
Global Const DCB_DIRTY = 2
Global Const DCB_SET = 3
Global Const DCB_ENABLE = 4
Global Const DCB_DISABLE = 8

''  GetDeviceCaps() Device Parameters
Global Const DRIVERVERSION = 0  '  Device driver version
Global Const TECHNOLOGY = 2     '  Device classification
Global Const HORZSIZE = 4       '  Horizontal size in millimeters
Global Const VERTSIZE = 6       '  Vertical size in millimeters
Global Const HORZRES = 8        '  Horizontal width in pixels
Global Const VERTRES = 10       '  Vertical width in pixels
Global Const BITSPIXEL = 12     '  Number of bits per pixel
Global Const PLANES = 14        '  Number of planes
Global Const NUMBRUSHES = 16    '  Number of brushes the device has
Global Const NUMPENS = 18       '  Number of pens the device has
Global Const NUMMARKERS = 20    '  Number of markers the device has
Global Const NUMFONTS = 22      '  Number of fonts the device has
Global Const NUMCOLORS = 24     '  Number of colors the device supports
Global Const PDEVICESIZE = 26   '  Size required for device descriptor
Global Const CURVECAPS = 28     '  Curve capabilities
Global Const LINECAPS = 30      '  Line capabilities
Global Const POLYGONALCAPS = 32 '  Polygonal capabilities
Global Const TEXTCAPS = 34      '  Text capabilities
Global Const CLIPCAPS = 36      '  Clipping capabilities
Global Const RASTERCAPS = 38    '  Bitblt capabilities
Global Const ASPECTX = 40       '  Length of the X leg
Global Const ASPECTY = 42       '  Length of the Y leg
Global Const ASPECTXY = 44      '  Length of the hypotenuse

Global Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Global Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Global Const SIZEPALETTE = 104  '  Number of entries in physical palette
Global Const NUMRESERVED = 106  '  Number of reserved entries in palette
Global Const COLORRES = 108     '  Actual color resolution


'  Device Capability Masks:

''  Device Technologies
Global Const DT_PLOTTER = 0     '  Vector plotter
Global Const DT_RASDISPLAY = 1  '  Raster display
Global Const DT_RASPRINTER = 2  '  Raster printer
Global Const DT_RASCAMERA = 3   '  Raster camera
Global Const DT_CHARSTREAM = 4  '  Character-stream, PLP
Global Const DT_METAFILE = 5    '  Metafile, VDM
Global Const DT_DISPFILE = 6    '  Display-file

''  Curve Capabilities
Global Const CC_NONE = 0        '  Curves not supported
Global Const CC_CIRCLES = 1     '  Can do circles
Global Const CC_PIE = 2 '  Can do pie wedges
Global Const CC_CHORD = 4       '  Can do chord arcs
Global Const CC_ELLIPSES = 8    '  Can do ellipese
Global Const CC_WIDE = 16       '  Can do wide lines
Global Const CC_STYLED = 32     '  Can do styled lines
Global Const CC_WIDESTYLED = 64 '  Can do wide styled lines
Global Const CC_INTERIORS = 128 '  Can do interiors

''  Line Capabilities
Global Const LC_NONE = 0        '  Lines not supported
Global Const LC_POLYLINE = 2    '  Can do polylines
Global Const LC_MARKER = 4      '  Can do markers
Global Const LC_POLYMARKER = 8  '  Can do polymarkers
Global Const LC_WIDE = 16       '  Can do wide lines
Global Const LC_STYLED = 32     '  Can do styled lines
Global Const LC_WIDESTYLED = 64 '  Can do wide styled lines
Global Const LC_INTERIORS = 128 '  Can do interiors

''  Polygonal Capabilities
Global Const PC_NONE = 0        '  Polygonals not supported
Global Const PC_POLYGON = 1     '  Can do polygons
Global Const PC_RECTANGLE = 2   '  Can do rectangles
Global Const PC_WINDPOLYGON = 4 '  Can do winding polygons
Global Const PC_TRAPEZOID = 4   '  Can do trapezoids
Global Const PC_SCANLINE = 8    '  Can do scanlines
Global Const PC_WIDE = 16       '  Can do wide borders
Global Const PC_STYLED = 32     '  Can do styled borders
Global Const PC_WIDESTYLED = 64 '  Can do wide styled borders
Global Const PC_INTERIORS = 128 '  Can do interiors

''  Polygonal Capabilities
Global Const CP_NONE = 0        '  No clipping of output
Global Const CP_RECTANGLE = 1   '  Output clipped to rects

''  Text Capabilities
Global Const TC_OP_CHARACTER = &H1      '  Can do OutputPrecision       CHARACTER
Global Const TC_OP_STROKE = &H2 '  Can do OutputPrecision       STROKE
Global Const TC_CP_STROKE = &H4 '  Can do ClipPrecision STROKE
Global Const TC_CR_90 = &H8     '  Can do CharRotAbility        90
Global Const TC_CR_ANY = &H10   '  Can do CharRotAbility        ANY
Global Const TC_SF_X_YINDEP = &H20      '  Can do ScaleFreedom  X_YINDEPENDENT
Global Const TC_SA_DOUBLE = &H40        '  Can do ScaleAbility  DOUBLE
Global Const TC_SA_INTEGER = &H80       '  Can do ScaleAbility  INTEGER
Global Const TC_SA_CONTIN = &H100       '  Can do ScaleAbility  CONTINUOUS
Global Const TC_EA_DOUBLE = &H200       '  Can do EmboldenAbility       DOUBLE
Global Const TC_IA_ABLE = &H400 '  Can do ItalisizeAbility      ABLE
Global Const TC_UA_ABLE = &H800 '  Can do UnderlineAbility      ABLE
Global Const TC_SO_ABLE = &H1000        '  Can do StrikeOutAbility      ABLE
Global Const TC_RA_ABLE = &H2000        '  Can do RasterFontAble        ABLE
Global Const TC_VA_ABLE = &H4000        '  Can do VectorFontAble        ABLE
Global Const TC_RESERVED = &H8000

''  Raster Capabilities
Global Const RC_BITBLT = 1      '  Can do standard BLT.
Global Const RC_BANDING = 2     '  Device requires banding support
Global Const RC_SCALING = 4     '  Device requires scaling support
Global Const RC_BITMAP64 = 8    '  Device can support >64K bitmap
Global Const RC_GDI20_OUTPUT = &H10     '  has 2.0 output calls
Global Const RC_DI_BITMAP = &H80        '  supports DIB to memory
Global Const RC_PALETTE = &H100 '  supports a palette
Global Const RC_DIBTODEV = &H200        '  supports DIBitsToDevice
Global Const RC_BIGFONT = &H400 '  supports >64K fonts
Global Const RC_STRETCHBLT = &H800      '  supports StretchBlt
Global Const RC_FLOODFILL = &H1000      '  supports FloodFill
Global Const RC_STRETCHDIB = &H2000     '  supports StretchDIBits

''  palette entry flags
Global Const PC_RESERVED = &H1  '  palette index used for animation
Global Const PC_EXPLICIT = &H2  '  palette index is explicit to device
Global Const PC_NOCOLLAPSE = &H4        '  do not match color to system palette

''  DIB color table identifiers
Global Const DIB_RGB_COLORS = 0 '  color table in RGBTriples
Global Const DIB_PAL_COLORS = 1 '  color table in palette indices

''  constants for Get/SetSystemPaletteUse()
Global Const SYSPAL_STATIC = 1
Global Const SYSPAL_NOSTATIC = 2

''  constants for CreateDIBitmap
Global Const CBM_INIT = &H4&    '  initialize bitmap

''  DrawText() Format Flags




''  ExtFloodFill style flags
Global Const FLOODFILLBORDER = 0
Global Const FLOODFILLSURFACE = 1


''  Scroll Bar Constants
Global Const SB_HORZ = 0
Global Const SB_VERT = 1
Global Const SB_CTL = 2
Global Const SB_BOTH = 3

''  Scroll Bar Commands
Global Const SB_LINEUP = 0
Global Const SB_LINEDOWN = 1
Global Const SB_PAGEUP = 2
Global Const SB_PAGEDOWN = 3
Global Const SB_THUMBPOSITION = 4
Global Const SB_THUMBTRACK = 5
Global Const SB_TOP = 6
Global Const SB_BOTTOM = 7
Global Const SB_ENDSCROLL = 8

''  ShowWindow() Commands
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_NORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9

''  Old ShowWindow() Commands
Global Const HIDE_WINDOW = 0
Global Const SHOW_OPENWINDOW = 1
Global Const SHOW_ICONWINDOW = 2
Global Const SHOW_FULLSCREEN = 3
Global Const SHOW_OPENNOACTIVATE = 4

''  Identifiers for the WM_SHOWWINDOW message
Global Const SW_PARENTCLOSING = 1
Global Const SW_OTHERZOOM = 2
Global Const SW_PARENTOPENING = 3
Global Const SW_OTHERUNZOOM = 4

'' RedrawWindow flags
Global Const RDW_INVALIDATE = &H1
Global Const RDW_INTERNALPAINT = &H2
Global Const RDW_ERASE = &H4
Global Const RDW_VALIDATE = &H8
Global Const RDW_NOINTERNALPAINT = &H10
Global Const RDW_NOERASE = &H20
Global Const RDW_NOCHILDREN = &H40
Global Const RDW_ALLCHILDREN = &H80
Global Const RDW_UPDATENOW = &H100
Global Const RDW_ERASENOW = &H200
Global Const RDW_FRAME = &H400
Global Const RDW_NOFRAME = &H800

'' ScrollWindowEx flags
Global Const SW_SCROLLCHILDREN = &H1
Global Const SW_INVALIDATE = &H2
Global Const SW_ERASE = &H4

''  Region Flags
Global Const ERRORAPI = 0
Global Const NULLREGION = 1
Global Const SIMPLEREGION = 2
Global Const COMPLEXREGION = 3

''  CombineRgn() Styles
Global Const RGN_AND = 1
Global Const RGN_OR = 2
Global Const RGN_XOR = 3
Global Const RGN_DIFF = 4
Global Const RGN_COPY = 5

''  Virtual Keys, Standard Set
Global Const VK_LBUTTON = &H1
Global Const VK_RBUTTON = &H2
Global Const VK_CANCEL = &H3
Global Const VK_MBUTTON = &H4
Global Const VK_BACK = &H8
Global Const VK_TAB = &H9
Global Const VK_CLEAR = &HC
Global Const VK_RETURN = &HD
Global Const VK_SHIFT = &H10
Global Const VK_CONTROL = &H11
Global Const VK_MENU = &H12
Global Const VK_PAUSE = &H13
Global Const VK_CAPITAL = &H14
Global Const VK_ESCAPE = &H1B
Global Const VK_SPACE = &H20
Global Const VK_PRIOR = &H21
Global Const VK_NEXT = &H22
Global Const VK_END = &H23
Global Const VK_HOME = &H24
Global Const VK_LEFT = &H25
Global Const VK_UP = &H26
Global Const VK_RIGHT = &H27
Global Const VK_DOWN = &H28
Global Const VK_SELECT = &H29
Global Const VK_PRINT = &H2A
Global Const VK_EXECUTE = &H2B
Global Const VK_SNAPSHOT = &H2C
Global Const VK_INSERT = &H2D
Global Const VK_DELETE = &H2E
Global Const VK_HELP = &H2F
Global Const VK_NUMPAD0 = &H60
Global Const VK_NUMPAD1 = &H61
Global Const VK_NUMPAD2 = &H62
Global Const VK_NUMPAD3 = &H63
Global Const VK_NUMPAD4 = &H64
Global Const VK_NUMPAD5 = &H65
Global Const VK_NUMPAD6 = &H66
Global Const VK_NUMPAD7 = &H67
Global Const VK_NUMPAD8 = &H68
Global Const VK_NUMPAD9 = &H69
Global Const VK_MULTIPLY = &H6A
Global Const VK_ADD = &H6B
Global Const VK_SEPARATOR = &H6C
Global Const VK_SUBTRACT = &H6D
Global Const VK_DECIMAL = &H6E
Global Const VK_DIVIDE = &H6F
Global Const VK_F1 = &H70
Global Const VK_F2 = &H71
Global Const VK_F3 = &H72
Global Const VK_F4 = &H73
Global Const VK_F5 = &H74
Global Const VK_F6 = &H75
Global Const VK_F7 = &H76
Global Const VK_F8 = &H77
Global Const VK_F9 = &H78
Global Const VK_F10 = &H79
Global Const VK_F11 = &H7A
Global Const VK_F12 = &H7B
Global Const VK_F13 = &H7C
Global Const VK_F14 = &H7D
Global Const VK_F15 = &H7E
Global Const VK_F16 = &H7F
Global Const VK_F17 = &H80
Global Const VK_F18 = &H81
Global Const VK_F19 = &H82
Global Const VK_F20 = &H83
Global Const VK_F21 = &H84
Global Const VK_F22 = &H85
Global Const VK_F23 = &H86
Global Const VK_F24 = &H87
Global Const VK_NUMLOCK = &H90
Global Const VK_SCROLL = &H91

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '0'


'' Queue Status

Global Const QS_KEY = 1
Global Const QS_MOUSEMOVE = 2
Global Const QS_MOUSEBUTTON = 4
Global Const QS_MOUSE = 6
Global Const QS_POSTMESSAGE = 8
Global Const QS_TIMER = &H10
Global Const QS_PAINT = &H20
Global Const QS_SENDMESSAGE = &H40
Global Const QS_ALLINPUT = &H7F

''  SetWindowsHook() codes
Global Const WH_MSGFILTER = (-1)
Global Const WH_JOURNALRECORD = 0
Global Const WH_JOURNALPLAYBACK = 1
Global Const WH_KEYBOARD = 2
Global Const WH_GETMESSAGE = 3
Global Const WH_CALLWNDPROC = 4
Global Const WH_CBT = 5
Global Const WH_SYSMSGFILTER = 6
Global Const WH_WINDOWMGR = 7
Global Const WH_HARDWARE = 8
Global Const WH_SHELL = 10

''  Hook Codes
Global Const HC_LPLPFNNEXT = (-2)
Global Const HC_LPFNNEXT = (-1)
Global Const HC_ACTION = 0
Global Const HC_GETNEXT = 1
Global Const HC_SKIP = 2
Global Const HC_NOREM = 3
Global Const HC_NOREMOVE = 3
Global Const HC_SYSMODALON = 4
Global Const HC_SYSMODALOFF = 5

''  CBT Hook Codes
Global Const HCBT_MOVESIZE = 0
Global Const HCBT_MINMAX = 1
Global Const HCBT_QS = 2

''  WH_MSGFILTER Filter Proc Codes
Global Const MSGF_DIALOGBOX = 0
Global Const MSGF_MESSAGEBOX = 1
Global Const MSGF_MENU = 2
Global Const MSGF_MOVE = 3
Global Const MSGF_SIZE = 4
Global Const MSGF_SCROLLBAR = 5
Global Const MSGF_NEXTWINDOW = 6

''  Window Manager Hook Codes
Global Const WC_INIT = 1
Global Const WC_SWP = 2
Global Const WC_DEFWINDOWPROC = 3
Global Const WC_MINMAX = 4
Global Const WC_MOVE = 5
Global Const WC_SIZE = 6
Global Const WC_DRAWCAPTION = 7


''  Window field offsets for GetWindowLong() and GetWindowWord()





'' GetWindowLong and and GetWindowWord dialog box constants
Global Const DWL_MSGRESULT = 0
Global Const DWL_DLGPROC = 4
Global Const DWL_USER = 8


''  Class field offsets for GetClassLong() and GetClassWord()
Global Const GCL_MENUNAME = (-8)
Global Const GCW_HBRBACKGROUND = (-10)
Global Const GCW_HCURSOR = (-12)
Global Const GCW_HICON = (-14)
Global Const GCW_HMODULE = (-16)
Global Const GCW_CBWNDEXTRA = (-18)
Global Const GCW_CBCLSEXTRA = (-20)
Global Const GCL_WNDPROC = (-24)
Global Const GCW_STYLE = (-26)
Global Const GCW_ATOM = (-32)

'' SendMessage Flag
Global Const HWND_BROADCAST = -1

''  Window Messages
Global Const WM_NULL = &H0
Global Const WM_CREATE = &H1
Global Const WM_DESTROY = &H2
Global Const WM_SIZE = &H5

Global Const WM_SETFOCUS = &H7
Global Const WM_KILLFOCUS = &H8
Global Const WM_ENABLE = &HA
Global Const WM_SETREDRAW = &HB
Global Const WM_SETTEXT = &HC
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_PAINT = &HF

Global Const WM_QUERYENDSESSION = &H11
Global Const WM_QUIT = &H12
Global Const WM_QUERYOPEN = &H13
Global Const WM_ERASEBKGND = &H14
Global Const WM_SYSCOLORCHANGE = &H15
Global Const WM_ENDSESSION = &H16
Global Const WM_SYSTEMERROR = &H17
Global Const WM_SHOWWINDOW = &H18
Global Const WM_CTLCOLOR = &H19
Global Const WM_WININICHANGE = &H1A
Global Const WM_DEVMODECHANGE = &H1B
Global Const WM_ACTIVATEAPP = &H1C
Global Const WM_FONTCHANGE = &H1D
Global Const WM_TIMECHANGE = &H1E
Global Const WM_CANCELMODE = &H1F
Global Const WM_SETCURSOR = &H20
Global Const WM_MOUSEACTIVATE = &H21
Global Const WM_CHILDACTIVATE = &H22
Global Const WM_QUEUESYNC = &H23
Global Const WM_GETMINMAXINFO = &H24
Global Const WM_PAINTICON = &H26
Global Const WM_ICONERASEBKGND = &H27
Global Const WM_NEXTDLGCTL = &H28
Global Const WM_SPOOLERSTATUS = &H2A
Global Const WM_DELETEITEM = &H2D
Global Const WM_VKEYTOITEM = &H2E
Global Const WM_CHARTOITEM = &H2F
Global Const WM_SETFONT = &H30
Global Const WM_GETFONT = &H31
Global Const WM_COMMNOTIFY = &H44
Global Const WM_QUERYDRAGICON = &H37
Global Const WM_COMPAREITEM = &H39
Global Const WM_COMPACTING = &H41
Global Const WM_WINDOWPOSCHANGING = &H46
Global Const WM_WINDOWPOSCHANGED = &H47
Global Const WM_POWER = &H48
Global Const WM_KEYFIRST = &H100
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Global Const WM_CHAR = &H102
Global Const WM_DEADCHAR = &H103
Global Const WM_SYSKEYDOWN = &H104
Global Const WM_SYSKEYUP = &H105
Global Const WM_SYSCHAR = &H106
Global Const WM_SYSDEADCHAR = &H107
Global Const WM_KEYLAST = &H108
Global Const WM_INITDIALOG = &H110
Global Const WM_COMMAND = &H111
Global Const WM_SYSCOMMAND = &H112
Global Const WM_TIMER = &H113
Global Const WM_HSCROLL = &H114
Global Const WM_VSCROLL = &H115
Global Const WM_INITMENU = &H116
Global Const WM_INITMENUPOPUP = &H117
Global Const WM_MENUSELECT = &H11F
Global Const WM_MENUCHAR = &H120
Global Const WM_ENTERIDLE = &H121
Global Const WM_MOUSEFIRST = &H200
Global Const WM_MOUSEMOVE = &H200
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONDOWN = &H204
Global Const WM_RBUTTONUP = &H205
Global Const WM_RBUTTONDBLCLK = &H206
Global Const WM_MBUTTONDOWN = &H207
Global Const WM_MBUTTONUP = &H208
Global Const WM_MBUTTONDBLCLK = &H209
Global Const WM_MOUSELAST = &H209
Global Const WM_PARENTNOTIFY = &H210
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
Global Const WM_DROPFILES = &H233
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Global Const WM_CLEAR = &H303
Global Const WM_UNDO = &H304
Global Const WM_RENDERFORMAT = &H305
Global Const WM_RENDERALLFORMATS = &H306
Global Const WM_DESTROYCLIPBOARD = &H307
Global Const WM_DRAWCLIPBOARD = &H308
Global Const WM_PAINTCLIPBOARD = &H309
Global Const WM_VSCROLLCLIPBOARD = &H30A
Global Const WM_SIZECLIPBOARD = &H30B
Global Const WM_ASKCBFORMATNAME = &H30C
Global Const WM_CHANGECBCHAIN = &H30D
Global Const WM_HSCROLLCLIPBOARD = &H30E
Global Const WM_QUERYNEWPALETTE = &H30F
Global Const WM_PALETTEISCHANGING = &H310
Global Const WM_PALETTECHANGED = &H311

'  NOTE: All Message Numbers below &H0400 are RESERVED.

'  Private Window Messages Start Here:
Global Const WM_USER = &H400
Global Const LB_SELECTSTRING = (WM_USER + 13)
Global Const LB_FINDSTRING = (WM_USER + 16)

''  WM_SYNCTASK Commands
Global Const ST_BEGINSWP = 0
Global Const ST_ENDSWP = 1


'' WM_ACTIVATE constants
Global Const WA_INACTIVE = 0
Global Const WA_ACTIVE = 1
Global Const WA_CLICKACTIVE = 2


''  WinWhere() Area Codes
Global Const HTERROR = (-2)
'Global Const HTTRANSPARENT = (-1)
'Global Const HTNOWHERE = 0
'Global Const HTCLIENT = 1
'Global Const HTCAPTION = 2
'Global Const HTSYSMENU = 3
'Global Const HTGROWBOX = 4
'Global Const HTSIZE = HTGROWBOX
'Global Const HTMENU = 5
'Global Const HTHSCROLL = 6
'Global Const HTVSCROLL = 7
'Global Const HTREDUCE = 8
'Global Const HTZOOM = 9
'Global Const HTLEFT = 10
'Global Const HTRIGHT = 11
'Global Const HTTOP = 12
'Global Const HTTOPLEFT = 13
'Global Const HTTOPRIGHT = 14
'Global Const HTBOTTOM = 15
'Global Const HTBOTTOMLEFT = 16
'Global Const HTBOTTOMRIGHT = 17
'Global Const HTSIZEFIRST = HTLEFT
'Global Const HTSIZELAST = HTBOTTOMRIGHT

''  WM_MOUSEACTIVATE Return Codes
Global Const MA_ACTIVATE = 1
Global Const MA_ACTIVATEANDEAT = 2
Global Const MA_NOACTIVATE = 3
Global Const MA_NOACTIVATEANDEAT = 4


''  Size Message Commands
Global Const SIZENORMAL = 0
Global Const SIZEICONIC = 1
Global Const SIZEFULLSCREEN = 2
Global Const SIZEZOOMSHOW = 3
Global Const SIZEZOOMHIDE = 4

''  Key State Masks for Mouse Messages
Global Const MK_LBUTTON = &H1
Global Const MK_RBUTTON = &H2
Global Const MK_SHIFT = &H4
Global Const MK_CONTROL = &H8
Global Const MK_MBUTTON = &H10

''  Window Styles
Global Const WS_OVERLAPPED = &H0&
Global Const WS_POPUP = &H80000000
Global Const WS_CHILD = &H40000000
Global Const WS_MINIMIZE = &H20000000
Global Const WS_VISIBLE = &H10000000
Global Const WS_DISABLED = &H8000000
Global Const WS_CLIPSIBLINGS = &H4000000
Global Const WS_CLIPCHILDREN = &H2000000
Global Const WS_MAXIMIZE = &H1000000
Global Const WS_CAPTION = &HC00000      '  WS_BORDER Or WS_DLGFRAME
Global Const WS_BORDER = &H800000
Global Const WS_DLGFRAME = &H400000
Global Const WS_VSCROLL = &H200000
Global Const WS_HSCROLL = &H100000
Global Const WS_SYSMENU = &H80000
Global Const WS_THICKFRAME = &H40000
Global Const WS_GROUP = &H20000
Global Const WS_TABSTOP = &H10000

Global Const WS_MINIMIZEBOX = &H20000
Global Const WS_MAXIMIZEBOX = &H10000

Global Const WS_TILED = WS_OVERLAPPED
Global Const WS_ICONIC = WS_MINIMIZE
Global Const WS_SIZEBOX = WS_THICKFRAME

'  Common Window Styles
Global Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Global Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Global Const WS_CHILDWINDOW = (WS_CHILD)
Global Const WS_TILEDWINDOW = (WS_OVERLAPPEDWINDOW)

'  Extended Window Styles
Global Const WS_EX_DLGMODALFRAME = &H1&
Global Const WS_EX_NOPARENTNOTIFY = &H4&
Global Const WS_EX_TOPMOST = &H8&
Global Const WS_EX_ACCEPTFILES = &H10&
Global Const WS_EX_TRANSPARENT = &H20&

' MDI style allows use of all child styles
Global Const MDIS_ALLCHILDSTYLES = &H1&

''  Class styles
Global Const CS_VREDRAW = &H1
Global Const CS_HREDRAW = &H2
Global Const CS_KEYCVTWINDOW = &H4
Global Const CS_DBLCLKS = &H8
Global Const CS_OWNDC = &H20
Global Const CS_CLASSDC = &H40
Global Const CS_PARENTDC = &H80
Global Const CS_NOKEYCVT = &H100
Global Const CS_NOCLOSE = &H200
Global Const CS_SAVEBITS = &H800
Global Const CS_BYTEALIGNCLIENT = &H1000
Global Const CS_BYTEALIGNWINDOW = &H2000
Global Const CS_GLOBALCLASS = &H4000    '  Global window class

''  Predefined Clipboard Formats


Global Const CF_METAFILEPICT = 3
Global Const CF_SYLK = 4
Global Const CF_DIF = 5
Global Const CF_TIFF = 6
Global Const CF_OEMTEXT = 7



Global Const CF_OWNERDISPLAY = &H80
Global Const CF_DSPTEXT = &H81
Global Const CF_DSPBITMAP = &H82
Global Const CF_DSPMETAFILEPICT = &H83

'  "Private" formats don't get GlobalFree()'d
Global Const CF_PRIVATEFIRST = &H200
Global Const CF_PRIVATELAST = &H2FF

'  "GDIOBJ" formats do get DeleteObject()'d
Global Const CF_GDIOBJFIRST = &H300
Global Const CF_GDIOBJLAST = &H3FF


''  Owner draw control types
Global Const ODT_MENU = 1
Global Const ODT_LISTBOX = 2
Global Const ODT_COMBOBOX = 3
Global Const ODT_BUTTON = 4

'  Owner draw actions


'  Owner draw state



''  PeekMessage() Options
Global Const PM_NOREMOVE = &H0
Global Const PM_REMOVE = &H1
Global Const PM_NOYIELD = &H2

''  Flags for _lopen
Global Const READAPI = 0
Global Const WRITEAPI = 1
Global Const READ_WRITE = 2


'' Window placement flags
Global Const CW_USEDEFAULT = &H8000
Global Const WPF_SETMINPOSITION = 1
Global Const WPF_RESTORETOMAXIMIZED = 2

''  SetWindowPos Flags


Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8

Global Const SWP_DRAWFRAME = &H20

Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOREPOSITION = &H200
Global Const SWP_NOSENDCHANGING = &H400
Global Const SWP_DEFERERASE = &H2000

'' SetWindowPos() hwndInsertAfter values
Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
Global Const DLGWINDOWEXTRA = 30        '  Window extra bytes needed for private dialog classes



''  GetSystemMetrics() codes



'' System parameters support

Global Const SPI_GETBEEP = 1
Global Const SPI_SETBEEP = 2
Global Const SPI_GETMOUSE = 3
Global Const SPI_SETMOUSE = 4
Global Const SPI_GETBORDER = 5
Global Const SPI_SETBORDER = 6
Global Const SPI_GETKEYBOARDSPEED = 10
Global Const SPI_SETKEYBOARDSPEED = 11
Global Const SPI_LANGDRIVER = 12
Global Const SPI_ICONHORIZONTALSPACING = 13
Global Const SPI_GETSCREENSAVETIMEOUT = 14
Global Const SPI_SETSCREENSAVETIMEOUT = 15
Global Const SPI_GETSCREENSAVEACTIVE = 16
Global Const SPI_SETSCREENSAVEACTIVE = 17
Global Const SPI_GETGRIDGRANULARITY = 18
Global Const SPI_SETGRIDGRANULARITY = 19
Global Const SPI_SETDESKWALLPAPER = 20
Global Const SPI_SETDESKPATTERN = 21
Global Const SPI_GETKEYBOARDDELAY = 22
Global Const SPI_SETKEYBOARDDELAY = 23
Global Const SPI_ICONVERTICALSPACING = 24
Global Const SPI_GETICONTITLEWRAP = 25
Global Const SPI_SETICONTITLEWRAP = 26
Global Const SPI_GETMENUDROPALIGNMENT = 27
Global Const SPI_SETMENUDROPALIGNMENT = 28
Global Const SPI_SETDOUBLECLKWIDTH = 29
Global Const SPI_SETDOUBLECLKHEIGHT = 30
Global Const SPI_GETICONTITLELOGFONT = 31
Global Const SPI_SETDOUBLECLICKTIME = 32
Global Const SPI_SETMOUSEBUTTONSWAP = 33
Global Const SPI_SETICONTITLELOGFONT = 34
Global Const SPI_GETFASTTASKSWITCH = 35
Global Const SPI_SETFASTTASKSWITCH = 36

' SystemParametersInfo flags
Global Const SPIF_UPDATEINIFILE = 1
Global Const SPIF_SENDWININICHANGE = 2
Global Const MB_ICONHAND = &H10
Global Const MB_ICONASTERISK = &H40
Global Const MB_TASKMODAL = &H2000
Global Const MB_NOFOCUS = &H8000
Global Const MB_TYPEMASK = &HF
Global Const MB_ICONMASK = &HF0
Global Const MB_DEFMASK = &HF00
Global Const MB_MODEMASK = &H3000
Global Const MB_MISCMASK = &HC000


''  Color Types
Global Const CTLCOLOR_MSGBOX = 0
Global Const CTLCOLOR_EDIT = 1
Global Const CTLCOLOR_LISTBOX = 2
Global Const CTLCOLOR_BTN = 3
Global Const CTLCOLOR_DLG = 4
Global Const CTLCOLOR_SCROLLBAR = 5
Global Const CTLCOLOR_STATIC = 6
Global Const CTLCOLOR_MAX = 8   '  three bits max







''  GetWindow() Constants
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5
Global Const GW_ProGGer = 2

'' GetDCEx flags

Global Const DCX_WINDOW = &H1&
Global Const DCX_CACHE = &H2&
Global Const DCX_CLIPCHILDREN = &H8&
Global Const DCX_CLIPSIBLINGS = &H10&
Global Const DCX_PARENTCLIP = &H20&
Global Const DCX_EXCLUDERGN = &H40&
Global Const DCX_INTERSECTRGN = &H80&
Global Const DCX_LOCKWINDOWUPDATE = &H400&
Global Const DCX_USESTYLE = &H10000



''  Menu flags for Add/Check/EnableMenuItem()
Global Const MF_INSERT = &H0
Global Const MF_CHANGE = &H80
Global Const MF_APPEND = &H100
Global Const MF_DELETE = &H200
Global Const MF_REMOVE = &H1000
Global Const MF_BYCOMMAND = &H0
Global Const MF_ENABLED = &H0
Global Const MF_GRAYED = &H1
Global Const MF_DISABLED = &H2
Global Const MF_UNCHECKED = &H0
Global Const MF_CHECKED = &H8
Global Const MF_USECHECKBITMAPS = &H200
Global Const MF_STRING = &H0
Global Const MF_BITMAP = &H4
Global Const MF_OWNERDRAW = &H100
Global Const MF_MENUBARBREAK = &H20
Global Const MF_MENUBREAK = &H40
Global Const MF_UNHILITE = &H0
Global Const MF_HILITE = &H80
Global Const MF_SYSMENU = &H2000
Global Const MF_HELP = &H4000
Global Const MF_MOUSESELECT = &H8000
Global Const MF_END = &H80

'' TrackPopupMenu flags
Global Const TPM_LEFTBUTTON = &H0
Global Const TPM_RIGHTBUTTON = &H2
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_CENTERALIGN = &H4
Global Const TPM_RIGHTALIGN = &H8

''  System Menu Command Values
Global Const SC_SIZE = &HF000
Global Const SC_MOVE = &HF010
Global Const SC_MINIMIZE = &HF020
Global Const SC_MAXIMIZE = &HF030
Global Const SC_NEXTWINDOW = &HF040
Global Const SC_PREVWINDOW = &HF050
Global Const SC_CLOSE = &HF060
Global Const SC_VSCROLL = &HF070
Global Const SC_HSCROLL = &HF080
Global Const SC_MOUSEMENU = &HF090
Global Const SC_KEYMENU = &HF100
Global Const SC_ARRANGE = &HF110
Global Const SC_RESTORE = &HF120
Global Const SC_TASKLIST = &HF130

Global Const SC_ICON = SC_MINIMIZE
Global Const SC_ZOOM = SC_MAXIMIZE

''  Standard Cursor IDs
Global Const IDC_ARROW = 32512&
Global Const IDC_IBEAM = 32513&
Global Const IDC_WAIT = 32514&
Global Const IDC_CROSS = 32515&
Global Const IDC_UPARROW = 32516&
Global Const IDC_SIZE = 32640&
Global Const IDC_ICON = 32641&
Global Const IDC_SIZENWSE = 32642&
Global Const IDC_SIZENESW = 32643&
Global Const IDC_SIZEWE = 32644&
Global Const IDC_SIZENS = 32645&


Global Const ORD_LANGDRIVER = 1 '  The ordinal number for the entry point of
                                '  language drivers.

''  Standard Icon IDs
Global Const IDI_APPLICATION = 32512&
Global Const IDI_HAND = 32513&
Global Const IDI_QUESTION = 32514&
Global Const IDI_EXCLAMATION = 32515&
Global Const IDI_ASTERISK = 32516&


'  Control Manager Structures and Definitions

''  Edit Control Styles
Global Const ES_LEFT = &H0&
Global Const ES_CENTER = &H1&
Global Const ES_RIGHT = &H2&
Global Const ES_MULTILINE = &H4&
Global Const ES_UPPERCASE = &H8&
Global Const ES_LOWERCASE = &H10&
Global Const ES_PASSWORD = &H20&
Global Const ES_AUTOVSCROLL = &H40&
Global Const ES_AUTOHSCROLL = &H80&
Global Const ES_NOHIDESEL = &H100&
Global Const ES_OEMCONVERT = &H400&
Global Const ES_READONLY = &H800&
Global Const ES_WANTRETURN = &H1000&

''  Edit Control Notification Codes
Global Const EN_SETFOCUS = &H100
Global Const EN_KILLFOCUS = &H200
Global Const EN_CHANGE = &H300
Global Const EN_UPDATE = &H400
Global Const EN_ERRSPACE = &H500
Global Const EN_MAXTEXT = &H501
Global Const EN_HSCROLL = &H601
Global Const EN_VSCROLL = &H602

''  Edit Control Messages
Global Const EM_GETSEL = WM_USER + 0
Global Const EM_SETSEL = WM_USER + 1
Global Const EM_GETRECT = WM_USER + 2
Global Const EM_SETRECT = WM_USER + 3
Global Const EM_SETRECTNP = WM_USER + 4
Global Const EM_SCROLL = WM_USER + 5
Global Const EM_LINESCROLL = WM_USER + 6
Global Const EM_GETMODIFY = WM_USER + 8
Global Const EM_SETMODIFY = WM_USER + 9
Global Const EM_GETLINECOUNT = WM_USER + 10
Global Const EM_LINEINDEX = WM_USER + 11
Global Const EM_GETTHUMB = WM_USER + 14
Global Const EM_LINELENGTH = WM_USER + 17
Global Const EM_REPLACESEL = WM_USER + 18
Global Const EM_SETFONT = WM_USER + 19
Global Const EM_GETLINE = WM_USER + 20
Global Const EM_LIMITTEXT = WM_USER + 21
Global Const EM_CANUNDO = WM_USER + 22
Global Const EM_UNDO = WM_USER + 23
Global Const EM_FMTLINES = WM_USER + 24
Global Const EM_LINEFROMCHAR = WM_USER + 25
Global Const EM_SETWORDBREAK = WM_USER + 26
Global Const EM_SETTABSTOPS = WM_USER + 27
Global Const EM_SETPASSWORDCHAR = WM_USER + 28
Global Const EM_EMPTYUNDOBUFFER = WM_USER + 29
Global Const EM_MSGMAX = WM_USER + 30
Global Const EM_GETFIRSTVISIBLELINE = (WM_USER + 30)
Global Const EM_SETREADONLY = (WM_USER + 31)
Global Const EM_SETWORDBREAKPROC = (WM_USER + 32)
Global Const EM_GETWORDBREAKPROC = (WM_USER + 33)
Global Const EM_GETPASSWORDCHAR = (WM_USER + 34)

Global Const WB_LEFT = 0
Global Const WB_RIGHT = 1
Global Const WB_ISDELIMITER = 2

''  Button Control Styles
Global Const BS_PUSHBUTTON = &H0&
Global Const BS_DEFPUSHBUTTON = &H1&
Global Const BS_CHECKBOX = &H2&
Global Const BS_AUTOCHECKBOX = &H3&
Global Const BS_RADIOBUTTON = &H4&
Global Const BS_3STATE = &H5&
Global Const BS_AUTO3STATE = &H6&
Global Const BS_GROUPBOX = &H7&
Global Const BS_USERBUTTON = &H8&
Global Const BS_AUTORADIOBUTTON = &H9&
Global Const BS_PUSHBOX = &HA&
Global Const BS_OWNERDRAW = &HB&
Global Const BS_LEFTTEXT = &H20&

''  User Button Notification Codes
Global Const BN_CLICKED = 0
Global Const BN_PAINT = 1
Global Const BN_HILITE = 2
Global Const BN_UNHILITE = 3
Global Const BN_DISABLE = 4
Global Const BN_DOUBLECLICKED = 5

''  Button Control Messages
Global Const BM_GETCHECK = WM_USER + 0
Global Const BM_SETCHECK = WM_USER + 1
Global Const BM_GETSTATE = WM_USER + 2
Global Const BM_SETSTATE = WM_USER + 3
Global Const BM_SETSTYLE = WM_USER + 4

''  Static Control Constants
Global Const SS_LEFT = &H0&
Global Const SS_CENTER = &H1&
Global Const SS_RIGHT = &H2&
Global Const SS_ICON = &H3&
Global Const SS_BLACKRECT = &H4&
Global Const SS_GRAYRECT = &H5&
Global Const SS_WHITERECT = &H6&
Global Const SS_BLACKFRAME = &H7&
Global Const SS_GRAYFRAME = &H8&
Global Const SS_WHITEFRAME = &H9&
Global Const SS_USERITEM = &HA&
Global Const SS_SIMPLE = &HB&
Global Const SS_LEFTNOWORDWRAP = &HC&
Global Const SS_NOPREFIX = &H80&        '  Don't do "&" character translation


''  Dialog Styles
Global Const DS_ABSALIGN = &H1&
Global Const DS_SYSMODAL = &H2&
Global Const DS_LOCALEDIT = &H20&       '  Edit items get Local storage.
Global Const DS_SETFONT = &H40& '  User specified font for Dlg controls
Global Const DS_MODALFRAME = &H80&      '  Can be combined with WS_CAPTION
Global Const DS_NOIDLEMSG = &H100&      '  WM_ENTERIDLE message will not be sent

Global Const DM_GETDEFID = WM_USER + 0
Global Const DM_SETDEFID = WM_USER + 1
Global Const DC_HASDEFID = &H534      '&H534B

''  Dialog Codes
Global Const DLGC_WANTARROWS = &H1      '  Control wants arrow keys
Global Const DLGC_WANTTAB = &H2 '  Control wants tab keys
Global Const DLGC_WANTALLKEYS = &H4     '  Control wants all keys
Global Const DLGC_WANTMESSAGE = &H4     '  Pass message to control
Global Const DLGC_HASSETSEL = &H8       '  Understands EM_SETSEL message
Global Const DLGC_DEFPUSHBUTTON = &H10  '  Default pushbutton
Global Const DLGC_UNDEFPUSHBUTTON = &H20        '  Non-default pushbutton
Global Const DLGC_RADIOBUTTON = &H40    '  Radio button
Global Const DLGC_WANTCHARS = &H80      '  Want WM_CHAR messages
Global Const DLGC_STATIC = &H100        '  Static item: don't include
Global Const DLGC_BUTTON = &H2000       '  Button item: can be checked

''  Listbox Return Values
Global Const LB_CTLCODE = 0&

Global Const LB_OKAY = 0
Global Const LB_ERR = (-1)
Global Const LB_ERRSPACE = (-2)

'
'  The idStaticPath parameter to DlgDirList can have the following values
'  ORed if the list box should show other details of the files along with
'  the name of the files;

'  all other details also will be returned

''  Listbox Notification Codes
Global Const LBN_ERRSPACE = (-2)
Global Const LBN_SELCHANGE = 1
Global Const LBN_DBLCLK = 2
Global Const LBN_SELCANCEL = 3
Global Const LBN_SETFOCUS = 4
Global Const LBN_KILLFOCUS = 5

''  Listbox messages
Global Const LB_ADDSTRING = (WM_USER + 1)
Global Const LB_INSERTSTRING = (WM_USER + 2)
Global Const LB_DELETESTRING = (WM_USER + 3)
Global Const LB_RESETCONTENT = (WM_USER + 5)
Global Const LB_SETSEL = (WM_USER + 6)
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_GETSEL = (WM_USER + 8)
Global Const LB_GETCURSEL = (WM_USER + 9)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const LB_DIR = (WM_USER + 14)
Global Const LB_GETTOPINDEX = (WM_USER + 15)
Global Const LB_GETSELCOUNT = (WM_USER + 17)
Global Const LB_GETSELITEMS = (WM_USER + 18)
Global Const LB_SETTABSTOPS = (WM_USER + 19)
Global Const LB_GETHORIZONTALEXTENT = (WM_USER + 20)
Global Const LB_SETHORIZONTALEXTENT = (WM_USER + 21)
Global Const LB_SETCOLUMNWIDTH = (WM_USER + 22)
Global Const LB_SETTOPINDEX = (WM_USER + 24)
Global Const LB_GETITEMRECT = (WM_USER + 25)
Global Const LB_GETITEMDATA = (WM_USER + 26)
Global Const LB_SETITEMDATA = (WM_USER + 27)
Global Const LB_SELITEMRANGE = (WM_USER + 28)
Global Const LB_MSGMAX = (WM_USER + 33)
Global Const LB_SETCARETINDEX = (WM_USER + 31)
Global Const LB_GETCARETINDEX = (WM_USER + 32)
Global Const LB_SETITEMHEIGHT = (WM_USER + 33)
Global Const LB_GETITEMHEIGHT = (WM_USER + 34)
Global Const LB_FINDSTRINGEXACT = (WM_USER + 35)

''  Listbox Styles




''  Combo Box return Values
Global Const CB_OKAY = 0
Global Const CB_ERR = (-1)
Global Const CB_ERRSPACE = (-2)

''  Combo Box Notification Codes
Global Const CBN_ERRSPACE = (-1)
Global Const CBN_SELCHANGE = 1
Global Const CBN_DBLCLK = 2
Global Const CBN_SETFOCUS = 3
Global Const CBN_KILLFOCUS = 4
Global Const CBN_EDITCHANGE = 5
Global Const CBN_EDITUPDATE = 6
Global Const CBN_DROPDOWN = 7
Global Const CBN_CLOSEUP = 8
Global Const CBN_SELENDOK = 9
Global Const CBN_SELENDCANCEL = 10

''  Combo Box styles
Global Const CBS_SIMPLE = &H1&
Global Const CBS_DROPDOWN = &H2&
Global Const CBS_DROPDOWNLIST = &H3&
Global Const CBS_OWNERDRAWFIXED = &H10&
Global Const CBS_OWNERDRAWVARIABLE = &H20&
Global Const CBS_AUTOHSCROLL = &H40&
Global Const CBS_OEMCONVERT = &H80&
Global Const CBS_SORT = &H100&
Global Const CBS_HASSTRINGS = &H200&
Global Const CBS_NOINTEGRALHEIGHT = &H400&
Global Const CBS_DISABLENOSCROLL = &H800&

''  Combo Box messages
Global Const CB_GETEDITSEL = (WM_USER + 0)
Global Const CB_LIMITTEXT = (WM_USER + 1)
Global Const CB_SETEDITSEL = (WM_USER + 2)
Global Const CB_ADDSTRING = (WM_USER + 3)
Global Const CB_DELETESTRING = (WM_USER + 4)
Global Const CB_DIR = (WM_USER + 5)
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_GETCURSEL = (WM_USER + 7)
Global Const CB_GETLBTEXT = (WM_USER + 8)
Global Const CB_GETLBTEXTLEN = (WM_USER + 9)
Global Const CB_INSERTSTRING = (WM_USER + 10)
Global Const CB_RESETCONTENT = (WM_USER + 11)
Global Const CB_FINDSTRING = (WM_USER + 12)
Global Const CB_SELECTSTRING = (WM_USER + 13)
Global Const CB_SETCURSEL = (WM_USER + 14)
Global Const CB_SHOWDROPDOWN = (WM_USER + 15)
Global Const CB_GETITEMDATA = (WM_USER + 16)
Global Const CB_SETITEMDATA = (WM_USER + 17)
Global Const CB_GETDROPPEDCONTROLRECT = (WM_USER + 18)
Global Const CB_SETITEMHEIGHT = (WM_USER + 19)
Global Const CB_GETITEMHEIGHT = (WM_USER + 20)
Global Const CB_SETEXTENDEDUI = (WM_USER + 21)
Global Const CB_GETEXTENDEDUI = (WM_USER + 22)
Global Const CB_GETDROPPEDSTATE = (WM_USER + 23)
Global Const CB_FINDSTRINGEXACT = (WM_USER + 24)

''  Scroll Bar Styles
Global Const SBS_HORZ = &H0&
Global Const SBS_VERT = &H1&
Global Const SBS_TOPALIGN = &H2&
Global Const SBS_LEFTALIGN = &H2&
Global Const SBS_BOTTOMALIGN = &H4&
Global Const SBS_RIGHTALIGN = &H4&
Global Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Global Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Global Const SBS_SIZEBOX = &H8&

''  Static messages
Global Const STM_SETICON = (WM_USER + 0)
Global Const STM_GETICON = (WM_USER + 1)

''  WaitSoundState() Constants
Global Const S_QUEUEEMPTY = 0
Global Const S_THRESHOLD = 1
Global Const S_ALLTHRESHOLD = 2

''  Accent Modes
Global Const S_NORMAL = 0
Global Const S_LEGATO = 1
Global Const S_STACCATO = 2

''  SetSoundNoise() Sources
Global Const S_PERIOD512 = 0    '  Freq = N/512 high pitch, less coarse hiss
Global Const S_PERIOD1024 = 1   '  Freq = N/1024
Global Const S_PERIOD2048 = 2   '  Freq = N/2048 low pitch, more coarse hiss
Global Const S_PERIODVOICE = 3  '  Source is frequency from voice channel (3)
Global Const S_WHITE512 = 4     '  Freq = N/512 high pitch, less coarse hiss
Global Const S_WHITE1024 = 5    '  Freq = N/1024
Global Const S_WHITE2048 = 6    '  Freq = N/2048 low pitch, more coarse hiss
Global Const S_WHITEVOICE = 7   '  Source is frequency from voice channel (3)
Global Const S_SERDVNA = (-1)   '  Device not available
Global Const S_SEROFM = (-2)    '  Out of memory
Global Const S_SERMACT = (-3)   '  Music active
Global Const S_SERQFUL = (-4)   '  Queue full
Global Const S_SERBDNT = (-5)   '  Invalid note
Global Const S_SERDLN = (-6)    '  Invalid note length
Global Const S_SERDCC = (-7)    '  Invalid note count
Global Const S_SERDTP = (-8)    '  Invalid tempo
Global Const S_SERDVL = (-9)    '  Invalid volume
Global Const S_SERDMD = (-10)   '  Invalid mode
Global Const S_SERDSH = (-11)   '  Invalid shape
Global Const S_SERDPT = (-12)   '  Invalid pitch
Global Const S_SERDFQ = (-13)   '  Invalid frequency
Global Const S_SERDDR = (-14)   '  Invalid duration
Global Const S_SERDSR = (-15)   '  Invalid source
Global Const S_SERDST = (-16)   '  Invalid state

''  COMM declarations
'
Global Const NOPARITY = 0
Global Const ODDPARITY = 1
Global Const EVENPARITY = 2
Global Const MARKPARITY = 3
Global Const SPACEPARITY = 4
Global Const ONESTOPBIT = 0
Global Const ONE5STOPBITS = 1
Global Const TWOSTOPBITS = 2
Global Const IGNORE = 0 '  Ignore signal
Global Const INFINITE = &HFFFF  '  Infinite timeout

''  COMM Error Flags
Global Const CE_RXOVER = &H1    '  Receive Queue overflow
Global Const CE_OVERRUN = &H2   '  Receive Overrun Error
Global Const CE_RXPARITY = &H4  '  Receive Parity Error
Global Const CE_FRAME = &H8     '  Receive Framing error
Global Const CE_BREAK = &H10    '  Break Detected
Global Const CE_CTSTO = &H20    '  CTS Timeout
Global Const CE_DSRTO = &H40    '  DSR Timeout
Global Const CE_RLSDTO = &H80   '  RLSD Timeout
Global Const CE_TXFULL = &H100  '  TX Queue is full
Global Const CE_PTO = &H200     '  LPTx Timeout
Global Const CE_IOE = &H400     '  LPTx I/O Error
Global Const CE_DNS = &H800     '  LPTx Device not selected
Global Const CE_OOP = &H1000    '  LPTx Out-Of-Paper
Global Const CE_MODE = &H8000   '  Requested mode unsupported
Global Const IE_BADID = (-1)    '  Invalid or unsupported id
Global Const IE_OPEN = (-2)     '  Device Already Open
Global Const IE_NOPEN = (-3)    '  Device Not Open
Global Const IE_MEMORY = (-4)   '  Unable to allocate queues
Global Const IE_DEFAULT = (-5)  '  Error in default parameters
Global Const IE_HARDWARE = (-10)        '  Hardware Not Present
Global Const IE_BYTESIZE = (-11)        '  Illegal Byte Size
Global Const IE_BAUDRATE = (-12)        '  Unsupported BaudRate

''  COMM Events
Global Const EV_RXCHAR = &H1    '  Any Character received
Global Const EV_RXFLAG = &H2    '  Received certain character
Global Const EV_TXEMPTY = &H4   '  Transmitt Queue Empty
Global Const EV_CTS = &H8           '  CTS changed state
Global Const EV_DSR = &H10          '  DSR changed state
Global Const EV_RLSD = &H20         '  RLSD changed state
Global Const EV_BREAK = &H40    '  BREAK received
Global Const EV_ERR = &H80          '  Line status error occurred
Global Const EV_RING = &H100    '  Ring signal detected
Global Const EV_PERR = &H200    '  Printer error occured
Global Const EV_CTSS = &H400  ' Current state of CTS
Global Const EV_DSRS = &H800  ' Current state of DSR
Global Const EV_RLSDS = &H1000 ' Current state of RLSD signal


''  COMM Escape Functions
Global Const SETXOFF = 1        '  Simulate XOFF received
Global Const SETXON = 2     '  Simulate XON received
Global Const SETRTS = 3     '  Set RTS high
Global Const CLRRTS = 4     '  Set RTS low
Global Const SETDTR = 5     '  Set DTR high
Global Const CLRDTR = 6     '  Set DTR low
Global Const RESETDEV = 7       '  Reset device if possible
Global Const GETMAXLPT = 8
Global Const GETMAXCOM = 9
Global Const GETBASEIRQ = 10
Global Const CBR_110 = &HFF10
Global Const CBR_300 = &HFF11
Global Const CBR_600 = &HFF12
Global Const CBR_1200 = &HFF13
Global Const CBR_2400 = &HFF14
Global Const CBR_4800 = &HFF15
Global Const CBR_9600 = &HFF16
Global Const CBR_14400 = &HFF17
Global Const CBR_19200 = &HFF18
Global Const CBR_38400 = &HFF1B
Global Const CBR_56000 = &HFF1F
Global Const CBR_128000 = &HFF23
Global Const CBR_256000 = &HFF27

'' COMM notifications on WM_COMMNOTIFY messages
Global Const CN_RECEIVE = &H1
Global Const CN_TRANSMIT = &H2
Global Const CN_EVENT = &H4

'' COMM status flags
Global Const CSTF_CTSHOLD = &H1
Global Const CSTF_DSRHOLD = &H2
Global Const CSTF_RLSDHOLD = &H4
Global Const CSTF_XOFFHOLD = &H8
Global Const CSTF_XOFFSENT = &H10
Global Const CSTF_EOF = &H20
Global Const CSTF_TXIM = &H40
Global Const LPTx = &H80        '  Set if ID is for LPT device


'   Help engine section.
Global Const HELP_MULTIKEY = &H201
Global Const HELP_SETWINPOS = &H203

'' Printer and device support constants

' field selection bits
Global Const DM_ORIENTATION = &H1&
Global Const DM_PAPERSIZE = &H2&
Global Const DM_PAPERLENGTH = &H4&
Global Const DM_PAPERWIDTH = &H8&
Global Const DM_SCALE = &H10&
Global Const DM_COPIES = &H100&
Global Const DM_DEFAULTSOURCE = &H200&
Global Const DM_PRINTQUALITY = &H400&
Global Const DM_COLOR = &H800&
Global Const DM_DUPLEX = &H1000&
Global Const DM_YRESOLUTION = &H2000&
Global Const DM_TTOPTION = &H4000&

'' printer orientation selections
Global Const DMORIENT_PORTRAIT = 1
Global Const DMORIENT_LANDSCAPE = 2

'' paper selections
' Warning: The PostScript driver mistakingly uses DMPAPER_ values between
' 50 and 56.  Don't use this range when defining new paper sizes.

Global Const DMPAPER_LETTER = 1         'Letter 8 1/2 x 11 in
Global Const DMPAPER_LETTERSMALL = 2    'Letter Small 8 1/2 x 11 in
Global Const DMPAPER_TABLOID = 3        'Tabloid 11 x 17 in
Global Const DMPAPER_LEDGER = 4         'Ledger 17 x 11 in
Global Const DMPAPER_LEGAL = 5          'Legal 8 1/2 x 14 in
Global Const DMPAPER_STATEMENT = 6      'Statement 5 1/2 x 8 1/2 in
Global Const DMPAPER_EXECUTIVE = 7              'Executive 7 1/4 x 10 1/2 in
Global Const DMPAPER_A3 = 8             'A3 297 x 420 mm
Global Const DMPAPER_A4 = 9             'A4 210 x 297 mm
Global Const DMPAPER_A4SMALL = 10       'A4 Small 210 x 297 mm
Global Const DMPAPER_A5 = 11            'A5 148 x 210 mm
Global Const DMPAPER_B4 = 12            'B4 250 x 354
Global Const DMPAPER_B5 = 13            'B5 182 x 257 mm
Global Const DMPAPER_FOLIO = 14         'Folio 8 1/2 x 13 in
Global Const DMPAPER_QUARTO = 15        'Quarto 215 x 275 mm
Global Const DMPAPER_10X14 = 16         '10x14 in
Global Const DMPAPER_11X17 = 17         '11x17 in
Global Const DMPAPER_NOTE = 18          'Note 8 1/2 x 11 in
Global Const DMPAPER_ENV_9 = 19         'Envelope #9 3 7/8 x 8 7/8
Global Const DMPAPER_ENV_10 = 20        'Envelope #10 4 1/8 x 9 1/2
Global Const DMPAPER_ENV_11 = 21        'Envelope #11 4 1/2 x 10 3/8
Global Const DMPAPER_ENV_12 = 22        'Envelope #12 4 \276 x 11
Global Const DMPAPER_ENV_14 = 23        'Envelope #14 5 x 11 1/2
Global Const DMPAPER_CSHEET = 24        'C size sheet
Global Const DMPAPER_DSHEET = 25        'D size sheet
Global Const DMPAPER_ESHEET = 26        'E size sheet
Global Const DMPAPER_ENV_DL = 27        'Envelope DL 110 x 220mm
Global Const DMPAPER_ENV_C5 = 28        'Envelope C5 162 x 229 mm
Global Const DMPAPER_ENV_C3 = 29        'Envelope C3  324 x 458 mm
Global Const DMPAPER_ENV_C4 = 30        'Envelope C4  229 x 324 mm
Global Const DMPAPER_ENV_C6 = 31        'Envelope C6  114 x 162 mm
Global Const DMPAPER_ENV_C65 = 32       'Envelope C65 114 x 229 mm
Global Const DMPAPER_ENV_B4 = 33        'Envelope B4  250 x 353 mm
Global Const DMPAPER_ENV_B5 = 34        'Envelope B5  176 x 250 mm
Global Const DMPAPER_ENV_B6 = 35        'Envelope B6  176 x 125 mm
Global Const DMPAPER_ENV_ITALY = 36     'Envelope 110 x 230 mm
Global Const DMPAPER_ENV_MONARCH = 37   'Envelope Monarch 3.875 x 7.5 in
Global Const DMPAPER_ENV_PERSONAL = 38  '6 3/4 Envelope 3 5/8 x 6 1/2 in
Global Const DMPAPER_FANFOLD_US = 39    'US Std Fanfold 14 7/8 x 11 in
Global Const DMPAPER_FANFOLD_STD_GERMAN = 40  'German Std Fanfold 8 1/2 x 12 in
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41  'German Legal Fanfold 8 1/2 x 13 in
Global Const DMPAPER_USER = 256

'' printer bin selections
Global Const DMBIN_UPPER = 1
Global Const DMBIN_ONLYONE = 1
Global Const DMBIN_LOWER = 2
Global Const DMBIN_MIDDLE = 3
Global Const DMBIN_MANUAL = 4
Global Const DMBIN_ENVELOPE = 5
Global Const DMBIN_ENVMANUAL = 6
Global Const DMBIN_AUTO = 7
Global Const DMBIN_TRACTOR = 8
Global Const DMBIN_SMALLFMT = 9
Global Const DMBIN_LARGEFMT = 10
Global Const DMBIN_LARGECAPACITY = 11
Global Const DMBIN_CASSETTE = 14
Global Const DMBIN_USER = 256  'device specific bins start here

'' print qualities
Global Const DMRES_DRAFT = -1
Global Const DMRES_LOW = -2
Global Const DMRES_MEDIUM = -3
Global Const DMRES_HIGH = -4

'' color enable/disable for color printers
Global Const DMCOLOR_MONOCHROME = 1
Global Const DMCOLOR_COLOR = 2

'' Printer duplex enable
Global Const DMDUP_SIMPLEX = 1
Global Const DMDUP_VERTICAL = 2
Global Const DMDUP_HORIZONTAL = 3

'' TrueType options
Global Const DMTT_BITMAP = 1    'print TT fonts as graphics
Global Const DMTT_DOWNLOAD = 2  'download TT fonts as soft fonts
Global Const DMTT_SUBDEV = 3    'substitute device fonts for TT fonts



'' device mode function modes
Global Const DM_UPDATE = 1
Global Const DM_COPY = 2
Global Const DM_PROMPT = 4
Global Const DM_MODIFY = 8
Global Const DM_IN_BUFFER = 8
Global Const DM_IN_PROMPT = 4
Global Const DM_OUT_BUFFER = 2
Global Const DM_OUT_DEFAULT = 1

'' device capabilities indices
Global Const DC_FIELDS = 1
Global Const DC_PAPERS = 2
Global Const DC_PAPERSIZE = 3
Global Const DC_MINEXTENT = 4
Global Const DC_MAXEXTENT = 5
Global Const DC_BINS = 6
Global Const DC_DUPLEX = 7
Global Const DC_SIZE = 8
Global Const DC_EXTRA = 9
Global Const DC_VERSION = 10
Global Const DC_DRIVER = 11
Global Const DC_BINNAMES = 12
Global Const DC_ENUMRESOLUTIONS = 13
Global Const DC_FILEDEPENDENCIES = 14
Global Const DC_TRUETYPE = 15
Global Const DC_PAPERNAMES = 16
Global Const DC_ORIENTATION = 17
Global Const DC_COPIES = 18

' DC_TRUETYPE bit fields
Global Const DCTT_BITMAP = &H1&
Global Const DCTT_DOWNLOAD = &H2&
Global Const DCTT_SUBDEV = &H4&

'' LZ encode constants
Global Const LZERROR_BADINHANDLE = -1   '   invalid input handle
Global Const LZERROR_BADOUTHANDLE = -2  '   invalid output handle
Global Const LZERROR_READ = -3          '   corrupt compressed file format
Global Const LZERROR_WRITE = -4         '   out of space for output file
Global Const LZERROR_GLOBALLOC = -5     '   insufficient memory for LZFile struct
Global Const LZERROR_GLOBLOCK = -6      '   bad global handle
Global Const LZERROR_BADVALUE = -7      '   input parameter out of range
Global Const LZERROR_UNKNOWNALG = -8    '   compression algorithm not recognized


'' Version Control Resources
Global Const VS_FILE_INFO = 16                          ' Version stamp res type
Global Const VS_VERSION_INFO = 1                ' Version stamp res ID
Global Const VS_USER_DEFINED = 100              ' User-defined res IDs

'' Version control flags
Global Const VS_FFI_SIGNATURE = &HFEEF04BD
Global Const VS_FFI_STRUCVERSION = &H10000
Global Const VS_FFI_FILEFLAGSMASK = &H3F&
Global Const VS_FF_DEBUG = &H1&
Global Const VS_FF_PRERELEASE = &H2&
Global Const VS_FF_PATCHED = &H4&
Global Const VS_FF_PRIVATEBUILD = &H8&
Global Const VS_FF_INFOINFERRED = &H10&
Global Const VS_FF_SPECIALBUILD = &H20&

'' Version control OS flags
Global Const VOS_UNKNOWN = &H0&
Global Const VOS_DOS = &H10000
Global Const VOS_OS216 = &H20000
Global Const VOS_OS232 = &H30000
Global Const VOS_NT = &H40000
Global Const VOS__BASE = &H0&
Global Const VOS__WINDOWS16 = &H1&
Global Const VOS__PM16 = &H2&
Global Const VOS__PM32 = &H3&
Global Const VOS__WINDOWS32 = &H4&
Global Const VOS_DOS_WINDOWS16 = &H10001
Global Const VOS_DOS_WINDOWS32 = &H10004
Global Const VOS_OS216_PM16 = &H20002
Global Const VOS_OS232_PM32 = &H30003
Global Const VOS_NT_WINDOWS32 = &H40004

'' Version control file types
Global Const VFT_UNKNOWN = &H0&
Global Const VFT_APP = &H1&
Global Const VFT_DLL = &H2&
Global Const VFT_DRV = &H3&
Global Const VFT_FONT = &H4&
Global Const VFT_VXD = &H5&
Global Const VFT_STATIC_LIB = &H7&

' VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV
Global Const VFT2_UNKNOWN = &H0&
Global Const VFT2_DRV_PRINTER = &H1&
Global Const VFT2_DRV_KEYBOARD = &H2&
Global Const VFT2_DRV_LANGUAGE = &H3&
Global Const VFT2_DRV_DISPLAY = &H4&
Global Const VFT2_DRV_MOUSE = &H5&
Global Const VFT2_DRV_NETWORK = &H6&
Global Const VFT2_DRV_SYSTEM = &H7&
Global Const VFT2_DRV_INSTALLABLE = &H8&
Global Const VFT2_DRV_SOUND = &H9&
Global Const VFT2_DRV_COMM = &HA&

' VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT
Global Const VFT2_FONT_RASTER = &H1&
Global Const VFT2_FONT_VECTOR = &H2&
Global Const VFT2_FONT_TRUETYPE = &H3&

'' VerFindFile() flags
Global Const VFFF_ISSHAREDFILE = &H1
Global Const VFF_CURNEDEST = &H1
Global Const VFF_FILEINUSE = &H2
Global Const VFF_BUFFTOOSMALL = &H4

'' VerInstallFile() flags
Global Const VIFF_FORCEINSTALL = &H1
Global Const VIFF_DONTDELETEOLD = &H2
Global Const VIF_TEMPFILE = &H1&
Global Const VIF_MISMATCH = &H2&
Global Const VIF_SRCOLD = &H4&
Global Const VIF_DIFFLANG = &H8&
Global Const VIF_DIFFCODEPG = &H10&
Global Const VIF_DIFFTYPE = &H20&
Global Const VIF_WRITEPROT = &H40&
Global Const VIF_FILEINUSE = &H80&
Global Const VIF_OUTOFSPACE = &H100&
Global Const VIF_ACCESSVIOLATION = &H200&
Global Const VIF_SHARINGVIOLATION = &H400&
Global Const VIF_CANNOTCREATE = &H800&
Global Const VIF_CANNOTDELETE = &H1000&
Global Const VIF_CANNOTRENAME = &H2000&
Global Const VIF_CANNOTDELETECUR = &H4000&
Global Const VIF_OUTOFMEMORY = &H8000&
Global Const VIF_CANNOTREADSRC = &H10000
Global Const VIF_CANNOTREADDST = &H20000
Global Const VIF_BUFFTOOSMALL = &H40000

'' WM_POWER window message and DRV_POWER driver notification */
Global Const PWR_OK = 1
Global Const PWR_FAIL = (-1)
Global Const PWR_SUSPENDREQUEST = 1
Global Const PWR_SUSPENDRESUME = 2
Global Const PWR_CRITICALRESUME = 3


'' Network operation return values
Global Const WN_SUCCESS = 0
Global Const WN_NOT_SUPPORTED = 1
Global Const WN_NET_ERROR = 2
Global Const WN_MORE_DATA = 3
Global Const WN_BAD_POINTER = 4
Global Const WN_BAD_VALUE = 5
Global Const WN_BAD_PASSWORD = 6
Global Const WN_ACCESS_DENIED = 7
Global Const WN_FUNCTION_BUSY = 8
Global Const WN_WINDOWS_ERROR = 9
Global Const WN_BAD_USER = &HA
Global Const WN_OUT_OF_MEMORY = &HB
Global Const WN_CANCEL = &HC
Global Const WN_CONTINUE = &HD

'' Network Connection errors
Global Const WN_NOT_CONNECTED = &H30
Global Const WN_OPEN_FILES = &H31
Global Const WN_BAD_NETNAME = &H32
Global Const WN_BAD_LOCALNAME = &H33
Global Const WN_ALREADY_CONNECTED = &H34
Global Const WN_DEVICE_ERROR = &H35
Global Const WN_CONNECTION_CLOSED = &H36
Declare Function SystemParametersInfo Lib "User" (ByVal uAction As Integer, ByVal uparam As Integer, lpvParam As Any, ByVal fuWinIni As Integer) As Integer
'WM_NCHITTEST return codes
'Global Const HTERROR = -2
Global Const HTTRANSPARENT = -1
Global Const HTNOWHERE = 0
Global Const HTCLIENT = 1
Global Const HTCAPTION = 2
Global Const HTSYSMENU = 3
Global Const HTSIZE = 4
Global Const HTMENU = 5
Global Const HTHSCROLL = 6
Global Const HTVSCROLL = 7
Global Const HTMINBUTTON = 8
Global Const HTMAXBUTTON = 9
Global Const HTLEFT = 10
Global Const HTRIGHT = 11
Global Const HTTOP = 12
Global Const HTTOPLEFT = 13
Global Const HTTOPRIGHT = 14
Global Const HTBOTTOM = 15
Global Const HTBOTTOMLEFT = 16
Global Const HTBOTTOMRIGHT = 17
Global Const HTBORDER = 18
Global Const HTGROWBOX = HTSIZE
Global Const HTREDUCE = HTMINBUTTON
Global Const HTZOOM = HTMAXBUTTON

Sub AddCombo(Combo As ComboBox, txt$)
For X = 0 To Combo.ListCount - 1
    If UCase$(Combo.List(X)) = UCase$(txt$) Then Exit Sub
Next
Combo.AddItem txt$
End Sub

Sub addcomboroom(Combo As ComboBox)
Chat% = findchatroom()
AolList% = FindChildByClass(Chat%, "_AOL_ListBox")
num = SendMessageByNum(AolList%, LB_GETCOUNT, 0, 0)
X = SetFocusAPI(Chat%)
For i% = 0 To num - 1
    NameZ$ = String$(11, " ")
    Ret = AOLGetList(i%, NameZ$)
    NameZ$ = Trim$(NameZ$)
    SN$ = usersn()
    
    If Trim$(UCase$(NameZ$)) = Trim$(UCase(SN$)) Then GoTo onn2
    Call AddCombo(Combo, NameZ$)
onn2:
Next

End Sub

Sub AddRoom(lst As ListBox)
lst.Clear
For index% = 0 To 25 'This will loop until everyone's added or it will stop at 25
NameZ$ = String$(256, " ")
Ret = AOLGetList(index%, NameZ$) & ErB$
If Len(Trim$(NameZ$)) <= 1 Then Exit For
NameZ$ = Left$(Trim$(NameZ$), Len(Trim(NameZ$)) - 1)
lst.AddItem (NameZ$)
Next index%
For i = 0 To lst.ListCount
    If lst.List(i) = GetSN() Then lst.RemoveItem i
    Next i
End Sub

Function AOLChild()
Dim Aol%, MDI%, Child%
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
Child% = FindChildByClass(MDI%, "AOL Child")
AOLChild = Child%
End Function

Function AOLGetUer()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUer = User
End Function

Function AOLGetUser()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function

Sub AOLKeyword(Text)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
keyw% = FindChildByTitle(MDI%, "Keyword")
kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending% = SendMessage(kedit%, 258, 13, 0)
pausing = DoEvents()
End Sub

Function AOLMDI()
Aol = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(Aol, "MDIClient")
End Function

Function aolversion()
num = 1
Aol% = FindWindow("AOL Frame25", 0&)
TOO% = FindChildByClass(Aol%, "AOL Toolbar")
ICC% = FindChildByClass(TOO%, "_AOL_Icon")
Do
c% = DoEvents()
ICC% = getwindow(ICC%, GW_HWNDNEXT)
num = num + 1
Loop Until ICC% = 0
Select Case num
Case 19
aolversion = 30
Case 21
aolversion = 25
End Select

End Function

Function AOLWindow()


End Function

Function Backwards(thestring)
End Function

Sub center(formname As Form)
On Error Resume Next
formname.Top = (Screen.Height - formname.Height) / 2
formname.Left = (Screen.Width - formname.Width) / 2
End Sub

Sub ChangeAOLHostName(AOLPath, NewName)
Aol% = FindWindow("AOL Frame25", 0&)
If Aol% = 0 Then
    MsgBox "Please run AOL so I can detect which version of America Online you are using.", 64, "Run America Online!"
    Exit Sub
End If
If aolversion() = 25 Then
    X = FileExists(AOLPath & "\tool\chat.aol")
    If X = 0 Then
        MsgBox AOLPath & "\tool\chat.aol could not be found.", 22, "Error!"
        Exit Sub
    End If
    FNum = FreeFile
    Open AOLPath & "\tool\chat.aol" For Binary As FNum
        Seek FNum, 6885
        Put FNum, , NewName
    Close FNum
    Exit Sub
End If
If aolversion() = 30 Then
    X = FileExists(AOLPath & "\tool\aolchat.aol")
    If X = 0 Then
        MsgBox AOLPath & "\tool\aolchat.aol could not be found.", 22, "Error!"
        Exit Sub
    End If
    FNum = FreeFile
    Open AOLPath & "\tool\aolchat.aol" For Binary As FNum
        Seek FNum, 4747
        FF = NewName & "hello lamer" & NewName
        Put FNum, , FF
    Close FNum
    Exit Sub
End If

End Sub

Function ChatView() As Integer
Chat% = findchatroom()
view% = FindChildByClass(Chat%, "_AOL_View")
ChatView = view%
End Function

Sub Click(button%)
D = SendMessageByNum(button%, WM_LBUTTONDOWN, 0, 0)
D = SendMessageByNum(button%, WM_LBUTTONUP, 0, 0)

End Sub

Function conelite(word$) As String
made$ = ""
For X = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, X, 1)
    leet$ = ""
    If letter$ = "a" Then leet$ = "@"
    If letter$ = "b" Then leet$ = "þ"
    If letter$ = "c" Then leet$ = "©"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then leet$ = "ë"
    If letter$ = "f" Then leet$ = "ƒ"
    If letter$ = "g" Then leet$ = "g"
    If letter$ = "h" Then leet$ = "h"
    If letter$ = "i" Then leet$ = "ï"
    If letter$ = "j" Then leet$ = "j"
    If letter$ = "k" Then leet$ = "k"
    If letter$ = "l" Then leet$ = "l"
    If letter$ = "m" Then leet$ = "m"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then leet$ = "ø"
    If letter$ = "p" Then leet$ = "p"
    If letter$ = "q" Then leet$ = "q"
    If letter$ = "r" Then leet$ = "®"
    If letter$ = "s" Then leet$ = "$"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then leet$ = "ü"
    If letter$ = "v" Then leet$ = "v"
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "x" Then leet$ = "×"
    If letter$ = "y" Then leet$ = "ý"
    If letter$ = "z" Then leet$ = "z"
    If letter$ = " " Then leet$ = " "
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "1" Then leet$ = "1"
    If letter$ = "2" Then leet$ = "2"
    If letter$ = "3" Then leet$ = "3"
    If letter$ = "4" Then leet$ = "4"
    If letter$ = "5" Then leet$ = "5"
    If letter$ = "6" Then leet$ = "6"
    If letter$ = "7" Then leet$ = "7"
    If letter$ = "8" Then leet$ = "8"
    If letter$ = "9" Then leet$ = "9"
    If letter$ = "0" Then leet$ = "0"
    If letter$ = "A" Then leet$ = "Å"
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "F" Then leet$ = "ƒ"
    If letter$ = "G" Then leet$ = "G"
    If letter$ = "H" Then leet$ = "/-/"
    If letter$ = "I" Then leet$ = "‡"
    If letter$ = "J" Then leet$ = "J"
    If letter$ = "K" Then leet$ = "]<"
    If letter$ = "L" Then leet$ = "£"
    If letter$ = "M" Then leet$ = "/\/\"
    If letter$ = "N" Then leet$ = "/\/"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "P" Then leet$ = "¶"
    If letter$ = "Q" Then leet$ = "`Q"
    If letter$ = "R" Then leet$ = "•R•"
    If letter$ = "S" Then leet$ = "§"
    If letter$ = "T" Then leet$ = "†"
    If letter$ = "U" Then leet$ = "Ü"
    If letter$ = "V" Then leet$ = "\/"
    If letter$ = "W" Then leet$ = "\\'"
    If letter$ = "X" Then leet$ = "><"
    If letter$ = "Y" Then leet$ = "¥"
    If letter$ = "Z" Then leet$ = "Z"
    If letter$ = "~" Then leet$ = "~"
    If letter$ = "`" Then leet$ = "`"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "@" Then leet$ = "ä"
    If letter$ = "#" Then leet$ = "#"
    If letter$ = "$" Then leet$ = "$"
    If letter$ = "%" Then leet$ = "%"
    If letter$ = "^" Then leet$ = "^"
    If letter$ = "&" Then leet$ = "&"
    If letter$ = "*" Then leet$ = "™"
    If letter$ = "(" Then leet$ = "("
    If letter$ = ")" Then leet$ = ")"
    If letter$ = "-" Then leet$ = "-"
    If letter$ = "_" Then leet$ = "_"
    If letter$ = "+" Then leet$ = "+"
    If letter$ = "=" Then leet$ = "="
    If letter$ = "[" Then leet$ = "["
    If letter$ = "]" Then leet$ = "]"
    If letter$ = "{" Then leet$ = "{"
    If letter$ = "}" Then leet$ = "}"
    If letter$ = ":" Then leet$ = ":"
    If letter$ = ";" Then leet$ = ";"
    If letter$ = "'" Then leet$ = "'"
    If letter$ = "," Then leet$ = ","
    If letter$ = "." Then leet$ = "."
    If letter$ = "<" Then leet$ = "<"
    If letter$ = ">" Then leet$ = ">"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    made$ = made$ & leet$
Next
conelite = made$
End Function

Function conleet(message As String) As String
        msgln = Len(message$)
        leet$ = message$
        message$ = ""
        lo = 1
        up = True
        For X = 1 To msgln
            letter$ = Mid$(leet$, lo, 1)
            If up = True Then
                    If UCase$(letter$) = "I" Then
                        message$ = message$ & "i"
                    ElseIf UCase$(letter$) = "O" Then
                        message$ = message$ & "0"
                    ElseIf UCase$(letter$) = "E" Then
                        message$ = message$ & "3"
                    ElseIf UCase$(letter$) = "A" Then
                        message$ = message$ & "4"
                    Else:
                        message$ = message$ & UCase$(letter$)
                    End If
                up = False
                ElseIf up = False Then
                If UCase$(letter$) = "L" Then
                        message$ = message$ & "L"
                    ElseIf UCase$(letter$) = "O" Then
                        message$ = message$ & "0"
                    ElseIf UCase$(letter$) = "E" Then
                        message$ = message$ & "3"
                    ElseIf UCase$(letter$) = "A" Then
                        message$ = message$ & "4"
                    Else:
                        message$ = message$ & LCase$(letter$)
                    End If
                up = True
            End If
            lo = lo + 1
        Next
conleet = message$
End Function

Sub CountMail()
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
NER% = FindChildByTitle(MDI%, "New Mail")
If NER% <> 0 Then GoTo J
TOO% = FindChildByClass(Aol%, "AOL Toolbar")
NEE% = FindChildByClass(TOO%, "_AOL_Icon")
D = SendMessageByNum(NEE%, WM_LBUTTONDOWN, 0, 0)
D = SendMessageByNum(NEE%, WM_LBUTTONUP, 0, 0)
Do
c% = DoEvents()
NER% = FindChildByTitle(MDI%, "New Mail")
Loop Until NER% <> 0
TRE% = FindChildByClass(NER%, "_AOL_Tree")
Do
U = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
Timeout (2)
f = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
Loop Until U = f
J:
TRE% = FindChildByClass(NER%, "_AOL_Tree")
f = SendMessageByNum(TRE%, LB_GETCOUNT, 0, 0)
MsgBox "You have " & f & " pieces of mail!", 64, "Mail Counter"

End Sub

Function Elit(thestring)
Elitee = thestring
Final = ""
For i = 1 To Len(Elitee)
    FF = Asc(Mid$(Elitee, i, 1))
    If FF = 32 Then
        Final = Final & Chr$(160)
    End If
    If FF = 33 Then
        Final = Final & Chr$(33)
    End If
    If FF = 34 Then
        Final = Final & Chr$(147)
    End If
    If FF = 35 Then
        Final = Final & Chr$(35)
    End If
    If FF = 36 Then
        Final = Final & Chr$(36)
    End If
    If FF = 37 Then
        Final = Final & Chr$(137)
    End If
    If FF = 38 Then
        Final = Final & Chr$(38)
    End If
    If FF = 39 Then
        Final = Final & Chr$(145)
    End If
    If FF = 40 Then
        Final = Final & Chr$(40)
    End If
    If FF = 41 Then
        Final = Final & Chr$(41)
    End If
    If FF = 42 Then
        Final = Final & Chr$(42)
    End If
    If FF = 43 Then
        Final = Final & Chr$(43)
    End If
    If FF = 44 Then
        Final = Final & Chr$(44)
    End If
    If FF = 45 Then
        Final = Final & Chr$(45)
    End If
    If FF = 46 Then
        Final = Final & Chr$(46)
    End If
    If FF = 47 Then
        Final = Final & Chr$(47)
    End If
    If FF = 48 Then
        Final = Final & Chr$(48)
    End If
    If FF = 49 Then
        Final = Final & Chr$(49)
    End If
    If FF = 50 Then
        Final = Final & Chr$(50)
    End If
    If FF = 51 Then
        Final = Final & Chr$(51)
    End If
    If FF = 52 Then
        Final = Final & Chr$(52)
    End If
    If FF = 53 Then
        Final = Final & Chr$(53)
    End If
    If FF = 54 Then
        Final = Final & Chr$(54)
    End If
    If FF = 55 Then
        Final = Final & Chr$(55)
    End If
    If FF = 56 Then
        Final = Final & Chr$(56)
    End If
    If FF = 57 Then
        Final = Final & Chr$(57)
    End If
    If FF = 58 Then
        Final = Final & Chr$(58)
    End If
    If FF = 59 Then
        Final = Final & Chr$(59)
    End If
    If FF = 60 Then
        Final = Final & Chr$(60)
    End If
    If FF = 61 Then
        Final = Final & Chr$(61)
    End If
    If FF = 62 Then
        Final = Final & Chr$(62)
    End If
    If FF = 63 Then
        Final = Final & Chr$(191)
    End If
    If FF = 64 Then
        Final = Final & Chr$(64)
    End If
    If FF = 65 Then
        Final = Final & Chr$(197)
    End If
    If FF = 66 Then
        Final = Final & Chr$(223)
    End If
    If FF = 67 Then
        Final = Final & Chr$(199)
    End If
    If FF = 68 Then
        Final = Final & Chr$(208)
    End If
    If FF = 69 Then
        Final = Final & Chr$(203)
    End If
    If FF = 70 Then
        Final = Final & Chr$(70)
    End If
    If FF = 71 Then
        Final = Final & Chr$(71)
    End If
    If FF = 72 Then
        Final = Final & "|-|"
    End If
    If FF = 73 Then
        Final = Final & Chr$(204)
    End If
    If FF = 74 Then
        Final = Final & "J"
    End If
    If FF = 75 Then
        Final = Final & "\<"
    End If
    If FF = 76 Then
        Final = Final & Chr$(163)
    End If
    If FF = 77 Then
        Final = Final & "]V["
    End If
    If FF = 78 Then
        Final = Final & Chr$(209)
    End If
    If FF = 79 Then
        Final = Final & Chr$(212)
    End If
    If FF = 80 Then
        Final = Final & "P"
    End If
    If FF = 81 Then
        Final = Final & "Q"
    End If
    If FF = 82 Then
        Final = Final & "I2"
    End If
    If FF = 83 Then
        Final = Final & Chr$(167)
    End If
    If FF = 84 Then
        Final = Final & Chr$(134)
    End If
    If FF = 85 Then
        Final = Final & Chr$(220)
    End If
    If FF = 86 Then
        Final = Final & "V"
    End If
    If FF = 87 Then
        Final = Final & "\\'"
    End If
    If FF = 88 Then
        Final = Final & "><"
    End If
    If FF = 89 Then
        Final = Final & Chr$(165)
    End If
    If FF = 90 Then
        Final = Final & "Z"
    End If
    If FF = 91 Then
        Final = Final & "["
    End If
    If FF = 92 Then
        Final = Final & "\"
    End If
    If FF = 93 Then
        Final = Final & "]"
    End If
    If FF = 94 Then
        Final = Final & Chr$(136)
    End If
    If FF = 95 Then
        Final = Final & "_"
    End If
    If FF = 96 Then
        Final = Final & "`"
    End If
    If FF = 97 Then
        Final = Final & Chr$(229)
    End If
    If FF = 98 Then
        Final = Final & Chr$(223)
    End If
    If FF = 99 Then
        Final = Final & Chr$(231)
    End If
    If FF = 100 Then
        Final = Final & Chr$(208)
    End If
    If FF = 101 Then
        Final = Final & Chr$(234)
    End If
    If FF = 102 Then
        Final = Final & Chr$(131)
    End If
    If FF = 103 Then
        Final = Final & "g"
    End If
    If FF = 104 Then
        Final = Final & "h"
    End If
    If FF = 105 Then
        Final = Final & Chr$(161)
    End If
    If FF = 106 Then
        Final = Final & ",j"
    End If
    If FF = 107 Then
        Final = Final & "k"
    End If
    If FF = 108 Then
        Final = Final & Chr$(163)
    End If
    If FF = 109 Then
        Final = Final & "/\/)"
    End If
    If FF = 110 Then
        Final = Final & Chr$(241)
    End If
    If FF = 111 Then
        Final = Final & Chr$(240)
    End If
    If FF = 112 Then
        Final = Final & "p"
    End If
    If FF = 113 Then
        Final = Final & "q"
    End If
    If FF = 114 Then
        Final = Final & Chr$(174)
    End If
    If FF = 115 Then
        Final = Final & Chr$(167)
    End If
    If FF = 116 Then
        Final = Final & Chr$(134)
    End If
    If FF = 117 Then
        Final = Final & Chr$(251)
    End If
    If FF = 118 Then
        Final = Final & "v"
    End If
    If FF = 119 Then
        Final = Final & "w"
    End If
    If FF = 120 Then
        Final = Final & Chr$(215)
    End If
    If FF = 121 Then
        Final = Final & Chr$(255)
    End If
    If FF = 122 Then
        Final = Final & "z"
    End If
    If FF = 123 Then
        Final = Final & Chr$(123)
    End If
    If FF = 124 Then
        Final = Final & Chr$(124)
    End If
    If FF = 125 Then
        Final = Final & Chr$(125)
    End If
    If FF = 126 Then
        Final = Final & Chr$(126)
    End If
    If FF = 127 Then
        Final = Final & Chr$(127)
    End If
    If FF = 128 Then
        Final = Final & Chr$(128)
    End If
    If FF = 129 Then
        Final = Final & Chr$(129)
    End If
    If FF = 130 Then
        Final = Final & Chr$(130)
    End If
    If FF = 131 Then
        Final = Final & Chr$(131)
    End If
    If FF = 132 Then
        Final = Final & Chr$(132)
    End If
    If FF = 133 Then
        Final = Final & Chr$(133)
    End If
    If FF = 134 Then
        Final = Final & Chr$(134)
    End If
    If FF = 135 Then
        Final = Final & Chr$(135)
    End If
    If FF = 136 Then
        Final = Final & Chr$(136)
    End If
    If FF = 137 Then
        Final = Final & Chr$(137)
    End If
    If FF = 138 Then
        Final = Final & Chr$(138)
    End If
    If FF = 139 Then
        Final = Final & Chr$(139)
    End If
    If FF = 140 Then
        Final = Final & Chr$(140)
    End If
    If FF = 141 Then
        Final = Final & Chr$(141)
    End If
    If FF = 142 Then
        Final = Final & Chr$(142)
    End If
    If FF = 143 Then
        Final = Final & Chr$(143)
    End If
    If FF = 144 Then
        Final = Final & Chr$(144)
    End If
    If FF = 145 Then
        Final = Final & Chr$(145)
    End If
    If FF = 146 Then
        Final = Final & Chr$(146)
    End If
    If FF = 147 Then
        Final = Final & Chr$(147)
    End If
    If FF = 148 Then
        Final = Final & Chr$(148)
    End If
    If FF = 149 Then
        Final = Final & Chr$(149)
    End If
    If FF = 150 Then
        Final = Final & Chr$(150)
    End If
    If FF = 151 Then
        Final = Final & Chr$(151)
    End If
    If FF = 152 Then
        Final = Final & Chr$(152)
    End If
    If FF = 153 Then
        Final = Final & Chr$(153)
    End If
    If FF = 154 Then
        Final = Final & Chr$(154)
    End If
    If FF = 155 Then
        Final = Final & Chr$(155)
    End If
    If FF = 156 Then
        Final = Final & Chr$(156)
    End If
    If FF = 157 Then
        Final = Final & Chr$(157)
    End If
    If FF = 158 Then
        Final = Final & Chr$(158)
    End If
    If FF = 159 Then
        Final = Final & Chr$(159)
    End If
    If FF = 160 Then
        Final = Final & Chr$(160)
    End If
    If FF = 161 Then
        Final = Final & Chr$(161)
    End If
    If FF = 162 Then
        Final = Final & Chr$(162)
    End If
    If FF = 163 Then
        Final = Final & Chr$(163)
    End If
    If FF = 164 Then
        Final = Final & Chr$(164)
    End If
    If FF = 165 Then
        Final = Final & Chr$(165)
    End If
    If FF = 166 Then
        Final = Final & Chr$(166)
    End If
    If FF = 167 Then
        Final = Final & Chr$(167)
    End If
    If FF = 168 Then
        Final = Final & Chr$(168)
    End If
    If FF = 169 Then
        Final = Final & Chr$(169)
    End If
    If FF = 170 Then
        Final = Final & Chr$(170)
    End If
    If FF = 171 Then
        Final = Final & Chr$(171)
    End If
    If FF = 172 Then
        Final = Final & Chr$(172)
    End If
    If FF = 173 Then
        Final = Final & Chr$(173)
    End If
    If FF = 174 Then
        Final = Final & Chr$(174)
    End If
    If FF = 175 Then
        Final = Final & Chr$(175)
    End If
    If FF = 176 Then
        Final = Final & Chr$(176)
    End If
    If FF = 177 Then
        Final = Final & Chr$(177)
    End If
    If FF = 178 Then
        Final = Final & Chr$(178)
    End If
    If FF = 179 Then
        Final = Final & Chr$(179)
    End If
    If FF = 180 Then
        Final = Final & Chr$(180)
    End If
    If FF = 181 Then
        Final = Final & Chr$(181)
    End If
    If FF = 182 Then
        Final = Final & Chr$(182)
    End If
    If FF = 183 Then
        Final = Final & Chr$(183)
    End If
    If FF = 184 Then
        Final = Final & Chr$(184)
    End If
    If FF = 185 Then
        Final = Final & Chr$(185)
    End If
    If FF = 186 Then
        Final = Final & Chr$(186)
    End If
    If FF = 187 Then
        Final = Final & Chr$(187)
    End If
    If FF = 188 Then
        Final = Final & Chr$(188)
    End If
    If FF = 189 Then
        Final = Final & Chr$(189)
    End If
    If FF = 190 Then
        Final = Final & Chr$(190)
    End If
    If FF = 191 Then
        Final = Final & Chr$(191)
    End If
    If FF = 192 Then
        Final = Final & Chr$(192)
    End If
    If FF = 193 Then
        Final = Final & Chr$(193)
    End If
    If FF = 194 Then
        Final = Final & Chr$(194)
    End If
    If FF = 195 Then
        Final = Final & Chr$(195)
    End If
    If FF = 196 Then
        Final = Final & Chr$(196)
    End If
    If FF = 197 Then
        Final = Final & Chr$(197)
    End If
    If FF = 198 Then
        Final = Final & Chr$(198)
    End If
    If FF = 199 Then
        Final = Final & Chr$(199)
    End If
    If FF = 200 Then
        Final = Final & Chr$(200)
    End If
    If FF = 201 Then
        Final = Final & Chr$(201)
    End If
    If FF = 202 Then
        Final = Final & Chr$(202)
    End If
    If FF = 203 Then
        Final = Final & Chr$(203)
    End If
    If FF = 204 Then
        Final = Final & Chr$(204)
    End If
    If FF = 205 Then
        Final = Final & Chr$(205)
    End If
    If FF = 206 Then
        Final = Final & Chr$(206)
    End If
    If FF = 207 Then
        Final = Final & Chr$(207)
    End If
    If FF = 208 Then
        Final = Final & Chr$(208)
    End If
    If FF = 209 Then
        Final = Final & Chr$(209)
    End If
    If FF = 210 Then
        Final = Final & Chr$(210)
    End If
    If FF = 211 Then
        Final = Final & Chr$(211)
    End If
    If FF = 212 Then
        Final = Final & Chr$(212)
    End If
    If FF = 213 Then
        Final = Final & Chr$(213)
    End If
    If FF = 214 Then
        Final = Final & Chr$(214)
    End If
    If FF = 215 Then
        Final = Final & Chr$(215)
    End If
    If FF = 216 Then
        Final = Final & Chr$(216)
    End If
    If FF = 217 Then
        Final = Final & Chr$(217)
    End If
    If FF = 218 Then
        Final = Final & Chr$(218)
    End If
    If FF = 219 Then
        Final = Final & Chr$(219)
    End If
    If FF = 220 Then
        Final = Final & Chr$(220)
    End If
    If FF = 221 Then
        Final = Final & Chr$(221)
    End If
    If FF = 222 Then
        Final = Final & Chr$(222)
    End If
    If FF = 223 Then
        Final = Final & Chr$(223)
    End If
    If FF = 224 Then
        Final = Final & Chr$(224)
    End If
    If FF = 225 Then
        Final = Final & Chr$(225)
    End If
    If FF = 226 Then
        Final = Final & Chr$(226)
    End If
    If FF = 227 Then
        Final = Final & Chr$(227)
    End If
    If FF = 228 Then
        Final = Final & Chr$(228)
    End If
    If FF = 229 Then
        Final = Final & Chr$(229)
    End If
    If FF = 230 Then
        Final = Final & Chr$(230)
    End If
    If FF = 231 Then
        Final = Final & Chr$(231)
    End If
    If FF = 232 Then
        Final = Final & Chr$(232)
    End If
    If FF = 233 Then
        Final = Final & Chr$(233)
    End If
    If FF = 234 Then
        Final = Final & Chr$(234)
    End If
    If FF = 235 Then
        Final = Final & Chr$(235)
    End If
    If FF = 236 Then
        Final = Final & Chr$(236)
    End If
    If FF = 237 Then
        Final = Final & Chr$(237)
    End If
    If FF = 238 Then
        Final = Final & Chr$(238)
    End If
    If FF = 239 Then
        Final = Final & Chr$(239)
    End If
    If FF = 240 Then
        Final = Final & Chr$(240)
    End If
    If FF = 241 Then
        Final = Final & Chr$(241)
    End If
    If FF = 242 Then
        Final = Final & Chr$(242)
    End If
    If FF = 243 Then
        Final = Final & Chr$(243)
    End If
    If FF = 244 Then
        Final = Final & Chr$(244)
    End If
    If FF = 245 Then
        Final = Final & Chr$(245)
    End If
    If FF = 246 Then
        Final = Final & Chr$(246)
    End If
    If FF = 247 Then
        Final = Final & Chr$(247)
    End If
    If FF = 248 Then
        Final = Final & Chr$(248)
    End If
    If FF = 249 Then
        Final = Final & Chr$(249)
    End If
    If FF = 250 Then
        Final = Final & Chr$(250)
    End If
    If FF = 251 Then
        Final = Final & Chr$(251)
    End If
    If FF = 252 Then
        Final = Final & Chr$(252)
    End If
    If FF = 253 Then
        Final = Final & Chr$(253)
    End If
    If FF = 254 Then
        Final = Final & Chr$(254)
    End If
    If FF = 255 Then
        Final = Final & Chr$(255)
    End If
Next i
Elit = Final
End Function

Sub enter(edt%)
X = SendMessageByNum(edt%, WM_CHAR, 13, 0)
End Sub

Sub exitprog()
ext = MsgBox("exit?", 68)
If ext = 6 Then
    Aol% = FindWindow("AOL Frame25", 0&)
    End
End If
End Sub

Function Extention(exe$, ext$)
If InStr(exe$, ".") Then
    Where = InStr(exe$, ".")
    filename$ = Mid$(exe$, 1, Where - 1)
    FileType$ = Mid$(exe$, Where + 2)
    If FileType$ <> ext$ Then FileType$ = ext$
    exe$ = filename$ + "." + FileType$
ElseIf InStr(exe$, ".") = False Then
    exe$ = exe$ + "." + ext$
End If
End Function

Function FileExists(filename As String) As Integer
Dim i As Integer
On Error Resume Next
i = Len(Dir$(filename))
If Err Or i = 0 Then
    FileExists = False
    FileExists = 0
Else
    FileExists = True
    FileExists = 1
End If

End Function

Function findchatroom() As Integer
Aol% = FindWindow("AOL Frame25", 0&)
lstbox% = FindChildByClass(Aol%, "_AOL_ListBox")
Parent% = GetParent(lstbox%)
view% = FindChildByClass(Parent%, "_AOL_View")
edi% = FindChildByClass(Parent%, "_AOL_Edit")
If edi% <> 0 Then
chatrm% = GetParent(view%)
findchatroom = chatrm%
End If

End Function

Function findmail()
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
Child% = FindChildByClass(MDI%, "AOL Child")
REP% = FindChildByTitle(Child%, "Reply")
SNO% = FindChildByTitle(Child%, "Send Now")
If REP% = 0 And SNO% = 0 Then
    Do: DoEvents
        Child% = getwindow(Child%, GW_HWNDNEXT)
        REP% = FindChildByTitle(Child%, "Reply")
        SNO% = FindChildByTitle(Child%, "Send Now")
    Loop Until REP% <> 0 Or SNO% <> 0 Or Child% = 0
    If Child% = 0 Then
        findmail = 0
        Exit Function
    End If
End If
findmail = Child%

End Function

Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function

Function Generate() As String
Randomize
    Phrases = 7
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = "Hello, "
      If phrase = 2 Then T$ = "Good evening, "
      If phrase = 3 Then T$ = "Hello, And welcome to America Online. "
      If phrase = 4 Then T$ = "Welcome to America Online. "
      If phrase = 5 Then T$ = "Excuse me, "
      If phrase = 6 Then T$ = "Dear User, "
      If phrase = 7 Then T$ = "What's Up! "
    Phrases = 7
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & "I am with America Online Billing "
      If phrase = 2 Then T$ = T$ & "I am with OTC (Online Techincal Consultants) "
      If phrase = 3 Then T$ = T$ & "I am with WWWC (World Wide Web Consultants) "
      If phrase = 4 Then T$ = T$ & "I am with AOL Techincal Staff "
      If phrase = 5 Then T$ = T$ & "I am with AOL System Security "
      If phrase = 6 Then T$ = T$ & "I am with AOL Resource Department "
      If phrase = 7 Then T$ = T$ & "I am with the AOL Community Action Team "
    Phrases = 7
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & "and due to Technical Failures "
      If phrase = 2 Then T$ = T$ & "and due to Billing errors "
      If phrase = 3 Then T$ = T$ & "and due to Line Noise "
      If phrase = 4 Then T$ = T$ & "and due to A virus in our database "
      If phrase = 5 Then T$ = T$ & "and due to Massive data flow in our database "
      If phrase = 6 Then T$ = T$ & "and due to data corruption "
      If phrase = 7 Then T$ = T$ & "and due to hackers by-passing out systems "

    Phrases = 5
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & "we seem to have lost your password. "
      If phrase = 2 Then T$ = T$ & "we seem to have lost your account information. "
      If phrase = 3 Then T$ = T$ & "we seem to have failed to recieve your logon password. "
      If phrase = 4 Then T$ = T$ & "we seem to have failed to recieve your account information. "
      If phrase = 5 Then T$ = T$ & "we have lost your Credit Card information. "
    Phrases = 4
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & "To correct this situation please click respond and enter "
      If phrase = 2 Then T$ = T$ & "Please click respond and enter your "
      If phrase = 3 Then T$ = T$ & "Please help us by entering your "
      If phrase = 4 Then T$ = T$ & "Please respond with your "
    Phrases = 3
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & "Your password information."
      If phrase = 2 Then T$ = T$ & "Your current account password."
      If phrase = 3 Then T$ = T$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    Phrases = 4
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & " Please respond within 2 minutes too keep this account active."
      If phrase = 2 Then T$ = T$ & " It is very important that you respond immediately."
      If phrase = 3 Then T$ = T$ & " Please respond as soon as possible "
      If phrase = 4 Then T$ = T$ & " Incooperation may lead to termination of your account."
    Phrases = 4
    phrase = Int(Rnd * Phrases + 1)
      If phrase = 1 Then T$ = T$ & " Thank you for using America Online."
      If phrase = 2 Then T$ = T$ & " Thank you for your time."
      If phrase = 3 Then T$ = T$ & " Thank you for your cooperation."
      If phrase = 4 Then T$ = T$ & " Thank you for your help and enjoy the service!"
Generate = T$
End Function

Function GetchatText() As String
view% = ChatView()
If view% = 0 Then Exit Function
i = SendMessageByNum(view%, WM_GETTEXTLENGTH, 0, 0)
ChatText$ = Space(i)
 
ChatText$ = Trim$(ChatText$)
For X = Len(ChatText$) To 1 Step -1
    letter$ = Mid$(ChatText$, X)
    letter$ = Mid$(letter$, 1, 1)
    If letter$ = Chr$(13) Then
        ChatText$ = Mid$(ChatText$, X, Len(ChatText$))
        Exit For
    End If
Next
GetchatText = ChatText$
End Function

Function GetFromINI(SectionHeader$, VarName$, filename$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
  'Get Requested Information
   GetFromINI = Left(RetStr, GetPrivateProfileString(SectionHeader$, ByVal VarName$, "", RetStr, Len(RetStr), filename$))
End Function

Function getroomname() As String
On Error Resume Next
Chat% = findchatroom()
X = GetWindowTextLength(Chat%)
Title$ = Space(X + 1)
X = GetWindowText(Chat%, Title$, X + 1)
Title$ = FixAPIString(Title$)
getroomname = Title$
End Function

Function GetScreenName()
Dim MDI%
Dim Aol%
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
Dim Welcome%
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
Dim yourname As String * 255
Dim yourname2$
X = GetWindowText(Welcome%, yourname, 255)
yourname = LTrim(RTrim(Trim(yourname)))
yourname = Right$(yourname, Len(yourname) - InStr(yourname, ", "))
yourname2 = Left$(yourname, 10)
yourname2 = Right$(yourname, Len(yourname) - 1)
If InStr(yourname2, "!") <> 0 Then yourname2 = Left$(yourname2, InStr(yourname2, "!") - 1)
GetScreenName = yourname2$
End Function

Function GetSN() As String
On Error Resume Next
Aol = FindWindow("aol frame25", 0&)
Wel = FindChildByTitle(Aol, "Welcome,")
If Wel = 0 Then GetSN = "You Are Not Online": Exit Function
namelen = SendMessage(Wel, WM_GETTEXTLENGTH, 0, 0)
buffer$ = String$(namelen, 0)
X = SendMessageByString(Wel, WM_GETTEXT, namelen, buffer$)
a = InStr(buffer$, ",")
SN$ = Mid$(buffer$, a + 2, (Len(buffer$) - (a + 1)))
SN$ = (SN$)
GetSN = SN$
End Function

Function GetTextFromRICHCNTL(hWindow As Integer)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
Msg$ = String$(255, 0)
XX% = SendMessageLong(hWindow, WM_ProGGer, 254, Msg$)
Msg$ = Trim$(Msg$)
GetTextFromRICHCNTL = Msg$

End Function

Function GetWinIni(Section, key)
Dim RetVal As String, AppName As String, worked As Integer
    RetVal = String$(255, 0)
    worked = GetProfileString(Section, key, "", RetVal, Len(RetVal))
    If worked = 0 Then
        GetWinIni = "unknown"
    Else
        GetWinIni = Left(RetVal, worked)
    End If

End Function

Sub hit_menu(Mnu_str$)
Aol% = FindWindow("AOL Frame25", 0&)
mnu% = GetMenu(Aol%)
MNU_Count% = GetMenuItemCount(mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Buff$ = Space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Buff$, 50, MF_BYPOSITION)  'As Integer
        Buff$ = Trim$(Buff$): Buff$ = Left(Buff$, Len(Buff$) - 1)
        If Buff$ = "" Then Buff$ = " -"
        If InStr(Buff$, Mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = SendMessageByNum(Aol%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Sub

Sub hit_menu_by_string(Mnu_str$)
Aol% = FindWindow("AOL Frame25", 0&)
mnu% = GetMenu(Aol%)
MNU_Count% = GetMenuItemCount(mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Buff$ = Space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Buff$, 50, MF_BYPOSITION)  'As Integer
        Buff$ = Trim$(Buff$): Buff$ = Left(Buff$, Len(Buff$) - 1)
        If Buff$ = "" Then Buff$ = " -"
        If InStr(Buff$, Mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = SendMessageByNum(Aol%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Sub

Sub Keyword(key$)
Aol% = FindWindow("AOL Frame25", 0&)
runmeu "AOL Frame25", "&Go To", "Keyword..."
Do
DoEvents
KeyWindow% = FindChildByTitle(Aol%, "Keyword")
DoEvents
Loop Until KeyWindow% <> 0
DoEvents
AolEdit% = FindChildByClass(KeyWindow%, "_AOL_Edit")
Call SetEdit(AolEdit%, key$)
enter (AolEdit%)
End Sub

Sub KillWait()
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
If aolversion() = 25 Then
    Aol% = FindWindow("AOL Frame25", 0&)
    MDI% = FindChildByClass(Aol%, "MDIClient")
     
    Do
        c% = DoEvents()
        EXI% = FindWindow("_AOL_Modal", "America Online")
    Loop Until EXI% <> 0
    c% = DoEvents()
    D = SendMessageByNum(EXI%, &H10, 0, 0)
End If
If aolversion() = 30 Then
    Aol% = FindWindow("AOL Frame25", 0&)
    MDI% = FindChildByClass(Aol%, "MDIClient")
    hit_menu "Preferences"
    Do
        c% = DoEvents()
        PRE% = FindChildByTitle(MDI%, "Preferences")
    Loop Until PRE% <> 0
    ICC% = FindChildByClass(PRE%, "_AOL_Icon")
    Click (ICC%)
    Do
        c% = DoEvents()
        MOO% = FindWindow("_AOL_Modal", "Chat Preferences")
    Loop Until MOO% <> 0
    D = SendMessageByNum(MOO%, &H10, 0, 0)
    D = SendMessageByNum(PRE%, &H10, 0, 0)
End If
c% = DoEvents()
End Sub

Sub KuRuPtadd(List As ListBox)
End Sub

Sub KuRuPtaddlist(TheList As ListBox)
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
Chatroom% = findchatroom()
If Chatroom% = 0 Then
    MsgBox "You must be in a chat room to use this feature.", 22, "No Chat Room!"
    Exit Sub
End If
AolList% = FindChildByClass(Chatroom%, "_AOL_Listbox")
yoursn = GetScreenName()
LBSize = SendMessage(AolList%, LB_GETCOUNT, 0, 0)
For GetNames% = 0 To LBSize - 1
    storage$ = String$(255, 0)
    operation = AOLGetList(GetNames%, storage$)
    storage$ = Left$(storage$, InStr(storage$, Chr$(0)) - 1)
    If UCase$(storage$) = UCase$(yoursn) Then GoTo ProGGer
    For BEBE = 0 To TheList.ListCount - 1
        If UCase$(storage$) = UCase$(TheList.List(BEBE)) Then
            GoTo ProGGer
        End If
    Next BEBE
    TheList.AddItem storage$
ProGGer:
Next GetNames%

End Sub

Sub Kuruptim(who, TBox)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
hit_menu "Send an Instant Message"
Do
    c% = DoEvents()
    IMM% = FindChildByTitle(MDI%, "Send Instant Message")
    TOO% = FindChildByTitle(IMM%, "To:")
Loop Until IMM% <> 0 And TOO% <> 0
edi% = FindChildByClass(IMM%, "_AOL_Edit")
X = SendMessageByString(edi%, WM_SETTEXT, 0&, who)
RIC% = FindChildByClass(IMM%, "RICHCNTL")
If RIC% <> 0 Then
    X = SendMessageByString(RIC%, WM_SETTEXT, 0&, TBox)
    SEND% = FindChildByClass(IMM%, "_AOL_Icon")
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
End If
If RIC% = 0 Then
    edi% = getwindow(edi%, GW_HWNDNEXT)
    X = SendMessageByString(edi%, WM_SETTEXT, 0&, TBox)
    SEND% = FindChildByTitle(IMM%, "Send")
End If

D = SendMessageByNum(SEND%, WM_LBUTTONDOWN, 0, 0)
D = SendMessageByNum(SEND%, WM_LBUTTONUP, 0, 0)

    c% = DoEvents()
    OFF% = FindWindow("#32770", "America Online")
    IMM% = FindChildByTitle(MDI%, "Send Instant Message")

End Sub

Sub KuRuPtList(List As ListBox, who$)
For X = 0 To List.ListCount - 1
    If UCase$(List.List(X)) Like UCase$("*" + who$ + "*") Then Exit Sub
Next
If who$ = "" Then Exit Sub
List.AddItem who$
End Sub

Sub KuRuPtRemove(KuRuPt As ListBox)
If KuRuPt = "" Then Exit Sub
KuRuPt.RemoveItem KuRuPt.ListIndex
End Sub

Function kuruptsn(Kur As String, ByVal rupt As String, Lug As String) As String
Dim LorD As String
Dim Rakwon As String
Dim Shocker As String

KuRuPt:
If InStr(UCase(Kur), UCase(rupt)) Then
        LorD$ = Left(Kur, InStr(UCase(Kur), UCase(rupt)) - 1)
        Rakwon$ = Mid(Kur, InStr(UCase(Kur), UCase(rupt)) + Len(rupt), Len(Kur))
        Shocker = LorD & KooL & Rakwon
    Else:
        Shocker = Kur
End If
kuruptsn = Shocker
End Function

Function LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        Where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, Where + Len(look$))
        LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function

Sub lobby()
Aol% = FindWindow("AOL Frame25", 0&)
runmeu "AOL Frame25", "&Go To", "Lobby"
Do
DoEvents
Loop Until findchatroom() <> 0
End Sub

Sub Pause(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Function Percent(Complete As Integer, Total As Integer, TotalOutput As Integer) As Integer
R = Int(Complete / Total * TotalOutput)
Percent = R
End Function

Function r_backward(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function r_elie(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo dustepp2
If crapp% > 0 Then GoTo dustepp2

If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "|V|"
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "º"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$

dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function

Function r_hacke(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
r_hacker = newsent$

End Function

Function r_sam(strin As String)
Let r_same = Trim(strr)
End Function

Function r_space(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
r_spaced = newsent$

End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub remove(List As ListBox)
If List = "" Then Exit Sub
List.RemoveItem List.ListIndex
End Sub

Sub renamehost(SN As String)

Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
End Sub

Function Replace(object$, RepWhat$, RepWit$) As String
Yay:
If InStr(UCase$(object$), UCase$(RepWhat$)) Then
        LeftSide$ = Left$(object$, InStr(UCase$(object$), UCase$(RepWhat$)) - 1)
        RightSide$ = Mid$(object$, InStr(UCase$(object$), UCase$(RepWhat$)) + Len(RepWhat$), Len(object$))
        Finish$ = LeftSide$ & RepWit$ & RightSide$
    Else:
        Finish$ = object$
End If
Replace = Finish$
End Function

Sub resetnew(SN As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
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
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub

Sub Resetsn(SN As String, aoldir As String, Replace As String)
DoEvents
On Error GoTo ThatPlace
TheNumber = Len(SN$)
If TheNumber = 1 Then SN$ = SN$ + "         "
If TheNumber = 2 Then SN$ = SN$ + "        "
If TheNumber = 3 Then SN$ = SN$ + "       "
If TheNumber = 4 Then SN$ = SN$ + "      "
If TheNumber = 5 Then SN$ = SN$ + "     "
If TheNumber = 6 Then SN$ = SN$ + "    "
If TheNumber = 7 Then SN$ = SN$ + "   "
If TheNumber = 8 Then SN$ = SN$ + "  "
If TheNumber = 9 Then SN$ = SN$ + " "
TheNumber = Len(Replace$)
If TheNumber = 1 Then Replace$ = Replace$ + "         "
If TheNumber = 2 Then Replace$ = Replace$ + "        "
If TheNumber = 3 Then Replace$ = Replace$ + "       "
If TheNumber = 4 Then Replace$ = Replace$ + "      "
If TheNumber = 5 Then Replace$ = Replace$ + "     "
If TheNumber = 6 Then Replace$ = Replace$ + "    "
If TheNumber = 7 Then Replace$ = Replace$ + "   "
If TheNumber = 8 Then Replace$ = Replace$ + "  "
If TheNumber = 9 Then Replace$ = Replace$ + " "
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

Sub RoomBust(roomname$)
    
End Sub

Sub RunMenuByString(Application, StringSearch)

MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:

End Sub

Sub runmeu(Main_Prog As String, Top_Position As String, Menu_String As String)
On Error GoTo stp
Dim Top_Position_Num As Integer
Dim buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim Aol As Integer
Dim Menu_Handle As Integer


Top_Position_Num = -1
Aol% = FindWindow(Main_Prog, 0&)
Menu_Handle = GetMenu(Aol%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, &H400)
    Trim_Buffer = (buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, &H400)
    Trim_Buffer = (buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(Aol, &H111, Get_ID%, 0&)
stp:
Resume Next

End Sub

Function scramble(TheWord) As String
On Error GoTo uhoh
Scram$ = TheWord
Do: DoEvents
    DD = InStr(Scram$, " ")
    EE$ = Scram$
    If DD <> 0 Then EE$ = Left$(Scram$, DD - 1)
    If DD <> 0 Then Scram$ = Mid$(Scram$, DD + 1, Len(Scram$))
    For i = 1 To Len(EE$)
        Randomize
        FF = Int(Len(EE$) * Rnd + 1)
        GG$ = GG$ + Mid$(EE$, FF, 1)
        HH$ = Left$(EE$, FF - 1)
        II$ = Mid$(EE$, FF + 1, Len(EE$))
        EE$ = HH$ + II$
    Next i
    GG$ = GG$ & " "
Loop Until DD = 0
uhoh:
scramble = Trim$(GG$)
Exit Function
End Function

Function ScrambleGame(thestring As String)
Dim bytestring As String

thestringcount = Len(thestring$)
If Not Mid(thestring$, thestringcount, 1) = " " Then thestring$ = thestring$ & " "
For Stringe = 1 To Len(thestring$)
characters$ = Mid(thestring$, Stringe, 1)
thestrings$ = thestrings$ & characters$

If characters$ = " " Then
smoked:
DoEvents
For Ensemble = 1 To Len(thestrings$) - 1
Randomize
randomstring = Int((Len(thestrings$) * Rnd) + 1)
If randomstring = Len(thestrings$) Then GoTo already
If bytesread Like "*" & randomstring & "*" Then GoTo already
stringrandom$ = Mid(thestrings$, randomstring, 1)
stringfound$ = stringfound$ & stringrandom$
bytesread = bytesread & randomstring
GoTo really
already:
Ensemble = Ensemble - 1
really:
Next Ensemble
If stringfound$ = thestrings$ Then stringfound$ = "": GoTo smoked
thestrings2$ = thestrings2$ & stringfound$ & " "
stringfound$ = ""
thestrings$ = ""
bytesread = ""
strngfound$ = ""
End If

Next Stringe
ScrambleGame = Mid(thestrings2$, 1, Len(thestring$) - 1)
End Function

Function ScrambleText(thetext)
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
ScrambleText = scrambled$

Exit Function
End Function

Sub sendim(who, TBox)
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
hit_menu "Send an Instant Message"
Do
    C1% = DoEvents()
    IMM% = FindChildByTitle(MDI%, "Send Instant Message")
    TOO% = FindChildByTitle(IMM%, "To:")
Loop Until IMM% <> 0 And TOO% <> 0
edi% = FindChildByClass(IMM%, "_AOL_Edit")
X = SendMessageByString(edi%, WM_SETTEXT, 0&, who)
RIC% = FindChildByClass(IMM%, "RICHCNTL")
If RIC% <> 0 Then
    X = SendMessageByString(RIC%, WM_SETTEXT, 0&, TBox)
    SEND% = FindChildByClass(IMM%, "_AOL_Icon")
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
    SEND% = getwindow(SEND%, GW_HWNDNEXT)
End If
If RIC% = 0 Then
    edi% = getwindow(edi%, GW_HWNDNEXT)
    X = SendMessageByString(edi%, WM_SETTEXT, 0&, TBox)
    SEND% = FindChildByTitle(IMM%, "Send")
End If

D = SendMessageByNum(SEND%, WM_LBUTTONDOWN, 0, 0)
D = SendMessageByNum(SEND%, WM_LBUTTONUP, 0, 0)
Do
    C1% = DoEvents()
    OFF% = FindWindow("#32770", "America Online")
    IMM% = FindChildByTitle(MDI%, "Send Instant Message")
Loop Until OFF% <> 0 Or IMM% = 0
If OFF% <> 0 Then
    OGG% = FindChildByTitle(OFF%, "OK")
    S = SendMessageByNum(OGG%, WM_LBUTTONDOWN, 0, 0)
    S = SendMessageByNum(OGG%, WM_LBUTTONUP, 0, 0)
    D = SendMessageByNum(IMM%, &H10, 0, 0)
    IMSOFF = 1
End If
End Sub

Sub sendmail(ToWho, Subject, thetext)
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")
hit_menu "Compose Mail"
Do
    c% = DoEvents()
    COP% = FindChildByTitle(MDI%, "Compose Mail")
Loop Until COP% <> 0
edi% = FindChildByClass(COP%, "_AOL_Edit")
X = SendMessageByString(edi%, WM_SETTEXT, 0&, ToWho)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
X = SendMessageByString(edi%, WM_SETTEXT, 0&, Subject)
RIC% = FindChildByClass(COP%, "RICHCNTL")
If RIC% <> 0 Then
    X = SendMessageByString(RIC%, WM_SETTEXT, 0&, thetext)
End If
If RIC% = 0 Then
    edi% = getwindow(edi%, GW_HWNDNEXT)
    edi% = getwindow(edi%, GW_HWNDNEXT)
    edi% = getwindow(edi%, GW_HWNDNEXT)
    X = SendMessageByString(edi%, WM_SETTEXT, 0&, thetext)
End If
ICC% = FindChildByClass(COP%, "_AOL_Icon")
Click (ICC%)

End Sub

Sub SENDTEXT(thetext)
Chatroom% = findchatroom()
AolEdit% = FindChildByClass(Chatroom%, "_AOL_Edit")
X = SendMessageByString(AolEdit%, WM_SETTEXT, 0&, thetext)
D = SendMessageByNum(AolEdit%, WM_CHAR, 13, 0)
End Sub

Sub SetEdit(edt%, txt$)
X = SendMessageByString(edt%, WM_SETTEXT, 0, txt$)
End Sub

Function SignedOn()
Aol% = FindWindow("AOL Frame25", 0&)
If Aol% = 0 Then
    SignedOn = 0
    Exit Function
End If
MDI% = FindChildByClass(Aol%, "MDIClient")
Wel% = FindChildByTitle(MDI%, "Welcome, ")
If Wel% = 0 Then
    SignedOn = 0
    Exit Function
End If
COMB% = FindChildByClass(Wel%, "_AOL_Combobox")
If COMB% <> 0 Then
    SignedOn = 0
    Exit Function
End If
SignedOn = 2
End Function

Sub SignOff()
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
Aol% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(Aol%, "MDIClient")

Do
c% = DoEvents()
EXX% = FindChildByTitle(MDI%, "Exit?")
EXI% = FindWindow("_AOL_Modal", "America Online")
Loop Until EXX% <> 0 Or EXI% <> 0
If EXX% <> 0 Then
    ICC% = FindChildByClass(EXX%, "_AOL_Icon")
    ICC% = getwindow(ICC%, GW_HWNDLAST)
    ICC% = getwindow(ICC%, GW_HWNDPREV)
    ICC% = getwindow(ICC%, GW_HWNDPREV)
    D = SendMessageByNum(ICC%, WM_LBUTTONDOWN, 0, 0)
    D = SendMessageByNum(ICC%, WM_LBUTTONUP, 0, 0)
End If
If EXI% <> 0 Then
    YEE% = FindChildByTitle(EXI%, "&Yes")
    D = SendMessageByNum(YEE%, WM_LBUTTONDOWN, 0, 0)
    D = SendMessageByNum(YEE%, WM_LBUTTONUP, 0, 0)
End If

End Sub

Sub stayontop(f As Form)
SetWindowPos f.hWnd, HWND_TOPMOST, -1, -1, -1, -1, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Sub Timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Function trimnull() As String


      
    
        GoTo NullDetect
   

NullDetect:
trimnull = Total$

End Function

Function usersn() As String
    On Error GoTo noton
    Aol% = FindWindow("AOL Frame25", 0&)
    Welcome% = FindChildByTitle(Aol%, "Welcome, ")
    Length = GetWindowTextLength(Welcome%)
    Title$ = Space(Length)
    X = GetWindowText(Welcome%, Title$, Length)
    comma = InStr(Title$, ", ")
    If comma = False Then Exit Function
    rest$ = Mid$(Title$, comma + 2)
    SN$ = Mid$(rest$, 1, Len(rest$) - 1)
    SN$ = Trim(SN$)
    usersn = SN$
noton:
Resume Next
End Function

Sub VBMsg1_WindowMessage(hWindow As Integer, Msg As Integer, wParam As Integer, lParam As Long, RetVal As Long, CallDefProc As Integer)
Dim ChatText$, colon%, SN$, Textsaid$, Kword$, a%, b%, c%, edithWnd%, e%, msgnam%, SEND%
    ChatText$ = aggetstringfromlpstr(lParam)
    colon% = InStr(1, ChatText$, ":")
    SN$ = Mid(ChatText$, 3, (colon% - 3))
    Textsaid$ = Right(ChatText$, (Len(ChatText$) - Len(SN$)) - 4)
    Kword$ = Text1
    If UCase$(Left$(Textsaid$, Len(Kword$))) = UCase$(Kword$) Then
        DoEvents
Call SENDTEXT("I got it")

End If
End Sub

Sub Wait4OK()
    Do
        DoEvents
        full% = FindWindow("#32770", 0&)
    Loop Until full% <> 0
    X = SendMessageByNum(full%, Wm_Close, 0, 0)
End Sub

Sub waitforok()
Do
DoEvents
okw% = FindWindow("#32770", "America Online")
DoEvents
If okw% <> 0 Then Exit Do
Loop
    OKB = FindChildByTitle(okw%, "OK")
    Okd = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub

