Attribute VB_Name = "Module2"

' This GenOziDe basic file was written by GenOziDe and consists of:
' bas files called... trident.bas master1.bas master2.bas master3.bas...
' master32.bas NOFX.bas SK8R.bas Progger.bas Fade.bas and my own personal code.

Global Const SWP_NOACTIVATE = &H10


'Menu Functions
Declare Sub DrawMenuBar Lib "User" Alias "Drawmenubar" (ByVal hWnd As Integer)
Declare Function AOLGetList% Lib "311.dll" (ByVal Index%, ByVal buf$)
Declare Function TrackPopupMenu Lib "User" (ByVal hMenu As Integer, ByVal wflags As Integer, ByVal X As Integer, ByVal y As Integer, ByVal nReserved As Integer, ByVal hWnd As Integer, lpReserved As Any) As Integer
Declare Function GetMenu Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function CreatePopupMenu Lib "User" () As Integer
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wflags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function DeleteMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wflags As Integer) As Integer

'Menu Constants
Global Const MF_POPUP = &H10
Global Const MF_STRING = &H0
Global Const MF_ENABLED = &H0
Global Const MF_SEPARATOR = &H800
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_RIGHTBUTTON = &H2






Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function ExitWindows Lib "User" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer
Global Const LB_GETTEXT = &H189
Global Const WM_GETDLGCODE = &H87
Global Const WM_CTLCOLOR = &H19
Declare Function ShowWindow Lib "User" Alias "showwindow" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Global Const LBN_DBLCLK = 2
Const c003E = 16 ' &H10%
Global Const WM_USER = &H400
Global Const LB_GETSEL = (WM_USER + 8)
Declare Function sndplaysound Lib "MMSystem" Alias "sndPlaySound" (ByVal lpWavName$, ByVal Flags%) As Integer '
Global Const WM_MOVE = &H3
Global Const SWP_NOREPOSITION = &H200
Declare Sub Movewindow Lib "User" (ByVal hWnd As Integer, ByVal X As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Global Const WM_CLOSE = &H10
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const WM_ACTIVATE = &H6

Global Const NUMBOXES = 5
Global Const SAVEFILE = 1, LOADFILE = 2
Global Const REPLACEFILE = 1, READFILE = 2, ADDTOFILE = 3
Global Const RANDOMFILE = 4, BINARYFILE = 5

' Define a data type to hold a record:
' Define global variables to hold the file number and record number
' of the current data file.
' Default file name to show in dialog boxes.
Global Const Err_DeviceUnavailable = 68
Global Const Err_DiskNotReady = 71, Err_FileAlreadyExists = 58
Global Const Err_TooManyFiles = 67, Err_RenameAcrossDisks = 74
Global Const Err_Path_FileAccessError = 75, Err_DeviceIO = 57
Global Const Err_DiskFull = 61, Err_BadFileName = 64
Global Const Err_BadFileNameOrNumber = 52, Err_FileNotFound = 53
Global Const Err_PathDoesNotExist = 76, Err_BadFileMode = 54
Global Const Err_FileAlreadyOpen = 55, Err_InputPastEndOfFile = 62
Global Const MB_EXCLAIM = 48, MB_STOP = 16


Global Const WM_KILLFOCUS = &H8
Global Const WM_SETFOCUS = &H7
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const WM_SIZE = &H5
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOW = 5
Global Const WM_LBUTTONDBLCLK = &H203
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

Type RECT
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type
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
  X As Integer
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type



'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Function findwindow Lib "User" Alias "FindWindow" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer

Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wflags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer




Declare Function extfn2668 Lib "VBWFind.Dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn16A8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn1868 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal X%, ByVal y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Declare Function extfn2630 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1AD0 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer

Declare Function FlashWindow Lib "User" (ByVal hWnd As Integer, ByVal bInvert As Integer) As Integer
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)

Declare Sub SetCursorPos Lib "User" (ByVal X As Integer, ByVal y As Integer)

Declare Function getfreesystemresources Lib "User" (ByVal fusysresource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
'Declare Function sendmessage& Lib "User" (ByVal hWnd%, ByVal wmsg%, ByVal wparam%, ByVal lParam As Any)
Declare Function CreateMenu% Lib "User" ()

Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)

Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function GetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetCursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Global Const BM_SETCHECK = WM_USER + 1
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowtextlength Lib "User" Alias "GetWindowTextLength" (ByVal hWnd As Integer) As Integer
Declare Function GettopWindow Lib "User" Alias "GetTopWindow" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wflags%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)
Declare Function SendMessageLong Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal hMsg As Integer, ByVal wParam As Integer, ByVal lParam As Any) As Long

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function getfreespace Lib "Kernel" (ByVal wflags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal Filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal maxsize As Integer, ByVal Filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer
Declare Function agGetStringFromlpstr$ Lib "APIGuide.Dll" Alias "agGetStringFromLPSTR" (ByVal lpString&)

'Subs and Functions for "GDI"
Declare Sub SetBKColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer

'Subs and Functions for "MMSystem"
Declare Function MciSendString& Lib "MMSystem" (ByVal Cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)

'Subs and Functions for "APIGuide.dll"
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
Declare Function agGetAddressForLPSTR& Lib "APIGuide.Dll" Alias "agGetAddressForObject" (ByVal lpString$)
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
Declare Function agVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal Value&)
Declare Function dwVBSetControlFlags& Lib "APIGuide.Dll" (ctl As Control, ByVal mask&, ByVal Value&)


'Subs and Functions for "VBMsg.Vbx"
Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal s As String) As Long
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
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal win%, ByVal iNum%)

'                   Sound Constants
'                   ---------------


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
Global Const R2_MASKNOTPEN = 3 'DPna
Global Const R2_NOTCOPYPEN = 4 'PN
Global Const R2_MASKPENNOT = 5 'PDna
Global Const R2_NOT = 6 'Dn
Global Const R2_XORPEN = 7 'DPx
Global Const R2_NOTMASKPEN = 8 'DPan
Global Const R2_MASKPEN = 9 'DPa
Global Const R2_NOTXORPEN = 10 'DPxn
Global Const R2_NOP = 11 'D
Global Const R2_MERGENOTPEN = 12 'DPno
Global Const R2_COPYPEN = 13 'P
Global Const R2_MERGEPENNOT = 14 'PDno
Global Const R2_MERGEPEN = 15 'DPo
Global Const R2_WHITE = 16 ' 1

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
Global Const TRANSPARENT = 1
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
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hdc%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function getfocus% Lib "User" ()
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Function GetCurrentTime& Lib "User" ()
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Declare Function destroywindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Sub closewindow Lib "User" (ByVal hWnd As Integer)
Declare Function GetMenuItemCount Lib "User" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function setfocusapi% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function GetMenuItemID Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function GetSubMenu Lib "User" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function test Lib "APIGuide.Dll" Alias "AgGetStringFromLPSTR" (ByVal p1&) As String



Global Const WM_MBUTTONDOWN = &H207
Global Const WM_MBUTTONUP = &H208
Global Const GW_HWNDNEXT = 2
Global Const WM_GETTEXTLENGTH = &HE
Global Const MF_BYPOSITION = &H400
Global Const WM_COMMAND = &H111
Global Const WM_CHAR = &H102
Global Const wm_gettext = &HD
Global Const WM_SETTEXT = &HC
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function getwindow Lib "User" Alias "GetWindow" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function sendmessage Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function SendMessageByNum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function SendMessageByString& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function DeleteBubble% Lib "bubble.dll" (ByVal wnd%)
Declare Function findchildbytitle% Lib "VBWFind.Dll" (ByVal Parent%, ByVal Title$)
Declare Function Findchildbyclass% Lib "VBWFind.Dll" (ByVal Parent%, ByVal Title$)

Declare Function getparent Lib "User" Alias "GetParent" (ByVal hWnd As Integer) As Integer

Sub eightball()
Dim Lst As String
Dim Text As String
Dim AOL As Integer
Dim cht As Integer
Dim txt As String
Dim nws As String
Dim who As String
Dim wht As String
Dim R As Integer
Dim e As Integer
Dim m As Integer
Dim X
Dim y
GeNo:
y = getuser()
AOL = findwindow("AOL Frame25", 0&)
cht = Findchildbyclass(AOL, "_AOL_View")
txt = WinTxt(cht)
If Lst = "" Then Lst = txt
If txt = Lst Then Exit Sub
Lst = txt
nws = GetLast(txt)
who = Mid(nws, 2, InStr(nws, ":") - 2)
wht = Mid(nws, Len(who) + 4, Len(nws) - Len(who))
If LCase(Trim(trm(y))) = LCase(Trim(trm(who))) Then GoTo GeNo
R = getparent(cht)
e = Findchildbyclass(R, "_AOL_Edit")
tixt = RandomNumber(11)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then: tixt = "Definately YES!"
ElseIf tixt = "3" Then: tixt = "Definately No!"
ElseIf tixt = "4" Then: tixt = "Not a FuKin chance"
ElseIf tixt = "5" Then: tixt = "HEEELLLLLLLLLLLLLLL nO"
ElseIf tixt = "6" Then: tixt = "HeLL yeA!"
ElseIf tixt = "7" Then: tixt = "Response HaZey try again."
ElseIf tixt = "8" Then: tixt = "ProbabLee"
ElseIf tixt = "9" Then: tixt = "yep yep"
ElseIf tixt = "10" Then: tixt = "I'm not suRe"
ElseIf tixt = "11" Then: tixt = "AbsolootLee yeZ"
End If
Text = wht$
w = InStr(LCase$(Text), LCase$("/8Ball"))
If w <> 0 Then
sendtext "•––•^v^•{‡ " & who & ", The 8-ball say: " & tixt
timeout 0.5
GoTo GeNo
End If

End Sub

Sub eightballgen(tixt As TextBox)
Randomize
tixt = Int((Val("11") * Rnd) + 1)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then: tixt = "Definately YES!"
ElseIf tixt = "3" Then: tixt = "Definately No!"
ElseIf tixt = "4" Then: tixt = "Not a FuKin chance"
ElseIf tixt = "5" Then: tixt = "HEEELLLLLLLLLLLLLLL nO"
ElseIf tixt = "6" Then: tixt = "HeLL yeA!"
ElseIf tixt = "7" Then: tixt = "Response HaZey try again."
ElseIf tixt = "8" Then: tixt = "ProbabLee"
ElseIf tixt = "9" Then: tixt = "yep yep"
ElseIf tixt = "10" Then: tixt = "I'm not suRe"
ElseIf tixt = "11" Then: tixt = "AbsolootLee yeZ"
End If
End Sub

Sub eMail(Lst As ListBox, Subj$, lst2 As ListBox)

Dim R, X%, C1%, C2%, C3%, C4%, C5%, C6%, C7%, C8%, C9%, C10%, C11%, C12%, C13%, C14%, C15%, C16%, C17%
Call runmenua(3, 1)
X% = WaitForWindow("Compose Mail", "AOL Child")
C1% = getwindow(X%, 5)  '|Send|
C2% = getwindow(C1%, 2) '<Send>
C3% = getwindow(C2%, 2) '|Send Later|
C4% = getwindow(C3%, 2) '<Send Later>
C5% = getwindow(C4%, 2) '|Attach|
C6% = getwindow(C5%, 2) '<Attach>
C7% = getwindow(C6%, 2) '|Address Book|
C8% = getwindow(C7%, 2) '<Address Book>
C9% = getwindow(C8%, 2) '|To:|
C10% = getwindow(C9%, 2) '<To:>
C11% = getwindow(C10%, 2) '|CC:|
C12% = getwindow(C11%, 2) '<CC:>
C13% = getwindow(C12%, 2) '|Subject:|
C14% = getwindow(C13%, 2) '<Subject:>
C15% = Findchildbyclass(X%, "#32770")
C16% = Findchildbyclass(gen%, "Edit")
C17% = Findchildbyclass(X%, "RICHCNTL") '<Message>
Do
Loop Until C10% <> 0
who$ = ""
For i = 0 To Lst.ListCount - 1
who$ = who$ + Lst.List(i) + ", "
y = SendMessageByString(C10%, WM_SETTEXT, 0, who$)
Next i
timeout 1
textset C14%, CStr(Subj$)
whoa$ = ""
For i = 0 To lst2.ListCount - 1
whoa$ = whoa$ + lst2.List(i) & Chr(13)
txt$ = "——§—— G e n O z i D e ——§——"
y = SendMessageByString(Dad%, WM_SETTEXT, 0, whoa$ & Chr(13) & Chr(13) & Chr(13) & txt$)
z = SendMessageByString(rich, WM_SETTEXT, 0, whoa$ & Chr(13) & Chr(13) & Chr(13) & txt$)
Next i
textset C17%, CStr(whoa$)
timeout 0.2

    Click (C2%)


End Sub

Sub emailbomb(numofbomb As TextBox, whotobom, Subj$, mesg$)
Do
Call Mail((whotobom), (Subj$), (mesg$))
numofbomb = Str(Val(numofbomb - 1))
Loop Until numofbomb = 0
AOL% = findwindow("aol Frame25", 0&)
compoze% = findchildbytitle(AOL%, "Compose Mail")
If compoze% <> 0 Then
Do
AOL% = findwindow("aol Frame25", 0&)
compoze% = findchildbytitle(AOL%, "Compose Mail")
X = SendMessageByNum(compoze%, WM_CLOSE, 0, 0)
Loop Until compoze% = 0
End If

End Sub

Function Encrypt2()
AO% = findwindow("AOL Frame25", 0&)
Cmpse% = findchildbytitle%(AO%, "Compose Mail")
EDT% = Findchildbyclass%(Cmpse%, "_AOL_Edit")
a% = getwindow(EDT%, GW_HWNDNEXT): CC% = getwindow(a%, GW_HWNDNEXT)
B% = getwindow(CC%, GW_HWNDNEXT): Subject% = getwindow(B%, GW_HWNDNEXT)
Sub1% = getwindow(Subject%, GW_HWNDNEXT)
'Sub2% = GetWindowtext(SubGW_HWNDNEXT)
Whole$ = ""
For X1 = 1 To 60
 Randomize
   valteger = Int(17 * Rnd + 1)
    If valteger = 1 Then
    vv = "š"
    ElseIf valteger = 2 Then: Whole$ = Whole$ + ""
    ElseIf valteger = 3 Then: Whole$ = Whole$ + "¢"
    ElseIf valteger = 4 Then: Whole$ = Whole$ + "±"
    ElseIf valteger = 5 Then: Whole$ = Whole$ + "¥"
    ElseIf valteger = 6 Then: Whole$ = Whole$ + "†"
    ElseIf valteger = 7 Then: Whole$ = Whole$ + "ø"
    ElseIf valteger = 8 Then: Whole$ = Whole$ + "‡"
    ElseIf valteger = 9 Then: Whole$ = Whole$ + "µ"
    ElseIf valteger = 10 Then: Whole$ = Whole$ + "þ"
    ElseIf valteger = 11 Then: Whole$ = Whole$ + "‰"
    ElseIf valteger = 12 Then: Whole$ = Whole$ + "¤"
    ElseIf valteger = 13 Then: Whole$ = Whole$ + "£"
    ElseIf valteger = 14 Then: Whole$ = Whole$ + "÷"
    ElseIf valteger = 15 Then: Whole$ = Whole$ + "º"
    ElseIf valteger = 16 Then: Whole$ = Whole$ + "©"
    ElseIf valteger = 17 Then: Whole$ = Whole$ + "*"
    End If
Next
X% = SendMessageByString(Subject%, WM_SETTEXT, 0, Whole$)
End Function

Function Extention(exe$, ext$)
If InStr(exe$, ".") Then
    where = InStr(exe$, ".")
    Filename$ = Mid$(exe$, 1, where - 1)
    FileType$ = Mid$(exe$, where + 2)
    If FileType$ <> ext$ Then FileType$ = ext$
    exe$ = Filename$ + "." + FileType$
ElseIf InStr(exe$, ".") = False Then
    exe$ = exe$ + "." + ext$
End If
End Function

Sub fastestfakeoh(room1$)
' change Command3d1 to whatever the start button is named
'Command3d1.Enabled = False

word:
Call keywor("aol://2719:2-2-" & (room1$))
AOL% = findwindow("aol Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
Room% = Findchildbyclass(mdi%, "AOL Child")
rooma% = findchildbytitle(mdi%, (room1$))
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
timeout 0.2
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "~~>GenOziDe "
timeout 0.2
'If Command3d1.Enabled = True Then
'x = waitforwindow((room1$), "AOL Child")
'killwin rooma%
'Exit Sub
'End If
GoTo word

End Sub

Function FileErrors(errVal As Integer) As Integer
' Return Value  Meaning     Return Value    Meaning
' 0             Resume      2               Unrecoverable error
' 1             Resume Next 3               Unrecognized error
Dim MsgType As Integer
Dim Response As Integer
Dim Action As Integer
Dim msg As String
MsgType = MB_EXCLAIM
Select Case errVal
    Case Err_DeviceUnavailable  ' Error #68
        msg = "That device appears to be unavailable."
        MsgType = MB_EXCLAIM + 5
    Case Err_DiskNotReady       ' Error #71
        msg = "The disk is not ready."
    Case Err_DeviceIO
        msg = "The disk is full."
    Case Err_BadFileName, Err_BadFileNameOrNumber   ' Errors #64 & 52
        msg = "That file name is illegal."
    Case Err_PathDoesNotExist                        ' Error #76
        msg = "That path doesn't exist."
    Case Err_BadFileMode                            ' Error #54
        msg = "Can't open your file for that type of access."
    Case Err_FileAlreadyOpen                        ' Error #55
        msg = "That file is already open."
    Case Err_InputPastEndOfFile                     ' Error #62
    msg = "This file has a nonstandard end-of-file marker,"
    msg = msg + "or an attempt was made to read beyond "
    msg = msg + "the end-of-file marker."
    Case Else
        FileErrors = 3
        Exit Function
    End Select
    Response = MsgBox(msg, MsgType, "File Error")
    Select Case Response
        Case 4          ' Retry button.
            FileErrors = 0
        Case 5          ' Ignore button.
            FileErrors = 1
        Case 1, 2, 3    ' Ok and Cancel buttons.
            FileErrors = 2
        Case Else
            FileErrors = 3
    End Select
End Function

Function FileOpener(NewFileName As String, Mode As Integer, RecordLen As Integer, Confirm As Integer) As Integer
     Dim NewFileNum As Integer
     Dim Action As Integer
     Dim FileExists As Integer
     Dim msg As String
     On Error GoTo OpenerError
     If NewFileName Like "*[;-?[* ]*" Or NewFileName Like "*]*" Then Error Err_BadFileName
     If Confirm Then
        If Dir(NewFileName) = "" Then
            FileExists = False
        Else
            FileExists = True
        End If
        If Mode = REPLACEFILE And FileExists Then
            msg = "Replace contents of " + NewFileName + "?"
            If MsgBox(msg, 49, "Replace File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
        If Not FileExists Then
            msg = "The file " + NewFileName + " does not exist. "
            msg = msg + "Do you want to create it?"
            If MsgBox(msg, 1, "Create File?") = 2 Then
                FileOpener = 0
                Exit Function
            End If
        End If
     End If
     NewFileNum = FreeFile
     Select Case Mode
          Case REPLACEFILE
            Open NewFileName For Output As NewFileNum
          Case READFILE
            Open NewFileName For Input As NewFileNum
          Case ADDTOFILE
            Open NewFileName For Append As NewFileNum
          Case RANDOMFILE
            Open NewFileName For Random As NewFileNum Len = RecordLen
          Case BINARYFILE
            Open NewFileName For Binary As NewFileNum
          Case Else
            Exit Function
     End Select
     FileOpener = NewFileNum
Exit Function
OpenerError:
     Action = FileErrors(Err)
     Select Case Action
        Case 0
            Resume
        Case Else
            FileOpener = 0
            Exit Function
     End Select
End Function

Function findaol()
Dim AOL
AOL = findwindow("AOL Frame25", 0&)
findaol = AOL
End Function

Function FindAOLChildByTitle(TitleText As String) As Integer
Dim X%
Dim ChildWnd As Integer
Dim MDIhWnd%
Dim AOLChildhWnd%
Dim RetClsName As String * 255
  
MDIhWnd% = getwindow(findwindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(MDIhWnd%, RetClsName$, 254)
  If InStr(RetClsName$, "MDIClient") Then AOLChildhWnd% = MDIhWnd%
  MDIhWnd% = getwindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
If TitleText = "MDIClient" Then FindAOLChildByTitle = AOLChildhWnd%
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  If InStr(WindowCaption(ChildWnd), TitleText) <> 0 Then
      FindAOLChildByTitle = ChildWnd
      Exit Do
  End If
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

Sub FindChatWerd(werd$)
start:
X$ = AOLLastChatLine()
y = InStr(LCase$(X$), LCase$(werd$))
sn$ = Left(X$, InStr(X$, ":") - 1)
SnSaid$ = Right(X$, Len(X$) - Len(sn$) - 2)
End Sub

Function findchatwnd() As Integer
  Dim MDIhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim X%
MDIhWnd% = getwindow(findwindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(MDIhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = MDIhWnd% 'Child window found!
    End If
  MDIhWnd% = getwindow(MDIhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While MDIhWnd% <> 0
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
findchatwnd = ChatWnd

End Function

Function findcomposemail()
Dim bb As Integer
Dim dis_win As Integer

dis_win = Findchildbyclass(aolhwnd(), "AOL Child")

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
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findtocomposemail = 0
   Exit Function
End If
GoTo begin_find_composemail
End Function

Function FindFwdWin() As Integer
Dim hWnds() As Integer
Dim R, X%, i, T, G%
DoEvents
X = 0
G% = getwindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Send") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Send Now") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Address" & Chr(13) & "Book") <> 0 Then X = X + 1
If X = 3 Then GoTo Founddddd
DoEvents
G% = getwindow(G%, 2)
DoEvents
X = 0
Loop While G% <> 0


Exit Function
Founddddd:
DoEvents
FindFwdWnd = G%
Exit Function
End Function

Function FindReadWnd() As Integer
Dim hWnds() As Integer
Dim R, X%, i, T, G%
X = 0
DoEvents
G% = getwindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Reply") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Forward") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Reply to All") <> 0 Then X = X + 1
DoEvents
If X = 3 Then GoTo Founddd
DoEvents
G% = getwindow(G%, 2)
DoEvents
X = 0
Loop While G% <> 0


Exit Function
Founddd:
DoEvents
FindReadWnd = G%
Exit Function


End Function

Function findsn()
Dim dis_win As Integer

dis_win = Findchildbyclass(aolhwnd(), "AOL Child")

begin_find_SN:

bb$ = WindowCaption(dis_win)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  findsn = Trim(sn$)
  Exit Function
End If
Let countt = 0
dis_win = getnextwindow(dis_win, 2)
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findsn = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

'
Function findtoim()
Dim bb As Integer
Dim dis_win As Integer
AOL% = findwindow("AOL Frame25", 0&)

dis_win = Findchildbyclass(AOL%, "AOL Child")

begin_find_To_im:

bb = findchildbytitle(dis_win, "Send")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "To:")
    If bb <> 0 Then Let countt = countt + 1

bb = findchildbytitle(dis_win, "Available?")
    If bb <> 0 Then Let countt = countt + 1

If countt = 3 Then
  findtoim = dis_win
  Exit Function
End If
Let countt = 0
dis_win = getnextwindow(dis_win, 2)
If dis_win = getwindow(dis_win, GW_HWNDLAST) Then
   findtoim = 0
   Exit Function
End If

GoTo begin_find_To_im
End Function

Function FindWindowLike(hWndArray() As Integer, ByVal hWndStart As Integer, WindowText As String, Classname As String, ID) As Integer
   Dim hWnd As Integer
   Dim sWindowText As String
   Dim sClassname As String
   Dim sID
   Dim R As Integer
   ' Hold the level of recursion:
   Static level As Integer
   ' Hold the number of matching windows:
   Static iFound As Integer

   ' Initialize if necessary:
   If level = 0 Then
      iFound = 0
      ReDim hWndArray(0 To 0)
      If hWndStart = 0 Then hWndStart = GetDeskTopWindow()
   End If

   ' Increase recursion counter:
   level = level + 1

   ' Get first child window:
   hWnd = getwindow(hWndStart, 5)

   Do Until hWnd = 0
      DoEvents ' Not necessary
      ' Search children by recursion:
      R = FindWindowLike(hWndArray(), hWnd, WindowText, Classname, ID)

      ' Get the window text and class name:
      sWindowText = Space(255)
      R = GetWindowText(hWnd, sWindowText, 255)
      sWindowText = Left(sWindowText, R)
      sClassname = Space(255)
      R = GetClassName(hWnd, sClassname, 255)
      sClassname = Left(sClassname, R)

      ' If window is a child get the ID:
      If getparent(hWnd) <> 0 Then
         R = GetWindowWord(hWnd, (-12))
         sID = CLng("&H" & Hex(R))
      Else
         sID = Null
      End If

      ' Check that window matches the search parameters:
      If sWindowText Like WindowText And sClassname Like Classname Then
         If IsNull(ID) Then
            ' If find a match, increment counter and add handle to array:
            iFound = iFound + 1
            ReDim Preserve hWndArray(0 To iFound)
            hWndArray(iFound) = hWnd
         ElseIf Not IsNull(sID) Then
            If sID = CLng(ID) Then
               ' If find a match increment counter and add handle to array:
               iFound = iFound + 1
               ReDim Preserve hWndArray(0 To iFound)
               hWndArray(iFound) = hWnd
            End If
         End If
      End If

      ' Get next child window:
      hWnd = getwindow(hWnd, 2)
   Loop

   ' Decrement recursion counter:
   level = level - 1

   ' Return the number of windows found:
   FindWindowLike = iFound

End Function

Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function

Function frac(whatever)
bah = Int(whatever)
valfrac = (Val(whatever) - Val(bah))
frac = valfrac
End Function

Sub FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Sub

Function gencc(prefix)
loogie:
a = 0
Randomize Timer
heh = 99 * Rnd
Do
If Val(heh) = 10 Then heh = heh - (Int(10 * Rnd))
If Val(heh) < 10 Then GoTo hhhf
If heh > 10 Then
heh = heh - Int(10 * Rnd)
End If
Loop Until heh < 10
hhhf:

g1 = 4
g2 = Int(Val(heh) * Rnd)
g3 = Int(Val(heh) * Rnd)
g4 = Int(Val(heh) * Rnd)
g5 = Int(Val(heh) * Rnd)
g6 = Int(Val(heh) * Rnd)
g7 = Int(Val(heh) * Rnd)
g8 = Int(Val(heh) * Rnd)
g9 = Int(Val(heh) * Rnd)
g10 = Int(Val(heh) * Rnd)
g11 = Int(Val(heh) * Rnd)
g12 = Int(Val(heh) * Rnd)
g13 = Int(Val(heh) * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)

g13 = Int(Val(heh) * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)

gee = 0
hehd:
Do
T1$ = g1
T2$ = g2
T3$ = g3
T4$ = g4
T5$ = g5
T6$ = g6
T7$ = g7
T8$ = g8
T9$ = g9
T10$ = g10
T11$ = g11
T12$ = g12
T13$ = g13
T14$ = g14
T15$ = g15
T16$ = g16
evens = Val(T2$) + Val(T4$) + Val(T6$) + Val(T8$) + Val(T10$) + Val(T12$) + Val(T14$) + Val(T16$)
C1 = T1$
C3 = T3$
C5 = T5$
C7 = T7$
C9 = T9$
C11 = T11$
C13 = T13$
C15 = T15$
C1 = Val(C1) + Val(C1)
C3 = Val(C3) + Val(C3)
C5 = Val(C5) + Val(C5)
C7 = Val(C7) + Val(C7)
C9 = Val(C9) + Val(C9)
C11 = Val(C11) + Val(C11)
C13 = Val(C13) + Val(C13)
C15 = Val(C15) + Val(C15)
If C1 > 9 Then C1 = C1 - 9
If C3 > 9 Then C3 = C3 - 9
If C5 > 9 Then C5 = C5 - 9
If C7 > 9 Then C7 = C7 - 9
If C9 > 9 Then C9 = C9 - 9
If C11 > 9 Then C11 = C11 - 9
If C13 > 9 Then C13 = C13 - 9
If C15 > 9 Then C15 = C15 - 9
a = 1 + a
If a = 10 Then GoTo loogie
odds = Val(C1) + Val(C3) + Val(C5) + Val(C7) + Val(C9) + Val(C11) + Val(C13) + Val(C15)
If Int((odds + evens) / 60) = (odds + evens) / 60 Then GoTo bob

hah = 60
If gee = 0 Then g16 = -1
g16 = g16 + 1
If g16 > 10 Then
g13 = Int(10 * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)
gee = 0
GoTo hehd
End If
gee = 1
Loop
bob:
bahq$ = g1 & g2 & g3 & g4 & "-" & g5 & g6 & g7 & g8 & "-" & g9 & g10 & g11 & g12 & "-" & g13 & g14 & g15 & g16
gencc = bahq$
End Function

Function Generate() As String
Randomize
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = "Hello, "
      If phrase = 2 Then T$ = "Good evening, "
      If phrase = 3 Then T$ = "Hello, And welcome to America Online. "
      If phrase = 4 Then T$ = "Welcome to America Online. "
      If phrase = 5 Then T$ = "Excuse me, "
      If phrase = 6 Then T$ = "Dear User, "
      If phrase = 7 Then T$ = "What's Up! "
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "I am with America Online Billing "
      If phrase = 2 Then T$ = T$ & "I am with OTC (Online Techincal Consultants) "
      If phrase = 3 Then T$ = T$ & "I am with WWWC (World Wide Web Consultants) "
      If phrase = 4 Then T$ = T$ & "I am with AOL Techincal Staff "
      If phrase = 5 Then T$ = T$ & "I am with AOL System Security "
      If phrase = 6 Then T$ = T$ & "I am with AOL Resource Department "
      If phrase = 7 Then T$ = T$ & "I am with the AOL Community Action Team "
    phrases = 7
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "and due to Technical Failures "
      If phrase = 2 Then T$ = T$ & "and due to Billing errors "
      If phrase = 3 Then T$ = T$ & "and due to Line Noise "
      If phrase = 4 Then T$ = T$ & "and due to A virus in our database "
      If phrase = 5 Then T$ = T$ & "and due to Massive data flow in our database "
      If phrase = 6 Then T$ = T$ & "and due to data corruption "
      If phrase = 7 Then T$ = T$ & "and due to hackers by-passing out systems "

    phrases = 5
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "we seem to have lost your password. "
      If phrase = 2 Then T$ = T$ & "we seem to have lost your account information. "
      If phrase = 3 Then T$ = T$ & "we seem to have failed to recieve your logon password. "
      If phrase = 4 Then T$ = T$ & "we seem to have failed to recieve your account information. "
      If phrase = 5 Then T$ = T$ & "we have lost your Credit Card information. "
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "To correct this situation please click respond and enter "
      If phrase = 2 Then T$ = T$ & "Please click respond and enter your "
      If phrase = 3 Then T$ = T$ & "Please help us by entering your "
      If phrase = 4 Then T$ = T$ & "Please respond with your "
    phrases = 3
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & "Your password information."
      If phrase = 2 Then T$ = T$ & "Your current account password."
      If phrase = 3 Then T$ = T$ & "Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number."
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & " Please respond within 2 minutes too keep this account active."
      If phrase = 2 Then T$ = T$ & " It is very important that you respond immediately."
      If phrase = 3 Then T$ = T$ & " Please respond as soon as possible "
      If phrase = 4 Then T$ = T$ & " Incooperation may lead to termination of your account."
    phrases = 4
    phrase = Int(Rnd * phrases + 1)
      If phrase = 1 Then T$ = T$ & " Thank you for using America Online."
      If phrase = 2 Then T$ = T$ & " Thank you for your time."
      If phrase = 3 Then T$ = T$ & " Thank you for your cooperation."
      If phrase = 4 Then T$ = T$ & " Thank you for your help and enjoy the service!"
Generate = T$
End Function

Function getaolwinb(Parent As Integer, ClassToFind As String, Num As Integer)
''——-=-By:GenOziDe-=-——''
Dim dood As Integer
dood = 0
Count = (Num) 'Count Out The Number Of Windows Over You Want To Look For
'If num = 0 Then MsgBox "Are You An Asshole": Exit Function'Won't Let You Look For Nothing
DoEvents
a% = Findchildbyclass(Parent%, ClassToFind$) 'Begin Your Search For Da Window
dood = dood + 1 'If You Find One Add 1 To Your Counter
Do
DoEvents
If dood = Num Then 'If Your Counter = Your Number Then Exit Function
getaolwinb = a% 'Declare The Function
Exit Function
End If
DoEvents
Do         'Begin a Do...Loop to look For The Class Name
DoEvents
a% = getwindow(a%, GW_HWNDNEXT)
bb$ = String(255, 0)
CC% = GetClassName(a%, bb$, 254) 'Use This To Get The Class Name Of The Window You Found
bb$ = Disc(bb$)
If bb$ = ClassToFind$ Then dood = dood + 1 'Then Compare
If dood = Num Then Exit Do
Loop Until bb$ = ClassToFind$ 'Loop Until You Find The Window With The Class Name You're Looking For
Loop Until dood = Num 'Loop Until Your Counter Is = To Your Number
getaolwinb = a% 'Declare The Function

End Function

Function GetCaption(hWnd)
hwndLength% = GetWindowtextlength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetchatText() As String
Dim chattext
View% = ChatView()
If View% = 0 Then Exit Function
i = SendMessageByNum(View%, WM_GETTEXTLENGTH, 0, 0)
chattext = Space(i)
 
chattext = Trim$(chattxt$)
For X = Len(chattxt$) To 1 Step -1
    letter$ = Mid$(chattxt$, X)
    letter$ = Mid$(letter$, 1, 1)
    If letter$ = Chr$(13) Then
        chattext = Mid$(chattxt$, X, Len(chattxt$))
        Exit For
    End If
Next
GetchatText = chattxt$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
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
                    X = "Other"
    End If

End Function

Function GetFileName(Prompt As String) As String
    GetFileName = LTrim$(RTrim$(UCase$(InputBox$(Prompt, "Enter File Name"))))
End Function

Function GetFreeGDI(X)
'Example: text5.text = "Free GDI Resources: " & sGetFreeGDI
    X = Format$(getfreesystemresources(GFSR_GDIRESOURCES)) + "%"

End Function

Function GetFreeSYS(X)
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    X = Format$(getfreesystemresources(GFSR_SYSTEMRESOURCES)) + "%"

End Function

Function GetFreeUser(X)
'Example: text4.text = "Free User Resources: " & sGetFreeUser
    X = Format$(getfreesystemresources(GFSR_USERRESOURCES)) + "%"

End Function

Function GetLast(ByVal txt As String)
Dim X As Integer
Do
X = X + 1
Loop Until Mid(txt, Len(txt) - X, 1) = Chr(13)
GetLast = Right(txt, X)
End Function

Sub getlisttext(Lst As ListBox)
On Error Resume Next
Do Until xx = Lst.ListCount
Let da_new$ = Lst.List(xx)
If Trim(da_new$) <> "" Then Let namms = namms & da_new$ & "," Else Exit Sub
Let xx = xx + 1
Loop

End Sub

Sub getnum(og%, a)
Do
If a = 0 Then Exit Sub
B = 1 + B
og% = getwindow(og%, GW_HWNDNEXT)
Loop Until B >= a - 1
End Sub

Sub GetRoom(Lst As ListBox)
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
List% = Findchildbyclass(AOL%, "_AOL_Listbox")
Room% = getparent(List%)
For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_add
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
ListCheck namez$, Lst
Next Index%
end_add:

End Sub

Function GetSNFromIM()
Dim mdi%
mdi% = findwindow("AOL Frame25", 0&)
Dim welcome%
welcome% = findchildbytitle(mdi%, ">Instant Message From: ")
Dim yourname As String * 255
Dim YourName2$
X = GetWindowText(welcome%, yourname, 255)
yourname = LTrim(RTrim(Trim(yourname)))

yourname = Right$(yourname, Len(yourname) - InStr(yourname, ": "))
YourName2 = Left$(yourname, 24)
YourName2 = Right$(yourname, Len(yourname) - 1)
If InStr(YourName2, "") <> 0 Then YourName2 = Left$(YourName2, InStr(YourName2, ""))
SName$ = YourName2$
GetSNFromIM = SName$

End Function

Function GetTextFromRICHCNTL(hWindow As Integer)
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
msg$ = String$(255, 0)
xx% = SendMessageLong(hWindow, WM_ProGGer, 254, msg$)
msg$ = Trim$(msg$)
GetTextFromRICHCNTL = msg$

End Function

Function getuser()

'Find the welcome window
wlcm% = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "Welcome, ")

'Get the caption of it
dacap = WindowCaption(wlcm%)

'Extract the user's screen name
If wlcm% <> 0 Then
    numba% = (InStr(dacap, "!") - 10)
    Pname$ = Mid$(dacap, 10, numba%)
    getuser = Pname$
Else
    getuser = "(unknown)"
End If

End Function

Function GetWindowDir()
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
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

Function GetWinText(hWnd As Integer) As String
lentos = sendmessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
X = SendMessageByString(hWnd, wm_gettext, lentos + 1, Buffer$)
GetWinText = Buffer$
End Function

Function GetWinVer(X)
'Example: text2.text = "Window version: " & sGetWinVer
Dim lVer As Long, iWinVer As Integer
    lVer = GetVersion()
    iWinVer = CInt(lVer And &HFFFF&)
    X = Format$(iWinVer And &HFF) + "." + Format$(CInt(iWinVer / 256))

End Function

Sub guidetos(who$, phrase$)
Run "Keyword"
AOL = findwindow("AOL Frame25", 0&)
Do
keyword = findchildbytitle(AOL, "Keyword")
timeout (0.001)
EditBox = Findchildbyclass(keyword, "_AOL_Edit")
Loop Until EditBox <> 0
EditBox = Findchildbyclass(keyword, "_AOL_Edit")
send EditBox, "Guidepager"
go% = Findchildbyclass(keyword, "_AOL_Icon")
Click go%

AOL = findwindow("AOL Frame25", 0&)
mow = findchildbytitle(AOL, "I NEED HELP")
ikqn% = Findchildbyclass(mow, "_AOL_Icon")
DoEvents

Click ikqn%
timeout (0.001)




timeout (0.001)


getnum whototosbox%, 7
send whototosbox%, who$
getnum whototosbox%, 3
End Sub

Sub HideAOL()
a = findwindow("aol frame25", 0&)
B = ShowWindow(a, SW_HIDE)
End Sub

Sub HideAOLToolbar()
a = findwindow("aol frame25", 0&)
B = Findchildbyclass(a, "AOL Toolbar")
C = ShowWindow(B, SW_HIDE)
End Sub

Sub HideBuddylist()
a = findwindow("aol frame25", 0&)
B = Findchildbyclass(a, "MDIClient")
C = findchildbytitle(B, "Buddy Lists")
D = ShowWindow(C, SW_HIDE)

End Sub

Sub hidewelcome()
Dim X
wlcm% = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "Welcome, ")
X = ShowWindow(wlcm%, SW_HIDE)
End Sub

Sub hidewin(windo)
X = ShowWindow(windo, SW_HIDE)
End Sub

Sub Hit_Menu(Mnu_str$)
AOL% = findwindow("AOL Frame25", 0&)
mnu% = GetMenu(AOL%)
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
            junk% = SendMessageByNum(AOL%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Sub

Function IFileExists(ByVal sFileName As String) As Integer
'Example: If Not IFileExists("win.com") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IFileExists = False
        Else
            IFileExists = True
    End If

End Function

Sub im_off()
AOLInstantMessage "$IM_OFF", "      ÷•¤•^v^v^•(• Broken Punter •)•^v^v^•¤•÷" & Chr(13) & "          ÷•¤•^v^v^•(• By KoRn •)•^v^v^•¤•÷"
waitforok
End Sub

Sub im_on()
AOLInstantMessage "$IM_ON", "      ÷•¤•^v^v^•(• Broken Punter •)•^v^v^•¤•÷" & Chr(13) & "          ÷•¤•^v^v^•(• By KoRn •)•^v^v^•¤•÷"
waitforok
End Sub

Sub imansweringmachine(wuttosay As TextBox)
'this code goes in a timer
IM% = findchildbytitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo begin
IM% = findchildbytitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo begin
Exit Sub
begin:
e = Findchildbyclass(IM%, "RICHCNTL")

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
Call AOLSetText(e2, ((wuttosay) & Chr(13) & "ÐeliriuM by GenOziDe"))
Click (e)
timeout 0.1
killwin (IM%)

End Sub

Sub imbomb(numofbom As TextBox, whotobom As TextBox, wuttosay As TextBox)
If numofbom = "0" Then
Exit Sub
End If
If numofbom < 0 Then
Exit Sub
End If
Do
Call AOLInstantMessage((whotobom), (wuttosay) & Chr(13) & "      ÷•¤•^v^v^•(• Broken Punter •)•^v^v^•¤•÷" & Chr(13) & "          ÷•¤•^v^v^•(• By KoRn •)•^v^v^•¤•÷")
numofbom = Str(Val(numofbom - 1))
Loop Until numofbom = 0
Exit Sub
End Sub

Function imTOS(who$)
Dim wer As Variant
If UCase$(who$) Like UCase$("GenOziDe") Then
MsgBox "LoL Nice Try! Say goodbye to your main index!", 55, "Cant TOS GenOziDe!!"
Exit Function
End If
sendtext "  "
timeout 0.5
Do
Dim hWnds() As Integer
kw% = findchildbytitle(AOL%, "Keyword")
timeout 0.2
Loop Until kw% <> 0
edi% = Findchildbyclass(kw%, "_AOL_Edit")
TOS = SendMessageByString(edi%, WM_SETTEXT, 0, "KO Help")

AOL% = findwindow("AOL Frame25", 0&)
Do
DoEvents
XTY% = findchildbytitle(AOL%, "I Need Help!")
Loop Until XTY% <> 0
R = FindWindowLike(hWnds(), Find%, "*", "_AOL_Icon", Null)
IM = hWnds(3)
X% = SendMessageByNum(IM, WM_LBUTTONDOWN, 0, 0&)
X% = SendMessageByNum(IM, WM_LBUTTONUP, 0, 0&)




Do
DoEvents
XTY% = findchildbytitle(AOL%, "Report A Violation")
Loop Until XTY% <> 0
R = FindWindowLike(hWnds(), Find%, "*", "_AOL_Icon", Null)
imv = hWnds(2)
X% = SendMessageByNum(imv, WM_LBUTTONDOWN, 0, 0&)
X% = SendMessageByNum(imv, WM_LBUTTONUP, 0, 0&)

Do
DoEvents
XP% = findchildbytitle(AOL%, "Violations via Instant Messages")
Loop Until XP% <> 0

asnm% = findchildbytitle(XP%, "Date")
 
anx% = getwindow(asnm%, GW_HWNDNEXT)
asm% = SendMessageByString(anx%, WM_SETTEXT, 0, "" & Date)

esnm% = findchildbytitle(XP%, "Category of Chat Room")
 
enx% = getwindow(esnm%, GW_HWNDNEXT)
esm% = SendMessageByString(enx%, WM_SETTEXT, 0, "Life")


qsnm% = findchildbytitle(XP%, "Time AM/PM")
 
qnx% = getwindow(qsnm%, GW_HWNDNEXT)
qsm% = SendMessageByString(qnx%, WM_SETTEXT, 0, "" & Time)

wsnm% = findchildbytitle(XP%, "Chat Room Name")
 
wnx% = getwindow(wsnm%, GW_HWNDNEXT)
wsm% = SendMessageByString(wnx%, WM_SETTEXT, 0, "Teen Chat5")






snm% = findchildbytitle(XP%, "CUT and PASTE a copy of the IM here")
 
nx% = getwindow(snm%, GW_HWNDNEXT)
sm% = SendMessageByString(nx%, WM_SETTEXT, 0, who$ + ":" & " " + text2 & Chr(10))
Sen% = findchildbytitle(XP%, "Send")
timeout 0.1
STAtc% = findchildbytitle(soltos%, "SEND")
waitforok
 
ax2 = sendmessage(XP%, WM_CLOSE, 0, 0)
SendKeys "{f4}"
sendtext "- – — IM Violation TOS complete — – -"
sendtext "– — —– —— ÐeliriuM by GenOziDe —— –— — –"
Exit Function


End Function

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Sub invitationbomb(towho As TextBox, numberof As TextBox)
Do
a = findsn()
eek% = findchildbytitle(mdi%, "Invitation From: " & a)
If eek% <> 0 Then
killwin eek%
End If
If numberof < 1 Then Exit Sub
'If command3d1.Enabled = True Then Exit Sub
AOL% = findwindow("aol Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
child% = Findchildbyclass(mdi%, "AOL Child")
titl% = findchildbytitle(mdi%, "Buddy Lists")
e = Findchildbyclass(titl%, "_AOL_Icon")
e2 = getnextwindow(e, 2)
e3% = getnextwindow(e2, 2)
e4% = getnextwindow(e3%, 2)
Click e4%
timeout 0.1
Do: DoEvents
buddy% = Findchildbyclass(mdi%, "AOL Child")
buddy2% = findchildbytitle(mdi%, "Buddy Chat")
EDT% = Findchildbyclass(buddy2%, "_AOL_Edit")
edt2% = getnextwindow(EDT%, 2)
edt3% = getnextwindow(edt2%, 2)
edt4% = getnextwindow(edt3%, 2)
edt5% = getnextwindow(edt4%, 2)
edt6% = getnextwindow(edt5%, 2)
edt7% = getnextwindow(edt6%, 2)
edt8% = getnextwindow(edt7%, 2)
edt9% = getnextwindow(edt8%, 2)
Loop Until buddy2% <> 0
timeout 0.5
textset EDT%, CStr((towho))
textset edt3%, CStr("INVITATION BOMB Your invited to HELL")
textset edt9%, CStr("°º*¤FATE Infinity¤*º° by GenOziDe")
icn% = Findchildbyclass(buddy2%, "_AOL_Icon")
Click icn%
timeout 1
a = findsn()
eek% = findchildbytitle(mdi%, "Invitation From: " & a)
killwin eek%
timeout 0.1
numberof = Str(Val(numberof - 1))
If numberof = Str(Val("0")) Then Exit Sub
Loop

End Sub

Sub keywor(where$)
Run "Keyword"
Do
AOL = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(AOL, "Keyword")
daedit = Findchildbyclass(bah, "_AOL_Edit")
timeout (0.001)
Loop Until daedit <> 0
daedit = Findchildbyclass(bah, "_AOL_Edit")
send daedit, where$
ico% = Findchildbyclass(bah, "_AOL_Icon")
Click ico%
End Sub

Sub KillModal()
Do
AOM = findwindow("_AOL_Modal", 0&)
killwin (AOM)
DoEvents
Loop Until AOM = 0

End Sub

Sub KillWait()
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
If aolversion() = 25 Then
    AOL% = findwindow("AOL Frame25", 0&)
    mdi% = Findchildbyclass(AOL%, "MDIClient")
     
    Do
        C% = DoEvents()
        EXI% = findwindow("_AOL_Modal", "America Online")
    Loop Until EXI% <> 0
    C% = DoEvents()
    D = SendMessageByNum(EXI%, &H10, 0, 0)
End If
If aolversion() = 30 Then
    AOL% = findwindow("AOL Frame25", 0&)
    mdi% = Findchildbyclass(AOL%, "MDIClient")
    Hit_Menu "Preferences"
    Do
        C% = DoEvents()
        PRE% = findchildbytitle(mdi%, "Preferences")
    Loop Until PRE% <> 0
    ICC% = Findchildbyclass(PRE%, "_AOL_Icon")
    Click (ICC%)
    Do
        C% = DoEvents()
        MOO% = findwindow("_AOL_Modal", "Chat Preferences")
    Loop Until MOO% <> 0
    D = SendMessageByNum(MOO%, &H10, 0, 0)
    D = SendMessageByNum(PRE%, &H10, 0, 0)
End If
C% = DoEvents()
End Sub

Sub killwin(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

Function KTEncrypt(ByVal password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

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
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)

    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

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
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
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

Function LineFromText(Text, theline)
theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
C = C + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = C Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")
LineFromText = thechatext$


End Function

Sub ListCheck(itm As String, Lst As ListBox)

If itm = Sccc Then Exit Sub
If Lst.ListCount = 0 Then Lst.AddItem itm: Exit Sub

Do Until xx = (Lst.ListCount)
Let diss_itm$ = Lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub

Sub listclick(button)
SendNow = SendMessageByNum(button, LBN_DBLCLK, &HD, 0)
End Sub

Sub loadaol()
Dim X
X = Shell("C:\aol30\waol.exe")
End Sub

Function LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, where + Len(look$))
        LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function

Sub LogOff_Quick()
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then Exit Sub
X = setfocusapi(AOL%)
X = SendMessageByNum(AOL%, &H11, 0, &H0)
Do
X = DoEvents()
Beav% = findwindow("_AOL_Modal", "America Online")
Loop Until Beav% <> 0
oK% = findchildbytitle%(Beav%, "&Yes")
Goob = SendMessageByNum(oK%, &H201, 0, 0&)
Goob = SendMessageByNum(oK%, &H202, 0, 0&)
End Sub

Sub Mail(sill$, Subject$, Text$)
Do
AOL% = findwindow("AOL Frame25", 0&)
heh = findchildbytitle(AOL%, "Compose Mail")
send heh, "Compose Mail"
DoEvents
Loop Until heh = 0
AOL% = findwindow("AOL Frame25", 0&)
Run "Compose"
Do
X = DoEvents()
chatlist% = findchildbytitle(AOL%, "Compose Mail")
chatedit% = Findchildbyclass(chatlist%, "_AOL_Edit")
Loop Until chatedit% <> 0

chatwin% = getparent(chatlist%)
button% = Findchildbyclass(chatlist%, "_AOL_Icon")
ChatViw% = Findchildbyclass(chatlist%, "_AOL_View")
chatedit% = Findchildbyclass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, sill$)
blah% = getwindow(chatedit%, GW_HWNDNEXT)
good% = getwindow(blah%, GW_HWNDNEXT)
bad% = getwindow(good%, GW_HWNDNEXT)
Sad% = getwindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, Subject$)
Nad% = getwindow(Sad%, GW_HWNDNEXT)
Wad% = getwindow(Nad%, GW_HWNDNEXT)
Dad% = getwindow(Wad%, GW_HWNDNEXT)
fad% = getwindow(Dad%, GW_HWNDNEXT)
Qad% = getwindow(fad%, GW_HWNDNEXT)
Ead% = getwindow(Qad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Dad%, WM_SETTEXT, 0, Text$ & " ")
rich = Findchildbyclass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, Text$ & " ")

timeout (0.5)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
timeout (0.1)
Do
AOL% = findwindow("AOL Frame25", 0&)
heh = findchildbytitle(AOL%, "Compose Mail")
send heh, "Compose Mail"
DoEvents
Loop Until heh = 0
End Sub

'
Sub mailopen()
If findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail") = False Then
X = runmenu(4, 3)
Do
DoEvents
box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
Loop Until box
List = Findchildbyclass(box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)

Call timeout(1)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(1)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
Else
    box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
    List = Findchildbyclass(box, "_AOL_Tree")
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
End If
End Sub

Sub maniptext(who$, wut$)
AOL% = findwindow("AOL Frame25", 0&)
chatwin% = Findchildbyclass(AOL%, "MDIClient")
blah% = Findchildbyclass(chatwin%, "AOL Child")
ChatVew% = Findchildbyclass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (who$) & ":" & Chr(9) & (wut$))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)

End Sub

Sub massinstantmessage(Lst As ListBox, Text$)
Do
For i = 0 To Lst.ListCount - 1
If m001C% = 1 Then Exit Sub
who$ = Lst.List(0)
Lst.ListIndex = 0
Next i
okw = findwindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
Run "Send an Instant Message"
Do
AOL = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(AOL, "Send Instant Message")
txt% = Findchildbyclass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
txt% = Findchildbyclass(bah, "_AOL_Edit")
Do
rich% = Findchildbyclass(bah, "RICHCNTL")
bahqw% = findchildbytitle(bah, "Send")
DoEvents
timeout (0.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
send txt%, who$
send rich%, ((Text$) & Chr(13) & "      <//=-——• FATE Infinity •——-=\\>" & Chr(13) & "        ƒ——by GenOziDe——ƒ")
timeout (0.001)
getnum rich%, 1
Click rich%
Else
send txt%, Person$
getnum txt%, 1
send txt%, tt$
getnum txt%, 1
Click txt%
End If
timeout (0.001)
X = SendMessageByNum(bah, WM_CLOSE, 0, 0)
a = Lst.List(0)
Call delitem(Lst, (a))
Loop Until Lst.ListCount = 0
If Lst.ListCount = 0 Then
Exit Sub
End If

End Sub

Sub MassMailBot(triggerwerd$, Lst As ListBox)
BeGinen:
X$ = AOLLastChatLine()
y = InStr(LCase$(X$), LCase$(triggerwerd$))
sn$ = Left(X$, InStr(X$, ":") - 1)
SnSaid$ = Right(X$, Len(X$) - Len(sn$) - 2)
w = InStr(y, (triggerwerd$))
If w <> 0 Then
Lst.AddItem sn$
sendtext "—ƒ" & sn$ & ", you have been added."
End If
GoTo BeGinen
End Sub

Function MENUTRIM(Ands) As String
Dim X, a, B
     For X = 1 To Len(Ands)
          If Mid$(Ands, X, 1) = "&" Then
               a = Mid$(Ands, 1, X - 1)
               B = Mid$(Ands, X + 1, Len(Ands))
               Ands = a & B
          End If
          
          If InStr(Ands, "Ctrl") <> 0 Then
               Ands = Mid$(Ands, 1, InStr(Ands, "Ctrl") - 1)
          End If
          
          If InStr(Ands, " Shift") Then
               Ands = Mid$(Ands, 1, InStr(Ands, " Shift") - 1)
          End If
          
          If InStr(Ands, Chr$(9)) <> 0 Then
               Ands = Mid$(Ands, 1, InStr(Ands, Chr$(9)) - 1)
          End If
          
     Next
     
     MENUTRIM = Ands


End Function

Function mysn()
OL = findwindow%("AOL Frame25", "America  Online")
Happy% = findchildbytitle%(AOL, "Welcome,")
joy$ = GetWinText((Happy%))
If InStr(joy$, "!") = 0 Then Exit Function
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
mysn = yoursn
End Function

Sub ohscroll()
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)
timeout 0.5
sendtext "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4)

End Sub

Sub Own_Host(Host As String, SayWhat As String)
AOL% = findwindow("AOL Frame25", 0&)
Room% = findchatwnd()
View% = Findchildbyclass(Room%, "_AOL_VIEW")
Sng$ = CStr(Chr(13) + Chr(10) + Chr(13) + Chr(10) + Host$ + ":" + Chr(9) + SayWhat$ + Chr(13) + Chr(10))
Q% = SendMessageByString(View%, WM_SETTEXT, 0, Sng$)
DoEvents
End Sub

Function Percent(Complete As Integer, total As Integer, TotalOutput As Integer) As Integer
R = Int(Complete / total * TotalOutput)
Percent = R
End Function

Sub phishgen2(txt As TextBox, txt2 As TextBox)
Randomize
a = Int((Val("26") * Rnd) + 1)
If a = "1" Then
a = ""
ElseIf a = "2" Then: a = "B"
ElseIf a = "3" Then: a = "C"
ElseIf a = "4" Then: a = "D"
ElseIf a = "5" Then: a = ""
ElseIf a = "6" Then: a = "F"
ElseIf a = "7" Then: a = "G"
ElseIf a = "8" Then: a = "H"
ElseIf a = "9" Then: a = ""
ElseIf a = "10" Then: a = "J"
ElseIf a = "11" Then: a = "K"
ElseIf a = "12" Then: a = "L"
ElseIf a = "13" Then: a = "M"
ElseIf a = "14" Then: a = "N"
ElseIf a = "15" Then: a = ""
ElseIf a = "16" Then: a = "P"
ElseIf a = "17" Then: a = "Q"
ElseIf a = "18" Then: a = "R"
ElseIf a = "19" Then: a = "S"
ElseIf a = "20" Then: a = "T"
ElseIf a = "21" Then: a = ""
ElseIf a = "22" Then: a = "V"
ElseIf a = "23" Then: a = "W"
ElseIf a = "24" Then: a = "Y"
ElseIf a = "25" Then: a = "X"
ElseIf a = "26" Then: a = "Z"
End If
txt = a

Randomize
B = Int((Val("37") * Rnd) + 1)
If B = "1" Then
B = "A"
ElseIf B = "2" Then: B = "B"
ElseIf B = "3" Then: B = "C"
ElseIf B = "4" Then: B = "D"
ElseIf B = "5" Then: B = "E"
ElseIf B = "6" Then: B = "F"
ElseIf B = "7" Then: B = "G"
ElseIf B = "8" Then: B = "H"
ElseIf B = "9" Then: B = "I"
ElseIf B = "10" Then: B = "J"
ElseIf B = "11" Then: B = "K"
ElseIf B = "12" Then: B = "L"
ElseIf B = "13" Then: B = "M"
ElseIf B = "14" Then: B = "N"
ElseIf B = "15" Then: B = "O"
ElseIf B = "16" Then: B = "P"
ElseIf B = "17" Then: B = "Q"
ElseIf B = "18" Then: B = "R"
ElseIf B = "19" Then: B = "S"
ElseIf B = "20" Then: B = "T"
ElseIf B = "21" Then: B = "U"
ElseIf B = "22" Then: B = "V"
ElseIf B = "23" Then: B = "W"
ElseIf B = "24" Then: B = "Y"
ElseIf B = "25" Then: B = "X"
ElseIf B = "26" Then: B = "Z"
ElseIf B = "27" Then: B = "0"
ElseIf B = "28" Then: B = "1"
ElseIf B = "29" Then: B = "2"
ElseIf B = "30" Then: B = "3"
ElseIf B = "31" Then: B = "4"
ElseIf B = "32" Then: B = "5"
ElseIf B = "33" Then: B = "6"
ElseIf B = "34" Then: B = "7"
ElseIf B = "35" Then: B = "8"
ElseIf B = "36" Then: B = "9"
ElseIf B = "37" Then: B = " "
End If
txt = a + B

Randomize
C = Int((Val("6") * Rnd) + 1)
If C = "1" Then
C = "A"
ElseIf C = "2" Then: C = "E"
ElseIf C = "3" Then: C = "I"
ElseIf C = "4" Then: C = "O"
ElseIf C = "5" Then: C = "U"
ElseIf C = "6" Then: C = " "
End If
txt = a + B + C

Randomize
D = Int((Val("37") * Rnd) + 1)
If D = "1" Then
D = " "
ElseIf D = "2" Then: D = "B"
ElseIf D = "3" Then: D = "C"
ElseIf D = "4" Then: D = "D"
ElseIf D = "5" Then: D = " "
ElseIf D = "6" Then: D = "F"
ElseIf D = "7" Then: D = "G"
ElseIf D = "8" Then: D = "H"
ElseIf D = "9" Then: D = " "
ElseIf D = "10" Then: D = "J"
ElseIf D = "11" Then: D = "K"
ElseIf D = "12" Then: D = "L"
ElseIf D = "13" Then: D = "M"
ElseIf D = "14" Then: D = "N"
ElseIf D = "15" Then: D = " "
ElseIf D = "16" Then: D = "P"
ElseIf D = "17" Then: D = "Q"
ElseIf D = "18" Then: D = "R"
ElseIf D = "19" Then: D = "S"
ElseIf D = "20" Then: D = "T"
ElseIf D = "21" Then: D = " "
ElseIf D = "22" Then: D = "V"
ElseIf D = "23" Then: D = "W"
ElseIf D = "24" Then: D = "Y"
ElseIf D = "25" Then: D = "X"
ElseIf D = "26" Then: D = "Z"
ElseIf D = "27" Then: D = "0"
ElseIf D = "28" Then: D = "1"
ElseIf D = "29" Then: D = "2"
ElseIf D = "30" Then: D = "3"
ElseIf D = "31" Then: D = "4"
ElseIf D = "32" Then: D = "5"
ElseIf D = "33" Then: D = "6"
ElseIf D = "34" Then: D = "7"
ElseIf D = "35" Then: D = "8"
ElseIf D = "36" Then: D = "9"
ElseIf D = "37" Then: D = " "

End If
txt = a + B + C + D

Randomize
e = Int((Val("6") * Rnd) + 1)
If e = "1" Then
e = "A"
ElseIf e = "2" Then: e = "E"
ElseIf e = "3" Then: e = "I"
ElseIf e = "4" Then: e = "O"
ElseIf e = "5" Then: e = "U"
ElseIf e = "6" Then: e = " "
End If
txt = a + B + C + D + e

Randomize
F = Int((Val("37") * Rnd) + 1)
If F = "1" Then
F = ""
ElseIf F = "2" Then: F = "B"
ElseIf F = "3" Then: F = "C"
ElseIf F = "4" Then: F = "D"
ElseIf F = "5" Then: F = ""
ElseIf F = "6" Then: F = "F"
ElseIf F = "7" Then: F = "G"
ElseIf F = "8" Then: F = "H"
ElseIf F = "9" Then: F = ""
ElseIf F = "10" Then: F = "J"
ElseIf F = "11" Then: F = "K"
ElseIf F = "12" Then: F = "L"
ElseIf F = "13" Then: F = "M"
ElseIf F = "14" Then: F = "N"
ElseIf F = "15" Then: F = ""
ElseIf F = "16" Then: F = "P"
ElseIf F = "17" Then: F = "Q"
ElseIf F = "18" Then: F = "R"
ElseIf F = "19" Then: F = "S"
ElseIf F = "20" Then: F = "T"
ElseIf F = "21" Then: F = ""
ElseIf F = "22" Then: F = "V"
ElseIf F = "23" Then: F = "W"
ElseIf F = "24" Then: F = "Y"
ElseIf F = "25" Then: F = "X"
ElseIf F = "26" Then: F = "Z"
ElseIf F = "27" Then: F = "0"
ElseIf F = "28" Then: F = "1"
ElseIf F = "29" Then: F = "2"
ElseIf F = "30" Then: F = "3"
ElseIf F = "31" Then: F = "4"
ElseIf F = "32" Then: F = "5"
ElseIf F = "33" Then: F = "6"
ElseIf F = "34" Then: F = "7"
ElseIf F = "35" Then: F = "8"
ElseIf F = "36" Then: F = "9"
ElseIf F = "37" Then: F = " "
End If
txt = a + B + C + D + e + F

Randomize
G = Int((Val("6") * Rnd) + 1)
If G = "1" Then
G = "A"
ElseIf G = "2" Then: G = "E"
ElseIf G = "3" Then: G = "I"
ElseIf G = "4" Then: G = "O"
ElseIf G = "5" Then: G = "U"
ElseIf G = "6" Then: G = " "
End If
txt = a + B + C + D + e + F + G

Randomize
h = Int((Val("37") * Rnd) + 1)
If h = "1" Then
h = ""
ElseIf h = "2" Then: h = "B"
ElseIf h = "3" Then: h = "C"
ElseIf h = "4" Then: h = "D"
ElseIf h = "5" Then: h = ""
ElseIf h = "6" Then: h = "F"
ElseIf h = "7" Then: h = "G"
ElseIf h = "8" Then: h = "H"
ElseIf h = "9" Then: h = ""
ElseIf h = "10" Then: h = "J"
ElseIf h = "11" Then: h = "K"
ElseIf h = "12" Then: h = "L"
ElseIf h = "13" Then: h = "M"
ElseIf h = "14" Then: h = "N"
ElseIf h = "15" Then: h = ""
ElseIf h = "16" Then: h = "P"
ElseIf h = "17" Then: h = "Q"
ElseIf h = "18" Then: h = "R"
ElseIf h = "19" Then: h = "S"
ElseIf h = "20" Then: h = "T"
ElseIf h = "21" Then: h = ""
ElseIf h = "22" Then: h = "V"
ElseIf h = "23" Then: h = "W"
ElseIf h = "24" Then: h = "Y"
ElseIf h = "25" Then: h = "X"
ElseIf h = "26" Then: h = "Z"
ElseIf h = "27" Then: h = "0"
ElseIf h = "28" Then: h = "1"
ElseIf h = "29" Then: h = "2"
ElseIf h = "30" Then: h = "3"
ElseIf h = "31" Then: h = "4"
ElseIf h = "32" Then: h = "5"
ElseIf h = "33" Then: h = "6"
ElseIf h = "34" Then: h = "7"
ElseIf h = "35" Then: h = "8"
ElseIf h = "36" Then: h = "9"
ElseIf h = "37" Then: h = " "
End If
txt = a + B + C + D + e + F + G + h

Randomize
i = Int((Val("6") * Rnd) + 1)
If i = "1" Then
i = "E"
ElseIf i = "2" Then: i = "A"
ElseIf i = "3" Then: i = "I"
ElseIf i = "4" Then: i = "O"
ElseIf i = "5" Then: i = "U"
ElseIf i = "6" Then: i = " "
End If
txt = a + B + C + D + e + F + G + h + i

Randomize
j = Int((Val("37") * Rnd) + 1)
If j = "1" Then
j = "A"
ElseIf j = "2" Then: j = "B"
ElseIf j = "3" Then: j = "C"
ElseIf j = "4" Then: j = "D"
ElseIf j = "5" Then: j = "E"
ElseIf j = "6" Then: j = "F"
ElseIf j = "7" Then: j = "G"
ElseIf j = "8" Then: j = "H"
ElseIf j = "9" Then: j = "I"
ElseIf j = "10" Then: j = "J"
ElseIf j = "11" Then: j = "K"
ElseIf j = "12" Then: j = "L"
ElseIf j = "13" Then: j = "M"
ElseIf j = "14" Then: j = "N"
ElseIf j = "15" Then: j = "O"
ElseIf j = "16" Then: j = "P"
ElseIf j = "17" Then: j = "Q"
ElseIf j = "18" Then: j = "R"
ElseIf j = "19" Then: j = "S"
ElseIf j = "20" Then: j = "T"
ElseIf j = "21" Then: j = "U"
ElseIf j = "22" Then: j = "V"
ElseIf j = "23" Then: j = "W"
ElseIf j = "24" Then: j = "Y"
ElseIf j = "25" Then: j = "X"
ElseIf j = "26" Then: j = "Z"
ElseIf j = "27" Then: j = "0"
ElseIf j = "28" Then: j = "1"
ElseIf j = "29" Then: j = "2"
ElseIf j = "30" Then: j = "3"
ElseIf j = "31" Then: j = "4"
ElseIf j = "32" Then: j = "5"
ElseIf j = "33" Then: j = "6"
ElseIf j = "34" Then: j = "7"
ElseIf j = "35" Then: j = "8"
ElseIf j = "36" Then: j = "9"
ElseIf j = "37" Then: j = " "
End If
txt = a + B + C + D + e + F + G + h + j


Randomize
k = Int((Val("37") * Rnd) + 1)
If k = "1" Then
k = "A"
ElseIf k = "2" Then: k = "B"
ElseIf k = "3" Then: k = "C"
ElseIf k = "4" Then: k = "D"
ElseIf k = "5" Then: k = "E"
ElseIf k = "6" Then: k = "F"
ElseIf k = "7" Then: k = "G"
ElseIf k = "8" Then: k = "H"
ElseIf k = "9" Then: k = "I"
ElseIf k = "10" Then: k = "J"
ElseIf k = "11" Then: k = "K"
ElseIf k = "12" Then: k = "L"
ElseIf k = "13" Then: k = "M"
ElseIf k = "14" Then: k = "N"
ElseIf k = "15" Then: k = "O"
ElseIf k = "16" Then: k = "P"
ElseIf k = "17" Then: k = "Q"
ElseIf k = "18" Then: k = "R"
ElseIf k = "19" Then: k = "S"
ElseIf k = "20" Then: k = "T"
ElseIf k = "21" Then: k = "U"
ElseIf k = "22" Then: k = "V"
ElseIf k = "23" Then: k = "W"
ElseIf k = "24" Then: k = "Y"
ElseIf k = "25" Then: k = "X"
ElseIf k = "26" Then: k = "Z"
ElseIf k = "27" Then: k = "0"
ElseIf k = "28" Then: k = "1"
ElseIf k = "29" Then: k = "2"
ElseIf k = "30" Then: k = "3"
ElseIf k = "31" Then: k = "4"
ElseIf k = "32" Then: k = "5"
ElseIf k = "33" Then: k = "6"
ElseIf k = "34" Then: k = "7"
ElseIf k = "35" Then: k = "8"
ElseIf k = "36" Then: k = "9"
ElseIf k = "37" Then: k = ""
End If
txt2 = k
Randomize
L = Int((Val("37") * Rnd) + 1)
If j = "1" Then
L = "A"
ElseIf L = "2" Then: L = "B"
ElseIf L = "3" Then: L = "C"
ElseIf L = "4" Then: L = "D"
ElseIf L = "5" Then: L = "E"
ElseIf L = "6" Then: L = "F"
ElseIf L = "7" Then: L = "G"
ElseIf L = "8" Then: L = "H"
ElseIf L = "9" Then: L = "I"
ElseIf L = "10" Then: L = "j"
ElseIf L = "11" Then: L = "k"
ElseIf L = "12" Then: L = "L"
ElseIf L = "13" Then: L = "M"
ElseIf L = "14" Then: L = "N"
ElseIf L = "15" Then: L = "O"
ElseIf L = "16" Then: L = "P"
ElseIf L = "17" Then: L = "Q"
ElseIf L = "18" Then: L = "R"
ElseIf L = "19" Then: L = "S"
ElseIf L = "20" Then: L = "T"
ElseIf L = "21" Then: L = "U"
ElseIf L = "22" Then: L = "V"
ElseIf L = "23" Then: L = "W"
ElseIf L = "24" Then: L = "Y"
ElseIf L = "25" Then: L = "X"
ElseIf L = "26" Then: L = "Z"
ElseIf L = "27" Then: L = "0"
ElseIf L = "28" Then: L = "1"
ElseIf L = "29" Then: L = "2"
ElseIf L = "30" Then: L = "3"
ElseIf L = "31" Then: L = "4"
ElseIf L = "32" Then: L = "5"
ElseIf L = "33" Then: L = "6"
ElseIf L = "34" Then: L = "7"
ElseIf L = "35" Then: L = "8"
ElseIf L = "36" Then: L = "9"
ElseIf L = "37" Then: L = ""
End If
txt2 = k + L
Randomize
m = Int((Val("37") * Rnd) + 1)
If j = "1" Then
m = "A"
ElseIf m = "2" Then: m = "B"
ElseIf m = "3" Then: m = "C"
ElseIf m = "4" Then: m = "D"
ElseIf m = "5" Then: m = "E"
ElseIf m = "6" Then: m = "F"
ElseIf m = "7" Then: m = "G"
ElseIf m = "8" Then: m = "H"
ElseIf m = "9" Then: m = "I"
ElseIf m = "10" Then: m = "k"
ElseIf m = "11" Then: m = "l"
ElseIf m = "12" Then: m = "m"
ElseIf m = "13" Then: m = "M"
ElseIf m = "14" Then: m = "N"
ElseIf m = "15" Then: m = "O"
ElseIf m = "16" Then: m = "P"
ElseIf m = "17" Then: m = "Q"
ElseIf m = "18" Then: m = "R"
ElseIf m = "19" Then: m = "S"
ElseIf m = "20" Then: m = "T"
ElseIf m = "21" Then: m = "U"
ElseIf m = "22" Then: m = "V"
ElseIf m = "23" Then: m = "W"
ElseIf m = "24" Then: m = "Y"
ElseIf m = "25" Then: m = "X"
ElseIf m = "26" Then: m = "Z"
ElseIf m = "27" Then: m = "0"
ElseIf m = "28" Then: m = "1"
ElseIf m = "29" Then: m = "2"
ElseIf m = "30" Then: m = "3"
ElseIf m = "31" Then: m = "4"
ElseIf m = "32" Then: m = "5"
ElseIf m = "33" Then: m = "6"
ElseIf m = "34" Then: m = "7"
ElseIf m = "35" Then: m = "8"
ElseIf m = "36" Then: m = "9"
ElseIf m = "37" Then: m = ""
End If
txt2 = k + L + m

Randomize
n = Int((Val("37") * Rnd) + 1)
If j = "1" Then
n = "A"
ElseIf n = "2" Then: n = "B"
ElseIf n = "3" Then: n = "C"
ElseIf n = "4" Then: n = "D"
ElseIf n = "5" Then: n = "E"
ElseIf n = "6" Then: n = "F"
ElseIf n = "7" Then: n = "G"
ElseIf n = "8" Then: n = "H"
ElseIf n = "9" Then: n = "I"
ElseIf n = "10" Then: n = "k"
ElseIf n = "11" Then: n = "l"
ElseIf n = "12" Then: n = "m"
ElseIf n = "13" Then: n = "n"
ElseIf n = "14" Then: n = "N"
ElseIf n = "15" Then: n = "O"
ElseIf n = "16" Then: n = "P"
ElseIf n = "17" Then: n = "Q"
ElseIf n = "18" Then: n = "R"
ElseIf n = "19" Then: n = "S"
ElseIf n = "20" Then: n = "T"
ElseIf n = "21" Then: n = "U"
ElseIf n = "22" Then: n = "V"
ElseIf n = "23" Then: n = "W"
ElseIf n = "24" Then: n = "Y"
ElseIf n = "25" Then: n = "X"
ElseIf n = "26" Then: n = "Z"
ElseIf n = "27" Then: n = "0"
ElseIf n = "28" Then: n = "1"
ElseIf n = "29" Then: n = "2"
ElseIf n = "30" Then: n = "3"
ElseIf n = "31" Then: n = "4"
ElseIf n = "32" Then: n = "5"
ElseIf n = "33" Then: n = "6"
ElseIf n = "34" Then: n = "7"
ElseIf n = "35" Then: n = "8"
ElseIf n = "36" Then: n = "9"
ElseIf n = "37" Then: n = ""
End If
txt2 = k + L + m + n

Randomize
o = Int((Val("37") * Rnd) + 1)
If j = "1" Then
o = "A"
ElseIf o = "2" Then: o = "B"
ElseIf o = "3" Then: o = "C"
ElseIf o = "4" Then: o = "D"
ElseIf o = "5" Then: o = "E"
ElseIf o = "6" Then: o = "F"
ElseIf o = "7" Then: o = "G"
ElseIf o = "8" Then: o = "H"
ElseIf o = "9" Then: o = "I"
ElseIf o = "10" Then: o = "k"
ElseIf o = "11" Then: o = "l"
ElseIf o = "12" Then: o = "m"
ElseIf o = "13" Then: o = "M"
ElseIf o = "14" Then: o = "N"
ElseIf o = "15" Then: o = "O"
ElseIf o = "16" Then: o = "P"
ElseIf o = "17" Then: o = "Q"
ElseIf o = "18" Then: o = "R"
ElseIf o = "19" Then: o = "S"
ElseIf o = "20" Then: o = "T"
ElseIf o = "21" Then: o = "U"
ElseIf o = "22" Then: o = "V"
ElseIf o = "23" Then: o = "W"
ElseIf o = "24" Then: o = "Y"
ElseIf o = "25" Then: o = "X"
ElseIf o = "26" Then: o = "Z"
ElseIf o = "27" Then: o = "0"
ElseIf o = "28" Then: o = "1"
ElseIf o = "29" Then: o = "2"
ElseIf o = "30" Then: o = "3"
ElseIf o = "31" Then: o = "4"
ElseIf o = "32" Then: o = "5"
ElseIf o = "33" Then: o = "6"
ElseIf o = "34" Then: o = "7"
ElseIf o = "35" Then: o = "8"
ElseIf o = "36" Then: o = "9"
ElseIf o = "37" Then: o = ""
End If
txt2 = k + L + m + n + o

End Sub

Sub physhgen(txt As TextBox, txt2 As TextBox)
Randomize   ' Seed random number generator.
X = Int(25 * Rnd + 1)
If X = "1" Then X = "DethMet" & y = "WaReZ"
If X = "2" Then X = "Xq311pX" & y = "kewl"
If X = "2" Then X = "Xq311pX" & y = "kewl"
If X = "3" Then X = "Xxo311oxX" & y = "pass"
If X = "4" Then X = "iil Cease" & y = "qwert"
If X = "5" Then X = "texasange1" & y = "horse"
If X = "6" Then X = "MattM350" & y = "dr6zg"
If X = "7" Then X = "Xx4tayxX" & y = "HeHeHe"
If X = "8" Then X = "RickG15" & y = "laser"
If X = "9" Then X = "Codogg69" & y = "SNOOP"
If X = "10" Then X = "NastyNoz16" & y = "Thugluve"
If X = "11" Then X = "MuscleBal1" & y = "Fila"
If X = "12" Then X = "PITBUL5" & y = "doodle"
If X = "14" Then X = "Topgun1983" & y = "4252"
If X = "15" Then X = "MMata39751" & y = "hello"
If X = "16" Then X = "THEFIRM SH" & y = "49311"
If X = "17" Then X = "AAmodei" & y = "rainfall"
If X = "18" Then X = "VileBillC" & y = "carpet2"
If X = "19" Then X = "R44js" & y = "X-MEN"
If X = "16" Then X = "CHEO8055" & y = "booo"
If X = "17" Then X = "TotalCaos5" & y = "chaos"
If X = "18" Then X = "DethMet" & y = "daemon"
If X = "19" Then X = "ILLumina01" & y = "banana"
If X = "20" Then X = "Crshovr921" & y = "0516"
If X = "21" Then X = "WiPe IcE" & y = "guild"
If X = "22" Then X = "DMath37689" & y = "trigger"
If X = "23" Then X = "Zantaclwz3" & y = "301644"
If X = "24" Then X = "ID HiTMaN" & y = "id2id2"
If X = "25" Then X = "Twiggy8989" & y = "metal"
txt.Text = "SN:  " & X
txt2.Text = "PW:  " & y

End Sub

Sub playwav(file)
SoundName$ = file
   wflags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndplaysound(SoundName$, wflags%)
End Sub

Sub preferences()
Call runmenua(4, 2)
X = WaitForWindow("Preferences", "AOL Child")
D% = Findchildbyclass(X, "_AOL_Icon")
Click D%
y = WaitForWindow("Chat Preferences", "_AOL_Modal")
B% = findchildbytitle(y, "Cancel")
Click B%
killwin X
End Sub

Function r_backwards(strin As TextBox)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
r_backwards = newsent$

End Function

Function r_elite(strin As String)
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

Function r_hacker(strin As TextBox)
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

Function r_same(strr As String)
Let r_same = Trim(strr)

End Function

Function r_spaced(strin As TextBox)
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

Function rcase(hehe$)
Do
flet$ = Left$(hehe$, 1)
If UCase$(flet$) = flet$ Then
joe$ = LCase$(flet$)
Else: joe$ = UCase$(flet$)
End If
bah$ = bah$ & joe$
hehe$ = Mid$(hehe$, 2)
DoEvents
Loop Until Len(hehe$) = 0
rcase = bah$

End Function

Function ReadINI(AppName, KeyName, Filename As String) As String
'Example: text4.text = ReadINI("DaProggy", "Lamers Name", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), Filename))

End Function

Sub renamehost(sn As String)
X = aolversion()
If X = "C:\aol30" Then
Open "C:\aol30\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , sn

Close #1
ElseIf X = "C:\aol25" Then
Open "C:\aol25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , sn

Close #1
End If
End Sub

Sub replace(windo%, gud$, word$)
mmm = GetWinText(windo%)
If InStr(mmm, gud$) Then
heh$ = Left$(mmm, InStr(mmm, gud$) - 1)
duh$ = Mid$(mmm, InStr(mmm, gud$) + 1)
send windo%, heh$ & word$ & duh$
End If

End Sub

Function ReplaceText(Text, charfind, charchange)
Dim replace1
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For replace1 = 1 To Len(Text)
thechar$ = Mid(Text, replace1, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next replace1

ReplaceText = thechars$

End Function

Sub resetnew(sn As String, tru_sn As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(sn)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
If Len(sn) < 7 Then MsgBox ("The screen name hasta be at least 7 characters long :)"): Exit Sub
tru_sn = tru_sn + String$(Len(sn) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
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

Sub Run(ByVal menuname As String)
Dim mhWnd As Integer
Dim MnuCnt As Integer
Dim Cnt As Integer
Dim SmhWnd As Integer
Dim SMnuCnt As Integer
Dim SCnt As Integer
Dim lint As Integer
Dim lpString As String
Dim AOLx
AOLx = findwindow%("AOL Frame25", 0&)
mhWnd = GetMenu%(AOLx)
MnuCnt = GetMenuItemCount%(mhWnd)
For Cnt = 0 To MnuCnt - 1
    SmhWnd = GetSubMenu%(mhWnd, Cnt)
    SMnuCnt = GetMenuItemCount%(SmhWnd)
    For SCnt = 0 To SMnuCnt - 1
    lint = 30
    lpString = Space(30)
    X = GetMenuString%(SmhWnd, SCnt, lpString, lint, MF_BYPOSITION)
    
    If InStr(UCase$(lpString), UCase$(menuname)) Then
        GMnID = GetMenuItemID%(SmhWnd, SCnt)
        Ret = SendMessageByNum(AOLx, WM_COMMAND, GMnID, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt

End Sub

Function runmenu(TopMenuPos As Integer, PopupPos As Integer)
Dim i, mnuhWnd, submnu, AOL, Menuid As Integer
Dim lParam As Long
Dim blah As String * 500
AOL = findwindow("ThunderMain", 0&)
mdi = Findchildbyclass(AOL, "wndclass_desked_gsk")
mnu = GetMenu(AOL)
X = GetMenuString(mnu, 0, blah, Len(blah) + 1, &H400)
If InStr(blah, "*") Then TopMenuPos = TopMenuPos + 1
submnu = GetSubMenu%(mnu, TopMenuPos)
Menuid = GetMenuItemID%(submnu, PopupPos)
i = SendMessageByNum(AOL, &H111, Menuid, 0)
End Function

Function runmenu2(Main_Prog As String, Top_Position As String, Menu_String As String)
Dim Top_Position_Num As Integer
Dim Buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim AOL As Integer
Dim Menu_Handle As Integer
End Function

Function runmenu3(TopMenuPos As Integer, PopupPos As Integer)
Dim i, mnuhWnd, submnu, Menuid  As Integer
Dim lParam As Long
Const MF_BYCOMMAND = &H0
mnuhWnd = GetMenu%(findwindow(0&, "America  Online"))
submnu = GetSubMenu%(mnuhWnd, TopMenuPos)
Menuid = GetMenuItemID%(submnu, PopupPos)
lParam = CLng(0) * &H10000 Or MF_BYCOMMAND
i = SendMessageByNum(findwindow(0&, "America  Online"), WM_COMMAND, Menuid, 0&)
End Function

Sub runmenua(horz, vert)
'Runs the specified AOL Menu (Horizonatl,Verticle)
'Each Position starts at 0 not 1

Dim F, gi, sm, m, a As Integer
a = findwindow("AOL Frame25", 0&)
m = GetMenu(a)
sm = GetSubMenu(m, horz)
gi = GetMenuItemID(sm, vert)
F = SendMessageByNum(a, WM_COMMAND, gi, 0)

End Sub

Sub runmenubystring(Application, StringSearch)
ToSearch% = GetMenu(Application)
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
RunTheMenu% = sendmessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Sub RunToolBar(IcoNum As Integer)
AOL% = findwindow("AOL Frame25", 0&)
ToolBar% = Findchildbyclass(AOL%, "AOL Toolbar")
T1% = Findchildbyclass(ToolBar%, "_AOL_Icon")
T2% = getwindow(T1%, GW_HWNDNEXT)
T3% = getwindow(T2%, GW_HWNDNEXT)
T31% = getwindow(T3%, GW_HWNDNEXT)
T4% = getwindow(T31%, GW_HWNDNEXT)
T5% = getwindow(T4%, GW_HWNDNEXT)
T6% = getwindow(T5%, GW_HWNDNEXT)
T7% = getwindow(T6%, GW_HWNDNEXT)
T8% = getwindow(T7%, GW_HWNDNEXT)
T9% = getwindow(T8%, GW_HWNDNEXT)
T10% = getwindow(T9%, GW_HWNDNEXT)
T11% = getwindow(T10%, GW_HWNDNEXT)
T12% = getwindow(T11%, GW_HWNDNEXT)
T13% = getwindow(T12%, GW_HWNDNEXT)
T14% = getwindow(T13%, GW_HWNDNEXT)
T15% = getwindow(T14%, GW_HWNDNEXT)
T16% = getwindow(T15%, GW_HWNDNEXT)
T17% = getwindow(T16%, GW_HWNDNEXT)
T18% = getwindow(T17%, GW_HWNDNEXT)
T19% = getwindow(T18%, GW_HWNDNEXT)
If IcoNum = 1 Then IcoNum2% = T1%
If IcoNum = 2 Then IcoNum2% = T2%
If IcoNum = 3 Then IcoNum2% = T3%
If IcoNum = 4 Then IcoNum2% = T4%
If IcoNum = 5 Then IcoNum2% = T5%
If IcoNum = 6 Then IcoNum2% = T6%
If IcoNum = 7 Then IcoNum2% = T7%
If IcoNum = 8 Then IcoNum2% = T8%
If IcoNum = 9 Then IcoNum2% = T9%
If IcoNum = 10 Then IcoNum2% = T10%
If IcoNum = 11 Then IcoNum2% = T11%
If IcoNum = 12 Then IcoNum2% = T12%
If IcoNum = 13 Then IcoNum2% = T13%
If IcoNum = 14 Then IcoNum2% = T14%
If IcoNum = 15 Then IcoNum2% = T15%
If IcoNum = 16 Then IcoNum2% = T16%
If IcoNum = 17 Then IcoNum2% = T17%
If IcoNum = 18 Then IcoNum2% = T18%
If IcoNum = 19 Then IcoNum2% = T19%
DoEvents
Click (IcoNum2%)
End Sub

Sub send(chatedit, sill$)
sndtext = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub Sendafuk(sendstring$)
Dim AOL%, aoledit%
AOL% = findwindow("AOL FRAME25", 0&)
If AOL% = 0 Then
Exit Sub
End If
aoledit% = Findchildbyclass(AOL%, "_AOL_EDIT")
If aoledit% = 0 Then
MsgBox "You have to be in a chatroom dumbfuck!", 56, "Dumbass!"
Exit Sub
End If
X = SendMessageByString(aoledit%, WM_SETTEXT, 0&, sendstring$)
DoEvents
X = SendMessageByNum(aoledit%, WM_CHAR, 13, 0&)
DoEvents
End Sub

Function sendchatroomtext(ByVal sendstring As String) As Integer
    'Find Send BuTToN
    zzzzz = findwindow("aol frame25", 0&)
    ChatWnd% = findchatwnd()
    EdtBoxAts% = Findchildbyclass(ChatWnd%, "_AOL_EDIT")
    SendBut% = getnextwindow(EdtBoxAts%, GW_HWNDNEXT)
    VC = SendMessageByString(EdtBoxAts%, WM_SETTEXT, 0, sendstring)
    DoEvents
    X = setfocusapi(SendBut%)
    X = SetActiveWindow(SendBut%)
    Call aolclick(SendBut%)
    DoEvents
    X = setfocusapi(zzzzz)
    'SendKeys "{ENTER}"

    'Tell the procedure from which this function was called that text was sent
    sendchatroomtext = True
    

End Function

Sub SENDCHATTXT(win, sill$)
SendNow% = SendMessageByString(win, WM_SETTEXT, 0, sill$)
X% = SendMessageByNum(win, WM_CHAR, 13, 0)
timeout (0.001)
End Sub

Sub Sendit(sill$)
AOL% = findwindow("AOL Frame25", 0&)
chatlist% = Findchildbyclass(AOL%, "_AOL_Edit")
SendNow% = SendMessageByString(chatlist%, WM_SETTEXT, 0, sill$)
X% = SendMessageByNum(chatlist%, WM_CHAR, 13, 0)
timeout (0.001)
End Sub

Sub ADD_AOL_LB(itm As String, Lst As ListBox)
'Add a list of names to a VB ListBox
'This is usually called by another one of my functions
Dim xx
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

Sub addaolmenu()
    'first find the handle of AOL
    AOL% = findwindow("AOL Frame25", 0&)
    
    'then the handle of AOL's menu
    aolmenu% = GetMenu(AOL%)

    'then we create 3 popup menus, one will be the main one added to AOL, the other two will be submenus of the main menu we add
    mainmenu% = CreatePopupMenu()
    submenu1% = CreatePopupMenu()
    submenu2% = CreatePopupMenu()
    
    'this adds 3 menu items to the first submenu
    X% = AppendMenu(submenu1%, MF_ENABLED Or MF_STRING, 56, "Menu Item &1")
    X% = AppendMenu(submenu1%, MF_ENABLED Or MF_STRING, 57, "Menu Item &2")
    X% = AppendMenu(submenu1%, MF_ENABLED Or MF_STRING, 58, "Menu Item &3")

    'this adds 3 menu items to the second submenu
    X% = AppendMenu(submenu2%, MF_ENABLED Or MF_STRING, 59, "Menu Item &4")
    X% = AppendMenu(submenu2%, MF_ENABLED Or MF_STRING, 60, "Menu Item &5")
    X% = AppendMenu(submenu2%, MF_ENABLED Or MF_STRING, 61, "Menu Item &6")
    
    'this "appends" the submenus onto our mainmenu
    X% = AppendMenu(mainmenu%, MF_STRING Or MF_POPUP, submenu1%, "Sub Menu &1")
    X% = AppendMenu(mainmenu%, MF_STRING Or MF_POPUP, submenu2%, "Sub Menu &2")
    
    'this adds a seperator onto our mainmenu
    X% = AppendMenu(mainmenu%, MF_SEPARATOR, 0&, 0&)
    
    'this adds there menu items to our mainmenu
    X% = AppendMenu(mainmenu%, MF_ENABLED Or MF_STRING, 62, "Menu Item &7")
    X% = AppendMenu(mainmenu%, MF_ENABLED Or MF_STRING, 63, "Menu Item &8")
    X% = AppendMenu(mainmenu%, MF_ENABLED Or MF_STRING, 64, "Menu Item &9")
    
    'this adds our main menu to aol's menu bar
    X% = AppendMenu(aolmenu%, MF_STRING Or MF_POPUP, mainmenu%, "&Main Menu")
    
    'this redraws aol's menu bar
    DrawMenuBar (AOL%)
    
    'this sets VBMsg to subclass AOL, allowing us to make our menu items do something

    MsgBox "The menu has been added"

End Sub

Sub addcmbchk(itm As String, Lst As ComboBox)
If Lst.ListCount = 0 Then Lst.AddItem itm: Exit Sub
Do Until xx = (Lst.ListCount)
Let diss_itm$ = Lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub

Sub addroom(Lst As ListBox)
'This calls a function in 311.dll that retreives the names
'from the AOL listbox.
'PLEASE NOTE THE FOLLOWING:
'1)  I don't support this dll..its hacked and illegal
'2)  This only works on 16 bit versions of AOL
'3)  Its a good idea to bring the chat room to the top
'    of the AOL client before doing this.  Sometimes it
'    gets text from other AOL listboxes


For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addr
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)

ADD_AOL_LB namez$, Lst
Next Index%
end_addr:

End Sub

Sub addroomwithoutme(Lst As ListBox)
Call addroom(Lst)
X = getuser()
For i = 0 To Lst.ListCount - 1
dood = Lst.List(i)
w = InStr(dood, X)
y = Lst.ListIndex
If w <> 0 Then
Lst.RemoveItem (i)
End If
Next i
End Sub

Sub addsta(Lst As ListBox)
AOL = findwindow%("AOL Frame25", "America  Online")
Happy% = findchildbytitle%(AOL, "Welcome,")
joy$ = GetWinText((Happy%))
If InStr(joy$, "!") = 0 Then Exit Sub
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
For Index% = 0 To 25 'This will loop until everyone's added or it will stop at 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then Exit For
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
If namez$ = yoursn Then GoTo bo2
If InStr(LCase$(namez$), "tos") Then GoTo bo2
If InStr(LCase$(namez$), "guide") Then GoTo bo2
If InStr(LCase$(namez$), "host") Then GoTo bo2
Lst.AddItem (namez$)
bo2:
Next Index%
End Sub

Function AlterAOL(txt As String)
Dim Weq, Ap As Integer
AlterAOL = Ap
End Function

Function AOhWnd()
AOL% = findwindow("AOL Frame25", 0&)

End Function

Sub aolclick(E1 As Integer)
If proG_STAT$ = "OFF" Then
Exit Sub
End If
do_wn = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
timeout 0.008
u_p = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
End Sub

Function AOLCountMail()
themail% = Findchildbyclass(AOLMDI(), "AOL Child")
thetree% = Findchildbyclass(themail%, "_AOL_Tree")
AOLCountMail = sendmessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Sub AOLCursor()
Call runmenubystring(AOLWindow(), "&About America Online")
Do: DoEvents
C% = findwindow("_AOL_Modal", 0&)
Loop Until a% <> 0
D% = sendmessage(C%, WM_CLOSE, 0, 0)

End Sub

Function AOLFindRoom()
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
firs% = getwindow(mdi%, 5)
listers% = Findchildbyclass(firs%, "_AOL_Edit")
listere% = Findchildbyclass(firs%, "_AOL_View")
listerb% = Findchildbyclass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone

firs% = getwindow(mdi%, GW_CHILD)
While firs%
firs% = getwindow(firs%, 2)
listers% = Findchildbyclass(firs%, "_AOL_Edit")
listere% = Findchildbyclass(firs%, "_AOL_View")
listerb% = Findchildbyclass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
firs% = getwindow(mdi%, 5)
listers% = Findchildbyclass(firs%, "_AOL_Edit")
listere% = Findchildbyclass(firs%, "_AOL_View")
listerb% = Findchildbyclass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
Wend

bone:
Room% = firs%
AOLFindRoom = Room%
End Function

Function AOLGetChat()
childs% = AOLFindRoom()
child = Findchildbyclass(childs%, "_AOL_View")


GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$
AOLGetChat = theview$
End Function

Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

AOLGetText = TrimSpace$
End Function

Function AOLGetUser()
On Error Resume Next
AOL% = findwindow("AOL Frame25", "America  Online")
mdi% = Findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowtextlength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function

Function aolhwnd()
a = findwindow("AOL Frame25", 0&)
aolhwnd = a
End Function

Sub AOLInstantMessage(Person$, tt$)
okw = findwindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
Run "Send an Instant Message"
Do
AOL = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(AOL, "Send Instant Message")
txt% = Findchildbyclass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
txt% = Findchildbyclass(bah, "_AOL_Edit")
Do
rich% = Findchildbyclass(bah, "RICHCNTL")
bahqw% = findchildbytitle(bah, "Send")
DoEvents
timeout (0.00001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
send txt%, Person$
send rich%, tt$
timeout (0.00001)
getnum rich%, 1
Click rich%
Else
send txt%, Person$
getnum txt%, 1
send txt%, tt$
getnum txt%, 1
Click txt%
End If
timeout (0.00001)
X = SendMessageByNum(bah, WM_CLOSE, 0, 0)
End Sub

Function AOLIsOnline()
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
welcome% = findchildbytitle(mdi%, "Welcome, ")
If welcome% = 0 Then
MsgBox "Please sign on before using this feature.", 64, "Online"
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = Findchildbyclass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = lastline
End Function

Function AOLMDI()
AOL% = findwindow("AOL Frame25", 0&)
AOLMDI = Findchildbyclass(AOL%, "MDIClient")
End Function

Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
keywor ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub

Sub AOLOpenMail(which)
If which = 1 Then
Call AOLRunMenuByString("Read &New Mail")
End If

If which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call AOLRunMenuByString("Check Mail You've &Sent")
End If

End Sub

Sub AOLRespondIM(message)
IM% = findchildbytitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo z
IM% = findchildbytitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo z
Exit Sub
z:
e = Findchildbyclass(IM%, "RICHCNTL")

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
Click (e)
End Sub

Function AOLRoomCount()
thechild% = AOLFindRoom()
lister% = Findchildbyclass(thechild%, "_AOL_Listbox")

getcount = sendmessage(lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLRunMenuByString(stringer As String)
Call runmenubystring(AOLWindow(), stringer)
End Sub

Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub AOLSignoff()
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
X = runmenu(7, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
pfc% = findchildbytitle(AOL%, "Sign Off?")
If pfc% <> 0 Then
Icon1% = Findchildbyclass(pfc%, "_AOL_Icon")
Icon1% = getwindow(Icon1%, 2)
Icon1% = getwindow(Icon1%, 2)
Icon1% = getwindow(Icon1%, 2)
Icon1% = getwindow(Icon1%, 2)
Icon1% = getwindow(Icon1%, 2)
clickicon% = sendmessage(Icon1%, WM_LBUTTONDOWN, 0, 0&)
clickicon% = sendmessage(Icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function aolversion()
AOL% = findwindow("AOL Frame25", 0&)
hMenu% = GetMenu(AOL%)

submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
aolversion = "C:\aol30"
Else
aolversion = "C:\aol25"
End If
End Function

Sub AOLWaitMail()
mailwin% = GettopWindow(AOLMDI())
aoltree% = Findchildbyclass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
timeout 3
secondcount = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub

Function AOLWindow()
AOL% = findwindow("AOL Frame25", 0&)
AOLWindow = AOL%
End Function

Sub AutoGreetPref()
a% = findwindow("AOL Frame25", 0&)
ssa% = aolversion()
If ssa% = 25 Then wh$ = "Set Preferences" + Chr$(9) + "Ctrl+="
If ssa% = 3 Then wh$ = "Preferences"
B% = runmenu2("AOL Frame25", "Mem&bers", wh$)
DoEvents
Do
DoEvents
C% = findchildbytitle(a%, "Preferences")
DoEvents
Loop Until C% <> 0
DoEvents
If ssa% = 25 Then D% = getaolwinb(C%, "_AOL_Icon", 3)
If ssa% = 3 Then D% = getaolwinb(C%, "_AOL_ICON", 1)
DoEvents
aolclick (D%)
DoEvents
Do
DoEvents
a% = findwindow("AOL Frame25", 0&)
e% = findwindow("_AOL_Modal", "Chat Preferences")
DoEvents
Loop Until e% <> 0
DoEvents
F% = findchildbytitle(e%, "Notify me when members arrive")
DoEvents
DoEvents
h% = SendMessageByNum(F%, BM_SETCHECK, True, 0&)
DoEvents
j% = findchildbytitle(e%, "OK")

DoEvents
aolclick (j%)
DoEvents
k% = SendMessageByNum(C%, WM_CLOSE, 0, 0&)
End Sub

Function bak(ward$)
Do
heh$ = Right$(ward$, 1)
mog$ = mog$ & heh$
bad$ = Left$(ward$, Len(ward$) - 1)
ward$ = bad$
DoEvents
Loop Until Len(ward$) = 0
bak = mog$

End Function

Sub BlankIM(who$)
Call AOLInstantMessage((who$), ("<a href=GenOziDe>"))
End Sub

Sub blankimbomb(numofbom As TextBox, whotobom As TextBox)
If numofbom = "0" Then
Exit Sub
End If
If numofbom < 0 Then
Exit Sub
End If
Do
Call AOLInstantMessage((whotobom), (" "))
numofbom = Str(Val(numofbom - 1))
Loop Until numofbom = 0
Exit Sub
End Sub

Sub bustinpriv(Room$)
electric1:
If proG_STAT$ = "OFF" Then
Exit Sub
End If

Let num_try = num_try + 1
a = findwindow("AOL Frame25", 0&)

Do
DoEvents
fcr = findchildbytitle(a, "Keyword")
DoEvents
Loop Until fcr <> 0

E1 = Findchildbyclass(fcr, "_AOL_Icon")
i1 = Findchildbyclass(fcr, "_AOL_Edit")
send_to = SendMessageByString(i1, WM_SETTEXT, 0, "aol://2719:2-2-" & Room$)
tta = SendMessageByNum(E1, WM_LBUTTONDOWN, 0, 0&)
ttr = SendMessageByNum(E1, WM_LBUTTONUP, 0, 0&)
waitforok
If OK_note = "NO" Then
Exit Sub
End If
timeout 0.2
GoTo electric1

End Sub

Sub centerform(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub

Sub Changefocus()
Dim aa, ab As Integer
aa = getfocus()
Do
  DoEvents
  ab = getfocus()
Loop While aa = ab
End Sub

Sub changewav(wav As String)
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub

Function char(What$, a)
u$ = Mid(What$, Val(a))
char = Left(u$, 1)
End Function

Function ChatRoomName()
ChatRoomName = WindowCaption(findchatwnd())

End Function

Function chattext()
Do
AOL = findwindow("AOL Frame25", 0&)
View% = Findchildbyclass(AOL, "_AOL_View")
bag = GetWinText(View%)
Do
AOL = findwindow("AOL Frame25", 0&)
View% = Findchildbyclass(AOL, "_AOL_View")
bad = GetWinText(View%)
DoEvents
Loop Until bad <> bag
chattext = Mid(bad, InStr(bad, bag) + Len(bag) - 1)
DoEvents
Loop

End Function

Function ChatView() As Integer
Chat% = AOLFindRoom()
View% = Findchildbyclass(Chat%, "_AOL_View")
ChatView = View%
End Function

Sub checkifalive(who$)
Do
AOL% = findwindow("AOL Frame25", 0&)
heh = findchildbytitle(AOL%, "Compose Mail")
send heh, "Compose Mail"
DoEvents
Loop Until heh = 0
AOL% = findwindow("AOL Frame25", 0&)
Run "Compose"
Do
X = DoEvents()
chatlist% = findchildbytitle(AOL%, "Compose Mail")
chatedit% = Findchildbyclass(chatlist%, "_AOL_Edit")
uh = ShowWindow(chatlist%, SW_HIDE)
Loop Until chatedit% <> 0
chatwin% = getparent(chatlist%)
button% = Findchildbyclass(chatlist%, "_AOL_Icon")
ChatVew% = Findchildbyclass(chatlist%, "_AOL_View")
chatedit% = Findchildbyclass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, who$ & ", '")
blah% = getwindow(chatedit%, GW_HWNDNEXT)
good% = getwindow(blah%, GW_HWNDNEXT)
bad% = getwindow(good%, GW_HWNDNEXT)
Sad% = getwindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, "TOS check")
Nad% = getwindow(Sad%, GW_HWNDNEXT)
Wad% = getwindow(Nad%, GW_HWNDNEXT)
Dad% = getwindow(Wad%, GW_HWNDNEXT)
fad% = getwindow(Dad%, GW_HWNDNEXT)
Qad% = getwindow(fad%, GW_HWNDNEXT)
Ead% = getwindow(Qad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Dad%, WM_SETTEXT, 0, "TOS Check")
rich = Findchildbyclass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, "TOS Check")
timeout (0.1)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
timeout (0.1)
Do
eroro = findchildbytitle(AOL%, "Error")
timeout (0.001)
Loop Until eroro <> 0
winwo% = Findchildbyclass(eroro, "_AOL_View")
teee = GetWinText(winwo%)

If InStr(LCase$(teee), LCase$(who$)) Then
checkifalive1 = 1
End If
If InStr(LCase$(teee), LCase$("mailbox")) Then
checkifalive1 = 0
End If
killwin (eroro)
killwin (chatlist%)
Do
AOL% = findwindow("AOL Frame25", 0&)
heh = findchildbytitle(AOL%, "Compose Mail")
send heh, "Compose Mail"
DoEvents
Loop Until heh = 0

End Sub

Sub Click(button%)
SendNow% = SendMessageByNum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub ClickAOLMenu(Menu_String As String, Top_Position As String)
Dim Top_Position_Num As Integer
Dim Buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim AOL As Integer
Dim Menu_Handle As Integer


Top_Position_Num = -1
AOL% = findwindow("AOL Frame25", 0&)
Menu_Handle = GetMenu(AOL%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    Buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, Buffer$, Len(Top_Position) + 1, &H400)
    Trim_Buffer = Trimnull(Buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    Buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, Buffer$, Len(Menu_String) + 1, &H400)
    Trim_Buffer = Trimnull(Buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(AOL, &H111, Get_ID%, 0&)

End Sub

Sub ClickKW25()
AOL% = findwindow("AOL Frame25", 0&)
tool% = Findchildbyclass(AOL%, "Aol Toolbar")
Icon1% = Findchildbyclass(tool%, "_Aol_Icon")
Icon2% = getwindow(Icon1%, 2)
Icon3% = getwindow(Icon2%, 2)
Icon4% = getwindow(Icon3%, 2)
Icon5% = getwindow(Icon4%, 2)
Icon6% = getwindow(Icon5%, 2)
Icon7% = getwindow(Icon6%, 2)
Icon8% = getwindow(Icon7%, 2)
Icon9% = getwindow(Icon8%, 2)
Icon10% = getwindow(Icon9%, 2)
Icon11% = getwindow(Icon10%, 2)
Icon12% = getwindow(Icon11%, 2)
Icon13% = getwindow(Icon12%, 2)
z% = sendmessage(Icon13%, WM_LBUTTONDOWN, 0&, 0&)
Q% = sendmessage(Icon13%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub ClickKW30()
AOL% = findwindow("AOL Frame25", 0&)
tool% = Findchildbyclass(AOL%, "Aol Toolbar")
Icon1% = Findchildbyclass(tool%, "_Aol_Icon")
Icon2% = getwindow(Icon1%, 2)
Icon3% = getwindow(Icon2%, 2)
Icon4% = getwindow(Icon3%, 2)
Icon5% = getwindow(Icon4%, 2)
Icon6% = getwindow(Icon5%, 2)
Icon7% = getwindow(Icon6%, 2)
Icon8% = getwindow(Icon7%, 2)
Icon9% = getwindow(Icon8%, 2)
Icon10% = getwindow(Icon9%, 2)
Icon11% = getwindow(Icon10%, 2)
Icon12% = getwindow(Icon11%, 2)
Icon13% = getwindow(Icon12%, 2)
Icon14% = getwindow(Icon13%, 2)
Icon15% = getwindow(Icon14%, 2)
Icon16% = getwindow(Icon15%, 2)
Icon17% = getwindow(Icon16%, 2)
Icon18% = getwindow(Icon17%, 2)
Icon19% = getwindow(Icon18%, 2)
z% = sendmessage(Icon18%, WM_LBUTTONDOWN, 0&, 0&)
Q% = sendmessage(Icon18%, WM_LBUTTONUP, 0&, 0&)
End Sub

Function countmail()
Call runmenua(3, 2)
Do
DoEvents
AO% = findwindow(0&, "America  Online")
Hand% = Findchildbyclass(AO%, "_AOL_TREE")
DoEvents
Loop Until Hand% <> 0
Buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
If Buffer > 1 And Buffer <> 550 Then
MsgBox "You have " & Buffer & " messages in your Mailbox...", 0
End If
If Buffer = 1 Then
MsgBox "You have " & Buffer & " message in your Mailbox...", 0
End If
If Buffer < 1 Then
MsgBox "You have no messages in your Mailbox.", 0
End If
If Buffer = 550 Then
MsgBox "Holy ShiT! Yer BoX Iz FuLL!"
End If

End Function

Sub countuermailz()
MsgBox "Wait For Hourglass To Clear..... And Click OK"
AO% = findwindow("AOL Frame25", 0&)
arf = findchildbytitle(AO%, "New Mail")
Hand% = Findchildbyclass(arf, "_AOL_TREE")
mails = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
If mails > 1 Then
MsgBox "You have " & mails & " messages in your E-Mailbox."
End If
If mails = 1 Then
MsgBox "You have one message in your E-Mailbox."
End If
If mails < 1 Then
MsgBox "You have zero messages in your E-Mailbox"
End If

End Sub

Sub deletmenu()
    AOL% = findwindow("AOL Frame25", 0&)
    aolmenu% = GetMenu(AOL%)
    a = 0
    Do
        X% = DoEvents()
        check% = GetSubMenu(aolmenu%, a)
        If check% <> 0 Then a = a + 1 Else Exit Do
    Loop
    For B = 0 To a - 1
        Buff$ = String(255, 0)
        X% = GetMenuString(aolmenu%, B, Buff$, 255, MF_BYPOSITION)
        Buff$ = Left(Buff$, X%)
        If Buff$ = "&FATE Infinity" Then
            X% = DeleteMenu(aolmenu%, B, MF_BYPOSITION)
            B = B - 1
            a = a - 1
        End If
    Next B

    DrawMenuBar (AOL%)
AO% = findwindow("aol Frame25", 0&)
y% = SendMessageByString(AO%, WM_SETTEXT, 0, "America  Online")
End
End Sub

Sub delitem(Lst As ListBox, item$)
Do
If LCase$(Lst.List(a)) = LCase$(item$) Then Lst.RemoveItem (a)
a = 1 + a
Loop Until a >= Lst.ListCount

End Sub

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
char2 = char1$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(char1$, 1, Len(char1$) - 1)
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
char2 = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function

Function Disc(ByVal Marbro As String) As String
On Error Resume Next
Disc = Left$(Marbro$, InStr(Marbro$, Chr$(0)) - 1)
End Function

Function dupekill()
Num = countmail()
DELNUM% = 0
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")
FNMB% = findchildbytitle(mdi%, "New Mail")

If FNMB% = 0 Then
RunToolBar (1)
timeout (0.4)
Do: DoEvents
    NMB% = findchildbytitle(mdi%, "New Mail")
    timeout (0.1)
Loop Until NMB% <> 0
waitmail
End If

Do: DoEvents
LSTTXT$ = ","
DELTXT$ = ","
btnDEL% = findchildbytitle(findchildbytitle(mdi%, "New Mail"), "Delete")
If countmail() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
List% = Findchildbyclass(findchildbytitle(mdi%, "New Mail"), "_AOL_Tree")
For i = 0 To countmail() - 1
Ln = sendmessage(List%, LB_GETTEXTLEN, i, 0)
If Ln = -1 And i >= countmail() Then
    Exit For
ElseIf Ln = -1 And i <= countmail() Then
    MAILTXT$ = String$(60, 0)
Else
    MAILTXT$ = String$(Ln, 0)
End If
GTTXT = SendMessageByString(List%, LB_GETTEXT, i, MAILTXT$)
MAILTXT$ = Right$(MAILTXT$, Len(MAILTXT$) - InStr(InStr(MAILTXT$, Chr$(9)) + 1, MAILTXT$, Chr$(9)))
If InStr(LSTTXT$, "," & MAILTXT$ & ",") And InStr(DELTXT$, "," & MAILTXT$ & ",") = 0 Then
            X = sendmessage(List%, LB_SETCURSEL, i, 0)
            Call Click(btnDEL%)
            DELNUM% = DELNUM% + 1
            Num = Num - 1
            i = i - 1
            DELTXT$ = DELTXT$ + MAILTXT$ + ","
Else
LSTTXT$ = LSTTXT$ + MAILTXT$ + ","
End If
Next i
Loop Until Len(DELTXT$) = 1
MsgBox "There were " & DELNUM% & " duplicate mails deleted.", 12, "Dupe Count"
dupekill = DELNUM%
End Function

Sub EchoBot(txt1 As TextBox)
Dim Lst As String
Dim AOL As Integer
Dim cht As Integer
Dim txt As String
Dim nws As String
Dim who As String
Dim wht$
Dim R As Integer
Dim e As Integer
Dim m As Integer
AOL = findwindow("AOL Frame25", 0&)
cht = Findchildbyclass(AOL, "_AOL_View")
txt = WinTxt(cht)
If Lst = "" Then Lst = txt
If txt = Lst Then Exit Sub
Lst = txt
nws = GetLast(txt)
who = Mid(nws, 2, InStr(nws, ":") - 2)
wht$ = Mid(nws, Len(who) + 4, Len(nws) - Len(who))
If LCase(Trim(trm(who))) = LCase(Trim(trm(txt1.Text))) Then
R = getparent(cht)
e = Findchildbyclass(R, "_AOL_Edit")
sendtext wht$
m = sendmessage(e, &H102, 13, 0&)
End If
End Sub

Sub sendtext(sill$)
AOL% = findwindow("AOL Frame25", 0&)
chatlist% = Findchildbyclass(AOL%, "_AOL_Listbox")
chatwin% = getparent(chatlist%)
ChatVew% = Findchildbyclass(chatwin%, "_AOL_View")
chatedit% = Findchildbyclass(chatwin%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, sill$)
SendNow% = SendMessageByNum(chatedit%, WM_CHAR, &HD, 0)
'Call Click(ChatSend%)
timeout 0.000001
End Sub

Sub SetBackPre()
Call runmenubystring(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = findchildbytitle(AOLMDI(), "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = getwindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
Click (mailbut%)

Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
aolcloses% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = sendmessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = sendmessage(aolconfirm%, BM_SETCHECK, 1, 0)

Click (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = sendmessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function SetChildFocus(child)
setchild% = setfocusapi(child)
End Function

Sub setpreference()
Call runmenubystring(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = findchildbytitle(AOLMDI(), "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = getwindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
Click (mailbut%)

Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
aolcloses% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = sendmessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = sendmessage(aolconfirm%, BM_SETCHECK, 0, 0)

Click (aolOK%)
Do: DoEvents
aolmod% = findwindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = sendmessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Sub ShowAOL()
a = findwindow("aol frame25", 0&)
B = ShowWindow(a, SW_SHOW)
End Sub

Sub ShowAOLToolbar()
a = findwindow("aol frame25", 0&)
B = Findchildbyclass(a, "AOL Toolbar")
C = ShowWindow(B, SW_SHOW)
End Sub

Sub showaolwins()
fc = Findchildbyclass(aolhwnd(), "AOL Child")
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

Sub ShowBuddyList()
a = findwindow("aol frame25", 0&)
B = Findchildbyclass(a, "MDIClient")
C = findchildbytitle(B, "Buddy Lists")
D = ShowWindow(C, SW_SHOW)
End Sub

Sub Showwin(windo)
X = ShowWindow(windo, SW_SHOW)
End Sub

Function SignedOn()
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then
    SignedOn = 0
    Exit Function
End If
mdi% = Findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(mdi%, "Welcome, ")
If WEL% = 0 Then
    SignedOn = 0
    Exit Function
End If
COMB% = Findchildbyclass(WEL%, "_AOL_Combobox")
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
AOL% = findwindow("AOL Frame25", 0&)
mdi% = Findchildbyclass(AOL%, "MDIClient")

Do
C% = DoEvents()
EXX% = findchildbytitle(mdi%, "Exit?")
EXI% = findwindow("_AOL_Modal", "America Online")
Loop Until EXX% <> 0 Or EXI% <> 0
If EXX% <> 0 Then
    ICC% = Findchildbyclass(EXX%, "_AOL_Icon")
    ICC% = getwindow(ICC%, GW_HWNDLAST)
    ICC% = getwindow(ICC%, GW_HWNDPREV)
    ICC% = getwindow(ICC%, GW_HWNDPREV)
    D = SendMessageByNum(ICC%, WM_LBUTTONDOWN, 0, 0)
    D = SendMessageByNum(ICC%, WM_LBUTTONUP, 0, 0)
End If
If EXI% <> 0 Then
    YEE% = findchildbytitle(EXI%, "&Yes")
    D = SendMessageByNum(YEE%, WM_LBUTTONDOWN, 0, 0)
    D = SendMessageByNum(YEE%, WM_LBUTTONUP, 0, 0)
End If

End Sub

Sub SpiralScroller(txt As TextBox)
X = txt.Text
thastart:
sendtext "•––•^v^•{‡ " & txt & " ‡}•^v^•—•"
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
timeout 0.5
If txt.Text = X Then
sendtext "•––•^v^•{‡ " & txt & " ‡}•^v^•—•"
timeout 0.5
sendtext "•––•^v^•{‡ ÐeliriuM by GenOziDe ‡}•^v^•—•"
Exit Sub
End If
GoTo thastart
End Sub

Function StayOnlinea()
modal% = findwindow("_AOL_Modal", 0&)
btn4% = Findchildbyclass(modal%, "_AOL_Button")
btn3% = findchildbytitle(modal%, "Yes")
If btn3% <> 0 Then
Click btn3%
End If
AOL% = findwindow("aol Frame25", 0&)
palette% = Findchildbyclass(AOL%, "_AOL_Palette")
Btn% = Findchildbyclass(palette%, "_AOL_Button")
btn2% = findchildbytitle(palette%, "OK")
If palette% <> 0 Then
Click btn2%
End If
End Function

Sub stayontop(frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub stevecase(y)
Randomize
y = Int((Val("10") * Rnd) + 1)
If y = "1" Then
y = "Yes, What?"
ElseIf y = "2" Then: y = "Don't bother me please."
ElseIf y = "3" Then: y = "Leave me alone"
ElseIf y = "4" Then: y = "Hi! Do I know you?"
ElseIf y = "5" Then: y = "Hello, age/sex/location."
ElseIf y = "6" Then: y = "Hey!"
ElseIf y = "7" Then: y = "Hi! =)"
ElseIf y = "8" Then: y = "What's up?"
ElseIf y = "9" Then: y = "MmMmmmM... suck my dick harder Billy gates!!!!!!" & Chr(13) & "Woah.. sorry wrong IM =)"
ElseIf y = "10" Then: y = "Do you want to cyber?"
End If
End Sub

Sub SteveCaseIM()
Dim y, X$, IM%
X = getuser()
Call stevecase(Q$)
Call AOLInstantMessage(X$, Q$)
y = WaitForWindow("Untitled", "AOL Child")
z = ShowWindow(y, SW_HIDE)
timeout 2
IM% = findchildbytitle(AOLMDI(), ">Instant Message From:")
rich% = Findchildbyclass(IM%, "RICHCNTL")
Call textset(IM%, ">Instant Message From:  Steve Case")
Call textset(rich%, "Steve Case:" & Chr(9) & Q$ & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13))
z = ShowWindow(y, SW_SHOW)
End Sub

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Sub textset(hWnd As Integer, What As String)
Dim R
R = SendMessageByString(hWnd, &HC, 0, What)

End Sub

Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
X = DoEvents()
Loop
End Sub

Sub TMan(sill$, bah$)
AOL% = findwindow("AOL Frame25", 0&)
chatlist% = Findchildbyclass(AOL%, "_AOL_Listbox")
chatwin% = getparent(chatlist%)
ChaView% = Findchildbyclass(chatwin%, "_AOL_View")
chatedit% = Findchildbyclass(chatwin%, "_AOL_Edit")
ChatSend% = findchildbytitle(chatwin%, "Send")
sndtext% = SendMessageByString(ChaView%, WM_SETTEXT, 0, Chr(13) & Chr(9) & sill$ & ":" & Chr(9) & bah$)
SendNow% = SendMessageByNum(ChaView%, WM_CHAR, &HD, 0)
'Call Click(ChatSend%)
timeout (0.1)

End Sub

Function TOSPhrases() As String
Randomize
phraZes = Int((Val("141") * Rnd) + 1)
If phraZes = "1" Then
phraZes = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
phraZes = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf phraZes = "3" Then
phraZes = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf phraZes = "4" Then
phraZes = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf phraZes = "5" Then
phraZes = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf phraZes = "6" Then
phraZes = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "7" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "8" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "9" Then
phraZes = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf phraZes = "10" Then
phraZes = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "11" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "12" Then
phraZes = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "13" Then
phraZes = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf phraZes = "14" Then
phraZes = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf phraZes = "15" Then
phraZes = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "16" Then
phraZes = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "17" Then
phraZes = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf phraZes = "18" Then
phraZes = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf phraZes = "19" Then
phraZes = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf phraZes = "20" Then
phraZes = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf phraZes = "21" Then
phraZes = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf phraZes = "22" Then
phraZes = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf phraZes = "23" Then
phraZes = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf phraZes = "24" Then
phraZes = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf phraZes = "25" Then
phraZes = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf phraZes = "26" Then
phraZes = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf phraZes = "27" Then
phraZes = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf phraZes = "28" Then
phraZes = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
phraZes = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
phraZes = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
phraZes = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
phraZes = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
phraZes = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf phraZes = "34" Then
phraZes = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf phraZes = "35" Then
phraZes = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf phraZes = "36" Then
phraZes = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf phraZes = "37" Then
phraZes = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf phraZes = "38" Then
phraZes = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf phraZes = "39" Then
phraZes = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf phraZes = "40" Then
phraZes = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf phraZes = "41" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "42" Then
phraZes = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "43" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "44" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "45" Then
phraZes = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf phraZes = "46" Then
phraZes = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf phraZes = "47" Then
phraZes = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf phraZes = "48" Then
phraZes = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf phraZes = "49" Then
phraZes = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf phraZes = "50" Then
phraZes = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf phraZes = "51" Then
phraZes = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf phraZes = "52" Then
phraZes = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf phraZes = "53" Then
phraZes = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "54" Then
phraZes = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf phraZes = "55" Then
phraZes = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "56" Then
phraZes = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf phraZes = "57" Then
phraZes = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "58" Then
phraZes = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "59" Then
phraZes = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf phraZes = "60" Then
phraZes = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf phraZes = "61" Then
phraZes = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf phraZes = "62" Then
phraZes = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf phraZes = "63" Then
phraZes = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "64" Then
phraZes = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "65" Then
phraZes = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf phraZes = "66" Then
phraZes = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf phraZes = "67" Then
phraZes = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf phraZes = "68" Then
phraZes = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf phraZes = "69" Then
phraZes = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf phraZes = "70" Then
phraZes = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf phraZes = "71" Then
phraZes = "TrmsReport:  Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf phraZes = "72" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf phraZes = "73" Then
phraZes = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf phraZes = "74" Then
phraZes = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf phraZes = "75" Then
phraZes = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf phraZes = "76" Then
phraZes = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf phraZes = "77" Then
phraZes = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf phraZes = "78" Then
phraZes = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf phraZes = "79" Then
phraZes = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf phraZes = "80" Then
phraZes = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf phraZes = "81" Then
phraZes = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
phraZes = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
phraZes = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
phraZes = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
l012A = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "86" Then
phraZes = "Hi, I'm with FCB (Federal Credit Bureau). Due to The type of payment you are using, we need to validate it. Please reply with your credit card number, expiration date, full name, address, and phone number. Thank you. "
ElseIf phraZes = "87" Then
phraZes = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
ElseIf phraZes = "88" Then
phraZes = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your billing info.  Please click respond and send your CC#, Expiration Date, and Billing Zip Code. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "89" Then
phraZes = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "90" Then
phraZes = "AOLBilling:    Hello from the AOL TOS Staff. We have lost your account info. Please click respond and send your correct log-on password to your account. We are sorry for this trouble. Your account has been credited 15 free hours :-)"
ElseIf phraZes = "91" Then
phraZes = "Hello I am from the AOL service department and there has been a minor problem with your accout. I see that we have lost your password information and need it as soon as possible. We cannot access any of your personal info without your password. If you could please type your full log-on password and click the send button with your response. Thank you and Enjoy America online."
ElseIf phraZes = "92" Then
phraZes = "Hello, I am from AOLnet and there has been a network overload on our computers. We have been trying fix this problem but have had no luck. All your account information has been lost. In order for AOL to fix this we need your full First Name and Last, Your phone number, Address, Zip code, Credit Card Number and experation date, as well as your Screen Name(s) and Log-On Password(s). If you fail to reply will result in account termination, thank you :-)"
ElseIf phraZes = "93" Then
phraZes = "Hello I am a member of the America Online Stratus Team. We have been experiancing some problems with your preticular account. It would help us if you gave is your Screen Name(s) and Log-on Password(s) so we can remedy this problem. Thank you AOLST. "
ElseIf phraZes = "94" Then
phraZes = "Hello , AOL Customer! I am sorry to Inform you account information Was not Correctly Processed into our Stratus System. We are now forced to ask you to respond with your personal log on password and you complete billing info (Name, Address, & CC Number Information) . Thank-you for your Time, And Continue To enjoy AOL!"
ElseIf phraZes = "95" Then
phraZes = "Guide HAP:   Hello I work for AOL.  When you signed on we received an INCORRECT password.  Please respond with the CORRECT password now or face possible account termination for safety reasons, Thank you. =]"
ElseIf phraZes = "96" Then
phraZes = "Hi, I'm with STC (Standard Tele-Comm. Inc.) Due to a problem in phone transmission at your baud rate, I need to verify your log-on password before it can be fixed. Failure to comply will result in billing errors.  "
ElseIf phraZes = "97" Then
phraZes = "I am OnlineHost #147; an On-Line operator for America Online. Recently, computer Hackers have been logging on to the Network with your account. I have been logging them off when they failed validation. Please validate yourself as a registered user by stating your Credit Card# and expiration date. Thank you for your cooperation."
ElseIf phraZes = "98" Then
phraZes = "Good evening. I work with AOL's Network Security Division. Over the past few days, computer 'Hackers' have been logging on to the system using your account.  Please verify yourself as a registered user by stating your password, or face immediate account invalidation."
ElseIf phraZes = "99" Then
phraZes = "Hello! I am currently standing in as the TOSAdvisor because he is on vacation. I am sorry to inform you that your Password has been altered by computer Hackers. Your account will be terminated unless we can verify you as a registered user. Please type in your Password for system indentification. Thanks! :)"
ElseIf phraZes = "100" Then
phraZes = "Hello, I am with the America Online Billing department.  Please disreagrd the red letters at the bottom of the Instant Message window.  We have that there due to hackers who recently have been soliciting for passwords.  But Due to some sort of line noise or error when you logged on we were not able to verify your log on password.  Please reply with your log on password so that we may verify you, and we are sorry for the inconvinience."
ElseIf phraZes = "101" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "102" Then
phraZes = "This is a message sent out to random America Online members for a catalog containing -cheap- Microsoft software. Please send your request for the catalog along with your credit card number fo the amount $19.95. After 3 weeks processing time you'll recive a copy of Microsoft Windows96 Nashville version. *Microsoft INC.*"
ElseIf phraZes = "103" Then
phraZes = "*America Online Screen Name Dept.* Sorry for this inconvenience but somehow your screen name has been duplicated by another member. Please change your screen name and E-Mail me with your new screen name and password, So you can be verified as a AOL member. Thank you, And you have 5 days to reply."
ElseIf phraZes = "104" Then
phraZes = "Hello, I am with Sprint United Telephone. There has been damages to the fiber optic network wich you are connected to America Online with. I must ask for your log-on password so AOL can verify your current membership status. Failure to do so will result in account termination. Continue to enjoy your use of AOL courtesy of Sprint-TelNet."
ElseIf phraZes = "105" Then
phraZes = "Hi, I am with the AOL inc. User Services Department. We are sorry to inform you that your account has not been logged correctly please re-enter your personal logon Password, So you can continue to stay with us on America Online...Hope to see you soon..."
ElseIf phraZes = "106" Then
phraZes = "Hello, I am a service representative of America On Line.  Due to a computer error in our password files, we ask that you change your current log-in password to  'Secure', until futher advised, for your account's security. For your inconvience we at America On Line would like to offer our thanks by adding 2 Free hours of online time to your account. Thank you for your time and inconvience."
ElseIf phraZes = "107" Then
phraZes = "Hello this is the America Online party service and it seems that you have one a new 566.6 modem. This is the fastest modem out to this day, for you to redeem this prize we need your full log-on password to make sure you are not a hacker on this account, Thank you and have a nice day!"
ElseIf phraZes = "108" Then
phraZes = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "109" Then
phraZes = "Hello. I work for the America On-Line User Database. Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password onlineHello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation.. Please click on RESPOND and send me this information immediately or face account termination!  Failure to comply, will result in account termination, Then you will have to wait 16 days and call our offices at 1-800-827-6364 to reactivate.Thank you for your cooperation, America On-Line User Services."
ElseIf phraZes = "110" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "111" Then
phraZes = "Good evening, I am with the America Online Billing Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "112" Then
phraZes = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "113" Then
phraZes = "Hello, I am form the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf phraZes = "114" Then
phraZes = "Hello, I am from the America Online Billing Department.  I am sorry to report that we have been having trouble keeping your Billing Information valid.  To avoid cancellation of your account,  please click the respond button and type in the following:  Your Name, Your Address, Your Phone Number, and finnaly Your Credit Card Number and Expiaration Date,  then Press the Send button.  We are sorry for this inconvenience.  Thank you for your time and have a nice day."
ElseIf phraZes = "115" Then
phraZes = "Greetings CyberSurfer! I Am A Representative For the Online Banking System on AOL. As You May Have Read Credit Card Fraud is Huge on the 'Net'. Whould You Like To Safe Guard Your Credit Card Info With AOL? Just Reply With Your Credit Card Number And Full Information So That We may Logg Your Card Into Our Computers So No one else But You May Access It Online! We Are Offering This Safety Measure At No Cost to You At All!"
ElseIf phraZes = "116" Then
phraZes = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf phraZes = "117" Then
phraZes = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "118" Then
phraZes = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "119" Then
phraZes = "Hello, this is the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you for your cooperation."
ElseIf phraZes = "120" Then
phraZes = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf phraZes = "121" Then
phraZes = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day."
ElseIf phraZes = "122" Then
phraZes = "Welcome, I Am with America Online Credit Validators. Your credit card Information Has gone out of date. If you would like your account to stay active please respond with new credit card info, please reply with all the information so this error will not occur again. We at America Online are deticated In helping our customers! please feel free to call our 1-800 Number about our new credit plan. Thank you."
ElseIf phraZes = "123" Then
phraZes = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf phraZes = "124" Then
phraZes = "Hello I am with the America Online Security department. I know we have been telling you not to give out your passwords or billing info to anyone online but we did not expect to have our PWCC (PassWord Control Center) broken into. Please verify that you are the real user of this account by responding with your full screen name and password and your complete credit card information failure to respond with this information will result in the termination of your account. Thank you and enjoy America online. :-)"
ElseIf phraZes = "125" Then
phraZes = "Hello I am a undercover America Online billing representative. We have detected ''hackers'' using this account and purchasing goods with it, normally you would get stuck with the bill. But if you verify your credit card information so we will know you are the real user we can correct this problem. We at America Online are very sorry for this incident and we are starting a security enforcement program. that will prevent this from happening again. Thank You and have a nice day"
ElseIf phraZes = "126" Then
phraZes = "Hi, I am Chris Douglas and I'm from the America Online Billing Department.  I am sorry to inform you that we are having problems locating your billing information.  We need you to hit the Respond button, and reply with your Full name, Address, Credit card number with Expiration date, and your telephone number.  We apologize for this inconvenience.  Thank you and have a nice day."
ElseIf phraZes = "127" Then
phraZes = "Welcome To America Online! Thank you for choosing us as your online provider. We need to verify your billing information, so if you could please respond with your full name, address, city, state, zip code, and credit card information we will update your system. Sorry for the incovenience. Thank You!"
ElseIf phraZes = "128" Then
phraZes = "Welcome to America Online! While you were entering your billing information into your account systems our server went down and lost all the information conserning your account. If you could please respond with your full name, address, city, state, zip, credit card information, and log-on Password we can quickly update your syetem. Sorry for the misshap and thanks in advance! :-)"
ElseIf phraZes = "129" Then
phraZes = "Good Evenning, I am with the CAT (Community Action Team) of AOL. We are search your records and your Screen name came up as a unregisterd user. Please respond with your name, address, phone #, city , state, zip, & all of your credit card information (Number, exp. date, ect...). Sorry for circumstances but this is **VERY IMPORTANT**!"
ElseIf phraZes = "130" Then
phraZes = " Due to our last system crash on 8/12/96 we have lost all infromation concerning your account. Please respond with your Full address and Credit Card Number. NOTE: 60 mins are credited to your account because of the crash also."
ElseIf phraZes = "131" Then
phraZes = "  We're sorry to report that your account information has been erased.  For you to stay online we need you to verify your password, Credit Card number, Expiration date, and last 4 digits of your phone number.  Thank you and sorry for the delay."
ElseIf phraZes = "132" Then
phraZes = "Hello, my name is Todd Stiger.  I work for the internet.  I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your credit card number and personal information.  Thanks!"
ElseIf phraZes = "133" Then
phraZes = "Sir, I have some bad news! We here at AOL have made a SERIOUS billing error. We have your credit card number as 5436-2854-3546-9374, If this is not correct, please respond with the corrct number and your Full name, Address, State, City, and Zip code.  Sorry for the trouble!"
ElseIf phraZes = "134" Then
phraZes = "I am afraid that you have been SHARING your account with other members! THAT IS AGAINST TOS AGREEMENT!! Please respond with your credit card number, Full Name, and FULL address! This is the only way to prove wether or not you are sharing your account!"
ElseIf phraZes = "135" Then
phraZes = "Hello user, my name is William Denbo with the America Online Password Dept.! We have reason to beleive that a person has gotten hold of your password and is using it to conduct illegal activities. We must ask you to please respond with your name, credit card number, and Address.This is very important.Do not delay!"
ElseIf phraZes = "136" Then
phraZes = "Welcome new AOL user,Our service was temporarily shut down this morning between the times of 6:30a.m. - 7:30a.m.,Due to a complicated computer install!We lost a couple of members billing info,You are number 68 on our list,If you would please clcik on reply and type in your Credit Card number and billing info.!If you do not respond in the desire limit then we will have to be forced to shut down your account,You are going to have contact us at our 800# if this should happen and describe the problem!Thank you for your time. :-)"
ElseIf phraZes = "137" Then
phraZes = "Hello, I am with the AHD (Anti-Hacker Department) and we beleive that Hackers may be using your account for illegal activites. To verify you are the rightful owner of that account; please respond with you name, password, and credit card information. Thank you and have a nice day using America Online."
ElseIf phraZes = "138" Then
phraZes = "I have a program that can kick you off-lin and kill your account.So give me your password and I wont do it.But,if you dont give me your password I'll ruin your credit card,and kill your account!So don't fuck around!!!!!!"
ElseIf phraZes = "139" Then
phraZes = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf phraZes = "140" Then
phraZes = "Good morning current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a virus in the system earlier and it may of caused errors in your billing,When you reply with the information you entered during sign-up so we can verify your account and send you your current months total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time."
ElseIf phraZes = "141" Then
phraZes = "There have been notable billing errors showing in our records. For us to correct this matter efficiently, please hit rrespond and give me your full name, your full address, and your credit card number with experation date.  I am sorry for the inconvienence!"
End If
End Function

Function Trimnull() As String
Dim X, total$
For X = 1 To Len(in)
    If (Mid$(in, X, 1) <> Chr$(0)) Then
        total$ = total$ + Mid$(in, X, 1)
    Else
        GoTo NullDetect
    End If
Next
NullDetect:
Trimnull = total$
End Function

Function trm(ByVal txt As String)
Dim X As Integer
Dim y As String
Dim z As String
For X = 1 To Len(txt)
y = Mid(txt, X, 1)
If y = Chr(0) Then y = ""
z = z & y
Next X
trm = z
End Function

Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
firs% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bon
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bon

While firs%
firss% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bon
firs% = getwindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bon
Wend
GoTo GoBack
FindClassLike = 0

bon:
Room% = firs%
UntilWindowClass = Room%
End Function

Sub vbmsgg()
chattxt$ = agGetStringFromlpstr$(lParam)
End Sub

Sub waitforchange()
Dim Old As Integer
Dim Boy As Integer
Old = getfocus()
Boy = getfocus()
Do Until Boy <> Old
X = DoEvents()
Boy = getfocus()
Loop
End Sub

Sub waitformok()
Do
DoEvents
okw = findwindow("_AOL_Modal", 0&)
DoEvents
Loop Until okw <> 0
okw = findwindow("_AOL_Modal", 0&)
killwin (okw)

End Sub

Sub waitforok()
Do
DoEvents
okw = findwindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
okb = findchildbytitle(okw, "OK")

DoEvents
Loop Until okb <> 0
Do
okw = findwindow("#32770", "America Online")
    okb = findchildbytitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
DoEvents
Loop Until okw = 0

End Sub

Function WaitForWindow(txt As String, clss As String) As Integer
Dim R
Dim hWnds() As Integer

Do
R = FindWindowLike(hWnds(), 0, txt, clss, Null)
DoEvents
Loop While R = 0
WaitForWindow = hWnds(1)

End Function

Sub waitmail()
Do
box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
timeout (0.1)
Loop Until box <> 0
List = Findchildbyclass(box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(0.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(0.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)


End Sub

Sub WatchForCatwatch()
For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo Genocide
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
w = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If w <> 0 Then
Call keywor("PC")
MsgBox "A Cat had entered the room."
End If
If X <> 0 Then
Call keywor("PC")
MsgBox "A Cat had entered the room."
End If
Next Index%
Genocide:

End Sub

Sub WatchForGuides()
For Index% = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call keywor("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
End Sub

Sub wate(howlong)
Do
a = 1 + a
Loop Until a > howlong

End Sub

Function WindowCaption(hWndd As Integer)
Dim WindowText As String * 255
Dim GetWinText As Integer
GetWinText = GetWindowText(hWndd, WindowText, 255)
WindowCaption = (WindowText)
End Function

Function WindowsDirectory() As String
Dim WinPath As String
WinPath = String(145, Chr(0))
WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))

End Function

Function WinTxt(ByVal hWnd As Integer)
Dim X As Integer
Dim y As String
Dim z As Integer
X = sendmessage(hWnd, &HE, 0&, 0&)
y = String(X + 1, " ")
z = SendMessageByString(hWnd, &HD, X + 1, y)
WinTxt = Left(y, X)
End Function

Sub WriteINI(sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim R As Integer
    R = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

