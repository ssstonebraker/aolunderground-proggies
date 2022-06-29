Declare Sub bringwindowtotop Lib "User" (ByVal hWnd As Integer)
Declare Function extfn2668 Lib "VBWFind.Dll" Alias "Findchildbyclass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn16A8 Lib "User" Alias "FindWindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn1868 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal X%, ByVal y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Declare Function extfn2630 Lib "VBWFind.Dll" Alias "Findchildbytitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1AD0 Lib "User" Alias "SetWindowText" (ByVal p1%, ByVal p2$) As Integer
'Declare Function getfreespace Lib "kernel" (ByVal wflags As Integer) As Long
'Declare Function getfreesystemresources Lib "user" (ByVal fusysresource As Integer) As Integer
'Declare Function Getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGuide.Dll" (ByVal lpString&)
Declare Function sendmessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

'
'Master1.bas updated by ePoD for
' aNnHi£aTioN
'

Option Compare Text
'                      Types
'                      -----
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
Type RECT
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type


'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Function GetMenuString Lib "User" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, LPRect As RECT)
'Declare Sub SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal F%)
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal X As Integer, ByVal y As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function getfreesystemresources Lib "User" (ByVal fusysresource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function getfocus% Lib "User" ()
Declare Function setfocusapi% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function getwindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function findwindow% Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any)
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function getparent% Lib "User" (ByVal hWnd As Integer)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
Declare Function sendmessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
'Declare Function SendMessage& Lib "User" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam As Any)
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function sendmessagebynum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function CreateMenu% Lib "User" ()
Declare Function AppendMenu% Lib "User" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem&)
Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wflags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function getwindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function GetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function GetCurrentTime& Lib "User" ()
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetCursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowtextlength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GettopWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetMenu% Lib "User" (ByVal hWnd%)
Declare Function GetMenuItemID% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuItemCount% Lib "User" (ByVal hMenu%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wflags%)
Declare Function GetSubMenu% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hdc%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function getfreespace Lib "Kernel" (ByVal wflags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal maxsize As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer

'Subs and Functions for "GDI"
Declare Sub SetBKColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal X As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer

'Subs and Functions for "MMSystem"
Declare Function sndPlaySound Lib "MMSystem" (ByVal lpWavName$, ByVal Flags%) As Integer '
Declare Function MciSendString& Lib "MMSystem" (ByVal Cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)
Declare Function findchildbytitle% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)
Declare Function findchildbyclass% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)

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
Declare Function agVBSendControlMsg& Lib "APIGuide.Dll" (ctl As Control, ByVal Msg%, ByVal wp%, ByVal lp&)
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
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (I As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal I As Long) As Integer
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
Declare Function GetNames Lib "311.dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)
Declare Function vbeNumChildWindow% Lib "VBStr.Dll" (ByVal win%, ByVal iNum%)

'                   Global Constants
'                   ----------------
'OpenFile() Flags
Global Const WM_USER = &H400
Global Const LB_GETCOUNT = (WM_USER + 12)

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

'Device Technologies
Global Const DT_PLOTTER = 0
Global Const DT_RASDISPLAY = 1
Global Const DT_RASPRINTER = 2
Global Const DT_RASCAMERA = 3
Global Const DT_CHARSTREAM = 4
Global Const DT_METAFILE = 5
Global Const DT_DISPFILE = 6

'Curve Capabilities
Global Const CC_NONE = 0
Global Const CC_CIRCLES = 1
Global Const CC_PIE = 2
Global Const CC_CHORD = 4
Global Const CC_ELLIPSES = 8
Global Const CC_WIDE = 16
Global Const CC_STYLED = 32
Global Const CC_WIDESTYLED = 64
Global Const CC_INTERIORS = 128

'Line Capabilities
Global Const LC_NONE = 0
Global Const LC_POLYLINE = 2
Global Const LC_MARKER = 4
Global Const LC_POLYMARKER = 8
Global Const LC_WIDE = 16
Global Const LC_STYLED = 32
Global Const LC_WIDESTYLED = 64
Global Const LC_INTERIORS = 128

'Polygonal Capabilities
Global Const PC_NONE = 0
Global Const PC_POLYGON = 1
Global Const PC_RECTANGLE = 2
Global Const PC_WINDPOLYGON = 4
Global Const PC_TRAPEZOID = 4
Global Const PC_SCANLINE = 8
Global Const PC_WIDE = 16
Global Const PC_STYLED = 32
Global Const PC_WIDESTYLED = 64
Global Const PC_INTERIORS = 128

'Polygonal Capabilities
Global Const CP_NONE = 0
Global Const CP_RECTANGLE = 1

'Text Capabilities
Global Const TC_OP_CHARACTER = &H1
Global Const TC_OP_STROKE = &H2
Global Const TC_CP_STROKE = &H4
Global Const TC_CR_90 = &H8
Global Const TC_CR_ANY = &H10
Global Const TC_SF_X_YINDEP = &H20
Global Const TC_SA_DOUBLE = &H40
Global Const TC_SA_INTEGER = &H80
Global Const TC_SA_CONTIN = &H100
Global Const TC_EA_DOUBLE = &H200
Global Const TC_IA_ABLE = &H400
Global Const TC_UA_ABLE = &H800
Global Const TC_SO_ABLE = &H1000
Global Const TC_RA_ABLE = &H2000
Global Const TC_VA_ABLE = &H4000
Global Const TC_RESERVED = &H8000

'Raster Capabilities
Global Const RC_BITBLT = 1
Global Const RC_BANDING = 2
Global Const RC_SCALING = 4
Global Const RC_BITMAP64 = 8
Global Const RC_GDI20_OUTPUT = &H10
Global Const RC_DI_BITMAP = &H80
Global Const RC_PALETTE = &H100
Global Const RC_DIBTODEV = &H200
Global Const RC_BIGFONT = &H400
Global Const RC_STRETCHBLT = &H800
Global Const RC_FLOODFILL = &H1000
Global Const RC_STRETCHDIB = &H2000

'palette entry flags
Global Const PC_RESERVED = &H1
Global Const PC_EXPLICIT = &H2
Global Const PC_NOCOLLAPSE = &H4

'DIB color table identifiers
Global Const DIB_RGB_COLORS = 0
Global Const DIB_PAL_COLORS = 1

'constants for Get/SetSystemPaletteUse()
Global Const SYSPAL_STATIC = 1
Global Const SYSPAL_NOSTATIC = 2

'constants for CreateDIBitmap
Global Const CBM_INIT = &H4&

'DrawText() Format Flags
Global Const DT_TOP = &H0
Global Const DT_LEFT = &H0
Global Const DT_CENTER = &H1
Global Const DT_RIGHT = &H2
Global Const DT_VCENTER = &H4
Global Const DT_BOTTOM = &H8
Global Const DT_WORDBREAK = &H10
Global Const DT_SINGLELINE = &H20
Global Const DT_EXPANDTABS = &H40
Global Const DT_TABSTOP = &H80
Global Const DT_NOCLIP = &H100
Global Const DT_EXTERNALLEADING = &H200
Global Const DT_CALCRECT = &H400
Global Const DT_NOPREFIX = &H800
Global Const DT_INTERNAL = &H1000

'ExtFloodFill style flags
Global Const FLOODFILLBORDER = 0
Global Const FLOODFILLSURFACE = 1


'Scroll Bar Constants
Global Const SB_HORZ = 0
Global Const SB_VERT = 1
Global Const SB_CTL = 2
Global Const SB_BOTH = 3

'Scroll Bar Commands
Global Const SB_LINEUP = 0
Global Const SB_LINEDOWN = 1
Global Const SB_PAGEUP = 2
Global Const SB_PAGEDOWN = 3
Global Const SB_THUMBPOSITION = 4
Global Const SB_THUMBTRACK = 5
Global Const SB_TOP = 6
Global Const SB_BOTTOM = 7
Global Const SB_ENDSCROLL = 8

'ShowWindow() Commands
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

'Old ShowWindow() Commands
Global Const HIDE_WINDOW = 0
Global Const SHOW_OPENWINDOW = 1
Global Const SHOW_ICONWINDOW = 2
Global Const SHOW_FULLSCREEN = 3
Global Const SHOW_OPENNOACTIVATE = 4

'Identifiers for the WM_SHOWWINDOW message
Global Const SW_PARENTCLOSING = 1
Global Const SW_OTHERZOOM = 2
Global Const SW_PARENTOPENING = 3
Global Const SW_OTHERUNZOOM = 4

'RedrawWindow flags
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

'ScrollWindowEx flags
Global Const SW_SCROLLCHILDREN = &H1
Global Const SW_INVALIDATE = &H2
Global Const SW_ERASE = &H4

'Region Flags
Global Const ERRORAPI = 0
Global Const NULLREGION = 1
Global Const SIMPLEREGION = 2
Global Const COMPLEXREGION = 3

'CombineRgn() Styles
Global Const RGN_AND = 1
Global Const RGN_OR = 2
Global Const RGN_XOR = 3
Global Const RGN_DIFF = 4
Global Const RGN_COPY = 5

'Virtual Keys, Standard Set
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

'Queue Status
Global Const QS_KEY = 1
Global Const QS_MOUSEMOVE = 2
Global Const QS_MOUSEBUTTON = 4
Global Const QS_MOUSE = 6
Global Const QS_POSTMESSAGE = 8
Global Const QS_TIMER = &H10
Global Const QS_PAINT = &H20
Global Const QS_SENDMESSAGE = &H40
Global Const QS_ALLINPUT = &H7F

'SetWindowHook() codes
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

'Hook Codes
Global Const HC_LPLPFNNEXT = (-2)
Global Const HC_LPFNNEXT = (-1)
Global Const HC_ACTION = 0
Global Const HC_GETNEXT = 1
Global Const HC_SKIP = 2
Global Const HC_NOREM = 3
Global Const HC_NOREMOVE = 3
Global Const HC_SYSMODALON = 4
Global Const HC_SYSMODALOFF = 5

'CBT Hook Codes
Global Const HCBT_MOVESIZE = 0
Global Const HCBT_MINMAX = 1
Global Const HCBT_QS = 2

'WH_MSGFILTER Filter Proc Codes
Global Const MSGF_DIALOGBOX = 0
Global Const MSGF_MESSAGEBOX = 1
Global Const MSGF_MENU = 2
Global Const MSGF_MOVE = 3
Global Const MSGF_SIZE = 4
Global Const MSGF_SCROLLBAR = 5
Global Const MSGF_NEXTWINDOW = 6

'Window Manager Hook Codes
Global Const WC_INIT = 1
Global Const WC_SWP = 2
Global Const WC_DEFWINDOWPROC = 3
Global Const WC_MINMAX = 4
Global Const WC_MOVE = 5
Global Const WC_SIZE = 6
Global Const WC_DRAWCAPTION = 7

'Window field offsets for GetWindowLong() and GetWindowWord()
Global Const GWL_WNDPROC = (-4)
Global Const GWW_HINSTANCE = (-6)
'Global Const GWW_HWNDPARENT = (-8)
Global Const GWW_ID = (-12)
Global Const GWL_STYLE = (-16)
Global Const GWL_EXSTYLE = (-20)

'GetWindowLong and and GetWindowWord dialog box constants
Global Const DWL_MSGRESULT = 0
Global Const DWL_DLGPROC = 4
Global Const DWL_USER = 8

'Class field offsets for GetClassLong() and GetClassWord()
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

'SendMessage Flag
Global Const HWND_BROADCAST = -1

'Window Messages
Global Const WM_NULL = &H0
Global Const WM_CREATE = &H1
Global Const WM_DESTROY = &H2
Global Const WM_MOVE = &H3
Global Const WM_SIZE = &H5
Global Const WM_ACTIVATE = &H6
Global Const WM_SETFOCUS = &H7
Global Const WM_KILLFOCUS = &H8
Global Const WM_ENABLE = &HA
Global Const WM_SETREDRAW = &HB


Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_PAINT = &HF
Global Const WM_CLOSE = &H10
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

Global Const WM_QUEUESYNC = &H23
Global Const WM_GETMINMAXINFO = &H24
Global Const WM_PAINTICON = &H26
Global Const WM_ICONERASEBKGND = &H27
Global Const WM_NEXTDLGCTL = &H28
Global Const WM_SPOOLERSTATUS = &H2A
Global Const WM_DRAWITEM = &H2B
Global Const WM_MEASUREITEM = &H2C
Global Const WM_DELETEITEM = &H2D
Global Const WM_VKEYTOITEM = &H2E
Global Const WM_CHARTOITEM = &H2F
Global Const WM_SETFONT = &H30

Global Const WM_COMMNOTIFY = &H44
Global Const WM_QUERYDRAGICON = &H37
Global Const WM_COMPAREITEM = &H39
Global Const WM_COMPACTING = &H41
Global Const WM_WINDOWPOSCHANGING = &H46
Global Const WM_WINDOWPOSCHANGED = &H47
Global Const WM_POWER = &H48
Global Const WM_NCCREATE = &H81
Global Const WM_NCDESTROY = &H82
Global Const WM_NCCALCSIZE = &H83
Global Const WM_NCHITTEST = &H84
Global Const WM_NCPAINT = &H85
Global Const WM_NCACTIVATE = &H86
Global Const WM_GETDLGCODE = &H87
Global Const WM_NCMOUSEMOVE = &HA0
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_NCLBUTTONUP = &HA2
Global Const WM_NCLBUTTONDBLCLK = &HA3
Global Const WM_NCRBUTTONDOWN = &HA4
Global Const WM_NCRBUTTONUP = &HA5
Global Const WM_NCRBUTTONDBLCLK = &HA6
Global Const WM_NCMBUTTONDOWN = &HA7
Global Const WM_NCMBUTTONUP = &HA8
Global Const WM_NCMBUTTONDBLCLK = &HA9
Global Const WM_KEYFIRST = &H100
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101

Global Const WM_DEADCHAR = &H103


Global Const WM_SYSCHAR = &H106
Global Const WM_SYSDEADCHAR = &H107
Global Const WM_KEYLAST = &H108
Global Const WM_INITDIALOG = &H110

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
'WM_SYNCTASK Commands
Global Const ST_BEGINSWP = 0
Global Const ST_ENDSWP = 1


'WM_ACTIVATE constants
Global Const WA_INACTIVE = 0
Global Const WA_ACTIVE = 1
Global Const WA_CLICKACTIVE = 2


'WinWhere() Area Codes
Global Const HTERROR = (-2)
Global Const HTTRANSPARENT = (-1)
Global Const HTNOWHERE = 0
Global Const HTCLIENT = 1
Global Const HTCAPTION = 2
Global Const HTSYSMENU = 3
Global Const HTGROWBOX = 4
Global Const HTSIZE = HTGROWBOX
Global Const HTMENU = 5
Global Const HTHSCROLL = 6
Global Const HTVSCROLL = 7
Global Const HTREDUCE = 8
Global Const HTZOOM = 9
Global Const HTLEFT = 10
Global Const HTRIGHT = 11
Global Const HTTOP = 12
Global Const HTTOPLEFT = 13
Global Const HTTOPRIGHT = 14
Global Const HTBOTTOM = 15
Global Const HTBOTTOMLEFT = 16
Global Const HTBOTTOMRIGHT = 17
Global Const HTSIZEFIRST = HTLEFT
Global Const HTSIZELAST = HTBOTTOMRIGHT

'WM_MOUSEACTIVATE Return Codes
Global Const MA_ACTIVATE = 1
Global Const MA_ACTIVATEANDEAT = 2
Global Const MA_NOACTIVATE = 3
Global Const MA_NOACTIVATEANDEAT = 4


'Size Message Commands
Global Const SIZENORMAL = 0
Global Const SIZEICONIC = 1
Global Const SIZEFULLSCREEN = 2
Global Const SIZEZOOMSHOW = 3
Global Const SIZEZOOMHIDE = 4

'Key State Masks for Mouse Messages
Global Const MK_LBUTTON = &H1
Global Const MK_RBUTTON = &H2
Global Const MK_SHIFT = &H4
Global Const MK_CONTROL = &H8
Global Const MK_MBUTTON = &H10

'Window Styles
Global Const WS_OVERLAPPED = &H0&
Global Const WS_POPUP = &H80000000
Global Const WS_CHILD = &H40000000
Global Const WS_MINIMIZE = &H20000000
Global Const WS_VISIBLE = &H10000000
Global Const WS_DISABLED = &H8000000
Global Const WS_CLIPSIBLINGS = &H4000000
Global Const WS_CLIPCHILDREN = &H2000000
Global Const WS_MAXIMIZE = &H1000000
Global Const WS_CAPTION = &HC00000
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
'Common Window Styles
Global Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Global Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Global Const WS_CHILDWINDOW = (WS_CHILD)
Global Const WS_TILEDWINDOW = (WS_OVERLAPPEDWINDOW)

'Extended Window Styles
Global Const WS_EX_DLGMODALFRAME = &H1&
Global Const WS_EX_NOPARENTNOTIFY = &H4&
Global Const WS_EX_TOPMOST = &H8&
Global Const WS_EX_ACCEPTFILES = &H10&
Global Const WS_EX_TRANSPARENT = &H20&

' MDI style allows use of all child styles
Global Const MDIS_ALLCHILDSTYLES = &H1&

'Class styles
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
Global Const CS_GLOBALCLASS = &H4000

'Predefined Clipboard Formats
Global Const CF_TEXT = 1
Global Const CF_BITMAP = 2
Global Const CF_METAFILEPICT = 3
Global Const CF_SYLK = 4
Global Const CF_DIF = 5
Global Const CF_TIFF = 6
Global Const CF_OEMTEXT = 7
Global Const CF_DIB = 8
Global Const CF_PALETTE = 9
Global Const CF_OWNERDISPLAY = &H80
Global Const CF_DSPTEXT = &H81
Global Const CF_DSPBITMAP = &H82
Global Const CF_DSPMETAFILEPICT = &H83

'"Private" formats don't get GlobalFree()'d
Global Const CF_PRIVATEFIRST = &H200
Global Const CF_PRIVATELAST = &H2FF

'"GDIOBJ" formats do get DeleteObject()'d
Global Const CF_GDIOBJFIRST = &H300
Global Const CF_GDIOBJLAST = &H3FF


'Owner draw control types
Global Const ODT_MENU = 1
Global Const ODT_LISTBOX = 2
Global Const ODT_COMBOBOX = 3
Global Const ODT_BUTTON = 4

'Owner draw actions
Global Const ODA_DRAWENTIRE = &H1
Global Const ODA_SELECT = &H2
Global Const ODA_FOCUS = &H4

'Owner draw state
Global Const ODS_SELECTED = &H1
Global Const ODS_GRAYED = &H2
Global Const ODS_DISABLED = &H4
Global Const ODS_CHECKED = &H8
Global Const ODS_FOCUS = &H10


'PeekMessage() Options
Global Const PM_NOREMOVE = &H0
Global Const PM_REMOVE = &H1
Global Const PM_NOYIELD = &H2

'Flags for _lopen
Global Const READAPI = 0
Global Const WRITEAPI = 1
Global Const READ_WRITE = 2


'Window placement flags
Global Const CW_USEDEFAULT = &H8000
Global Const WPF_SETMINPOSITION = 1
Global Const WPF_RESTORETOMAXIMIZED = 2

'SetWindowPos Flags
'Global Const SWP_NOSIZE = &H1
'Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_DRAWFRAME = &H20
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOREPOSITION = &H200
Global Const SWP_NOSENDCHANGING = &H400
Global Const SWP_DEFERERASE = &H2000

'SetWindowPos() hwndInsertAfter values
Global Const HWND_TOP = 0
Global Const HWND_BOTTOM = 1
'Global Const HWND_TOPMOST = -1
'Global Const HWND_NOTOPMOST = -2
Global Const DLGWINDOWEXTRA = 30

'GetSystemMetrics() codes
Global Const SM_CXSCREEN = 0
Global Const SM_CYSCREEN = 1
Global Const SM_CXVSCROLL = 2
Global Const SM_CYHSCROLL = 3
Global Const SM_CYCAPTION = 4
Global Const SM_CXBORDER = 5
Global Const SM_CYBORDER = 6
Global Const SM_CXDLGFRAME = 7
Global Const SM_CYDLGFRAME = 8
Global Const SM_CYVTHUMB = 9
Global Const SM_CXHTHUMB = 10
Global Const SM_CXICON = 11
Global Const SM_CYICON = 12
Global Const SM_CXCURSOR = 13
Global Const SM_CYCURSOR = 14
Global Const SM_CYMENU = 15
Global Const SM_CXFULLSCREEN = 16
Global Const SM_CYFULLSCREEN = 17
Global Const SM_CYKANJIWINDOW = 18
Global Const SM_MOUSEPRESENT = 19
Global Const SM_CYVSCROLL = 20
Global Const SM_CXHSCROLL = 21
Global Const SM_DEBUG = 22
Global Const SM_SWAPBUTTON = 23
Global Const SM_RESERVED1 = 24
Global Const SM_RESERVED2 = 25
Global Const SM_RESERVED3 = 26
Global Const SM_RESERVED4 = 27
Global Const SM_CXMIN = 28
Global Const SM_CYMIN = 29
Global Const SM_CXSIZE = 30
Global Const SM_CYSIZE = 31
Global Const SM_CXFRAME = 32
Global Const SM_CYFRAME = 33
Global Const SM_CXMINTRACK = 34
Global Const SM_CYMINTRACK = 35
Global Const SM_CXDOUBLECLK = 36
Global Const SM_CYDOUBLECLK = 37
Global Const SM_CXICONSPACING = 38
Global Const SM_CYICONSPACING = 39
Global Const SM_MENUDROPALIGNMENT = 40
Global Const SM_PENWindow = 41
Global Const SM_DBCSENABLED = 42

'System parameters support

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
Global Const SPI_GETSCREENSAVEPause = 14
Global Const SPI_SETSCREENSAVEPause = 15
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

'SystemParametersInfo flags
Global Const SPIF_UPDATEINIFILE = 1
Global Const SPIF_SENDWININICHANGE = 2
'MessageBox() Flags
Global Const MB_OK = &H0
Global Const MB_OKCANCEL = &H1
Global Const MB_ABORTRETRYIGNORE = &H2
Global Const MB_YESNOCANCEL = &H3
Global Const MB_YESNO = &H4
Global Const MB_RETRYCANCEL = &H5
Global Const MB_ICONHAND = &H10
Global Const MB_ICONQUESTION = &H20
Global Const MB_ICONEXCLAMATION = &H30
Global Const MB_ICONASTERISK = &H40
Global Const MB_ICONINFORMATION = MB_ICONASTERISK
Global Const MB_ICONSTOP = MB_ICONHAND
Global Const MB_DEFBUTTON1 = &H0
Global Const MB_DEFBUTTON2 = &H100
Global Const MB_DEFBUTTON3 = &H200
Global Const MB_APPLMODAL = &H0
Global Const MB_SYSTEMMODAL = &H1000
Global Const MB_TASKMODAL = &H2000
Global Const MB_NOFOCUS = &H8000
Global Const MB_TYPEMASK = &HF
Global Const MB_ICONMASK = &HF0
Global Const MB_DEFMASK = &HF00
Global Const MB_MODEMASK = &H3000
Global Const MB_MISCMASK = &HC000

'Color Types
Global Const CTLCOLOR_MSGBOX = 0
Global Const CTLCOLOR_EDIT = 1
Global Const CTLCOLOR_LISTBOX = 2
Global Const CTLCOLOR_BTN = 3
Global Const CTLCOLOR_DLG = 4
Global Const CTLCOLOR_SCROLLBAR = 5
Global Const CTLCOLOR_STATIC = 6
Global Const CTLCOLOR_MAX = 8 'three bits max
Global Const COLOR_SCROLLBAR = 0
Global Const COLOR_BACKGROUND = 1
Global Const COLOR_ACTIVECAPTION = 2
Global Const COLOR_INACTIVECAPTION = 3
Global Const COLOR_MENU = 4
Global Const COLOR_WINDOW = 5
Global Const COLOR_WINDOWFRAME = 6
Global Const COLOR_MENUTEXT = 7
Global Const COLOR_WINDOWTEXT = 8
Global Const COLOR_CAPTIONTEXT = 9
Global Const COLOR_ACTIVEBORDER = 10
Global Const COLOR_INACTIVEBORDER = 11
Global Const COLOR_APPWORKSPACE = 12
Global Const COLOR_HIGHLIGHT = 13
Global Const COLOR_HIGHLIGHTTEXT = 14
Global Const COLOR_BTNFACE = 15
Global Const COLOR_BTNSHADOW = 16
Global Const COLOR_GRAYTEXT = 17
Global Const COLOR_BTNTEXT = 18
Global Const COLOR_INACTIVECAPTIONTEXT = 19
Global Const COLOR_BTNHIGHLIGHT = 20

'GetWindow() Constants


Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5

'GetDCEx flags
Global Const DCX_WINDOW = &H1&
Global Const DCX_CACHE = &H2&
Global Const DCX_CLIPCHILDREN = &H8&
Global Const DCX_CLIPSIBLINGS = &H10&
Global Const DCX_PARENTCLIP = &H20&
Global Const DCX_EXCLUDERGN = &H40&
Global Const DCX_INTERSECTRGN = &H80&
Global Const DCX_LOCKWINDOWUPDATE = &H400&
Global Const DCX_USESTYLE = &H10000

'Menu flags for Add/Check/EnableMenuItem()
Global Const MF_INSERT = &H0
Global Const MF_CHANGE = &H80
Global Const MF_APPEND = &H100
Global Const MF_DELETE = &H200
Global Const MF_REMOVE = &H1000

Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800
Global Const MF_ENABLED = &H0
Global Const MF_GRAYED = &H1
Global Const MF_DISABLED = &H2
Global Const MF_UNCHECKED = &H0
Global Const MF_CHECKED = &H8
Global Const MF_USECHECKBITMAPS = &H200
Global Const MF_STRING = &H0
Global Const MF_BITMAP = &H4
Global Const MF_OWNERDRAW = &H100
Global Const MF_POPUP = &H10
Global Const MF_MENUBARBREAK = &H20
Global Const MF_MENUBREAK = &H40
Global Const MF_UNHILITE = &H0
Global Const MF_HILITE = &H80
Global Const MF_SYSMENU = &H2000
Global Const MF_HELP = &H4000
Global Const MF_MOUSESELECT = &H8000
Global Const MF_END = &H80

'TrackPopupMenu flags
Global Const TPM_LEFTBUTTON = &H0
Global Const TPM_RIGHTBUTTON = &H2
Global Const TPM_LEFTALIGN = &H0
Global Const TPM_CENTERALIGN = &H4
Global Const TPM_RIGHTALIGN = &H8

'System Menu Command Values
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

'Standard Cursor IDs
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
Global Const ORD_LANGDRIVER = 1

'Standard Icon IDs
Global Const IDI_APPLICATION = 32512&
Global Const IDI_HAND = 32513&
Global Const IDI_QUESTION = 32514&
Global Const IDI_EXCLAMATION = 32515&
Global Const IDI_ASTERISK = 32516&

'Dialog Box Command IDs
Global Const IDOK = 1
Global Const IDCANCEL = 2
Global Const IDABORT = 3
Global Const IDRETRY = 4
Global Const IDIGNORE = 5
Global Const IDYES = 6
Global Const IDNO = 7

'Edit Control Styles
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

'Edit Control Notification Codes
Global Const EN_SETFOCUS = &H100
Global Const EN_KILLFOCUS = &H200
Global Const EN_CHANGE = &H300
Global Const EN_UPDATE = &H400
Global Const EN_ERRSPACE = &H500
Global Const EN_MAXTEXT = &H501
Global Const EN_HSCROLL = &H601
Global Const EN_VSCROLL = &H602
Global Const WB_LEFT = 0
Global Const WB_RIGHT = 1
Global Const WB_ISDELIMITER = 2

'Button Control Styles
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

'User Button Notification Codes
Global Const BN_CLICKED = 0
Global Const BN_PAINT = 1
Global Const BN_HILITE = 2
Global Const BN_UNHILITE = 3
Global Const BN_DISABLE = 4
Global Const BN_DOUBLECLICKED = 5

'Static Control Constants
Global Const SS_LEFT = &H0&


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
Global Const SS_NOPREFIX = &H80&


'Dialog Styles
Global Const DS_ABSALIGN = &H1&
Global Const DS_SYSMODAL = &H2&
Global Const DS_LOCALEDIT = &H20&
Global Const DS_SETFONT = &H40&
Global Const DS_MODALFRAME = &H80&
Global Const DS_NOIDLEMSG = &H100&
Global Const DC_HASDEFID = &H534

'Dialog Codes
Global Const DLGC_WANTARROWS = &H1
Global Const DLGC_WANTTAB = &H2
Global Const DLGC_WANTALLKEYS = &H4
Global Const DLGC_WANTMESSAGE = &H4
Global Const DLGC_HASSETSEL = &H8
Global Const DLGC_DEFPUSHBUTTON = &H10
Global Const DLGC_UNDEFPUSHBUTTON = &H20
Global Const DLGC_RADIOBUTTON = &H40
Global Const DLGC_WANTCHARS = &H80
Global Const DLGC_STATIC = &H100
Global Const DLGC_BUTTON = &H2000

'Scroll Bar Styles
Global Const SBS_HORZ = &H0&
Global Const SBS_VERT = &H1&
Global Const SBS_TOPALIGN = &H2&
Global Const SBS_LEFTALIGN = &H2&
Global Const SBS_BOTTOMALIGN = &H4&
Global Const SBS_RIGHTALIGN = &H4&
Global Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Global Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Global Const SBS_SIZEBOX = &H8&

'WaitSoundState() Constants
Global Const S_QUEUEEMPTY = 0
Global Const S_THRESHOLD = 1
Global Const S_ALLTHRESHOLD = 2

'Accent Modes
Global Const S_NORMAL = 0
Global Const S_LEGATO = 1
Global Const S_STACCATO = 2

'SetSoundNoise() Sources
Global Const S_PERIOD512 = 0 '
Global Const S_PERIOD1024 = 1
Global Const S_PERIOD2048 = 2
Global Const S_PERIODVOICE = 3
Global Const S_WHITE512 = 4
Global Const S_WHITE1024 = 5
Global Const S_WHITE2048 = 6
Global Const S_WHITEVOICE = 7
Global Const S_SERDVNA = (-1)
Global Const S_SEROFM = (-2)
Global Const S_SERMACT = (-3)
Global Const S_SERQFUL = (-4)
Global Const S_SERBDNT = (-5)
Global Const S_SERDLN = (-6)
Global Const S_SERDCC = (-7)
Global Const S_SERDTP = (-8)
Global Const S_SERDVL = (-9)
Global Const S_SERDMD = (-10)
Global Const S_SERDSH = (-11)
Global Const S_SERDPT = (-12)
Global Const S_SERDFQ = (-13)
Global Const S_SERDDR = (-14)
Global Const S_SERDSR = (-15)
Global Const S_SERDST = (-16)

'COMM declarations
Global Const NOPARITY = 0
Global Const ODDPARITY = 1
Global Const EVENPARITY = 2
Global Const MARKPARITY = 3
Global Const SPACEPARITY = 4
Global Const ONESTOPBIT = 0
Global Const ONE5STOPBITS = 1
Global Const TWOSTOPBITS = 2
Global Const IGNORE = 0
Global Const INFINITE = &HFFFF

'COMM Error Flags
Global Const CE_RXOVER = &H1
Global Const CE_OVERRUN = &H2
Global Const CE_RXPARITY = &H4
Global Const CE_FRAME = &H8
Global Const CE_BREAK = &H10
Global Const CE_CTSTO = &H20
Global Const CE_DSRTO = &H40
Global Const CE_RLSDTO = &H80
Global Const CE_TXFULL = &H100
Global Const CE_PTO = &H200
Global Const CE_IOE = &H400
Global Const CE_DNS = &H800
Global Const CE_OOP = &H1000
Global Const CE_MODE = &H8000
Global Const IE_BADID = (-1)
Global Const IE_OPEN = (-2)
Global Const IE_NOPEN = (-3)
Global Const IE_MEMORY = (-4)
Global Const IE_DEFAULT = (-5)
Global Const IE_HARDWARE = (-10)
Global Const IE_BYTESIZE = (-11)
Global Const IE_BAUDRATE = (-12)

'COMM Events
Global Const EV_RXCHAR = &H1
Global Const EV_RXFLAG = &H2
Global Const EV_TXEMPTY = &H4
Global Const EV_CTS = &H8
Global Const EV_DSR = &H10
Global Const EV_RLSD = &H20
Global Const EV_BREAK = &H40
Global Const EV_ERR = &H80
Global Const EV_RING = &H100
Global Const EV_PERR = &H200
Global Const EV_CTSS = &H400
Global Const EV_DSRS = &H800
Global Const EV_RLSDS = &H1000


'COMM Escape Functions
Global Const SETXOFF = 1 'Simulate XOFF received
Global Const SETXON = 2 'Simulate XON received
Global Const SETRTS = 3 'Set RTS high
Global Const CLRRTS = 4 'Set RTS low
Global Const SETDTR = 5 'Set DTR high
Global Const CLRDTR = 6 'Set DTR low
Global Const RESETDEV = 7 'Reset device if possible
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

'COMM notifications on WM_COMMNOTIFY messages
Global Const CN_RECEIVE = &H1
Global Const CN_TRANSMIT = &H2
Global Const CN_EVENT = &H4

'COMM status flags
Global Const CSTF_CTSHOLD = &H1
Global Const CSTF_DSRHOLD = &H2
Global Const CSTF_RLSDHOLD = &H4
Global Const CSTF_XOFFHOLD = &H8
Global Const CSTF_XOFFSENT = &H10
Global Const CSTF_EOF = &H20
Global Const CSTF_TXIM = &H40
Global Const LPTx = &H80

'Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_INDEX = &H3
Global Const HELP_HELPONHELP = &H4
Global Const HELP_SETINDEX = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105
Global Const HELP_MULTIKEY = &H201
Global Const HELP_SETWINPOS = &H203

'Field selection bits
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

'Printer orientation selections
Global Const DMORIENT_PORTRAIT = 1
Global Const DMORIENT_LANDSCAPE = 2

'Paper selections
Global Const DMPAPER_LETTER = 1
Global Const DMPAPER_LETTERSMALL = 2
Global Const DMPAPER_TABLOID = 3
Global Const DMPAPER_LEDGER = 4
Global Const DMPAPER_LEGAL = 5
Global Const DMPAPER_STATEMENT = 6
Global Const DMPAPER_EXECUTIVE = 7
Global Const DMPAPER_A3 = 8
Global Const DMPAPER_A4 = 9
Global Const DMPAPER_A4SMALL = 10
Global Const DMPAPER_A5 = 11
Global Const DMPAPER_B4 = 12
Global Const DMPAPER_B5 = 13
Global Const DMPAPER_FOLIO = 14
Global Const DMPAPER_QUARTO = 15
Global Const DMPAPER_10X14 = 16
Global Const DMPAPER_11X17 = 17
Global Const DMPAPER_NOTE = 18
Global Const DMPAPER_ENV_9 = 19
Global Const DMPAPER_ENV_10 = 20
Global Const DMPAPER_ENV_11 = 21
Global Const DMPAPER_ENV_12 = 22
Global Const DMPAPER_ENV_14 = 23
Global Const DMPAPER_CSHEET = 24
Global Const DMPAPER_DSHEET = 25
Global Const DMPAPER_ESHEET = 26
Global Const DMPAPER_ENV_DL = 27
Global Const DMPAPER_ENV_C5 = 28
Global Const DMPAPER_ENV_C3 = 29
Global Const DMPAPER_ENV_C4 = 30
Global Const DMPAPER_ENV_C6 = 31
Global Const DMPAPER_ENV_C65 = 32
Global Const DMPAPER_ENV_B4 = 33
Global Const DMPAPER_ENV_B5 = 34
Global Const DMPAPER_ENV_B6 = 35
Global Const DMPAPER_ENV_ITALY = 36
Global Const DMPAPER_ENV_MONARCH = 37
Global Const DMPAPER_ENV_PERSONAL = 38
Global Const DMPAPER_FANFOLD_US = 39
Global Const DMPAPER_FANFOLD_STD_GERMAN = 40
Global Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Global Const DMPAPER_USER = 256

'Printer bin selections
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
Global Const DMBIN_USER = 256

'Print qualities
Global Const DMRES_DRAFT = -1
Global Const DMRES_LOW = -2
Global Const DMRES_MEDIUM = -3
Global Const DMRES_HIGH = -4

'Color enable/disable for color printers
Global Const DMCOLOR_MONOCHROME = 1
Global Const DMCOLOR_COLOR = 2

'Printer duplex enable
Global Const DMDUP_SIMPLEX = 1
Global Const DMDUP_VERTICAL = 2
Global Const DMDUP_HORIZONTAL = 3

'TrueType options
Global Const DMTT_BITMAP = 1
Global Const DMTT_DOWNLOAD = 2
Global Const DMTT_SUBDEV = 3

'Device mode function modes
Global Const DM_UPDATE = 1
Global Const DM_COPY = 2
Global Const DM_PROMPT = 4
Global Const DM_MODIFY = 8
Global Const DM_IN_BUFFER = 8
Global Const DM_IN_PROMPT = 4
Global Const DM_OUT_BUFFER = 2
Global Const DM_OUT_DEFAULT = 1

'Device capabilities indices
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

'DC_TRUETYPE bit fields
Global Const DCTT_BITMAP = &H1&
Global Const DCTT_DOWNLOAD = &H2&
Global Const DCTT_SUBDEV = &H4&

'LZ encode constants
Global Const LZERROR_BADINHANDLE = -1
Global Const LZERROR_BADOUTHANDLE = -2
Global Const LZERROR_READ = -3
Global Const LZERROR_WRITE = -4
Global Const LZERROR_GLOBALLOC = -5
Global Const LZERROR_GLOBLOCK = -6
Global Const LZERROR_BADVALUE = -7
Global Const LZERROR_UNKNOWNALG = -8

'Version Control Resources
Global Const VS_FILE_INFO = 16
Global Const VS_VERSION_INFO = 1
Global Const VS_USER_DEFINED = 100

'Version control flags
Global Const VS_FFI_SIGNATURE = &HFEEF04BD
Global Const VS_FFI_STRUCVERSION = &H10000
Global Const VS_FFI_FILEFLAGSMASK = &H3F&
Global Const VS_FF_DEBUG = &H1&
Global Const VS_FF_PRERELEASE = &H2&
Global Const VS_FF_PATCHED = &H4&
Global Const VS_FF_PRIVATEBUILD = &H8&
Global Const VS_FF_INFOINFERRED = &H10&
Global Const VS_FF_SPECIALBUILD = &H20&

'Version control OS flags
Global Const VOS_UNKNOWN = &H0&
Global Const VOS_DOS = &H10000
Global Const VOS_OS216 = &H20000
Global Const VOS_OS232 = &H30000
Global Const VOS_NT = &H40000
Global Const VOS__BASE = &H0&
Global Const VOS__Window16 = &H1&
Global Const VOS__PM16 = &H2&
Global Const VOS__PM32 = &H3&
Global Const VOS__Window32 = &H4&
Global Const VOS_DOS_Window16 = &H10001
Global Const VOS_DOS_Window32 = &H10004
Global Const VOS_OS216_PM16 = &H20002
Global Const VOS_OS232_PM32 = &H30003
Global Const VOS_NT_Window32 = &H40004

'Version control file types
Global Const VFT_UNKNOWN = &H0&
Global Const VFT_APP = &H1&
Global Const VFT_DLL = &H2&
Global Const VFT_DRV = &H3&
Global Const VFT_FONT = &H4&
Global Const VFT_VXD = &H5&
Global Const VFT_STATIC_LIB = &H7&

' VS_VERSION.dwFileSubtype for VFT_Window_DRV
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

' VS_VERSION.dwFileSubtype for VFT_Window_FONT
Global Const VFT2_FONT_RASTER = &H1&
Global Const VFT2_FONT_VECTOR = &H2&
Global Const VFT2_FONT_TRUETYPE = &H3&

'VerFindFile() flags
Global Const VFFF_ISSHAREDFILE = &H1
Global Const VFF_CURNEDEST = &H1
Global Const VFF_FILEINUSE = &H2
Global Const VFF_BUFFTOOSMALL = &H4

'VerInstallFile() flags
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

'WM_POWER window message and DRV_POWER driver notification
Global Const PWR_OK = 1
Global Const PWR_FAIL = (-1)
Global Const PWR_SUSPENDREQUEST = 1
Global Const PWR_SUSPENDRESUME = 2
Global Const PWR_CRITICALRESUME = 3

'Network operation return values
Global Const WN_SUCCESS = 0
Global Const WN_NOT_SUPPORTED = 1
Global Const WN_NET_ERROR = 2
Global Const WN_MORE_DATA = 3
Global Const WN_BAD_POINTER = 4
Global Const WN_BAD_VALUE = 5
Global Const WN_BAD_PASSWORD = 6
Global Const WN_ACCESS_DENIED = 7
Global Const WN_FUNCTION_BUSY = 8
Global Const WN_Window_ERROR = 9
Global Const WN_BAD_USER = &HA
Global Const WN_OUT_OF_MEMORY = &HB
Global Const WN_CANCEL = &HC
Global Const WN_CONTINUE = &HD

'Network Connection errors
Global Const WN_NOT_CONNECTED = &H30
Global Const WN_OPEN_FILES = &H31
Global Const WN_BAD_NETNAME = &H32
Global Const WN_BAD_LOCALNAME = &H33
Global Const WN_ALREADY_CONNECTED = &H34
Global Const WN_DEVICE_ERROR = &H35
Global Const WN_CONNECTION_CLOSED = &H36

'Play Sounds

'                       Globals
'                       -------
Global Xsound As String
Global Info(1 To 5) As String
Global ClickNum As Integer
Global DialogCaption As String
Global Trk
Global TotalTrk
Global Flag
Global NewCount
Global CNCL As Integer

Declare Function SendMsglParamStr Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Long
Declare Function SetCursor Lib "User" (ByVal hCursor As Integer) As Integer
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetWindowtextlength% Lib "User" (ByVal hWnd%)
Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
Declare Function enumchildwindows% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer



'WM_Commands used by event handler


'Window Finding Constants

Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDLAST = 1
Global Const Gw_hwndFirst = 0
'Message Constants
Global Const wm_gettext = &HD
Global Const WM_SETTEXT = &HC
Global Const WM_CHAR = &H102

Global Const LB_GETITEMDATA = (WM_USER + 26)

Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const wm_SYSKEYDOWN = &H104
Global Const wm_Syskeyup = &H105
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const EM_SETREADONLY = (WM_USER + 31)
Global Const EM_LINELENGTH = WM_USER + 17
Global Const EM_SETFONT = WM_USER + 19
Global Const WM_GETFONT = &H31
Global Const WM_CHILDACTIVATE = &H22
Global Const SS_RIGHT = &H2&
Global Const EM_LIMITTEXT = WM_USER + 21
Global Const EM_REPLACESEL = WM_USER + 18
Global Const FW_BOLD = 700
Global Const WM_COMMAND = &H111
Global Const MF_BYCOMMAND = &H0
Global Const FF_ROMAN = 16  '  Variable stroke width, serifed.
 Global Const SS_CENTER = &H1&
 Global Const SET_BACKGROUND_COLOR = 4103
'GetTopWindow
Declare Function GettopWindow% Lib "User" (ByVal hWnd%)
'GetWindow
Declare Function getwindow% Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer)
     
'My globals
Global Chatwindow As Integer
Global txttosay
Global timeleft As Integer

'MM GLOBALS
Global namestring As String
Global cmts As String
Global Keep As Integer
Global soff As Integer
Global sall As Integer
Global sselect As Integer
Global rmv As Integer
Global fulsn As String
     
  '---More----------
Declare Function SetCursor Lib "User" (ByVal hCursor As Integer) As Integer
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function SetWindowText% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function getparent Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function lstrlen Lib "Kernel" (ByVal lpString As Any) As Integer
Declare Function enumchildwindows% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)
Declare Function SetActiveWindow% Lib "User" (ByVal hWnd%)
Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer



'WM_Commands used by event handler


'Message Constants
'GetTopWindow
Declare Function GettopWindow% Lib "User" (ByVal hWnd%)
'GetWindow
Declare Function getwindow% Lib "User" (ByVal hWnd As Integer, ByVal wCmd As Integer)
     

'MM GLOBALS
Global Recording
Global Requester
Global TimeStarted

Declare Function AOLGetList% Lib "311.dll" (ByVal Index%, ByVal Buf$)

Sub addroom (Lst As ListBox)


For Index% = 0 To 23
namez$ = String$(256, " ")
Ret = AOLGetList(Index%, namez$) & ErB$
If Len(Trim$(namez$)) <= 1 Then GoTo Croop
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)

Lst.AddItem namez$
Next Index%
Croop:

End Sub

Function AOhWnd ()
AOhWnd = findwindow("AOL Frame25", "America  Online")
DoEvents

End Function

Sub aolclick (Hand%)
Ret = sendmessagebynum(Hand%, &H201, 0, 0&)
Ret = sendmessagebynum(Hand%, &H202, 0, 0&)
End Sub

Sub aolkeyword (keyword As String)
Dim AOL%, X%, G$, z, Run
Dim hWnds() As Integer
AOL% = findwindow("AOL Frame25", 0&)
Run = runmenu(2, 5)
Do
For Chewy = 1 To 25
    DoEvents
Next Chewy
DaNewWinda% = GetActiveWindow()
X% = findchildbytitle(AOL%, "Keyword")
Loop Until X%


EditBox% = findchildbyclass(X%, "_AOL_EDIT")
TextSet EditBox%, keyword
Run = sendmessagebynum(EditBox%, &H102, 13, 0)

End Sub

Sub AOLSendMail (PERSON, Subject, MESSAGE)


AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then
    MsgBox "Must Be Online"
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
subjec% = getwindow(subjt%, 2)
mess% = findchildbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz%, WM_SETTEXT, 0, PERSON)
a = sendmessagebystring(subjec%, WM_SETTEXT, 0, Subject)
a = sendmessagebystring(mess%, WM_SETTEXT, 0, MESSAGE)



Do: DoEvents
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
mailwin% = findchildbytitle(MDI%, "Compose Mail")
erro% = findchildbytitle(MDI%, "Error")
aolw% = findwindow("#32770", "America Online")
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = sendmessage(aolw%, WM_CLOSE, 0, 0)
a = sendmessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
If erro% <> 0 Then
a = sendmessage(erro%, WM_CLOSE, 0, 0)
a = sendmessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub

Sub CenterForm (F As Form)
F.Top = (screen.Height * .85) / 2 - F.Height / 2
F.Left = screen.Width / 2 - F.Width / 2

End Sub

Function ChatRoomName ()
ChatRoomName = windowcaption(FindChatWnd())

End Function

Function childwithstring (Parent As Integer, TitleText As String)
Dim X%
Dim ChildWnd As Integer
Dim MDIhWnd%
Dim AOLChildhWnd%
Dim RetClsName As String * 255
  
MDIhWnd% = Parent
If MDIhWnd% = 0 Then
    chidwithstring = 0
    Exit Function
End If
ChildWnd = getwindow(AOLChildhWnd%, GW_CHILD)
Do
  If InStr(windowcaption(ChildWnd), TitleText) <> 0 Then
      childwithstring = ChildWnd
      Exit Do
  End If
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

Function ClassFromSibling (Sib As Integer, Class$) As Integer
Buf$ = String$(255, 0)
First = getwindow(Sib, 0)
X = GetClassName(First, Buf$, 255)
If Class$ = TrimNuLL(Buf$) Then ClassFromSibling = First: Exit Function

PrevhWnd = First
Do
Buf$ = String$(255, 0)
ThishWnd = getwindow(PrevhWnd, 2)
X = GetClassName(ThishWnd, Buf$, 255)
Debug.Print Buf$
If Class$ = TrimNuLL(Buf$) Then
    ClassFromSibling = ThishWnd: Exit Function
End If
PrevhWnd = ThishWnd
Loop While ThishWnd <> 0

ClassFromSibling = 0

End Function

Sub click (btn)
    X = setfocusapi(btn)
    X = SetActiveWindow(btn)
    SD% = sendmessage(btn, WM_KEYDOWN, VK_SPACE, 0&)
    SU% = sendmessage(btn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub ClickAOLMenu (Menu_String As String, Top_Position As String)
Dim Top_Position_Num As Integer
Dim buffer As String
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
    buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, &H400)
    Trim_Buffer = TrimNuLL(buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, &H400)
    Trim_Buffer = TrimNuLL(buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
Loop
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = sendmessagebynum(AOL, &H111, Get_ID%, 0&)

End Sub

Sub ClickButton (hWnd As Integer)
Dim R
R = sendmessagebynum(getparent(hWnd), &H111, (GetWindowWord(hWnd, (-12))), ByVal CLng(hWnd))
    DoEvents

End Sub

Sub ClickListBox (hWnd As Integer)
Dim R
R = sendmessagebynum(hWnd, &H203, 0, 0&)

End Sub

Sub CloseWin (HAN%)
Dim XZ%
XZ% = sendmessage(HAN%, WM_CLOSE, 0, 0)
End Sub

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

Sub eMail (T$, CC$, Subj$, Mssg$, Attch$, Snd%)

Dim R, X%, C1%, C2%, C3%, C4%, C5%, C6%, C7%, C8%, C9%, C10%, C11%, C12%, C13%, C14%, C15%, C16%, C17%
R = runmenu(3, 1)
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
C15% = getwindow(C14%, 2) '|File:|
C16% = getwindow(C15%, 2) '|(filename)|
C17% = getwindow(C16%, 2) '<Message>
TextSet C10%, CStr(T$)
TextSet C12%, CStr(CC$)
TextSet C14%, CStr(Subj$)
TextSet C17%, CStr(Mssg$)
If Attch$ <> "" Then
    ClickButton (C6%): DoEvents
    z% = WaitForWindow("Attach File", "*")
    TextSet findchildbyclass(z%, "Edit"), CStr(Attch$)
    ClickButton (findchildbytitle(z%, "OK")): DoEvents
End If
If Snd% = 1 Then
    ClickButton (C2%)
End If


End Sub

Sub Explode (frm As Form, CFlag As Integer)
Const STEPS = 150 'Lower Number Draws Faster, Higher Number Slower
Dim FRect As RECT
Dim FWidth, FHeight As Integer
Dim I, X, y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush

' If CFlag = True, then explode from center of form, otherwise
' explode from upper left corner.
    GetWindowRect frm.hWnd, FRect
    FWidth = (FRect.Right - FRect.Left)
    FHeight = FRect.Bottom - FRect.Top
    
' Create brush with Form's background color.
    hScreen = GetDC(0)
    Brush = CreateSolidBrush(frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    
' Draw rectangles in larger sizes filling in the area to be occupied
' by the form.
    For I = 1 To STEPS
	cx = FWidth * (I / STEPS)
	cy = FHeight * (I / STEPS)
	If CFlag Then
	    X = FRect.Left + (FWidth - cx) / 2
	    y = FRect.Top + (FHeight - cy) / 2
	Else
	    X = FRect.Left
	    y = FRect.Top
	End If
	Rectangle hScreen, X, y, X + cx, y + cy
    Next I
    
' Release the device context to free memory.
' Make the Form visible

    If ReleaseDC(0, hScreen) = 0 Then
	MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    DeleteObject (Brush)
    frm.Show

End Sub

Function FindAOLChildByTitle (TitleText As String) As Integer
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
  If InStr(windowcaption(ChildWnd), TitleText) <> 0 Then
      FindAOLChildByTitle = ChildWnd
      Exit Do
  End If
  ChildWnd = getwindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
End Function

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

start:
C = findchildbyclass(b, "_AOL_VIEW")
If C = 0 Then GoTo nextwnd
d = findchildbyclass(b, "_AOL_EDIT")
If d = 0 Then GoTo nextwnd
e = findchildbyclass(b, "_AOL_LISTBOX")
If e = 0 Then GoTo nextwnd
'We've found it
findchatroom = b
Exit Function

nextwnd:
b = getnextwindow(b, 2)
If b = getwindow(b, GW_HWNDLAST) Then Exit Function
GoTo start


End Function

Function FindChatWnd () As Integer
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
FindChatWnd = ChatWnd

End Function

'
Function FindFwdWnd () As Integer
Dim hWnds() As Integer
Dim R, X%, I, T, G%
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

'
Function FindMailWnd () As Integer
Dim hWnds() As Integer
Dim R, X%, I, T, G%
X = 0
DoEvents
G% = getwindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Read") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Ignore") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Keep As New") <> 0 Then X = X + 1
DoEvents
If findchildbytitle(G%, "Delete") <> 0 Then X = X + 1
DoEvents
If findchildbyclass(G%, "_AOL_Tree") <> 0 Then X = X + 1
If X = 5 Then GoTo Foundd
DoEvents
G% = getwindow(G%, 2)
DoEvents
X = 0
Loop While G% <> 0


Exit Function
Foundd:
DoEvents
FindMailWnd = G%
Exit Function

End Function

'
Function FindReadWnd () As Integer
Dim hWnds() As Integer
Dim R, X%, I, T, G%
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
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  findsn = Trim(sn$) '_win
  Exit Function
End If
Let countt = 0
dis_win2 = getnextwindow(dis_win2, 2)
If dis_win2 = getwindow(dis_win2, GW_HWNDLAST) Then
   findsn = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

'
Function FindWindowLike (hWndArray() As Integer, ByVal hWndStart As Integer, WindowText As String, Classname As String, ID) As Integer
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

Function GetAOL ()
GetAOL = findwindow("AOL Frame25", "America  Online")
End Function

Function GetControl () As String
ActivehWnd = getfocus()
Dim buffer As String
buffer = String$(255, 0)
X = GetWindowText(ActivehWnd, buffer, 255)
GetControl = TrimNuLL(buffer)

End Function

Function GetCPUType () As String
'Example: text9.text = "Your system's CPU type is: " & sGetCPUType
Dim lWinFlags As Long

    lWinFlags = GetWinFlags()

    If lWinFlags And WF_CPU486 Then
	GetCPUType = "486"
	ElseIf lWinFlags And WF_CPU386 Then
	    GetCPUType = "386"
	    ElseIf lWinFlags And WF_CPU286 Then
		GetCPUType = "286"
		Else
		    GetCPUType = "Other"
    End If

End Function

Function GetFreeGDI () As String
'Example: text5.text = "Free GDI Resources: " & sGetFreeGDI
    GetFreeGDI = Format$(getfreesystemresources(GFSR_GDIRESOURCES)) + "%"

End Function

Function GetFreeSYS () As String
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    GetFreeSYS = Format$(getfreesystemresources(GFSR_SYSTEMRESOURCES)) + "%"

End Function

Function GetFreeUser () As String
'Example: text4.text = "Free User Resources: " & sGetFreeUser
    GetFreeUser = Format$(getfreesystemresources(GFSR_USERRESOURCES)) + "%"

End Function

Function GetSN ()

Dim MDI%
MDI% = findwindow("AOL Frame25", 0&)
Dim welcome%
welcome% = findchildbytitle(MDI%, "Welcome, ")
Dim yourname As String * 255
Dim YourName2$
X = GetWindowText(welcome%, yourname, 255)
yourname = LTrim(RTrim(Trim(yourname)))
yourname = Right$(yourname, Len(yourname) - InStr(yourname, ", "))
YourName2 = Left$(yourname, 10)
YourName2 = Right$(yourname, Len(yourname) - 1)
If InStr(YourName2, "!") <> 0 Then YourName2 = Left$(YourName2, InStr(YourName2, "!") - 1)
GetSN = YourName2$

End Function

Function getuser ()

'Find the welcome window
wlcm% = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "Welcome, ")

'Get the caption of it
dacap = windowcaption(wlcm%)

'Extract the user's screen name
If wlcm% <> 0 Then
    numba% = (InStr(dacap, "!") - 10)
    Pname$ = Mid$(dacap, 10, numba%)
    getuser = Pname$
Else
    getuser = "(unknown)"
End If

End Function

Function GetWinDir () As String
buffer$ = String$(255, 0)
X = GetWindowDirectory(buffer$, 255)
Trm$ = TrimNuLL(buffer$)
If Right$(Trm$, 1) <> "\" Then Trm$ = Trm$ + "\"
GetWinDir = Trm$

End Function

Function GetWindowAct () As String
Focus = getfocus()
ActivehWnd = getparent(Focus)
Dim buffer As String
buffer = String$(255, 0)
I = GetWindowText(ActivehWnd, buffer, 255)
d = buffer
GetWindowAct = TrimNuLL(buffer)

End Function

Function GetWindowFromClass (Parent As Integer, Class$) As Integer
Lst = getwindow(Parent, 5)
this = getwindow(Parent, 0)
in$ = String$(255, 0)
X = GetClassName(this, in$, 255)
Out$ = TrimNuLL(in$)
If Class$ = Out$ Then GoTo found
Do
this = getwindow(Parent, 2)
in$ = String$(255, 0)
X = GetClassName(this, in$, 255)
Out$ = TrimNuLL(in$)
If Class$ = Out$ Then GoTo found
Loop Until this = Lst
GetWindowFromClass = 0
Exit Function
found:
GetWindowFromClass = this

End Function

Function GetWindowhWnd () As Integer
Focus = getfocus()
GetWindowhWnd = getparent(Focus)

End Function

Function GetWinVer () As String
'Example: text2.text = "Window version: " & sGetWinVer
Dim lVer As Long, iWinVer As Integer
    lVer = GetVersion()
    iWinVer = CInt(lVer And &HFFFF&)
    GetWinVer = Format$(iWinVer And &HFF) + "." + Format$(CInt(iWinVer / 256))

End Function

Function GetYours () As String
ActivehWnd = getfocus()
Dim buffer As String
buffer = String$(255, 0)
I = GetWindowText(ActivehWnd, buffer, 255)
d = buffer
GetYours = TrimNuLL(buffer)

End Function

Sub Hit_Menu (Mnu_str$)
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
	    junk% = sendmessagebynum(AOL%, WM_COMMAND, Mnu_ID%, 0)
	End If
    Next Sub_level%
Next Top_Level%
End Sub

Function IFileExists (ByVal sFileName As String) As Integer
'Example: If Not IFileExists("win.com") then...
Dim I As Integer
On Error Resume Next

    I = Len(Dir$(sFileName))
    
    If Err Or I = 0 Then
	IFileExists = False
	Else
	    IFileExists = True
    End If

End Function

Sub Im_off ()

Dim nambox As Integer
Dim imwnd As Integer
Dim txtbx As Integer
Dim btnsend As Integer
AOL = findwindow(0&, "America  Online")
sfdlk = setfocusapi(AOL)
X = runmenu(4, 3)
Do
For I = 1 To 25
DoEvents
Next I
imwnd = findchildbytitle(AOL, "Send Instant Message")
Loop Until imwnd
DoEvents
nambox = findchildbyclass(imwnd, "_AOL_Edit")
Call TextSet(nambox, "$IM_OFF")
txtbx = getnextwindow(nambox, GW_HWNDNEXT)
Call TextSet(txtbx, "RiPP: By BucK And ShaLraTH")
btnsend = findchildbytitle(imwnd, "Send")
Call click(btnsend)
Do
    DoEvents
    msger% = GetActiveWindow()
    OKbuton% = findchildbytitle(msger%, "OK")
Loop Until OKbuton%
Call click(OKbuton%)
ADSKF = setfocusapi(imwnd)
SendKeys "^{F4}", True
End Sub

Sub Im_on ()
Dim nambox As Integer
Dim imwnd As Integer
Dim txtbx As Integer
Dim btnsend As Integer
AOL = findwindow(0&, "America  Online")
sfdlk = setfocusapi(AOL)
X = runmenu(4, 3)
Do
For I = 1 To 25
DoEvents
Next I
imwnd = findchildbytitle(AOL, "Send Instant Message")
Loop Until imwnd
DoEvents
nambox = findchildbyclass(imwnd, "_AOL_Edit")
Call TextSet(nambox, "$IM_ON")
txtbx = getnextwindow(nambox, GW_HWNDNEXT)
Call TextSet(txtbx, "RiPP: By BucK And ShaLrAtH")
btnsend = findchildbytitle(imwnd, "Send")
Call click(btnsend)
Do
    DoEvents
    msger% = GetActiveWindow()
    OKbuton% = findchildbytitle(msger%, "OK")
Loop Until OKbuton%
Call click(OKbuton%)
ADSKF = setfocusapi(imwnd)
SendKeys "^{F4}", True
End Sub

Function IM_Send (who$, WhatToSay$)

Dim nambox As Integer
Dim imwnd As Integer
Dim txtbx As Integer
Dim btnsend As Integer
AOL = findwindow(0&, "America  Online")
sfdlk = setfocusapi(AOL)
X = runmenu(4, 3)
Do
For I = 1 To 25
DoEvents
Next I
imwnd = findchildbytitle(AOL, "Send Instant Message")
Loop Until imwnd
DoEvents
nambox = findchildbyclass(imwnd, "_AOL_Edit")
Call TextSet(nambox, who$)
txtbx = getnextwindow(nambox, GW_HWNDNEXT)
Call TextSet(txtbx, WhatToSay$)
btnsend = findchildbytitle(imwnd, "Send")
Call click(btnsend)
poopapoop = Timer
Do
    lala = Timer
    If (lala - poopapoop) > 1.5 Then GoTo SkipDaShiT
    DoEvents
    msger% = GetActiveWindow()
    OKbuton% = findchildbytitle(msger%, "OK")
Loop Until OKbuton%
Call click(OKbuton%)
If OKbuton% Then
    IM_Send = 1
    ADSKF = setfocusapi(imwnd)
    SendKeys "^{F4}", True
End If

SkipDaShiT:


End Function

Function IMtxt ()
    Dim NumChars As Integer
    Dim CRText As String '

    NumChars = sendmessagebynum(ViewWnd, &HE, 0, 0&)
    CRText = Space$(NumChars) 'load up CRText to NumChars
    '
    'Fill up CRText w/ IM text
    Dim X As Integer
    X = sendmessagebystring(ViewWnd, &HD, NumChars, CRText)
    
    If Value = 1 Then
	txtim = Left$(CRText, X)
    Else
      txtim = Right$(Left$(CRText, X), 128)
    End If


End Function

Function ison () As Integer
AOL = findwindow("AOL Frame25", "America  Online")
wlcm = findchildbytitle(AOL, "Welcome, ")
If wlcm <> 0 Then ison = True
If wlcm = 0 Then
    ison = False
    MsgBox "You must be signed on to AOL to do this.", 48, "You are not signed on!"
End If
End Function

Function KTEncrypt (ByVal password, ByVal strng, force%)
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
    ToChange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(ToChange)
    
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

Sub LetterExplode (lbl As Label, maxsize As Integer, spd As Integer)
lbl.Visible = True
Do
For I = 1 To spd
DoEvents
Next I
lbl.FontSize = lbl.FontSize + 1
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize >= maxsize
'MsgBox lbl.FontSize
End Sub

Sub LetterImplode (lbl As Label, minsize As Integer, spd As Integer)
Do
For I = 1 To spd
DoEvents
Next I
lbl.FontSize = lbl.FontSize - 1
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize <= minsize
lbl.Visible = False
'MsgBox lbl.FontSize
End Sub

Sub lstfromstr (Lst As ListBox, strng As String)
cma = InStr(strng, ",")
If cma Then
    dasn = Mid$(strng, 1, cma - 1)
    Lst.AddItem dasn
Else
    Lst.AddItem strng
    Exit Sub
End If
Do
cma2 = InStr(cma + 1, strng, ",")
If cma2 Then
    stpplc = Len(strng) - (cma + 1) - (Len(strng) - cma2)
    dasn = Mid$(strng, cma + 1, stpplc)
    Lst.AddItem dasn
ElseIf cma2 = False Then
    asn = Mid$(strng, cma + 1)
    If asn = "" Then Exit Sub
    Lst.AddItem asn
End If
cma = cma2
Loop Until cma = False
End Sub

Sub mailcount ()

AO% = findwindow(0&, "America  Online")
Hand% = findchildbyclass(AO%, "_AOL_TREE")
buffer = sendmessagebynum(Hand%, LB_GETCOUNT, 0, 0)
If buffer > 1 And buffer <> 550 Then
MsgBox "You have " & buffer & " messages in your Mailbox...", 0
End If
If buffer = 1 Then
MsgBox "You have " & buffer & " message in your Mailbox...", 0
End If
If buffer < 1 Then
MsgBox "You have no messages in your Mailbox...", 0
End If
If buffer = 550 Then
MsgBox "Holy ShiT! Yer BoX Iz FuLL!"
End If

End Sub

Function mailopen ()
If findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail") = False Then
X = runmenu(4, 3)
Do
DoEvents
box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
Loop Until box
List = findchildbyclass(box, "_AOL_Tree")
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
    List = findchildbyclass(box, "_AOL_Tree")
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
End If
mailopen = mailnum
End Function

Function MsgBoxText () As String
Dim TophWnd%
Dim X%
Dim BabyhWnd%
Dim RetClsName As String * 255
Dim MsgText As String, MsgLen As Integer
TophWnd% = GetActiveWindow()
BabyhWnd% = getwindow(TophWnd%, GW_CHILD)
X% = GetClassName(TophWnd%, RetClsName$, 254)
Do
  X% = GetClassName(BabyhWnd%, RetClsName$, 254)
  If InStr(UCase$(RetClsName$), "STATIC") Then
      BabyhWnd% = getwindow(BabyhWnd%, GW_HWNDNEXT)
      MsgLen = GetWindowtextlength(BabyhWnd%)
      MsgText = String$(MsgLen, " ")
      X = GetWindowText(BabyhWnd%, MsgText, MsgLen)
      MsgBoxText = Trim$(MsgText)
      Exit Do
  End If
  BabyhWnd% = getwindow(BabyhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While BabyhWnd% <> 0
End Function

Sub pause (duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub playsound (Xsound As String)
Debug.Print "Xsound  " & Xsound
Dim X%
X% = sndPlaySound(Xsound, 1)

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

Function ReadINI (AppName, KeyName, filename As String) As String
'Example: text4.text = ReadINI("DaProggy", "Lamers Name", app.path + "\Prog.ini")
Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))

End Function

Function runmenu (TopMenuPos As Integer, PopupPos As Integer)
Dim I, mnuhWnd, submnu, AOL, Menuid As Integer
Dim lParam As Long
Dim RioT As String * 500
AOL = findwindow("AOL Frame25", 0&)
MDI = findchildbyclass(AOL, "MDICLIENT")
mnu = GetMenu(AOL)
X = GetMenuString(mnu, 0, RioT, Len(RioT) + 1, &H400)
If InStr(RioT, "*") Then TopMenuPos = TopMenuPos + 1
submnu = GetSubMenu%(mnu, TopMenuPos)
Menuid = GetMenuItemID%(submnu, PopupPos)
I = sendmessagebynum(AOL, &H111, Menuid, 0)
End Function

Sub runmenu2 (Main_Prog As String, Top_Position As String, Menu_String As String)
Dim Top_Position_Num As Integer
Dim buffer As String
Dim Look_For_Menu_String As Integer
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Integer
Dim BY_POSITION As Integer
Dim Get_ID As Integer
Dim Click_Menu_Item As Integer
Dim Menu_Parent As Integer
Dim AOL As Integer
Dim Menu_Handle As Integer
End Sub

Function runmenu3 (TopMenuPos As Integer, PopupPos As Integer)
Dim I, mnuhWnd, submnu, Menuid  As Integer
Dim lParam As Long
Const MF_BYCOMMAND = &H0
mnuhWnd = GetMenu%(findwindow(0&, "America  Online"))
submnu = GetSubMenu%(mnuhWnd, TopMenuPos)
Menuid = GetMenuItemID%(submnu, PopupPos)
lParam = CLng(0) * &H10000 Or MF_BYCOMMAND
I = sendmessagebynum(findwindow(0&, "America  Online"), WM_COMMAND, Menuid, 0&)
End Function

Sub RunMenuByString (ApplicationOfMenu, STringToSearchFor)
'This runs an application's menu by its text.  This
'includes & signs (for underlined letters)

SearchString$ = STringToSearchFor
hMenu = GetMenu(ApplicationOfMenu)
Cnt = GetMenuItemCount(hMenu)
For I = 0 To Cnt - 1
PopUphMenu = GetSubMenu(hMenu, I)
Cnt2 = GetMenuItemCount(PopUphMenu)
For O = 0 To Cnt2 - 1
    hMenuID = GetMenuItemID(PopUphMenu, O)
    MenuString$ = String$(100, " ")
    X = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
	SendtoID = hMenuID
	GoTo Initiate
    End If
Next O
Next I
Initiate:
X = sendmessagebynum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub

Sub send (p0162 As Variant)
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
child% = findchildbyclass(MDI%, "AOL Child")
AolEdit% = findchildbyclass(child%, "_AOL_Edit")
K = sendmessagebystring(AolEdit%, WM_SETTEXT, 0, p0162)
z% = sendmessage(AolEdit%, WM_CHAR, 13, 0)
End Sub

Sub send_IM (who$, txt$)
AOL = findwindow("AOL Frame25", "America  Online")
X = runmenu(5, 6)
Do
DoEvents
imwnd = findchildbytitle(AOL, "Send Instant Message")
Loop Until imwnd
DoEvents
nmbx% = findchildbyclass(imwnd, "_AOL_Edit")
txtbx% = getnextwindow(nmbx%, GW_HWNDNEXT)
Call TextSet(nmbx%, who$)
Call TextSet(txtbx%, txt$)
DoEvents
btnsend = findchildbytitle(imwnd, "Send")
Call click(btnsend)
End Sub

Function sendchatroomtext (ByVal sendstring As String) As Integer
    'Find Send BuTToN
    zzzzz = findwindow("aol frame25", 0&)
    ChatWnd% = FindChatWnd()
    EdtBoxAts% = findchildbyclass(ChatWnd%, "_AOL_EDIT")
    SendBut% = getnextwindow(EdtBoxAts%, GW_HWNDNEXT)
    VC = sendmessagebystring(EdtBoxAts%, WM_SETTEXT, 0, sendstring)
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

Sub sendim (PERSON As String, MESSAGE As String, Snd As Integer)
On Error Resume Next
Dim hWnds() As Integer
Dim rWnds() As Integer
Dim R, L, X, d, IM As Integer
X = findwindow("AOL Frame25", 0&)
z = runmenu(4, 3)
IM = WaitForWindow("Send Instant Message", "AOL Child")
R = FindWindowLike(hWnds(), IM, "*", "_AOL_Edit", Null)
L = findchildbytitle(IM, "Send")
R = sendmessagebystring(hWnds(1), &HC, 0, PERSON)
R = sendmessagebystring(hWnds(2), &HC, 0, MESSAGE)
If Snd <> 0 Then ClickButton (L)
End Sub

Sub SendMail (ToWho, Subject, TheText)
X = SignedOn()
If X = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If
AOL% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(AOL%, "MDIClient")
Hit_Menu "Compose Mail"
Do
    C% = DoEvents()
    COP% = findchildbytitle(MDI%, "Compose Mail")
Loop Until COP% <> 0
edi% = findchildbyclass(COP%, "_AOL_Edit")
X = sendmessagebystring(edi%, WM_SETTEXT, 0&, ToWho)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
edi% = getwindow(edi%, GW_HWNDNEXT)
X = sendmessagebystring(edi%, WM_SETTEXT, 0&, Subject)
RIC% = findchildbyclass(COP%, "RICHCNTL")
If RIC% <> 0 Then
    X = sendmessagebystring(RIC%, WM_SETTEXT, 0&, TheText)
End If
If RIC% = 0 Then
    edi% = getwindow(edi%, GW_HWNDNEXT)
    edi% = getwindow(edi%, GW_HWNDNEXT)
    edi% = getwindow(edi%, GW_HWNDNEXT)
    X = sendmessagebystring(edi%, WM_SETTEXT, 0&, TheText)
End If
ICC% = findchildbyclass(COP%, "_AOL_Icon")
click (ICC%)

End Sub

Sub sendtext (st$)
Dim CCLLPP As String
CCLLPP = Clipboard.GetText()
Clipboard.SetText st$
AppActivate "America  Online"
SendKeys "^v", True
SendKeys "{enter}", True
Clipboard.SetText CCLLPP

End Sub

Function SignedOn ()
AOL% = findwindow("AOL Frame25", 0&)
If AOL% = 0 Then
    SignedOn = 0
    Exit Function
End If
MDI% = findchildbyclass(AOL%, "MDIClient")
WEL% = findchildbytitle(MDI%, "Welcome, ")
If WEL% = 0 Then
    SignedOn = 0
    Exit Function
End If
COMB% = findchildbyclass(WEL%, "_AOL_Combobox")
If COMB% <> 0 Then
    SignedOn = 0
    Exit Function
End If
SignedOn = 2
End Function

Sub sndtxt (EDT%, Text2Send$)
XZ = sendmessagebystring(EDT%, WM_SETTEXT, 0, Text2Send$)
End Sub

Sub stayontop (F As Form)
j% = SetWindowPos(F.hWnd, -1, 0, 0, 0, 0, 3)

End Sub

Function StrFromLst (lstname As ListBox) As String
On Error GoTo daerror
For I = 1 To lstname.ListCount
    dastr = dastr & lstname.List(I - 1) & ","
Next I
StrFromLst = dastr
Exit Function
daerror:
StrFromLst = ""
End Function

Sub TextSet (hWnd As Integer, What As String)
Dim R
R = sendmessagebystring(hWnd, &HC, 0, What)

End Sub

Sub timeout (duration)
starttime = Timer
Do While Timer - starttime < duration
X = DoEvents()
Loop
End Sub

Function TrimNuLL (in$) As String
For X = 1 To Len(in$)
    If (Mid$(in$, X, 1) <> Chr$(0)) Then
	Total$ = Total$ + Mid$(in$, X, 1)
    Else
	GoTo NullDetect
    End If
Next
NullDetect:
TrimNuLL = Total$

End Function

Sub TurnIMs (Setting)
z = runmenu(4, 3)
DoEvents
Clipboard.Clear
Clipboard.SetText "$im_" + Setting
z = getfocus()
DoEvents
Do
ll = getfocus()
DoEvents
Loop Until ll <> z
SendKeys "^V", True
Clipboard.Clear
Clipboard.SetText "Zero is KiNG!"
SendKeys "{TAB}", True
SendKeys "^V", True
SendKeys "{TAB}", True
SendKeys " ", True
Do
pi = GetControl()
DoEvents
Loop Until pi = "OK"
SendKeys " ", True
SendKeys "%-", True
SendKeys "C", True

End Sub

Sub waitformail ()
Dim timr As Long
Dim begin As Double
Dim ending As Double
Dim dfc As Double
timr = GetCurrentTime()
begin = Time
Do
listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
kewl = lpmaxpos%
pause (1.5)
listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
If kewl = lpmaxpos% Then
    listWnd = GetWindowFromClass(getfocus(), "_AOL_Tree")       'Auto "New Mail" Fill Detector
    GetScrollRange listWnd, 1, Lpminpos%, lpmaxpos%
    kewl = lpmaxpos%
    pause (1.5)
    If kewl = lpmaxpos% Then Exit Sub
    End If
 Loop
	'If lstlp = lpmaxpos% Then
	'  Call Pause(7)
	'    If lstlp = lpmaxpos% Then
	'    Exit Do
	'    End If
	'  Else
	'  lstlp = lpmaxpos%
	'End If
'If Cou >= 1400 Then Exit Do   'Change 1600 to more or less
'If GetCurrentTime() >= Timr + (30 * 1000) Then Exit Do
'Call Pause(.0001)             'This number determines the
X = DoEvents()                'number of loops with identical
			      'lpMaxPos's to assume the box
			      'is no longer being updated.

End Sub

Function WaitForWindow (txt As String, clss As String) As Integer
Dim R
Dim hWnds() As Integer

Do
R = FindWindowLike(hWnds(), 0, txt, clss, Null)
DoEvents
Loop While R = 0
WaitForWindow = hWnds(1)

End Function

Sub waitmail ()
Do
box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
timeout (.1)
Loop Until box <> 0
List = findchildbyclass(box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call timeout(.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)


End Sub

Function windowcaption (hWnd As Integer) As String
Dim WindowText As String * 255
Dim GetWinText As Integer
GetWinText = GetWindowText(hWnd, WindowText, 254)
windowcaption = TrimNuLL(WindowText)

End Function

Sub WriteINI (sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim R As Integer
    R = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

