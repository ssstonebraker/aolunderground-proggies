Declare Function extfn1608 Lib "fate01.dll" Alias "AOLgetlist" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnEC8 Lib "User" Alias "appendmenu" () As Integer
Declare Function extfn10C0 Lib "fate02.dll" Alias "FindChildByClass" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfnD78 Lib "User" Alias "getclassname" () As Integer
Declare Function extfnF00 Lib "User" Alias "deletemenu" () As Integer
Declare Function extfnDB0 Lib "User" Alias "DestroyWindow" () As Integer
Declare Sub extsubE58 Lib "User" Alias "DrawMenuBar" ()
Declare Function extfn10F8 Lib "fate02.dll" Alias "FindChildByTitle" (ByVal p1%, ByVal p2$) As Integer
Declare Function extfn1018 Lib "User" Alias "findwindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn12B8 Lib "User" Alias "GetClipboardData" () As Integer
Declare Function extfnCD0 Lib "User" Alias "GetDesktopHwnd" () As Integer
Declare Function extfn1280 Lib "User" Alias "getmenuitemcount" (ByVal p1%) As Integer
Declare Function extfnF38 Lib "User" Alias "getmenustring" (ByVal p1%, ByVal p2%, ByVal p3$, ByVal p4%, ByVal p5%) As Integer
Declare Function extfnD08 Lib "Kernel" Alias "getversion" () As Long
Declare Function extfn1248 Lib "User" Alias "GetWindow" () As Integer
Declare Function extfnDE8 Lib "Kernel" Alias "GetWindowsDirectory" () As Integer
Declare Function extfnE90 Lib "User" Alias "insertmenu" () As Integer
Declare Function extfnFE0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn11A0 Lib "User" Alias "SendMessage" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4$) As Long
Declare Function extfnE20 Lib "User" Alias "modifymenu" () As Integer
Declare Sub extsubF70 Lib "User" Alias "movewindow" ()
Declare Function extfn1168 Lib "User" Alias "GETnextwindow" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1088 Lib "User" Alias "GetParent" (ByVal p1%) As Integer
Declare Function extfn1210 Lib "mmsystem" Alias "sndplaysound" (ByVal p1$, ByVal p2%) As Integer
Declare Function extfn1130 Lib "User" Alias "SetFocus" (ByVal p1%) As Integer
Declare Function extfnD40 Lib "User" Alias "SetMenu" () As Integer
Declare Function extfn11D8 Lib "User" Alias "setparent" () As Integer
Declare Sub extsub1050 Lib "User" Alias "setwindowpos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfnFA8 Lib "User" Alias "GetWindowText" (ByVal p1%, ByVal p2$, ByVal p3%) As Integer


Declare Function extfn10C0 Lib "fate02.dll" Alias "FindChildByClass" (ByVal p1%, ByVal p2$) As Integer


Declare Sub extsub1050 Lib "User" Alias "setwindowpos" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4%, ByVal p5%, ByVal p6%, ByVal p7%)
Declare Function extfn17C8 Lib "User" Alias "getsubmenu" (ByVal p1%, ByVal p2%) As Integer
Declare Function extfn1800 Lib "User" Alias "getmenuitemid" (ByVal p1%, ByVal p2%) As Integer

Declare Function extfn1790 Lib "User" Alias "getmenu" (ByVal p1%) As Integer
Declare Function extfn1018 Lib "User" Alias "findwindow" (ByVal p1 As Any, ByVal p2 As Any) As Integer
Declare Function extfn10F8 Lib "fate02.dll" Alias "FindChildByTitle" (ByVal p1%, ByVal p2$) As Integer

Declare Function GetFocus Lib "User" () As Integer
Declare Function GetMenu Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function GetSubMenu Lib "user" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function ModifyMenu Lib "user" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpString As Any) As Integer
Declare Function GetMenuItemID Lib "user" (ByVal hMenu As Integer, ByVal nPos As Integer) As Integer
Declare Function removemenu Lib "user" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Sub DrawMenuBar Lib "user" (ByVal hWnd As Integer)
Global Const MF_SEPARATOR = &H800
Global Const MF_POPUP = &H10
Global Const MF_BYPOSITION = &H400
Global DidMen As Integer
Global IsDil As Integer
Declare Function CreateMenu% Lib "user" ()
Declare Function AppendMenu% Lib "user" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem&)
Declare Function AppendMenuByString% Lib "user" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "user" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function GetCurrentTime& Lib "user" ()
Declare Function GetFreeSystemResources Lib "user" (ByVal fuSysResource%) As Integer
Declare Function lstrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal maxsize As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer
Declare Sub GetScrollRange Lib "user" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Function ReleaseDC Lib "user" (ByVal hWnd%, ByVal hdc%) As Integer
Type RECT
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
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
Declare Sub SetBKColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Sub ReleaseCapture Lib "user" ()
Declare Sub SetWindowText Lib "user" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetParent% Lib "user" (ByVal hWnd As Integer)
Declare Sub closewindow Lib "user" (ByVal hWnd As Integer)
Declare Sub MoveWindow Lib "user" (ByVal hWnd As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Integer)
Declare Sub bringwindowtotop Lib "user" (ByVal hWnd As Integer)
Declare Function GetCursor Lib "user" () As Integer
Declare Function GetWindowtextlength Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function EnableMenuItem Lib "user" (ByVal hMenu As Integer, ByVal wIDEnableItem As Integer, ByVal wEnable As Integer) As Integer
Declare Function DestroyMenu Lib "user" (ByVal hMenu As Integer) As Integer
Declare Function GetWindowWord Lib "user" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function GetDC Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function getwindowtask Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function enablewindow Lib "user" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function GetActiveWindow Lib "user" () As Integer
Declare Function destroywindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function createwindow% Lib "user" (ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal hwndparent%, ByVal hMenu%, ByVal hInstance%, ByVal lpParam$)
Declare Function CreatePopupMenu Lib "user" () As Integer
Declare Function SetActiveWindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function SetFocusAPI Lib "user" Alias "SetFocus" (ByVal hWnd As Integer) As Integer
Declare Function showwindow Lib "user" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function sndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer
Declare Function waveOutGetNumDevs Lib "MMSystem" () As Integer
Declare Function DeleteMenu Lib "user" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetWindow Lib "user" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer
Declare Function GetWindowText Lib "user" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Declare Function GetClassName Lib "user" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function getnextwindow Lib "user" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetMenuItemCount Lib "user" (ByVal hMenu As Integer) As Integer
Declare Function GetMenuState Lib "user" (ByVal hMenu As Integer, ByVal wId As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetMenuString Lib "user" (ByVal hMenu As Integer, ByVal wIDItem As Integer, ByVal lpString As String, ByVal nMaxCount As Integer, ByVal wFlag As Integer) As Integer
Declare Function findwindow Lib "user" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GettopWindow Lib "user" (ByVal hWnd As Integer) As Integer
Declare Function sendmessagebynum& Lib "user" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
Declare Function sendmessagebystring& Lib "user" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam$)
Declare Function sendmessage Lib "user" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function ExitWindows Lib "user" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
Declare Function AOLGetList% Lib "311.Dll" (ByVal index%, ByVal buf$)
Declare Function AOLGetcombo% Lib "311.Dll" (ByVal index%, ByVal buf$)
Declare Function GetNames Lib "311.Dll" Alias "AOLGetList" (ByVal p1%, ByValp2$) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGUIDE.DLL" (ByVal lpString&)
Declare Function ptGetStringFromAddress Lib "VBMSG.VBX" () As String
Declare Function findchildbytitle% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Declare Function findchildbyclass% Lib "vbwfind.dll" (ByVal Parent%, ByVal Title$)
Global Const WM_USER = &H400
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Global Const LB_DELETESTRING = (WM_USER + 3)
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_FINDSTRING = (WM_USER + 16)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETCOUNT = (WM_USER + 12)
Global Const GW_HWNDNEXT = 2
Global Const GW_CHILD = 5
Global Const WM_CLOSE = &H10
Global Const WM_DESTROY = &H2
Global Const WM_SETTEXT = &HC
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_CHAR = &H102
Global Const WM_COMMAND = &H111
Global Const MB_ICONSTOP = 16
Global Const MB_OK = 0
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const MF_BYCommand = &H0
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Global Const WM_GETTEXT = &HD
Global Const CB_ADDSTRING = (WM_USER + 3)
Global Const CB_GETCOUNT = (WM_USER + 6)
Global Const CB_DELETESTRING = (WM_USER + 4)
Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SW_SHOW = 5
Global Const BM_GETSTATE = WM_USER + 1

'**************************************************'
'*    -=- *     RoYaL CRoWN.1 Bas ©     * -=-     *'
'*    -=- *     By:PeET, Spy & TiNY     * -=-     *'
'*    -=- *     Many Features added     * -=-     *'
'**************************************************'

Declare Function IsWindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Option Compare Text
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

'                   API Subs and Functions
'                   **********************


'*(users) subs and functions*'




Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, LPRect As RECT)
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function GetFocus Lib "User" () As Integer
Declare Function FindWindowByNum Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&) As Integer
Declare Function FindWindowByString Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&) As Integer
Declare Function ExitWindow Lib "User" (ByVal dwReturnCode&, ByVal wReserved%) As Integer
Declare Function Setparent Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
Declare Function GetMessage Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function IsWindowVisible% Lib "User" (ByVal hWnd%)
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lParam&)

'*(kernal)  Functions & Subs*'

Declare Function lstrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal maxsize As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer

'*(GDI) subs & Functions*'

Declare Sub SetBKColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Function GetDeviceCaps Lib "GDI" (ByVal hdc%, ByVal nIndex%) As Integer
Declare Function TextOut Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal lpString As String, ByVal nCount As Integer) As Integer
Declare Function FloodFill Lib "GDI" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal crColor As Long) As Integer
Declare Function SetTextColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer

'*(MMsystem) Subs and Functions *'

Declare Function MciSendString& Lib "MMSystem" (ByVal Cmd$, ByVal Returnstr As Any, ByVal returnlen%, ByVal hcallback%)

'*(VBWFind.Dll)Subs and Functions*'

Declare Function FindChild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal Title$)

'*(APIGuide.Dll) Subs and Functions*'

Declare Sub agCopyData Lib "APIGuide.Dll" (source As Any, dest As Any, ByVal nCount%)
Declare Sub agCopyDataBynum Lib "APIGuide.Dll" Alias "agCopyData" (ByVal source&, ByVal dest&, ByVal nCount%)
Declare Sub agDWordTo2Integers Lib "APIGuide.Dll" (ByVal l&, lw%, lh%)
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


'*(VBMsg.vbx) Subs and Functions*'

Declare Sub ptGetTypeFromAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptCopyTypeToAddress Lib "VBMsg.Vbx" (ByVal lAddress As Long, lpType As Any, ByVal cbBytes As Integer)
Declare Sub ptSetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL)
Declare Function ptGetVariableAddress Lib "VBMsg.Vbx" (Var As Any) As Long
Declare Function ptGetTypeAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (Var As Any) As Long
Declare Function ptGetStringAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (ByVal S As String) As Long
Declare Function ptGetLongAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (l As Long) As Long
Declare Function ptGetIntegerAddress Lib "VBMsg.Vbx" Alias "ptGetVariableAddress" (i As Integer) As Long
Declare Function ptGetIntegerFromAddress Lib "VBMsg.Vbx" (ByVal i As Long) As Integer
Declare Function ptGetLongFromAddress Lib "VBMsg.Vbx" (ByVal l As Long) As Long
Declare Function ptMakelParam Lib "VBMsg.Vbx" (ByVal wLow As Integer, ByVal wHigh As Integer) As Long
Declare Function ptLoWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptHiWord Lib "VBMsg.Vbx" (ByVal lParam As Long) As Integer
Declare Function ptMakeUShort Lib "VBMsg.Vbx" (ByVal LongVal As Long) As Integer
Declare Function ptConvertUShort Lib "VBMsg.Vbx" (ByVal ushortVal As Integer) As Long
Declare Function ptMessagetoText Lib "VBMsg.Vbx" (ByVal uMsgID As Long, ByVal bFlag As Integer) As String
Declare Function ptRecreateControlHwnd Lib "VBMsg.Vbx" (ctl As Control) As Long
Declare Function ptGetControlModel Lib "VBMsg.Vbx" (ctl As Control, lpm As MODEL) As Long
Declare Function ptGetControlName Lib "VBMsg.Vbx" (ctl As Control) As String
'                   Global Shit
'                   ***********
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
'* GetDriveType return values *'
Global Const DRIVE_REMOVABLE = 2
Global Const DRIVE_FIXED = 3
Global Const DRIVE_REMOTE = 4
'*Global Memory Flags*'
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

'*(in addition to GMEM_DISCARDABLE) Flags returned by GlobalFlags*'
Global Const GMEM_DISCARDED = &H4000
Global Const GMEM_LOCKCOUNT = &HFF
'*Predefined Resource Types*'
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
'*GetFreeSystemResources*'
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
'*GetWinFlags*'
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
'*Parameter error checking*'
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
'*KERNEL parameter errors*'
Global Const ERR_BAD_HINSTANCE = &H6020
Global Const ERR_BAD_HMODULE = &H6021
Global Const ERR_BAD_GLOBAL_HANDLE = &H6022
Global Const ERR_BAD_LOCAL_HANDLE = &H6023
Global Const ERR_BAD_ATOM = &H6024
Global Const ERR_BAD_HFILE = &H6025
'*USER parameter errors*'
Global Const ERR_BAD_HWND = &H6040
Global Const ERR_BAD_HMENU = &H6041
Global Const ERR_BAD_HCURSOR = &H6042
Global Const ERR_BAD_HICON = &H6043
Global Const ERR_BAD_HDWP = &H6044
Global Const ERR_BAD_CID = &H6045
Global Const ERR_BAD_HDRVR = &H6046
'*GDI parameter errors*'
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
'*KERNEL errors*'
Global Const ERR_GALLOC = &H1
Global Const ERR_GREALLOC = &H2
Global Const ERR_GLOCK = &H3
Global Const ERR_LALLOC = &H4
Global Const ERR_LREALLOC = &H5
Global Const ERR_LLOCK = &H6
Global Const ERR_ALLOCRES = &H7
Global Const ERR_LOCKRES = &H8
Global Const ERR_LOADMODULE = &H9
'*USER errors*'
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
'*GDI errors*'
Global Const ERR_CREATEDC = &H80
Global Const ERR_CREATEMETA = &H81
Global Const ERR_DELOBJSELECTED = &H82
Global Const ERR_SELBITMAP = &H83
'*Exit Window parameters*'
Global Const EW_RESTARTWindow = &H42
Global Const EW_REBOOTSYSTEM = &H43
'*Stock system bitmaps*'
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
'*Stock system Icons*'
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
'*(Binary) Raster-ops *'
Global Const R2_BLACK = 1 ' *0
Global Const R2_NOTMERGEPEN = 2 '*DPon
Global Const R2_MASKNOTPEN = 3'*DPna
Global Const R2_NOTCOPYPEN = 4'*PN
Global Const R2_MASKPENNOT = 5'*PDna
Global Const R2_NOTMASKPEN = 8'*DPan
Global Const R2_MASKPEN = 9 '*DPa
Global Const R2_NOTXORPEN = 10'*DPxn
Global Const R2_NOP = 11'*D
Global Const R2_MERGENOTPEN = 12'*DPno
Global Const R2_COPYPEN = 13'*P
Global Const R2_MERGEPENNOT = 14'*PDno
Global Const R2_MERGEPEN = 15 '*DPo
Global Const R2_WHITE = 16'* 1
'*(Ternary) Raster-ops *'
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
'*StretchBlt() Modes*'
Global Const BLACKONWHITE = 1
Global Const WHITEONBLACK = 2
Global Const COLORONCOLOR = 3
'*PolyFill() Modes*'
Global Const ALTERNATE = 1
Global Const WINDING = 2
'*Text Alignment Options*'
Global Const TA_NOUPDATECP = 0
Global Const TA_UPDATECP = 1
Global Const TA_LEFT = 0
Global Const TA_RIGHT = 2
Global Const TA_CENTER = 6
Global Const TA_TOP = 0
Global Const TA_BOTTOM = 8
Global Const TA_BASELINE = 24
'*ExtTextOut flags*'
Global Const ETO_GRAYED = 1
Global Const ETO_OPAQUE = 2
Global Const ETO_CLIPPED = 4
'*SetMapperFlags*'
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

'constants for Get or SetSystemPaletteUse()
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
Global Const GWW_HWNDPARENT = (-8)
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
Global Const HWND_BROADCAST = -1
Global Const WM_NULL = &H0
Global Const WM_CREATE = &H1
Global Const WM_MOVE = &H3
Global Const WM_SIZE = &H5
Global Const WM_ACTIVATE = &H6
Global Const WM_SETFOCUS = &H7
Global Const WM_KILLFOCUS = &H8
Global Const WM_ENABLE = &HA
Global Const WM_SETREDRAW = &HB


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
Global Const MK_LBUTTON = &H1
Global Const GW_HWNDPREV = 3
Global Const BM_SETCHECK = WM_USER + 1

Sub c (SendString$)
Dim aol%, aoledit%
aol% = findwindow("AOL FRAME25", 0&)
If aol% = 0 Then
Exit Sub
End If
aoledit% = findchildbyclass(aol%, "_AOL_EDIT")
If aoledit% = 0 Then
MsgBox "You have to be in a chatroom dumbfuck!", 56, "Dumbass!"
Exit Sub
End If
x = sendmessagebystring(aoledit%, WM_SETTEXT, 0&, SendString$)
DoEvents
x = sendmessagebynum(aoledit%, WM_CHAR, 13, 0&)
DoEvents
End Sub

Sub Changefocus ()
Dim aa, ab As Integer
aa = GetFocus()
Do
  DoEvents
  ab = GetFocus()
Loop While aa = ab
End Sub

Sub changewav (wav As String)
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub

Function Chathwnd () As Integer

aol% = findwindow("AOL Frame25", 0&)
ChatList% = findchildbyclass(aol%, "_AOL_ListBox")
Chat% = GetParent(ChatList%)
ChatText% = findchildbyclass(Chat%, "_AOL_Edit")
ChatSnd% = getnextwindow(ChatText%, 2)
ThePlaza% = findchildbytitle(Chat%, "The Plaza")
AOLLive% = findchildbytitle(Chat%, "AOL Live")
MemDir% = findchildbytitle(Chat%, "Member Directory")
ListRooms% = findchildbytitle(Chat%, "List Rooms")
If ThePlaza% = 0 Or AOLLive% = 0 Or MemDir% = 0 Or ListRooms% = 0 Then Chathwnd = 0: Exit Function
If ChatText% = 0 Or Chat% = 0 Or ChatList% = 0 Then Chathwnd = 0: Exit Function
Chathwnd = Chat%

End Function

Function ChatRoomName ()


End Function

Function ClassFromSibling (Sib As Integer, Class$) As Integer
Dim buf$
Dim first, x, thishwnd, prevhwnd As Integer

buf$ = String$(255, 0)
first = GetWindow(Sib, 0)
x = GetClassName(first, buf$, 255)
If Class$ = (buf$) Then ClassFromSibling = first: Exit Function

prevhwnd = first
Do
buf$ = String$(255, 0)
thishwnd = GetWindow(prevhwnd, 2)
x = GetClassName(thishwnd, buf$, 255)
Debug.Print buf$
If Class$ = (buf$) Then
    ClassFromSibling = thishwnd: Exit Function
End If
prevhwnd = thishwnd
Loop While thishwnd <> 0

ClassFromSibling = 0

End Function

Sub Click (btn)
    x = SetFocusAPI(btn)
    x = SetActiveWindow(btn)
    SD% = sendmessage(btn, WM_KEYDOWN, VK_SPACE, 0&)
    SU% = sendmessage(btn, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub click1 (button%)
SendNow% = sendmessagebynum(button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = sendmessagebynum(button%, WM_LBUTTONUP, &HD, 0)

End Sub

Sub ClickButton (hWnd As Integer)
Dim r
r = sendmessagebynum(GetParent(hWnd), &H111, (GetWindowWord(hWnd, (-12))), ByVal CLng(hWnd))
    DoEvents

End Sub

Sub ClickKW25 ()
aol% = findwindow("AOL Frame25", 0&)
Tool% = findchildbyclass(aol%, "Aol Toolbar")
Icon1% = findchildbyclass(Tool%, "_Aol_Icon")
Icon2% = GetWindow(Icon1%, 2)
Icon3% = GetWindow(Icon2%, 2)
Icon4% = GetWindow(Icon3%, 2)
Icon5% = GetWindow(Icon4%, 2)
Icon6% = GetWindow(Icon5%, 2)
Icon7% = GetWindow(Icon6%, 2)
Icon8% = GetWindow(Icon7%, 2)
Icon9% = GetWindow(Icon8%, 2)
Icon10% = GetWindow(Icon9%, 2)
Icon11% = GetWindow(Icon10%, 2)
Icon12% = GetWindow(Icon11%, 2)
Icon13% = GetWindow(Icon12%, 2)
z% = sendmessage(Icon13%, WM_LBUTTONDOWN, 0&, 0&)
q% = sendmessage(Icon13%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub ClickKW30 ()
aol% = findwindow("AOL Frame25", 0&)
Tool% = findchildbyclass(aol%, "Aol Toolbar")
Icon1% = findchildbyclass(Tool%, "_Aol_Icon")
Icon2% = GetWindow(Icon1%, 2)
Icon3% = GetWindow(Icon2%, 2)
Icon4% = GetWindow(Icon3%, 2)
Icon5% = GetWindow(Icon4%, 2)
Icon6% = GetWindow(Icon5%, 2)
Icon7% = GetWindow(Icon6%, 2)
Icon8% = GetWindow(Icon7%, 2)
Icon9% = GetWindow(Icon8%, 2)
Icon10% = GetWindow(Icon9%, 2)
Icon11% = GetWindow(Icon10%, 2)
Icon12% = GetWindow(Icon11%, 2)
Icon13% = GetWindow(Icon12%, 2)
Icon14% = GetWindow(Icon13%, 2)
Icon15% = GetWindow(Icon14%, 2)
Icon16% = GetWindow(Icon15%, 2)
Icon17% = GetWindow(Icon16%, 2)
Icon18% = GetWindow(Icon17%, 2)
Icon19% = GetWindow(Icon18%, 2)
z% = sendmessage(Icon18%, WM_LBUTTONDOWN, 0&, 0&)
q% = sendmessage(Icon18%, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub ClosehWnd (hWnd As Integer)
x% = sendmessagebynum(hWnd, WM_CLOSE, 0&, 0&)
End Sub

Sub CountMail ()
AO% = findwindow("AOL Frame25", 0&)
arf = findchildbytitle(AO%, "New Mail")
Hand% = findchildbyclass(arf, "_AOL_TREE")
mails = sendmessagebynum(Hand%, LB_GETCOUNT, 0, 0)
'If mails > 1 Then
'MsgBox "You have " & mails & " messages in your E-Mailbox."
'End If
'If mails = 1 Then
'MsgBox "You have one message in your E-Mailbox."
'E'nd If
'If mails < 1 Then
'MsgBox "You have zero messages in your E-Mailbox"
'End If

End Sub

Sub CountMail1 ()
aol% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(aol%, "MDIClient")
NER% = findchildbytitle(MDI%, "New Mail")
If NER% <> 0 Then GoTo J
TOO% = findchildbyclass(aol%, "AOL Toolbar")
NEE% = findchildbyclass(TOO%, "_AOL_Icon")
d = sendmessagebynum(NEE%, WM_LBUTTONDOWN, 0, 0)
d = sendmessagebynum(NEE%, WM_LBUTTONUP, 0, 0)
Do
G% = DoEvents()
NER% = findchildbytitle(MDI%, "New Mail")
Loop Until NER% <> 0
TRE% = findchildbyclass(NER%, "_AOL_Tree")
Do
U = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
Timeout (2)
F = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
Loop Until U = F
J:
TRE% = findchildbyclass(NER%, "_AOL_Tree")
F = sendmessagebynum(TRE%, LB_GETCOUNT, 0, 0)
MsgBox "You have " & F & " pieces of mail!", 64, "Mail Counter"



End Sub

Sub disablemodal ()
AOM = findwindow("_AOL_Modal", 0&)
x = enablewindow(AOM, 0)
aol = findwindow("AOL Frame25", 0&)
x = enablewindow(aol, 1)
End Sub

Sub ED ()

End Sub

Sub ELD ()

End Sub

Sub EnableAOL ()
Dim bb As Integer
Dim dis_win As Integer
CessPit = enablewindow(findwindow("AOL Frame25", 0&), 1)

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

Sub Enablemodal ()
AOM = findwindow("_AOL_Modal", 0&)
x = enablewindow(AOM, 1)
aol = findwindow("AOL Frame25", 0&)
x = enablewindow(aol, 0)
End Sub

Function Encrypt2 ()
AO% = findwindow("AOL Frame25", 0&)
Cmpse% = findchildbytitle%(AO%, "Compose Mail")
edt% = findchildbyclass%(Cmpse%, "_AOL_Edit")
a% = GetWindow(edt%, GW_HWNDNEXT): CC% = GetWindow(a%, GW_HWNDNEXT)
b% = GetWindow(CC%, GW_HWNDNEXT): Subject% = GetWindow(b%, GW_HWNDNEXT)
Sub1% = GetWindow(Subject%, GW_HWNDNEXT)
'Sub2% = GetWindowtext(SubGW_HWNDNEXT)
Whole$ = ""
For X1 = 1 To 60
 Randomize
   valteger = Int(17 * Rnd + 1)
    If valteger = 1 Then
    vv = "š"
    ElseIf valteger = 2 Then Whole$ = Whole$ + ""
    ElseIf valteger = 3 Then Whole$ = Whole$ + "¢"
    ElseIf valteger = 4 Then Whole$ = Whole$ + "±"
    ElseIf valteger = 5 Then Whole$ = Whole$ + "¥"
    ElseIf valteger = 6 Then Whole$ = Whole$ + "†"
    ElseIf valteger = 7 Then Whole$ = Whole$ + "ø"
    ElseIf valteger = 8 Then Whole$ = Whole$ + "‡"
    ElseIf valteger = 9 Then Whole$ = Whole$ + "µ"
    ElseIf valteger = 10 Then Whole$ = Whole$ + "þ"
    ElseIf valteger = 11 Then Whole$ = Whole$ + "‰"
    ElseIf valteger = 12 Then Whole$ = Whole$ + "¤"
    ElseIf valteger = 13 Then Whole$ = Whole$ + "£"
    ElseIf valteger = 14 Then Whole$ = Whole$ + "÷"
    ElseIf valteger = 15 Then Whole$ = Whole$ + "º"
    ElseIf valteger = 16 Then Whole$ = Whole$ + "©"
    ElseIf valteger = 17 Then Whole$ = Whole$ + "*"
    End If
Next
x% = sendmessagebystring(Subject%, WM_SETTEXT, 0, Whole$)
End Function

Function findchatroom ()
'Finds the handle of the AOL Chatroom by looking for a
'Window with a ListBox (Chat ScreenNames), Edit Box,
'(Where you type chat text), and an _AOL_VIEW.  If another
'AOL window is present that also has these 3 controls, it
'may find the wrong window.  I have never seen another AOL
'window with these 3 controls at once

aol = findwindow("AOL Frame25", 0&)
If aol = 0 Then Exit Function
b = findchildbyclass(aol, "AOL Child")

start:
h = findchildbyclass(b, "_AOL_VIEW")
If h = 0 Then GoTo nextwnd
d = findchildbyclass(b, "_AOL_EDIT")
If d = 0 Then GoTo nextwnd
E = findchildbyclass(b, "_AOL_LISTBOX")
If E = 0 Then GoTo nextwnd
'We've found it
findchatroom = b
Exit Function

nextwnd:
b = getnextwindow(b, 2)
If b = GetWindow(b, GW_HWNDLAST) Then Exit Function
GoTo start


End Function

Function FindCht ()
aol = findwindow("AOL Frame25", "America  Online")
lstrms = findchildbytitle(aol, "List Rooms")
If lstrms = 0 Then FindCht = 0: Exit Function
FindCht = GetParent(lstrms)
End Function

Function findcomposemail ()
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

Function FindFwdWnd () As Integer
Dim HWNDS() As Integer
Dim r, x%, i, T, G%
DoEvents
x = 0
G% = GetWindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Send") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Send Now") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Address" & Chr(13) & "Book") <> 0 Then x = x + 1
If x = 3 Then GoTo Founddddd
DoEvents
G% = GetWindow(G%, 2)
DoEvents
x = 0
Loop While G% <> 0


Exit Function
Founddddd:
DoEvents
FindFwdWnd = G%
Exit Function

End Function

Function FindMailWnd () As Integer
Dim HWNDS() As Integer
Dim r, x%, i, T, G%
x = 0
DoEvents
G% = GetWindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Read") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Ignore") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Keep As New") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Delete") <> 0 Then x = x + 1
DoEvents
If findchildbyclass(G%, "_AOL_Tree") <> 0 Then x = x + 1
If x = 5 Then GoTo Foundd
DoEvents
G% = GetWindow(G%, 2)
DoEvents
x = 0
Loop While G% <> 0


Exit Function
Foundd:
DoEvents
FindMailWnd = G%
Exit Function

End Function

Function FindReadWnd () As Integer
Dim HWNDS() As Integer
Dim r, x%, i, T, G%
x = 0
DoEvents
G% = GetWindow(AOhWnd(), 5)
DoEvents
Do
If findchildbytitle(G%, "Reply") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Forward") <> 0 Then x = x + 1
DoEvents
If findchildbytitle(G%, "Reply to All") <> 0 Then x = x + 1
DoEvents
If x = 3 Then GoTo Founddd
DoEvents
G% = GetWindow(G%, 2)
DoEvents
x = 0
Loop While G% <> 0


Exit Function
Founddd:
DoEvents
FindReadWnd = G%
Exit Function


End Function

Function FindSN ()
'Finds the user's Screen Name...they must be signed on!

Dim dis_win2 As Integer
a = findwindow("AOL Frame25", 0&)
dis_win2 = findchildbyclass(a, "AOL Child")

begin_find_SN:

bb$ = WindowCaption(dis_win2)
    If Left(bb$, 9) = "Welcome, " Then Let countt = countt + 1
If countt = 1 Then
  val1 = InStr(bb$, " ")
  val2 = InStr(bb$, "!")
  Let sn$ = Mid$(bb$, val1 + 1, val2 - val1 - 1)
  FindSN = Trim(sn$) '_win
  Exit Function
End If
Let countt = 0
dis_win2 = getnextwindow(dis_win2, 2)
If dis_win2 = GetWindow(dis_win2, GW_HWNDLAST) Then
   FindSN = 0
   Exit Function
End If

GoTo begin_find_SN

End Function

Function FindWindowLike (hWndArray() As Integer, ByVal hWndStart As Integer, WindowText As String, Classname As String, ID) As Integer
   Dim hWnd As Integer
   Dim sWindowText As String
   Dim sClassname As String
   Dim sID
   Dim r As Integer
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
   hWnd = GetWindow(hWndStart, 5)

   Do Until hWnd = 0
      DoEvents ' Not necessary
      ' Search children by recursion:
      r = FindWindowLike(hWndArray(), hWnd, WindowText, Classname, ID)

      ' Get the window text and class name:
      sWindowText = Space(255)
      r = GetWindowText(hWnd, sWindowText, 255)
      sWindowText = Left(sWindowText, r)
      sClassname = Space(255)
      r = GetClassName(hWnd, sClassname, 255)
      sClassname = Left(sClassname, r)

      ' If window is a child get the ID:
      If GetParent(hWnd) <> 0 Then
         r = GetWindowWord(hWnd, (-12))
         sID = CLng("&H" & Hex(r))
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
      hWnd = GetWindow(hWnd, 2)
   Loop

   ' Decrement recursion counter:
   level = level - 1

   ' Return the number of windows found:
   FindWindowLike = iFound

End Function

Sub getcombonames (lst As ListBox)
lst.Clear
For index% = 0 To 25
names$ = String$(256, " ")
ret = AOLGetcombo(index%, names$) & erb$
If Len(Trim$(names$)) <= 1 Then Exit For
names$ = Left$(Trim$(names$), Len(Trim(names$)) - 1)
lst.AddItem (names$)
Next index%
End Sub

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
    GetFreeGDI = Format$(GetFreeSystemResources(GFSR_GDIRESOURCES)) + "%"

End Function

Function GetFreeSys () As String
'Example: text3.text = "Free System Resources: " & sGetFreeSys
    GetFreeSys = Format$(GetFreeSystemResources(GFSR_SYSTEMRESOURCES)) + "%"

End Function

Sub getnum (og%, a)
Do
If a = 0 Then Exit Sub
b = 1 + b
og% = GetWindow(og%, GW_HWNDNEXT)
Loop Until b >= a - 1
End Sub

Sub GetRoom (lst As ListBox)
aol% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(aol%, "MDIClient")
List% = findchildbyclass(aol%, "_AOL_Listbox")
room% = GetParent(List%)
For index% = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetList(index%, namez$) & erb$
If Len(Trim$(namez$)) <= 1 Then GoTo end_add
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
ListCheck namez$, lst
Next index%
end_add:

End Sub

Function GetWindowFromClass (Parent As Integer, Class$) As Integer
lst = GetWindow(Parent, 5)
This = GetWindow(Parent, 0)
in$ = String$(255, 0)
x = GetClassName(This, in$, 255)
Out$ = (in$)
If Class$ = Out$ Then GoTo Found
Do
This = GetWindow(Parent, 2)
in$ = String$(255, 0)
x = GetClassName(This, in$, 255)
Out$ = (in$)
If Class$ = Out$ Then GoTo Found
Loop Until This = lst
GetWindowFromClass = 0
Exit Function
Found:
GetWindowFromClass = This

End Function

Function getwintext (hWnd As Integer) As String
lentos = sendmessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
x = sendmessagebystring(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
getwintext = Buffer$
End Function

Function GetYours () As String
activehwnd = GetFocus()
Dim Buffer As String
Buffer = String$(255, 0)
i = GetWindowText(activehwnd, Buffer, 255)
d = Buffer
GetYours = (Buffer)

End Function

Sub hit_menu (Mnu_str$)
aol% = findwindow("AOL Frame25", 0&)
Mnu% = GetMenu(aol%)
MNU_Count% = GetMenuItemCount(Mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(Mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Buff$ = Space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Buff$, 50, MF_BYPOSITION)  'As Integer
        Buff$ = Trim$(Buff$): Buff$ = Left(Buff$, Len(Buff$) - 1)
        If Buff$ = "" Then Buff$ = " -"
        If InStr(Buff$, Mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = sendmessagebynum(aol%, WM_COMMAND, Mnu_ID%, 0)
            Exit Sub
        End If
    Next Sub_level%
Next Top_Level%
End Sub

Sub IM (person$, tt$)
okw = findwindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
run "Send an Instant Message"
Do
aol = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(aol, "Send Instant Message")
txt% = findchildbyclass(bah, "_AOL_Edit")
DoEvents
Loop Until txt% <> 0
nah = findchildbytitle(aol, "To:")
x = sendmessagebystring(nah, WM_SETTEXT, 0, "To: ")
txt% = findchildbyclass(bah, "_AOL_Edit")
Do
rich% = findchildbyclass(bah, "RICHCNTL")
bahqw% = findchildbytitle(bah, "Send")
DoEvents
Timeout (.001)
Loop Until rich% <> 0 Or bahqw% <> 0
If rich% <> 0 Then
Send txt%, person$
Send rich%, tt$
Timeout (.001)
getnum rich%, 1
Click rich%
Else
Send txt%, person$
getnum txt%, 1
Send txt%, tt$
getnum txt%, 1
Click txt%
End If
Timeout (.001)
Do
aol = findwindow("AOL Frame25", 0&)
okw = findwindow("#32770", "America Online")
ba1h = findchildbytitle(aol, "To: ")
DoEvents
Loop Until ba1h = 0 Or okw <> 0
If okw <> 0 Then Waitforok: killwin (bah)
End Sub

Sub IM_OFF ()

AppActivate "America  Online"
Timeout .01
SendKeys "^i"
Timeout .1
Waitforok
SendKeys "$IM_OFF"
Timeout .1
Waitforok
SendKeys "{TAB}"
SendKeys "<-¤->"
SendKeys "{TAB}{ }"
Timeout .1
Waitforok
SendKeys "{ }"

End Sub

Sub IM_ON ()
AppActivate "America  Online"
Timeout .01
SendKeys "^i"
Timeout .1
SendKeys "$IM_ON"
Timeout .1
SendKeys "{TAB}"
SendKeys "<-¤->"
SendKeys "{TAB}{ }"
Timeout .1
SendKeys "{ }"

End Sub

Function imtos (Text)
Dim wer As Variant
If UCase$(Text1) Like UCase$("*PeET*") Then
MsgBox "LoL Nice Try ! Say goodbye to your main index !", 55, "Cant Tos PeET!! "
Exit Function
End If

c "‹«¦ ° Phucker TOSer  v1.o   By:PeET ° ¦»›"
c "‹«¦ ° Phucker TOSer  Found Name ° ¦»›"
c "‹«¦ ° Phucker TOSer  Via IM Viol ° ¦»›"

Do
Dim HWNDS() As Integer
kw% = findchildbytitle(aol%, "Keyword")
Timeout .2
Loop Until kw% <> 0
edi% = findchildbyclass(kw%, "_AOL_Edit")
Tos = sendmessagebystring(edi%, WM_SETTEXT, 0, "KO Help")

aol% = findwindow("AOL Frame25", 0&)
Do
DoEvents
XTY% = findchildbytitle(aol%, "I Need Help!")
Loop Until XTY% <> 0
r = FindWindowLike(HWNDS(), find%, "*", "_AOL_Icon", Null)
k = HWNDS(3)
x% = sendmessagebynum(k, WM_LBUTTONDOWN, 0, 0&)
x% = sendmessagebynum(k, WM_LBUTTONUP, 0, 0&)




Do
DoEvents
XTY% = findchildbytitle(aol%, "Report A Violation")
Loop Until XTY% <> 0
r = FindWindowLike(HWNDS(), find%, "*", "_AOL_Icon", Null)
imv = HWNDS(2)
x% = sendmessagebynum(imv, WM_LBUTTONDOWN, 0, 0&)
x% = sendmessagebynum(imv, WM_LBUTTONUP, 0, 0&)

Do
DoEvents
XP% = findchildbytitle(aol%, "Violations via Instant Messages")
Loop Until XP% <> 0

asnm% = findchildbytitle(XP%, "Date")
 
anx% = GetWindow(asnm%, GW_HWNDNEXT)
asm% = sendmessagebystring(anx%, WM_SETTEXT, 0, "" & Date)

esnm% = findchildbytitle(XP%, "Category of Chat Room")
 
enx% = GetWindow(esnm%, GW_HWNDNEXT)
esm% = sendmessagebystring(enx%, WM_SETTEXT, 0, "Life")


qsnm% = findchildbytitle(XP%, "Time AM/PM")
 
qnx% = GetWindow(qsnm%, GW_HWNDNEXT)
qsm% = sendmessagebystring(qnx%, WM_SETTEXT, 0, "" & Time)

wsnm% = findchildbytitle(XP%, "Chat Room Name")
 
wnx% = GetWindow(wsnm%, GW_HWNDNEXT)
wsm% = sendmessagebystring(wnx%, WM_SETTEXT, 0, "Teen Chat5")






snm% = findchildbytitle(XP%, "CUT and PASTE a copy of the IM here")
 
nx% = GetWindow(snm%, GW_HWNDNEXT)
sm% = sendmessagebystring(nx%, WM_SETTEXT, 0, Text1 + ":" & " " + text2 & Chr(10))
sen% = findchildbytitle(XP%, "Send")
Timeout .1
STAtc% = findchildbytitle(soltos%, "SEND")
Waitforok
 
ax2 = sendmessage(XP%, WM_CLOSE, 0, 0)
SendKeys "{f4}"


c "‹«¦ ° Phucker TOSer  IM Violation Completed ° ¦»›"
c "‹«¦ ° (Whata Whore, Die Bitch!!!) ° ¦»›"



 
Exit Function
    

End Function

Sub IMz ()
Activate
SendKeys "^i"
Waitforok
End Sub

Sub Keywor (where$)
run "Keyword"
Do
aol = findwindow("AOL Frame25", 0&)
bah = findchildbytitle(aol, "Keyword")
daedit = findchildbyclass(bah, "_AOL_Edit")
Timeout (.001)
Loop Until daedit <> 0
daedit = findchildbyclass(bah, "_AOL_Edit")
Send daedit, where$
ico% = findchildbyclass(bah, "_AOL_Icon")
Click ico%
End Sub

Sub Keyword (Shit)
Call Runmenubystring(a, "Keyword...")
    x = findchildbytitle(a, "Keyword") 'Find the KW window
    b = findchildbyclass(x, "_AOL_EDIT")  'Find the edit
    kw = Shit  'Set up Variable
    asdf = sendmessagebystring(b, WM_SETTEXT, 0, kw) 'Put text
    d = findchildbyclass(x, "_AOL_ICON") 'Find "GO"
    Timeout .5
x = sendmessagebynum(d, WM_LBUTTONDOWN, 0, 0&) 'Click Down
pause .5
x = sendmessagebynum(d, WM_LBUTTONUP, 0, 0&)   'Click UP
pause 2

End Sub

Sub killmodal ()
Do
AOM = findwindow("_AOL_Modal", 0&)
killwin (AOM)
DoEvents
Loop Until AOM = 0

End Sub

Sub killwin (windo)
x = sendmessagebynum(windo, WM_CLOSE, 0, 0)
End Sub

Sub KOtos (who$, phrase$)
Keywor ("Kohelp")
Do
aol = findwindow("AOL Frame25", 0&)
mow = findchildbytitle(aol, "I NEED HELP")
ikqn% = findchildbyclass(mow, "_AOL_Icon")
DoEvents
password = findchildbytitle(aol, "Report Password Solicitations")
Click ikqn%
Timeout (.001)
Loop Until password <> 0

Do
password = findchildbytitle(aol, "Report Password Solicitations")
Timeout (.001)
Loop Until password <> 0
whototosbox% = findchildbyclass(password, "_AOL_Edit")
getnum whototosbox%, 3248
Send whototosbox%, who$
getnum whototosbox%, 3008
Send whototosbox%, phrase$
Sendbutton% = findchildbytitle(password, "Send")
Click Sendbutton%
Waitforok
x = sendmessagebynum(guidepager, WM_CLOSE, 0, 0)

End Sub

Function KTEncrypt (ByVal password, ByVal Strng, force%)
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
  If Len(Strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(Strng, 4) + Right$(Strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      Strng = Mid$(Strng, 5, Len(Strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, Strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(Strng, look + 1)) - 1)
          Strng = Left$(Strng, look - 1) + Addin$ + Mid$(Strng, look + 2)
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
    Strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + Strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(Strng)

    'Alter character code
    ToChange = Asc(Mid$(Strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(Strng, Looper, 1) = Chr$(ToChange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, Strng, Chr$(1))
      If look > 0 Then
        Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(Strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(Strng, Chr$(0))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(Strng, Chr$(10))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(Strng, Chr$(13))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(Strng, Chr$(26))
      If look > 0 Then Strng = Left$(Strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(Strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    Strng = Chr$(1) + "KT" + Chr$(1) + Strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(Strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      Strng = Mid$(Strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = Strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Sub LetterExplode (lbl As Label, maxsize As Integer, spd As Integer)
lbl.Visible = True
Do
For i = 1 To spd
DoEvents
Next i
lbl.FontSize = lbl.FontSize + 1
'lbl.Top = lbl.Top + 10
'lbl.Left = lbl.Left + 10
Loop Until lbl.FontSize >= maxsize
'MsgBox lbl.FontSize
End Sub

Sub ListCheck (itm As String, lst As ListBox)

If itm = Sccc Then Exit Sub
If lst.ListCount = 0 Then lst.AddItem itm: Exit Sub

Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub mailman ()
Activate
SendKeys "^m"
Waitforok

End Sub

Function mailopen ()
If findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail") = False Then
Do
DoEvents
Box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
Loop Until Box
List = findchildbyclass(Box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)

Call Timeout(.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
Else
    Box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "NewMail ")
    List = findchildbyclass(Box, "_AOL_Tree")
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
End If
mailopen = mailnum
End Function

Function MYSN ()
aol = findwindow%("AOL Frame25", "America  Online")
Happy% = findchildbytitle%(aol, "Welcome,")
joy$ = getwintext((Happy%))
If InStr(joy$, "!") = 0 Then Exit Function
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
MYSN = yoursn
End Function

Sub NewMail ()
Activate
SendKeys "{F10}{RIGHT}{RIGHT}{RIGHT}{DOWN}{DOWN}{DOWN}"
Timeout .001
SendKeys "{ENTER}"

End Sub

Sub OH ()
a = findwindow("AOL Frame25", 0&)
If a = 0 Then
    MsgBox "Not Signed On", 16, "Error"
    Exit Sub
End If

b = findchildbytitle(a, "Welcome")
If b = 0 Then
    MsgBox "Not Signed On", 16, "Error"
    Exit Sub
End If

d = findchildbyclass(b, "_AOL_COMBOBOX")
If d <> 0 Then
    MsgBox "Not Signed On", 16, "Error"
    Exit Sub
End If
c "                                                                                                                        "
c "                                                                                                                        "
c "                                                                                                                        "
c "                                                                                                                        "
End Sub

Sub OldMail ()
Activate
SendKeys "{F10}{RIGHT}{RIGHT}{RIGHT}{DOWN}{DOWN}{DOWN}{DOWN}"
Timeout .001
SendKeys "{ENTER}"


End Sub

Sub ontop (frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub pause (duratn As Integer)
'This pauses for duratn seconds
Let curent = Timer

Do Until Timer - curent >= duratn
DoEvents
Loop

End Sub

Function PeET () As Variant
PeET = " ]}–»"
End Function

Sub PlayWav (da_file As String)
Waver = sndPlaySound(ByVal CStr(da_file), 1)
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

Sub RoomAdd (lst As ListBox)
aol = findwindow%("AOL Frame25", 0&)
Happy% = findchildbytitle%(aol, "Welcome,")
joy$ = getwintext((Happy%))
If InStr(joy$, "!") = 0 Then Exit Sub
joy$ = Left$(joy$, InStr(joy$, "!") - 1)
yoursn = Right$(joy$, Len(joy$) - InStr(joy$, ",") - 1)
lst.Clear
For index% = 0 To 25
names$ = String$(256, " ")
ret = AOLGetList(index%, names$) & erb$
If Len(Trim$(names$)) <= 1 Then Exit For
names$ = Left$(Trim$(names$), Len(Trim(names$)) - 1)
If names$ = yoursn Then GoTo bob2
lst.AddItem (names$)
bob2:
Next index%


End Sub

Sub run (ByVal menuname As String)
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
    x = GetMenuString%(SmhWnd, SCnt, lpString, lint, MF_BYPOSITION)
    
    If InStr(UCase$(lpString), UCase$(menuname)) Then
        GMnID = GetMenuItemID%(SmhWnd, SCnt)
        ret = sendmessagebynum(AOLx, WM_COMMAND, GMnID, 0)
        DoEvents
        Exit Sub
    End If
    DoEvents
    Next SCnt
Next Cnt

End Sub

Sub Runmenubystring (ApplicationOfMenu, STringToSearchFor)
'This runs an application's menu by its text.  This
'includes & signs (for underlined letters)

SearchString$ = STringToSearchFor
hMenu = GetMenu(ApplicationOfMenu)
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

Sub Send (chatedit, sill$)
sndtext = sendmessagebystring(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub sendim (person As String, message As String, Snd As Integer)

End Sub

Sub Sendmail (ToWho, Subject, TheText)
x = SignedOn()
If x = 0 Then
    MsgBox "You must be signed on to use this feature.", 22, "Not Signed On!"
    Exit Sub
End If

aol% = findwindow("AOL Frame25", 0&)
MDI% = findchildbyclass(aol%, "MDIClient")
hit_menu "Compose Mail"
Do
    E% = DoEvents()
    COP% = findchildbytitle(MDI%, "Compose Mail")
Loop Until COP% <> 0
edi% = findchildbyclass(COP%, "_AOL_Edit")
x = sendmessagebystring(edi%, WM_SETTEXT, 0&, ToWho)
edi% = GetWindow(edi%, GW_HWNDNEXT)
edi% = GetWindow(edi%, GW_HWNDNEXT)
edi% = GetWindow(edi%, GW_HWNDNEXT)
edi% = GetWindow(edi%, GW_HWNDNEXT)
x = sendmessagebystring(edi%, WM_SETTEXT, 0&, Subject)
RIC% = findchildbyclass(COP%, "RICHCNTL")
If RIC% <> 0 Then
    x = sendmessagebystring(RIC%, WM_SETTEXT, 0&, TheText)
End If
If RIC% = 0 Then
    edi% = GetWindow(edi%, GW_HWNDNEXT)
    edi% = GetWindow(edi%, GW_HWNDNEXT)
    edi% = GetWindow(edi%, GW_HWNDNEXT)
    x = sendmessagebystring(edi%, WM_SETTEXT, 0&, TheText)
End If
ICC% = findchildbyclass(COP%, "_AOL_Icon")
Click (ICC%)

End Sub

Sub SentMail ()
Activate
SendKeys "{F10}{RIGHT}{RIGHT}{RIGHT}{DOWN}{DOWN}{DOWN}{DOWN}{DOWN}"
Timeout .001
SendKeys "{ENTER}"

End Sub

Sub Setedit (edt%, txt$)
x = sendmessagebystring(edt%, WM_SETTEXT, 0, txt$)


End Sub

Function SignedOn ()
aol% = findwindow("AOL Frame25", 0&)
If aol% = 0 Then
    SignedOn = 0
    Exit Function
End If
MDI% = findchildbyclass(aol%, "MDIClient")
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

Sub Sound (da_file As String)
Waver = sndPlaySound(ByVal CStr(da_file), 1)

End Sub

Sub Stayontop (frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Function StrFromLst (lstname As ListBox) As String
On Error GoTo daerror
For i = 1 To lstname.ListCount
    dastr = dastr & lstname.List(i - 1) & ","
Next i
StrFromLst = dastr
Exit Function
daerror:
StrFromLst = ""
Exit Function
On Error Resume Next
Dim HWNDS() As Integer
Dim rWnds() As Integer
Dim r, l, x, d, PeET As Integer
x = findwindow("AOL Frame25", 0&)
z = (4)
PeET = WaitForWindow("Send Instant Message", "AOL Child")
r = FindWindowLike(HWNDS(), PeET, "*", "_AOL_Edit", Null)
l = findchildbytitle(PeET, "Send")
r = sendmessagebystring(HWNDS(1), &HC, 0, person)
r = sendmessagebystring(HWNDS(2), &HC, 0, message)
If Snd <> 0 Then ClickButton (l)

End Function

Sub textsend (Hands$)

Dim Rets
Dim aol As Variant
Dim Room1 As Variant
Dim Room2 As Variant
Dim Room3 As Variant
Dim Room4 As Variant
Dim FindChat
aol = findwindow("AOL FRAME25", 0&)
Room1 = findchildbytitle(aol, "List Rooms")
Room2 = findchildbytitle(aol, "AOL Live!")
Room3 = findchildbytitle(aol, "The Plaza")
Room4 = findchildbytitle(aol, "Member Directory")
If Room1 <> 0 And GetParent(Room1) = GetParent(Room2) And GetParent(Room2) = GetParent(Room3) And GetParent(Room3) = GetParent(Room4) And GetParent(Room4) = GetParent(Room1) Then
FindChat = GetParent(Room4)
Dim a
Dim b%
a = FindChat
b% = findchildbyclass(a, "_AOL_EDIT")
Rets = sendmessagebystring(b%, &HC, 0, Hands$)
hfasj = sendmessagebynum(b%, WM_CHAR, 13, 0)
'
'to use type TextSend("text to send goes here")
End If
'
End Sub

Sub Timeout (duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub vbmsgg ()
ChatText$ = agGetStringFromLPSTR$(lParam)
End Sub

Sub waitforchange ()
Dim old As Integer
Dim boy As Integer
old = GetFocus()
boy = GetFocus()
Do Until boy <> old
x = DoEvents()
boy = GetFocus()
Loop

End Sub

Sub Activate ()
a = findwindow("AOL Frame25", 0&)
End Sub

Sub ADD_AOL_LB (itm As String, lst As ListBox)
'Add a list of names to a VB ListBox
'This is usually called by another one of my functions

If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub addcmbchk (itm As String, lst As ComboBox)
If lst.ListCount = 0 Then lst.AddItem itm: Exit Sub
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub addnames (lst As ComboBox)
aol% = findwindow("AOL Frame25", 0&)
WEL% = findchildbytitle(WEL%, "Welcome")
List% = findchildbyclass(aol%, "_AOL_ComboBox")
For index% = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetcombo(index%, namez$) & erb$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addz
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
addcmbchk namez$, lst
Next index%
end_addz:

End Sub

Sub addroom (lst As ListBox)
'This calls a function in 311.dll that retreives the names
'from the AOL listbox.
'PLEASE NOTE THE FOLLOWING:
'1)  I don't support this dll..its hacked and illegal
'2)  This only works on 16 bit versions of AOL
'3)  Its a good idea to bring the chat room to the top
'    of the AOL client before doing this.  Sometimes it
'    gets text from other AOL listboxes


For index% = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetList(index%, namez$) & erb$
If Len(Trim$(namez$)) <= 1 Then GoTo end_addr
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)

ADD_AOL_LB namez$, lst
Next index%
end_addr:

  End Sub

Function AlterAOL (txt As String)
Dim Weq, Ap As Integer
AlterAOL = Ap
End Function

Function AOhWnd ()
AOhWnd = findwindow("AOL Frame25", "America  Online")
DoEvents

End Function

Function aolhwnd () As Integer
aol% = findwindow("AOL Frame25", 0&)
aolhwnd% = aol%
End Function

Sub bustinpriv ()
electric1:
If proG_STAT$ = "OFF" Then
Exit Sub
End If

Let num_try = num_try + 1
a = findwindow("AOL Frame25", 0&)
2

Do
DoEvents
fcr = findchildbytitle(a, "Keyword")
DoEvents
Loop Until fcr <> 0

E1 = findchildbyclass(fcr, "_AOL_Icon")
i1 = findchildbyclass(fcr, "_AOL_Edit")
send_to = sendmessagebystring(i1, WM_SETTEXT, 0, "aol://2719:2-2-api")
tta = sendmessagebynum(E1, WM_LBUTTONDOWN, 0, 0&)
ttr = sendmessagebynum(E1, WM_LBUTTONUP, 0, 0&)
Waitforok
If OK_note = "NO" Then
Exit Sub
End If
Timeout .2
GoTo electric1

End Sub

Sub Waitforok ()

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

Function WaitForWindow (txt As String, clss As String) As Integer
Dim r
Dim HWNDS() As Integer

Do
r = FindWindowLike(HWNDS(), 0, txt, clss, Null)
DoEvents
Loop While r = 0
WaitForWindow = HWNDS(1)

End Function

Sub waitmail ()
Do
Box = findchildbytitle(findwindow("AOL Frame25", "America  Online"), "New Mail")
Timeout (.1)
Loop Until Box <> 0
List = findchildbyclass(Box, "_AOL_Tree")
Do
DoEvents
mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum2 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Call Timeout(.5)
mailnum3 = sendmessage(List, LB_GETCOUNT, 0, 0&)
Loop Until mailnum = mailnum2 And mailnum2 = mailnum3
    mailnum = sendmessage(List, LB_GETCOUNT, 0, 0&)


End Sub

Sub WaitWindow ()
A0L% = findwindow("AOL Frame25", "America  Online")
Dim grl
Dim boy
End Sub

Function WindowCaption (hWndd As Integer)
Dim WindowText As String * 255
Dim getwintext As Integer
getwintext = GetWindowText(hWndd, WindowText, 255)
WindowCaption = (WindowText)
End Function

Function WindowsDirectory () As String
Dim WinPath As String
WinPath = String(145, Chr(0))
WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))

End Function

Sub WriteINI (sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

