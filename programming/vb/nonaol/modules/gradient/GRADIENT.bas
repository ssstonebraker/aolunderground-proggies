Attribute VB_Name = "modGradient"
'*****************************************
'*       G   r    a   d   i   e   n   t *
'*       Suplied by Brokengroove@usa.net *
'*              A.K.A on AOL : Thunder   *
'*****************************************

' To use call "gradient form me"

Option Explicit
DefInt A-Z

Global Const WM_USER = &H400

' Window Messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5

Public Const WM_ACTIVATE = &H6
'
'  WM_ACTIVATE state values
Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
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
Public Const WM_OTHERWINDOWCREATED = &H42    'no longer supported
Public Const WM_OTHERWINDOWDESTROYED = &H43  'no longer supported
Public Const WM_COMMNOTIFY = &H44            'no longer supported

' notifications passed in low word of lParam on WM_COMMNOTIFY messages
Public Const CN_RECEIVE = &H1
Public Const CN_TRANSMIT = &H2
Public Const CN_EVENT = &H4

Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const WM_POWER = &H48
'
'  wParam for WM_POWER window message and DRV_POWER driver notification

Public Const PWR_OK = 1
Public Const PWR_FAIL = (-1)
Public Const PWR_SUSPENDREQUEST = 1
Public Const PWR_SUSPENDRESUME = 2
Public Const PWR_CRITICALRESUME = 3

Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B

Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9

Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
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
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
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
Public Const WM_CLEAR = &H303
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






Dim GradhWnd As Long, GradIcon As Long
Dim OldGradProc As Long
Dim DrawDC As Long, tmpDC As Long
Dim hRgn As Long
Dim tmpGradFont As Long

Public Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Dim CaptionFont As LOGFONT
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const GWL_STYLE = (-16)
Public Const GCL_WNDPROC = (-24)
Public Const GCL_HICON = (-14)
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000 'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Public Const DT_SINGLELINE = &H20
Public Const DT_VCENTER = &H4
Public Const DT_END_ELLIPSIS = &H8000&
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100 'shorts
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CMETRICS = 44
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXBORDER = 5
Public Const SM_CXCURSOR = 13
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CXFRAME = 32
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CXHSCROLL = 21
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CXICONSPACING = 38
Public Const SM_CXMIN = 28
Public Const SM_CXMINTRACK = 34
Public Const SM_CXSCREEN = 0
Public Const SM_CXSMSIZE = 30
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CXVSCROLL = 2
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const SM_CYCURSOR = 14
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CYFRAME = 33
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYHSCROLL = 3
Public Const SM_CYICON = 12
Public Const SM_CYICONSPACING = 39
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_CYMENU = 15
Public Const SM_CYMIN = 29
Public Const SM_CYMINTRACK = 35
Public Const SM_CYSCREEN = 1
Public Const SM_CYSMSIZE = 31
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_CYVSCROLL = 20
Public Const SM_CYVTHUMB = 9
Public Const SM_DBCSENABLED = 42
Public Const SM_DEBUG = 22
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_MOUSEPRESENT = 19
Public Const SM_PENWINDOWS = 41
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_SWAPBUTTON = 23

Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Const DFC_CAPTION = 1
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONHELP = &H4
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_INACTIVE = &H100

Public Function GradientCallback(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim OldBMP As Long, NewBMP As Long
Dim rcWnd As RECT

Select Case wMsg
Case WM_NCACTIVATE, WM_MDIACTIVATE
  GetWindowRect GradhWnd, rcWnd
  'This will cause a slight flicker because we
  'let Windows paint the caption before we do.
  'We don't let this happen in the WM_NCPAINT message,
  'which is called more often than NCACTIVATE.
  GradientCallback = CallWindowProc(OldGradProc, hwnd, wMsg, wParam, lParam)
  'Create memory DC to draw the titlebar in.
  tmpDC = GetWindowDC(GradhWnd)
  DrawDC = CreateCompatibleDC(tmpDC)
  NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
  OldBMP = SelectObject(DrawDC, NewBMP)
  With rcWnd
   hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
   SelectClipRgn tmpDC, hRgn
   OffsetClipRgn tmpDC, -.Left, -.Top
  End With
  'Find out what color the titlebar needs
  'to be...
  If wParam And GetParent(GradhWnd) = 0 Then
   DrawGradient 0, GetSysColor(COLOR_ACTIVECAPTION)
  ElseIf wParam = GradhWnd And GetParent(GradhWnd) <> 0 Then
   DrawGradient 0, GetSysColor(COLOR_INACTIVECAPTION)
  ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, 0, 0) = GradhWnd Then
   DrawGradient 0, GetSysColor(COLOR_ACTIVECAPTION)
  Else
   DrawGradient 0, GetSysColor(COLOR_INACTIVECAPTION)
  End If
  'Cleanup
  SelectObject DrawDC, OldBMP
  DeleteObject NewBMP
  DeleteDC DrawDC
  OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
  GetClipRgn tmpDC, hRgn
  ReleaseDC GradhWnd, tmpDC
  DeleteObject hRgn
  tmpDC = 0
  Exit Function

Case WM_NCPAINT
  'Basically same as above.
  GetWindowRect GradhWnd, rcWnd
  tmpDC = GetWindowDC(GradhWnd)
  DrawDC = CreateCompatibleDC(tmpDC)
  NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.Left, 50)
  OldBMP = SelectObject(DrawDC, NewBMP)
  With rcWnd
   hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
   SelectClipRgn tmpDC, hRgn
   OffsetClipRgn tmpDC, -.Left, -.Top
  End With
  'Get the color to paint the caption with.
  If GetActiveWindow() = GradhWnd Then
   DrawGradient 0, GetSysColor(COLOR_ACTIVECAPTION)
  ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, 0, 0) = GradhWnd Then
   DrawGradient 0, GetSysColor(COLOR_ACTIVECAPTION)
  Else
   DrawGradient 0, GetSysColor(COLOR_INACTIVECAPTION)
  End If
  'Cleanup
  SelectObject DrawDC, OldBMP
  DeleteObject NewBMP
  DeleteDC DrawDC
  OffsetClipRgn tmpDC, rcWnd.Left, rcWnd.Top
  GetClipRgn tmpDC, hRgn
  'Call the old proc. This will only
  'draw the titlebar's min/max/close buttons
  'because we told it not to do the rest (this
  'eliminates flicker.
  GradientCallback = CallWindowProc(OldGradProc, hwnd, WM_NCPAINT, hRgn, lParam)
  ReleaseDC GradhWnd, tmpDC
  DeleteObject hRgn
  tmpDC = 0
  Exit Function
Case WM_SIZE
  'Whoa, we need to paint the caption.
  If hwnd = GradhWnd Then SendMessage GradhWnd, WM_NCPAINT, 0, 0
End Select

GradientCallback = CallWindowProc(OldGradProc, hwnd, wMsg, wParam, lParam)

End Function

Public Sub GradientForm(frm As Form)

If OldGradProc <> 0 Then Exit Sub

GradhWnd = frm.hwnd
GradIcon = frm.Icon
OldGradProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf GradientCallback)
GradientGetCapsFont

End Sub

Public Sub GradientReleaseForm()

If OldGradProc = 0 Or GradhWnd = 0 Then Exit Sub

SetWindowLong GradhWnd, GWL_WNDPROC, OldGradProc
OldGradProc = 0
GradhWnd = 0

End Sub

Public Sub GradientGetCapsFont()



Dim NCM As NONCLIENTMETRICS
Dim lfNew As LOGFONT

NCM.cbSize = Len(NCM)
Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)

If NCM.iCaptionHeight = 0 Then
 CaptionFont.lfHeight = 0
Else
 CaptionFont = NCM.lfCaptionFont
End If

End Sub

Private Function DrawGradient(ByVal Color1 As Long, ByVal Color2 As Long) As Long

Dim i As Integer
Dim DestWidth As Long, DestHeight As Long
Dim StartPnt As Integer, EndPnt As Integer
Dim PixelStep As Long, XBorder As Long
Dim WndRect As RECT
Dim OldFont As Long
Dim fText As String

On Error Resume Next

GetWindowRect GradhWnd, WndRect
With WndRect
 DestWidth = .Right - .Left
End With

'Get height of caption bar
DestHeight = GetSystemMetrics(SM_CYCAPTION)
'Get the text of the form's caption
fText = Space$(255)
Call GetWindowText(GradhWnd, fText, 255)
fText = Trim$(fText)
'Get the width of the border
XBorder = GetSystemMetrics(SM_CXFRAME)
'The width of the area we need to paint:
DestWidth = DestWidth - (XBorder * 2) - (GetSystemMetrics(SM_CXSMSIZE) * 3) + 6
'Where the painting begins:
StartPnt = XBorder
'Where the painting ends:
EndPnt = XBorder + DestWidth - 4

'How many steps do we need to
'paint the titlebar?
PixelStep = DestWidth \ 8
ReDim Colors(PixelStep) As Long
'Create gradient colors in the array
GradateColors Colors(), Color1, Color2

Dim rct As RECT
Dim hBr As Long

With rct
 .Top = XBorder
 .Left = XBorder
 .Right = XBorder + (DestWidth \ PixelStep)
 .Bottom = XBorder + DestHeight - 1
For i = 0 To PixelStep - 1
 'Paint the titlebar in increments, increasing
 'the color index with each iteration.
 hBr = CreateSolidBrush(Colors(i))
 FillRect DrawDC, rct, hBr
 'Cleanup
 DeleteObject hBr
 'Prepare for the next iteration
 OffsetRect rct, (DestWidth \ PixelStep), 0
 If i = PixelStep - 2 Then .Right = EndPnt
Next
If GradIcon <> 0 Then 'Paint the icon
 'Move the caption text's start point over
 'to make room for the icon
 .Left = XBorder + GetSystemMetrics(SM_CXSMSIZE) + 2
 DrawIconEx DrawDC, XBorder + 1, XBorder + 1, GradIcon, GetSystemMetrics(SM_CXSMSIZE) - 2, GetSystemMetrics(SM_CYSMSIZE) - 2, ByVal 0&, ByVal 0&, 2
Else
 'No icon
 .Left = XBorder
End If
'If getting the caption font failed, use the font
'from the gradient caption form.
If CaptionFont.lfHeight = 0 And tmpGradFont = 0 Then
 tmpGradFont = SendMessage(GradhWnd, WM_GETFONT, 0, 0)
ElseIf tmpGradFont = 0 Then
 tmpGradFont = CreateFontIndirect(CaptionFont)
End If
OldFont = SelectObject(DrawDC, tmpGradFont)
'This is like setting FontTransparent on a Form to True:
SetBkMode DrawDC, 1
'Use a white caption, since the background is black
'on the left side
SetTextColor DrawDC, RGB(255, 255, 255)
.Left = .Left + 2
.Right = .Right - 10
'Draw the caption text
DrawText DrawDC, fText, Len(fText) - 1, rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
'Cleanup
SelectObject DrawDC, OldFont
DeleteObject tmpGradFont
tmpGradFont = 0

.Left = XBorder
.Right = .Right + 12
If tmpDC <> 0 Then
 'Blit our work from the memory DC to the form's
 'window DC to finish the job.
 BitBlt tmpDC, .Left, .Top, .Right - .Left - 10, .Bottom - .Top, DrawDC, .Left, .Top, vbSrcCopy
 'Tell windows that we already painted most of
 'the titlebar.
 ExcludeClipRect tmpDC, XBorder, XBorder, .Right - .Left - 8, .Bottom - .Top + 4
End If

End With

End Function

Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)



Dim i As Integer
Dim dblR As Double, dblG As Double, dblB As Double
Dim addR As Double, addG As Double, addB As Double
Dim bckR As Double, bckG As Double, bckB As Double

dblR = CDbl(Color1 And &HFF)
dblG = CDbl(Color1 And &HFF00&) / 255
dblB = CDbl(Color1 And &HFF0000) / &HFF00&
bckR = CDbl(Color2 And &HFF&)
bckG = CDbl(Color2 And &HFF00&) / 255
bckB = CDbl(Color2 And &HFF0000) / &HFF00&

addR = (bckR - dblR) / UBound(Colors)
addG = (bckG - dblG) / UBound(Colors)
addB = (bckB - dblB) / UBound(Colors)

For i = 0 To UBound(Colors)
 dblR = dblR + addR
 dblG = dblG + addG
 dblB = dblB + addB
 If dblR > 255 Then dblR = 255
 If dblG > 255 Then dblG = 255
 If dblB > 255 Then dblB = 255
 If dblR < 0 Then dblR = 0
 If dblG < 0 Then dblG = 0
 If dblG < 0 Then dblB = 0
 Colors(i) = RGB(dblR, dblG, dblB)
Next

End Sub


