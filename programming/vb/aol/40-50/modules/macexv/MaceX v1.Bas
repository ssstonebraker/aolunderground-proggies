Attribute VB_Name = "MaceXv¹"
'MaceXv1 for VB4+ and AOL4.0
'Sup Peeps this is my first bas
' has tons of fades and couple
'other funtions should be good fer a small prog
'Look fer MaceXv2 it'll have a ton of options
'Have any probs
' Email me @ MaceX2000@juno.com
' or visit my site at Http://macex.cjb.net
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

Sub XAOL4_ClearChat()
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Sub XAOL4_InstantMessage(person, Message)
Call XAOL4_Keyword("aol://9293:" & person)
TimeOut (2)
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
Public Sub IMsOff()
Call XAOL4_InstantMessage("$IM_OFF", "MaCE X Ownz U ")
End Sub
Public Sub IMsOn()
Call XAOL4_InstantMessage("$IM_ON", "MaCE X Ownz U ")
End Sub

Public Sub AOLKillWindow(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

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

Sub BoldFadeBlack(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
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
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & f$ & "<FONT COLOR=#777777>" & b$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & h$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & K$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next W
SendChat (PC$)

Sub BoldFadeBlue(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
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
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & f$ & "<FONT COLOR=#0000BE>" & b$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & h$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & K$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeGreen(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
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
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & f$ & "<FONT COLOR=#008800>" & b$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & h$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & K$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendChat (PC$)
End Sub

Sub BoldFadeRed(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
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
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & f$ & "<FONT COLOR=#330000>" & b$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & h$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & K$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendChat (PC$)


End Sub
Sub BoldFadeYellow(thetext As String)
A = Len(thetext)
For W = 1 To A Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
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
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & f$ & "<FONT COLOR=#333300>" & b$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & h$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & K$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendChat (PC$)

End Sub

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

Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub

Sub BoldWavY(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    
Next W
BoldSendChat (p$)

End Sub


Sub BoldWavyChatBlueBlack(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
     
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorbluegree(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
  
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblack(thetext)

G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblue(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4

    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
   
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

Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
A% = GetWindowTextB(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
End Function

Function FindForwardWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
ChildFocus% = GetWindow(MDI%, 5)

While ChildFocus%
listers% = FindChildByTitle(ChildFocus%, "Send Now")
Listere% = FindChildByClass(ChildFocus%, "_AOL_Icon")
Listerb% = FindChildByClass(ChildFocus%, "_AOL_Button")

If listers% <> 0 And Listere% <> 0 And Listerb% <> 0 Then FindForwardWindow = ChildFocus%: Exit Function
ChildFocus% = GetWindow(ChildFocus%, 2)
Wend
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

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function



Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function


Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
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

Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub

Sub Keyword(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call TimeOut(0.05)
ClickIcon (AoIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Function LastChatLine()
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
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

Sub PhreakyAttention(text)

SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
SendChat ("<B>" & text)
SendChat ("<I>" & text)
SendChat ("<U>" & text)
SendChat ("<S>" & text)
SendChat ("<b>¤</b><i> ¤</i><u> ¤</u><s> ¤</s> " & text & " <s>¤</s><u> ¤</u><i> ¤</i><b> ¤</b>")
End Sub


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
BoldBlackBlueBlack (newsent$)


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

ub RespondIM(Message)
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

'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    A = Hex(RGB)
    b = Len(A)
    If b = 5 Then A = "0" & A
    If b = 4 Then A = "00" & A
    If b = 3 Then A = "000" & A
    If b = 2 Then A = "0000" & A
    If b = 1 Then A = "00000" & A
    RGBtoHEX = A
End Function




Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub



Sub SendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

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


Sub SendMail(Recipiants, Subject, Message)

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

Sub PlayWaV(file)
'This will play a WAV file
SoundName$ = file
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub StopWaV()
'This will stop a WAV that is playing
Call PlayWaV(" ")
End Sub


Sub CooLWavyBlueBlack(thetext)
G$ = thetext
A = Len(G$)
For W = 1 To A Step 4
    r$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup><s><i>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
BoldSendChat (p$)
End Sub


Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
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

Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub


Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub

Sub AOLIcon(icon%)
Clck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)

WinCaption = WinTitle$
End Function

