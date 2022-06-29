Option Explicit
Dim WinVersion As Integer, SoundAvailable As Integer
Global VisibleFrame As Frame

Global Const TWIPS = 1
Global Const PIXELS = 3
Global Const RES_INFO = 2
Global Const MINIMIZED = 1

Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Declare Function GetWindowsDirectory Lib "Kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetSystemDirectory Lib "kernel" (ByVal P$, ByVal S%) As Integer
Declare Function GetWinFlags Lib "kernel" () As Long
Global Const WF_CPU286 = &H2&
Global Const WF_CPU386 = &H4&
Global Const WF_CPU486 = &H8&
Global Const WF_STANDARD = &H10&
Global Const WF_ENHANCED = &H20&
Global Const WF_80x87 = &H400&

Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetSystemMetrics Lib "User" (ByVal n As Integer) As Integer
Global Const SM_MOUSEPRESENT = 19

Declare Function GetDeviceCaps Lib "GDI" (ByVal hDC%, ByVal nIndex%) As Integer

Declare Function GlobalCompact Lib "kernel" (ByVal flag&) As Long
Declare Function GetFreeSpace Lib "kernel" (ByVal flag%) As Long
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Global Const GFSR_SYSTEMRESOURCES = &H0
Global Const GFSR_GDIRESOURCES = &H1
Global Const GFSR_USERRESOURCES = &H2

Declare Function sndPlaySound Lib "MMSystem" (lpsound As Any, ByVal flag As Integer) As Integer
Declare Function waveOutGetNumDevs Lib "MMSystem" () As Integer

Declare Function TrackPopupMenu Lib "user" (ByVal hMenu%, ByVal wFlags%, ByVal x%, ByVal y%, ByVal r2%, ByVal hWd%, r As Rect) As Integer
Declare Function GetMenu Lib "user" (ByVal hWd%) As Integer
Declare Function GetSubMenu Lib "user" (ByVal hMenu%, ByVal nPos%) As Integer
Declare Function InsertMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Global Const MF_POPUP = &H10
Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800

Declare Function GetDeskTopWindow Lib "User" () As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Sub ReleaseDC Lib "User" (ByVal hWnd%, ByVal hDC%)
Declare Function BitBlt Lib "GDI" (ByVal destDC%, ByVal x%, ByVal y%, ByVal w%, ByVal h%, ByVal srchDC%, ByVal srcX%, ByVal srcY%, ByVal rop&) As Integer
Global Const SRCCOPY = &HCC0020
Global Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046
Global Const SRCAND = &H8800C6

Declare Sub SetWindowPos Lib "User" (ByVal h1%, ByVal h2%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%)
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40

Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer

Function DeviceColors (hDC As Integer) As Long
Const PLANES = 14
Const BITSPIXEL = 12
    DeviceColors = GetDeviceCaps(hDC, PLANES) * 2 ^ GetDeviceCaps(hDC, BITSPIXEL)
End Function

Function DosVersion ()
Dim Ver As Long, DosVer As Long
    Ver = GetVersion()
    DosVer = Ver \ &H10000
    DosVersion = Format((DosVer \ 256) + ((DosVer Mod 256) / 100), "Fixed")
End Function

Function GetSysIni (section, key)
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetPrivateProfileString(section, key, "", retVal, Len(retVal), "System.ini")
    If worked = 0 Then
	GetSysIni = "unknown"
    Else
	GetSysIni = Left(retVal, worked)
    End If
End Function

Function GetWinIni (section, key)
Dim retVal As String, AppName As String, worked As Integer
    retVal = String$(255, 0)
    worked = GetProfileString(section, key, "", retVal, Len(retVal))
    If worked = 0 Then
	GetWinIni = "unknown"
    Else
	GetWinIni = Left(retVal, worked)
    End If
End Function

Function SystemDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    SystemDirectory = Left(WinPath, GetSystemDirectory(WinPath, Len(WinPath)))

End Function

Function WindowsDirectory () As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WindowsDirectory = Left(WinPath, GetWindowsDirectory(WinPath, Len(WinPath)))
End Function

Function WindowsVersion ()
Dim Ver As Long, WinVer As Long
    Ver = GetVersion()
    WinVer = Ver And &HFFFF&
    WindowsVersion = Format((WinVer Mod 256) + ((WinVer \ 256) / 100), "Fixed")
End Function

