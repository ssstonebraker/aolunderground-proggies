Attribute VB_Name = "WinXP"
' sup thanx to AssMaN and Heat for helping on some of
' the subs most of these i had to rewrite due to the
' fact they were not made for windows xp i'm pretty
' sure all of these are working but if you find a bug
' please let me know :
' shaggy@darcfx.com or darcfx.com/shaggy/
Public Const vbBabyBlue = 16560384
Public Const vbOrange = 33023
Public Const vbNavyBlue = 9896450
Public Const vbDarkGreen = 49152
Public Const vbDarkRed = 192
Public Const vbSilver = 80000005
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function sndPlaySoundA Lib "c:\WINDOWS\SYSTEM\WINMM.DLL" (ByVal lpszSoundName$, ByVal ValueFlags As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
      Private Const SPI_GETDRAGFULLWINDOWS = 38
      Private Const SPI_SETDRAGFULLWINDOWS = 37
      Private Const SPIF_SENDWININICHANGE = 2
Private Declare Function CreateEllipticRgn Lib "GDI32" _
 (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
 ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" _
 (ByVal hwnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long
 Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const WM_SETTEXT = &HC
'-------------------
Private Type Registers
  RegBX As Long
  RegDX As Long
  RegCX As Long
  RegAX As Long
  RegDI As Long
  RegSI As Long
  RegFlags As Long
End Type

Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type

Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const VWin32_DIOC_DOS_IOCTL = 1

Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long
Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'-------------------

Public TargetName As String
Public TargetHwnd As Long
Public Caption As String
'-------------------
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SW_MAXIMIZE = 3
Public Const SW_HIDE = 0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const SW_SHOW = 5
Public Const WM_CLOSE = &H10
Public Const SW_MINIMIZE = 6
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = 1
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const VK_SPACE = &H20
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
   Global Const SND_SYNC = &H0
   Global Const SND_ASYNC = &H1
   Global Const SND_NODEFAULT = &H2
   Global Const SND_LOOP = &H8
   Global Const SND_NOSTOP = &H10
Private Target As String
Public Const WM_SYSCOMMAND = &H112
'-------------------
Dim dAngle As Double
Const NUM_TURNS = 36
Const PI = 3.14159265358979
Const CENTER_X = 4000
Const SRCCOPY = &HCC0020
Private Declare Function StretchBlt Lib "GDI32" _
 (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
 ByVal nWidth As Long, ByVal nHeight As Long, _
 ByVal hSrcDC As Long, ByVal XSrc As Long, _
 ByVal YSrc As Long, ByVal nSrcWidth As Long, _
 ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "GDI32" _
 (ByVal hDestDC As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal hSrcDC As Long, _
 ByVal XSrc As Long, ByVal YSrc As Long, _
 ByVal dwRop As Long) As Long
      Option Base 0

      Private Type PALETTEENTRY
         peRed As Byte
         peGreen As Byte
         peBlue As Byte
         peFlags As Byte
      End Type

      Private Type LOGPALETTE
         palVersion As Integer
         palNumEntries As Integer
         palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
      End Type

      Private Type GUID
         Data1 As Long
         Data2 As Integer
         Data3 As Integer
         Data4(7) As Byte
      End Type

      #If Win32 Then

         Private Const RASTERCAPS As Long = 38
         Private Const RC_PALETTE As Long = &H100
         Private Const SIZEPALETTE As Long = 104

         Private Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
         End Type

         Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function CreateCompatibleBitmap Lib "GDI32" ( _
            ByVal hDC As Long, ByVal nWidth As Long, _
            ByVal nHeight As Long) As Long
         Private Declare Function GetDeviceCaps Lib "GDI32" ( _
            ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
         Private Declare Function GetSystemPaletteEntries Lib "GDI32" ( _
            ByVal hDC As Long, ByVal wStartIndex As Long, _
            ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) _
            As Long
         Private Declare Function CreatePalette Lib "GDI32" ( _
            lpLogPalette As LOGPALETTE) As Long
         Private Declare Function SelectObject Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hObject As Long) As Long
         
         Private Declare Function DeleteDC Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function GetForegroundWindow Lib "user32" () _
            As Long
         Private Declare Function SelectPalette Lib "GDI32" ( _
            ByVal hDC As Long, ByVal hPalette As Long, _
            ByVal bForceBackground As Long) As Long
         Private Declare Function RealizePalette Lib "GDI32" ( _
            ByVal hDC As Long) As Long
         Private Declare Function GetWindowDC Lib "user32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetDC Lib "user32" ( _
            ByVal hwnd As Long) As Long
         Private Declare Function GetWindowRect Lib "user32" ( _
            ByVal hwnd As Long, lpRect As RECT) As Long
         Private Declare Function ReleaseDC Lib "user32" ( _
            ByVal hwnd As Long, ByVal hDC As Long) As Long
         Private Declare Function GetDesktopWindow Lib "user32" () As Long

         Private Type PicBmp
            size As Long
            Type As Long
            hBmp As Long
            hPal As Long
            Reserved As Long
         End Type

         Private Declare Function OleCreatePictureIndirect _
            Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

      #ElseIf Win16 Then

         Private Const RASTERCAPS As Integer = 38
         Private Const RC_PALETTE As Integer = &H100
         Private Const SIZEPALETTE As Integer = 104

         Private Type RECT
            Left As Integer
            Top As Integer
            Right As Integer
            Bottom As Integer
         End Type

         Private Declare Function CreateCompatibleDC Lib "GDI" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function CreateCompatibleBitmap Lib "GDI" ( _
            ByVal hDC As Integer, ByVal nWidth As Integer, _
            ByVal nHeight As Integer) As Integer
         Private Declare Function GetDeviceCaps Lib "GDI" ( _
            ByVal hDC As Integer, ByVal iCapabilitiy As Integer) As Integer
         Private Declare Function GetSystemPaletteEntries Lib "GDI" ( _
            ByVal hDC As Integer, ByVal wStartIndex As Integer, _
            ByVal wNumEntries As Integer, _
            lpPaletteEntries As PALETTEENTRY) As Integer
         Private Declare Function CreatePalette Lib "GDI" ( _
            lpLogPalette As LOGPALETTE) As Integer
         Private Declare Function SelectObject Lib "GDI" ( _
            ByVal hDC As Integer, ByVal hObject As Integer) As Integer
         Private Declare Function BitBlt Lib "GDI" ( _
            ByVal hDCDest As Integer, ByVal XDest As Integer, _
            ByVal YDest As Integer, ByVal nWidth As Integer, _
            ByVal nHeight As Integer, ByVal hDCSrc As Integer, _
            ByVal XSrc As Integer, ByVal YSrc As Integer, _
            ByVal dwRop As Long) As Integer
         Private Declare Function DeleteDC Lib "GDI" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function GetForegroundWindow Lib "User" _
            Alias "GetActiveWindow" () As Integer
         Private Declare Function SelectPalette Lib "User" ( _
            ByVal hDC As Integer, ByVal hPalette As Integer, ByVal _
            bForceBackground As Integer) As Integer
         Private Declare Function RealizePalette Lib "User" ( _
            ByVal hDC As Integer) As Integer
         Private Declare Function GetWindowDC Lib "User" ( _
            ByVal hwnd As Integer) As Integer
         Private Declare Function GetDC Lib "User" ( _
            ByVal hwnd As Integer) As Integer
         Private Declare Function GetWindowRect Lib "User" ( _
            ByVal hwnd As Integer, lpRect As RECT) As Integer
         Private Declare Function ReleaseDC Lib "User" ( _
            ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
         Private Declare Function GetDesktopWindow Lib "User" () As Integer

         Private Type PicBmp
            size As Integer
            Type As Integer
            hBmp As Integer
            hPal As Integer
            Reserved As Integer
         End Type

         Private Declare Function OleCreatePictureIndirect _
            Lib "oc25.dll" (PictDesc As PicBmp, RefIID As GUID, _
            ByVal fPictureOwnsHandle As Integer, IPic As IPicture) _
            As Integer
            
            
Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, _
                                                                                                                             ByVal lpBuffer As String, _
                                                                                                                             ByVal nSize As Long) As Long


Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As String, ByVal lpvSource As String, cbLen As Long)

Declare Function GetEnvironmentStrings Lib "kernel32" Alias "GetEnvironmentStringsA" () As Long

Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                                                                                              nSize As Long) As Long

Declare Function GetTickCount Lib "kernel32" () As Long

'----------------------------
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
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

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Public Const JOY_CAL_READ5 = &H400000
Public Const JOY_CAL_READ6 = &H800000
Public Const JOY_CAL_READZONLY = &H1000000
Public Const JOY_CAL_READUONLY = &H4000000
Public Const JOY_CAL_READVONLY = &H8000000
Type JOYINFOEX
        dwSize As Long                 '  size of structure
        dwFlags As Long                 '  flags to indicate what to return
        dwXpos As Long                '  x position
        dwYpos As Long                '  y position
        dwZpos As Long                '  z position
        dwRpos As Long                 '  rudder/4th axis position
        dwUpos As Long                 '  5th axis position
        dwVpos As Long                 '  6th axis position
        dwButtons As Long             '  button states
        dwButtonNumber As Long        '  current button number pressed
        dwPOV As Long                 '  point of view state
        dwReserved1 As Long                 '  reserved for communication between winmm driver
        dwReserved2 As Long                 '  reserved for future expansion
End Type
'-----------------------------
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
      #End If
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const EM_GETSEL = &HB0
Private Const EM_SETSEL = &HB1
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEFROMCHAR = &HC9
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPIB
    X As Long
    Y As Long
End Type

Sub Clkicon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub EjectMedia(Drive As String)
' With this sub you eject removable media that support it, like Zip-drives
'  and CDs.
' this function is questionable to it's acuracy
' Call:
'   EjectMedia "D:"
' where D: is, of course, the letter of the drive to eject.
Dim SecAttr As SECURITY_ATTRIBUTES
Dim ErrorResult
Dim hDevice As Long, Regs As Registers, RB As Long
  
  hDevice = CreateFile("\\.\vwin32", 0, 0, SecAttr, 0, FILE_FLAG_DELETE_ON_CLOSE, 0)
  If hDevice = -1 Then
    ErrorResult = -1
    Exit Sub
  End If
  With Regs
    .RegAX = &H220D
    .RegBX = Asc(Left(Drive, 1)) - 64
    .RegCX = &H849
  End With
  ErrorResult = DeviceIoControl(hDevice, VWin32_DIOC_DOS_IOCTL, Regs, Len(Regs), Regs, Len(Regs), RB, 0)
  ErrorResult = CloseHandle(hDevice)
End Sub

Public Sub Button(but%)
'I placed this on here just in case u wanted it to be used with
'the MouseOverHwnd feature. Its purpose is to click on a buttons's handle
'Example:   Call Button (MouseOverHwnd) <---clicks on whatever button
'                                           the mouse is over
clickicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
clickicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function FindChildByClass(Parent, child As String) As Integer
Dim childfocus As Long, Buffer$, classbuffer%
childfocus = GetWindow(Parent, vbNull)

While childfocus
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByClass = childfocus: Exit Function
childfocus = GetWindow(childfocus, 2)
Wend

End Function
Sub ErrorMsg(Msg)
Call MsgBox("An error has occured." + vbNewLine + "Error Distription: " + Msg + vbNewLine + vbNewLine + "Please retrace your steps on this program to asure" + vbNewLine + "you are doing everything correctly.", 16)
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call MciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Sub CloseCDR()
MciSendString "Set CDAudio Door Closed Wait", _
        0&, 0&, 0&
End Sub



Sub CursorThink(frm As Form)
frm.MousePointer = 11
End Sub

Sub CursorThinkNot(frm As Form)
frm.MousePointer = 0
End Sub








Sub Graphic_DarkenPic(Brightness, picFrom As PictureBox, picTo As PictureBox)
'birghtness must be between 0.1 and 0.5 or it will fail
'NOTE: 0.1 is the most dark and 0.5 is the least darkest
'Example:  Call Graphic_DarkenPic(.3, Picture1, Picture2)
Dim clr As Long
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim X As Integer
Dim Y As Integer
Dim fraction As Single
If Brightness > 0.5 Then MsgBox "A picture can only be darkened at the values between 0.1 and 0.5.", 16, "Winass.bas Runtime Error": Exit Sub
If Brightness > 0.5 Then MsgBox "A picture can only be darkened at the values between 0.1 and 0.5.", 16, "Winass.bas Runtime Error": Exit Sub
    DoEvents

    fraction = CSng(Brightness)

    picTo.AutoRedraw = True
    picTo.Width = picFrom.Width
    picTo.Height = picFrom.Height
    picTo.ScaleMode = vbPixels
    picFrom.ScaleMode = vbPixels

    For Y = 0 To picFrom.ScaleHeight
        For X = 0 To picFrom.ScaleWidth
            ' Get the source pixel's color components.
            clr = picFrom.Point(X, Y)
            r = clr Mod 256
            g = (clr \ 256) Mod 256
            b = clr \ 256 \ 256

            ' Decrease the brightness.
            r = r * fraction
            g = g * fraction
            b = b * fraction

            ' Write the new pixel.
            picTo.PSet (X, Y), RGB(r, g, b)
        Next X
        DoEvents
    Next Y

    ' Make the changes permanent.
    picTo.picture = picTo.Image

End Sub



Sub SetText(win, Txt)

TheText% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Function GetCaption(hwnd)
'returns the caption of "hWnd" window
'example:  Call GetCaption(MouseOverHwnd)
'in this example, the program will retrieve that caption of whatever
'object the mouse is over.
hwndlength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndlength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndlength% + 1))

GetCaption = hwndTitle$
End Function
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CreateBitmapPicture
      '    - Creates a bitmap type Picture object from a bitmap and
      '      palette.
      '
      ' hBmp
      '    - Handle to a bitmap.
      '
      ' hPal
      '    - Handle to a Palette.
      '    - Can be null if the bitmap doesn't use a palette.
      '
      ' Returns
      '    - Returns a Picture object containing the bitmap.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      #If Win32 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Long, _
            ByVal hPal As Long) As picture

            Dim r As Long
      #ElseIf Win16 Then
         Public Function CreateBitmapPicture(ByVal hBmp As Integer, _
            ByVal hPal As Integer) As picture

            Dim r As Integer
      #End If
         Dim pic As PicBmp
         ' IPicture requires a reference to "Standard OLE Types."
         Dim IPic As IPicture
         Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID.
         With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With

         ' Fill Pic with necessary parts.
         With pic
            .size = Len(pic)          ' Length of structure.
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
            .hBmp = hBmp              ' Handle to bitmap.
            .hPal = hPal              ' Handle to palette (may be null).
         End With

         ' Create Picture object.
         r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object.
         Set CreateBitmapPicture = IPic
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureWindow
      '    - Captures any portion of a window.
      '
      ' hWndSrc
      '    - Handle to the window to be captured.
      '
      ' Client
      '    - If True CaptureWindow captures from the client area of the
      '      window.
      '    - If False CaptureWindow captures from the entire window.
      '
      ' LeftSrc, TopSrc, WidthSrc, HeightSrc
      '    - Specify the portion of the window to capture.
      '    - Dimensions need to be specified in pixels.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the specified
      '      portion of the window that was captured.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''
      '
      #If Win32 Then
         Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As picture

            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim r As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long
      #ElseIf Win16 Then
         Public Function CaptureWindow(ByVal hWndSrc As Integer, _
            ByVal client As Boolean, ByVal LeftSrc As Integer, _
            ByVal TopSrc As Integer, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As picture

            Dim hDCMemory As Integer
            Dim hBmp As Integer
            Dim hBmpPrev As Integer
            Dim r As Integer
            Dim hDCSrc As Integer
            Dim hPal As Integer
            Dim hPalPrev As Integer
            Dim RasterCapsScrn As Integer
            Dim HasPaletteScrn As Integer
            Dim PaletteSizeScrn As Integer
      #End If
         Dim LogPal As LOGPALETTE

         ' Depending on the value of Client get the proper device context.
         If client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
         Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window.
         End If

         ' Create a memory device context for the copy process.
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC.
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)

         ' Get screen properties.
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            ' capabilities.
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              ' support.
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette.

         ' If the screen has a palette make a copy and realize it.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette.
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it.
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
         End If

         ' Copy the on-screen image into the memory DC.
         r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)

      ' Remove the new copy of the  on-screen image.
         hBmp = SelectObject(hDCMemory, hBmpPrev)

         ' If the screen has a palette get back the palette that was
         ' selected in previously.
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         End If

         ' Release the device context resources back to the system.
         r = DeleteDC(hDCMemory)
         r = ReleaseDC(hWndSrc, hDCSrc)

         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles. Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureScreen
      '    - Captures the entire screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the screen.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Picture1.picture = CaptureScreen
      Public Function CaptureScreen() As picture
         #If Win32 Then
            Dim hWndScreen As Long
         #ElseIf Win16 Then
            Dim hWndScreen As Integer
         #End If

         ' Get a handle to the desktop window.
         hWndScreen = GetDesktopWindow()

         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object.

         Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureForm
      '    - Captures an entire form including title bar and border.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the entire
      '      form.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Picture1.picture = CaptureForm(Me)
      Public Function CaptureForm(frmSrc As Form) As picture
         ' Call CaptureWindow to capture the entire form given its window
         ' handle and then return the resulting Picture object.
         Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, _
            frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
            frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureClient
      '    - Captures the client area of a form.
      '
      ' frmSrc
      '    - The Form object to capture.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the form's
      '      client area.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Example:  Picture1.picture = CaptureClient(Form2)
      Public Function CaptureClient(frmSrc As Form) As picture
         ' Call CaptureWindow to capture the client area of the form given
         ' its window handle and return the resulting Picture object.
         Set CaptureClient = CaptureWindow(frmSrc.hwnd, True, 0, 0, _
            frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
            frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
      End Function

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' CaptureActiveWindow
      '    - Captures the currently active window on the screen.
      '
      ' Returns
      '    - Returns a Picture object containing a bitmap of the active
      '      window.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      'Example:  Picture1.picture = CaptureActiveWindow
      Public Function CaptureActiveWindow() As picture
         #If Win32 Then
            Dim hWndActive As Long
            Dim r As Long
         #ElseIf Win16 Then
            Dim hWndActive As Integer
            Dim r As Integer
         #End If
         Dim RectActive As RECT

         ' Get a handle to the active/foreground window.
         hWndActive = GetForegroundWindow()

         ' Get the dimensions of the window.
         r = GetWindowRect(hWndActive, RectActive)

         ' Call CaptureWindow to capture the active window given its
         ' handle and return the Resulting Picture object.
      Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
            RectActive.Right - RectActive.Left, _
            RectActive.Bottom - RectActive.Top)
      End Function

Sub OpenDefaultBrowser(URL, Form)
'example: Call OpenDefaultBrowser("http://come.to/nwoonline", Form1)
   Dim ret As Long

  ret = ShellExecute(Form.hwnd, "Open", URL, vbNullString, vbNullString, SW_SHOWMAXIMIZED)
End Sub
Public Sub AddFonts2Combo(Combo As ComboBox)
Dim X As Long
 For X = 0 To Printer.FontCount - 1  'get number of printer fonts
  Combo.AddItem Printer.Fonts(X)   'add font(x) to combo
 Next X                                          'continue until all fonts are added
End Sub
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      ' PrintPictureToFitPage
      '    - Prints a Picture object as big as possible.
      '
      ' Prn
      '    - Destination Printer object.
      '
      ' Pic
      '    - Source Picture object.
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '
      Public Sub PrintPictureToFitPage(Prn As Printer, pic As picture)
         Const vbHiMetric As Integer = 8
         Dim PicRatio As Double
         Dim PrnWidth As Double
         Dim PrnHeight As Double
         Dim PrnRatio As Double
         Dim PrnPicWidth As Double
         Dim PrnPicHeight As Double

         ' Determine if picture should be printed in landscape or portrait
         ' and set the orientation.
         If pic.Height >= pic.Width Then
            Prn.Orientation = vbPRORPortrait   ' Taller than wide.
         Else
            Prn.Orientation = vbPRORLandscape  ' Wider than tall.
         End If

         ' Calculate device independent Width-to-Height ratio for picture.
         PicRatio = pic.Width / pic.Height

         ' Calculate the dimentions of the printable area in HiMetric.
         PrnWidth = Prn.ScaleX(Prn.ScaleWidth, Prn.ScaleMode, vbHiMetric)
         PrnHeight = Prn.ScaleY(Prn.ScaleHeight, Prn.ScaleMode, vbHiMetric)
         ' Calculate device independent Width to Height ratio for printer.
         PrnRatio = PrnWidth / PrnHeight

         ' Scale the output to the printable area.
         If PicRatio >= PrnRatio Then
            ' Scale picture to fit full width of printable area.
            PrnPicWidth = Prn.ScaleX(PrnWidth, vbHiMetric, Prn.ScaleMode)
            PrnPicHeight = Prn.ScaleY(PrnWidth / PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         Else
            ' Scale picture to fit full height of printable area.
            PrnPicHeight = Prn.ScaleY(PrnHeight, vbHiMetric, Prn.ScaleMode)
            PrnPicWidth = Prn.ScaleX(PrnHeight * PicRatio, vbHiMetric, _
               Prn.ScaleMode)
         End If

         ' Print the picture using the PaintPicture method.
         Prn.PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
      End Sub
      '-----------------------------------------------------------



Sub CheckIfMouseMove(label As label)
Static done_before As Boolean
Static last_point As POINTAPI

Dim cur_point As POINTAPI

    ' If we have done this before, compare the
    ' current mouse position to the previous one.
    If done_before Then
        GetCursorPos cur_point
        If (cur_point.X <> last_point.X) Or _
           (cur_point.Y <> last_point.Y) _
        Then
            label.Caption = "True"
        Else
            label.Caption = "False"
        End If
        
        ' Record the cursor position.
        last_point = cur_point
    Else
        done_before = True
        
        ' Just record the cursor position.
        GetCursorPos last_point
    End If
End Sub

Sub CircleizeForm(Form As Form)
Form.Show 'The form!
SetWindowRgn Form.hwnd, _
  CreateEllipticRgn(0, 0, 300, 200), _
  True
End Sub

Sub Graphic_SpinPic(picture, pic1, pic2, pic3, pic4, TrueOrFalse)
'this is a REALLY cool tool to make a picture spin like it's 3d.
'the position of the picture is determined by pic4. For the
'TrueorFalse part,  you could either put true or false,  i prefer
'false because it looks better
'Example
'Call Graphic_SpinPic(PictureMain, Picture1, Picture2, Picture3, Picture4, False)
If TrueOrFalse = True Then
pic1.picture = picture
Else
pic1.Visible = False
End If
pic2.picture = picture
pic3.picture = picture
pic4.picture = picture
pic2.AutoSize = True
pic3.AutoSize = True
pic1.Width = pic2.Width
pic1.Height = pic2.Height
pic4.Width = pic2.Width
pic4.Height = pic2.Height
pic1.Visible = False
pic2.Visible = False
pic3.Visible = False
pic1.AutoRedraw = True
pic2.AutoRedraw = True
pic3.AutoRedraw = True
pic4.AutoRedraw = True
pic2.BorderStyle = 0
pic3.BorderStyle = 0
pic1.BorderStyle = 0
pic4.BorderStyle = 0

pic1.Cls

If Cos(dAngle * PI / 180) >= 0 Then
    Call StretchBlt(pic1.hDC, _
    (pic2.Width - Abs(Cos(dAngle * PI / 180) _
    * pic2.Width)) / (2 * Screen.TwipsPerPixelX), _
    0, Abs(Cos(dAngle * PI / 180) * pic2.Width) _
    / Screen.TwipsPerPixelX, pic2.Height / _
    Screen.TwipsPerPixelY, pic2.hDC, 0, 0, _
    pic2.Width / Screen.TwipsPerPixelX, _
    pic2.Height / Screen.TwipsPerPixelY, SRCCOPY)
ElseIf Cos(dAngle * PI / 180) < 0 Then
    Call StretchBlt(pic1.hDC, _
    (pic3.Width - Abs(Cos(dAngle * PI / 180) _
    * pic3.Width)) / (2 * Screen.TwipsPerPixelX), _
    0, Abs(Cos(dAngle * PI / 180) * pic3.Width) / _
    Screen.TwipsPerPixelX, pic3.Height / _
    Screen.TwipsPerPixelY, pic3.hDC, 0, 0, _
    pic3.Width / Screen.TwipsPerPixelX, _
    pic3.Height / Screen.TwipsPerPixelY, SRCCOPY)
End If

Call BitBlt(pic4.hDC, 0, 0, _
  pic1.Width / Screen.TwipsPerPixelX, _
  pic1.Height / Screen.TwipsPerPixelY, _
  pic1.hDC, 0, 0, SRCCOPY)

pic4.Refresh

'increment angle and make sure it stays
'between 0 and 360

dAngle = dAngle + 360 / NUM_TURNS
dAngle = dAngle Mod 360
End Sub

Private Function IsFullWindowDragOn() As Boolean
          Dim result As Long

          'Call API and check for successful call.
          If SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0&, result, 0&) _
                 <> 0 Then
              'Feature supported now check value of result.
              If result = 0 Then
                  IsFullWindowDragOn = False
              Else
                  IsFullWindowDragOn = True
              End If

          'Call failed, feature not supported.
          Else
              IsFullWindowDragOn = False
          End If
          
End Function



Sub mess(Message, TypeOfMess)
'Simple way of displayin message boxes
'Example:  Call mess ("Visit http://come.to/nwoonline", important)
If TypeOfMess = Error Then
MsgBox Message, 16, App.Title
Exit Sub
End If
If TypeOfMess = important Then
MsgBox Message, 48, App.Title
Exit Sub
End If
If TypeOfMess = critical Then
MsgBox Message, 64, App.Title
Exit Sub
End If
End Sub


Sub NoWindowOutlineToggle()
'Toggles between the motion outline of a window, or the complete motion
'of the entire window itself.
'Example: NoWindowOutlineToggle
Dim result As Long

          'Toggle the setting.
          If IsFullWindowDragOn Then
              'Turn full window drag off.
              result = SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, 0&, _
                 ByVal vbNullString, SPIF_SENDWININICHANGE)
          Else
              'Turn full window drag on.
              result = SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, 1&, _
                 ByVal vbNullString, SPIF_SENDWININICHANGE)
          End If
          
End Sub



Sub Text_GetLineCount(textbox As textbox, label As label)
'Example: Call Text_GetLineCount (Text1, Label1)
'In this example label1 will indicate how many lines are
'present in Text1
Dim LineCount As Long
LineCount = SendMessageLong(textbox.hwnd, EM_GETLINECOUNT, 0, 0&)
    label = label.Caption + CStr(LineCount)
End Sub

Sub Text_GetCharacterPosition(Text As textbox, label As label)
'Example: Call Text_GetCharacterPosition(Text1, Label1)
'In this example label1 will indicate on which number
'character the cursor is blinking on
Dim currLinePos As Long
    overallCursorPos = SendMessageLong(Text.hwnd, EM_GETSEL, 0, 0&) \ &H10000
    shit = overallCursorPos
    label = label.Caption + CStr(shit + 1)
End Sub



Sub TextFilePrint(nameoffile)
'needed
Dim file_name As String

    file_name = nameoffile
    If Right$(file_name, 1) <> "\" Then _
        file_name = file_name & "\"
    file_name = file_name

    PrintTextFile file_name
End Sub

Sub PrintTextFile(fname As String)

Dim fnum As Integer
Dim Txt As String
Dim pos As Integer
Dim para As String
Dim Word As String

    ' Load the file into the string txt.
    fnum = FreeFile
    Open fname For Input As fnum
    Txt = Input(LOF(fnum), fnum)
    Close fnum
    
    ' Print the file.
    Do While Len(Txt) > 0
        ' Get the next paragraph.
        pos = InStr(Txt, vbCrLf)
        If pos = 0 Then
            para = Txt
            Txt = ""
        Else
            para = Left$(Txt, pos - 1)
            Txt = Mid$(Txt, pos + 2)
        End If
        
        ' Print the paragraph.
        Do While Len(para) > 0
            ' Get the next word.
            pos = InStr(para, " ")
            If pos = 0 Then
                Word = para
                para = ""
            Else
                Word = Left$(para, pos)
                para = Mid$(para, pos + 1)
            End If
            
            ' Print the word.
            If Printer.CurrentX + _
                Printer.TextWidth(Word) <= _
                Printer.ScaleWidth _
            Then
                ' The word fits on this line.
                Printer.Print Word;
            Else
                ' Start a new line.
                Printer.Print
                ' Start a new page if needed.
                If Printer.CurrentY + _
                    Printer.Font.size > _
                    Printer.ScaleHeight _
                Then Printer.NewPage
                
                Printer.Print Word;
            End If
        Loop
        ' End the paragraph with a new line.
        Printer.Print
    Loop
    ' Close the document.
    Printer.EndDoc
End Sub

Sub SetNewBlinkRate(NewRate)
If NewRate <> "" Then
    SetCaretBlinkTime CLng(NewRate)
End If
End Sub

Function Text_StripLetter(Txt As String, textbox As textbox, which As String)
'This takes out a certain letter
'Which is the letter you take out(its in number value)
'For example..in the work Khan if I wanted to
'take out the H I would use
'Text_StripLetter("Khan", 2)
txtlen = Len(Txt)
before = Left$(Txt, which - 1)
textbox.Text = before
beforelen = Len(before)
afterthat = txtlen - beforelen - 1
After = Right$(Txt, afterthat)
textbox.Text = After
Text_StripLetter = before & After
End Function
Public Sub TextColor_Blue(Txt As textbox)
'Same procedures as LabelColor_Blue, but replaced with a textbox
Txt.ForeColor = &HFFFF00
Pause 0.1
Txt.ForeColor = &HFF0000
Pause 0.1
Txt.ForeColor = &HC00000
Pause 0.1
Txt.ForeColor = &H800000
Pause 0.1
Txt.ForeColor = &H400000
Pause 0.1
End Sub
Public Sub LabelColor_Blue(label As label)
'Example: LabelColor_Blue Label1
label.ForeColor = &HFFFF00
Pause 0.1
label.ForeColor = &HFF0000
Pause 0.1
label.ForeColor = &HC00000
Pause 0.1
label.ForeColor = &H800000
Pause 0.1
label.ForeColor = &H400000
Pause 0.1
End Sub
Public Sub TextColor_Green(Txt As textbox)
'Same procedures as LabelColor_Green, but replaced with a textbox
Txt.ForeColor = &HFF00&
Pause 0.1
Txt.ForeColor = &HC000&
Pause 0.1
Txt.ForeColor = &H8000&
Pause 0.1
Txt.ForeColor = &H4000&
Pause 0.1
End Sub
Public Sub LabelColor_Green(label As label)
'Example: LabelColor_Green Label1
label.ForeColor = &HFF00&
Pause 0.1
label.ForeColor = &HC000&
Pause 0.1
label.ForeColor = &H8000&
Pause 0.1
label.ForeColor = &H4000&
Pause 0.1
End Sub
Public Sub TextColor_Red(Txt As textbox)
'Same procedures as LabelColor_Red, but replaced with a textbox
Txt.ForeColor = &HFF&
Pause 0.1
Txt.ForeColor = &HC0&
Pause 0.1
Txt.ForeColor = &H80&
Pause 0.1
Txt.ForeColor = &H40&
Pause 0.1
End Sub
Public Sub LabelColor_Red(label As label)
'Example: LabelColor_Red Label1
label.ForeColor = &HFF&
Pause 0.1
label.ForeColor = &HC0&
Pause 0.1
label.ForeColor = &H80&
Pause 0.1
label.ForeColor = &H40&
Pause 0.1

End Sub
Public Sub TextColor_Teal(Txt As textbox)
'Same procedures as LabelColor_Teal, but replaced with a textbox
Txt.ForeColor = &HFFFF00
Pause 0.1
Txt.ForeColor = &HC0C000
Pause 0.1
Txt.ForeColor = &H808000
Pause 0.1
Txt.ForeColor = &H404000
Pause 0.1
End Sub
Public Sub LabelColor_Teal(label As label)
'Example: LabelColor_Teal Label1
label.ForeColor = &HFFFF00
Pause 0.1
label.ForeColor = &HC0C000
Pause 0.1
label.ForeColor = &H808000
Pause 0.1
label.ForeColor = &H404000
Pause 0.1
End Sub
Public Sub TextColor_Yellow(Txt As textbox)
'Same procedures as LabelColor_Yellow, but replaced with a textbox
Txt.ForeColor = &HFFFF&
Pause 0.1
Txt.ForeColor = &HC0C0&
Pause 0.1
Txt.ForeColor = &H8080&
Pause 0.1
Txt.ForeColor = &H4040&
Pause 0.1
End Sub

Public Sub LabelColor_Yellow(label As label)
'Example: LabelColor_Yellow Label1
label.ForeColor = &HFFFF&
Pause 0.1
label.ForeColor = &HC0C0&
Pause 0.1
label.ForeColor = &H8080&
Pause 0.1
label.ForeColor = &H4040&
Pause 0.1
End Sub
Sub PlayMIDI(MIDI As String)
'Example: PlayMIDI ("c:\windows\song.mid")
Dim sn As Long
file$ = MIDI
Snd = mciExecute("play " & file$)
End Sub
Sub IfWinExits(Winname)
'Checks if a window Exists
'Example:  IfWinExists ("Windows Explorer")
    TargetName = Winname
    TargetHwnd = 0
    
    ' Examine the window names.
    EnumWindows AddressOf WindowEnumerator, 0

    ' See if we got an hwnd.
    If TargetHwnd = 0 Then
        MsgBox "The window does not exist."
    Else
        MsgBox "The window exists."
    End If
End Sub







' Return False to stop the enumeration.
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Function WindowEnumerator(ByVal app_hwnd As Long, ByVal lParam As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if the title contains the target.
    If InStr(Title, TargetName) > 0 Then
        ' Save the hwnd and end the enumeration.
        TargetHwnd = app_hwnd
        WindowEnumerator = False
    Else
        ' Continue the enumeration.
        WindowEnumerator = True
    End If
End Function
Public Sub TerminateTask(app_name As String)
'Closes a program by given Window caption
'Example: TerminateTask ("Windows Explorer")
    Target = app_name
    EnumWindows AddressOf EnumCallback, 0
End Sub


Sub CheckForSoundCard()
 Dim rtn As Integer 'declare the needed variables

 rtn = waveOutGetNumDevs() 'check for a sound card

 If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
  MsgBox "Your system supports a sound card."
 Else 'otherwise no sound card found
  MsgBox "Your system cannot play Sound Files."
 End If
End Sub

Sub FallFormDown(frm As Form, Speed)
'Example: FallFormDown Me, 100
'Speed may be whatever u like,  100 is pretty fast
Do
frm.Top = Val(frm.Top) + Speed
Loop Until frm.Top > Screen.Height

End Sub

Sub FallFormUp(frm As Form, Speed)
'Example: FallFormUp Me, 100
'Speed may be whatever u like,  100 is pretty fast
shit = frm.Height - Screen.Height
Do
frm.Top = Val(frm.Top) - Speed
Loop Until frm.Top < shit
End Sub
Sub FallFormLeft(frm As Form, Speed)
'Example: FallFormLeft Me, 100
'Speed may be whatever u like,  100 is pretty fast
shit = frm.Width - Screen.Width
Do
frm.Left = Val(frm.Left) - Speed
Loop Until frm.Left < shit
End Sub
Sub FallFormRight(frm As Form, Speed)
'Example: FallFormRight Me, 100
'Speed may be whatever u like,  100 is pretty fast
Do
frm.Left = Val(frm.Left) + Speed
Loop Until frm.Left > Screen.Width
End Sub
Function ReverseText(Text)
'Example: Text1.text = ReverseText(Text1.text)
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words


End Function


Function Text_Encrypt(strin As String)
'Returns the strin encrypted
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)

If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "A" Then Let NextChr$ = "~"
If NextChr$ = "a" Then Let NextChr$ = "`"
If NextChr$ = "B" Then Let NextChr$ = "!"
If NextChr$ = "C" Then Let NextChr$ = "@"
If NextChr$ = "c" Then Let NextChr$ = "#"
If NextChr$ = "D" Then Let NextChr$ = "$"
If NextChr$ = "d" Then Let NextChr$ = "%"
If NextChr$ = "E" Then Let NextChr$ = "^"
If NextChr$ = "e" Then Let NextChr$ = "&"
If NextChr$ = "f" Then Let NextChr$ = "*"
If NextChr$ = "H" Then Let NextChr$ = "("
If NextChr$ = "I" Then Let NextChr$ = ")"
If NextChr$ = "i" Then Let NextChr$ = "-"
If NextChr$ = "k" Then Let NextChr$ = "_"
If NextChr$ = "L" Then Let NextChr$ = "+"
If NextChr$ = "M" Then Let NextChr$ = "="
If NextChr$ = "m" Then Let NextChr$ = "["
If NextChr$ = "N" Then Let NextChr$ = "]"
If NextChr$ = "n" Then Let NextChr$ = "{"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "o" Then Let NextChr$ = "\"
If NextChr$ = "P" Then Let NextChr$ = "|"
If NextChr$ = "p" Then Let NextChr$ = ";"
If NextChr$ = "r" Then Let NextChr$ = "'"
If NextChr$ = "S" Then Let NextChr$ = ":"
If NextChr$ = "s" Then Let NextChr$ = """"
If NextChr$ = "t" Then Let NextChr$ = ","
If NextChr$ = "U" Then Let NextChr$ = "."
If NextChr$ = "u" Then Let NextChr$ = "/"
If NextChr$ = "V" Then Let NextChr$ = "<"
If NextChr$ = "W" Then Let NextChr$ = ">"
If NextChr$ = "w" Then Let NextChr$ = "?"
If NextChr$ = "X" Then Let NextChr$ = "¥"
If NextChr$ = "x" Then Let NextChr$ = "Ä"
If NextChr$ = "Y" Then Let NextChr$ = "ƒ"
If NextChr$ = "y" Then Let NextChr$ = "Ü"
If NextChr$ = "!" Then Let NextChr$ = "¶"
If NextChr$ = "?" Then Let NextChr$ = "£"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "æ"
If NextChr$ = "1" Then Let NextChr$ = "q"
If NextChr$ = "%" Then Let NextChr$ = "w"
If NextChr$ = "2" Then Let NextChr$ = "e"
If NextChr$ = "3" Then Let NextChr$ = "r"
If NextChr$ = "_" Then Let NextChr$ = "t"
If NextChr$ = "-" Then Let NextChr$ = "y"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$

dustepp2:
If Crap% > 0 Then Let Crap% = Crap% - 1
DoEvents
Loop
r_encrypt = newsent$
End Function

Function Text_Decrypt(strin As String)
'Returns the strin encrypted
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)

If crapp% > 0 Then GoTo dustepp2

If NextChr$ = "~" Then Let NextChr$ = "A"
If NextChr$ = "`" Then Let NextChr$ = "a"
If NextChr$ = "!" Then Let NextChr$ = "B"
If NextChr$ = "@" Then Let NextChr$ = "c"
If NextChr$ = "#" Then Let NextChr$ = "c"
If NextChr$ = "$" Then Let NextChr$ = "D"
If NextChr$ = "%" Then Let NextChr$ = "d"
If NextChr$ = "^" Then Let NextChr$ = "E"
If NextChr$ = "&" Then Let NextChr$ = "e"
If NextChr$ = "*" Then Let NextChr$ = "f"
If NextChr$ = "(" Then Let NextChr$ = "H"
If NextChr$ = ")" Then Let NextChr$ = "I"
If NextChr$ = "-" Then Let NextChr$ = "i"
If NextChr$ = "_" Then Let NextChr$ = "k"
If NextChr$ = "+" Then Let NextChr$ = "L"
If NextChr$ = "=" Then Let NextChr$ = "M"
If NextChr$ = "[" Then Let NextChr$ = "m"
If NextChr$ = "]" Then Let NextChr$ = "N"
If NextChr$ = "{" Then Let NextChr$ = "n"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "\" Then Let NextChr$ = "o"
If NextChr$ = "|" Then Let NextChr$ = "P"
If NextChr$ = ";" Then Let NextChr$ = "p"
If NextChr$ = "'" Then Let NextChr$ = "r"
If NextChr$ = ":" Then Let NextChr$ = "S"
If NextChr$ = """" Then Let NextChr$ = "s"
If NextChr$ = "," Then Let NextChr$ = "t"
If NextChr$ = "." Then Let NextChr$ = "U"
If NextChr$ = "/" Then Let NextChr$ = "u"
If NextChr$ = "<" Then Let NextChr$ = "V"
If NextChr$ = ">" Then Let NextChr$ = "v"
If NextChr$ = "?" Then Let NextChr$ = "w"
If NextChr$ = "¥" Then Let NextChr$ = "x"
If NextChr$ = "Ä" Then Let NextChr$ = "X"
If NextChr$ = "ƒ" Then Let NextChr$ = "Y"
If NextChr$ = "Ü" Then Let NextChr$ = "y"
If NextChr$ = "¶" Then Let NextChr$ = "!"
If NextChr$ = "£" Then Let NextChr$ = "?"
If NextChr$ = "…" Then Let NextChr$ = "."
If NextChr$ = "æ" Then Let NextChr$ = ","
If NextChr$ = "q" Then Let NextChr$ = "1"
If NextChr$ = "w" Then Let NextChr$ = "%"
If NextChr$ = "e" Then Let NextChr$ = "2"
If NextChr$ = "r" Then Let NextChr$ = "3"
If NextChr$ = "t" Then Let NextChr$ = "_"
If NextChr$ = "y" Then Let NextChr$ = "-"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$

dustepp2:
If cra% > 0 Then Let cra% = cra% - 1
DoEvents
Loop
r_decrypt = newsent$
End Function
Function Text_Elite(strin As String)
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
If NextChr$ = "M" Then Let NextChr$ = "(\/)"
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
Text_Elite = newsent$
End Function
Sub DestroyFile(sFileName As String)
'this is the same as deleting a file but it does'nt send to the
'recycling bin, it just deletes it automatically
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    'Create two buffers with a specified 'wipe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wipe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
        Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
        For iLoop = 1 To Blocks
            offset = Seek(hFileHandle)
            Put hFileHandle, , Block1
            Put hFileHandle, offset, Block2
        Next iLoop
    Close hFileHandle
    'Now you can delete the file, which contains no sensitive data
    Kill sFileName
End Sub


Public Sub TransferListToTextBox(Lst As ListBox, Txt As textbox)
'This moves the individual highlighted part of a
'listbox to a textbox
Ind = Lst.ListIndex
daname$ = Lst.List(Ind)
Txt.Text = ""
Txt.Text = daname$
End Sub
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

Sub RemoveDuplicateNames(Lst As Control)
'Removes Duplicate Names from a Listbox
'Ex....: Call RemoveDuplicateNames(List1)
For i = 0 To Lst.ListCount - 1
For nig = 0 To Lst.ListCount - 1

If LCase(Lst.List(i)) Like LCase(Lst.List(nig)) And i <> nig Then

Lst.RemoveItem (nig)
End If

Next nig
Next i
End Sub
Sub RemoveItemFromListbox(Lst As ListBox, item$)
'SelfEx.
Do
nofreeze% = DoEvents()
If LCase$(Lst.List(a)) = LCase$(item$) Then Lst.RemoveItem (a)
a = 1 + a
Loop Until a >= Lst.ListCount
End Sub

Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'Writing to a INI:
'R% = WritePrivateProfileString("Type", "Name", "Value", App.Path + "\KnK.ini")

'Read:
'Name$ = GetFromINI("Type", "Name", App.Path + "\KnK.ini")
'If Name$ = "Value" Then

End Function
Public Function FileExists(sFileName As String) As Boolean
'Example:
'If FileExists ("c:\windows\win.ini") then msgbox "Exist" else Msgbox "Doesn't Exist"
On Error Resume Next
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
    If err.Number <> 0 Then Exit Function
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Sub FileSetHidden(TheFile As String)
'Example: FileSetHidden ("c:\windows\win.ini")
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Sub FileSetNormal(TheFile As String)
'Example: FileSetNormal ("c:\windows\win.ini")
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub
Public Sub FileSetReadOnly(TheFile As String)
'Example: FileSetReadOnly("c:\windows\win.ini")
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
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
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
'Example: Call LoadListBox("c:\windows\codes.txt", List1)
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
Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
'Example:  Call LoadComboBox("c:\windows\codes.txt", Combo1)
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

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Sub File_ReName(file$, newname$)
'Example:
'  Call File_ReName("c:\important.txt", "c:\notimportant.txt")
Name file$ As newname$
nofreeze% = DoEvents()
End Sub
Sub FormFadeBlink(TheForm As Form)
'Example:  FormFadeBlink Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
TheForm.BackColor = &H0&
TheForm.DrawStyle = 6
TheForm.DrawMode = 13

TheForm.DrawWidth = 2
TheForm.ScaleMode = 3
TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
TheForm.Line (0, 0)-(TheForm.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i

End Sub
Sub FormFadeToBlack(TheForm As Form)
'Example:  FormFadeToBlack Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.

TheForm.BackColor = &H0&
TheForm.DrawStyle = 6
TheForm.DrawMode = 13

TheForm.DrawWidth = 2
TheForm.ScaleMode = 3
TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
TheForm.Line (0, 0)-(TheForm.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i
TheForm.BackColor = &H0&
TheForm.DrawStyle = 6
TheForm.DrawMode = 13

TheForm.DrawWidth = 2
TheForm.ScaleMode = 3
TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
TheForm.Line (0, 0)-(TheForm.Width, Y + 2), RGB(i + 3, i, i * 3), BF
Y = Y + 2
Next i

End Sub
Sub FormFadeBW(TheForm)
'Example:  FormFadeBW Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
TheForm.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
TheForm.Line (0, b)-(TheForm.Width, b + 1), RGB(a + 1, a, a * 1), BF
b = b + 2
Next a

End Sub
Sub SaveTextFile(nameoffile, TextToSave)

MyString = TextToSave
filenum = FreeFile
Open (nameoffile) For Output As filenum
Write #filenum, MyString
Close filenum
End Sub
Sub OpenTextFile(nameoffile As String, TextToOpen)
    Dim TextString As String
    On Error Resume Next
    Open nameoffile$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    TextToOpen.Text = TextString$
End Sub
Sub Panel3DIn(Parent As Form, who As Control)
'Example: Call Panel3DIn(Form1, Command1)
'NOTE: Will not work in the Forms Form_Load()
'      If wanted to start this procedure off at the begginin, please
'      do it in the form's Form_Resize()
'Top Dark Gray
Parent.Line (who.Left + who.Width, who.Top - 10)-(who.Left - 30, who.Top - 10), RGB(127, 127, 127)

'Left Dark Gray
Parent.Line (who.Left - 10, who.Top)-(who.Left - 10, who.Top + who.Height), RGB(127, 127, 127)

'Bottom White
Parent.Line -(who.Left + who.Width, who.Top + who.Height), RGB(255, 255, 255)

'Right White
Parent.Line -(who.Left + who.Width, who.Top - 30), RGB(255, 255, 255)

End Sub
Sub Panel3DOff(Parent As Form, who As Control)
'Example: Call Panel3DOff(Form1, Command1)
'NOTE: Will not work in the Forms Form_Load()
'      If wanted to start this procedure off at the begginin, please
'      do it in the form's Form_Resize()
'Top Dark Gray
Parent.Line (who.Left + who.Width, who.Top - 10)-(who.Left - 30, who.Top - 10), RGB(191, 191, 191)

'Left Dark Gray
Parent.Line (who.Left - 10, who.Top)-(who.Left - 10, who.Top + who.Height), RGB(191, 191, 191)

'Bottom White
Parent.Line -(who.Left + who.Width, who.Top + who.Height), RGB(191, 191, 191)

'Right White
Parent.Line -(who.Left + who.Width, who.Top - 30), RGB(191, 191, 191)

End Sub

Sub Panel3DOut(Parent As Form, who As Control)
'Example: Call Panel3DOut(Form1, Command1)
'NOTE: Will not work in the Forms Form_Load()
'      If wanted to start this procedure off at the begginin, please
'      do it in the form's Form_Resize()
'Top White
Parent.Line (who.Left + who.Width, who.Top - 10)-(who.Left - 30, who.Top - 10), RGB(255, 255, 255)

'Left White
Parent.Line (who.Left - 10, who.Top)-(who.Left - 10, who.Top + who.Height), RGB(255, 255, 255)

'Bottom Dark Gray
Parent.Line -(who.Left + who.Width, who.Top + who.Height), RGB(127, 127, 127)

'Right Dark Gray
Parent.Line -(who.Left + who.Width, who.Top - 30), RGB(127, 127, 127)


End Sub

Function ScanFile(FileName As String, SearchString As String) As Long
'ScanFile("C:\FileName.???","String to Search for")
Free = FreeFile
Dim where As Long
Open FileName$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
    Text$ = Space(32000)
    Get #Free, X, Text$
    Debug.Print X
    If InStr(1, Text$, SearchString$, 1) Then
        where = InStr(1, Text$, SearchString$, 1)
        ScanFile = (where + X) - 1
        Close #Free
        Exit For
    End If
    Next X
Close #Free
End Function
Sub TextFromFileToTextBox(Fle As String, Txt As textbox)
'Call TextFromFileToTextBox("C:\????\Filename.txt,Text1)
     Dim FileName As String
     Dim F As Integer

     FileName = Fle

        F = FreeFile                   'Get a file handle
        Open FileName For Input As F   'Open the file
        Txt.Text = Input$(LOF(F), F) 'Read entire file into text box
        Close F                        'Close the file.

End Sub
 Sub PlayCD(TRack$)
'Plays the given track of a cd
'Example: PlayCD (1)
     Dim lRet As Long
     Dim nCurrentTrack As Integer

     'Open the device
     lRet = MciSendString("open cdaudio alias cd wait", 0&, 0, 0)

     'Set the time format to Tracks (default is milliseconds)
     lRet = MciSendString("set cd time format tmsf", 0&, 0, 0)

     'Then to play from the beginning
     lRet = MciSendString("play cd", 0&, 0, 0)

     'Or to play from a specific track, say track 4
     nCurrentTrack = TRack
     lRet = MciSendString("play cd from" & Str(nCurrentTrack), 0&, 0, 0)

     End Sub

     Sub StopCD()
     Dim lRet As Long

     'Stop the playback
     lRet = MciSendString("stop cd wait", 0&, 0, 0)

     DoEvents  'Let Windows process the event

     'Close the device
     lRet = MciSendString("close cd", 0&, 0, 0)

     End Sub
Sub Form3D(frmForm As Form)
'Example:  Form3D Me
'*WARNING:  This procedure will not work in the Form_Load() of a form
'           it must be done in the Form_resize() part of a form if u
'           wish to start the program of with this procedure.
       Const cPi = 3.1415926
       Dim intLineWidth As Integer
       intLineWidth = 5
       '     'save scale mode
       Dim intSaveScaleMode As Integer
       intSaveScaleMode = frmForm.ScaleMode
       frmForm.ScaleMode = 3
       Dim intScaleWidth As Integer
       Dim intScaleHeight As Integer
       intScaleWidth = frmForm.ScaleWidth
       intScaleHeight = frmForm.ScaleHeight
       '     'clear form
       frmForm.Cls
       '     'draw white lines
       frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
       frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
       '     'draw grey lines
       frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
       frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
       '     'draw triangles(actually circles) at corners
       Dim intCircleWidth As Integer
       intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
       frmForm.FillStyle = 0
       frmForm.FillColor = QBColor(15)
       frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
       frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
       '     'draw black frame
       frmForm.Line (0, intScaleHeight)-(0, 0), 0
       frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
       frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
       frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
       '     'restore scale mode
       frmForm.ScaleMode = intSaveScaleMode
End Sub

Sub FadeFormBlue(vForm As Form)
'Example:  FadeFormBlue Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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

Sub FadeFormFire(vForm As Object)
'Example:  FadeFormFire Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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
Sub FadeFormGreen(vForm As Form)
'Example:  FadeFormGreen Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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
'Example:  FadeFormGrey Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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

Sub FadeFormIce(vForm As Object)
'Example:  FadeFormIce Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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

Sub FadeFormPlatinum(vForm As Object)
'Example:  FadeFormPlatinum Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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
Sub FadeFormPurple(vForm As Form)
'Example:  FadeFormPurple Me
'*WARNING: This function cannot be done in Form_Load (). You may use
'          Form_Resize() as a replacement.
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

Sub PLAYWAVE(file)
'Example: PLAYWAVE ("C:\windows\sound.wav")
SoundName$ = file
   ValueFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySoundA(SoundName$, ValueFlags%)

End Sub
' Ask Windows for the list of tasks.
Public Sub ChangeAllCaption(app_name As String, mycaption As String)
'this will change all if not most of the captions
'of the programs you have loaded, kinda fun if
'your messing with someone

    Target = app_name
    Caption = mycaption
    EnumWindows AddressOf EnumCallback6, 0
End Sub

Public Sub ChangeCaption(app_name As String, mycaption As String)
'changes the caption of window

    Target = app_name
    Caption = mycaption
    EnumWindows AddressOf EnumCallback6b, 0
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

Public Sub ShowTask(app_name As String)
'unhides a given window name
'Example: UnhideTask ("Windows Explorer")
    Target = app_name
    EnumWindows AddressOf EnumCallback5, 0
End Sub
Public Sub HideTask(app_name As String)
'Hides a window by the given Caption Text
'Example: HideTask ("Windows Explorer")
    Target = app_name
    EnumWindows AddressOf EnumCallback2, 0
End Sub

Public Sub MinimizeTask(app_name As String)
'Minimizes a window by the given Caption Text
'Example: MinimizeTask ("Windows Explorer")
    Target = app_name
    EnumWindows AddressOf EnumCallback3, 0
End Sub
Public Sub MaximizeTask(app_name As String)
'Maximizes a window by the given Caption Text
'Example: Maximize ("Windows Explorer")
    Target = app_name
    EnumWindows AddressOf EnumCallback4, 0
End Sub
' Check a returned task to see if we should
' kill it.
' this is needed for the terminate task function
' of this basfile
Public Function EnumCallback(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
    If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        SendMessage app_hwnd, WM_CLOSE, 0, 0
    End If
    
    ' Continue searching.
    EnumCallback = 1
End Function
'this is needed for other functions
Public Function EnumCallback3(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
    If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        MiniWindow (app_hwnd)
    End If
    
    ' Continue searching.
    EnumCallback3 = 1
End Function

'this is needed for other functions
Public Function EnumCallback4(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
    If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        MaxWindow (app_hwnd)
    End If
    
    ' Continue searching.
    EnumCallback4 = 1
End Function
'this is needed for other functions
Public Function EnumCallback5(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
    If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        ShowWin (app_hwnd)
    End If
    
    ' Continue searching.
    EnumCallback5 = 1
End Function

Public Function EnumCallback6(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
   
        ' Kill the window.
        Call SetText(app_hwnd, Caption)
 
    
    ' Continue searching.
    EnumCallback6 = 1
End Function
Public Function EnumCallback6b(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
       If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        Call SetText(app_hwnd, Caption)
    End If
 

 
    
    ' Continue searching.
    EnumCallback6b = 1
End Function
'this is needed for other functions
Public Function EnumCallback2(ByVal app_hwnd As Long, ByVal param As Long) As Long
Dim buf As String * 256
Dim Title As String
Dim Length As Long

    ' Get the window's title.
    Length = GetWindowText(app_hwnd, buf, Len(buf))
    Title = Left$(buf, Length)

    ' See if this is the target window.
    If InStr(Title, Target) <> 0 Then
        ' Kill the window.
        HideWindow (app_hwnd)
    End If
    
    ' Continue searching.
    EnumCallback2 = 1
End Function

'--------------







'this centers an object on  your form
'Example: Call CenterObject(Command1, Form1)
Public Sub CenterObject(object As Object, Form As Form)
   With object
      .Left = (Form.Width - .Width) / 2
      .Top = (Form.Height - .Height) / 2
   End With
End Sub
Public Sub TopLeftObject(object As Object)
   With object
      .Left = 0
      .Top = 0
   End With
End Sub
Sub ChangeWallpaper(PictureFilePath)
'changes wallpaper
'Example:   Call ChangeWallpaper("C:\windows\win.bmp")
Dim upd As Integer

    
    upd = SPIF_UPDATEINIFILE

    SystemParametersInfo SPI_SETDESKWALLPAPER, _
        0, PictureFilePath, upd
End Sub

Sub ClickStartMenu()
'clicks the startmenu
Const MENU_KEYCODE = 91

    ' Press the button.
    keybd_event MENU_KEYCODE, 0, 0, 0
    DoEvents

    ' Release the button.
    keybd_event MENU_KEYCODE, 0, KEYEVENTF_KEYUP, 0
    DoEvents
End Sub







Sub OpenCDR()
'Example: OpenCDR
    MciSendString "Set CDAudio Door Open Wait", _
        0&, 0&, 0&
End Sub
Sub RecycleFile(FileName)
'Example: RecycleFile ("c:\windows\desktop\blowme.wav")
Dim OP As SHFILEOPSTRUCT

    With OP
        .wFunc = FO_DELETE
        .pFrom = FileName
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation OP
End Sub

Sub shutdown()

StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)

End Sub


Sub ShowWin(hwnd)
Dim hi
hi = showwindow(hwnd, SW_SHOW)
End Sub



Function HideTaskbar()
Dim Hw1&
Hw1& = FindWindow("Shell_TrayWnd", vbNullString)
HideWindow (Hw1)
End Function
Function ShowTaskbar()
Dim Hw1&
Hw1& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWin (Hw1)
End Function




Function CloseTasks()
'this closes the new systay in win xp
Dim Hw1&, Hw2&
Hw1& = FindWindow("ToolbarWindow32", vbNullString)
Hw2& = FindChildByClass(Hw1&, "SysPager")
HideWindow (Hw2)
End Function

Sub AcidTrip(frm As Form)
'Example:   AcidTrip Me
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
Sub ForceShutdown()
'Example:  ForceShutDown
ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub
Sub RestartComputer()
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Function GetWindowDir()
'finds the window's directory
'Exmaple:  Msgbox GetWindowDir
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function
Sub Pause(interval)
'Example: Pause 3
'Will Pause for 3 seconds
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Sub do3d(Obj As Control, style%, Thick%)
On Error Resume Next
Obj.Parent.AutoRedraw = True
    If Thick <= 0 Then Thick = 1
    If Thick > 8 Then Thick = 8
    OldMode = Obj.Parent.ScaleMode
    OldWidth = Obj.Parent.DrawWidth
    Obj.Parent.ScaleMode = 3
    Obj.Parent.DrawWidth = 1
    ObjHeight = Obj.Height
    ObjWidth = Obj.Width
    ObjLeft = Obj.Left
    ObjTop = Obj.Top
    
    Select Case style
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
            Obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), TLshade
            Obj.Parent.Line -Step(0, CurHigh), BRshade
            Obj.Parent.Line -Step(-CurWide, 0), BRshade
            Obj.Parent.Line -Step(0, -CurHigh), TLshade
        Next i
        If Thick > 2 Then
            CurLeft = ObjLeft - Thick - 1
            CurTop = ObjTop - Thick - 1
            CurWide = ObjWidth + ((Thick + 1) * 2) - 1
            CurHigh = ObjHeight + ((Thick + 1) * 2) - 1
            Obj.Parent.Line (CurLeft, CurTop)-Step(CurWide, 0), QBColor(0)
            Obj.Parent.Line -Step(0, CurHigh), QBColor(0)
            Obj.Parent.Line -Step(-CurWide, 0), QBColor(0)
            Obj.Parent.Line -Step(0, -CurHigh), QBColor(0)
        End If
    Obj.Parent.ScaleMode = OldMode
    Obj.Parent.DrawWidth = OldWidth
End Sub
Function MouseOverHwnd()
'Returns the handle of everything the mouse is over
'Example:  Call GetCaption (MouseOverHwnd)
'that will return the caption of everything the mouse if over
    ' Declares
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      MouseOverHwnd = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function
Sub MiniWindow(hwnd)
'Minimizes a window with the given Handle
'Example:  MiniWindow (MouseOverHwnd)
mi = showwindow(hwnd, SW_MINIMIZE)
End Sub
Sub CloseWindow(winew)
'This will close a window to a giving Handle

closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Sub StayOnTop(the As Form)
'sets your form to be the topmost window all the
'time. Example:  StayOnTop Me
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub NotOnTop(the As Form)
'Example: NotOnTop Form1
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub MoveFormNoCaption(frm As Form)
'Example: MoveFormNoCaption Form1
'Should be used in  an objects Mouse_Down() procedure
frm.MousePointer = 15
ReleaseCapture
g% = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, 2, 0)
Call SnapCheck(frm)
frm.MousePointer = 1
End Sub


Sub MaxWindow(hwnd)
'makes "hWnd" window Maximized
'Example:  MaxWindow (MouseOverHwnd)
ma = showwindow(hwnd, SW_MAXIMIZE)
End Sub

Sub HideWindow(hwnd)
'hides the "hWnd" window
'Example: HideWindow (MouseOverHwnd)
hi = showwindow(hwnd, SW_HIDE)
End Sub

Public Sub CenterForm(frmForm As Form)
' this will center you form in the very center of
'the users screen
'Example:  CenterForm Me
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Public Sub CenterFormTop(frm As Form)
'this will center the form in the top center of
'the user's screen
'Example CenterFormTop Me
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

Public Sub SnapCheck(frm As Form)
    'snaps form to screen
    'sort of like winamp
    If frm.Left < 0& Then
        Do
            DoEvents
            frm.Left = frm.Left + 10
        Loop Until frm.Left >= 0&
        frm.Left = 0&
    End If
    
    If frm.Top < 0& Then
        Do
            DoEvents
            frm.Top = frm.Top + 10
        Loop Until frm.Top >= 0&
        frm.Top = 0&
    End If
    
    If frm.Top + frm.Height > Screen.Height Then
        Do
            DoEvents
            frm.Top = frm.Top - 10
        Loop Until frm.Top <= Screen.Height - frm.Height
        frm.Top = Screen.Height - frm.Height
    End If
    
    If frm.Left + frm.Width > Screen.Width Then
        Do
            DoEvents
            frm.Left = frm.Left - 10
        Loop Until frm.Left <= Screen.Width - frm.Width
        frm.Left = Screen.Width - frm.Width
    End If
    
    If frm.Left - 400 < 0& Then
        Do
            DoEvents
            frm.Left = frm.Left - 10
        Loop Until frm.Left <= 0&
        frm.Left = 0&
    End If
    
    If frm.Top - 400 < 0& Then
        Do
            DoEvents
            frm.Top = frm.Top - 10
        Loop Until frm.Top <= 0&
        frm.Top = 0&
    End If
    
    If (frm.Left + frm.Width) + 400 > Screen.Width Then
        Do
            DoEvents
            frm.Left = frm.Left + 10
        Loop Until frm.Left + frm.Width >= Screen.Width
        frm.Left = Screen.Width - frm.Width
    End If
    
    If (frm.Top + frm.Height) + 400 > Screen.Height Then
        Do
            DoEvents
            frm.Top = frm.Top + 10
        Loop Until frm.Top + frm.Height >= Screen.Height
        frm.Top = Screen.Height - frm.Height
    End If
End Sub

Public Function AOL()
AOL = FindWindow("aol frame25", vbNullString)
End Function

Public Function MDI()
MDI = FindWindowEx(AOL, 0&, "mdiclient", vbNullString)
End Function

Public Function FindRoom() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe As Long, mdiclient As Long, aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long
FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
       Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid6 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
       Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) And (Winkid7 <> 0) And (Winkid8 <> 0) And (Winkid9 <> 0) Then
              find_aolchild = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
find_aolchild = 0
' example on how to use:
' Dim TheWin As Long
' TheWin = find_aolchild()
' If TheWin <> 0 Then
' What to do if window is there
' End If
End Function
Public Function FindIM() As Long
' If this function finds the window, it will return it's
' handle. If it doesn't find it, it will return 0.
Dim aolframe As Long, mdiclient As Long, aolchild As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, FindOtherWin As Long
FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "_aol_edit", vbNullString)
       Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid4 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) Then
              find_aolchild = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
find_aolchild = 0
' example on how to use:
' Dim TheWin As Long
' TheWin = find_aolchild()
' If TheWin <> 0 Then
' What to do if window is there
' End If
End Function

Public Sub ChatSend(Chat As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub FadeBy2(picbox As Object, firstcolor As Long, secondcolor As Long)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, J&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim firstcolorRed%, firstcolorGreen%, firstcolorBlue%
Dim secondcolorRed%, secondcolorGreen%, secondcolorBlue%
Dim ColorDifRed, ColorDifGreen, ColorDifBlue
SaveScale = picbox.ScaleMode: SaveStyle = picbox.DrawStyle
SaveRedraw = picbox.AutoRedraw: picbox.ScaleMode = 3
firstcolorRed = firstcolor And 255
  firstcolorGreen = (firstcolor And 65280) / 256
    firstcolorBlue = (firstcolor And 16711680) / 65536
secondcolorRed = secondcolor And 255
  secondcolorGreen = (secondcolor And 65280) / 256
    secondcolorBlue = (secondcolor And 16711680) / 65536
      aRed = firstcolorRed
      aGreen = firstcolorGreen
      aBlue = firstcolorBlue
      pixels = picbox.ScaleWidth
    If pixels <= 0 Then Exit Sub
        ColorDifRed = (secondcolorRed - firstcolorRed)
        ColorDifGreen = (secondcolorGreen - firstcolorGreen)
        ColorDifBlue = (secondcolorBlue - firstcolorBlue)
          RedDelta = ColorDifRed / pixels
          GreenDelta = ColorDifGreen / pixels
          BlueDelta = ColorDifBlue / pixels
        picbox.DrawStyle = 5
        picbox.AutoRedraw = True
For Y = 0 To pixels + 1
        aRed = aRed + RedDelta
            If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
            If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
            If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
            If ThisColor > -1 Then
                picbox.Line (Y - 2, -2)-(Y - 2, picbox.Height + 2), ThisColor, BF
            End If
    Next Y
picbox.ScaleMode = SaveScale
picbox.DrawStyle = SaveStyle
picbox.AutoRedraw = SaveRedraw
End Sub

