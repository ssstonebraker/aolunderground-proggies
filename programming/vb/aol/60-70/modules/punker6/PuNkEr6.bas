Attribute VB_Name = "PuNkEr6"
Attribute VB_Description = "Written by PuNkDuDe"
'PuNkEr6.bas : VB6
'By PuNkDuDe
'PuNkDuDe OnLiNe - http://punkdude.cjb.net

'AOL 6.0 (AOL 5.0/4.0  procedures are commented)
'AIM 4.3

'Over 200 procedures

' º•º~› GREETZ
' •¤•~› Progee
' ¤•¤~› Dev
' •¤•~› tru
' ¤•¤~› k2
' •¤•~› MaestrO
' ¤•¤~› wn
' •¤•~› Phear
' ¤•¤~› rudd
' •¤•~› TeXx
' ¤•¤~› duck
' •¤•~› bone
' ¤•¤~› Circle Of Doom
' •¤•~› Rit Man
' ¤•¤~› Sl4sH3d
' •¤•~› ChRoMe
' ¤•¤~› Beav

Option Explicit 'Requires that all variables are defined

'Variable
Dim StopBust As Boolean 'Room Bust Stop
Dim nid As NOTIFYICONDATA  'Tray icon

'Declaration
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function setactivewindow Lib "user32" Alias "SetActiveWindow" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Constant
'WM=window message, LB=listbox, MF=menu, CB=combobox, VK=virtual key, BM=checkbox
'SND=sndPlaySound, SWP=SetWindowPos, SW=ShowWindow, SPI=SystemParametersInfo
'GWL=GetWindowLong

Public Const PW_CHAR = "*" 'CUSTOM CONST: SetPWChar; by PuNkDuDe

Public Const WM_ACTIVATE = &H6
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_SETFONT = &H30
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SYSCOMMAND = &H112

Public Const LB_ADDSTRING = &H180
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETCURSEL = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_RESETCONTENT = &H184
Public Const LB_SETCURSEL = &H186
Public Const LB_GETTOPINDEX = &H18E

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

Public Const CB_ADDSTRING = &H143
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_RESETCONTENT = &H14B
Public Const CB_SETCURSEL = &H14E

Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SNAPSHOT = &H2C
Public Const VK_SPACE = &H20

Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_SYNC = &H0

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_GETPASSWORDCHAR = &HD2

Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MINIMIZE = &H20000000

Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1

Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPI_SETDRAGFULLWINDOWS = 37

Public Const GWL_STYLE = (-16)

Public Const MAX_PATH = 260

Public Const SRCCOPY = &HCC0020

Public Const MF_BYPOSITION = &H400&

Public Const KEYEVENTF_KEYUP = &H2

Public Const SHERB_NOCONFIRMATION = &H1
Public Const SHERB_NOPROGRESSUI = &H2
Public Const SHERB_NOSOUND = &H4
'Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type LIST_INFO   'CUSTOM TYPE: GetListInfo; by PuNkDuDe
    Count As Long 'List count
    Cursel As Long 'Selected index
    ItemData As Long 'Item data
    ItemHeight As Long 'Item height
    ItemRect As RECT 'Item RECT
    SelCount As Long 'Select count
    SelText As String 'Selected text
    TextLen As Long 'Selected text lenght
    TopIndex As Long 'Index of the topmost visible index
End Type

Public Type AOL_INFO   'CUSTOM TYPE: GetAOLInfo; by PuNkDuDe
    sCaption As String 'AOL caption
    sDir As String 'AOL directory
    hwnd As Long 'AOL Frame25 hwnd
    hDC As Long 'DC of AOLs main window
    hRoomHwnd As Long 'Chatroom hwnd
    hMDI As Long 'MDIClient hwnd
    iLaunches As Integer 'Number of times AOL has been launched since installation
    iSignOns As Integer 'Number of times AOL has been signed on since installation
    sUserSN As String 'User SN
    bOnline As Boolean 'Wether or not the user is signed on
    iVersion As Integer 'AOL version
End Type

Public Type AIM_INFO   'CUSTOM TYPE: GetAIMInfo; by PuNkDuDe
    lCount As Long 'Buddy list count
    sCaption As String 'AIM caption
    sUserName As String 'User SN
    hwnd As Long 'AIM hwnd
    hDC As Long 'DC of AIMs main window
End Type

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    UID As Long
    uFlags As Long
    uCallBackmessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Enum AIM_AD 'CUSTOM ENUM: AIM_GetAd; by PuNkDuDe
    topAd
    BottomAd
End Enum

Public Enum CHAT_TYPE 'CUSTOM ENUM: EnterRoom; by PuNkDuDe
    cPrivate
    cPublic
    cMember
    cRestricted
End Enum

Public Enum CHAT_PREF 'CUSTOM ENUM: SetChatPref; by PuNkDuDe
    Notify_Arrive = 1
    Notify_Leave = 2
    Double_Space = 3
    Alpha_List = 4
    Room_Sounds = 5
End Enum

Public Enum ImOn_Off 'CUSTOM ENUM: ImOnOff; by PuNkDuDe
    imON
    imOff
End Enum

Public Enum WINDOW_STATE 'CUSTOM ENUM: GetWindowState; by PuNkDuDe
    Maximized
    Minimized
    Normal
End Enum

Public Enum ASCII_TYPE 'CUSTOM ENUM: GenerateAscii; by PuNkDuDe
    aLeft
    aRight
    aOther
End Enum

Public Enum MAIL_TAB 'CUSTOM ENUM: MailGotoTab; by PuNkDuDe
    mNew = 0
    mOld = 1
    mSent = 2
End Enum

Public Enum TERM_LANG 'CUSTOM ENUM: TranslateTerm; by PuNkDuDe
    lGerman
    lSpanish
    lJapanese
End Enum

Public Enum TRAY_STATE 'CUSTOM ENUM: CD_Tray; by PuNkDuDe
    Tray_Open
    Tray_Close
End Enum

Public Enum CD_PAUSERESUME 'CUSTOM ENUM: CD_Pause; by PuNkDuDe
    cdPause
    cdResume
End Enum

Public Enum CENTER_TYPE 'CUSTOM ENUM: Form_CenterControl; by PuNkDuDe
    oVertical
    oHorizontal
    oBoth
End Enum

Public Enum EMPTY_RECYCLE 'CUSTOM ENUM: EmptyRecycleBin; by PuNkDuDe
    NoPreference = 0
    NoConfirm = SHERB_NOCONFIRMATION
    NoProgress = SHERB_NOPROGRESSUI
    NoSound = SHERB_NOSOUND
End Enum
'Procedures : PuNkDuDe
Public Function GetText(Win As Long) As String
    'Get window text
    Dim Length As Long
    Dim sLength As String
    Length = SendMessage(Win&, WM_GETTEXTLENGTH, 0&, 0&)
    sLength = String$(Length, 0&)
    Call SendMessageByString(Win&, WM_GETTEXT, Length& + 1, sLength$)
    GetText$ = sLength$
End Function

Public Function GetListText(lnglist As Long, index As Integer) As String
    'Return list text
    On Error Resume Next
    Dim Ta&, Ta2&
    Dim m_AOLThreadID As Long, m_AOLProcessID As Long
    Dim hAOLProcess As Long, lAddrOfItemData As Long, lAddrOfName As Long, lBytesRead As Long
    Dim sBuffer As String, ScreenName As String
    If lnglist <> 0 Then
        m_AOLThreadID = GetWindowThreadProcessId(lnglist, m_AOLProcessID)
        hAOLProcess = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, m_AOLProcessID)
        If hAOLProcess Then
            sBuffer$ = String$(4, vbNullChar)
            lAddrOfItemData = SendMessage(lnglist, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
            lAddrOfItemData = lAddrOfItemData + 24
            Call ReadProcessMemory(hAOLProcess, lAddrOfItemData, sBuffer$, 4, lBytesRead&)
            Call CopyMemory(lAddrOfName, ByVal sBuffer$, Len(sBuffer$))
            lAddrOfName = lAddrOfName + 6
            sBuffer$ = String$(16, vbNullChar)
            Call ReadProcessMemory(hAOLProcess, lAddrOfName, sBuffer$, Len(sBuffer$), lBytesRead&)
            GetListText$ = Left$(sBuffer$, InStr(sBuffer$, vbNullChar) - 1)
            Call CloseHandle(hAOLProcess)
        End If
    End If
End Function

Public Function GetComboText(lnglist As Long, index As Integer) As String
    'Return combo text
    On Error Resume Next
    Dim m_AOLThreadID As Long, m_AOLProcessID As Long
    Dim hAOLProcess As Long, lAddrOfItemData As Long, lAddrOfName As Long, lBytesRead As Long
    Dim sBuffer As String
    If lnglist <> 0 Then
        m_AOLThreadID = GetWindowThreadProcessId(lnglist, m_AOLProcessID)
        hAOLProcess = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, m_AOLProcessID)
        If hAOLProcess Then
            sBuffer$ = String$(4, vbNullChar)
            lAddrOfItemData = SendMessage(lnglist, CB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
            lAddrOfItemData = lAddrOfItemData + 24
            Call ReadProcessMemory(hAOLProcess, lAddrOfItemData, sBuffer$, 4, lBytesRead&)
            Call CopyMemory(lAddrOfName, ByVal sBuffer$, Len(sBuffer$))
            lAddrOfName = lAddrOfName + 6
            sBuffer$ = String$(16, vbNullChar)
            Call ReadProcessMemory(hAOLProcess, lAddrOfName, sBuffer$, Len(sBuffer$), lBytesRead&)
            GetComboText$ = Left$(sBuffer$, (InStr(sBuffer$, vbNullChar) - 1))
            Call CloseHandle(hAOLProcess)
        End If
    End If
End Function

Public Function FindRoom() As Long
    'Returns the handle of an AOL Chatroom window (if found)
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, ChildRICH As Long, ChildRICH2 As Long
    Dim ChildList As Long
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    ChildRICH& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    ChildRICH2& = FindWindowEx(AOLChild&, ChildRICH&, "RICHCNTL", vbNullString)
    ChildList& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    If (AOLChild& <> 0&) And (ChildRICH& <> 0&) And (ChildRICH2& <> 0&) And (ChildList& <> 0&) Then
        FindRoom& = AOLChild&
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            ChildRICH& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
            ChildRICH2& = FindWindowEx(AOLChild&, ChildRICH&, "RICHCNTL", vbNullString)
            ChildList& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
            If (AOLChild& <> 0&) And (ChildRICH& <> 0&) And (ChildRICH2& <> 0&) And (ChildList& <> 0&) Then
                FindRoom& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindRoom& = AOLChild&
End Function

Public Function FindUploadWindow() As Long
    'Returns the handle of an Upload window (if found)
    Dim AOLModal As Long, AolStatic As Long
    Dim Caption As String
    AOLModal& = FindWindow("_AOL_Modal", vbNullString)
    AolStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    If (GetCaption(AOLModal&) = "File Transfer") And (InStr(GetText(AolStatic&), "Uploading") > 0&) Then
        FindUploadWindow& = AOLModal&
        Exit Function
    Else
        Do
            AOLModal& = FindWindow("_AOL_Modal", vbNullString)
            AolStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
            If (GetCaption(AOLModal&) = "File Transfer") And (InStr(GetText(AolStatic&), "Uploading") > 0&) Then
                FindUploadWindow& = AOLModal&
                Exit Function
            End If
        Loop Until AOLModal& = 0&
    End If
    FindUploadWindow& = AOLModal&
End Function

Public Function FindSentIM(SN As String) As Long
    'Returns the handle of a sent IM window (if found)
    Dim AOLIM As Long
    Dim Caption As String, nCaption As String, nSN As String
    AOLIM& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(AOLIM&)
    nCaption$ = RemoveSpaces(LCase$(Caption$))
    nSN$ = RemoveSpaces(LCase$(SN$))
    If ((Left$(Caption$, 7) = " IM To:") Or (Left$(Caption$, 9) = ">IM From:")) And (InStr(nCaption$, nSN$) > 0&) Then
        FindSentIM& = AOLIM&
        Exit Function
    Else
        Do
            AOLIM& = FindWindowEx(AOLMDI&, AOLIM&, "AOL Child", vbNullString)
            Caption$ = GetCaption(AOLIM&)
            nCaption$ = RemoveSpaces(LCase$(Caption$))
            If ((Left$(Caption$, 7) = " IM To:") Or (Left$(Caption$, 9) = ">IM From:")) And (InStr(nCaption$, nSN$) > 0&) Then
                FindSentIM& = AOLIM&
                Exit Function
            End If
        Loop Until AOLIM& = 0&
    End If
    FindSentIM& = AOLIM&
End Function

Public Function FindBuddySetup() As Long
    'Returns the handle of an Buddy List Setup window (if found)
    Dim AOLChild As Long
    Dim Caption As String
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(AOLChild&)
    If Caption$ = "Buddy List Setup" Then
        FindBuddySetup& = AOLChild&
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            Caption$ = GetCaption(AOLChild&)
            If Caption$ = "Buddy List Setup" Then
                FindBuddySetup& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindBuddySetup& = AOLChild&
End Function

Public Function FindBuddyPref() As Long
    'Returns the handle of the Buddy List Preferences window (if found)
    Dim AOLChild As Long, AOLTab As Long
    Dim Caption As String
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    AOLTab& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
    Caption$ = GetCaption(AOLChild&)
    If (Caption$ = "Buddy List Preferences") And (AOLTab& <> 0&) Then
        FindBuddyPref& = AOLChild&
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            AOLTab& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
            Caption$ = GetCaption(AOLChild&)
            If (Caption$ = "Buddy List Preferences") And (AOLTab& <> 0&) Then
                FindBuddyPref& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindBuddyPref& = AOLChild&
End Function

Public Function FindIMSend() As Long
    'Returns the handle of an IM window (if found)
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, ChildEdit As Long, ChildRICH As Long
    Dim ChildList As Long, ChildIcon1 As Long, ChildIcon2 As Long
    Dim ChildCaption As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    ChildEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    ChildRICH& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    ChildList& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    ChildIcon1& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    ChildIcon2& = FindWindowEx(AOLChild&, ChildIcon1&, "_AOL_Icon", vbNullString)
    ChildCaption$ = GetText(AOLChild&)
    If (AOLChild& <> 0&) And (ChildEdit& <> 0&) And (ChildRICH& <> 0&) And (ChildList& <> 0&) And (ChildIcon1& <> 0&) And (ChildIcon2& <> 0&) And (ChildCaption$ = "Send Instant Message") Then
        FindIMSend& = AOLChild&
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            ChildEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
            ChildRICH& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
            ChildList& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
            ChildIcon1& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
            ChildIcon2& = FindWindowEx(AOLChild&, ChildIcon1&, "_AOL_Icon", vbNullString)
            ChildCaption$ = GetText(AOLChild&)
            If (AOLChild& <> 0&) And (ChildEdit& <> 0&) And (ChildRICH& <> 0&) And (ChildList& <> 0&) And (ChildIcon1& <> 0&) And (ChildIcon2& <> 0&) And (ChildCaption$ = "Send Instant Message") Then
                FindIMSend& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindIMSend& = AOLChild&
End Function

Public Function FindMailBox() As Long
    'Returns the handle of an AOL Mail Box window (if found)
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long
    Dim TabControl As Long, TabPage As Long, ChildEdit As Long
    Dim ChildCaption As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, ChildEdit&, "_AOL_TabPage", vbNullString)
    ChildCaption$ = GetText(AOLChild&)
    If (AOLChild& <> 0&) And (TabControl& <> 0&) And (TabPage& <> 0&) And (InStr(ChildCaption$, "'s Online Mailbox")) Then
        FindMailBox& = AOLChild&
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, ChildEdit&, "_AOL_TabPage", vbNullString)
            ChildCaption$ = GetText(AOLChild&)
            If (AOLChild& <> 0&) And (TabControl& <> 0&) And (TabPage& <> 0&) And (InStr(ChildCaption$, "'s Online Mailbox")) Then
                FindMailBox& = AOLChild&
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    FindMailBox& = AOLChild&
End Function

Public Function ChatSend(Message As String) As Boolean
    'AOL 4, 5, 6
    Dim ChildRICH As Long, ChildRICH2 As Long, ChildIcon As Long
    Dim OldWin As Long
    Dim OldText As String
    If FindRoom& = 0& Then ChatSend = False: Exit Function
    ChildRICH& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    ChildRICH2& = FindWindowEx(FindRoom&, ChildRICH&, "RICHCNTL", vbNullString)
    ChildIcon& = FindWindowEx(FindRoom&, ChildRICH2&, "_AOL_Icon", vbNullString)
    OldWin& = GetActiveWindow&
    OldText$ = GetText(ChildRICH2&): DoEvents
    Call SendMessageByString(ChildRICH2&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(ChildRICH2&, WM_SETTEXT, 0&, ("" & Message$ & ""))
    Do: DoEvents
        Call SendMessage(ChildIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(ChildIcon&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessage(ChildIcon&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(ChildIcon&, WM_KEYUP, VK_SPACE, 0&)
    Loop Until GetText(ChildRICH2&) = ""
    If Len(OldText$) > 0& Then Call SetText(ChildRICH2&, OldText$)
    If OldWin& <> 0& Then setactivewindow (OldWin&)
    ChatSend = CBool(FindRoom&)
End Function

Public Sub IMSend(Screen_Name As String, Message As String)
    'AOL 4, 5, 6
    Dim X As Long
    X& = FindSentIM&(Screen_Name$)
    If X& <> 0& Then
        Dim RICH As Long, Ico As Long
        RICH& = FindWindowEx(X&, 0&, "RICHCNTL", vbNullString)
        RICH& = FindWindowEx(X&, RICH&, "RICHCNTL", vbNullString)
        Call SendMessageByString(RICH&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(RICH&, WM_SETTEXT, 0&, "" & Message$ & "")
        Ico& = FindWindowEx(X&, 0&, "_AOL_Icon", vbNullString)
        Call SendMessage(Ico&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessage(Ico&, WM_CHAR, VK_RETURN, 0&)
        Exit Sub
    End If
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLCombo As Long, AOLEdit As Long
    Dim AOLIMWindow As Long, IMto As Long, IMmessage As Long, IMIcon As Long
    Dim i As Integer
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLCombo& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit& = FindWindowEx(AOLCombo&, 0&, "Edit", vbNullString)
    If AOLEdit& = 0& Then Exit Sub
    If GetText(AOLEdit&) <> "im" Then
        Do Until (GetText(AOLEdit&) = "im"): DoEvents
            Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
            Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "im")
        Loop
    End If
    DoEvents
    Call SendMessage(AOLEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessage(AOLEdit&, WM_CHAR, VK_RETURN, 0&)
    DoEvents
    Do: DoEvents
        AOLIMWindow& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", "Send Instant Message")
        IMto& = FindWindowEx(AOLIMWindow&, 0&, "_AOL_Edit", vbNullString)
        IMmessage& = FindWindowEx(AOLIMWindow&, 0&, "RICHCNTL", vbNullString)
    Loop Until ((AOLIMWindow& <> 0&) And (IMto& <> 0&) And (IMmessage& <> 0&))
    Call SendMessageByString(IMto&, WM_SETTEXT, 0&, "" & Screen_Name$ & "")
    Call SendMessageByString(IMmessage&, WM_SETTEXT, 0&, "" & Message$ & "")
    IMIcon& = FindWindowEx(AOLIMWindow&, 0&, "_AOL_Icon", vbNullString)
    If AOLVersion2 = 6 Then
        For i = 1 To 9
          IMIcon& = FindWindowEx(AOLIMWindow&, IMIcon&, "_AOL_Icon", vbNullString)
        Next i
    Else
        For i = 1 To 8
          IMIcon& = FindWindowEx(AOLIMWindow&, IMIcon&, "_AOL_Icon", vbNullString)
        Next i
    End If
    DoEvents
    Call ClickIcon(IMIcon&)
End Sub

Public Sub MailSend(Screen_Name As String, subject As String, Message As String)
    'AOL 4, 5, 6
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLTIcon As Long
    Dim MailWindow As Long, MailTo As Long, MailSubject As Long, MailMessage As Long, MailIcon As Long
    Dim i As Integer
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLTIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
    AOLTIcon& = FindWindowEx(AOLToolbar2&, AOLTIcon&, "_AOL_Icon", vbNullString)
    If AOLVersion2 = 6 Then AOLTIcon& = FindWindowEx(AOLToolbar2&, AOLTIcon&, "_AOL_Icon", vbNullString)
    Call SendMessage(AOLTIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(AOLTIcon&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(AOLTIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(AOLTIcon&, WM_KEYUP, VK_SPACE, 0&)
    DoEvents
    Do: DoEvents
        MailWindow& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", "Write Mail")
        MailTo& = FindWindowEx(MailWindow&, 0&, "_AOL_Edit", vbNullString)
        MailSubject& = FindWindowEx(MailWindow&, MailTo&, "_AOL_Edit", vbNullString)
        MailSubject& = FindWindowEx(MailWindow&, MailSubject&, "_AOL_Edit", vbNullString)
        MailMessage& = FindWindowEx(MailWindow&, 0&, "RICHCNTL", vbNullString)
    Loop Until ((MailWindow& <> 0&) And (MailTo& <> 0&) And (MailSubject& <> 0&) And (MailMessage& <> 0&))
    Call SendMessageByString(MailTo&, WM_SETTEXT, 0&, "" & Screen_Name$ & "")
    Call SendMessageByString(MailSubject&, WM_SETTEXT, 0&, "" & subject$ & "")
    Call SendMessageByString(MailMessage&, WM_SETTEXT, 0&, "" & Message$ & "")
    MailIcon& = FindWindowEx(MailWindow&, 0&, "_AOL_Icon", vbNullString)
    If AOLVersion2 = 6 Then
        For i = 1 To 17
            MailIcon& = FindWindowEx(MailWindow&, MailIcon&, "_AOL_Icon", vbNullString)
        Next i
    ElseIf AOLVersion2 = 5 Then
        For i = 1 To 15
            MailIcon& = FindWindowEx(MailWindow&, MailIcon&, "_AOL_Icon", vbNullString)
        Next i
    ElseIf AOLVersion2 = 4 Then
        For i = 1 To 13
            MailIcon& = FindWindowEx(MailWindow&, MailIcon&, "_AOL_Icon", vbNullString)
        Next i
    End If
    DoEvents
    Call SendMessage(MailIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(MailIcon&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(MailIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(MailIcon&, WM_KEYUP, VK_SPACE, 0&)
    WaitForOKorModal
End Sub

Public Sub ListSave(ListB As ListBox, FileN As String)
    'Save text from (ListB) into a text file (FileN)
    Dim X As Long
    Dim i As Integer
    Dim Txt As String
    On Error Resume Next
    X = FreeFile
    For i = 0 To (ListB.ListCount - 1)
        Txt$ = Txt$ & ListB.List(i)
    Next i
    Open FileN$ For Output As #X
        Print #X, Txt$
    Close X
End Sub

Public Sub ListLoad(ListB As ListBox, FileN As String)
    'Load text from (FileN) into a listbox (ListB)
    Dim X As Long
    Dim Txt As String
    On Error Resume Next
    X = FreeFile
    Open FileN$ For Input As #X
        While Not EOF(X)
            Input #X, Txt$
            ListB.AddItem (Txt$)
            DoEvents
        Wend
    Close X
End Sub

Public Sub Text_Save(TextB As TextBox, FileN As String)
    'Save text (TextB) into a text file (FileN)
    Dim X As Long
    Dim Txt As String
    On Error Resume Next
    X = FreeFile
    Open FileN$ For Output As #X
        Print #X, TextB.Text
    Close X
End Sub

Public Sub Text_Load(TextB As TextBox, FileN As String)
    'Load text from (FileN) to a textbox (TextB)
    Dim X As Long
    Dim Txt As String
    On Error Resume Next
    X = FreeFile
    Open FileN$ For Input As #X
        While Not EOF(X): DoEvents
            Input #X, Txt$
            TextB.Text = Txt$
        Wend
    Close X
End Sub

Public Function ShortFileName(LongFile As String) As String
    'Returns the DOS/short file name
    Dim ret As Long
    Dim buffer As String
    buffer$ = String$(256, 0)
    ret = GetShortPathName(LongFile$, buffer$, Len(buffer$))
    If Len(ret) > 0& Then
        ShortFileName$ = Left$(buffer$, ret)
    Else
        ShortFileName$ = ""
    End If
End Function
Public Function FullFileName(ShortFile As String) As String
    'Returns the long file name from a DOS/short file name
    Dim ret As Long
    Dim buffer As String
    buffer$ = Space$(255)
    ret = GetFullPathName(ShortFile$, 255, buffer$, "")
    If Len(ret) > 0& Then
        FullFileName$ = Left$(buffer$, Len(ret))
    Else
        FullFileName$ = ""
    End If
End Function

Public Function UserSN() As String
    'Get the user's screen name from the Welcome window
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLChild As Long
    Dim Caption As String
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", vbNullString)
    Caption$ = GetText(AOLChild&)
    If (Mid$(Caption$, 1, 9) = "Welcome, ") And (Right$(Caption$, 1) = "!") Then
        UserSN$ = Trim$(Mid$(Caption$, 10, InStr(Caption$, "!") - 10))
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDiHandle&, AOLChild&, "AOL Child", vbNullString)
            Caption$ = GetText(AOLChild&)
            If (Mid$(Caption$, 1, 9) = "Welcome, ") And (Right$(Caption$, 1) = "!") Then
                UserSN$ = Trim$(Mid$(Caption$, 10, InStr(Caption$, "!") - 10))
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    UserSN$ = ""
End Function

Public Sub RoomBust(name As String, Optional Pause)
    'Call the RoomBustStop sub to end the RoomBust
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLCombo As Long, AOLEdit As Long
    Dim Count As Integer
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLCombo& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit& = FindWindowEx(AOLCombo&, 0&, "Edit", vbNullString)
    If AOLEdit& = 0& Then Exit Sub
    If CBool(IsMissing(Pause)) = True Then Pause = 0.8
    Call SendMessage(FindRoom&, WM_CLOSE, 0&, 0&)
    StopBust = False
    Count = 0
    Do: DoEvents
        Count = Count + 1
        If GetText(AOLEdit&) <> ("aol://2719:2-2-" & name$) Then
            Do: DoEvents
                Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
                Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ("aol://2719:2-2-" & name$))
            Loop Until (GetText(AOLEdit&) = ("aol://2719:2-2-" & name$))
        End If
        Call SendMessage(AOLEdit&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessage(AOLEdit&, WM_CHAR, VK_RETURN, 0&)
        If FindWindow("#32770", "America Online") Then
            Do: DoEvents
                If FindWindow("#32770", "America Online") <> 0& Then
                    Call SendMessage(FindWindow("#32770", "America Online"), WM_CLOSE, 0&, 0&)
                End If
            Loop Until (FindWindow("#32770", "America Online") = 0&)
        End If
        If Count <> 4 Then TimeOut (Pause) Else TimeOut (1): Count = 0
    Loop Until (FindRoom& <> 0&) Or (StopBust = True)
End Sub

Public Sub RoomBustStop()
    'Stops the RoomBust sub
    StopBust = True
End Sub

Public Function ChatIgnoreByIndex(index As Integer, Ignore As Boolean, Optional ReturnSN As String) As Boolean
    'AOL 5.0 & 4.0
    Dim AOLHandle As Long, AOLMDiHandle As Long, ChatList As Long, tIndex As Long
    Dim ChatigWin As Long, ChatigCheck1 As Long, ChatigCheck2 As Long, ChatigCheck3 As Long, ChatigCheck As Long, ChatigChecked As Long
    Dim SN As String, FullSn As String
    Dim i As Integer
    If (FindRoom& = 0&) Or (AOLVersion2 = 6) Then Exit Function
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDiClient", vbNullString)
    ChatList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    If index > (SendMessage(ChatList&, LB_GETCOUNT, 0&, 0) - 1) Then Exit Function
    Call SendMessage(ChatList&, LB_SETCURSEL, CLng(index), 0&)
    DoEvents
    SN$ = GetListText(ChatList&, index)
    If LCase$(SN$) <> LCase$(UserSN$) Then
        FullSn$ = SN$
        If CBool(IsMissing(ReturnSN)) = False Then ReturnSN$ = FullSn$
    End If
    DoEvents
    Call SendMessageByNum(ChatList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do: DoEvents
        ChatigWin& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", FullSn$)
        ChatigCheck& = FindWindowEx(ChatigWin&, 0&, "_AOL_Checkbox", vbNullString)
        ChatigCheck1& = FindWindowEx(ChatigWin&, 0&, "_AOL_icon", vbNullString)
        ChatigCheck2& = FindWindowEx(ChatigWin&, ChatigCheck1&, "_AOL_icon", vbNullString)
        ChatigCheck3& = FindWindowEx(ChatigWin&, ChatigCheck2&, "_AOL_icon", vbNullString)
    Loop Until ((ChatigWin& <> 0&) And (ChatigCheck& <> 0&) And (ChatigCheck1& <> 0&) And (ChatigCheck2& <> 0&) And (ChatigCheck3& <> 0&))
    ChatigChecked& = SendMessage(ChatigCheck&, BM_GETCHECK, 0&, 0&)
    ChatIgnoreByIndex = False
    If Not (ChatigChecked& = Ignore) Then
        Call SendMessage(ChatigCheck&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(ChatigCheck&, WM_LBUTTONUP, 0&, 0&): ChatIgnoreByIndex = True
    End If
    Do: DoEvents
        Call SendMessage(ChatigWin&, WM_CLOSE, 0&, 0&)
        ChatigWin& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", FullSn$)
    Loop Until (ChatigWin& = 0&)
End Function

Public Function ChatIgnoreByName(Screen_Name As String, Ignore As Boolean, Optional ReturnSN As String) As Boolean
    'AOL 5.0 & 4.0
    Dim AOLHandle As Long, AOLMDiHandle As Long, ChatList As Long, tIndex As Long
    Dim ChatigWin As Long, ChatigCheck1 As Long, ChatigCheck2 As Long, ChatigCheck3 As Long, ChatigCheck As Long, ChatigChecked As Long
    Dim SN As String, FullSn As String
    Dim i As Integer
    If FindRoom& = 0& Then Exit Function
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    ChatList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(ChatList&, LB_SETCURSEL, 0, 0)
    For i = 0 To (SendMessage(ChatList&, LB_GETCOUNT, 0&, 0) - 1)
        If i > (SendMessage(ChatList&, LB_GETCOUNT, 0&, 0) - 1) Then Exit Function
        Call SendMessage(ChatList&, LB_SETCURSEL, i, 0)
        DoEvents
        SN$ = GetListText(ChatList&, i)
        If (LCase$(SN$) <> LCase$(UserSN$)) And (InStr(ShortString$(SN$), ShortString$(Screen_Name$)) > 0&) Then
            FullSn$ = SN$
            If CBool(IsMissing(ReturnSN$)) = False Then ReturnSN$ = FullSn$
            Exit For
        End If
        If i = (SendMessage(ChatList&, LB_GETCOUNT, 0&, 0) - 1) Then Exit Function
    Next i
    If i = -1 Then Exit Function
    Call SendMessage(ChatList&, LB_SETCURSEL, i, 0)
    DoEvents
    Call SendMessageByNum(ChatList&, WM_LBUTTONDBLCLK, 0, 0&)
    Do: DoEvents
        ChatigWin& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", FullSn$)
        ChatigCheck& = FindWindowEx(ChatigWin&, 0&, "_AOL_Checkbox", vbNullString)
        ChatigCheck1& = FindWindowEx(ChatigWin&, 0&, "_AOL_Icon", vbNullString)
        ChatigCheck2& = FindWindowEx(ChatigWin&, ChatigCheck1&, "_AOL_Icon", vbNullString)
        ChatigCheck3& = FindWindowEx(ChatigWin&, ChatigCheck2&, "_AOL_Icon", vbNullString)
    Loop Until ((ChatigWin& <> 0&) And (ChatigCheck& <> 0&) And (ChatigCheck1& <> 0&) And (ChatigCheck2& <> 0&) And (ChatigCheck3& <> 0&))
    ChatigChecked& = SendMessage(ChatigCheck&, BM_GETCHECK, 0&, 0&)
    ChatIgnoreByName = False
    If Not (ChatigChecked& = Ignore) Then
        Call SendMessage(ChatigCheck&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(ChatigCheck&, WM_LBUTTONUP, 0&, 0&): ChatIgnoreByName = True
    End If
    Do: DoEvents
        Call SendMessage(ChatigWin&, WM_CLOSE, 0&, 0&)
        ChatigWin& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", FullSn$)
    Loop Until (ChatigWin& = 0&)
End Function

Public Sub Form_OnTop(Frm As Form, TopMost As Boolean)
    'Sets a form to *top_most* or *not_top_most*
    Dim xFlag As Long
    If TopMost = True Then xFlag = HWND_TOPMOST Else xFlag = HWND_NOTOPMOST
    Call SetWindowPos(Frm.hwnd, xFlag, 0, 0, 0, 0, FLAGS)
End Sub

Public Function GetCaption(Win As Long) As String
    'Gets the caption of a window
    Dim WindowCapLen As Long, Cap As String
    WindowCapLen = GetWindowTextLength(Win&)
    Cap$ = String$(WindowCapLen, 0&)
    Call GetWindowText(Win&, Cap$, (WindowCapLen + 1))
    GetCaption$ = Cap$
End Function

Public Sub AOLImageKill()
    'Closes all open AOL Images
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, img As Long
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    img& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    If img& = 0& Then Exit Sub
    Do Until (img& = 0&): DoEvents
        Call SendMessage(img&, WM_CLOSE, 0, 0&)
        AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
        img& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    Loop
End Sub

Public Sub RestrictedRoom(PR$)
    'restricted room enter (EX.  RestrictedRoom("hack")   )
    'aol://2719/L1:2-2-S%A0e%A0r%A0v%A0e%A0r
    Dim za As String, i As Integer
    za$ = ""
    For i = 1 To Len(PR$)
        za$ = (za$ & Mid(PR, i, 1) & "%A0")
    Next i
    Call Keyword("aol://2719/L1:2-2-" & Mid$(za$, 1, Len(za$) - 3))
End Sub

Public Sub ScreenShot(Optional Pic As PictureBox, Optional sFile As String)
    'Takes a screen shot by pressing the Print Screen key
    Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&) 'Execute Print Screen key
    If CBool(IsMissing(Pic)) = False Then TimeOut (0.3): Pic.Picture = Clipboard.GetData(vbCFBitmap) Else Exit Sub 'Place in Picturebox
    If CBool(sFile$ = "") = False Then DoEvents: Call SavePicture(Pic.Picture, sFile$) 'Save Picturebox image
End Sub

Public Sub ScreenShotWindow(Win As Long, Optional Pic As PictureBox, Optional sFile As String)
    'Takes a screen shot of a window by pressing the Print Screen key
    On Error Resume Next
    AppActivate (GetCaption(Win&)): DoEvents
    Call SendMessage(Win&, WM_KEYDOWN, VK_MENU, 0&)
    Call keybd_event(VK_SNAPSHOT, 0, 0&, 0&) 'Execute Print Screen key
    Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0&) 'Execute Print Screen key
    Call SendMessage(Win&, WM_KEYUP, VK_MENU, 0&)
    If CBool(IsMissing(Pic) = True) = False Then
        TimeOut (0.3)
        Pic.Picture = Clipboard.GetData(vbCFBitmap) 'Place in Picturebox
    Else
        Exit Sub
    End If
    If CBool(sFile$ = "") = False Then DoEvents: Call SavePicture(Pic.Picture, sFile$) 'Save Picturebox image
End Sub

Public Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
    'Gives a precent
    On Error Resume Next
    Percent = Int(Complete / Total * TotalOutput)
End Function
Public Sub PercentBar(Shape As Control, Done As Integer, Total As Long)
    'A Precent Bar
    On Error Resume Next
    Dim X As String
    Shape.AutoRedraw = True
    Shape.FillStyle = 0
    Shape.DrawStyle = 0
    Shape.FontName = "Arial Narrow"
    Shape.FontSize = 8
    Shape.FontBold = False
    X = (Done / Total * Shape.Width)
    Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
    Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
    Shape.CurrentX = ((Shape.Width / 2) - 100)
    Shape.CurrentY = ((Shape.Height / 2) - 125)
    Shape.ForeColor = RGB(255, 0, 0)
    Shape.Print Percent(Done, Total&, 100) & "%"
End Sub

Public Function GetFromINI(lSection As String, lKey As String, lDirectory As String) As String
    'Gets the key value from a *.ini file
    Dim lstrBuffer As String
    lstrBuffer = String$(750, Chr$(0))
    lKey$ = LCase$(lKey$)
    GetFromINI$ = Left$(lstrBuffer, GetPrivateProfileString(lSection$, ByVal lKey$, "", lstrBuffer, Len(lstrBuffer), lDirectory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyVal As String, Directory As String)
    'Writes a key value to a *.ini file
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyVal$, Directory$)
End Sub

Public Function WindowsDirectory() As String
    'Returns the path of the installed Windows directory
    Dim buffer As String
    Dim sLength As Long
    buffer$ = Space$(255)
    sLength = GetWindowsDirectory(buffer$, 255)
    WindowsDirectory$ = Left$(buffer$, sLength)
End Function

Public Function AOLDirectory() As String
    'Returns the path of AOL's installed directory
    On Error Resume Next
    AOLDirectory$ = ""
    AOLDirectory$ = GetFromINI("WAOL", LCase$("AppPath"), (WindowsDirectory$ & "\win.ini"))
End Function
Public Function AOLVersion() As Integer
    'Retrieves version from AOL version.inf file (latest installed version)
    'If AOLVersion2 = 0 then the version was not found
    On Error Resume Next
    AOLVersion = 0
    AOLVersion = GetFromINI("Application Information", "Version", (AOLDirectory$ & "\version.inf"))
End Function

Public Function AOLVersion2() As Integer
    'Retrieves version from AOL Help menu
    'If AOLVersion2 = 0 then the version was not found or is less than 4.0
    Dim AOLWindow As Long, AOLMenu As Long, MenuCount As Long
    Dim LookFor As Long, SubMenu As Long, SubCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOLVersion2 = 0
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMenu& = GetMenu(AOLWindow&)
    If AOLMenu& = 0& Then Exit Function
    MenuCount& = GetMenuItemCount(AOLMenu&)
    For LookFor& = 0& To MenuCount& - 1
        SubMenu& = GetSubMenu(AOLMenu&, LookFor&)
        SubCount& = GetMenuItemCount(SubMenu&)
        For LookSub& = 0 To (SubCount& - 1)
            sID& = GetMenuItemID(SubMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(SubMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase$(sString$), LCase$("&What's New in AOL")) > 0& Then
                sString$ = Trim$(sString$)
                AOLVersion2 = Val(Mid(sString$, Len(sString$) - 3, Len(sString$) - 2))
                Exit Function
            End If
        Next LookSub&
    Next LookFor&
End Function

Public Function GetFreeDiskSpace(Drive As String) As Currency
    'Returns the free disk space of a sepcified drive
    Dim UserBytes As ULARGE_INTEGER
    Dim TotalBytes As ULARGE_INTEGER
    Dim FreeBytes As ULARGE_INTEGER
    Dim TempVal As Currency
    Dim RetVal As Long
    RetVal = GetDiskFreeSpaceEx(Drive$, UserBytes, TotalBytes, FreeBytes)
    Call CopyMemory(TempVal, FreeBytes, 8)
    GetFreeDiskSpace = (TempVal * 10000)
End Function

Public Sub Window_Enable(Win As Long, Enable As Boolean)
    'Enables a given window
    If CBool(IsWindowEnabled(Win&)) <> Enable Then Call EnableWindow(Win&, Enable)
End Sub

Public Sub Window_Close(Win As Long)
    'Closes a given window
    Call SendMessage(Win&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub Window_Minimize(Win As Long)
    'Minimizes a given window
    Call ShowWindow(Win&, SW_MINIMIZE)
End Sub

Public Sub Window_Maximize(Win As Long)
    'Maximizes a given window
    Call ShowWindow(Win&, SW_MAXIMIZE)
End Sub

Public Sub Window_Restore(Win As Long)
    'Retores a given window
    Call ShowWindow(Win&, SW_RESTORE)
End Sub

Public Sub Window_Hide(Win As Long, Hide As Boolean)
    'Hides/Shows a given window
    Dim xFlag As Long
    If Hide = True Then xFlag = SW_HIDE Else xFlag = SW_SHOW
    Call ShowWindow(Win&, xFlag)
End Sub

Public Function PlayWave(file As String, LoopWave As Boolean) As Long
    'Play a *.wav file
    Dim wFlags As Long
    If LoopWave = True Then wFlags = SND_SYNC And SND_LOOP Else wFlags = SND_SYNC
    PlayWave = sndPlaySound(file$, wFlags)
End Function

Public Sub HideAOLPlus(Hide As Boolean)
    'Hide the AOL Plus window
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, hFlag As Long
    Dim PlusCap As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    PlusCap$ = GetCaption(AOLChild&)
    If Trim$(PlusCap$) = Trim$("AOL Plus") Then
        AOLChild& = AOLChild&
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            PlusCap$ = GetCaption(AOLChild&)
            If Trim$(PlusCap$) = Trim$("AOL Plus") Then
                AOLChild& = AOLChild&
                Exit Do
            End If
        Loop Until AOLChild& = 0&
    End If
    If Hide = True Then hFlag = SW_HIDE Else hFlag = SW_SHOW
    Call ShowWindow(AOLChild&, hFlag)
End Sub

Public Sub HideWelcome(Hide As Boolean)
    'Hide AOL's Welcome window
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, hFlag As Long
    Dim Caption As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(AOLChild&)
    If (Mid$(Caption$, 1, 9) = "Welcome, ") And (Right$(Caption$, 1) = "!") Then
        AOLChild& = AOLChild&
    Else
        Do
            AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
            Caption$ = GetCaption(AOLChild&)
            If (Mid$(Caption$, 1, 9) = "Welcome, ") And (Right$(Caption$, 1) = "!") Then
                AOLChild& = AOLChild&
                Exit Do
            End If
        Loop Until AOLChild& = 0&
    End If
    If Hide = True Then hFlag = SW_HIDE Else hFlag = SW_SHOW
    Call ShowWindow(AOLChild&, hFlag)
End Sub

Public Sub LinkSend(Description As String, URL As String)
    'Sends a link to an AOL chatroom
    ChatSend ("< A HREF=" & URL$ & ">" & Description$ & "</a>")
End Sub

Public Sub KillGlyph()
    'Close the spinning AOL icon in the toolbar
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLGlyph As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLGlyph& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Glyph", vbNullString)
    Call SendMessage(AOLGlyph&, WM_CLOSE, 0&, 0&)
End Sub

Public Function UserOnline() As Boolean
    'Retrieves online status; true or false
    Dim AOLWindow As Long, AOLMenu As Long, MenuCount As Long
    Dim LookFor As Long, SubMenu As Long, SubCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    UserOnline = False
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMenu& = GetMenu(AOLWindow&)
    If AOLMenu& = 0& Then Exit Function
    MenuCount& = GetMenuItemCount(AOLMenu&)
    For LookFor& = 0& To MenuCount& - 1
        SubMenu& = GetSubMenu(AOLMenu&, LookFor&)
        SubCount& = GetMenuItemCount(SubMenu&)
        For LookSub& = 0 To SubCount& - 1
            sID& = GetMenuItemID(SubMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(SubMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase$(sString$), LCase$("Sign &Off")) > 0& Then
                UserOnline = True
                Exit Function
            End If
        Next LookSub&
    Next LookFor&
    UserOnline = False
End Function

Public Sub Form_Move(Frm As Form)
    'Moves a form
    Dim ret As Long
    DoEvents
    ReleaseCapture
    ret = SendMessage(Frm.hwnd, &HA1, 2, 0&)
End Sub

Public Function FileExists(FileName As String) As Boolean
    'Checks for the existense of a file; true or false
    If Len(FileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(FileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Sub Form_UnloadAll(sFormName As String)
    'Unloads all forms in project
    Dim Form As Form
    For Each Form In Forms
        If Form.name <> sFormName$ Then
              Unload Form
              Set Form = Nothing
         End If
     Next Form
 End Sub

Public Sub TimeOut(interval)
    'This pauses a program
    Dim current As Long
    current = Timer
    Do While Timer - current < Val(interval)
        DoEvents
    Loop
End Sub

Public Sub OpenAOL()
    'Retrieves path from win.ini file
    'Shells aol.exe
    On Error Resume Next
    Call Shell((AOLDirectory$ & "\aol.exe"), vbMaximizedFocus)
End Sub

Public Sub ChatClear()
    'Removes all text from an AOL chatroom
    Call SendMessageByString(FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString), WM_SETTEXT, 0&, "")
End Sub

Public Sub Keyword(KW As String)
    'AOL Keyword
    If KW$ = "" Then Exit Sub
    Dim AOLHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLCombo As Long, AOLEdit As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLCombo& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit& = FindWindowEx(AOLCombo&, 0&, "Edit", vbNullString)
    If AOLEdit& = 0& Then Exit Sub
    Do Until (GetText(AOLEdit&) = KW$): DoEvents
        Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "" & KW$ & "")
    Loop
    Call SendMessage(AOLEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessage(AOLEdit&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub SearchAOL(KW As String)
    'Search from AOL's toolbar
    If KW$ = "" Then Exit Sub
    Dim AOLHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLEdit As Long
    Dim i As Integer
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLEdit& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Edit", vbNullString)
    For i = 1 To 4
        AOLEdit& = FindWindowEx(AOLToolbar2&, AOLEdit&, "_AOL_Edit", vbNullString)
    Next i
    If AOLEdit& = 0& Then Exit Sub
    Do Until (GetText(AOLEdit&) = KW$): DoEvents
        Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "" & KW$ & "")
    Loop
    Call SendMessage(AOLEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessage(AOLEdit&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Function RoomCount() As Integer
    'Returns an AOL chatroom, room count; from the screen name listbox
    RoomCount = Int(SendMessage(FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString), LB_GETCOUNT, 0&, 0&))
End Function

Public Function RoomCount2() As Integer
    'Returns an AOL chatroom, room count; from text in the chat window
    Dim ChatStat As Long
    Dim i As Integer
    ChatStat& = FindWindowEx(FindRoom&, 0&, "_AOL_Static", vbNullString)
    For i = 1 To 2
        ChatStat& = FindWindowEx(FindRoom&, ChatStat&, "_AOL_Static", vbNullString)
    Next i
    RoomCount2 = Val(GetText(ChatStat&))
End Function

Public Sub AntiIdle()
    'Use in timer
    Dim AOLModal As Long, AolStatic As Long, Icn As Long
    Dim X As String
    AOLModal& = FindWindow("_AOL_Modal", vbNullString)
    X$ = Mid$(GetText(FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)), 1, 2)
    If X$ = "Do" Then
        Icn& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
        ClickIcon (Icn&)
    End If
End Sub

Public Sub Kill45Timer()
    'Closes AOL's 45 Minute timer
    Dim X As Long
    X& = FindWindow("_AOL_Timer", vbNullString)
    If X& <> 0& Then
        Call SendMessage(X&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub AddIMSNtoList(Lst As ListBox, AddAtAOLdotCOM As Boolean)
    'Add the Screen Name of all open IMs to a Listbox
    Dim AOLWindow As Long, AOLMDI As Long, AOLChild As Long, hFlag As Long
    Dim PlusCap As String, AOLdotCOM As String, Txt As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLWindow&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(AOLMDI&, 0&, "AOL Child", vbNullString)
    Txt$ = GetCaption(AOLChild&)
    If (Left$(Txt$, 3) = " IM" Or Left$(Txt$, 3) = ">IM") And (InStr(Txt$, ":") > 0&) Then
        AOLChild& = AOLChild&
    End If
    If AddAtAOLdotCOM = True Then AOLdotCOM = "@aol.com" Else AOLdotCOM = ""
    Do: DoEvents
        Lst.AddItem ((Trim$(Mid$(Txt$, InStr(Txt$, ":") + 1)) & AOLdotCOM$))
        AOLChild& = FindWindowEx(AOLMDI&, AOLChild&, "AOL Child", vbNullString)
    Loop Until (AOLChild& <> 0&)
End Sub

Public Function PrivateRoom(rName As String) As Boolean
    'Enters a private AOL chatroom
    'Returns true or false depending on the success of the function
    If rName$ = "" Then Exit Function
    Call Keyword("aol://2719:2-2-" & rName$)
    If (FindRoom& = 0&) Or (Not (GetCaption(FindRoom&) Like rName$)) Then PrivateRoom = False Else PrivateRoom = True
End Function

Public Function EnterRoom(rName As String, RoomType As CHAT_TYPE) As Boolean
    'Enters a private, public, member, or restricted AOL chatroom
    'Returns true or false depending on the success of the function
    If rName$ = "" Then Exit Function
    Select Case RoomType
        Case CHAT_TYPE.cPrivate
            Call Keyword("aol://2719:2-2-" & Trim$(rName$))
        Case CHAT_TYPE.cPublic
            Call Keyword("aol://2719:21-2-" & Trim$(rName$))
        Case CHAT_TYPE.cMember
            Call Keyword("aol://2719:61-2-" & Trim$(rName$))
        Case CHAT_TYPE.cRestricted
            Call RestrictedRoom(Trim$(rName$))
    End Select
    WaitForOKorRoom
    EnterRoom = Not CBool((FindRoom& = 0&) Or (RemoveSpaces(GetCaption(FindRoom&)) Like RemoveSpaces(rName$)))
End Function

Public Sub ImOnOff(imStat As ImOn_Off)
    'Turns instant messages on or off
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLIMWindow As Long
    Dim OnOff As String
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    If (AOLHandle& = 0&) Or (AOLMDiHandle& = 0&) Then Exit Sub
    If imStat = imOff Then
        OnOff$ = "$IM_ON"
    ElseIf imStat = imON Then
        OnOff$ = "$IM_OFF"
    End If
    Call IMSend(OnOff$, "PuNkDuDe")
    DoEvents
    Do: DoEvents
        AOLIMWindow& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", (" IM To: " & Trim$(LCase$(OnOff$))))
    Loop Until (AOLIMWindow& <> 0&)
    Call SendMessage(AOLIMWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ImIgnore(SN As String, imStat As ImOn_Off)
    'Turns instant messages on or off
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLIMWindow As Long
    Dim OnOff As String
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    If (AOLHandle& = 0&) Or (AOLMDiHandle& = 0&) Then Exit Sub
    If imStat = imOff Then
        OnOff$ = "$IM_ON " & Trim$(SN$)
    ElseIf imStat = imON Then
        OnOff$ = "$IM_OFF " & Trim$(SN$)
    End If
    Call IMSend(OnOff$, "PuNkDuDe")
    DoEvents
    Do: DoEvents
        AOLIMWindow& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", (" IM To: " & Trim$(LCase$(OnOff$))))
    Loop Until (AOLIMWindow& <> 0&)
    Call SendMessage(AOLIMWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub HideCtrlAltDel(Hide As Boolean)
    'Hide program from CTRL ALT DEL window
    Dim regserv As Long, xFlag As Long
    If Hide = True Then xFlag = RSP_SIMPLE_SERVICE Else xFlag = RSP_UNREGISTER_SERVICE
    regserv = RegisterServiceProcess(GetCurrentProcessId(), xFlag)
End Sub

Public Sub DisableCtrlAltDel(Disable As Boolean)
    'Disables/Enables Ctrl+Alt+Del
    Dim ret As Long
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, CLng(Disable), 0&, 0&)
End Sub

Public Function AOLWin() As Long
    'Returns the handle of AOL's main window
    AOLWin& = FindWindow("AOL Frame25", vbNullString)
End Function

Public Function AOLMDI() As Long
    'Returns the handle of AOL's MDIClient window
    AOLMDI& = FindWindowEx(FindWindow("AOL Frame25", vbNullString), 0&, "MDIClient", vbNullString)
End Function

Public Function Random(Number As Integer) As Integer
    'Returns a random integer value
    Randomize
    Random = Int((Val(Number) * Rnd) + 1)
End Function

Public Sub ClickIcon(Icn As Long)
    'Clicks AOL icons
    Call SendMessage(Icn&, WM_LBUTTONDOWN, 0&, 0&) 'Activate button (required for AOL 6.0)
    Call SendMessage(Icn&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(Icn&, WM_KEYDOWN, VK_SPACE, 0&) 'Click button
    Call SendMessage(Icn&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub ClickRadio(RadioButton As Long)
    'Clicks a radio/option button
    Call SendMessageLong(RadioButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(RadioButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub ClickCheck(CheckBox As Long)
    'Clicks a checkbox
    Call SendMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub FillAOLMDI(cRed As Long, cGreen As Long, cBlue As Long)
    'Fill AOL's MDI window with a given color value
    Dim rct As RECT
    Dim hBrush As Long
    Call GetWindowRect(AOLMDI&, rct)
    rct.Top = 0: rct.Left = 0: rct.Right = (rct.Right - rct.Left): rct.Bottom = (rct.Bottom - rct.Top)
    hBrush& = CreateSolidBrush(RGB(cRed, cGreen, cBlue))
    Call FillRect(GetDC(AOLMDI&), rct, hBrush&)
End Sub

Public Sub FillAOLMdiPic(Pic As PictureBox)
    'Fills (tiles) AOL's MDI with a picture
    If AOLMDI& = 0& Then Exit Sub
    On Error Resume Next
    Dim rct As RECT
    Dim AoMDI As Long
    Dim d As Integer, a As Integer
    Dim iWidth As Long, iHeight As Long
    Pic.AutoSize = True
    Pic.ScaleMode = 3
    Call GetWindowRect(AOLMDI&, rct)
    iWidth = (rct.Right - rct.Left)
    iHeight = (rct.Bottom - rct.Top)
    AoMDI& = AOLMDI&
    For d = 0 To iWidth Step Pic.ScaleWidth
        For a = 0 To iHeight Step Pic.ScaleHeight
            Call BitBlt(GetDC(AoMDI&), ByVal CLng(d), ByVal CLng(a), Pic.ScaleWidth, Pic.ScaleHeight, Pic.hDC, 0&, 0&, SRCCOPY)
        Next a
    Next d
    Pic.AutoSize = False
End Sub

Public Sub FillWindowRect(Window As Long, cRed As Long, cGreen As Long, cBlue As Long)
    'Fills a window with the given color
    Dim rct As RECT
    Dim hBrush As Long
    Call GetWindowRect(Window&, rct)
    rct.Top = 0: rct.Left = 0: rct.Right = (rct.Right - rct.Left): rct.Bottom = (rct.Bottom - rct.Top)
    hBrush& = CreateSolidBrush(RGB(cRed, cGreen, cBlue))
    Call FillRect(GetDC(Window&), rct, hBrush&)
End Sub

Public Sub LocateMember(SN As String)
    'Gets the location of a member
    Call Keyword("aol://3548:" & RemoveSpaces$(SN$))
End Sub

Public Function LocateMember2(SN As String) As Boolean
    'Returns the location of a member; true or false
    Dim AOLMsg As Long, AOLLocate As Long
    Dim AOLMsgText As String
    If SN$ <> "" Then SN$ = Trim$(SN$) Else Exit Function
    Call Keyword("aol://3548:" & SN$)
    Do: DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        AOLLocate& = FindWindowEx(AOLMDI&, 0&, "AOL Child", ("Locate " & SN$))
    Loop Until ((AOLMsg& <> 0&) Or (AOLLocate& <> 0&))
    If (AOLMsg& <> 0&) And (AOLLocate = 0&) Then
        Call SendMessage(AOLMsg&, WM_CLOSE, 0&, 0&)
        LocateMember2 = False
    ElseIf (AOLLocate& <> 0&) And (AOLMsg& = 0&) Then
        Call SendMessage(AOLLocate&, WM_CLOSE, 0&, 0&)
        LocateMember2 = True
    End If
End Function

Public Function AOLLaunches() As Integer
    'Number of times AOL has been launched
    On Error Resume Next
    AOLLaunches = 0
    AOLLaunches = GetFromINI("Client Info", "nLaunches", (AOLDirectory$ & "\Status.ini"))
End Function

Public Function AOLSignOns() As Integer
    'Number of times AOL has been signed on
    On Error Resume Next
    AOLSignOns = 0
    AOLSignOns = GetFromINI("Client Info", "nSignOns", (AOLDirectory$ & "\Status.ini"))
End Function

Public Sub waitforok()
    'Waits until an AOL msgbox is opened and the closes it
    Dim AOLMsg As Long
    Do: DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        If AOLMsg& <> 0& Then
            Call SendMessage(AOLMsg&, WM_CLOSE, 0&, 0&): Exit Do
        End If
    Loop
End Sub

Public Sub WaitForOKorRoom()
    Dim AOLMsg As Long
    Do: DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        If AOLMsg& <> 0& Then
            Call SendMessage(AOLMsg&, WM_CLOSE, 0&, 0&): Exit Do
        End If
        If FindRoom& <> 0& Then Exit Do
    Loop
End Sub

Public Sub WaitForOKorModal()
    Dim MailWindow As Long, AOLMsg As Long, AOLModal As Long, AOLModalIcon As Long
    Do: DoEvents
        AOLModal& = FindWindow("_AOL_Modal", vbNullString)
        AOLModalIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
        AOLMsg& = FindWindowEx(AOLWin&, 0&, "#32770", "America Online")
        If AOLModalIcon& <> 0& Then ClickIcon (AOLModalIcon&): Exit Do
        If AOLMsg& <> 0& Then Call SendMessage(AOLMsg&, WM_CLOSE, 0&, 0&): Exit Do
    Loop
End Sub

Public Sub WaitForRoom()
    'Loops until a AOL chatroom is found
    Do: DoEvents
    Loop Until (FindRoom& <> 0&)
End Sub

Public Sub AddRoomToList(lLst As ListBox, AddUser As Boolean)
    'AOL 4.0 & 5.0
    Dim rList As Long
    Dim i As Integer
    Dim lText As String
    rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    For i = 0 To (SendMessage(rList&, LB_GETCOUNT, 0&, 0&) - 1)
        If i > (SendMessage(rList&, LB_GETCOUNT, 0&, 0&) - 1) Then Exit Sub
        lText$ = GetListText(rList&, i)
        If (lText$ = UserSN$) And (AddUser = False) Then Resume Next Else lLst.AddItem (lText$)
    Next i
End Sub

Public Sub PreventMultipleInstance()
    'If program is opened more than once, this will close the newely opened one
    'Place in Form_Load event of main form
    If App.PrevInstance = True Then End
End Sub

Public Sub AOLClearHistory()
    'Clears the content of the web address bar
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long, AOLToolbar2 As Long, AOLCombo As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
    AOLCombo& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
    If AOLCombo& = 0& Then Exit Sub
    Call SendMessage(AOLCombo&, CB_RESETCONTENT, 0&, 0&)
End Sub

Public Sub HideToolbar(Hide As Boolean)
    'Hides/shows AOL's toolbar
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLToolbar As Long
    Dim xFlag As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLHandle&, 0&, "AOL Toolbar", vbNullString)
    If AOLToolbar& = 0& Then Exit Sub
    If Hide = True Then xFlag = SW_HIDE Else xFlag = SW_SHOW
    Call ShowWindow(AOLToolbar&, xFlag)
End Sub

Public Sub RemovePWChar()
    'Remove the Password character of the password box of AOL's Sign On window
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLSign As Long, AOLEdit As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLSign& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", "Sign On")
    AOLEdit& = FindWindowEx(AOLSign&, 0&, "_AOL_Edit", vbNullString)
    If AOLEdit& = 0& Then Exit Sub
    Call SendMessage(AOLEdit&, EM_SETPASSWORDCHAR, 0&, 0&)
End Sub

Public Sub SetPWChar(Optional Char As String)
    'Set the Password character of the password box of AOL's Sign On window
    Dim AOLHandle As Long, AOLMDiHandle As Long, AOLSign As Long, AOLEdit As Long
    AOLHandle& = FindWindow("AOL Frame25", vbNullString)
    AOLMDiHandle& = FindWindowEx(AOLHandle&, 0&, "MDIClient", vbNullString)
    AOLSign& = FindWindowEx(AOLMDiHandle&, 0&, "AOL Child", "Sign On")
    AOLEdit& = FindWindowEx(AOLSign&, 0&, "_AOL_Edit", vbNullString)
    If CBool(IsMissing(Char)) = True Then Char$ = PW_CHAR
    Call SendMessage(AOLEdit&, EM_SETPASSWORDCHAR, ByVal CLng(Asc(Char$)), 0&)
End Sub

Public Sub WinZipPW(Txt As String)
    'Enter the password for a winzip file
    Dim WinPW As Long, edit As Long, Button As Long
    WinPW& = FindWindow("#32770", "Password")
    edit& = FindWindowEx(WinPW&, 0&, "Edit", vbNullString)
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, Txt$)
    Button& = FindWindowEx(WinPW&, 0&, "Button", "OK")
    If Button& = 0& Then Exit Sub
    Do Until (CBool(IsWindowEnabled(Button&)) = True): DoEvents
    Loop
    Call AppActivate(GetCaption(WinPW&))
    Call ClickIcon(Button&)
End Sub

Public Function Text_Wavy(Txt As String, Optional WavyUp As Boolean) As String
    'If WavyUp=True then the wavy starts low; If WavyUp=False then the wavy starts high
    Dim tTxt As String, nTxt As String, lTxt As String, Script As String
    Dim X As Double
    tTxt$ = Txt$: lTxt$ = "": X = 0
    Do While (Len(tTxt$) > 0&): DoEvents
        X = X + 1
        nTxt$ = Left$(tTxt$, 1): tTxt$ = Right$(tTxt$, (Len(tTxt$) - 1))
        If (IsMissing(WavyUp) = True) Or (WavyUp = True) Then
            If X Mod 2 = 0 Then Script$ = "<sup>" Else Script$ = "</sup>"
        Else
            If X Mod 2 = 0 Then Script$ = "<sub>" Else Script$ = "</sub>"
        End If
        lTxt$ = (lTxt$ & Script$ & nTxt$)
    Loop
    Text_Wavy$ = lTxt$
End Function

Public Function ReplaceString(CheckString As String, FindIn As String, NewString As String) As String
    'Replace one string within another
    'EX:   X = ReplaceString("dude", "punkdude", "guy")
    Dim tTxt As String, nTxt As String
    Dim X As Integer, i As Integer
    nTxt$ = FindIn$
    Do Until InStr(nTxt$, CheckString$) = 0&: DoEvents
        nTxt$ = Replace(FindIn$, CheckString$, NewString$)
    Loop
    ReplaceString$ = nTxt$
End Function

Public Function ReverseString(Txt As String) As String
    'Returns the string reversed
    Dim nTxt As String, tTxt As String
    Dim i As Integer
    tTxt$ = Txt$
    For i = 1 To Len(Txt$)
        nTxt$ = (nTxt$ & Right$(tTxt$, 1))
        tTxt$ = Left$(Txt$, Len(Txt$) - i)
    Next i
    ReverseString$ = nTxt$
End Function

Public Function RemoveSpaces(Txt As String) As String
    'Remove all of the spaces from a string
    If Txt$ = "" Then Exit Function
    Dim nTxt As String
    nTxt$ = Trim$(Txt$)
    Do Until (InStr(nTxt$, " ") = 0&): DoEvents
        nTxt$ = Replace(nTxt$, " ", "")
        If InStr(nTxt$, " ") > 0& Then Exit Do
    Loop
    RemoveSpaces$ = nTxt$
End Function

Public Function RemoveNull(Txt As String) As String
    'Remove all of the spaces from a string
    Dim nTxt As String
    nTxt$ = Trim$(Txt$)
    Do While (InStr(nTxt$, vbNullChar) > 0&): DoEvents
        nTxt$ = Replace(nTxt$, vbNullChar, "")
    Loop
    RemoveNull$ = nTxt$
End Function

Public Function Text_Lag(Txt As String) As String
    'Returns a srting containing <html> and </html>
    Dim nTxt As String, tTxt As String, oTxt As String
    Dim i As Integer
    nTxt$ = Trim$(Txt$): oTxt$ = ""
    For i = 1 To Len(nTxt$)
        tTxt$ = Mid$(nTxt$, i, 1)
        oTxt$ = (oTxt$ & "<html>" & tTxt$ & "</html>")
    Next i
    Text_Lag$ = Trim$(oTxt$)
End Function

Public Sub AddNames(AOLList As Long, Lst As ListBox, AddAtAOL As Boolean)
    'By Bone
    'AddAtAOL=True will add @AOL.COM to to the list along with the screen name
    On Error Resume Next
    Dim xFlag As String
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim Ta As Long, Ta2 As Long
    rList& = AOLList&
    If rList& = 0& Then Exit Sub
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To (SendMessage(rList&, LB_GETCOUNT, 0, 0) - 1)
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            Ta& = InStr(1, ScreenName$, Chr(9))
            Ta2& = InStr((Ta& + 1), ScreenName$, Chr(9))
            ScreenName$ = Mid$(ScreenName$, (Ta& + 1), (Ta2& - 2))
            ScreenName$ = Right$(ScreenName$, Len(ScreenName$) - InStr(ScreenName$, Chr(9)))
            If AddAtAOL = True Then xFlag$ = "@AOL.COM" Else xFlag$ = ""
            Lst.AddItem (RemoveNull(ScreenName$) & xFlag$)
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub PictureToDesktop(Pic As PictureBox, Optional X As Single, Optional Y As Single)
    'Copies the contents of a picturebox to the desktop
    If IsMissing(X) = True Then X = 0: If IsMissing(Y) = True Then Y = 0
    Pic.AutoRedraw = True
    Pic.ScaleMode = 3
    Call BitBlt(GetWindowDC(GetDesktopWindow&), ByVal CLng(X), ByVal CLng(Y), Pic.ScaleWidth, Pic.ScaleHeight, Pic.hDC, 0, 0, SRCCOPY)
End Sub

Public Sub Window_TopMost(Window As Long, TopMost As Boolean)
    'Sets a window to the topmost position
    If TopMost = True Then
        Call SetWindowPos(Window&, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        Call SetWindowPos(Window&, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Public Sub SignOnGuest(SN As String, PW As String)
    'Signs on AOL as a guest
    On Error GoTo guest_error
    Dim SignOn As Long, Combo As Long, ret As Integer
    Dim Modal As Long, Stat As Long, edit As Long, Edit2 As Long, Icon As Long
    If (UserOnline = True) Or (Len(SN$) < 3) Or (PW$ = "") Then Exit Sub
    If FindWindowEx(AOLMDI&, 0&, "AOL Child", "Goodbye from America Online!") Then Call SendMessage(FindWindowEx(AOLMDI&, 0&, "AOL Child", "Goodbye from America Online!"), WM_CLOSE, 0&, 0&): setactivewindow (AOLWin&): SendKeys ("&OO"): DoEvents 'Re-open sign on window
    SignOn& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Sign On")
    Combo& = FindWindowEx(SignOn&, 0&, "_AOL_Combobox", vbNullString)
    If Combo& = 0& Then Exit Sub
    ret = SendMessage(Combo&, CB_GETCOUNT, 0&, 0&)
    Call SendMessage(Combo&, CB_SETCURSEL, (CLng(ret) - 1&), 0&)
    ClickIcon (Combo&)
    Do: DoEvents: Loop Until CBool(IsWindowVisible(FindWindowEx(SignOn&, 0&, "_AOL_Edit", vbNullString))) = False
    Call SendMessage(Combo&, WM_CHAR, VK_RETURN, 0&)
    Do: DoEvents
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        Stat& = FindWindowEx(Modal&, 0&, "_AOL_Static", "Guest Sign-On:")
    Loop Until Modal& <> 0& And Stat& <> 0&
    edit& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, SN$)
    Edit2& = FindWindowEx(Modal&, edit&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(Edit2&, WM_SETTEXT, 0&, PW$)
    Icon& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Icon&)
    Exit Sub
guest_error: Exit Sub
End Sub

Public Function AIM_Win() As Long
    'Returns the handle of AIM's buddylist window
    AIM_Win& = FindWindow("_Oscar_BuddyListWin", vbNullString)
End Function
Public Function AIM_UserSn() As String
    'Returns AIM's screen name from the caption of the main window
    Dim AIMWindow As Long
    Dim tSN As String
    AIMWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If AIMWindow& = 0& Then AIM_UserSn = "": Exit Function
    tSN$ = GetCaption(AIMWindow&)
    AIM_UserSn$ = Trim$(Left$(tSN$, (InStr(tSN$, "'s Buddy List Window") - 1)))
End Function

Public Sub AIM_IMSend(SN As String, Message As String)
    'Sends an AIM instant message
    If (Len(SN$) < 3) Or (Message$ = "") Then Exit Sub
    Dim aCap As String
    Dim AIMWindow As Long, TabGroup As Long, Tree As Long, AIMIcon As Long
    Dim IMWindow As Long, Combo As Long, edit As Long, WndAte As Long, Ate32 As Long, Icon As Long
    AIMWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabGroup& = FindWindowEx(AIMWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    Tree& = FindWindowEx(TabGroup&, 0&, "_Oscar_Tree", vbNullString)
    If Tree& = 0& Then Exit Sub
    Do: DoEvents
        IMWindow& = FindWindowEx(IMWindow&, 0&, "AIM_IMessage", vbNullString)
        aCap$ = Trim$(LCase$(RemoveSpaces(Left$(GetCaption(IMWindow&), InStr(GetCaption(IMWindow&), " - Instant Message")))))
        If aCap$ = Trim$(LCase$(RemoveSpaces(SN$))) Then Exit Do
    Loop While (IMWindow& <> 0&)
    If aCap$ <> Trim$(LCase$(RemoveSpaces(SN$))) Then
        Call SendMessage(Tree&, LB_SETCURSEL, 0&, 0&)
        DoEvents
        AIMIcon& = FindWindowEx(TabGroup&, 0&, "_Oscar_IconBtn", vbNullString)
        ClickIcon (AIMIcon&)
        Do: DoEvents
            IMWindow& = FindWindow("AIM_IMessage", "Instant Message")
        Loop Until (IMWindow& <> 0&)
        Combo& = FindWindowEx(IMWindow&, 0&, "_Oscar_PersistantCombo", vbNullString)
        edit& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
        Call SendMessageByString(edit&, WM_SETTEXT, 0&, SN$)
    End If
    WndAte& = FindWindowEx(IMWindow&, 0&, "WndAte32Class", "AteWindow")
    WndAte& = FindWindowEx(IMWindow&, WndAte&, "WndAte32Class", "AteWindow")
    Ate32& = FindWindowEx(WndAte&, 0&, "Ate32Class", vbNullString)
    Call SendMessageByString(Ate32&, WM_SETTEXT, 0&, Message$)
    Icon& = FindWindowEx(IMWindow&, 0&, "_Oscar_IconBtn", vbNullString)
    ClickIcon (Icon&)
End Sub

Public Sub AIM_OpenChatInvite()
    'Opens an AIM chat invitation
    Dim AIMWindow As Long, TabGroup As Long, Tree As Long, AIMIcon As Long
    AIMWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabGroup& = FindWindowEx(AIMWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    Tree& = FindWindowEx(TabGroup&, 0&, "_Oscar_Tree", vbNullString)
    If Tree& = 0& Then Exit Sub
    AIMIcon& = FindWindowEx(TabGroup&, 0&, "_Oscar_IconBtn", vbNullString)
    AIMIcon& = FindWindowEx(TabGroup&, AIMIcon&, "_Oscar_IconBtn", vbNullString)
    ClickIcon (AIMIcon&)
End Sub

Public Sub AIM_SignOn(SN As String, PW As String)
    'Signs on to AIM
    Dim SignOn As Long, ComboB As Long, cEdit As Long, edit As Long, Icon As Long
    Dim i As Integer
    SignOn& = FindWindow("AIM_CSignOnWnd", vbNullString)
    ComboB& = FindWindowEx(SignOn&, 0&, "ComboBox", vbNullString)
    cEdit& = FindWindowEx(ComboB&, 0&, "Edit", vbNullString)
    If cEdit& = 0& Then Exit Sub
    Call SendMessage(ComboB&, CB_SETCURSEL, 1, 0&)
    For i = 0 To (SendMessage(ComboB&, CB_GETCOUNT, 0&, 0&) - 1)
        Call SendMessage(ComboB&, CB_SETCURSEL, ByVal CLng(i), 0&)
        If GetText(cEdit&) Like SN$ Then Exit For
    Next i
    Do Until CBool(IsWindowEnabled(cEdit&)) = True: DoEvents: Loop
    edit& = FindWindowEx(SignOn&, 0&, "Edit", vbNullString)
    Call SendMessageByString(edit&, WM_SETTEXT, 0&, ByVal PW$)
    Do: DoEvents: Loop Until (GetText(cEdit&) Like SN$)
    Icon& = FindWindowEx(SignOn&, 0&, "_Oscar_IconBtn", vbNullString)
    Icon& = FindWindowEx(SignOn&, Icon&, "_Oscar_IconBtn", vbNullString)
    Icon& = FindWindowEx(SignOn&, Icon&, "_Oscar_IconBtn", vbNullString)
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub AIM_HideAd()
    'Hides AIM's ad
    Dim AIMWindow As Long, WndAte As Long, Ate32 As Long
    AIMWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    WndAte& = FindWindowEx(AIMWindow&, 0&, "WndAte32Class", "AteWindow")
    If WndAte& = 0& Then Exit Sub
    Call SendMessage(WndAte&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub SetText(Window As Long, Txt As String)
    'Sets a string to a window
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, ByVal "" & Txt$ & "")
End Sub

Public Sub RunAOLMenu(Parent As String, SubS As String, Optional SubChild As String = "")
    'Runs an AOL menu by it's shortcut keys
    AppActivate (GetCaption(AOLWin&)): SendKeys ("^" & Parent$ & SubS$ & SubChild$)
End Sub

Public Sub AwayMessage(Title As String, Message As String)
    'AOL 6.0
    'Activate the Away Message
    If AOLVersion2 <> 6 Then Exit Sub
    Dim buddywin As Long, BuddyIcon As Long
    Dim AwayWin As Long, AwayList As Long, NewIcon As Long, TempIcon As Long, OkIcon As Long
    Dim NewAway As Long, NewTitle As Long, NewMess As Long, NewSave As Long
    Dim i As Integer
    Dim InList As Boolean
    Call Keyword("BuddyView")
    Do: DoEvents
        buddywin& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Buddy List")
    Loop Until (buddywin& <> 0&)
    BuddyIcon& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
    BuddyIcon& = FindWindowEx(buddywin&, BuddyIcon&, "_AOL_Icon", vbNullString)
    BuddyIcon& = FindWindowEx(buddywin&, BuddyIcon&, "_AOL_Icon", vbNullString)
    BuddyIcon& = FindWindowEx(buddywin&, BuddyIcon&, "_AOL_Icon", vbNullString)
    ClickIcon (BuddyIcon&)
    Do: DoEvents
        AwayWin& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Away Message")
    Loop Until (AwayWin& <> 0&)
    AwayList& = FindWindowEx(AwayWin&, 0&, "_AOL_Listbox", vbNullString)
    InList = False
    For i = 0 To (SendMessage(AwayList&, LB_GETCOUNT, 0&, 0&) - 1&)
        If GetListText(AwayList&, i) Like Title$ Then InList = True: Exit For
    Next i
    NewIcon& = FindWindowEx(AwayWin&, 0&, "_AOL_Icon", vbNullString) 'New icon
    TempIcon& = FindWindowEx(AwayWin&, NewIcon&, "_AOL_Icon", vbNullString)
    TempIcon& = FindWindowEx(AwayWin&, TempIcon&, "_AOL_Icon", vbNullString)
    OkIcon& = FindWindowEx(AwayWin&, TempIcon&, "_AOL_Icon", vbNullString) 'OK icon
    If InList = True Then
        Call SendMessage(AwayList&, LB_SETCURSEL, ByVal CLng(i), 0&)
        ClickIcon (OkIcon&): Exit Sub
    Else
        ClickIcon (NewIcon&)
        Do: DoEvents
            NewAway& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "New Away Message")
        Loop Until (NewAway& <> 0&)
        NewTitle& = FindWindowEx(NewAway&, 0&, "_AOL_Edit", vbNullString)
        Call SendMessageByString(NewTitle&, WM_SETTEXT, 0&, ByVal Title$)
        NewMess& = FindWindowEx(NewAway&, 0&, "RICHCNTL", vbNullString)
        Call SendMessageByString(NewMess&, WM_SETTEXT, 0&, ByVal Message$)
        NewSave& = FindWindowEx(NewAway&, 0&, "_AOL_Icon", vbNullString)
        ClickIcon (NewSave&)
        Do: DoEvents
            setactivewindow (AwayWin&)
        Loop Until (GetActiveWindow& = AwayWin&)
        ClickIcon (OkIcon&): Exit Sub
    End If
End Sub

Public Sub DisableX(Window As Long)
    'Removes the X box in the titlebar by deleting the Close menu in the system menu
    Call RemoveMenu(GetSystemMenu(Window&, 0&), GetMenuItemCount(GetSystemMenu(Window&, 0&)) - 1, MF_BYPOSITION)
End Sub

Public Function RoomName() As String
    'Returns an AOL chatroom name/caption
    If FindRoom& <> 0& Then RoomName$ = GetCaption(FindRoom&) Else RoomName$ = ""
End Function

Public Sub SetRoomFont(FontName As String)
    'Sets the chatroom font using the combobox
    Dim ComboB As Long
    Dim i As Integer
    ComboB& = FindWindowEx(FindRoom&, 0&, "_AOL_Combobox", vbNullString)
    For i = 0 To (SendMessage(ComboB&, CB_GETCOUNT, 0&, 0&) - 1&)
        Call SendMessage(ComboB&, CB_SETCURSEL, ByVal CLng(i), 0&)
        If EqualString(GetText(ComboB&), FontName$) = True Then Call SendMessage(ComboB&, CB_SETCURSEL, ByVal CLng(i), 0&): Exit For
    Next i
End Sub

Public Sub SetRoomTextOps(Optional Bold As Boolean, Optional Italic As Boolean, Optional UnderL As Boolean)
    'Activates Bold, Italic, or Underline in a chat
    If (IsMissing(Bold) = True) And (IsMissing(Italic) = True) And (IsMissing(UnderL) = True) Then Exit Sub
    Dim bBut As Long, iBut As Long, uBut As Long
    bBut& = FindWindowEx(FindRoom&, 0&, "_AOL_Icon", vbNullString)
    iBut& = FindWindowEx(FindRoom&, bBut&, "_AOL_Icon", vbNullString)
    uBut& = FindWindowEx(FindRoom&, iBut&, "_AOL_Icon", vbNullString)
    If (IsMissing(Bold) = False) And (Bold = True) Then
        Call SendMessage(bBut&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(bBut&, WM_KEYUP, VK_SPACE, 0&)
    End If
    DoEvents
    If (IsMissing(Italic) = False) And (Italic = True) Then
        Call SendMessage(iBut&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(iBut&, WM_KEYUP, VK_SPACE, 0&)
    End If
    DoEvents
    If (IsMissing(UnderL) = False) And (UnderL = True) Then
        Call SendMessage(uBut&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(uBut&, WM_KEYUP, VK_SPACE, 0&)
    End If
End Sub

Public Sub CloseIE()
    'Closes all open Internet Explorer windows
    Do While (FindWindow("IEFrame", vbNullString) <> 0&): DoEvents
        Call SendMessage(FindWindow("IEFrame", vbNullString), WM_CLOSE, 0&, 0&)
    Loop
End Sub

Public Function ComputerUserName() As String
    'Returns the user name of the local computer
    Dim Temp As String, name As String
    Temp$ = String$(255, 0&)
    Call GetUserName(Temp$, Len(Temp$))
    ComputerUserName$ = Trim$(Temp$)
End Function

Public Function ComputerName() As String
    'Returns the name of the local computer
    Dim Temp As String, name As String
    Temp$ = String$(255, 0&)
    Call GetComputerName(Temp$, Len(Temp$))
    ComputerName$ = Trim$(Temp$)
End Function

Public Function GetTotalMemory() As Long
    'Get total physical memory
    Dim mem As MEMORYSTATUS
    GlobalMemoryStatus mem
    GetTotalMemory = (mem.dwTotalPhys \ 1024)
End Function

Public Function GetFreeMemory() As Long
    'Get free physical memory
    Dim mem As MEMORYSTATUS
    GlobalMemoryStatus mem
    GetFreeMemory = (mem.dwAvailPhys \ 1024)
End Function

Public Function GetFreeMemoryPercent() As Long
    'Get free physical memory percent
    Dim mem As MEMORYSTATUS
    GlobalMemoryStatus mem
    GetFreeMemoryPercent = (100 - 100 * mem.dwAvailPhys \ mem.dwTotalPhys)
End Function

Public Function DriveType(Drive As String) As String
    'Returns a drive type of a given drive
    Select Case GetDriveType(Drive$)
        Case 1
            DriveType$ = "Doesn't Exist"
        Case DRIVE_CDROM
            DriveType$ = "CD-ROM"
        Case DRIVE_FIXED
            DriveType$ = "Hard Drive"
        Case DRIVE_RAMDISK
            DriveType$ = "RAM Disk"
        Case DRIVE_REMOTE
            DriveType$ = "Network Drive"
        Case DRIVE_REMOVABLE
            DriveType$ = "Removable"
    End Select
End Function

Public Sub CursorShow(Show As Boolean)
    'Shows/hides the mouse cursor
    Call ShowCursor(ByVal CLng(Show))
End Sub

Public Sub SetFullDrag(Full As Boolean)
    'Full=True; shows the entire contents of a window while it is being dragged
    'Full=False; shows the outline of a window while it is being dragged
    Call SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, ByVal CLng(Full), ByVal 0, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Sub

Public Sub Tray_AddIcon(Frm As Form, Icn, Optional ToolTip As String)
    'Adds an icon to the system tray (near the clock)
    If IsMissing(ToolTip) = True Then ToolTip$ = ""
    With nid
        .cbSize = Len(nid)
        .hwnd = Frm.hwnd
        .hIcon = Icn
        .szTip = ToolTip$ & vbNullChar
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallBackmessage = WM_MOUSEMOVE
        .UID = 1
    End With
    Call Shell_NotifyIcon(NIM_ADD, nid)
End Sub

Public Sub Tray_DeleteIcon()
    'Deletes an icon from the system tray (near the clock)
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Public Function AOLSoundName(Num As Integer) As String
    'Returns the AOL sound name of a given number
    If (Num <= 0) Or (Num > GetFromINI("General", "NumSounds", AOLDirectory$ & "\sounds\sounds.ini")) Then AOLSoundName = "": Exit Function
    Dim Txt As String
    AOLSoundName$ = GetFromINI(("Sound" & CStr(Num)), "Soundname", AOLDirectory$ & "\sounds\sounds.ini")
End Function
Public Sub ChangeAOLSound(SoundName As String, NewSound As String)
    'Changes AOL's default sounds
    Dim i As Integer, X As Integer
    Dim Found As Boolean
    Found = False
    X = Val(GetFromINI("General", "NumSounds", AOLDirectory$ & "\sounds\sounds.ini"))
    For i = 1 To X
       If AOLSoundName(i) = SoundName$ Then Found = True: Exit For
       If i = 10 Then Exit Sub
    Next i
    Call WriteToINI(("Sound" & CStr(i)), "Filename", NewSound$, AOLDirectory$ & "\sounds\sounds.ini")
End Sub

Public Function ListFindString(Lst As Long, Txt As String) As Long
    'Returns the index of a string within a list
    ListFindString = SendMessage(Lst&, LB_FINDSTRING, -1, ByVal CStr(Txt$))
End Function

Public Sub ListClear(Lst As Long)
    'Clears the contents of a list
    Call SendMessage(Lst&, LB_RESETCONTENT, 0&, 0&)
End Sub

Public Sub GetListInfo(Lst As Long, lType As LIST_INFO)
    'Example:
    'Dim inf as LIST_INFO
    'Call GetListInfo(inf)
    'Msgbox inf.SelText$
    lType.SelText = ""
    With lType
        .Count = SendMessage(Lst&, LB_GETCOUNT, ByVal 0&, ByVal 0&)
        .Cursel = SendMessage(Lst&, LB_GETCURSEL, ByVal 0&, ByVal 0&)
        .ItemData = SendMessage(Lst&, LB_GETITEMDATA, ByVal .Cursel, ByVal 0&)
        .ItemHeight = SendMessage(Lst&, LB_GETITEMHEIGHT, ByVal .Cursel, ByVal 0&)
        Call SendMessage(Lst&, LB_GETITEMRECT, ByVal .Cursel, ByVal .ItemRect)
        .SelCount = SendMessage(Lst&, LB_GETSELCOUNT, ByVal 0&, ByVal 0&)
        .SelText = GetListText(Lst&, CInt(.Cursel))
        .TextLen = SendMessage(Lst&, LB_GETTEXTLEN, ByVal 0&, ByVal 0&)
        .TopIndex = SendMessage(Lst&, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
    End With
End Sub

Public Sub GetAOLInfo(iType As AOL_INFO)
    'Example:
    'Dim inf as AOL_INFO
    'Call GetAOLInfo(inf)
    'Msgbox inf.sCaption$
    With iType
        .sCaption = GetCaption(AOLWin&)
        .sDir$ = AOLDirectory$
        .iLaunches = AOLLaunches
        .iSignOns = AOLSignOns
        .iVersion = AOLVersion
        .hRoomHwnd = FindRoom&
        .hMDI = AOLMDI&
        .hwnd = FindWindow("AOL Frame25", vbNullString)
        .hDC = GetDC(.hwnd)
        .sUserSN$ = UserSN$
        .bOnline = UserOnline
    End With
End Sub

Public Sub GetAIMInfo(iType As AIM_INFO)
    'Example:
    'Dim inf as AIM_INFO
    'Call GetAIMInfo(inf)
    'Msgbox inf.sCaption$
    Dim AIMWindow As Long, TabGroup As Long, Tree As Long
    AIMWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabGroup& = FindWindowEx(AIMWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    Tree& = FindWindowEx(TabGroup&, 0&, "_Oscar_Tree", vbNullString)
    With iType
        .sCaption$ = GetCaption(AIMWindow&)
        .sUserName$ = AIM_UserSn$
        .hwnd = AIMWindow&
        .hDC = GetDC(.hwnd&)
        .lCount = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
    End With
End Sub

Public Sub TrayClockColor(BackC As Long)
    'Changes the backcolor of the system clock
    Dim rct As RECT
    Dim iWidth As Long, iHeight As Long
    Dim STrayWnd As Long, TrayNotify As Long, TrayClock As Long
    Dim backBrush As Long, textBrush As Long 'Object handles
    STrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
    TrayNotify& = FindWindowEx(STrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
    TrayClock& = FindWindowEx(TrayNotify&, 0&, "TrayClockWClass", vbNullString)
    Call GetWindowRect(TrayClock&, rct)
    iWidth = (rct.Right - rct.Left)
    iHeight = (rct.Bottom - rct.Top)
    rct.Top = 0: rct.Left = 0: rct.Right = iWidth: rct.Bottom = iHeight
    Call FillRect(GetDC(TrayClock&), rct, CreateSolidBrush(BackC))
    Call TextOut(GetDC(TrayClock&), 2, 2, GetText(TrayClock&), Len(GetText(TrayClock&)))
End Sub

Public Sub Mp3_Play(FileName As String, Alias As String)
    'Want more MP3 functions? Get PunkMp3.OCX
    'http://punkdude.cjb.net
    Dim sFile As String
    sFile$ = ShortFileName(FileName$)
    If sFile$ = "" Then Exit Sub
    Call MciSendString("close " & Alias$, 0&, 0&, 0&)
    Call MciSendString("open " & sFile$ & " type MPEGVideo alias " & Alias$, 0&, 0&, 0&)
    Call MciSendString("play " & Alias$, 0&, 0&, 0&)
End Sub

Public Sub Mp3_Pause(Alias As String)
    Call MciSendString("pause " & Alias$, 0&, 0&, 0&)
End Sub

Public Sub Mp3_Resume(Alias As String)
    Call MciSendString("resume " & Alias$, 0&, 0&, 0&)
End Sub

Public Sub Mp3_Stop(Alias As String)
    Call MciSendString("stop " & Alias$, 0&, 0&, 0&)
End Sub

Public Sub Mp3_Close(Alias As String)
    Call MciSendString("close " & Alias$, 0&, 0&, 0&)
End Sub

Public Sub WriteTextFile(FileName As String, tText As String)
    'Creates a file containing the given text
    On Error Resume Next
    Dim X As Long
    Dim sFile As String
    sFile$ = ShortFileName(FileName$)
    If sFile$ = "" Then Exit Sub
    X = FreeFile
    Open sFile$ For Output As #X
        Print #X, tText$
    Close X
End Sub

Public Sub UpChat(bOn As Boolean)
    'Allows users to modify windows within AOL while uploading
    Call EnableWindow(AOLWin&, ByVal CLng(bOn))
    Call EnableWindow(FindUploadWindow&, ByVal CLng(Not bOn))
End Sub

Public Sub ListKillDupe(Lst As ListBox)
    'Removes duplicate strings from a listbox
    Dim i As Integer, X As Integer
    For i = 0 To (Lst.ListCount - 1)
        For X = (i + 1) To (Lst.ListCount - 1)
            If Lst.List(i) = Lst.List(X) Then Lst.RemoveItem (X): X = (X - 1)
        Next X
    Next i
End Sub

Public Function ListToString(Lst As ListBox, Optional Delimeter As String = " ") As String
    'Creates a string containing all string from a list
    'Each string is seperated by the Delimeter
    Dim Txt As String
    Dim i As Integer
    Txt$ = ""
    For i = 0 To (Lst.ListCount - 1)
        Txt$ = (Txt$ & Delimeter$ & Lst.List(i))
    Next i
    ListToString$ = Txt$
End Function

Public Sub ListFromString(Lst As ListBox, Txt As String, Optional Delimeter As String = " ")
    'Adds all portions of a string seperated by the Delimeter to a listbox
    Dim Arry
    Dim i As Integer
    Arry = Split(Txt$, Delimeter$)
    For i = LBound(Arry) To UBound(Arry)
        Lst.AddItem (Arry(i))
    Next i
End Sub

Public Function ListSearch(Lst As ListBox, Txt As String) As Integer
    'Traverses (loops through a list) until the string (Txt) is found
    'Returns the index of the string
    Dim i As Integer
    For i = 0 To (Lst.ListCount - 1)
        If InStr(LCase$(Lst.List(i)), LCase$(Txt$)) > 0& Then ListSearch = i
    Next i
End Function

Public Function RoomIsPrivate() As Boolean
    'Returns whether or not the room is private; private=true, public=false
    If FindRoom& = 0& Then RoomIsPrivate = False: Exit Function
    Dim i As Integer
    Dim AOLIcon As Long
    AOLIcon& = 0
    For i = 1 To 13
        AOLIcon& = FindWindowEx(FindRoom&, AOLIcon&, "_AOL_Image", vbNullString)
    Next i
    ClickIcon (AOLIcon&)
    RoomIsPrivate = Not CBool(IsWindowVisible(AOLIcon&))
End Function

Public Function UploadStatus() As Integer
    'Returns the upload status from AOL's upload window
    If FindUploadWindow& = 0& Then Exit Function
    Dim Txt As String
    Txt$ = GetCaption(FindUploadWindow&)
    UploadStatus = Val(Mid$(Txt$, InStr(Txt$, "File Transfer - "), InStr(Txt$, "%") - 1))
End Function

Public Function ExtractText(Txt As String) As String
    'Returns the alpha (non-numeric) characters from a string
    Dim sTxt As String, oTxt As String
    Dim i As Integer
    sTxt$ = ""
    For i = 1 To Len(Txt$)
        oTxt$ = Mid$(Txt$, i, 1)
        If IsNumeric(oTxt$) = False Then sTxt$ = (sTxt$ & oTxt$)
    Next i
    ExtractText$ = sTxt$
End Function

Public Function ExtractNumeric(Txt As String) As String
    'Returns the numeric characters from a string
    Dim sTxt As String, oTxt As String
    Dim i As Integer
    sTxt$ = ""
    For i = 1 To Len(Txt$)
        oTxt$ = Mid$(Txt$, i, 1)
        If IsNumeric(oTxt$) = True Then sTxt$ = (sTxt$ & oTxt$)
    Next i
    ExtractNumeric$ = sTxt$
End Function

Public Sub Form_TileImage(Frm As Form, Pic As PictureBox)
    'Tiles an image on a form.
    'If form AutoRedraw property is set to true,
    'the tile will be perminant (until next change)
    On Error Resume Next
    Dim a As Integer, d As Integer
    For a = 0 To Frm.ScaleWidth Step Pic.Width
        For d = 0 To Frm.ScaleHeight Step Pic.Height
            Frm.PaintPicture Pic.Picture, a, d
        Next d
    Next a
End Sub

Public Sub Form_FadeBlue(Frm As Form)
    On Error Resume Next
    Dim i As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For i = 0 To 255
        Frm.Line (0, i)-(Screen.Width, i - 1), RGB(0, 0, 255 - i), B
    Next i
End Sub

Public Sub Form_FadeRed(Frm As Form)
    On Error Resume Next
    Dim i As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For i = 0 To 255
        Frm.Line (0, i)-(Screen.Width, i - 1), RGB(255 - i, 0, 0), B
    Next i
End Sub

Public Sub Form_FadeGreen(Frm As Form)
    On Error Resume Next
    Dim i As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For i = 0 To 255
        Frm.Line (0, i)-(Screen.Width, i - 1), RGB(0, 255 - i, 0), B
    Next i
End Sub

Public Sub RunMenuByString(Txt As String)
    'Runs an AOL menu containing the given string (Txt)
    'EX: Call RunMenuByString("AOL &Help")
    Dim AOLWindow As Long, AOLMenu As Long, MenuCount As Long
    Dim LookFor As Long, SubMenu As Long, SubCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLMenu& = GetMenu(AOLWindow&)
    If AOLMenu& = 0& Then Exit Sub
    MenuCount& = GetMenuItemCount(AOLMenu&)
    For LookFor& = 0& To MenuCount& - 1
        SubMenu& = GetSubMenu(AOLMenu&, LookFor&)
        SubCount& = GetMenuItemCount(SubMenu&)
        For LookSub& = 0 To SubCount& - 1
            sID& = GetMenuItemID(SubMenu&, LookSub&)
            sString$ = Space$(100)
            Call GetMenuString(SubMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase$(sString$), LCase$(Txt$)) > 0& Then
                Call SendMessageLong(AOLWindow&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Sub KillWait()
    'Gets rid of AOL's hourglass cursor
    If AOLWin& = 0& Then Exit Sub
    Dim AOLModal As Long, AOLIcon As Long
    RunMenuByString ("&About America Online")
    Do: DoEvents
        AOLModal& = FindWindow("_AOL_Modal", vbNullString)
        AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until (AOLIcon <> 0&)
    ClickIcon (AOLIcon&)
End Sub

Public Sub Combo_AddFonts(Comb As ComboBox)
    'Adds installed system fonts to a combobox
    Dim i As Integer
    For i = 1 To Screen.FontCount
        Comb.AddItem (Screen.Fonts(i))
    Next i
End Sub

Public Sub ListAddFonts(Lst As ListBox)
    'Adds installed system fonts to a listbox
    Dim i As Integer
    For i = 1 To Screen.FontCount
        Lst.AddItem (Screen.Fonts(i))
    Next i
End Sub

Public Sub CloseActiveWin()
    'Closes the active window
    Call SendMessage(GetActiveWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ScrollText(tBox As TextBox, Txt As String, Optional Delay = 0.8)
    'Scrolls the text (Txt) in a textbox (tBox)
    Dim i As Integer
    tBox.Text = ""
    For i = 1 To Len(Txt$)
        DoEvents
        tBox.Text = Left$(Txt$, i)
        TimeOut (Delay)
        If i = Len(Txt$) Then i = 0: tBox.Text = ""
    Next i
End Sub

Public Sub ScrollCaption(Win As Long, Txt As String, Optional Delay = 0.8)
    'Scrolls the caption (Txt) of a window (Win)
    Dim i As Integer
    Call SetText(Win&, "")
    For i = 1 To Len(Txt$)
        DoEvents
        Call SetText(Win&, Left$(Txt$, i))
        TimeOut (Delay)
        If i = Len(Txt$) Then i = 0: Call SetText(Win&, "")
    Next i
End Sub

Public Sub Array_Save(Arr, FileName As String)
    'Saves an array to a file
    On Error Resume Next
    Dim X As Long
    Dim i As Integer
    Dim Txt As String
    Txt$ = ""
    X = FreeFile
    For i = LBound(Arr) To UBound(Arr)
        Txt$ = (Txt$ & Arr(i))
    Next i
    Open ShortFileName$(FileName$) For Output As #X
        For i = LBound(Arr) To UBound(Arr): DoEvents
            Print #X, CStr(Arr(i))
        Next i
    Close X
End Sub

Public Sub Array_Load(Arr, FileName As String)
    'Loads a file into an array
    On Error Resume Next
    Dim i As Integer
    Dim X As Long
    X = FreeFile
    i = 0
    Open ShortFileName$(FileName$) For Input As #X
        While Not EOF(X): DoEvents
            i = i + 1
            Input #X, Arr(i)
        Wend
    Close X
End Sub

Public Sub Array_ToList(Arr, Lst As ListBox)
    'Fills a listbox with the contents of an array
    Dim i As Integer
    For i = LBound(Arr) To UBound(Arr)
        Lst.AddItem (CStr(Arr(i)))
    Next i
End Sub

Public Sub Array_FromList(Arr, Lst As ListBox)
    'Fills an array with the contents of a listbox
    Dim i As Integer
    ReDim Arr(Lst.ListCount - 1)
    For i = 0 To (Lst.ListCount - 1)
        Arr(i) = Lst.List(i)
    Next i
End Sub

Public Sub SetDesktopWallPaper(Optional FileName As String)
    'If FileName parameter is missing, the wallpaper will be set to default
    Dim sFile As String
    If CBool(IsMissing(FileName)) = False Then
        sFile$ = ShortFileName$(FileName$)
        If sFile$ = "" Then Exit Sub
        Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFile$, 0&)
    Else
        Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, Null, 0&)
    End If
End Sub

Public Sub AIM_GetAd(Pic As PictureBox, TopOrBottom As AIM_AD)
    'Puts AIM's ad image in a picturebox
    Dim rct As RECT
    Dim hWnd1 As Long, hWnd2 As Long, hWnd3 As Long
    Dim bMax As Boolean
    hWnd1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    hWnd2& = FindWindowEx(hWnd1&, 0&, "WndAte32Class", vbNullString)
    If TopOrBottom = topAd Then
        hWnd2& = FindWindowEx(hWnd1&, 0&, "WndAte32Class", vbNullString)
    ElseIf TopOrBottom = BottomAd Then
        hWnd2& = FindWindowEx(hWnd1&, 0&, "WndAte32Class", vbNullString)
        hWnd2& = FindWindowEx(hWnd1&, hWnd2&, "WndAte32Class", vbNullString)
    End If
    hWnd3& = FindWindowEx(hWnd2&, 0&, "Ate32Class", vbNullString)
    bMax = IsWindowVisible(hWnd1&)
    Do Until CBool(IsWindowVisible(hWnd1&)) = True: DoEvents
        Call ShowWindow(hWnd1&, SW_SHOW)
    Loop
    Call GetWindowRect(hWnd3&, rct)
    DoEvents
    Pic.ScaleMode = 3: Pic.AutoRedraw = True
    DoEvents
    Call BitBlt(Pic.hDC, 0, 0, (rct.Right - rct.Left), (rct.Bottom - rct.Top), GetDC(hWnd3&), 0, 0, SRCCOPY)
    DoEvents
    Call Pic.PaintPicture(Pic.Image, Pic.Height, Pic.Width)
    If bMax = False Then Call ShowWindow(hWnd1&, SW_HIDE)
End Sub

Public Sub GetGlyph(Pic As PictureBox)
    'Put AOL's glyph image in a picturebox
    Dim rct As RECT
    Dim X As Integer
    Dim hWnd1 As Long, hWnd2 As Long, hWnd3 As Long
    Dim bMax As Boolean
    hWnd1& = FindWindowEx(AOLWin&, 0&, "AOL Toolbar", vbNullString)
    hWnd2& = FindWindowEx(hWnd1&, 0&, "_AOL_Toolbar", vbNullString)
    hWnd3& = FindWindowEx(hWnd2&, 0&, "_AOL_Glyph", vbNullString)
    X = Window_GetState(AOLWin&)
    Call Window_SetState(AOLWin&, Maximized)
    Call GetWindowRect(hWnd3&, rct)
    DoEvents
    Pic.ScaleMode = 3: Pic.AutoRedraw = True
    DoEvents
    Call BitBlt(Pic.hDC, 0, 0, (rct.Right - rct.Left), (rct.Bottom - rct.Top), GetDC(hWnd3&), 0, 0, SRCCOPY)
    DoEvents
    Call Pic.PaintPicture(Pic.Image, Pic.Height, Pic.Width): DoEvents
    Call Window_SetState(AOLWin&, Window_IntToState(X))
End Sub

Public Sub AIM_ImSNtoList(Lst As ListBox)
    'Adds the screen name from all open AIM IM's to a listbox
    Dim IMWin As Long
    Do: DoEvents
        IMWin& = FindWindowEx(0&, IMWin&, "AIM_IMessage", vbNullString)
        If IMWin& = 0& Then Exit Sub
        Lst.AddItem (Trim$(Left$(GetCaption(IMWin&), InStr(GetCaption(IMWin&), " - Instant Message"))))
    Loop Until IMWin& = 0&
End Sub

Public Sub AIM_RunMenuByString(Txt As String)
    'Runs an AIM menu containing the given string (Txt)
    'EX: Call AIM_RunMenuByString("Edit &Profile...")
    Dim AIMMenu As Long, MenuCount As Long
    Dim LookFor As Long, SubMenu As Long, SubCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AIMMenu& = GetMenu(AIM_Win&)
    If AIMMenu& = 0& Then Exit Sub
    MenuCount& = GetMenuItemCount(AIMMenu&)
    For LookFor& = 0& To (MenuCount& - 1)
        SubMenu& = GetSubMenu(AIMMenu&, LookFor&)
        SubCount& = GetMenuItemCount(SubMenu&)
        For LookSub& = 0 To SubCount& - 1
            sID& = GetMenuItemID(SubMenu&, LookSub&)
            sString$ = Space$(100)
            Call GetMenuString(SubMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase$(sString$), LCase$(Txt$)) > 0& Then
                Call SendMessageLong(AIM_Win&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Sub SwitchValues(Value1, Value2)
    'Returns Value1 containing the value of Value2 & Value2 containing the value of Value2
    Dim tVal
    tVal = Value1
    Value1 = Value2
    Value2 = tVal
End Sub

Public Sub SetAOLChild(Frm As Form)
    'Set a form as an AOL MDI child
    Call SetParent(Frm.hwnd, AOLMDI&)
End Sub

Public Sub ClickList(Lst As Long)
    'Clicks the selected list item
    Call PostMessage(Lst&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub

Public Function Text_Oo(Txt As String) As String
    'Replaces O with Ø and o with ø
    Dim tTxt As String
    tTxt$ = ReplaceString("O", Txt$, "Ø")
    Text_Oo$ = ReplaceString("o", tTxt$, "ø")
End Function

Public Function Text_uFirst(Txt As String) As String
    'Converts the first character after every space to upper case
    Text_uFirst$ = StrConv(Txt$, vbProperCase)
End Function

Public Function DecodeSN(SN As String) As String
    'Replaces I with i and l with L
    Dim tTxt As String
    tTxt$ = ReplaceString("I", SN$, "i")
    DecodeSN$ = ReplaceString("l", tTxt$, "L")
End Function

Public Function DirectoryExists(TheDirectory As String) As Boolean
    'Checks if a directory exists
    Dim Check As Integer
    On Error Resume Next
    If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
    Check = Len(Dir$(TheDirectory))
    If (Err) Or (Check = 0) Then
        DirectoryExists = False
    Else
        DirectoryExists = True
    End If
End Function

Public Sub OpenDirectory(Dir As String)
    'Opens a given directory
    Call Shell(Dir$)
End Sub

Public Sub SignOff()
    'Signs off AOL by running the Exit menu
    Call RunMenuByString("E&xit")
End Sub

Public Sub FTPSpace(SN As String)
    'View the AOL FTP space of an AOL member
    'EX:  Call FTPSpace("SteveCase")  'Works too :)
    Call Keyword("aol://5862:144/members.aol.com:/" + Trim$(SN$))
End Sub

Public Sub FakeRoom(RoomName As String)
    'Makes some chaters think they entered a new room
    Call ChatSend(" ")
    Call ChatSend("<font face=Arial color=#000000>" & "*** You are in " & Chr$(34) & RoomName$ & Chr$(34) & ". ***")
    Call ChatSend(" ")
End Sub

Public Sub CloseCurrentRoom()
    'Closes the currently open AOL chatroom
    Call Window_Close(FindRoom&)
End Sub

Public Sub LinkColor(HexValue As Long)
    'Sends an HTML tag to an AOL chatroom which will change the
    'color of the forthcoming link text
    Call ChatSend("<body link=#" & CStr(HexValue) & "><html></html>")
End Sub

Public Sub AddSystemMenu(SysMenu As Long, Pos As Long, NewMenu As Long, newcaption As String)
    'Adds a menu to a programs system menu
    Call InsertMenu(SysMenu&, Pos&, MF_BYPOSITION, NewMenu&, newcaption$)
End Sub

Public Function Window_GetState(Win As Long) As WINDOW_STATE
    'Returns if a window is maximized, minimized, or normal (restored)
    Dim gws As Long
    gws = GetWindowLong(Win&, GWL_STYLE)
    If (gws And WS_MAXIMIZE) = WS_MAXIMIZE Then
        Window_GetState = Maximized
    ElseIf (gws And WS_MINIMIZE) = WS_MINIMIZE Then
        Window_GetState = Minimized
    Else
        Window_GetState = Normal
    End If
End Function

Public Sub Window_SetState(Win As Long, State As WINDOW_STATE)
    'Sets a window to maximized, minimized, or normal (restored)
    Select Case Int(State)
        Case 0: Window_Maximize (Win&)
        Case 1: Window_Minimize (Win&)
        Case 2: Window_Restore (Win&)
    End Select
End Sub

Public Function Window_IntToState(Num As Integer) As WINDOW_STATE
    'Changes an integer value to a WINDOW_STATE constant
    Select Case Int(Num)
        Case 0: Window_IntToState = Maximized
        Case 1: Window_IntToState = Minimized
        Case 2: Window_IntToState = Normal
    End Select
End Function

Public Sub MailBoxOpen()
    'Opens the AOL mailbox by clicking the Read icon on the toolbar
    Dim tBar As Long, tBar2 As Long, Icn As Long
    tBar& = FindWindowEx(AOLWin&, 0&, "AOL Toolbar", vbNullString)
    tBar2& = FindWindowEx(tBar&, 0&, "_AOL_Toolbar", vbNullString)
    Icn& = FindWindowEx(tBar2&, 0&, "_AOL_Icon", vbNullString)
    Icn& = FindWindowEx(tBar2&, Icn&, "_AOL_Icon", vbNullString)
    ClickIcon (Icn&)
End Sub
Public Sub MailOpenIndex(index As Long)
    'Opens the mailbox then opens the mail of a certain index
    Dim TabControl As Long, TabPage As Long, Tree As Long
    MailBoxOpen
    Do Until FindMailBox& <> 0&: DoEvents: Loop
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    Tree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    If index > SendMessage(Tree&, LB_GETCOUNT, 0&, 0&) Then Exit Sub
    Call SendMessage(Tree&, LB_SETCURSEL, ByVal CLng(index), 0&)
    DoEvents
    Call ClickList(Tree&)
End Sub

Public Sub Ghost(GhostOn As Boolean)
    'This will make it look like your not online, even if you are.
    'NOTE: Your screen name will still show up in chat rooms.
    Dim i As Integer
    Dim PrivIcon As Long, Pref As Long, PrefTab As Long, PrefPage As Long
    Dim PrefRadio As Long, PrefIcon As Long
    Keyword ("buddy")
    Do Until FindBuddySetup& <> 0&: DoEvents: Loop
    i = 0
    For i = 1 To 6
        PrivIcon& = FindWindowEx(FindBuddySetup&, PrivIcon&, "_AOL_Icon", vbNullString)
    Next i
    Do: DoEvents
        ClickIcon (PrivIcon&): TimeOut (0.4): Pref& = FindBuddyPref& 'Waits for Preferences window
    Loop Until Pref& <> 0&
    PrefTab& = FindWindowEx(Pref&, 0&, "_AOL_TabControl", vbNullString)
    i = 0
    For i = 1 To 2
        Call SendMessageLong(PrefTab&, WM_KEYDOWN, VK_RIGHT, 0&)
        Call SendMessageLong(PrefTab&, WM_KEYUP, VK_RIGHT, 0&)
        TimeOut (0.3)
    Next i
    DoEvents
    PrefPage& = FindWindowEx(PrefTab&, 0&, "_AOL_TabPage", vbNullString)
    PrefPage& = FindWindowEx(PrefTab&, PrefPage&, "_AOL_TabPage", vbNullString)
    PrefPage& = FindWindowEx(PrefTab&, PrefPage&, "_AOL_TabPage", vbNullString)
    Do: DoEvents
        PrefRadio& = FindWindowEx(PrefPage&, 0&, "_AOL_RadioBox", vbNullString)
    Loop Until PrefRadio& <> 0&
    If GhostOn = True Then
        i = 0
        For i = 1 To 4
            PrefRadio& = FindWindowEx(PrefPage&, PrefRadio&, "_AOL_RadioBox", vbNullString)
        Next i
        ClickRadio (PrefRadio&)
        i = 0
        For i = 1 To 2
            PrefRadio& = FindWindowEx(PrefPage&, PrefRadio&, "_AOL_RadioBox", vbNullString)
        Next i
        ClickRadio (PrefRadio&)
        i = 0
        For i = 1 To 3
            PrefIcon& = FindWindowEx(PrefPage&, PrefIcon&, "_AOL_Icon", vbNullString)
        Next i
    Else
        ClickRadio (PrefRadio&)
    End If
        ClickIcon (FindWindowEx(Pref&, 0&, "_AOL_Icon", vbNullString))
End Sub

Public Sub CopyImage(Win As Long, Pic As PictureBox)
    'Copies an Image from a control on a window into a picturebox
    Dim rct As RECT
    Dim bMax As Boolean
    bMax = IsWindowVisible(Win&)
    Do Until CBool(IsWindowVisible(Win&)) = True: DoEvents
        Call ShowWindow(Win&, SW_SHOW)
    Loop
    Call GetWindowRect(Win&, rct)
    DoEvents
    Pic.ScaleMode = 3: Pic.AutoRedraw = True
    DoEvents
    Call BitBlt(Pic.hDC, 0, 0, (rct.Right - rct.Left), (rct.Bottom - rct.Top), GetDC(Win&), 0, 0, SRCCOPY)
    DoEvents
    Call Pic.PaintPicture(Pic.Image, Pic.Height, Pic.Width)
    If bMax = False Then Call ShowWindow(Win&, SW_HIDE)
End Sub

Public Sub SetChatPref(cOption As CHAT_PREF, cChecked As Boolean)
    'Sets an AOL Chat Preferences option to Checked or Unchecked
    'EX:  Call SetChatPref(Notify_Arrive, True)
    If FindRoom& = 0& Then Exit Sub
    Dim i As Integer, X As Integer
    Dim AOLIcon As Long, AOLmPref As Long, AOLmOp As Long, AOLmIcon As Long
    AOLIcon& = 0&
    For i = 1 To 13
        AOLIcon& = FindWindowEx(FindRoom&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i
    ClickIcon (AOLIcon&)
    Do: DoEvents
        AOLmPref& = FindWindow("_AOL_Modal", "Chat Preferences")
    Loop Until AOLmPref& <> 0&
    For X = 1 To cOption: DoEvents
        AOLmOp& = FindWindowEx(AOLmPref&, AOLmOp&, "_AOL_Checkbox", vbNullString)
    Next X
    If Not (SendMessage(AOLmOp&, BM_GETCHECK, 0&, 0&) = cChecked) Then
        Call SendMessage(AOLmOp&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(AOLmOp&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents: Loop Until SendMessage(AOLmOp&, BM_GETCHECK, 0&, 0&) = cChecked
    AOLmIcon& = FindWindowEx(AOLmPref&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (AOLmIcon&)
End Sub

Public Function MaskToolTip(Link As String) As String
    'If sent with a hyperlink on AOL, makes the tooltip description say,
    ' "On AOL Only"
    MaskToolTip$ = ("aol://1223:0/" & Link$)
End Function

Public Function EqualString(Str1 As String, Str2 As String) As Boolean
    'Returns True or False depending on whether the converted strings are equal
    Dim a As String, b As String
    a$ = Trim$(LCase$(RemoveSpaces(Str1$))): b$ = Trim$(LCase$(RemoveSpaces(Str2$)))
    If a$ = b$ Then EqualString = True Else EqualString = False
End Function

Public Function ShortString(sString As String) As String
    'Trims, lowercases, and removes spaces from a string
    ShortString$ = Trim$(LCase$(RemoveSpaces(sString$)))
End Function

Public Function GenerateAscii(AsciType As ASCII_TYPE) As String
    'Generates decerative ASCII strings
    'EX:  Txt$ = GenerateAscii(aRight)
    Dim Arr(0 To 4) As String
    If AsciType = aLeft Then
        Arr(0) = "· ·÷(`‹›"
        Arr(1) = "· ··•["
        Arr(2) = "«¬­"
        Arr(3) = "«¬~"
        Arr(4) = "«v^×"
    ElseIf AsciType = aRight Then
        Arr(0) = "‹›´)÷· ·"
        Arr(1) = "]•·· ·"
        Arr(2) = ".´·)v›"
        Arr(3) = "•^v›"
        Arr(4) = "×^v»"
    ElseIf AsciType = aOther Then
        Arr(0) = "‹›"
        Arr(1) = "«›"
        Arr(2) = "«±»"
        Arr(3) = "‹^›"
        Arr(4) = "•¤•"
    End If
End Function

Public Sub WaitForKeyPress(Key As Long)
    'Waits until the specified key stroke
    'EX:  WaitForKeyPress(VK_RETURN)
    Do: DoEvents: Loop Until GetAsyncKeyState(Key) = -32767
End Sub

Public Function WaitForKeyPress2() As Long
    'Waits until any key stroke; Returns the stroked key
    'EX:  MsgBox WaitForKeyPress2
    Dim X As Long
    X = 32
    Do: DoEvents
        X = X + 1
        If X > 255 Then X = 32
    Loop Until GetAsyncKeyState(X) = -32767
    WaitForKeyPress2 = X
End Function

Public Sub ListRemoveSel(Lst As ListBox)
    'Removes the selected list item from a listbox
    If Lst.ListIndex = -1 Then Exit Sub
    Lst.RemoveItem (Lst.ListIndex)
End Sub

Public Sub MailGotoTab(mTab As MAIL_TAB)
    'Opens the mailbox and moves to one of the tabs; New, Old, or Sent
    Dim i As Integer
    Dim MailTab As Long
    MailBoxOpen
    Do: DoEvents: Loop Until FindMailBox& <> 0&
    MailTab& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    For i = 0 To mTab
        Call SendMessageLong(MailTab&, WM_KEYDOWN, VK_RIGHT, 0&)
        Call SendMessageLong(MailTab&, WM_KEYUP, VK_RIGHT, 0&)
        TimeOut (0.3)
    Next i
End Sub

Public Sub MailKeepAllAsNew()
    'Keeps all of the new mail As New
    Dim X As Integer
    Dim MailTab As Long, MailPage As Long, MailTree As Long, MailIcon As Long
    If FindMailBox& = 0& Then Exit Sub
    MailTab& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    MailPage& = FindWindowEx(MailTab&, 0&, "_AOL_TabPage", vbNullString)
    MailTree& = FindWindowEx(MailPage&, 0&, "_AOL_Tree", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, MailIcon&, "_AOL_Icon", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, MailIcon&, "_AOL_Icon", vbNullString)
    For X = 0 To SendMessage(MailTree&, LB_GETCOUNT, 0&, 0&) - 1
        Call SendMessage(MailTree&, LB_SETCURSEL, ByVal CLng(X), 0&)
        Do: DoEvents: Loop Until SendMessage(MailTree&, LB_GETCURSEL, 0&, 0&) = X
        ClickIcon (MailIcon&)
    Next X
End Sub

Public Sub MailDeleteAll(mType As MAIL_TAB)
    'Delete All Mail in a specified tab; New, Old, or Sent
    Dim MailTab As Long, MailPage As Long, MailTree As Long, MailIcon As Long
    Dim i As Integer, X As Integer
    MailBoxOpen
    Do: DoEvents: Loop Until FindMailBox& <> 0&
    MailTab& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    For i = 0 To mType
        Call SendMessageLong(MailTab&, WM_KEYDOWN, VK_RIGHT, 0&)
        Call SendMessageLong(MailTab&, WM_KEYUP, VK_RIGHT, 0&)
        TimeOut (0.3)
    Next i
    MailPage& = FindWindowEx(MailTab&, 0&, "_AOL_TabPage", vbNullString)
    MailTree& = FindWindowEx(MailPage&, 0&, "_AOL_Tree", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, MailIcon&, "_AOL_Icon", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, MailIcon&, "_AOL_Icon", vbNullString)
    MailIcon& = FindWindowEx(FindMailBox&, MailIcon&, "_AOL_Icon", vbNullString)
    For X = 0 To SendMessage(MailTree&, LB_GETCOUNT, 0&, 0&) - 1
        Call SendMessage(MailTree&, LB_SETCURSEL, ByVal CLng(X), 0&)
        Do: DoEvents: Loop Until SendMessage(MailTree&, LB_GETCURSEL, 0&, 0&) = X
        ClickIcon (MailIcon&)
    Next X
End Sub

Public Function GetColorFromCursor() As Long
    'Returns the RGB color of the pixel the cursor is pointing to
    'EX:  MsgBox GetColorFromCursor
    Dim Temp As Long
    Dim Curp As POINTAPI
    Dim ret As String
    Call GetCursorPos(Curp)
    Temp = GetPixel(GetWindowDC(WindowFromPoint(Curp.X, Curp.Y)), Curp.X, Curp.Y)
    If Temp = -1 Then
        Call GetClassName(WindowFromPoint(Curp.X, Curp.Y), ret, 256)
        Temp = GetPixel(GetWindowDC(FindWindowEx(WindowFromPoint(Curp.X, Curp.Y), 0&, ret, vbNullString)), Curp.X, Curp.Y)
    End If
    GetColorFromCursor = Temp
End Function

Public Function GetColorFromPoint(X As Long, Y As Long) As Long
    'Returns the RGB color of the pixel in the position of the given coordinates (X and Y)
    'EX:  MsgBox GetColorFromPoint(100, 100)
    Dim Temp As Long
    Dim ret As String
    Temp = GetPixel(GetWindowDC(WindowFromPoint(X, Y)), X, Y)
    If Temp = -1 Then
        Call GetClassName(WindowFromPoint(X, Y), ret, 256)
        Temp = GetPixel(GetWindowDC(FindWindowEx(WindowFromPoint(X, Y), 0&, ret, vbNullString)), X, Y)
    End If
    GetColorFromPoint = Temp
End Function

Public Function GetColorFromPointAPI(Coords As POINTAPI) As Long
    'Returns the RGB color of the pixel in the position of the given coordinates (Coords)
    'EX:  Dim Curps As POINTAPI
    '     Call GetCursorPos(Curps)
    '     MsgBox GetColorFromPointAPI(Curps)
    Dim Temp As Long
    Dim ret As String
    Temp = GetPixel(GetWindowDC(WindowFromPoint(Coords.X, Coords.Y)), Coords.X, Coords.Y)
    If Temp = -1 Then
        Call GetClassName(WindowFromPoint(Coords.X, Coords.Y), ret, 256)
        Temp = GetPixel(GetWindowDC(FindWindowEx(WindowFromPoint(Coords.X, Coords.Y), 0&, ret, vbNullString)), Coords.X, Coords.Y)
    End If
    GetColorFromPointAPI = Temp
End Function

Public Function MailCount(mType As MAIL_TAB) As Integer
    'Returns the mail count of a given mail tab; New, Old, or Sent
    Dim MailTab As Long, MailPage As Long, MailTree As Long
    Dim i As Integer, X As Integer
    MailBoxOpen
    Do: DoEvents: Loop Until FindMailBox& <> 0&
    MailTab& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    For i = 0 To mType
        Call SendMessageLong(MailTab&, WM_KEYDOWN, VK_RIGHT, 0&)
        Call SendMessageLong(MailTab&, WM_KEYUP, VK_RIGHT, 0&)
        TimeOut (0.3)
    Next i
    MailPage& = FindWindowEx(MailTab&, 0&, "_AOL_TabPage", vbNullString)
    MailTree& = FindWindowEx(MailPage&, 0&, "_AOL_Tree", vbNullString)
    MailCount = SendMessage(MailTree&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function TranslateTerm(Lang As TERM_LANG, Term As String) As String
    'Translate popular AOL/Programming terms into German, Spanish, & Japanese
    If Lang = lGerman Then
        Select Case LCase$(Term$)
            Case "elite"
                TranslateTerm$ = "die Elite"
            Case "chat"
                TranslateTerm$ = "der Schwatz"
            Case "room"
                TranslateTerm$ = "das Zimmer"
            Case "instant message"
                TranslateTerm$ = "der Moment Assage"
            Case "mail"
                TranslateTerm$ = "die Post"
            Case "text"
                TranslateTerm$ = "der Text"
            Case "bot"
                TranslateTerm$ = "der Automat"
            Case "programmer"
                TranslateTerm$ = "der Programmierer"
            Case "disclaimer"
                TranslateTerm$ = "das Dementi"
            Case "punt"
                TranslateTerm$ = "staken"
            Case "ignore"
                TranslateTerm$ = "ignorieren"
            Case "hack"
                TranslateTerm$ = "die Hacke"
            Case "screen"
                TranslateTerm$ = "das Raster"
            Case "name"
                TranslateTerm$ = "der Ruf"
            Case "hide"
                TranslateTerm$ = "wegtun"
            Case "show"
                TranslateTerm$ = "der Shein"
        End Select
    ElseIf Lang = lSpanish Then
        Select Case LCase$(Term$)
            Case "elite"
                TranslateTerm$ = "elite"
            Case "chat"
                TranslateTerm$ = "charlar"
            Case "room"
                TranslateTerm$ = "espacio"
            Case "instant message"
                TranslateTerm$ = "momento mensaje"
            Case "mail"
                TranslateTerm$ = "correo"
            Case "text"
                TranslateTerm$ = "texto"
            Case "bot"
                TranslateTerm$ = "robot"
            Case "programmer"
                TranslateTerm$ = "programador"
            Case "disclaimer"
                TranslateTerm$ = "limitación de responsabilidad"
            Case "punt"
                TranslateTerm$ = "batea"
            Case "ignore"
                TranslateTerm$ = "desentenderse de"
            Case "hack"
                TranslateTerm$ = "háck"
            Case "screen"
                TranslateTerm$ = "pantalla"
            Case "name"
                TranslateTerm$ = "nombre"
            Case "hide"
                TranslateTerm$ = "piel"
            Case "show"
                TranslateTerm$ = "designar"
        End Select
    ElseIf Lang = lJapanese Then
        Select Case LCase$(Term$)
            Case "elite"
                TranslateTerm$ = "eri-to"
            Case "chat"
                TranslateTerm$ = "chatto"
            Case "room"
                TranslateTerm$ = "ruumu"
            Case "instant message"
                TranslateTerm$ = "setsuna shourei"
            Case "mail"
                TranslateTerm$ = "meiru"
            Case "text"
                TranslateTerm$ = "bun"
            Case "bot"
                TranslateTerm$ = "robotto"
            Case "programmer"
                TranslateTerm$ = "purogurama"
            Case "disclaimer"
                TranslateTerm$ = "mensekijoukou"
            Case "punt"
                TranslateTerm$ = "panto"
            Case "ignore"
                TranslateTerm$ = "keishi"
            Case "hack"
                TranslateTerm$ = "hakku"
            Case "screen"
                TranslateTerm$ = "tsuitate"
            Case "name"
                TranslateTerm$ = "mei"
            Case "hide"
                TranslateTerm$ = "toku"
            Case "show"
                TranslateTerm$ = "tosho"
        End Select
    End If
End Function

Public Sub CD_Tray(OpenClose As TRAY_STATE)
    'Opens/Closes the cd tray
    If OpenClose = Tray_Open Then
        Call MciSendString("set cd door open", 0, 0&, 0&)
    Else
        Call MciSendString("set cd door closed", 0, 0&, 0&)
    End If
End Sub

Public Function CD_TrackCount() As Long
    'Returns the number of tracks from the cd media (if any)
    Dim sRet As String
    sRet$ = String$(50, " ")
    Call MciSendString("status cd number of tracks wait", sRet, Len(sRet), 0&)
    CD_TrackCount = CLng(Trim(sRet))
End Function

Public Sub CD_Play(Optional StartTrack = 0, Optional LastTrack = 0)
    'Plays a cd from the given starting track to the to the given ending track
    '(If neither value if given, the entire cd is played)
    If (StartTrack = 0) And (LastTrack = 0) Then
        Call MciSendString("play cd", 0, 0&, 0&)
    ElseIf (StartTrack <> 0) And (LastTrack = 0) Then
        Call MciSendString("play cd from " & CStr(StartTrack) & " to " & CStr(CD_TrackCount), 0, 0&, 0&)
    Else
        Call MciSendString("play cd from " & CStr(StartTrack) & " to " & CStr(LastTrack), 0, 0&, 0&)
    End If
End Sub

Public Sub CD_Stop()
    'Stops a cd's current track (if any)
    Call MciSendString("stop cd wait", 0, 0&, 0&)
End Sub

Public Sub CD_Pause(PauseResume As CD_PAUSERESUME)
    'Pauses/Resumes a cd's current track (if any)
    If PauseResume = cdPause Then
        Call MciSendString("pause cd", 0, 0&, 0&)
    Else
        Call MciSendString("resume cd", 0, 0&, 0&)
    End If
End Sub

Public Function RGBtoHEX(cRGB) As String
    'Converts an RGB color value to a Hex (HTML) color value
    'EX:  ChatSend ("<FONT COLOR=#" & RGBtoHEX(vbBlack) & ">PuNkDuDe")
    Dim a, b
    a = Hex(cRGB)
    b = Len(a)
    RGBtoHEX$ = (String$(6 - b, "0") & CStr(a))
End Function

Public Sub ListAddAscii(Lst As ListBox)
    'Adds all ASCII characters to a listbox
    Dim i As Integer
    For i = 33 To 255
        Lst.AddItem (Chr$(i) & vbNullChar)
    Next i
End Sub

Public Sub ListKillBlank(Lst As ListBox)
    'Removes all blank lines from a listbox
    Dim i As Integer
    For i = 0 To (Lst.ListCount - 1)
        If Lst.List(i) = "" Then Lst.RemoveItem (i)
    Next i
End Sub

Public Sub ScrollCredits(Win As Form, Lbl As Label, Optional interval = 0.6)
    'Scrolls text on a form, just like at the movies
    Dim a(1 To 3) As String 'Change 3 to number of people
    Dim X As Integer
    Dim textX As Long
    Win.ScaleMode = vbPixels
    a(1) = "punkdude": a(2) = "progee": a(3) = "assmonkey" 'Change names
    Lbl.Top = Win.Height
    For X = LBound(a) To UBound(a)
        Lbl.Caption = Lbl.Caption & vbCrLf & a(X)
    Next X
    textX = Win.ScaleHeight
    Lbl.Top = textX
    Lbl.Left = (Win.ScaleWidth / 2) - (Lbl.Width / 2)
    Do Until (Lbl.Top + Lbl.Height) <= 0: DoEvents
        textX = textX - 15
        Lbl.Top = textX
        Lbl.Left = (Win.ScaleWidth / 2) - (Lbl.Width / 2)
        TimeOut (interval)
        If (Lbl.Top + Lbl.Height) <= 0 Then Exit Sub
    Loop
End Sub

Public Sub Form_CenterControl(Parent As Form, Cntrl As Object, oPosition As CENTER_TYPE)
    'Centers a control on a form; Verticaly, Horizontaly, or both
    If oPosition = oHorizontal Then
        Cntrl.Left = (Parent.ScaleWidth / 2) - (Cntrl.Width / 2)
    ElseIf oPosition = oVertical Then
        Cntrl.Top = (Parent.ScaleHeight / 2) - (Cntrl.Height / 2)
    Else
        Cntrl.Left = (Parent.ScaleWidth / 2) - (Cntrl.Width / 2)
        Cntrl.Top = (Parent.ScaleHeight / 2) - (Cntrl.Height / 2)
    End If
End Sub

Public Function AOLMenuKeyDatabase(MenuString As String) As String
    Dim ret As String
    Select Case LCase$(Trim$(MenuString$))
        Case "read mail" 'Mail menu
            ret = "mr"
        Case "write mail"
            ret = "mw"
        Case "address book"
            ret = "ma"
        Case "mail center"
            ret = "mm"
        Case "recently deleted mail"
            ret = "md"
        Case "filing cabinet"
            ret = "mf"
        Case "mail waiting to be sent"
            ret = "mb"
        Case "automatic aol"
            ret = "mu"
        Case "mail signatures"
            ret = "ms"
        Case "mail controls"
            ret = "mc"
        Case "mail preferences"
            ret = "mp"
        Case "greetings & mail extras"
            ret = "mg"
        Case "newsletters"
            ret = "mn"
        Case "send instant message" 'People menu
            ret = "pi"
        Case "chat (people connection)"
            ret = "pc"
        Case "chat now"
            ret = "pn"
        Case "find a chat"
            ret = "pf"
        Case "start your own chat"
            ret = "ps"
        Case "live events"
            ret = "pv"
        Case "buddy list"
            ret = "pb"
        Case "get directory listing"
            ret = "pg"
        Case "locate member online"
            ret = "pl"
        Case "send message to pager"
            ret = "pm"
        Case "sign on a friend"
            ret = "po"
        Case "aol hometown"
            ret = "ph"
        Case "groups@aol"
            ret = "pa"
        Case "invitations"
            ret = "pv"
        Case "people directory"
            ret = "pd"
        Case "personals"
            ret = "pp"
        Case "white pages"
            ret = "pw"
        Case "yellow pages"
            ret = "py"
        Case "shop@aol" 'AOL Services menu
            ret = "as"
        Case "internet"
            ret = "ai"
        Case "add to my calendar"
            ret = "aa"
        Case "aol help"
            ret = "ah"
        Case "calendar"
            ret = "ac"
        Case "car buying"
            ret = "ab"
        Case "download center"
            ret = "ad"
        Case "government guide"
            ret = "au"
        Case "homework help"
            ret = "ak"
        Case "maps & directions"
            ret = "am"
        Case "medical references"
            ret = "ar"
        Case "member rewards"
            ret = "ae"
        Case "movie showtimes"
            ret = "aw"
        Case "online greetings"
            ret = "ag"
        Case "personals"
            ret = "ap"
        Case "recipe finder"
            ret = "af"
        Case "sports scores"
            ret = "ao"
        Case "stock portofilos"
            ret = "al"
        Case "stock quotes"
            ret = "aq"
        Case "travel reservations"
            ret = "av"
        Case "tv listings"
            ret = "at"
        Case "aol anywhere" 'Settings menu
            ret = "sA"
        Case "preferences"
            ret = "sp"
        Case "parental controls"
            ret = "sc"
        Case "my directory listing"
            ret = "sm"
        Case "screen names"
            ret = "ss"
        Case "passwords"
            ret = "sa"
        Case "biling center"
            ret = "sb"
        Case "aol quick checkout"
            ret = "sq"
        Case "favorite places" 'Favorites menu
            ret = "vf"
        Case "add top window to favorites"
            ret = "va"
        Case "go to keyword..."
            ret = "vg"
    End Select
    AOLMenuKeyDatabase$ = ret$
End Function

Public Sub AOLRunMenuByString(sString As String)
    'NOT FINISHED
    Dim ret As String, a As String, b As String
    ret$ = AOLMenuKeyDatabase(LCase$(Trim$(sString$)))
    a$ = Left$(ret$, 1): b$ = Right$(ret$, 1)
    
    Call SendMessage(AOLWin&, WM_KEYDOWN, VK_MENU, 0&)
    
    Call SendMessage(AOLWin&, WM_KEYUP, VK_MENU, 0&)
End Sub

Public Sub RunMenuByString6(Txt As String)
    'AOL 6.0
    'NOT FINISHED
    If AOLVersion2 <> 6 Then Exit Sub
    Dim i As Integer
    Dim AOLWindow As Long, AOLTool1 As Long, AOLTool2 As Long, AOLIcon As Long
    Dim AOLMenuHandle As Long, AOLMenu As Long, MenuCount As Long
    Dim Found As Long
    Dim LookFor As Long, SubMenu As Long, SubCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOLWindow& = FindWindow("AOL Frame25", vbNullString)
    AOLTool1& = FindWindowEx(AOLWindow&, 0&, "AOL Toolbar", vbNullString)
    AOLTool2& = FindWindowEx(AOLTool1&, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon& = 0&
    For i = 0 To 4
        AOLIcon& = FindWindowEx(AOLTool2&, AOLIcon&, "_AOL_Icon", vbNullString)
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            AOLMenuHandle& = FindWindow("#32768", vbNullString)
            Found = IsWindowVisible(AOLMenuHandle&)
        Loop Until Found <> 0
        MenuCount& = GetMenuItemCount(AOLMenuHandle&)
        For LookFor& = 0& To MenuCount& - 1
            SubMenu& = GetSubMenu(AOLMenuHandle&, 0&)
            SubCount& = GetMenuItemCount(SubMenu&)
            For LookSub& = 0 To SubCount& - 1
                sID& = GetMenuItemID(SubMenu&, LookSub&)
                sString$ = Space$(100)
                Call GetMenuString(SubMenu&, sID&, sString$, 100&, 1&)
                If InStr(LCase$(sString$), LCase$(Txt$)) > 0& Then
                    Call SendMessageLong(AOLWindow&, WM_COMMAND, sID&, 0&)
                    Exit Sub
                End If
            Next LookSub&
        Next LookFor&
    Next i
End Sub

Public Function HexProgram(ExeToHex As String, NewExe As String, FindString As String, ReplaceString As String) As Integer
    'Replaces every occurance of a string with an exe (ExeToHex) and
    'creates a new exe (NewExe) with the changes
    'Returns -1 if error occured or 0 if successful
    'EX:  Call HexProgram("C:\test1.exe", "C:\test2.exe", "punk", "dude")
    On Error Resume Next
    If FileExists(ExeToHex) = False Then HexProgram = -1: Exit Function
    Dim ExeText As String, ExeText2 As String, NewText As String
    Dim X As Integer
    Open NewExe$ For Output As #1
    MsgBox ShortFileName(ExeToHex$)
    Open ShortFileName(ExeToHex$) For Binary As #2
    Do While Not EOF(2): DoEvents
        ExeText2$ = Input(8000, #2)
        Do: DoEvents
            X = InStr(ExeText2$, FindString$)
            If X <> 0 Then
                NewText$ = NewText$ & Mid$(ExeText2$, 1, X - 1) & ReplaceString$
                ExeText2$ = Mid(ExeText2$, X + Len(ReplaceString$))
            End If
            ExeText2$ = NewText$ & ExeText2$
            NewText$ = ""
        Loop Until (X = 0)
        NewText$ = NewText$ & ExeText2$
        Print #1, NewText$
        NewText$ = ""
        If Len(ExeText2$) > 8000 Then ExeText2$ = ""
    Loop
    Close #2
    Close #1
    HexProgram = 0
End Function

Public Sub Form_Print(Frm As Form)
    'Prints a form from the local printer
    Frm.PrintForm
End Sub

Public Function GetOSVersion() As Long
    'Gets the running Windows version
    'Return Values:  1=95/98/ME, 2=NT/2000
    Dim ver As OSVERSIONINFO
    Dim ret As Integer
    ver.dwOSVersionInfoSize = 148
    ver.szCSDVersion = Space$(128)
    ret = GetVersionEx(ver)
    GetOSVersion = ver.dwPlatformId
End Function

Public Sub EmptyRecycleBin(hwnd As Long, Preference As EMPTY_RECYCLE)
    'Empties the recycle bin; for a normal empty, set Preference = NoPreference
    Dim ret As Long
    ret = SHEmptyRecycleBin(hwnd, "", Preference)
    If ret <> 0 Then
        Call SHUpdateRecycleBinIcon
    End If
End Sub

Public Sub OpenFile(FileName As String, Optional WinFocus As VbAppWinStyle = vbNormalFocus)
    'Opens a file
    If FileExists(FileName$) = False Then Exit Sub
    Call Shell(ShortFileName(FileName$), WinFocus)
End Sub

Public Function ColorBar(RedBar As Control, GreenBar As Control, BlueBar As Control)
    'Returns the RGB value of 3 controls (scroll bars)
    'EX:  X = ColorBar(HScroll1, HScroll2, HScroll3)
    ColorBar = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function

Public Sub DestroyControl(hObject As Control)
    'Destroys a control (Command button, Form, etc.)
    'EX:  DestroyControl(Command1)
    DestroyWindow (hObject.hwnd)
End Sub

Public Sub CurrentPosition(hDC As Long, xy As POINTAPI)
    'Returns the current position or the specified DC
    'EX:  Dim point as POINTAPI
    '     Call CurrentPosition(Picture1.hDC, point)
    '     Msgbox (point.x)
    Call GetCurrentPositionEx(hDC, xy)
End Sub
